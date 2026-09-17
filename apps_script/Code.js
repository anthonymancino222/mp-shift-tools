/* Every write in doPostLocked() below (appendRow/deleteRow) runs WITHOUT any
   serialization of its own — Apps Script Web Apps run each incoming request
   in its own separate execution, and nothing about SpreadsheetApp calls
   automatically queues concurrent ones against each other. Under real
   concurrent load (multiple tablets, or just several rapid taps close
   together), two executions' appendRow() calls can race and one write can
   silently clobber/lose the other — each execution still independently
   returns success() since neither one ever finds out about the other, so
   the CLIENT correctly believes it synced (entry.synced = true) even
   though the row never durably landed. Found 2026-09-15: a tablet's own
   sync status read "✅ All synced" while the server was missing the large
   majority of its pallets/entries — this is that bug, not a network
   timeout (a real, separate fetch-timeout fix already exists client-side,
   but it can't fix data that the server itself silently dropped).
   LockService.getScriptLock() forces every request to fully finish before
   the next one starts, at the cost of some queueing latency under
   concurrent load — the right tradeoff for a system whose whole job is
   recording accurate production counts. 30s is generous — if the lock
   genuinely can't be acquired in that window (something else stuck badly),
   surfacing a "Server busy" error the client will retry is far better than
   two requests racing on the sheet. */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: 'Server busy — please retry' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    return doPostLocked(e);
  } finally {
    lock.releaseLock();
  }
}

function doPostLocked(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ---- Live entry log (every "Submit Shift Count" on any device) ----
    // Written immediately so other tablets can see it via doGet before the
    // job is ever finished/sent. No email here — only the final report emails.
    // Idempotent on entryId: weak tablet WiFi means the same entry can get
    // POSTed more than once (a slow request still in flight when a retry
    // timer fires again) — skip the write if this entryId is already here,
    // so duplicates can never land on the sheet regardless of the cause.
    if (data.action === 'log_entry') {
      var activeSheet = getOrCreateActiveEntriesSheet(ss);
      if (!entryIdExists(activeSheet, data.entryId, 12)) {
        activeSheet.appendRow([
          new Date(), data.po, data.product, data.station, data.target,
          data.shift, data.date, data.time, data.good, data.reject, data.flags, data.entryId,
          data.pass || ''
        ]);
      }
      return success();
    }

    // ---- Live pallet log (every "+1 Pallet Complete" tap on any device) ----
    // Same idempotent-on-entryId pattern as log_entry, its own sheet since a
    // pallet tap has no good/reject/flags — just "one more pallet left the line."
    if (data.action === 'log_pallet') {
      var palletsSheet = getOrCreatePalletsSheet(ss);
      if (!entryIdExists(palletsSheet, data.entryId, 8)) {
        palletsSheet.appendRow([
          new Date(), data.po, data.product, data.station,
          data.shift, data.date, data.time, data.entryId,
          data.pass || ''
        ]);
      }
      return success();
    }

    // ---- Notes (job-level or entry-level) ----
    // Same idempotent-on-id pattern as log_entry/log_pallet. Notes are
    // append-only/locked by design — there is no edit or delete action for
    // them, on purpose (see the frontend's confirm-before-adding flow).
    // Scope is 'job' (EntryId blank) or 'entry' (EntryId set to the specific
    // ActiveEntries row this note is attached to).
    if (data.action === 'add_note') {
      var notesSheet = getOrCreateNotesSheet(ss);
      if (!entryIdExists(notesSheet, data.noteId, 6)) {
        notesSheet.appendRow([
          new Date(), data.po, data.pass || '', data.scope, data.entryId || '', data.noteId, data.text
        ]);
      }
      return success();
    }

    // ---- Saved Pallet Counts (Packing Setup presets) ----
    // Shared globally across every tablet — a packer on one device saves a
    // Pieces/Box x Boxes/Pallet setup once, and it's immediately available
    // to anyone on any other device, not just the one that saved it. Same
    // idempotent-on-id pattern as log_entry/log_pallet/add_note.
    if (data.action === 'save_pallet_preset') {
      var presetsSheet = getOrCreatePalletPresetsSheet(ss);
      if (!entryIdExists(presetsSheet, data.presetId, 2)) {
        presetsSheet.appendRow([
          new Date(), data.presetId, data.label, data.piecesPerBox, data.boxesPerPallet
        ]);
      }
      return success();
    }

    // ---- Delete a saved pallet preset ---- Not PIN-gated: unlike delete_job/
    // delete_entry (which erase real production counts), a preset is just a
    // remembered shortcut, low-stakes for any packer to remove.
    if (data.action === 'delete_pallet_preset') {
      deleteRowByEntryId(getOrCreatePalletPresetsSheet(ss), 2, data.presetId);
      return success();
    }

    // ---- Admin delete: an entire stray/junk job (all its ActiveEntries and
    // Pallets rows for this PO+Pass) ----
    // PIN is checked HERE, server-side, against a Script Property that is
    // never sent to the browser — unlike the rest of this file, which all
    // ships publicly on GitHub Pages, this value only ever lives in Apps
    // Script. A wrong PIN just gets a generic error; the real PIN is never
    // revealed either way. Deliberately scoped to ActiveEntries/Pallets only,
    // never Reports — a job that's already been finished and reported can't
    // be erased through this action.
    if (data.action === 'delete_job') {
      if (!checkPin(data.pin)) return pinError();
      removeActiveEntriesForPO(ss, data.po, data.pass);
      removePalletsForPO(ss, data.po, data.pass);
      return success();
    }

    // ---- Admin delete: one specific entry or pallet row, by its own
    // EntryId ---- Same PIN check as delete_job, just scoped to a single row
    // instead of an entire PO+Pass — for correcting one bad duplicate/mistake
    // inside an otherwise-legitimate still-running job.
    if (data.action === 'delete_entry') {
      if (!checkPin(data.pin)) return pinError();
      var targetSheetName = data.entryType === 'pallet' ? 'Pallets' : 'ActiveEntries';
      var targetEntryIdCol = data.entryType === 'pallet' ? 8 : 12;
      deleteRowByEntryId(ss.getSheetByName(targetSheetName), targetEntryIdCol, data.entryId);
      return success();
    }

    // ---- Admin reopen: bring an already-finished job (a Reports row) back
    // into Live Jobs ---- Same PIN as delete_job/delete_entry. Reconstructs
    // ActiveEntries rows straight from that report's own Full Log text (the
    // same "date time | Shift N | Good: x | Reject: y [| Flags: z]" format
    // sendReport() writes), so Good/Reject/date/time/shift/flags come back
    // exactly as they were. Only the AGGREGATE pallet count survives a
    // finished report though — individual pallet timestamps were never kept
    // past finishing — so Total Pallets worth of new Pallets rows are
    // recreated instead, stamped with today's date/time rather than their
    // original ones. The Reports row is removed so History stops showing it
    // as already finished; finishing it again later just appends a fresh
    // Reports row like normal. Matched by PO+Pass+Timestamp (not PO+Pass
    // alone) since the same PO+Pass can legitimately have more than one
    // finished report over time, and only the exact one being reopened
    // should be touched.
    if (data.action === 'reopen_job') {
      if (!checkPin(data.pin)) return pinError();
      var reportsSheetR = ss.getSheets()[0];
      var reportRow = findReportRow(reportsSheetR, data.po, data.pass, data.timestamp);
      if (!reportRow) {
        return ContentService
          .createTextOutput(JSON.stringify({ result: 'error', message: 'No matching report found' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      // Station comes from the client's choice (Anthony wanted the option to
      // bring a job back on either Long or Short Gluer, not just whichever it
      // ran on before) — falls back to the report's original station only if
      // the client somehow didn't send one.
      var reopenStation = data.station || reportRow.station;
      var activeSheetR = getOrCreateActiveEntriesSheet(ss);
      var now = new Date();
      var logLines = String(reportRow.fullLog || '').split('\n').filter(function (l) { return l.trim(); });
      logLines.forEach(function (line, idx) {
        var parsed = parseLogLine(line);
        activeSheetR.appendRow([
          now, reportRow.po, reportRow.product, reopenStation, reportRow.target,
          parsed.shift, parsed.date, parsed.time, parsed.good, parsed.reject, parsed.flags,
          'reopen_' + now.getTime() + '_' + idx, reportRow.pass
        ]);
      });
      var totalPallets = Number(reportRow.totalPallets) || 0;
      if (totalPallets > 0) {
        var palletsSheetR = getOrCreatePalletsSheet(ss);
        var tz = Session.getScriptTimeZone();
        var todayDate = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
        var todayTime = Utilities.formatDate(now, tz, 'hh:mm a');
        for (var pIdx = 0; pIdx < totalPallets; pIdx++) {
          palletsSheetR.appendRow([
            now, reportRow.po, reportRow.product, reopenStation,
            '1', todayDate, todayTime, 'reopen_pallet_' + now.getTime() + '_' + pIdx, reportRow.pass
          ]);
        }
      }
      reportsSheetR.deleteRow(reportRow.rowIndex);
      return success();
    }

    // ---- Final report ("Finish, Send Report") ----
    var reportsSheet = ss.getSheets()[0];
    ensureReportsHeaders(reportsSheet);
    appendRowByHeaders(reportsSheet, {
      'Timestamp': new Date(),
      'PO': data.po,
      'Product': data.product,
      'Station': data.station,
      'Target': data.target,
      'Total Good': data.totalGood,
      // Reject/Hold tracking was removed from the app 2026-09-17 (Anthony:
      // this app is for productivity, not quality holds) — the frontend no
      // longer sends totalReject at all. Column stays (old reports still
      // have real values in it) but every new row just gets a blank cell.
      'Total Reject': data.totalReject || '',
      'Total Pallets': data.totalPallets,
      'Full Log': data.fullLog,
      'Pass': data.pass || '',
      'OversPercent': data.oversPercent || 0,
      'EffectiveTarget': data.effectiveTarget || data.target || 0,
      'PiecesPerBox': data.piecesPerBox || '',
      'BoxesPerPallet': data.boxesPerPallet || '',
      'PalletsNeeded': data.palletsNeeded || '',
      'PctShortVsTarget': data.pctShortVsTarget,
      'PctShortVsEffectiveTarget': data.pctShortVsEffectiveTarget,
      'Shifts': data.shifts || ''
    });

    // Job is done — drop it out of the live "still active" list on every
    // device. Pass is included: a PO's 1st Pass and 2nd Pass run are now
    // fully separate jobs, so finishing one must not touch the other's
    // still-in-progress rows under the same PO. Notes are already captured
    // in this report's email body (see sendReport() in index.html) by the
    // time this runs, so clearing them here doesn't lose anything.
    removeActiveEntriesForPO(ss, data.po, data.pass);
    removePalletsForPO(ss, data.po, data.pass);
    removeNotesForPO(ss, data.po, data.pass);

    // QA group can be added later: to: "jayro@moquinpress.com,qualitygroup@moquinpress.com"
    MailApp.sendEmail({
      to: "jayro@moquinpress.com",
      subject: data.subject,
      body: data.body
    });

    return success();

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Read access for the app's Dashboard (History / Live Jobs / Data tabs) ----
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var activeSheet = ss.getSheetByName('ActiveEntries');
    var palletsSheet = ss.getSheetByName('Pallets');
    var notesSheet = ss.getSheetByName('Notes');

    var result = {
      result: 'success',
      active: activeSheet ? sheetToObjects(activeSheet) : [],
      pallets: palletsSheet ? sheetToObjects(palletsSheet) : [],
      notes: notesSheet ? sheetToObjects(notesSheet) : []
    };

    // "fast" mode (see fetchLive() in index.html) — the Live Jobs tab only
    // ever reads active/pallets/notes above, never Reports or PalletPresets,
    // so skip both here. Reports in particular grows by one row every time
    // ANY job finishes, forever — reading it in full on every poll is what
    // makes a short interval expensive; this lets the Dashboard poll Live
    // Jobs every ~25s (Anthony wants a new job visible within 30s) without
    // that cost scaling with all-time history. A plain (non-fast) call
    // still returns everything, unchanged, for History/Data/exports.
    if (!(e && e.parameter && e.parameter.fast === '1')) {
      var reportsSheet = ss.getSheets()[0];
      var presetsSheet = ss.getSheetByName('PalletPresets');
      result.reports = sheetToObjects(reportsSheet);
      result.palletPresets = presetsSheet ? sheetToObjects(presetsSheet) : [];
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ActiveEntries columns: Timestamp, PO, Product, Station, Target, Shift,
// Date, Time, Good, Reject, Flags, EntryId, Pass — EntryId is column 12.
// Pallets columns: Timestamp, PO, Product, Station, Shift, Date, Time,
// EntryId, Pass — EntryId is column 8. entryIdCol lets both sheets share this
// check. Pass is appended AFTER EntryId in both sheets (not inserted earlier)
// so this column index never has to change.
function entryIdExists(sheet, entryId, entryIdCol) {
  if (!entryId || sheet.getLastRow() < 2) return false;
  var ids = sheet.getRange(2, entryIdCol, sheet.getLastRow() - 1, 1).getValues();
  return ids.some(function (row) { return String(row[0]) === String(entryId); });
}

function getOrCreateActiveEntriesSheet(ss) {
  var sheet = ss.getSheetByName('ActiveEntries');
  if (!sheet) {
    sheet = ss.insertSheet('ActiveEntries');
    sheet.appendRow(['Timestamp', 'PO', 'Product', 'Station', 'Target', 'Shift', 'Date', 'Time', 'Good', 'Reject', 'Flags', 'EntryId', 'Pass']);
  }
  return sheet;
}

function getOrCreatePalletsSheet(ss) {
  var sheet = ss.getSheetByName('Pallets');
  if (!sheet) {
    sheet = ss.insertSheet('Pallets');
    sheet.appendRow(['Timestamp', 'PO', 'Product', 'Station', 'Shift', 'Date', 'Time', 'EntryId', 'Pass']);
  }
  return sheet;
}

// Notes columns: Timestamp, PO, Pass, Scope, EntryId, NoteId, Text.
// Scope is 'job' or 'entry'; EntryId is blank for job-level notes, or the
// ActiveEntries row's own EntryId for entry-level ones. NoteId is column 6.
function getOrCreateNotesSheet(ss) {
  var sheet = ss.getSheetByName('Notes');
  if (!sheet) {
    sheet = ss.insertSheet('Notes');
    sheet.appendRow(['Timestamp', 'PO', 'Pass', 'Scope', 'EntryId', 'NoteId', 'Text']);
  }
  return sheet;
}

// PalletPresets columns: Timestamp, PresetId, Label, PiecesPerBox,
// BoxesPerPallet. Not scoped to any PO/job — these are reusable packing
// setups (e.g. "Standard 12-pack case"), so anyone on any device can save,
// see, and apply one to whatever job they're currently working on.
function getOrCreatePalletPresetsSheet(ss) {
  var sheet = ss.getSheetByName('PalletPresets');
  if (!sheet) {
    sheet = ss.insertSheet('PalletPresets');
    sheet.appendRow(['Timestamp', 'PresetId', 'Label', 'PiecesPerBox', 'BoxesPerPallet']);
  }
  return sheet;
}

// A PO typed/stored as "45401" and "045401" is the same real job — this sheet
// has rows under both spellings from before the frontend normalized leading
// zeros. Without normalizing here too, finishing a job only clears rows that
// exactly match the spelling it was finished under, leaving the other
// spelling's rows stuck in ActiveEntries forever (Live Jobs keeps showing a
// job that was already finished).
function normalizePO(po) {
  var trimmed = String(po == null ? '' : po).trim();
  return /^0+[0-9]+$/.test(trimmed) ? trimmed.replace(/^0+/, '') : trimmed;
}

// Removes every ActiveEntries row for a PO (and, if given, matching Pass too)
// once it's been finished/sent. Deletes bottom-up so row indices don't shift
// mid-loop. Pass is column 13 (0-based index 12) — see the layout comment
// above entryIdExists. If pass is falsy (jobs created before Pass existed),
// falls back to matching by PO alone, same as the original behavior.
function removeActiveEntriesForPO(ss, po, pass) {
  var sheet = ss.getSheetByName('ActiveEntries');
  if (!sheet || sheet.getLastRow() < 2) return;
  var target = normalizePO(po);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var passMatches = !pass || String(data[i][12] || '') === String(pass);
    if (normalizePO(data[i][1]) === target && passMatches) {
      sheet.deleteRow(i + 1);
    }
  }
}

// Same cleanup as removeActiveEntriesForPO, for the Pallets sheet. Pass is
// column 9 (0-based index 8) there — see getOrCreatePalletsSheet's header.
function removePalletsForPO(ss, po, pass) {
  var sheet = ss.getSheetByName('Pallets');
  if (!sheet || sheet.getLastRow() < 2) return;
  var target = normalizePO(po);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var passMatches = !pass || String(data[i][8] || '') === String(pass);
    if (normalizePO(data[i][1]) === target && passMatches) {
      sheet.deleteRow(i + 1);
    }
  }
}

// Same cleanup as removeActiveEntriesForPO, for the Notes sheet. PO is
// column 2 (0-based index 1), Pass is column 3 (0-based index 2).
function removeNotesForPO(ss, po, pass) {
  var sheet = ss.getSheetByName('Notes');
  if (!sheet || sheet.getLastRow() < 2) return;
  var target = normalizePO(po);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var passMatches = !pass || String(data[i][2] || '') === String(pass);
    if (normalizePO(data[i][1]) === target && passMatches) {
      sheet.deleteRow(i + 1);
    }
  }
}

// Finds the exact Reports row a reopen_job call refers to. Matched by
// PO+Pass+Timestamp (within 1s, to absorb any float/precision rounding in
// the ISO round-trip) rather than PO+Pass alone, since the same PO+Pass can
// have more than one finished report over time (finished, later restarted,
// finished again) — only the specific one being reopened should be touched.
// Reports columns: Timestamp, PO, Product, Station, Target, Total Good,
// Total Reject, Total Pallets, Full Log, Pass.
function findReportRow(sheet, po, pass, timestamp) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  var data = sheet.getDataRange().getValues();
  var target = normalizePO(po);
  var targetTs = timestamp ? new Date(timestamp).getTime() : null;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (normalizePO(row[1]) !== target) continue;
    if (pass && String(row[9] || '') !== String(pass)) continue;
    if (targetTs != null) {
      var rowTs = row[0] instanceof Date ? row[0].getTime() : new Date(row[0]).getTime();
      if (Math.abs(rowTs - targetTs) > 1000) continue;
    }
    return {
      rowIndex: i + 1,
      po: row[1], product: row[2], station: row[3], target: row[4],
      totalGood: row[5], totalReject: row[6], totalPallets: row[7],
      fullLog: row[8], pass: row[9]
    };
  }
  return null;
}

// Parses one line of a report's Full Log back into structured fields — same
// line format sendReport() writes client-side (index.html's own
// parseFullLogText does the identical parse for display, this is the
// server-side equivalent for reconstruction).
function parseLogLine(line) {
  var parts = line.split('|').map(function (p) { return p.trim(); });
  var dateTime = (parts[0] || '').split(' ');
  var date = dateTime[0] || '';
  var time = dateTime.slice(1).join(' ');
  var shiftMatch = (parts[1] || '').match(/(\d)/);
  var goodMatch = (parts[2] || '').match(/([\d,]+)/);
  // Reject/Flags are optional trailing fields found by their own label, not
  // a fixed index — a line logged before Reject tracking was removed from
  // the app (2026-09-17) has one extra "| Reject: N" field ahead of Flags
  // that a newer line won't, so a fixed-position read would silently
  // swallow Flags when reopening a job finished after that date.
  var rejectPart = null, flagsPart = '';
  parts.forEach(function (p) {
    if (/^Reject:/i.test(p)) rejectPart = p;
    if (/^Flags:/i.test(p)) flagsPart = p.replace(/^Flags:\s*/i, '');
  });
  var rejectMatch = rejectPart ? rejectPart.match(/([\d,]+)/) : null;
  return {
    date: date,
    time: time,
    shift: shiftMatch ? shiftMatch[1] : '1',
    good: goodMatch ? parseInt(goodMatch[1].replace(/,/g, ''), 10) : 0,
    reject: rejectMatch ? parseInt(rejectMatch[1].replace(/,/g, ''), 10) : 0,
    flags: flagsPart
  };
}

// Checks a submitted PIN against the DELETE_PIN Script Property (Project
// Settings > Script Properties in the Apps Script editor — set/changed there
// directly, never in this file, so it's never part of the public frontend
// code). Returns false (not an error) if no PIN has been configured yet, so
// delete actions simply refuse to work until one is set up.
function checkPin(pin) {
  var stored = PropertiesService.getScriptProperties().getProperty('DELETE_PIN');
  return !!stored && String(pin || '') === stored;
}

function pinError() {
  return ContentService
    .createTextOutput(JSON.stringify({ result: 'error', message: 'Incorrect PIN' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Deletes the one row matching entryId in the given sheet/column (1-based,
// same convention as entryIdExists). Silently no-ops if the sheet doesn't
// exist or no row matches — deleting something already gone isn't an error.
function deleteRowByEntryId(sheet, entryIdCol, entryId) {
  if (!sheet || !entryId || sheet.getLastRow() < 2) return;
  var ids = sheet.getRange(2, entryIdCol, sheet.getLastRow() - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(entryId)) {
      sheet.deleteRow(i + 2); // +2: 0-based array index -> 1-based row, plus the header row
      return;
    }
  }
}

function sheetToObjects(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  return data.slice(1).map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

// The Reports sheet's full column list — as of 2026-09-16 this includes the
// extra history fields (packing setup, pallets needed, shortfall %, shifts)
// Anthony asked to have captured alongside every finished job. Real
// production history already exists under just the first 10 of these
// columns; ensureReportsHeaders() appends whichever of the rest are still
// missing to the END of row 1, exactly once, and never touches/reorders any
// existing header — sheetToObjects() maps by header NAME, not position, so
// old rows simply read back blank for the new ones, no migration needed.
var REPORTS_HEADERS = ['Timestamp', 'PO', 'Product', 'Station', 'Target', 'Total Good', 'Total Reject', 'Total Pallets', 'Full Log', 'Pass',
  'OversPercent', 'EffectiveTarget', 'PiecesPerBox', 'BoxesPerPallet', 'PalletsNeeded', 'PctShortVsTarget', 'PctShortVsEffectiveTarget', 'Shifts'];
function ensureReportsHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(REPORTS_HEADERS);
    return;
  }
  var lastCol = sheet.getLastColumn();
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var have = {};
  existing.forEach(function (h) { have[h] = true; });
  var missing = REPORTS_HEADERS.filter(function (h) { return !have[h]; });
  if (missing.length) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }
}

// Appends one row built from a {headerName: value} map, in whatever column
// order the sheet's OWN header row currently has — safer than a fixed
// positional array once a sheet's columns can grow over time (see
// ensureReportsHeaders above). Any header the map doesn't mention is left
// blank for that row.
function appendRowByHeaders(sheet, valuesByHeader) {
  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = headerRow.map(function (h) { return (h in valuesByHeader) ? valuesByHeader[h] : ''; });
  sheet.appendRow(row);
}

function success() {
  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}
