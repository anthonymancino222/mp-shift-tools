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
        appendRowByHeaders(activeSheet, {
          'Timestamp': new Date(), 'PO': data.po, 'Product': data.product, 'Station': data.station,
          'Target': data.target, 'Shift': data.shift, 'Date': data.date, 'Time': data.time,
          'Good': data.good, 'Reject': data.reject, 'Flags': data.flags, 'EntryId': data.entryId,
          'Pass': data.pass || '', 'DeviceId': data.deviceId || ''
        });
      }
      return success();
    }

    // ---- Live pallet log (every "+1 Pallet Complete" tap on any device) ----
    // Same idempotent-on-entryId pattern as log_entry, its own sheet since a
    // pallet tap has no good/reject/flags — just "one more pallet left the line."
    if (data.action === 'log_pallet') {
      var palletsSheet = getOrCreatePalletsSheet(ss);
      if (!entryIdExists(palletsSheet, data.entryId, 8)) {
        appendRowByHeaders(palletsSheet, {
          'Timestamp': new Date(), 'PO': data.po, 'Product': data.product, 'Station': data.station,
          'Shift': data.shift, 'Date': data.date, 'Time': data.time, 'EntryId': data.entryId,
          'Pass': data.pass || '', 'DeviceId': data.deviceId || '',
          // The Good-count entry THIS pallet auto-created client-side (see
          // createAutoEntryFromPallet in index.html) — stored so a later
          // pull on ANY device (a second tablet, this same device after a
          // cache clear, or a reopened job) can reconstruct the same
          // pallet<->entry link a live, never-synced job already has,
          // letting Undo Last Pallet correctly reverse the Good Count it
          // added instead of only decrementing the pallet tally.
          'LinkedEntryId': data.linkedEntryId || ''
        });
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
        appendRowByHeaders(notesSheet, {
          'Timestamp': new Date(), 'PO': data.po, 'Pass': data.pass || '', 'Scope': data.scope,
          'EntryId': data.entryId || '', 'NoteId': data.noteId, 'Text': data.text, 'DeviceId': data.deviceId || ''
        });
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

    // ---- Clear THIS DEVICE's own still-active rows for one PO+Pass ----
    // No PIN: scoped strictly to the calling device's own DeviceId, same
    // trust model finish_report already uses (a device can always clean up
    // rows it pushed itself — never anyone else's). Added 2026-09-17 for
    // Start Over: that action is local-only by design (no PIN, no server
    // call at all, so it stays instant) — but if the job being cleared had
    // already pushed real pallets/entries, those rows sat orphaned in Live
    // Jobs forever with the only cleanup tool (delete_job) being an
    // unscoped wipe of the ENTIRE PO+Pass — unsafe now that a different
    // device can legitimately be running its own separate job under that
    // exact same PO+Pass. deviceId is REQUIRED (unlike finish_report, where
    // it's expected but tolerated blank) — without one, this would just be
    // an unauthenticated wipe of a whole PO+Pass, which is exactly what the
    // PIN-gated delete_job exists for instead.
    if (data.action === 'clear_own_active') {
      if (!data.deviceId) return success(); // nothing to scope to — safe no-op
      removeActiveEntriesForPO(ss, data.po, data.pass, data.deviceId);
      removePalletsForPO(ss, data.po, data.pass, data.deviceId);
      removeNotesForPO(ss, data.po, data.pass, data.deviceId);
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
      var palletsSheetR = getOrCreatePalletsSheet(ss);
      var now = new Date();
      var logLines = String(reportRow.fullLog || '').split('\n').filter(function (l) { return l.trim(); });
      logLines.forEach(function (line, idx) {
        var parsed = parseLogLine(line);
        var entryId = 'reopen_' + now.getTime() + '_' + idx;
        // +idx ms keeps every reconstructed row's Timestamp in the SAME
        // relative order the original log lines were in — matters because
        // the client's "Undo Last Pallet" picks whichever pallet has the
        // latest Timestamp; giving every row the exact same `now` would
        // make that a coin flip instead of actually the last one.
        var rowTs = new Date(now.getTime() + idx);
        // Tagged with the reopening device's own DeviceId (not the original
        // logger's, long gone once a job finishes) — these rows now belong
        // to whichever device reopened the job, same as any freshly-created
        // job would be.
        appendRowByHeaders(activeSheetR, {
          'Timestamp': rowTs, 'PO': reportRow.po, 'Product': reportRow.product, 'Station': reopenStation,
          'Target': reportRow.target, 'Shift': parsed.shift, 'Date': parsed.date, 'Time': parsed.time,
          'Good': parsed.good, 'Reject': parsed.reject, 'Flags': parsed.flags,
          'EntryId': entryId, 'Pass': reportRow.pass, 'DeviceId': data.deviceId || ''
        });
        // A pallet row is only ever recreated for a log line that was
        // ITSELF a pallet-completion entry (the "| Pallet: full/partial"
        // marker sendReport() now writes — see parseLogLine). Tagging it
        // with LinkedEntryId (this same line's entryId, just above) is what
        // lets the client's undoPallet() correctly reverse the Good Count
        // that specific pallet added, exactly like a live job — the old
        // code instead created `totalPallets` blank, disconnected rows with
        // no way to tie back to a specific entry at all.
        if (parsed.fromPallet) {
          appendRowByHeaders(palletsSheetR, {
            'Timestamp': rowTs, 'PO': reportRow.po, 'Product': reportRow.product, 'Station': reopenStation,
            'Shift': parsed.shift, 'Date': parsed.date, 'Time': parsed.time,
            'EntryId': 'reopen_pallet_' + now.getTime() + '_' + idx, 'Pass': reportRow.pass, 'DeviceId': data.deviceId || '',
            'LinkedEntryId': entryId
          });
        }
      });
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

    // Job is done — drop THIS DEVICE's rows out of the live "still active"
    // list. Pass is included: a PO's 1st Pass and 2nd Pass run are fully
    // separate jobs, so finishing one must not touch the other's still-in-
    // progress rows under the same PO. DeviceId is included too (2026-09-17):
    // another device can legitimately be running its own separate job under
    // this exact same PO+Pass now, and finishing this one must never delete
    // that device's still-in-progress rows. Notes are already captured in
    // this report's email body (see sendReport() in index.html) by the time
    // this runs, so clearing them here doesn't lose anything.
    removeActiveEntriesForPO(ss, data.po, data.pass, data.deviceId);
    removePalletsForPO(ss, data.po, data.pass, data.deviceId);
    removeNotesForPO(ss, data.po, data.pass, data.deviceId);

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

// DeviceId is appended LAST on all three sheets (never inserted earlier) so
// its column index never has to change, same reasoning as Pass before it.
// Added 2026-09-17 so two devices logging the SAME PO+Pass are never merged
// into one job (Anthony: "treat it as a separate entry... no device should
// ever auto-merge") — every write is tagged with the device that made it,
// and index.html's own sync code only ever pulls rows back matching its own
// DeviceId into a local job. Live Jobs (doGet, unfiltered) is unaffected —
// it's meant to show every device's activity, tagged or not.
var ACTIVE_ENTRIES_HEADERS = ['Timestamp', 'PO', 'Product', 'Station', 'Target', 'Shift', 'Date', 'Time', 'Good', 'Reject', 'Flags', 'EntryId', 'Pass', 'DeviceId'];
// LinkedEntryId (2026-09-21) is appended LAST, same reasoning as DeviceId
// before it — an existing sheet gets it auto-added at the end by
// ensureHeaders(), so entryIdExists()'s hardcoded EntryId column index (8)
// never has to change.
var PALLETS_HEADERS = ['Timestamp', 'PO', 'Product', 'Station', 'Shift', 'Date', 'Time', 'EntryId', 'Pass', 'DeviceId', 'LinkedEntryId'];
var NOTES_HEADERS = ['Timestamp', 'PO', 'Pass', 'Scope', 'EntryId', 'NoteId', 'Text', 'DeviceId'];

function getOrCreateActiveEntriesSheet(ss) {
  var sheet = ss.getSheetByName('ActiveEntries');
  if (!sheet) {
    sheet = ss.insertSheet('ActiveEntries');
    sheet.appendRow(ACTIVE_ENTRIES_HEADERS);
  } else {
    ensureHeaders(sheet, ACTIVE_ENTRIES_HEADERS);
  }
  return sheet;
}

function getOrCreatePalletsSheet(ss) {
  var sheet = ss.getSheetByName('Pallets');
  if (!sheet) {
    sheet = ss.insertSheet('Pallets');
    sheet.appendRow(PALLETS_HEADERS);
  } else {
    ensureHeaders(sheet, PALLETS_HEADERS);
  }
  return sheet;
}

// Notes columns: Timestamp, PO, Pass, Scope, EntryId, NoteId, Text, DeviceId.
// Scope is 'job' or 'entry'; EntryId is blank for job-level notes, or the
// ActiveEntries row's own EntryId for entry-level ones. NoteId is column 6.
function getOrCreateNotesSheet(ss) {
  var sheet = ss.getSheetByName('Notes');
  if (!sheet) {
    sheet = ss.insertSheet('Notes');
    sheet.appendRow(NOTES_HEADERS);
  } else {
    ensureHeaders(sheet, NOTES_HEADERS);
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
// deviceId (optional) additionally requires column 14 (0-based index 13) to
// match — added 2026-09-17 so that finishing ONE device's job never deletes
// a different device's still-in-progress rows for the exact same PO+Pass,
// now that two devices are explicitly allowed to run the same PO+Pass as
// fully separate jobs (see the DeviceId comment on getOrCreateActiveEntriesSheet).
// Omit deviceId (delete_job, an admin action from the aggregated Live Jobs
// view) to keep the old blunt "wipe this PO+Pass for everyone" behavior.
// A row with a BLANK DeviceId (any row logged before that column existed)
// is always eligible for cleanup regardless of who's finishing — found
// 2026-09-17, hours after DeviceId shipped: a real job logged before that
// deploy had blank DeviceId on every row, so finishing it from any device
// (a real, non-blank id) never matched, and the "already finished" job sat
// in Live Jobs forever even though the Report itself sent and saved
// correctly. Blank rows have no exclusive owner to protect in the first
// place, unlike a row genuinely tagged by a DIFFERENT device.
function removeActiveEntriesForPO(ss, po, pass, deviceId) {
  var sheet = ss.getSheetByName('ActiveEntries');
  if (!sheet || sheet.getLastRow() < 2) return;
  var target = normalizePO(po);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var passMatches = !pass || String(data[i][12] || '') === String(pass);
    var rowDeviceId = String(data[i][13] || '');
    var deviceMatches = !deviceId || !rowDeviceId || rowDeviceId === String(deviceId);
    if (normalizePO(data[i][1]) === target && passMatches && deviceMatches) {
      sheet.deleteRow(i + 1);
    }
  }
}

// Same cleanup as removeActiveEntriesForPO, for the Pallets sheet. Pass is
// column 9 (0-based index 8), DeviceId column 10 (0-based index 9).
function removePalletsForPO(ss, po, pass, deviceId) {
  var sheet = ss.getSheetByName('Pallets');
  if (!sheet || sheet.getLastRow() < 2) return;
  var target = normalizePO(po);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var passMatches = !pass || String(data[i][8] || '') === String(pass);
    var rowDeviceId = String(data[i][9] || '');
    var deviceMatches = !deviceId || !rowDeviceId || rowDeviceId === String(deviceId);
    if (normalizePO(data[i][1]) === target && passMatches && deviceMatches) {
      sheet.deleteRow(i + 1);
    }
  }
}

// Same cleanup as removeActiveEntriesForPO, for the Notes sheet. PO is
// column 2 (0-based index 1), Pass is column 3 (0-based index 2), DeviceId
// column 8 (0-based index 7).
function removeNotesForPO(ss, po, pass, deviceId) {
  var sheet = ss.getSheetByName('Notes');
  if (!sheet || sheet.getLastRow() < 2) return;
  var target = normalizePO(po);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var passMatches = !pass || String(data[i][2] || '') === String(pass);
    var rowDeviceId = String(data[i][7] || '');
    var deviceMatches = !deviceId || !rowDeviceId || rowDeviceId === String(deviceId);
    if (normalizePO(data[i][1]) === target && passMatches && deviceMatches) {
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
  // "| Pallet: full" / "| Pallet: partial" (added 2026-09-21) marks a line
  // that was originally an auto-entry created by +1 Pallet Complete/Add
  // Partial Pallet, not a manually-typed Good count — reopen_job uses this
  // to know which lines need a matching Pallets row recreated, tied back
  // via LinkedEntryId. Absent on any line logged before this existed, or on
  // a genuinely manual entry — both correctly parse as fromPallet: false.
  var rejectPart = null, flagsPart = '', palletPart = null;
  parts.forEach(function (p) {
    if (/^Reject:/i.test(p)) rejectPart = p;
    if (/^Flags:/i.test(p)) flagsPart = p.replace(/^Flags:\s*/i, '');
    if (/^Pallet:/i.test(p)) palletPart = p;
  });
  var rejectMatch = rejectPart ? rejectPart.match(/([\d,]+)/) : null;
  return {
    date: date,
    time: time,
    shift: shiftMatch ? shiftMatch[1] : '1',
    good: goodMatch ? parseInt(goodMatch[1].replace(/,/g, ''), 10) : 0,
    reject: rejectMatch ? parseInt(rejectMatch[1].replace(/,/g, ''), 10) : 0,
    flags: flagsPart,
    fromPallet: !!palletPart,
    partial: !!(palletPart && /partial/i.test(palletPart))
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
  ensureHeaders(sheet, REPORTS_HEADERS);
}

// Generic version of the above — appends whichever of `headers` a sheet is
// still missing to the END of row 1, exactly once, never touching/reordering
// any existing header. Used for Reports (History fields, 2026-09-16) and now
// ActiveEntries/Pallets/Notes (DeviceId, 2026-09-17). Safe to call on every
// request; a sheet that already has every header is untouched.
function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }
  var lastCol = sheet.getLastColumn();
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var have = {};
  existing.forEach(function (h) { have[h] = true; });
  var missing = headers.filter(function (h) { return !have[h]; });
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
