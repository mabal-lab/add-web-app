// =============================================================================
// Focus Task Manager — Google Apps Script Backend
// -----------------------------------------------------------------------------
// Paste this entire file into your Apps Script editor, deploy as a web app,
// then copy the /exec URL into SHEETS_URL in index.html.
// =============================================================================

// Column order must exactly match the spreadsheet column order.
const TASK_COLS = [
  'id', 'name', 'description', 'category', 'effort',
  'startDate', 'dueDate', 'status', 'createdAt', 'lastModified'
];
const SUBTASK_COLS = [
  'id', 'taskId', 'scheduledDate', 'title', 'notes',
  'estimatedMinutes', 'completed', 'completedAt', 'sortOrder'
];

// These columns hold date-only strings (YYYY-MM-DD); Sheets auto-parses them
// into Date objects, so we convert back in readSheet().
const TASK_DATE_COLS    = ['startDate', 'dueDate'];
const TASK_DT_COLS      = ['createdAt', 'lastModified'];
const SUBTASK_DATE_COLS = ['scheduledDate'];
const SUBTASK_DT_COLS   = ['completedAt'];

// ── GET: return all tasks and subtasks as JSON ────────────────────────────────
function doGet(e) {
  try {
    const ss            = SpreadsheetApp.getActiveSpreadsheet();
    const tasksSheet    = ss.getSheetByName('Tasks');
    const subtasksSheet = ss.getSheetByName('Subtasks');

    if (!tasksSheet || !subtasksSheet) {
      return respond({
        error: 'Sheets named "Tasks" and "Subtasks" not found. ' +
               'Double-check the sheet tab names (case-sensitive).'
      });
    }

    const tasks    = readSheet(tasksSheet,    TASK_COLS,    TASK_DATE_COLS,    TASK_DT_COLS);
    const subtasks = readSheet(subtasksSheet, SUBTASK_COLS, SUBTASK_DATE_COLS, SUBTASK_DT_COLS);

    return respond({ ok: true, tasks: tasks, subtasks: subtasks });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ── POST: replace all data in both sheets ─────────────────────────────────────
// Body: { action: "replaceAll", tasks: [...], subtasks: [...] }
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    if (body.action === 'replaceAll') {
      const tasksSheet    = ss.getSheetByName('Tasks');
      const subtasksSheet = ss.getSheetByName('Subtasks');
      rewriteSheet(tasksSheet,    TASK_COLS,    body.tasks    || []);
      rewriteSheet(subtasksSheet, SUBTASK_COLS, body.subtasks || []);
      return respond({ ok: true });
    }

    return respond({ error: 'Unknown action: ' + body.action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ── Clear all data rows in a sheet and rewrite from the given array ───────────
function rewriteSheet(sheet, cols, rows) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    // Clear existing data (keep header in row 1)
    sheet.getRange(2, 1, lastRow - 1, cols.length).clearContent();
  }
  if (rows.length === 0) return;

  var data = rows.map(function(obj) {
    return cols.map(function(col) {
      var v = obj[col];
      return (v === undefined || v === null) ? '' : v;
    });
  });

  sheet.getRange(2, 1, data.length, cols.length).setValues(data);
}

// ── Read a sheet into an array of plain objects ───────────────────────────────
function readSheet(sheet, cols, dateCols, dtCols) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, cols.length).getValues();

  return data
    .map(function(row) {
      var obj = {};
      cols.forEach(function(col, i) {
        obj[col] = parseCell(row[i], col, dateCols, dtCols);
      });
      return obj;
    })
    .filter(function(obj) { return !!obj.id; }); // skip blank rows
}

// ── Convert a Sheets cell value to the right JS type ─────────────────────────
// Sheets auto-converts YYYY-MM-DD and ISO datetime strings to Date objects.
// We convert them back to strings here so the client receives plain JSON.
function parseCell(v, col, dateCols, dtCols) {
  if (v === '' || v === null || v === undefined) return null;

  if (v instanceof Date) {
    if (dtCols && dtCols.indexOf(col) !== -1) {
      // Full ISO datetime (e.g. createdAt, lastModified)
      return v.toISOString();
    }
    if (dateCols && dateCols.indexOf(col) !== -1) {
      // Date-only string (e.g. scheduledDate, startDate)
      return Utilities.formatDate(v, 'UTC', 'yyyy-MM-dd');
    }
    return v.toISOString(); // fallback
  }

  return v; // booleans, numbers, and plain strings pass through unchanged
}

// ── Wrap response as JSON with CORS headers ───────────────────────────────────
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
