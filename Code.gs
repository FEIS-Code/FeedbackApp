// ============================================================
// Feedback App — Google Apps Script Backend
// Five Elements International School — Team Suggestors
// ============================================================

const SPREADSHEET_ID = '1grU9PMVJSaF6mWG2KpX8_eDZyQyNl2qnNOALg1TnpR8';
const FEEDBACK_SHEET = 'Feedback';
const USERS_SHEET = 'Users';
const CATEGORIES_SHEET = 'Categories';

function getSheet(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toLowerCase() === name.trim().toLowerCase()) return sheets[i];
  }
  return null;
}

// --- Web App Entry Points ---

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'all';
  var result;
  switch (action) {
    case 'all': result = getFeedback(); break;
    case 'categories': result = getCategories(); break;
    case 'stats': result = getStats(); break;
    case 'users':
      result = getUsers(e.parameter.u, e.parameter.p);
      break;
    default: result = { error: 'Unknown action' };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var result;
  switch (data.action) {
    case 'submit': result = submitFeedback(data); break;
    case 'login': result = login(data.username, data.password); break;
    case 'respond': result = respondToFeedback(data); break;
    case 'delete': result = deleteFeedback(data); break;
    case 'saveCategories': result = saveCategories(data); break;
    case 'saveUsers': result = saveUsers(data); break;
    case 'setupData': 
      var auth = login(data.username, data.password);
      if (auth.success && auth.role === 'admin') { setupData(); result = {success:true}; }
      else result = {success:false, message:'Unauthorized'};
      break;
    default: result = { error: 'Unknown action' };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// --- Data Functions ---

function getFeedback() {
  var sheet = getSheet(FEEDBACK_SHEET);
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
    rows.push(obj);
  }
  return rows.reverse();
}

function submitFeedback(data) {
  var sheet = getSheet(FEEDBACK_SHEET);
  if (!sheet) return { success: false, message: 'Sheet not found' };
  var id = 'FB-' + Date.now().toString(36).toUpperCase();
  var now = new Date();
  var date = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a');
  sheet.appendRow([
    id,
    date,
    data.name || 'Anonymous',
    data.category || 'General',
    parseInt(data.rating) || 0,
    (data.feedback || '').trim(),
    '', // Admin Response
    'New' // Status
  ]);
  return { success: true, id: id };
}

function respondToFeedback(data) {
  var auth = login(data.auth.username, data.auth.password);
  if (!auth.success || auth.role !== 'admin') return { success: false, message: 'Unauthorized' };
  var sheet = getSheet(FEEDBACK_SHEET);
  var all = sheet.getDataRange().getDisplayValues();
  var headers = all[0];
  var idCol = headers.indexOf('ID');
  var respCol = headers.indexOf('Response');
  var statusCol = headers.indexOf('Status');
  for (var i = 1; i < all.length; i++) {
    if (all[i][idCol] === data.id) {
      if (respCol >= 0) sheet.getRange(i + 1, respCol + 1).setValue(data.response);
      if (statusCol >= 0) sheet.getRange(i + 1, statusCol + 1).setValue('Responded');
      return { success: true };
    }
  }
  return { success: false, message: 'Not found' };
}

function deleteFeedback(data) {
  var auth = login(data.auth.username, data.auth.password);
  if (!auth.success || auth.role !== 'admin') return { success: false, message: 'Unauthorized' };
  var sheet = getSheet(FEEDBACK_SHEET);
  var all = sheet.getDataRange().getDisplayValues();
  var idCol = all[0].indexOf('ID');
  for (var i = 1; i < all.length; i++) {
    if (all[i][idCol] === data.id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'Not found' };
}

function getCategories() {
  var sheet = getSheet(CATEGORIES_SHEET);
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) result.push(data[i][0].trim());
  }
  return result;
}

function saveCategories(data) {
  var auth = login(data.auth.username, data.auth.password);
  if (!auth.success || auth.role !== 'admin') return { success: false, message: 'Unauthorized' };
  var sheet = getSheet(CATEGORIES_SHEET) || SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(CATEGORIES_SHEET);
  sheet.clear();
  sheet.appendRow(['Category']);
  (data.categories || []).forEach(function(c) { sheet.appendRow([c]); });
  return { success: true };
}

function getStats() {
  var all = getFeedback();
  var total = all.length;
  var avgRating = 0, catCount = {}, statusCount = {New:0, Responded:0};
  for (var i = 0; i < all.length; i++) {
    avgRating += parseInt(all[i].Rating) || 0;
    var cat = all[i].Category || 'General';
    catCount[cat] = (catCount[cat] || 0) + 1;
    var st = all[i].Status || 'New';
    statusCount[st] = (statusCount[st] || 0) + 1;
  }
  return { total: total, avgRating: total ? (avgRating / total).toFixed(1) : 0, categories: catCount, statuses: statusCount };
}

function login(username, password) {
  var sheet = getSheet(USERS_SHEET);
  if (!sheet) return { success: false, message: 'Users sheet not found' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(username).trim() &&
        String(data[i][1]).trim() === String(password).trim()) {
      return { success: true, role: String(data[i][2]).trim(), displayName: String(data[i][3]).trim(), username: String(data[i][0]).trim() };
    }
  }
  return { success: false, message: 'Invalid credentials' };
}

function getUsers(username, password) {
  var auth = login(username, password);
  if (!auth.success || auth.role !== 'admin') return { error: 'Unauthorized' };
  var sheet = getSheet(USERS_SHEET);
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) result.push({ username: String(data[i][0]), password: String(data[i][1]), role: String(data[i][2]), displayName: String(data[i][3]) });
  }
  return result;
}

function saveUsers(data) {
  var auth = login(data.auth.username, data.auth.password);
  if (!auth.success || auth.role !== 'admin') return { success: false, message: 'Unauthorized' };
  var sheet = getSheet(USERS_SHEET);
  sheet.clear();
  sheet.appendRow(['Username', 'Password', 'Role', 'DisplayName']);
  (data.users || []).forEach(function(u) { sheet.appendRow([u.username, u.password, u.role, u.displayName]); });
  return { success: true };
}

// --- Setup ---

function setupData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var us = ss.getSheetByName(USERS_SHEET) || ss.insertSheet(USERS_SHEET);
  us.clear();
  us.appendRow(['Username', 'Password', 'Role', 'DisplayName']);
  us.appendRow(['admin', 'admin123', 'admin', 'Administrator']);
  us.appendRow(['srilatha', 'teach123', 'teacher', 'Ms. Srilatha']);

  var cs = ss.getSheetByName(CATEGORIES_SHEET) || ss.insertSheet(CATEGORIES_SHEET);
  cs.clear();
  cs.appendRow(['Category']);
  ['Teaching', 'Facilities', 'Food/Canteen', 'Events', 'Sports', 'Library', 'Transport', 'Safety', 'Cleanliness', 'General'].forEach(function(c) { cs.appendRow([c]); });

  var fs = ss.getSheetByName(FEEDBACK_SHEET) || ss.insertSheet(FEEDBACK_SHEET);
  fs.clear();
  fs.appendRow(['ID', 'Date', 'Name', 'Category', 'Rating', 'Feedback', 'Response', 'Status']);

  Logger.log('Setup complete');
}
