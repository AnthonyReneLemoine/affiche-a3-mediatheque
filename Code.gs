function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Créateur d\'affiches Médiathèque A3');
}

function createPoster(payload) {
  var data = JSON.parse(payload);
  var html = buildPosterHtml(data);

  var pdfBlob = HtmlService.createHtmlOutput(html).getAs(MimeType.PDF);
  var dateForName = formatDateForFilename(data.dates || '');
  var titleForName = sanitizeFilename(data.title || 'sans-titre');
  pdfBlob.setName('MED-SARZEAU-' + dateForName + '-' + titleForName + '.pdf');
  var targetFolder = ensurePosterFolder();
  var pdfFile = targetFolder.createFile(pdfBlob);

  var entryId = data.entryId || '';
  var entryRecord = savePosterEntry(entryId, data, pdfFile.getUrl());

  return {
    entryId: entryRecord.id,
    pdfUrl: pdfFile.getUrl()
  };
}

function listPosterEntries() {
  var sheet = ensurePosterSheet();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  var headers = values[0];
  var tz = Session.getScriptTimeZone();
  var rows = values.slice(1);
  return rows.map(function(row) {
    var entry = {};
    headers.forEach(function(header, index) {
      var value = row[index];
      if (value instanceof Date) {
        entry[header] = (header === 'createdAt' || header === 'updatedAt')
          ? Utilities.formatDate(value, tz, 'yyyy-MM-dd HH:mm')
          : Utilities.formatDate(value, tz, 'dd/MM/yyyy');
      } else {
        entry[header] = value !== null && value !== undefined ? String(value) : '';
      }
    });
    return entry;
  }).filter(function(e) { return e.id; }).reverse();
}

function ensurePosterSpreadsheet() {
  var sheetName = 'AFFICHE MEDIATHEQUE A3';
  var propertyKey = 'AFFICHE_MEDIATHEQUE_A3_SHEET_ID';
  var storedId = PropertiesService.getScriptProperties().getProperty(propertyKey);
  if (storedId) {
    try { return SpreadsheetApp.openById(storedId); }
    catch (e) { PropertiesService.getScriptProperties().deleteProperty(propertyKey); }
  }
  var files = DriveApp.getFilesByName(sheetName);
  if (files.hasNext()) {
    var existing = files.next();
    PropertiesService.getScriptProperties().setProperty(propertyKey, existing.getId());
    return SpreadsheetApp.open(existing);
  }
  var created = SpreadsheetApp.create(sheetName);
  PropertiesService.getScriptProperties().setProperty(propertyKey, created.getId());
  var targetFolder = ensurePosterFolder();
  var createdFile = DriveApp.getFileById(created.getId());
  targetFolder.addFile(createdFile);
  DriveApp.getRootFolder().removeFile(createdFile);
  return created;
}

function ensurePosterSheet() {
  var sheetName = 'AFFICHE MEDIATHEQUE A3';
  var ss = ensurePosterSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  ensureSheetHeaders(sheet);
  return sheet;
}

function ensureSheetHeaders(sheet) {
  var headers = ['id','createdAt','updatedAt','title','titleLine2','subtitle','description','dates','infos','publicCible','siteUrl','pdfUrl'];
  var existing = sheet.getLastColumn() >= headers.length
    ? sheet.getRange(1, 1, 1, headers.length).getValues()[0] : [];
  if (existing.length !== headers.length || existing.some(function(v,i) { return v !== headers[i]; })) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function ensurePosterFolder() {
  var folderName = 'AFFICHE MEDIATHEQUE A3';
  var folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

function savePosterEntry(entryId, data, pdfUrl) {
  var sheet = ensurePosterSheet();
  var now = new Date();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var payload = {
    id: entryId || Utilities.getUuid(), createdAt: now, updatedAt: now,
    title: data.title || '', titleLine2: data.titleLine2 || '',
    subtitle: data.subtitle || '', description: data.description || '',
    dates: data.dates || '', infos: data.infos || '',
    publicCible: data.publicCible || '', siteUrl: data.siteUrl || '', pdfUrl: pdfUrl || ''
  };
  var rowIndex = findEntryRow(sheet, payload.id);
  if (rowIndex) {
    payload.createdAt = sheet.getRange(rowIndex, 2).getValue();
    payload.updatedAt = now;
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([headers.map(function(h) { return payload[h] || ''; })]);
  } else {
    sheet.appendRow(headers.map(function(h) { return payload[h] || ''; }));
  }
  return payload;
}

function findEntryRow(sheet, entryId) {
  if (!entryId) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === entryId) return i + 2;
  }
  return 0;
}

// =============================================================
// Construction du HTML pour le PDF — fidèle au modèle IDML
// =============================================================
function buildPosterHtml(data) {
  // Contenu
  var titleLines = [data.title, data.titleLine2].filter(function(v) { return v; });
  var titleHtml = titleLines.length
    ? '<span class="title-text">' + escapeHtml(titleLines.join('\n')).replace(/\n/g, '<br>') + '</span>'
    : '';

  var subtitleHtml = data.subtitle
    ? '<div class="subtitle">' + escapeHtml(data.subtitle).replace(/\n/g, '<br>') + '</div>' : '';

  var datesHtml = data.dates
    ? '<div class="dates"><span class="bullet">&#8226;</span><span class="date-text">' + escapeHtml(data.dates).replace(/\n/g, ' ') + '</span></div>'
    : '';

  var bodyLines = [];
  if (data.description) bodyLines.push(escapeHtml(data.description));
  if (data.infos) bodyLines.push(escapeHtml(data.infos));
  var bodyHtml = bodyLines.length
    ? '<div class="body-text">' + bodyLines.join('\n').replace(/\n/g, '<br>') + '</div>'
    : '';

  var siteUrlHtml = data.siteUrl
    ? '<div class="url-bar"><span class="url-text">' + escapeHtml(data.siteUrl) + '</span></div>' : '';

  // Public cible
  var publicLines = (data.publicCible || '').split(/\n/).map(function(l) {
    return escapeHtml(l.trim());
  }).filter(function(l) { return l.length > 0; });
  var publicHtml = publicLines.length > 0
    ? '<div class="public-triangle"></div><div class="public-text">' +
      '<span class="public-line main">' + publicLines[0] + '</span>' +
      (publicLines[1] ? '<span class="public-line secondary">' + publicLines.slice(1).join(' ') + '</span>' : '') +
      '</div>'
    : '';

  var topLogo = data.topLogo || '';

  // Polices
  var ff = '';
  if (data.fontRias) ff += '@font-face{font-family:"Rias";src:url("'+data.fontRias+'");font-weight:700;font-style:normal}';
  if (data.fontRobotoBold) ff += '@font-face{font-family:"Roboto";src:url("'+data.fontRobotoBold+'");font-weight:700;font-style:normal}';
  if (data.fontVagRounded) ff += '@font-face{font-family:"VAG Rounded Std";src:url("'+data.fontVagRounded+'");font-weight:700;font-style:normal}';

  var css =
    ff +
    '@page{size:297mm 420mm;margin:0}' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'html,body{width:297mm;height:420mm;-webkit-print-color-adjust:exact;print-color-adjust:exact}' +

    // === POSTER ===
    '.poster{width:297mm;height:420mm;position:relative;overflow:hidden;background:#fff}' +

    // === CADRE GRIS ===
    '.page-frame{' +
      'position:absolute;left:7.89mm;top:6.4mm;' +
      'width:281mm;height:402.08mm;' +
      'background:#e6e6e6;z-index:1' +
    '}' +

    // === LOGO ===
    '.top-logo{' +
      'position:absolute;left:39.54mm;top:11.88mm;' +
      'width:54mm;height:86.12mm;' +
      'object-fit:contain;z-index:3' +
    '}' +

    // === BANDEAU VIOLET ===
    '.content-band{' +
      'position:absolute;left:23.85mm;top:264.04mm;' +
      'width:235.03mm;height:83.06mm;' +
      'background:#6666ff;z-index:2;' +
      'clip-path:polygon(0 0,100% 21%,100% 100%,0 100%);' +
      '-webkit-clip-path:polygon(0 0,100% 21%,100% 100%,0 100%)' +
    '}' +

    // === TITRE ===
    '.title{' +
      'position:absolute;left:29.85mm;bottom:126.96mm;' +
      'max-width:190mm;' +
      'font-family:"Rias","Segoe Print","Comic Sans MS",cursive;' +
      'font-size:47.69pt;font-weight:700;' +
      'line-height:1.0;color:#fff;text-transform:uppercase;' +
      'z-index:4' +
    '}' +

    // === SOUS-TITRE ===
    '.subtitle{' +
      'position:absolute;left:30mm;bottom:118.88mm;' +
      'max-width:150mm;' +
      'font-family:"VAG Rounded Std","Arial Rounded MT Bold","Helvetica Neue",Arial,sans-serif;' +
      'font-size:24.26pt;font-weight:700;' +
      'line-height:1;color:#fff;z-index:4' +
    '}' +

    // === DATES ===
    '.dates{' +
      'position:absolute;left:154.07mm;bottom:118.88mm;' +
      'display:flex;align-items:baseline;gap:4.88mm;' +
      'color:#fff;z-index:4' +
    '}' +
    '.dates .bullet{font-size:16.14pt;font-weight:700;font-family:"VAG Rounded Std","Arial Rounded MT Bold","Helvetica Neue",Arial,sans-serif}' +
    '.dates .date-text{font-size:26.65pt;font-weight:700;font-family:"Roboto","Helvetica Neue",Arial,sans-serif;text-transform:none}' +

    // === TEXTE PRINCIPAL ===
    '.body-text{' +
      'position:absolute;left:29.85mm;bottom:111.24mm;' +
      'max-width:210mm;' +
      'font-family:"VAG Rounded Std","Arial Rounded MT Bold","Helvetica Neue",Arial,sans-serif;' +
      'font-size:18.86pt;font-weight:700;' +
      'line-height:1.15;color:#000;z-index:4' +
    '}' +

    // === TRIANGLE PUBLIC ===
    '.public-triangle{' +
      'position:absolute;left:224.47mm;bottom:123.01mm;' +
      'width:37.05mm;height:35.17mm;' +
      'background:#ffffff;z-index:3;' +
      'clip-path:polygon(100% 100%,0 62.9%,70.6% 0);' +
      '-webkit-clip-path:polygon(100% 100%,0 62.9%,70.6% 0)' +
    '}' +
    '.public-text{' +
      'position:absolute;left:233.34mm;bottom:138.4mm;' +
      'display:flex;flex-direction:column;gap:1.2mm;' +
      'font-family:"Roboto","Helvetica Neue",Arial,sans-serif;' +
      'font-weight:700;color:#6666ff;z-index:4' +
    '}' +
    '.public-text .public-line.main{font-size:15.53pt}' +
    '.public-text .public-line.secondary{font-size:10.35pt}' +

    // === TRIANGLE DÉCORATIF ===
    '.small-triangle{' +
      'position:absolute;left:50.18mm;bottom:52.81mm;' +
      'width:8.19mm;height:8.54mm;' +
      'background:#6666ff;z-index:3;' +
      'clip-path:polygon(100% 100%,0 49.7%,86.1% 0);' +
      '-webkit-clip-path:polygon(100% 100%,0 49.7%,86.1% 0)' +
    '}' +

    // === BANDE URL ===
    '.url-bar{' +
      'position:absolute;left:23.85mm;bottom:59.83mm;' +
      'width:119.46mm;height:13.61mm;' +
      'background:#000;z-index:3;' +
      'display:flex;align-items:center;' +
      'padding-left:9.8mm' +
    '}' +
    '.url-text{' +
      'font-family:"Roboto","Helvetica Neue",Arial,sans-serif;' +
      'font-size:23.25pt;font-weight:700;' +
      'color:#6666ff;line-height:1' +
    '}';

  var html =
    '<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><style>' + css + '</style></head><body>' +
    '<div class="poster">' +

    '<div class="page-frame"></div>' +
    (topLogo ? '<img class="top-logo" src="' + topLogo + '" />' : '') +
    '<div class="content-band"></div>' +
    (titleHtml ? '<div class="title">' + titleHtml + '</div>' : '') +
    subtitleHtml +
    datesHtml +
    bodyHtml +
    publicHtml +
    siteUrlHtml +
    '<div class="small-triangle"></div>' +

    '</div></body></html>';

  return html;
}

function escapeHtml(text) {
  if (!text) return '';
  return text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function formatDateForFilename(input) {
  var trimmed = (input || '').trim();
  if (!trimmed) return '00-00-00';
  var monthMap = {
    'janvier':'01','février':'02','mars':'03','avril':'04',
    'mai':'05','juin':'06','juillet':'07','août':'08',
    'septembre':'09','octobre':'10','novembre':'11','décembre':'12'
  };
  var foundMonths = [];
  for (var name in monthMap) {
    if (trimmed.toLowerCase().indexOf(name) !== -1) foundMonths.push(monthMap[name]);
  }
  if (foundMonths.length > 0) { foundMonths.sort(); return foundMonths.join('-'); }
  var match = trimmed.match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);
  if (match) {
    var year = parseInt(match[3],10);
    if (year < 100) year += 2000;
    return String(year).slice(-2)+'-'+('0'+match[2]).slice(-2)+'-'+('0'+match[1]).slice(-2);
  }
  return '00-00-00';
}

function sanitizeFilename(text) {
  return text.replace(/[\/\\:*?"<>|]/g,'-').replace(/\s+/g,' ').trim().substring(0,80);
}
