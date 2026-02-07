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
  var titleLine2Html = data.titleLine2
    ? '<span class="title-line">' + escapeHtml(data.titleLine2) + '</span>' : '';
  var titleHtml = '<span class="title-line">' + escapeHtml(data.title) + '</span>' + titleLine2Html;

  var subtitleHtml = data.subtitle
    ? '<p class="subtitle">' + escapeHtml(data.subtitle).replace(/\n/g, '<br>') + '</p>' : '';
  var descHtml = data.description
    ? '<p class="description">' + escapeHtml(data.description).replace(/\n/g, '<br>') + '</p>' : '';
  var datesHtml = data.dates
    ? '<p class="dates">' + escapeHtml(data.dates).replace(/\n/g, '<br>') + '</p>' : '';
  var infosHtml = data.infos
    ? '<p class="infos">' + escapeHtml(data.infos).replace(/\n/g, '<br>') + '</p>' : '';
  var siteUrlHtml = '';

  // Public cible
  var publicLines = (data.publicCible || '').split(/\n/).map(function(l) {
    return escapeHtml(l.trim());
  }).filter(function(l) { return l.length > 0; });
  var publicHtml = '';
  if (publicLines.length > 0) {
    publicHtml = '<div class="triangle-wrap"><div class="triangle-bg"></div><div class="triangle-text">';
    publicHtml += '<span class="public-line-primary">' + publicLines[0] + '</span>';
    if (publicLines.length > 1) {
      publicHtml += '<span class="public-line-secondary">' + publicLines.slice(1).join(' ') + '</span>';
    }
    publicHtml += '</div></div>';
  }

  var mainImage = data.mainImage || '';
  var topLogo = data.topLogo || '';
  var bottomLogo = data.bottomLogo || '';
  var imageBgCss = mainImage ? 'background-image: url(\'' + mainImage + '\');' : '';

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

    // === PHOTO DE FOND (cadre IDML) ===
    '.image-bg{' +
      'position:absolute;left:3.25mm;top:5mm;' +
      'width:290.75mm;height:399.75mm;' +
      'background-size:cover;background-position:center center;' +
      'z-index:1' +
    '}' +
    '.image-overlay{' +
      'position:absolute;left:7.89mm;top:6.39mm;' +
      'width:280.99mm;height:402.08mm;' +
      'background:rgba(0,0,0,0.10);' +
      'z-index:2' +
    '}' +

    // === LOGO HAUT ===
    '.top-logo{' +
      'position:absolute;left:39.54mm;top:11.63mm;' +
      'width:54mm;height:86.12mm;' +
      'object-fit:contain;' +
      'z-index:6' +
    '}' +

    // === TRIANGLE PUBLIC ===
    '.triangle-wrap{' +
      'position:absolute;left:166.57mm;top:275.67mm;' +
      'width:37.05mm;height:35.17mm;' +
      'z-index:12' +
    '}' +
    '.triangle-bg{' +
      'position:absolute;inset:0;' +
      'background:#ffffff;' +
      'clip-path:polygon(100% 0, 0 100%, 100% 100%);' +
      '-webkit-clip-path:polygon(100% 0, 0 100%, 100% 100%)' +
    '}' +
    '.triangle-text{' +
      'position:absolute;inset:0;' +
      'display:flex;flex-direction:column;justify-content:center;align-items:center;' +
      'text-align:center;' +
      'font-family:"Roboto","Helvetica Neue",Arial,sans-serif;' +
      'color:#6666ff;' +
      'line-height:1;' +
      'letter-spacing:0.5pt;' +
      'z-index:13' +
    '}' +
    '.triangle-text .public-line-primary{font-size:15.53pt;font-weight:700;line-height:14.5pt;}' +
    '.triangle-text .public-line-secondary{font-size:10.35pt;font-weight:700;line-height:14.5pt;}' +

    // === BANDEAU VIOLET ===
    '.content-band{' +
      'position:absolute;' +
      'left:37.35mm;top:261.04mm;' +
      'width:235.03mm;height:83.07mm;' +
      'background:#6666ff;' +
      'z-index:10' +
    '}' +
    '.content-inner{' +
      'position:absolute;' +
      'left:6mm;top:16.38mm;right:1.84mm;bottom:11.84mm;' +
      'z-index:11' +
    '}' +

    // === TITRE (Rias) — blanc, uppercase ===
    '.title{' +
      'font-family:"Rias","Segoe Print","Comic Sans MS",cursive;' +
      'font-size:47.7pt;font-weight:700;' +
      'color:#fff;' +
      'line-height:31.24pt;' +
      'margin:0 0 2mm 0;' +
      'text-transform:uppercase;' +
      'letter-spacing:0.2mm' +
    '}' +
    '.title-line{display:block}' +

    // === SOUS-TITRE (VAG Rounded) — noir ===
    '.subtitle{' +
      'font-family:"VAG Rounded Std","Arial Rounded MT Bold","Helvetica Neue",Arial,sans-serif;' +
      'font-size:18.86pt;font-weight:700;' +
      'color:#000;' +
      'line-height:18.83pt;' +
      'margin:0 0 2mm 0' +
    '}' +

    // === DATES (Roboto Bold) — blanc, uppercase ===
    '.dates{' +
      'font-family:"Roboto","Helvetica Neue",Arial,sans-serif;' +
      'font-size:26.65pt;font-weight:700;' +
      'color:#fff;' +
      'line-height:14.06pt;' +
      'margin:0 0 1mm 0;' +
      'text-transform:uppercase' +
    '}' +

    // === INFOS PRATIQUES (VAG Rounded) — noir ===
    '.infos{' +
      'font-family:"VAG Rounded Std","Arial Rounded MT Bold","Helvetica Neue",Arial,sans-serif;' +
      'font-size:18.86pt;font-weight:700;' +
      'color:#000;' +
      'line-height:18.83pt;' +
      'margin:0 0 2mm 0' +
    '}' +

    // === DESCRIPTION (Roboto) — blanc ===
    '.description{' +
      'font-family:"Roboto","Helvetica Neue",Arial,sans-serif;' +
      'font-size:22.05pt;font-weight:700;' +
      'color:#fff;' +
      'line-height:14.06pt;' +
      'margin:0 0 1mm 0' +
    '}' +

    // === URL (footer) ===
    '.footer-url{' +
      'position:absolute;left:47.14mm;top:347.61mm;' +
      'font-family:"Uni Neue","Roboto","Helvetica Neue",Arial,sans-serif;' +
      'font-size:23.49pt;font-weight:700;' +
      'color:#6666ff;' +
      'z-index:20' +
    '}' +

    // === LOGO BAS ===
    '.footer-logo{' +
      'position:absolute;left:205.5mm;top:366.82mm;' +
      'width:80.12mm;height:45.2mm;' +
      'object-fit:contain;' +
      'z-index:20' +
    '}';

  var html =
    '<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><style>' + css + '</style></head><body>' +
    '<div class="poster">' +

    // Photo
    '<div class="image-bg" style="' + imageBgCss + '"></div>' +
    '<div class="image-overlay"></div>' +

    // Logo haut
    (topLogo ? '<img class="top-logo" src="' + topLogo + '" />' : '') +

    // Triangle public
    publicHtml +

    // Bandeau violet
    '<div class="content-band">' +
      '<div class="content-inner">' +
      '<h1 class="title">' + titleHtml + '</h1>' +
      subtitleHtml +
      datesHtml +
      infosHtml +
      descHtml +
      siteUrlHtml +
      '</div>' +
    '</div>' +

    // URL + logo bas
    (data.siteUrl ? '<div class="footer-url">' + escapeHtml(data.siteUrl) + '</div>' : '') +
    (bottomLogo ? '<img class="footer-logo" src="' + bottomLogo + '" />' : '') +

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
