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
  var siteUrlHtml = data.siteUrl
    ? '<div class="site-url-box">' + escapeHtml(data.siteUrl) + '</div>' : '';

  // Public cible
  var publicLines = (data.publicCible || '').split(/\n/).map(function(l) {
    return escapeHtml(l.trim());
  }).filter(function(l) { return l.length > 0; });
  var publicHtml = publicLines.length > 0
    ? '<div class="triangle-wrap"><div class="triangle-bg"></div><div class="triangle-text">' + publicLines.join('<br>') + '</div></div>'
    : '';

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

    // === PHOTO DE FOND ===
    // Occupe les 2/3 supérieurs, fond couvrant
    '.image-bg{' +
      'position:absolute;top:0;left:0;' +
      'width:297mm;height:290mm;' +
      'background-size:cover;background-position:center center;' +
      'z-index:1' +
    '}' +

    // === TRIANGLE BLANC — coin haut-droit ===
    // Triangle rectangle : sommet haut-droit, base le long du bord droit
    // Hypoténuse va du coin haut-gauche du triangle vers le bas-droit
    '.triangle-wrap{' +
      'position:absolute;top:0;right:0;' +
      'width:140mm;height:120mm;' +
      'z-index:5' +
    '}' +
    '.triangle-bg{' +
      'position:absolute;top:0;right:0;' +
      'width:140mm;height:120mm;' +
      'background:#ffffff;' +
      'clip-path:polygon(100% 0, 25% 0, 100% 100%);' +
      '-webkit-clip-path:polygon(100% 0, 25% 0, 100% 100%)' +
    '}' +
    '.triangle-text{' +
      'position:absolute;' +
      'top:18mm;right:8mm;' +
      'text-align:right;' +
      'font-family:"Roboto","Helvetica Neue",Arial,sans-serif;' +
      'font-size:22pt;font-weight:700;font-style:italic;' +
      'color:#7c5cbf;' +
      'line-height:1.3;' +
      'z-index:6' +
    '}' +

    // === FOND VIOLET ZONE BASSE ===
    // Remplit l'espace entre la photo et le footer
    '.violet-bg{' +
      'position:absolute;' +
      'left:0;top:250mm;' +
      'width:297mm;bottom:30mm;' +
      'background:rgba(138,118,205,0.82);' +
      'z-index:2' +
    '}' +

    // === BANDEAU VIOLET SEMI-TRANSPARENT ===
    // Superposé sur la partie basse de la photo
    // Bord supérieur légèrement incliné, va jusqu'au footer
    '.content-band{' +
      'position:absolute;' +
      'left:0;top:215mm;' +
      'width:235mm;bottom:30mm;' +
      'background:rgba(138,118,205,0.82);' +
      'clip-path:polygon(0 5%,70% 0,100% 3%,100% 100%,0 100%);' +
      '-webkit-clip-path:polygon(0 5%,70% 0,100% 3%,100% 100%,0 100%);' +
      'z-index:10;' +
      'padding:20mm 14mm 8mm 14mm' +
    '}' +

    // === TITRE (Rias) — blanc, uppercase ===
    '.title{' +
      'font-family:"Rias","Segoe Print","Comic Sans MS",cursive;' +
      'font-size:36pt;font-weight:700;' +
      'color:#fff;' +
      'line-height:1.0;' +
      'margin:0 0 3mm 0;' +
      'text-transform:uppercase;' +
      'letter-spacing:0.3mm' +
    '}' +
    '.title-line{display:block}' +

    // === SOUS-TITRE (VAG Rounded) — brun foncé ===
    '.subtitle{' +
      'font-family:"VAG Rounded Std","Arial Rounded MT Bold","Helvetica Neue",Arial,sans-serif;' +
      'font-size:10.5pt;font-weight:700;' +
      'color:#1a1512;' +
      'line-height:1.25;' +
      'margin:0 0 3mm 0' +
    '}' +

    // === DATES (Roboto Bold) — blanc, uppercase ===
    '.dates{' +
      'font-family:"Roboto","Helvetica Neue",Arial,sans-serif;' +
      'font-size:14pt;font-weight:700;' +
      'color:#fff;' +
      'line-height:1.15;' +
      'margin:0 0 2mm 0;' +
      'text-transform:uppercase' +
    '}' +

    // === INFOS PRATIQUES (VAG Rounded) — brun foncé ===
    '.infos{' +
      'font-family:"VAG Rounded Std","Arial Rounded MT Bold","Helvetica Neue",Arial,sans-serif;' +
      'font-size:9pt;font-weight:700;' +
      'color:#1a1512;' +
      'line-height:1.3;' +
      'margin:0 0 3mm 0' +
    '}' +

    // === DESCRIPTION (VAG Rounded) — brun foncé ===
    '.description{' +
      'font-family:"VAG Rounded Std","Arial Rounded MT Bold","Helvetica Neue",Arial,sans-serif;' +
      'font-size:9pt;font-weight:700;' +
      'color:#1a1512;' +
      'line-height:1.3;' +
      'margin:0 0 2mm 0' +
    '}' +

    // === ENCADRÉ URL ===
    '.site-url-box{' +
      'display:inline-block;' +
      'font-family:"VAG Rounded Std","Arial Rounded MT Bold","Helvetica Neue",Arial,sans-serif;' +
      'border:1.5pt solid #1a1512;' +
      'padding:1.5mm 5mm;' +
      'font-size:10.5pt;font-weight:700;' +
      'color:#1a1512;' +
      'margin-top:2mm' +
    '}' +

    // === FOOTER ===
    '.footer{' +
      'position:absolute;bottom:0;left:0;' +
      'width:297mm;height:30mm;' +
      'background:#f0edf3;' +
      'z-index:20;' +
      'display:flex;justify-content:space-between;align-items:center;' +
      'padding:0 12mm' +
    '}' +
    '.footer-url{' +
      'font-size:8pt;color:#333;' +
      'font-family:"Helvetica Neue",Helvetica,Arial,sans-serif' +
    '}' +
    '.footer-url .arr{color:#7c5cbf;font-size:7pt;margin-right:1mm}' +
    '.footer-logo{height:20mm}';

  var html =
    '<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><style>' + css + '</style></head><body>' +
    '<div class="poster">' +

    // Photo
    '<div class="image-bg" style="' + imageBgCss + '"></div>' +

    // Fond violet zone basse (remplissage entre image et footer)
    '<div class="violet-bg"></div>' +

    // Triangle public
    publicHtml +

    // Bandeau violet
    '<div class="content-band">' +
      '<h1 class="title">' + titleHtml + '</h1>' +
      subtitleHtml +
      datesHtml +
      infosHtml +
      descHtml +
      siteUrlHtml +
    '</div>' +

    // Footer
    '<div class="footer">' +
      '<div class="footer-url"><span class="arr">&#9654;</span>golfedumorbihan-vannesagglomeration.bzh</div>' +
      (bottomLogo ? '<img class="footer-logo" src="' + bottomLogo + '" />' : '') +
    '</div>' +

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
