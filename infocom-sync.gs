// ══════════════════════════════════════════════════
// INFOCOM PRO — Google Apps Script Sync
// Déployer comme Web App (accès : Tout le monde)
// ══════════════════════════════════════════════════

var SHEET_NAME = 'Prospects';
var CALLS_SHEET = 'Appels';
var RDV_SHEET = 'RDV';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (data.action === 'sync') {
      syncProspects(ss, data.companies || []);
      syncCalls(ss, data.calls || []);
      syncRdv(ss, data.rdvs || []);
      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        synced: (data.companies||[]).length,
        timestamp: new Date().toISOString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false, error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    ok: true, status: 'Infocom Pro Sync actif'
  })).setMimeType(ContentService.MimeType.JSON);
}

function syncProspects(ss, companies) {
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.clearContents();
  
  // En-têtes
  var headers = ['Entreprise','Secteur','Gérant','Téléphone','Email','Adresse',
                 'Statut','Température','Dossier','Dernière relance','Ajouté le'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Style en-têtes
  var hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setBackground('#0f2044').setFontColor('#ffffff')
        .setFontWeight('bold').setFontSize(11);
  
  if (!companies.length) return;
  
  var rows = companies.map(function(c) {
    var gerant = c.gerant && c.gerant.nom ? c.gerant.nom : (c.contact || '');
    var temp = c.temp === 'hot' ? '🔥 Chaud' : c.temp === 'warm' ? '🌤 Tiède' : '❄️ Froid';
    return [
      c.name || '',
      c.sector || '',
      gerant,
      c.phone || '',
      c.email || '',
      c.addr || '',
      c.lastStatus || '',
      temp,
      c.folderName || c.folderId || '',
      c.lastFollowup || '',
      c.addedAt ? new Date(c.addedAt).toLocaleDateString('fr-FR') : ''
    ];
  });
  
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  
  // Alternance couleurs lignes
  for (var i = 0; i < rows.length; i++) {
    if (i % 2 === 0) {
      sheet.getRange(i + 2, 1, 1, headers.length).setBackground('#f8f9fa');
    }
  }
  
  // Auto-resize colonnes
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

function syncCalls(ss, calls) {
  var sheet = ss.getSheetByName(CALLS_SHEET) || ss.insertSheet(CALLS_SHEET);
  sheet.clearContents();
  
  var headers = ['Date','Entreprise','Contact','Téléphone','Statut','Notes','Relance'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
       .setBackground('#1a3a7c').setFontColor('#ffffff').setFontWeight('bold');
  
  if (!calls.length) return;
  
  var rows = calls.map(function(c) {
    return [
      c.date ? new Date(c.date).toLocaleString('fr-FR') : '',
      c.company || '',
      c.contact || '',
      c.phone || '',
      c.status || '',
      c.notes || '',
      c.followup || ''
    ];
  });
  
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

function syncRdv(ss, rdvs) {
  var sheet = ss.getSheetByName(RDV_SHEET) || ss.insertSheet(RDV_SHEET);
  sheet.clearContents();
  
  var headers = ['Date RDV','Heure','Entreprise','Contact','Adresse','Notes','Statut'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
       .setBackground('#15803d').setFontColor('#ffffff').setFontWeight('bold');
  
  if (!rdvs.length) return;
  
  var rows = rdvs.map(function(r) {
    return [
      r.date || '',
      r.time || '',
      r.company || '',
      r.contact || '',
      r.addr || '',
      r.notes || '',
      r.status || ''
    ];
  });
  
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}
