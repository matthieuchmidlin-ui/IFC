// ══════════════════════════════════════════════════
// INFOCOM PRO — Google Apps Script Sync v2
// Compatible CORS / no-cors depuis Safari/iPhone
// ══════════════════════════════════════════════════

var SECRET_TOKEN = 'infocom2024matthieu'; // ⚠️ Change ce token

var SHEET_NAME = 'Prospects';
var CALLS_SHEET = 'Appels';
var RDV_SHEET = 'RDV';

function doPost(e) {
  // Vérification token
  var token = e.parameter.token || '';
  if (token !== SECRET_TOKEN) {
    return buildResponse({ok: false, error: 'Non autorise'});
  }

  try {
    // Accepte text/plain ET application/json (pour mode no-cors)
    var raw = e.postData.contents;
    var data = JSON.parse(raw);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.action === 'sync') {
      syncProspects(ss, data.companies || []);
      syncCalls(ss, data.calls || []);
      syncRdv(ss, data.rdvs || []);
      return buildResponse({
        ok: true,
        synced: (data.companies||[]).length,
        timestamp: new Date().toISOString()
      });
    }

  } catch(err) {
    return buildResponse({ok: false, error: err.message});
  }
}

function doGet(e) {
  var token = e.parameter.token || '';
  if (token !== SECRET_TOKEN) {
    return buildResponse({ok: false, error: 'Non autorise'});
  }
  return buildResponse({ok: true, status: 'Infocom Pro Sync actif'});
}

function buildResponse(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function syncProspects(ss, companies) {
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.clearContents();

  var headers = ['Entreprise','Secteur','Gérant','Téléphone','Email','Adresse',
                 'Statut','Température','Dossier','Dernière relance','Ajouté le'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
       .setBackground('#0f2044').setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(11);

  if (!companies.length) return;

  var rows = companies.map(function(c) {
    var gerant = c.gerant && c.gerant.nom ? c.gerant.nom : (c.contact || '');
    var temp = c.temp === 'hot' ? 'Chaud' : c.temp === 'warm' ? 'Tiède' : 'Froid';
    return [
      c.name || '', c.sector || '', gerant,
      c.phone || '', c.email || '', c.addr || '',
      c.lastStatus || '', temp, c.folderName || '',
      c.lastFollowup || '',
      c.addedAt ? new Date(c.addedAt).toLocaleDateString('fr-FR') : ''
    ];
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  for (var i = 0; i < rows.length; i++) {
    if (i % 2 === 0)
      sheet.getRange(i+2, 1, 1, headers.length).setBackground('#f8f9fa');
  }

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
      c.company || '', c.contact || '', c.phone || '',
      c.status || '', c.notes || '', c.followup || ''
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
      r.date || '', r.time || '', r.company || '',
      r.contact || '', r.addr || '', r.notes || '', r.status || ''
    ];
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}
