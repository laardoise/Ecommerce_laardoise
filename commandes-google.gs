/**
 * L.A Ardoise – service de commandes (Google Apps Script)
 *
 * MISE EN PLACE (5 minutes, gratuit) :
 * 1. Allez sur https://sheets.new : créez une feuille Google nommée « Commandes L.A Ardoise ».
 * 2. Menu Extensions > Apps Script. Effacez le contenu, collez TOUT ce fichier.
 * 3. Remplacez CODE_SECRET ci-dessous par un code long que vous inventez (lettres et chiffres).
 * 4. Cliquez Déployer > Nouveau déploiement > type « Application Web ».
 *      - Exécuter en tant que : Moi
 *      - Qui a accès : Tout le monde
 *    Autorisez l'accès quand Google le demande, puis copiez l'URL qui finit par /exec.
 * 5. Dans index.html, collez cette URL dans ORDER_ENDPOINT (début du dernier script).
 * 6. Dans admin.html > Réglages > « Commandes en direct » : collez la même URL et votre code secret.
 *
 * Si vous modifiez ce script plus tard : Déployer > Gérer les déploiements > Modifier > Nouvelle version.
 */
var CODE_SECRET = 'CHANGEZ-MOI-PAR-UN-CODE-LONG';

function feuille_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Commandes');
  if (!sh) {
    sh = ss.insertSheet('Commandes');
    sh.appendRow(['id', 'date', 'total', 'lignes', 'client', 'statut', 'langue']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function reponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Le site envoie ici chaque nouvelle commande. Le dashboard envoie ici les changements de statut.
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var sh = feuille_();
    var ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();

    if (d.action === 'update') {
      if (d.token !== CODE_SECRET) return reponse_({ ok: false, error: 'token' });
      if (['new', 'paid', 'cancelled'].indexOf(d.status) < 0) return reponse_({ ok: false });
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(d.id)) { sh.getRange(i + 2, 6).setValue(d.status); return reponse_({ ok: true }); }
      }
      return reponse_({ ok: false, error: 'introuvable' });
    }

    // Nouvelle commande (publique). Contrôles de base contre les abus.
    if (d.hp) return reponse_({ ok: true });
    var total = Number(d.total);
    if (!(total > 0) || total > 5000000) return reponse_({ ok: false });
    var id = String(d.id || '').replace(/[^a-z0-9]/gi, '').slice(0, 24) || Utilities.getUuid().slice(0, 8);
    for (var j = 0; j < ids.length; j++) { if (String(ids[j][0]) === id) return reponse_({ ok: true, doublon: true }); }
    var lignes = (d.lines || []).slice(0, 30).map(function (s) { return String(s).slice(0, 140); });
    sh.appendRow([id, new Date().toISOString(), Math.round(total), JSON.stringify(lignes),
      String(d.who || '').slice(0, 80), 'new', String(d.lang || 'fr').slice(0, 2)]);
    return reponse_({ ok: true });
  } catch (err) {
    return reponse_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Le dashboard lit ici la liste des commandes (protégé par le code secret).
function doGet(e) {
  if (!e || !e.parameter || e.parameter.token !== CODE_SECRET) return reponse_({ ok: false, error: 'token' });
  var sh = feuille_();
  var n = sh.getLastRow() - 1;
  var orders = [];
  if (n > 0) {
    var rows = sh.getRange(2, 1, n, 7).getValues().slice(-300);
    orders = rows.map(function (r) {
      var lignes = []; try { lignes = JSON.parse(r[3]); } catch (x) {}
      return { id: String(r[0]), date: String(r[1]), total: Number(r[2]), lines: lignes, who: String(r[4]), status: String(r[5] || 'new'), lang: String(r[6]) };
    });
  }
  return reponse_({ ok: true, orders: orders });
}
