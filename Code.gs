// Journée Carrière UIST / 2ITS — backend Google Apps Script
// Mot de passe admin : Paramètres du projet > Propriétés du script > ADMIN_PASSWORD
// Mot de passe admin par défaut (utilisé si ADMIN_PASSWORD n'est pas défini dans les propriétés du script)
const DEFAULT_PASSWORD = 'UIST-2ITS-2026';
const VERSION = 'v3-mot-de-passe-par-defaut';
const SHEET = 'Inscriptions';
const HEAD = ['Numéro','Date','Nom','Prénoms','Sexe','Téléphone','E-mail','Établissement','Niveau','Filière','Domaine'];
const KEYS = ['num','date','nom','pre','sexe','tel','mail','etab','niv','fil','dom'];
const EDITABLE = ['nom','pre','sexe','tel','mail','etab','niv','fil','dom'];

function sh_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.appendRow(HEAD);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#0a8fe0').setFontColor('#ffffff');
    s.getRange('A:B').setNumberFormat('@');
    s.getRange('F:F').setNumberFormat('@');
  }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function clean_(v, max) { return String(v == null ? '' : v).trim().slice(0, max || 120); }
function phone_(v) { return String(v || '').replace(/[\s.\-()]/g, ''); }

function doGet() { return out_({ ok: true, service: 'Journée Carrière UIST / 2ITS', version: VERSION }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'register') return out_(register_(d.data || {}));

    const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || DEFAULT_PASSWORD;
    if (String(d.password || '').trim() !== String(pw).trim()) return out_({ ok: false, error: 'Mot de passe incorrect.' });

    if (d.action === 'login') return out_({ ok: true });
    if (d.action === 'list') return out_({ ok: true, rows: list_() });
    if (d.action === 'delete') return out_(delete_(d.num));
    if (d.action === 'update') return out_(update_(d.num, d.data || {}));
    return out_({ ok: false, error: 'Action inconnue.' });
  } catch (err) {
    return out_({ ok: false, error: 'Erreur serveur : ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

function register_(x) {
  if (x.website) return { ok: true, num: 'JC-26-0000' }; // anti-spam (champ piège)
  const r = {
    nom: clean_(x.nom).toUpperCase(), pre: clean_(x.pre), sexe: clean_(x.sexe, 10),
    tel: phone_(x.tel), mail: clean_(x.mail), etab: clean_(x.etab), niv: clean_(x.niv, 40),
    fil: clean_(x.fil, 40), dom: clean_(x.dom, 60)
  };
  if (!r.nom || !r.pre || !r.tel || !r.etab || !r.niv || !r.dom || !['Masculin','Féminin'].includes(r.sexe))
    return { ok: false, error: 'Champs obligatoires manquants.' };
  if (!/^\+?\d{8,15}$/.test(r.tel)) return { ok: false, error: 'Numéro de téléphone invalide.' };
  if (r.mail && !/^\S+@\S+\.\S+$/.test(r.mail)) return { ok: false, error: 'Adresse e-mail invalide.' };
  if (x.consent !== true) return { ok: false, error: 'Le consentement est obligatoire.' };

  const s = sh_();
  const n = s.getLastRow();
  if (n > 1) {
    const tels = s.getRange(2, 6, n - 1, 1).getValues().map(function (v) { return phone_(v[0]); });
    if (tels.indexOf(r.tel) !== -1) return { ok: false, error: 'Ce numéro de téléphone est déjà inscrit.' };
  }
  const props = PropertiesService.getScriptProperties();
  const c = Number(props.getProperty('COUNTER') || 0) + 1;
  props.setProperty('COUNTER', String(c));
  r.num = 'JC-26-' + ('0000' + c).slice(-4);
  r.date = new Date().toISOString();
  s.appendRow(KEYS.map(function (k) { return r[k]; }));
  return { ok: true, num: r.num, pre: r.pre };
}

function list_() {
  const v = sh_().getDataRange().getValues();
  return v.slice(1).map(function (row) {
    const o = {};
    KEYS.forEach(function (k, i) { o[k] = String(row[i] == null ? '' : row[i]); });
    return o;
  });
}
function findRow_(num) {
  const s = sh_(), n = s.getLastRow();
  if (n < 2) return { s: s, row: -1 };
  const nums = s.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < nums.length; i++) if (String(nums[i][0]) === String(num)) return { s: s, row: i + 2 };
  return { s: s, row: -1 };
}
function delete_(num) {
  const f = findRow_(num);
  if (f.row < 0) return { ok: false, error: 'Inscription introuvable.' };
  f.s.deleteRow(f.row);
  return { ok: true };
}
function update_(num, data) {
  const f = findRow_(num);
  if (f.row < 0) return { ok: false, error: 'Inscription introuvable.' };
  EDITABLE.forEach(function (k) {
    if (data[k] === undefined) return;
    let val = clean_(data[k]);
    if (k === 'nom') val = val.toUpperCase();
    if (k === 'tel') val = phone_(val);
    f.s.getRange(f.row, KEYS.indexOf(k) + 1).setValue(val);
  });
  return { ok: true };
}