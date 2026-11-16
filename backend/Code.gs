// Backend Transfer Surge #1105 — Google Apps Script legato a un Google Sheet.
// Usato da transfer-surge-1105 (candidature) e dash-transfer-leader (gestione).
// Deploy: Estensioni → Apps Script → incolla → Esegui il deployment → App web,
// "Esegui come: me", "Chi ha accesso: chiunque". Copia l'URL /exec nelle due app.

const HEAD = ['ts','lang','nickname','power','server','seat','intent','alliance','allianceName','bio','status','notes'];
const NAMES = {NgC:'Night Commando', DJN:'THE RISE', LdKy:'GottisUnitedKnights', VNEM:'True Unity', FmLy:'Family01'};

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(head); }
  return sh;
}
const cand_ = () => sheet_('Candidature', HEAD);
const texts_ = () => sheet_('TestiAlleanze', ['tag','json']);
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function readTexts_() {
  const v = texts_().getDataRange().getValues().slice(1), o = {};
  v.forEach(r => { if (r[0]) try { o[r[0]] = JSON.parse(r[1]); } catch (e) {} });
  return o;
}

function doGet(e) {
  const p = e.parameter || {};
  try {
    if (p.action === 'getAll') {
      const rows = cand_().getDataRange().getValues().slice(1)
        .map(r => r.map(c => c instanceof Date ? c.toISOString() : c));
      return out_({success:true, rows, allianceTexts: readTexts_()});
    }
    if (p.action === 'translateTexts') {
      const src = p.sourceLang || 'it', translations = {};
      (p.targetLangs || 'fr,en,de').split(',').forEach(l => {
        const tr = s => s ? LanguageApp.translate(s, src, l) : '';
        translations[l] = {req: tr(p.sourceReq), desc: tr(p.sourceDesc)};
      });
      return out_({success:true, translations});
    }
    return out_({success:false, error:'azione sconosciuta'});
  } catch (err) { return out_({success:false, error:String(err)}); }
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'updateStatus') {
      const row = parseInt(d.rowIndex, 10), sh = cand_();
      if (!(row >= 2 && row <= sh.getLastRow())) return out_({success:false, error:'riga non valida'});
      sh.getRange(row, HEAD.indexOf('status') + 1).setValue(String(d.status || '').slice(0, 40));
      return out_({success:true});
    }
    if (d.action === 'saveAllianceTexts') {
      if (!NAMES[d.tag]) return out_({success:false, error:'alleanza sconosciuta'});
      const sh = texts_(), v = sh.getDataRange().getValues();
      const i = v.findIndex(r => r[0] === d.tag), json = JSON.stringify(d.texts || {});
      if (i > 0) sh.getRange(i + 1, 2).setValue(json); else sh.appendRow([d.tag, json]);
      return out_({success:true});
    }
    // candidatura dal player
    if (!d.nickname || !d.server || !NAMES[d.alliance] || !d.seat) return out_({success:false, error:'campi mancanti'});
    // apostrofo davanti a = + - @ : il foglio non lo esegue come formula
    const s = (x, n) => { const v = String(x == null ? '' : x).slice(0, n); return /^[=+\-@]/.test(v) ? "'" + v : v; };
    cand_().appendRow([new Date().toISOString(), s(d.lang,2), s(d.nickname,40), s(d.power,15), s(d.server,10),
      s(d.seat,15), s(d.intent,20), d.alliance, NAMES[d.alliance], s(d.bio,1000), 'In attesa', '']);
    return out_({success:true});
  } catch (err) { return out_({success:false, error:String(err)}); }
  finally { lock.releaseLock(); }
}
