// Hospital Veterinario — Apps Script API
// Sheets ID: 11Y9nusH47nWcW1ONWl-N0OKP6f0Gsj9NIGYJOhNnsKM
// Deploy: Ejecutar como "Yo", Acceso "Cualquier persona"

var SPREADSHEET_ID = '11Y9nusH47nWcW1ONWl-N0OKP6f0Gsj9NIGYJOhNnsKM';

function doGet(e) {
  var params = {};
  try { params = JSON.parse(decodeURIComponent(e.parameter.d || '{}')); } catch(ex) {}
  return handle(params);
}

function handle(p) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    var a = p.action || '';
    var result;
    if      (a === 'getAll')       result = getAll();
    else if (a === 'savePatient')  result = savePatient(p.patient);
    else if (a === 'saveMed')      result = saveMed(p.med);
    else if (a === 'suspMed')      result = suspMed(p.id);
    else if (a === 'recordDose')   result = recordDose(p.patientId, p.medId, p.ts);
    else if (a === 'saveEvent')    result = saveEvent(p.ev);
    else if (a === 'saveMon')      result = saveMon(p.mon);
    else if (a === 'saveLabs')     result = saveLabs(p.patientId, p.labDate, p.labs);
    else if (a === 'updateStatus') result = updateStatus(p.id, p.status, p.discharge);
    else result = { error: 'Unknown action: ' + a };
    out.setContent(JSON.stringify({ ok: true, data: result }));
  } catch (err) {
    out.setContent(JSON.stringify({ ok: false, error: err.toString() }));
  }
  return out;
}

// ── Helpers ──────────────────────────────────────────────

function ss() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

function rows(name) {
  var sh = ss().getSheetByName(name);
  if (!sh) return [];
  var d = sh.getDataRange().getValues();
  if (d.length < 2) return [];
  var h = d[0];
  return d.slice(1).map(function(r) {
    var o = {};
    h.forEach(function(k, i) { o[k] = (r[i] === null || r[i] === undefined) ? '' : r[i]; });
    return o;
  });
}

function nextId(name) {
  var r = rows(name);
  if (!r.length) return 1;
  var max = 0;
  r.forEach(function(row) { var n = parseInt(row.id) || 0; if (n > max) max = n; });
  return max + 1;
}

function upsert(name, data, key) {
  var sh = ss().getSheetByName(name);
  if (!sh) throw new Error('Hoja no encontrada: ' + name);
  var d = sh.getDataRange().getValues();
  var h = d[0];
  var row = h.map(function(k) { return data[k] !== undefined ? data[k] : ''; });
  var keyIdx = h.indexOf(key);
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][keyIdx]) === String(data[key])) {
      sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return data[key];
    }
  }
  sh.appendRow(row);
  return data[key];
}

// ── getAll ────────────────────────────────────────────────

var LAB_KEYS = ['HTO','HG','LEUCO','PLTS','NEU','LIN','MON','ALT','CREAT','BUN','ALP','CA','P','ALB','BIL','CHOL','NH3','PT','PTT','GGT'];

function getAll() {
  var pats  = rows('Pacientes');
  var meds  = rows('Medicamentos');
  var doses = rows('Dosis');
  var mons  = rows('Monitoreo');
  var evs   = rows('Eventos');
  var labs  = rows('Labs');

  return pats.map(function(p) {
    var pid = String(p.id);

    var patMeds = meds
      .filter(function(m) { return String(m.patientId) === pid; })
      .map(function(m) {
        return {
          id:    parseInt(m.id)        || 0,
          name:  String(m.name         || ''),
          dose:  String(m.dose         || ''),
          ml:    parseFloat(m.ml)      || 0,
          via:   String(m.via          || 'IV'),
          freq:  parseInt(m.freq)      || 0,
          price: parseInt(m.price)     || 0,
          susp:  m.susp === true || String(m.susp).toUpperCase() === 'TRUE',
          vsId:  String(m.vsId         || '')
        };
      });

    var admins = {};
    doses
      .filter(function(d) { return String(d.patientId) === pid; })
      .forEach(function(d) {
        var mid = String(d.medId);
        if (!admins[mid]) admins[mid] = [];
        admins[mid].push(String(d.timestamp));
      });

    var mon = mons
      .filter(function(m) { return String(m.patientId) === pid; })
      .map(function(m) {
        var data = {};
        try { data = JSON.parse(String(m.data || '{}')); } catch(e) {}
        return { date: m.date, hour: m.hour, data: data };
      });

    var evArr = evs
      .filter(function(e) { return String(e.patientId) === pid; })
      .map(function(e) {
        return { time: String(e.time), type: String(e.type), text: String(e.text), by: String(e.by) };
      });

    var patLabs = labs.filter(function(l) { return String(l.patientId) === pid; });
    var labDates = [];
    patLabs.forEach(function(l) {
      var dt = String(l.labDate);
      if (labDates.indexOf(dt) === -1) labDates.push(dt);
    });
    labDates.sort();

    var labsObj = {};
    LAB_KEYS.forEach(function(key) {
      labsObj[key] = labDates.map(function(date) {
        for (var i = 0; i < patLabs.length; i++) {
          if (String(patLabs[i].labDate) === date && String(patLabs[i].paramKey) === key)
            return String(patLabs[i].value);
        }
        return '';
      });
    });

    return {
      id:            parseInt(p.id)            || 0,
      name:          String(p.name             || ''),
      owner:         String(p.owner            || ''),
      doc:           String(p.doc              || ''),
      phone:         String(p.phone            || ''),
      species:       String(p.species          || ''),
      breed:         String(p.breed            || ''),
      sex:           String(p.sex              || ''),
      weight:        parseFloat(p.weight)      || 0,
      age:           String(p.age              || ''),
      status:        String(p.status           || 'hospitalizado'),
      vet:           String(p.vet              || ''),
      vsId:          String(p.vsId             || ''),
      admitDate:     String(p.admitDate        || ''),
      dischargeDate: String(p.dischargeDate    || ''),
      reason:        String(p.reason           || ''),
      diagnosis:     String(p.diagnosis        || ''),
      emoji:         String(p.emoji            || '🐾'),
      dayRate:       parseInt(p.dayRate)       || 0,
      medications:   patMeds,
      admins:        admins,
      monitoring:    mon,
      events:        evArr,
      labDates:      labDates,
      labs:          labsObj
    };
  });
}

// ── Mutations ─────────────────────────────────────────────

function savePatient(pat) {
  if (!pat.id) pat.id = nextId('Pacientes');
  upsert('Pacientes', pat, 'id');
  return parseInt(pat.id);
}

function saveMed(med) {
  if (!med.id) med.id = nextId('Medicamentos');
  upsert('Medicamentos', med, 'id');
  return parseInt(med.id);
}

function suspMed(id) {
  var sh = ss().getSheetByName('Medicamentos');
  var d = sh.getDataRange().getValues();
  var h = d[0];
  var idIdx   = h.indexOf('id');
  var suspIdx = h.indexOf('susp');
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][idIdx]) === String(id)) {
      sh.getRange(i + 1, suspIdx + 1).setValue('TRUE');
      return true;
    }
  }
  return false;
}

function recordDose(patientId, medId, ts) {
  var id = nextId('Dosis');
  ss().getSheetByName('Dosis').appendRow([id, patientId, medId, ts]);
  return id;
}

function saveEvent(ev) {
  if (!ev.id) ev.id = nextId('Eventos');
  upsert('Eventos', ev, 'id');
  return parseInt(ev.id);
}

function saveMon(mon) {
  if (!mon.id) mon.id = nextId('Monitoreo');
  var copy = {};
  Object.keys(mon).forEach(function(k) { copy[k] = mon[k]; });
  if (typeof copy.data === 'object') copy.data = JSON.stringify(copy.data);
  upsert('Monitoreo', copy, 'id');
  return parseInt(mon.id);
}

function saveLabs(patientId, labDate, labs) {
  var sh = ss().getSheetByName('Labs');
  var d = sh.getDataRange().getValues();
  var h = d[0];
  var pidIdx  = h.indexOf('patientId');
  var dateIdx = h.indexOf('labDate');
  var toDel = [];
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][pidIdx]) === String(patientId) && String(d[i][dateIdx]) === String(labDate))
      toDel.push(i + 1);
  }
  for (var j = toDel.length - 1; j >= 0; j--) sh.deleteRow(toDel[j]);
  var id = nextId('Labs');
  Object.keys(labs).forEach(function(key) {
    if (String(labs[key]).trim() !== '') sh.appendRow([id++, patientId, labDate, key, labs[key]]);
  });
  return true;
}

function updateStatus(id, status, discharge) {
  var sh = ss().getSheetByName('Pacientes');
  var d = sh.getDataRange().getValues();
  var h = d[0];
  var idIdx   = h.indexOf('id');
  var stIdx   = h.indexOf('status');
  var discIdx = h.indexOf('dischargeDate');
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][idIdx]) === String(id)) {
      sh.getRange(i + 1, stIdx + 1).setValue(status);
      if (discharge && discIdx >= 0) sh.getRange(i + 1, discIdx + 1).setValue(discharge);
      return true;
    }
  }
  return false;
}
