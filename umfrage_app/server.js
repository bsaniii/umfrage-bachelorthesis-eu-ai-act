const express = require('express');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_FILE = path.join(__dirname, 'data', 'responses.json');

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

if (!fs.existsSync(path.join(__dirname, 'data'))) {
  fs.mkdirSync(path.join(__dirname, 'data'));
}
if (!fs.existsSync(DATA_FILE)) {
  fs.writeFileSync(DATA_FILE, JSON.stringify([]));
}

function loadData() {
  try { return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); }
  catch { return []; }
}

function saveData(data) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));
}

// Zwischenspeichern (jede Seite)
app.post('/api/save', (req, res) => {
  const { sessionId, step, answers, completed } = req.body;
  if (!sessionId) return res.status(400).json({ error: 'sessionId fehlt' });

  const data = loadData();
  const idx = data.findIndex(r => r.sessionId === sessionId);
  const entry = {
    sessionId,
    step: step || 0,
    answers: answers || {},
    completed: completed || false,
    lastUpdated: new Date().toISOString(),
    createdAt: idx === -1 ? new Date().toISOString() : data[idx].createdAt
  };

  if (idx === -1) data.push(entry);
  else data[idx] = entry;

  saveData(data);
  res.json({ ok: true });
});

// Admin: Excel-Export
app.get('/admin/export', (req, res) => {
  const token = req.query.token;
  if (token !== process.env.ADMIN_TOKEN && token !== 'admin123') {
    return res.status(401).send('Nicht autorisiert');
  }

  const data = loadData();

  const workbook = new ExcelJS.Workbook();

  // Sheet 1: Alle Antworten
  const ws = workbook.addWorksheet('Alle Antworten');

  const labelMap = {
    email: 'E-Mail (Gewinnspiel)',
    company_size: 'Unternehmensgröße',
    role: 'Funktion',
    sector: 'Branche',
    ai_use: 'KI-Einsatz',
    know_aiact: 'Kenntnisstand EU AI Act',
    know_riskclass: 'Risikoklassifizierung bekannt',
    compliance_status: 'Compliance-Status',
    bb_awareness: 'Black-Box-Bewusstsein',
    xai_use: 'Genutzte XAI-Methoden',
    main_barrier: 'Größte Hürde',
    support_needed: 'Gewünschte Unterstützung',
    feasibility: 'Technische Umsetzbarkeit',
    regulation_opinion: 'Bewertung EU AI Act',
    open_comment: 'Offener Kommentar'
  };

  const likertLabels = {
    likert_challenges: [
      'Intransparenz erschwert EU AI Act',
      'XAI-Methoden reichen für EU AI Act',
      'Black-Box mindert Kundenvertrauen',
      'Bereit Leistung für Interpretierbarkeit zu opfern'
    ],
    likert_barriers: [
      'Fehlende tech. Standards',
      'Unklare rechtliche Anforderungen',
      'Mangelndes internes Fachwissen',
      'Hoher Dokumentationsaufwand',
      'Fehlende Compliance-Tools',
      'Wirtschaftliche Kosten'
    ],
    likert_future: [
      'XAI wird Standard in 3 Jahren',
      'EU AI Act verbessert KI-Qualität',
      'Nicht-Compliance = Wettbewerbsnachteil'
    ]
  };

  // Header-Zeile aufbauen
  const headers = ['Session-ID', 'Erstellt', 'Zuletzt aktualisiert', 'Schritt', 'Abgeschlossen'];
  Object.values(labelMap).forEach(l => headers.push(l));
  Object.entries(likertLabels).forEach(([, items]) => items.forEach(i => headers.push(i)));

  ws.addRow(headers);
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEDFE' } };

  data.forEach(r => {
    const a = r.answers || {};
    const row = [
      r.sessionId,
      r.createdAt ? new Date(r.createdAt).toLocaleString('de-DE') : '',
      r.lastUpdated ? new Date(r.lastUpdated).toLocaleString('de-DE') : '',
      r.step,
      r.completed ? 'Ja' : 'Nein'
    ];
    Object.keys(labelMap).forEach(k => {
      const v = a[k];
      row.push(Array.isArray(v) ? v.join('; ') : (v || ''));
    });
    Object.keys(likertLabels).forEach(k => {
      const items = likertLabels[k];
      const obj = a[k] || {};
      items.forEach(item => row.push(obj[item] || ''));
    });
    ws.addRow(row);
  });

  ws.columns.forEach(c => { c.width = 28; });

  // Sheet 2: Zusammenfassung
  const ws2 = workbook.addWorksheet('Zusammenfassung');
  const completed = data.filter(r => r.completed);
  const partial = data.filter(r => !r.completed);
  ws2.addRow(['Kennzahl', 'Wert']);
  ws2.getRow(1).font = { bold: true };
  ws2.addRow(['Gesamt Teilnahmen', data.length]);
  ws2.addRow(['Davon abgeschlossen', completed.length]);
  ws2.addRow(['Davon abgebrochen/unvollständig', partial.length]);
  ws2.addRow(['Export erstellt', new Date().toLocaleString('de-DE')]);
  ws2.columns = [{ width: 36 }, { width: 20 }];

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=umfrage_export_' + new Date().toISOString().slice(0,10) + '.xlsx');
  workbook.xlsx.write(res).then(() => res.end());
});

// Admin: JSON-Rohdaten
app.get('/admin/data', (req, res) => {
  const token = req.query.token;
  if (token !== process.env.ADMIN_TOKEN && token !== 'admin123') {
    return res.status(401).send('Nicht autorisiert');
  }
  res.json(loadData());
});

app.listen(PORT, () => {
  console.log('Server läuft auf http://localhost:' + PORT);
  console.log('Admin-Export: http://localhost:' + PORT + '/admin/export?token=admin123');
});
