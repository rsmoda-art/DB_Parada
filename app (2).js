const API = 'https://script.google.com/macros/s/AKfycbwFNenq_AQRFDt2ZxfWr4u4DPBtpGRVZHpYQifvKuI8jelD3lxIB7EJrPc1QCugUFvG/exec';
let tasks = [], map = {}, chart;
const $ = id => document.getElementById(id);
$('user').value = localStorage.user || '';
$('user').onchange = e => localStorage.user = e.target.value;

async function load() {
  const r = await (await fetch(API)).json();
  tasks = r.data; calc(); render();
}

function calc() {
  map = Object.fromEntries(tasks.map(t => [t.id, t]));
  tasks.forEach(t => {
    t.level = t.id.split('.').length;
    t.leaf = !tasks.some(c => c.id.startsWith(t.id + '.'));
  });
  tasks.forEach(t => {
    const leaves = t.leaf ? [t] : tasks.filter(c => c.leaf && c.id.startsWith(t.id + '.'));
    t.total = leaves.length; t.done = leaves.filter(l => l.exec).length;
    t.pct = t.total ? Math.round(t.done / t.total * 100) : 0;
  });
  tasks.forEach(t => {
    t.preds = (t.pred || '').split(/[;,]/).map(s => s.trim()).filter(Boolean);
    t.blocked = t.preds.filter(p => !map[p] || map[p].pct < 100);
    t.succ = tasks.filter(s => (s.pred || '').split(/[;,]/).map(x => x.trim()).includes(t.id));
    t.status = t.pct === 100 ? 'Concluído' : t.blocked.length ? 'Bloqueado' : 'Liberado';
  });
}

function render() {
  const f = $('filtro').value, leaves = tasks.filter(t => t.leaf);
  const all = leaves.length, ok = leaves.filter(t => t.exec).length;
  $('kGeral').textContent = Math.round(ok / all * 100) + '%';
  $('kOk').textContent = ok;
  $('kLib').textContent = leaves.filter(t => t.status === 'Liberado').length;
  $('kBlq').textContent = leaves.filter(t => t.status === 'Bloqueado').length;

  $('tb').innerHTML = tasks.filter(t => !f || t.status === f).map(t => `
    <tr class="${t.leaf ? '' : 'topic'} st-${t.status}">
      <td>${t.id}</td>
      <td style="padding-left:${(t.level - 1) * 18 + 8}px">${t.desc}</td>
      <td title="${t.blocked.length ? 'Aguardando: ' + t.blocked.join(', ') : ''}">${t.pred || ''}</td>
      <td><div class="bar"><i style="width:${t.pct}%"></i></div>${t.pct}%</td>
      <td>${t.status}</td>
      <td>${t.leaf ? `<input type="checkbox" ${t.exec ? 'checked' : ''}
        ${t.blocked.length && !t.exec ? 'disabled' : ''} onchange="toggle('${t.id}',this)">` : ''}</td>
    </tr>`).join('');
  drawChart();
}

async function toggle(id, el) {
  const t = map[id], user = $('user').value.trim();
  if (!user) { alert('Informe seu nome'); el.checked = !el.checked; return; }
  if (!el.checked && t.succ.some(s => s.exec)) {
    alert('Sucessoras já executadas: ' + t.succ.filter(s => s.exec).map(s => s.id).join(', '));
    el.checked = true; return;
  }
  if (el.checked && !confirm(`Confirmar execução de ${id}?`)) { el.checked = false; return; }
  t.exec = el.checked; calc(); render();
  await fetch(API, { method: 'POST', body: JSON.stringify({ id, exec: t.exec, user }) });
  load();
}

function drawChart() {
  const tops = tasks.filter(t => t.level === 1);
  const cfg = { labels: tops.map(t => `${t.id} ${t.desc.slice(0, 30)}`), data: tops.map(t => t.pct) };
  if (chart) { chart.data.labels = cfg.labels; chart.data.datasets[0].data = cfg.data; return chart.update(); }
  chart = new Chart($('chart'), {
    type: 'bar',
    data: { labels: cfg.labels, datasets: [{ label: '% Executado', data: cfg.data, backgroundColor: '#2e6b4f' }] },
    options: { indexAxis: 'y', animation: false, scales: { x: { max: 100 } } }
  });
}

async function exportXlsx() {
  const wb = new ExcelJS.Workbook(), now = new Date().toLocaleString('pt-BR');
  const hdr = r => r.eachCell(c => { c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E6B4F' } }; });
  const color = { 'Concluído': 'FFC6EFCE', 'Liberado': 'FFFFEB9C', 'Bloqueado': 'FFFFC7CE' };
  const leaves = tasks.filter(t => t.leaf), ok = leaves.filter(t => t.exec).length;

  // Dashboard
  const d = wb.addWorksheet('Dashboard');
  d.columns = [{ width: 12 }, { width: 55 }, { width: 14 }, { width: 14 }, { width: 14 }];
  d.addRow(['PARADA DE FÁBRICA – NESTLÉ MONTES CLAROS | SUBESTAÇÃO 138 kV']).font = { bold: true, size: 14 };
  d.addRow([`Atualizado em ${now}`]);
  d.addRow([]);
  hdr(d.addRow(['Geral', 'Concluídas', 'Liberadas', 'Bloqueadas', 'Total']));
  d.addRow([ok / leaves.length, ok, leaves.filter(t => t.status === 'Liberado').length,
    leaves.filter(t => t.status === 'Bloqueado').length, leaves.length]).getCell(1).numFmt = '0%';
  d.addRow([]);
  hdr(d.addRow(['ID', 'Tópico', 'Executadas', 'Total', '% Avanço']));
  tasks.filter(t => !t.leaf || t.level === 1).forEach(t => {
    const r = d.addRow([t.id, t.desc, t.done, t.total, t.pct / 100]);
    r.getCell(5).numFmt = '0%'; if (t.level === 1) r.font = { bold: true };
  });
  const img = wb.addImage({ base64: $('chart').toDataURL('image/png'), extension: 'png' });
  d.addImage(img, { tl: { col: 6, row: 3 }, ext: { width: 650, height: 320 } });

  // Atividades
  const a = wb.addWorksheet('Atividades');
  a.columns = [{ width: 10 }, { width: 60 }, { width: 14 }, { width: 8 }, { width: 12 }, { width: 18 }, { width: 18 }];
  hdr(a.addRow(['ID', 'Descrição', 'Predecessora', '%', 'Status', 'Data/Hora', 'Usuário']));
  tasks.forEach(t => {
    const r = a.addRow([t.id, t.desc, t.pred, t.pct / 100, t.status, t.data, t.user]);
    r.getCell(4).numFmt = '0%'; if (!t.leaf) r.font = { bold: true };
    r.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color[t.status] } };
  });

  // Bloqueadas
  const b = wb.addWorksheet('Bloqueadas');
  b.columns = [{ width: 10 }, { width: 60 }, { width: 25 }];
  hdr(b.addRow(['ID', 'Descrição', 'Aguardando']));
  leaves.filter(t => t.status === 'Bloqueado').forEach(t => b.addRow([t.id, t.desc, t.blocked.join('; ')]));

  const blob = new Blob([await wb.xlsx.writeBuffer()]);
  const l = document.createElement('a');
  l.href = URL.createObjectURL(blob);
  l.download = `Parada_SE138kV_${new Date().toISOString().slice(0, 10)}.xlsx`; l.click();
}

load(); setInterval(load, 30000);
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js');
