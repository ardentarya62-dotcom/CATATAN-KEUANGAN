<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#4f46e5">
<title>Financial Tracker — Keuangan Pribadi</title>
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<style>
  * { -webkit-tap-highlight-color: transparent; }
  body { padding-bottom: max(24px, env(safe-area-inset-bottom)); }
  input, select, button { font: inherit; }
  input, select { min-height: 44px; }
  button { touch-action: manipulation; }
  .table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  table { min-width: 620px; }
  @media (max-width: 640px) {
    .page-header { padding-top: 18px; padding-bottom: 18px; }
    .page-header-inner { align-items: flex-start; gap: 12px; }
    .page-header h1 { font-size: 1.2rem; line-height: 1.5rem; }
    .page-header p { font-size: .75rem; line-height: 1rem; max-width: 220px; }
    .export-btn { padding: 10px 12px; white-space: nowrap; }
    main { padding-left: 12px !important; padding-right: 12px !important; }
    .panel { padding: 16px !important; border-radius: 16px; }
    .summary-number { font-size: 1.05rem !important; overflow-wrap: anywhere; }
  }
</style>
</head>
<body class="bg-slate-100 min-h-screen text-slate-800 font-sans">
<header class="page-header bg-indigo-600 text-white py-6 shadow-lg mb-6 sm:mb-8">
  <div class="page-header-inner max-w-6xl mx-auto px-4 flex justify-between items-center">
    <div>
      <h1 class="text-2xl font-bold flex items-center gap-2"><i class="fa-solid fa-wallet"></i> Financial Tracker</h1>
      <p class="text-indigo-200 text-sm mt-1">Catat pemasukan & pengeluaran dengan mudah</p>
    </div>
    <button type="button" onclick="exportToExcel()" class="export-btn bg-emerald-500 hover:bg-emerald-600 text-white font-semibold rounded-lg shadow transition flex items-center gap-2 text-sm">
      <i class="fa-solid fa-file-excel"></i><span>Export Excel</span>
    </button>
  </div>
</header>

<main class="max-w-6xl mx-auto px-4 grid grid-cols-1 lg:grid-cols-3 gap-5 sm:gap-8">
  <section class="space-y-5 sm:space-y-6">
    <div class="panel bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
      <h2 class="text-lg font-semibold mb-3 text-slate-700 flex items-center justify-between gap-2">
        <span><i class="fa-solid fa-vault text-indigo-500 mr-2"></i>Saldo Awal</span>
        <button type="button" onclick="toggleEditSaldo()" class="text-sm text-indigo-600 hover:underline py-2 px-1">Edit</button>
      </h2>
      <div id="saldoDisplay" class="text-2xl font-extrabold text-indigo-600">Rp 0</div>
      <div id="saldoInputContainer" class="hidden mt-3">
        <label for="saldoAwalInput" class="block text-xs font-medium text-slate-600 mb-1">Saldo awal bulan (Rp)</label>
        <input type="number" id="saldoAwalInput" min="0" step="1" inputmode="numeric" placeholder="Contoh: 1000000" class="w-full border rounded-lg p-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none mb-2">
        <button type="button" onclick="simpanSaldoAwal()" class="w-full bg-indigo-600 text-white py-2 rounded-lg text-sm font-medium hover:bg-indigo-700">Simpan Saldo</button>
      </div>
    </div>

    <div class="panel bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
      <h2 class="text-lg font-semibold mb-4 text-slate-700 flex items-center gap-2"><i class="fa-solid fa-pen-to-square text-indigo-500"></i>Tambah Transaksi</h2>
      <form id="transactionForm" class="space-y-4">
        <div>
          <label for="jenis" class="block text-xs font-medium text-slate-600 mb-1">Jenis Transaksi</label>
          <select id="jenis" class="w-full border rounded-lg p-2.5 text-sm bg-slate-50 focus:ring-2 focus:ring-indigo-500 outline-none">
            <option value="pengeluaran">Pengeluaran</option><option value="pemasukan">Pemasukan</option>
          </select>
        </div>
        <div>
          <label for="tanggal" class="block text-xs font-medium text-slate-600 mb-1">Tanggal</label>
          <input type="date" id="tanggal" required class="w-full border rounded-lg p-2.5 text-sm bg-slate-50 focus:ring-2 focus:ring-indigo-500 outline-none">
        </div>
        <div>
          <label for="kategori" class="block text-xs font-medium text-slate-600 mb-1">Kategori</label>
          <select id="kategori" required class="w-full border rounded-lg p-2.5 text-sm bg-slate-50 focus:ring-2 focus:ring-indigo-500 outline-none"></select>
        </div>
        <div>
          <label for="nominal" class="block text-xs font-medium text-slate-600 mb-1">Nominal (Rp)</label>
          <input type="number" id="nominal" placeholder="Contoh: 50000" required min="1" step="1" inputmode="numeric" class="w-full border rounded-lg p-2.5 text-sm bg-slate-50 focus:ring-2 focus:ring-indigo-500 outline-none">
        </div>
        <div>
          <label for="keterangan" class="block text-xs font-medium text-slate-600 mb-1">Keterangan</label>
          <input type="text" id="keterangan" maxlength="200" placeholder="Contoh: Beli bensin / makan siang" class="w-full border rounded-lg p-2.5 text-sm bg-slate-50 focus:ring-2 focus:ring-indigo-500 outline-none">
        </div>
        <button type="submit" class="w-full bg-indigo-600 hover:bg-indigo-700 active:bg-indigo-800 text-white font-semibold py-3 rounded-lg shadow transition">Tambah Transaksi</button>
      </form>
      <p class="text-xs text-slate-400 mt-3">Data disimpan di browser/perangkat yang digunakan.</p>
    </div>
  </section>

  <section class="lg:col-span-2 space-y-5 sm:space-y-6">
    <div class="grid grid-cols-1 sm:grid-cols-3 gap-3 sm:gap-4">
      <div class="bg-white p-4 sm:p-5 rounded-xl border border-slate-200 shadow-sm"><p class="text-xs text-slate-500 font-medium">Total Pemasukan</p><p id="totalPemasukan" class="summary-number text-xl font-bold text-emerald-600 mt-1">Rp 0</p></div>
      <div class="bg-white p-4 sm:p-5 rounded-xl border border-slate-200 shadow-sm"><p class="text-xs text-slate-500 font-medium">Total Pengeluaran</p><p id="totalPengeluaran" class="summary-number text-xl font-bold text-rose-600 mt-1">Rp 0</p></div>
      <div class="bg-white p-4 sm:p-5 rounded-xl border border-slate-200 shadow-sm"><p class="text-xs text-slate-500 font-medium">Sisa Saldo</p><p id="saldoAkhir" class="summary-number text-xl font-bold text-indigo-600 mt-1">Rp 0</p></div>
    </div>

    <div class="panel bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
      <div class="flex flex-col gap-4 mb-5">
        <h2 class="text-lg font-semibold text-slate-700 flex items-center gap-2"><i class="fa-solid fa-chart-pie text-indigo-500"></i>Analisis Pengeluaran</h2>
        <div class="grid grid-cols-1 sm:grid-cols-[1fr_auto_1fr_auto] gap-2 items-center text-sm">
          <div><label for="filterStart" class="block text-xs text-slate-500 mb-1">Dari tanggal</label><input type="date" id="filterStart" class="w-full border rounded-lg p-2 text-xs bg-slate-50" onchange="renderApp()"></div>
          <span class="hidden sm:block text-slate-400 text-center pt-4">s/d</span>
          <div><label for="filterEnd" class="block text-xs text-slate-500 mb-1">Sampai tanggal</label><input type="date" id="filterEnd" class="w-full border rounded-lg p-2 text-xs bg-slate-50" onchange="renderApp()"></div>
          <button type="button" onclick="resetFilter()" class="bg-slate-200 hover:bg-slate-300 text-slate-700 px-4 py-2.5 rounded-lg text-xs mt-0 sm:mt-4">Reset Filter</button>
        </div>
      </div>
      <div id="chartEmpty" class="hidden text-center text-sm text-slate-400 py-8">Belum ada data pengeluaran untuk rentang tanggal ini.</div>
      <div class="w-full max-w-sm mx-auto relative"><canvas id="expenseChart" aria-label="Diagram pengeluaran per kategori"></canvas></div>
      <div id="categoryPercentages" class="mt-5 space-y-3"></div>
    </div>

    <div class="panel bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
      <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
        <h2 class="text-lg font-semibold text-slate-700 flex items-center gap-2"><i class="fa-solid fa-list-check text-indigo-500"></i>Riwayat Transaksi</h2>
        <p id="transactionCount" class="text-xs text-slate-400"></p>
      </div>
      <div class="table-wrap rounded-lg border border-slate-100">
        <table class="w-full text-left text-sm">
          <thead class="bg-slate-50 text-slate-500 border-b"><tr><th class="p-3">Tanggal</th><th class="p-3">Kategori</th><th class="p-3">Keterangan</th><th class="p-3 text-right">Nominal</th><th class="p-3 text-center">Aksi</th></tr></thead>
          <tbody id="transactionTable" class="divide-y divide-slate-100"></tbody>
        </table>
      </div>
      <p class="text-xs text-slate-400 mt-2">Geser tabel ke kiri/kanan jika dibuka dari layar kecil.</p>
    </div>
  </section>
</main>

<script>
const categories = [
  "Perpuluhan/Persembahan", "Sewa Kost", "Konsumsi", "Transportasi",
  "Tagihan Listrik", "Tagihan Pulsa/Internet", "Body Care", "Hobi/Hiburan",
  "Jajan/Cemilan", "Dinner With GF", "Buat Keluarga", "Maintenance Kendaraan",
  "Dana Tidak Terduga", "Gaji/Pemasukan Lain"
];
function readNumber(key) { const n = Number(localStorage.getItem(key)); return Number.isFinite(n) ? n : 0; }
let saldoAwal = readNumber('saldoAwal');
let transactions = [];
try { const saved = JSON.parse(localStorage.getItem('transactions') || '[]'); transactions = Array.isArray(saved) ? saved : []; } catch (_) { transactions = []; }
let chartInstance = null;

function localDateISO(date = new Date()) {
  const y = date.getFullYear(), m = String(date.getMonth()+1).padStart(2,'0'), d = String(date.getDate()).padStart(2,'0');
  return `${y}-${m}-${d}`;
}
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('tanggal').value = localDateISO();
  populateCategories();
  renderApp();
});
function populateCategories() {
  document.getElementById('kategori').innerHTML = categories.map(cat => `<option value="${escapeHtml(cat)}">${escapeHtml(cat)}</option>`).join('');
}
function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
}
function toggleEditSaldo() {
  const box = document.getElementById('saldoInputContainer');
  box.classList.toggle('hidden');
  document.getElementById('saldoAwalInput').value = saldoAwal || '';
}
function simpanSaldoAwal() {
  const raw = document.getElementById('saldoAwalInput').value;
  if (raw === '' || !Number.isFinite(Number(raw)) || Number(raw) < 0) { alert('Masukkan saldo awal yang valid (minimal Rp 0).'); return; }
  saldoAwal = Number(raw);
  localStorage.setItem('saldoAwal', String(saldoAwal));
  document.getElementById('saldoInputContainer').classList.add('hidden');
  renderApp();
}
document.getElementById('transactionForm').addEventListener('submit', e => {
  e.preventDefault();
  const jenis = document.getElementById('jenis').value;
  const tanggal = document.getElementById('tanggal').value;
  const kategori = document.getElementById('kategori').value;
  const nominal = Number(document.getElementById('nominal').value);
  const keterangan = document.getElementById('keterangan').value.trim();
  if (!tanggal || !kategori || !Number.isFinite(nominal) || nominal <= 0) { alert('Periksa kembali tanggal, kategori, dan nominal transaksi.'); return; }
  transactions.push({ id: Date.now(), jenis, tanggal, kategori, nominal, keterangan });
  saveAndRender();
  document.getElementById('nominal').value = '';
  document.getElementById('keterangan').value = '';
});
function hapusTransaksi(id) {
  if (!confirm('Yakin ingin menghapus transaksi ini?')) return;
  transactions = transactions.filter(t => t.id !== id);
  saveAndRender();
}
function saveAndRender() {
  try {
    localStorage.setItem('transactions', JSON.stringify(transactions));
    renderApp();
  } catch (err) {
    alert('Browser tidak dapat menyimpan data. Export laporan secara berkala agar data tidak hilang.');
  }
}
function getFilteredTransactions() {
  const start = document.getElementById('filterStart').value;
  const end = document.getElementById('filterEnd').value;
  if (start && end && start > end) { return []; }
  return transactions.filter(t => (!start || t.tanggal >= start) && (!end || t.tanggal <= end));
}
function resetFilter() {
  document.getElementById('filterStart').value = '';
  document.getElementById('filterEnd').value = '';
  renderApp();
}
function formatRp(number) {
  return new Intl.NumberFormat('id-ID', { style:'currency', currency:'IDR', maximumFractionDigits:0 }).format(Number(number) || 0);
}
function displayDate(date) {
  if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return date || '-';
  const [y,m,d] = date.split('-');
  return `${d}/${m}/${y}`;
}
function renderApp() {
  const filteredTx = getFilteredTransactions().slice().sort((a,b) => b.tanggal.localeCompare(a.tanggal) || b.id-a.id);
  document.getElementById('saldoDisplay').textContent = formatRp(saldoAwal);
  const totalIn = filteredTx.filter(t => t.jenis === 'pemasukan').reduce((sum,t) => sum + Number(t.nominal || 0), 0);
  const totalOut = filteredTx.filter(t => t.jenis === 'pengeluaran').reduce((sum,t) => sum + Number(t.nominal || 0), 0);
  document.getElementById('totalPemasukan').textContent = formatRp(totalIn);
  document.getElementById('totalPengeluaran').textContent = formatRp(totalOut);
  document.getElementById('saldoAkhir').textContent = formatRp(saldoAwal + totalIn - totalOut);
  document.getElementById('transactionCount').textContent = `${filteredTx.length} transaksi ditampilkan`;
  const tbody = document.getElementById('transactionTable');
  if (!filteredTx.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="text-center p-5 text-slate-400">Belum ada transaksi.</td></tr>';
  } else {
    tbody.innerHTML = filteredTx.map(t => `<tr class="hover:bg-slate-50">
      <td class="p-3 whitespace-nowrap text-slate-600">${escapeHtml(displayDate(t.tanggal))}</td>
      <td class="p-3 font-medium text-slate-700">${escapeHtml(t.kategori)}</td>
      <td class="p-3 text-slate-500">${escapeHtml(t.keterangan || '-')}</td>
      <td class="p-3 text-right whitespace-nowrap font-bold ${t.jenis === 'pemasukan' ? 'text-emerald-600' : 'text-rose-600'}">${t.jenis === 'pemasukan' ? '+' : '−'} ${escapeHtml(formatRp(t.nominal))}</td>
      <td class="p-3 text-center"><button type="button" onclick="hapusTransaksi(${Number(t.id)})" aria-label="Hapus transaksi" class="text-rose-500 hover:text-rose-700 p-2"><i class="fa-solid fa-trash"></i></button></td>
    </tr>`).join('');
  }
  renderChart(filteredTx);
}
function renderChart(data) {
  const expenseByCat = {};
  data.filter(t => t.jenis === 'pengeluaran').forEach(t => expenseByCat[t.kategori] = (expenseByCat[t.kategori] || 0) + Number(t.nominal || 0));
  const labels = Object.keys(expenseByCat);
  const values = Object.values(expenseByCat);
  document.getElementById('chartEmpty').classList.toggle('hidden', values.length > 0);
  document.getElementById('expenseChart').parentElement.classList.toggle('hidden', values.length === 0);
  if (chartInstance) { chartInstance.destroy(); chartInstance = null; }
  if (values.length) {
    chartInstance = new Chart(document.getElementById('expenseChart').getContext('2d'), {
      type:'doughnut',
      data:{ labels, datasets:[{ data:values, backgroundColor:['#f87171','#38bdf8','#fbbf24','#4ade80','#a78bfa','#f472b6','#fb923c','#2dd4bf','#818cf8','#c084fc','#f43f5e','#10b981','#6366f1','#94a3b8'] }] },
      options:{ responsive:true, maintainAspectRatio:true, plugins:{ legend:{ position:'bottom', labels:{ boxWidth:12, padding:14, font:{size:11} } }, tooltip:{ callbacks:{ label: ctx => `${ctx.label}: ${formatRp(ctx.raw)} (${(ctx.raw / values.reduce((a,b)=>a+b,0)*100).toFixed(1)}%)` } } } }
    });
  }
  const total = values.reduce((a,b)=>a+b,0);
  const sorted = labels.map((label,i)=>({label,value:values[i]})).sort((a,b)=>b.value-a.value);
  document.getElementById('categoryPercentages').innerHTML = total ? sorted.map(item => {
    const pct = item.value / total * 100;
    return `<div><div class="flex justify-between gap-3 text-xs mb-1"><span class="text-slate-600">${escapeHtml(item.label)}</span><span class="font-semibold text-slate-700 whitespace-nowrap">${pct.toFixed(1)}% · ${escapeHtml(formatRp(item.value))}</span></div><div class="h-2 bg-slate-100 rounded-full overflow-hidden"><div class="h-full bg-indigo-500 rounded-full" style="width:${pct}%"></div></div></div>`;
  }).join('') : '';
}
function exportToExcel() {
  if (typeof XLSX === 'undefined') { alert('Fitur Excel memerlukan internet. Hubungkan perangkat ke internet lalu coba lagi.'); return; }
  const rows = transactions.slice().sort((a,b)=>a.tanggal.localeCompare(b.tanggal)).map(t => ({
    Tanggal:t.tanggal, Jenis:t.jenis.toUpperCase(), Kategori:t.kategori, Keterangan:t.keterangan || '-', Nominal:Number(t.nominal || 0)
  }));
  if (!rows.length) rows.push({Tanggal:'', Jenis:'', Kategori:'', Keterangan:'Belum ada transaksi', Nominal:0});
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [{wch:14},{wch:16},{wch:28},{wch:36},{wch:18}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Riwayat Transaksi');
  const summary = [['Ringkasan Keuangan','Nilai'],['Saldo Awal',saldoAwal],['Total Pemasukan',transactions.filter(t=>t.jenis==='pemasukan').reduce((s,t)=>s+Number(t.nominal||0),0)],['Total Pengeluaran',transactions.filter(t=>t.jenis==='pengeluaran').reduce((s,t)=>s+Number(t.nominal||0),0)]];
  summary.push(['Saldo Akhir',summary[1][1]+summary[2][1]-summary[3][1]]);
  const catTotals = {};
  transactions.filter(t=>t.jenis==='pengeluaran').forEach(t=>catTotals[t.kategori]=(catTotals[t.kategori]||0)+Number(t.nominal||0));
  const totalOut = Object.values(catTotals).reduce((s,n)=>s+n,0);
  summary.push([],['Kategori Pengeluaran','Total','Persentase']);
  Object.entries(catTotals).sort((a,b)=>b[1]-a[1]).forEach(([cat,amount])=>summary.push([cat,amount,totalOut ? amount/totalOut : 0]));
  const sws = XLSX.utils.aoa_to_sheet(summary);
  sws['!cols'] = [{wch:30},{wch:20},{wch:16}];
  Object.keys(sws).filter(k=>/^[A-Z]+\d+$/.test(k) && Number(k.match(/\d+/)[0])>=7 && k[0]==='C').forEach(k=>sws[k].z='0.0%');
  XLSX.utils.book_append_sheet(wb, sws, 'Ringkasan');
  XLSX.writeFile(wb, `Laporan_Keuangan_${localDateISO()}.xlsx`);
}
</script>
</body>
</html>
