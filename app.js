import { profile, loadDocx, repairDocx, demoDocx } from './docx.js';

const $ = id => document.getElementById(id);
const state = { file: null, docx: null, filter: 'all', limit: 60, busy: false, modalReturnFocus: null, resultUrl: null };
const escapeHtml = value => String(value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
const formatBytes = size => size < 1024 * 1024 ? (size / 1024).toFixed(0) + ' KB' : (size / 1024 / 1024).toFixed(1) + ' MB';
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
let toastTimer;

function toast(message, error = false) {
  const item = $('toast');
  item.textContent = message;
  item.classList.toggle('error', error);
  item.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => item.classList.add('hidden'), 5000);
}
function setStep(step) {
  [...$('step-indicator').children].forEach((item, index) => item.classList.toggle('step-active', index === step));
}
function openModal() {
  state.modalReturnFocus = document.activeElement;
  $('preset-modal').classList.remove('hidden');
  $('preset-select').focus();
}
function closeModal() {
  $('preset-modal').classList.add('hidden');
  state.modalReturnFocus?.focus?.();
}
function chooseFile(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.docx')) return toast('Hanya file .docx yang dapat diproses.', true);
  if (file.size > 25 * 1024 * 1024) return toast('Ukuran file maksimal 25 MB.', true);
  state.file = file;
  openModal();
}
function showAnalysis() {
  $('upload-view').classList.add('hidden');
  $('analysis-view').classList.remove('hidden');
  $('file-name').textContent = state.file.name;
  const position = $('page-number-position').value;
  const label = { keep: 'nomor halaman tetap', 'bottom-center': 'nomor bawah tengah', 'bottom-right': 'nomor bawah kanan' }[position];
  $('file-meta').textContent = formatBytes(state.file.size) + ' · DOCX · ' + label;
  $('profile-name').textContent = profile.shortName;
  $('workspace').scrollIntoView({ behavior: 'smooth', block: 'start' });
  setStep(1);
}
function clearResult() {
  if (state.resultUrl) URL.revokeObjectURL(state.resultUrl);
  state.resultUrl = null;
  $('download-result').removeAttribute('href');
  $('download-result').classList.add('hidden');
  $('result-banner').classList.add('hidden');
  $('action-title').textContent = 'Siap dirapikan?';
  $('action-description').textContent = 'Perubahan akan dibuat pada file baru yang Anda unduh.';
}
function setProgress(percent) {
  $('progress-track').classList.remove('hidden');
  $('progress-fill').style.width = percent + '%';
}
function stopProgress() {
  $('progress-fill').style.width = '100%';
  setTimeout(() => $('progress-track').classList.add('hidden'), 450);
}
function renderPreview(report) {
  const root = $('document-preview');
  if (!report.paragraphs.length) {
    root.innerHTML = '<div class="preview-placeholder">Dokumen tidak memiliki paragraf teks yang dapat ditampilkan.</div>';
    return;
  }
  const marked = new Set(report.findings.filter(f => f.paragraphIndex !== null).map(f => f.paragraphIndex));
  let html = '', page = null;
  for (const item of report.paragraphs.slice(0, 250)) {
    if (item.page !== page) {
      if (page !== null) html += '</div>';
      page = item.page;
      html += `<div class="document-page"><span class="page-label">HALAMAN ${page} (PERKIRAAN)</span>`;
    }
    const heading = /heading|title/i.test(item.styleId) || (item.index === 0 && item.text.length < 100);
    html += `<p id="preview-p-${item.index}" class="${heading ? 'heading-text ' : ''}${marked.has(item.index) ? 'highlighted' : ''}">${escapeHtml(item.text)}</p>`;
  }
  html += '</div>';
  if (report.paragraphs.length > 250) html += '<div class="preview-placeholder">Pratinjau dibatasi pada 250 paragraf pertama. Seluruh dokumen tetap diperiksa.</div>';
  root.innerHTML = html;
}
function filteredFindings() {
  const findings = state.docx?.report.findings || [];
  return state.filter === 'all' ? findings : findings.filter(item => item.category === state.filter);
}
function renderFindings() {
  const root = $('findings-list');
  const findings = filteredFindings();
  $('findings-badge').textContent = findings.length + ' temuan';
  if (!findings.length) {
    root.innerHTML = '<div class="empty-findings"><span>✓</span><strong>Tidak ada temuan</strong><p>Format yang diperiksa sudah sesuai dengan profil ini.</p></div>';
    return;
  }
  const icon = { page: '▣', font: 'Aa', spacing: '↕', paragraph: '¶', structure: '1.1' };
  root.innerHTML = findings.slice(0, state.limit).map(item => `<div class="finding" data-category="${item.category}" data-paragraph="${item.paragraphIndex ?? ''}" role="button" tabindex="0" aria-label="${escapeHtml(item.title)}, halaman perkiraan ${item.page}"><span class="finding-icon">${icon[item.category]}</span><div class="finding-body"><div class="finding-top"><strong>${escapeHtml(item.title)}</strong><small>Hlm ${item.page}*</small></div><p>${escapeHtml(item.detail)}</p>${item.snippet ? `<span class="finding-snippet">“${escapeHtml(item.snippet)}”</span>` : ''}</div></div>`).join('');
  if (findings.length > state.limit) root.insertAdjacentHTML('beforeend', `<button class="show-more" type="button">Tampilkan ${Math.min(60, findings.length - state.limit)} temuan lagi</button>`);
}
function renderReport() {
  const report = state.docx.report;
  $('finding-count').textContent = report.findings.length.toLocaleString('id-ID');
  $('paragraph-count').textContent = report.paragraphs.length.toLocaleString('id-ID');
  $('summary-title').textContent = report.findings.length ? report.findings.length + ' hal yang perlu dirapikan' : 'Format sudah sesuai';
  $('summary-description').textContent = report.findings.length ? 'Kami menemukan beberapa pengaturan yang belum mengikuti profil pilihan Anda.' : 'Semua aturan yang diperiksa sudah cocok dengan profil pilihan Anda.';
  renderPreview(report);
  renderFindings();
}
async function analyze() {
  if (!state.file || state.busy) return;
  state.busy = true;
  clearResult();
  closeModal();
  showAnalysis();
  $('summary-title').textContent = 'Memeriksa dokumen...';
  $('summary-description').textContent = 'Membaca struktur dan pengaturan dokumen.';
  $('findings-list').innerHTML = '<div class="preview-placeholder">Memindai format...</div>';
  $('fix-button').disabled = true;
  setProgress(12);
  const start = performance.now();
  try {
    await delay(150);
    setProgress(34);
    state.docx = await loadDocx(state.file);
    setProgress(83);
    await delay(Math.max(0, 750 - (performance.now() - start)));
    state.filter = 'all'; state.limit = 60;
    [...$('filter-bar').children].forEach(button => button.classList.toggle('active', button.dataset.filter === 'all'));
    renderReport();
    stopProgress();
    $('fix-button').disabled = false;
  } catch (error) {
    $('analysis-view').classList.add('hidden');
    $('upload-view').classList.remove('hidden');
    setStep(0);
    toast(error.message || 'Dokumen tidak dapat diperiksa.', true);
    $('progress-track').classList.add('hidden');
  } finally { state.busy = false; }
}
async function fix() {
  if (!state.docx || state.busy) return;
  state.busy = true;
  const button = $('fix-button'), original = button.innerHTML;
  button.disabled = true; button.textContent = 'Merapikan dokumen...';
  setProgress(15);
  try {
    const before = state.docx.report;
    const position = $('page-number-position').value;
    const blob = await repairDocx(state.docx, { pageNumberPosition: position });
    setProgress(90);
    const filename = state.file.name.replace(/\.docx$/i, '') + '_rapi.docx';
    const checked = await loadDocx(new File([blob], filename, { type: blob.type }));
    if (checked.report.findings.length) throw new Error('File hasil masih memiliki temuan format. Unduhan dibatalkan agar Anda tidak menerima hasil yang keliru.');
    state.docx = checked;
    renderReport();
    if (state.resultUrl) URL.revokeObjectURL(state.resultUrl);
    state.resultUrl = URL.createObjectURL(blob);
    const link = $('download-result');
    link.href = state.resultUrl;
    link.download = filename;
    link.classList.remove('hidden');
    const count = before.findings.length;
    $('result-title').textContent = count ? count + ' temuan telah diperbaiki' : 'File hasil telah dibuat';
    $('result-description').textContent = 'File DOCX baru sudah diperiksa ulang. Buka file hasil untuk melihat tata letak sebenarnya.';
    const bullets = [];
    const categories = new Set(before.findings.map(item => item.category));
    if (categories.has('page')) bullets.push('Margin dan ukuran kertas disesuaikan dengan profil.');
    if (categories.has('font')) bullets.push('Jenis dan ukuran huruf disetel ke Times New Roman 12 pt.');
    if (categories.has('spacing')) bullets.push('Spasi baris disetel ke 1,5.');
    if (categories.has('paragraph')) bullets.push('Paragraf isi diratakan kiri-kanan, awal baris menjorok 1,25 cm, dan jarak antarapagraf dinormalkan.');
    if (categories.has('structure')) bullets.push('Urutan judul BAB dan subbab 1.1 disesuaikan.');
    if (position !== 'keep') bullets.push('Nomor halaman ditempatkan di ' + (position === 'bottom-center' ? 'bawah tengah.' : 'bawah kanan.'));
    if (!bullets.length) bullets.push('Format yang diperiksa sudah sesuai.');
    $('result-changes').innerHTML = bullets.map(text => '<li>' + escapeHtml(text) + '</li>').join('');
    $('result-banner').classList.remove('hidden');
    $('action-title').textContent = 'File hasil siap';
    $('action-description').textContent = 'Klik “Unduh file hasil” untuk menyimpan DOCX yang sudah diperbaiki.';
    setStep(2);
    link.click();
    toast('File hasil siap. Jika belum terunduh, klik “Unduh file hasil”.');
  } catch (error) {
    toast(error.message || 'Gagal membuat dokumen baru.', true);
  } finally {
    stopProgress();
    button.disabled = false; button.innerHTML = original;
    state.busy = false;
  }
}
function reset() {
  clearResult();
  state.file = null; state.docx = null; state.filter = 'all';
  $('file-input').value = '';
  $('analysis-view').classList.add('hidden');
  $('upload-view').classList.remove('hidden');
  setStep(0);
  $('workspace').scrollIntoView({ behavior: 'smooth', block: 'start' });
}
const dropZone = $('drop-zone');
dropZone.addEventListener('click', () => $('file-input').click());
dropZone.addEventListener('keydown', event => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); $('file-input').click(); } });
$('file-input').addEventListener('change', event => chooseFile(event.target.files[0]));
for (const name of ['dragenter', 'dragover']) dropZone.addEventListener(name, event => { event.preventDefault(); dropZone.classList.add('drag-over'); });
for (const name of ['dragleave', 'drop']) dropZone.addEventListener(name, event => { event.preventDefault(); dropZone.classList.remove('drag-over'); });
dropZone.addEventListener('drop', event => chooseFile(event.dataTransfer.files[0]));
document.addEventListener('dragover', event => event.preventDefault());
document.addEventListener('drop', event => event.preventDefault());
$('demo-button').addEventListener('click', async () => { try { chooseFile(await demoDocx()); } catch { toast('Dokumen contoh tidak dapat dibuat.', true); } });
$('modal-close').addEventListener('click', closeModal);
$('preset-modal').addEventListener('click', event => { if (event.target === $('preset-modal')) closeModal(); });
document.addEventListener('keydown', event => { if (event.key === 'Escape' && !$('preset-modal').classList.contains('hidden')) closeModal(); });
$('analyze-button').addEventListener('click', analyze);
$('change-file').addEventListener('click', reset);
$('change-preset').addEventListener('click', openModal);
$('fix-button').addEventListener('click', fix);
$('filter-bar').addEventListener('click', event => {
  const button = event.target.closest('[data-filter]');
  if (!button) return;
  state.filter = button.dataset.filter; state.limit = 60;
  [...$('filter-bar').children].forEach(item => item.classList.toggle('active', item === button));
  renderFindings();
});
$('findings-list').addEventListener('click', event => {
  if (event.target.closest('.show-more')) { state.limit += 60; renderFindings(); return; }
  const item = event.target.closest('.finding');
  if (!item) return;
  const paragraph = $('preview-p-' + item.dataset.paragraph);
  if (paragraph) { paragraph.scrollIntoView({ behavior: 'smooth', block: 'center' }); paragraph.classList.add('selected'); setTimeout(() => paragraph.classList.remove('selected'), 1600); }
});
$('findings-list').addEventListener('keydown', event => {
  if (event.key === 'Enter' || event.key === ' ') { const item = event.target.closest('.finding'); if (item) { event.preventDefault(); item.click(); } }
});
