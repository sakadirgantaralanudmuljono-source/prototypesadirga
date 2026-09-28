let notificationRefreshTimer = null;
let notificationFetchVersion = 0;

// Tetap berfungsi ketika penyimpanan browser diblokir/penuh.
const browserStorage = (() => {
  const memory = new Map();
  return {
    getItem(key) {
      if (memory.has(key)) return memory.get(key);
      try { return window.localStorage.getItem(key); } catch (_) { return null; }
    },
    setItem(key, value) {
      memory.set(key, String(value));
      try { window.localStorage.setItem(key, String(value)); } catch (_) {}
    },
    removeItem(key) {
      memory.set(key, null);
      try { window.localStorage.removeItem(key); } catch (_) {}
    }
  };
})();

const UI_VERSION = '3.8.2';
let systemLoadingTimer = null;
let dataLoadGeneration = 0;
let moduleLoadRequests = {};
let dashboardLoadRequest = null;
let modalLoadGeneration = 0;
let xlsxLoadPromise = null;
let searchTimer = null;
let batchIzinLoadGeneration = 0;
let penilaianLoadRequest = null;
let pendingCustomDialog = null;

const EXCEL_TEMPLATE_COLUMNS = {
  Anggota: [
    'ID','NTA','Nama','JenisKelamin','TempatLahir','TanggalLahir','Alamat',
    'Telepon','SekolahInstansi','Krida','Jabatan','Status','TanggalGabung','TanggalPelantikan'
  ],
  Kegiatan: [
    'ID','NamaKegiatan','Tanggal','Jenis','Lokasi','PenanggungJawab','Status','Keterangan'
  ],
  Absensi: [
    'ID','KegiatanID','KegiatanNama','Tanggal','AnggotaID','NamaAnggota','NTA',
    'StatusKehadiran','Catatan','Metode'
  ],
  Kas: ['ID','Tanggal','Jenis','Kategori','Keterangan','Nominal','Petugas','NoBukti','KegiatanID'],
  Inventaris: [
    'ID','KodeBarang','NamaBarang','Kategori','Jumlah','Kondisi','Lokasi',
    'PenanggungJawab','Keterangan'
  ],
  Surat: [
    'ID','NomorSurat','Tanggal','Jenis','Perihal','AsalTujuan','Status','LinkFile','Keterangan'
  ],
  Pengurus: [
    'ID','AnggotaID','NTA','Nama','Jabatan','Bidang','Periode','Urutan','Status'
  ],
  Users: ['ID','Username','Nama','Role','Status','AnggotaID','Password']
};


const MODULE_NAMES = [
  'anggota','kegiatan','absensi','kas','inventaris',
  'kegiatanInventaris','surat','pengurus','users'
];

// Hanya modul ini yang memengaruhi KPI/ringkasan Dashboard.
// Pengurus dan Users tidak perlu membuat Dashboard kotor.
const DASHBOARD_DATA_MODULES = new Set([
  'anggota','kegiatan','absensi','kas','inventaris','kegiatanInventaris','surat'
]);


// Perubahan pada tiga sumber ini dapat mengubah skor bulanan.
const PENILAIAN_SOURCE_MODULES = new Set(['anggota','kegiatan','absensi']);

const VIEW_MODULE_DEPENDENCIES = {
  anggota: ['anggota'],
  kegiatan: ['kegiatan'],
  absensi: ['absensi'],
  kas: ['kas','kegiatan'],
  inventaris: ['inventaris','kegiatanInventaris'],
  surat: ['surat'],
  pengurus: ['pengurus'],
  users: ['users']
};

const FORM_MODULE_DEPENDENCIES = {
  anggota: ['anggota'],
  kegiatan: ['kegiatan','pengurus'],
  absensi: ['absensi','kegiatan','anggota'],
  kas: ['kas','kegiatan'],
  inventaris: ['inventaris'],
  surat: ['surat'],
  pengurus: ['pengurus','anggota'],
  users: ['users']
};

const state = {
  token: browserStorage.getItem('saka_v2_token') || '',
  user: null,
  permissions: {},
  userMemberOptions: [],
  userMemberOptionsLoaded: false,
  userMemberOptionsRequest: null,
  modules: createModuleState(),
  dashboardMeta: createDashboardState(),
  penilaian: createPenilaianState(),
  data: {
    anggota: [],
    kegiatan: [],
    absensi: [],
    kas: [],
    inventaris: [],
    surat: [],
    pengurus: [],
    kegiatanInventaris: [],
    users: [],
    dashboard: {}
  },
  view: 'dashboard',
  search: '',
  pagination: {
    anggota: { page: 1, pageSize: 10 },
    kegiatan: { page: 1, pageSize: 10 },
    absensi: { page: 1, pageSize: 10, mode: 'rows' },
    penilaian: { page: 1, pageSize: 10 }
  },
  ui: {
    sidebarCollapsed: browserStorage.getItem('saka_sidebar_collapsed') === '1',
    displayMode: normalizeDisplayModePreference(browserStorage.getItem('saka_display_mode'))
  }
};

document.addEventListener('DOMContentLoaded', boot);


document.addEventListener('DOMContentLoaded', initPwa);

let deferredPwaInstallPrompt = null;

function isPwaStandalone() {
  return window.matchMedia('(display-mode: standalone)').matches ||
    window.navigator.standalone === true;
}

function isIosDevice() {
  return /iphone|ipad|ipod/i.test(window.navigator.userAgent || '');
}

function isSafariBrowser() {
  const ua = String(window.navigator.userAgent || '');
  return /safari/i.test(ua) && !/chrome|crios|fxios|edgios|opr\//i.test(ua);
}

function updatePwaInstallButton() {
  const button = document.getElementById('pwaInstallSettingButton');
  if (!button) return;
  const iosManualInstall = isIosDevice() && isSafariBrowser() && !isPwaStandalone();
  button.hidden = isPwaStandalone() || (!deferredPwaInstallPrompt && !iosManualInstall);
}

function updatePwaConnectivity() {
  const banner = document.getElementById('pwaOfflineBanner');
  if (!banner) return;
  banner.hidden = window.navigator.onLine;
}

async function initPwa() {
  updatePwaConnectivity();
  updatePwaInstallButton();
  window.addEventListener('online', () => {
    updatePwaConnectivity();
    if (typeof showToast === 'function') showToast('Koneksi internet kembali aktif.', 'success');
  });
  window.addEventListener('offline', updatePwaConnectivity);

  window.addEventListener('beforeinstallprompt', event => {
    event.preventDefault();
    deferredPwaInstallPrompt = event;
    updatePwaInstallButton();
  });

  window.addEventListener('appinstalled', () => {
    deferredPwaInstallPrompt = null;
    updatePwaInstallButton();
    if (typeof showToast === 'function') showToast('SAKA Dirgantara berhasil diinstal.', 'success');
  });

  if ('serviceWorker' in navigator) {
    try {
      await navigator.serviceWorker.register('/sw.js', { scope: '/' });
    } catch (error) {
      console.warn('Service worker gagal didaftarkan:', error);
    }
  }
}

async function installPwa() {
  if (isPwaStandalone()) return;

  if (deferredPwaInstallPrompt) {
    const prompt = deferredPwaInstallPrompt;
    deferredPwaInstallPrompt = null;
    await prompt.prompt();
    try { await prompt.userChoice; } catch (_) {}
    updatePwaInstallButton();
    return;
  }

  if (isIosDevice()) {
    const html = `
      <div class="pwa-install-help">
        <div class="pwa-install-help-icon"><span class="material-symbols-rounded">install_mobile</span></div>
        <h3>Instal SAKA Dirgantara</h3>
        <p>Di iPhone/iPad, instalasi dilakukan melalui menu Safari.</p>
        <ol>
          <li>Tekan tombol <strong>Share</strong> pada Safari.</li>
          <li>Pilih <strong>Add to Home Screen</strong> / <strong>Tambahkan ke Layar Utama</strong>.</li>
          <li>Tekan <strong>Add</strong> / <strong>Tambah</strong>.</li>
        </ol>
        <p class="module-period-note">Setelah terpasang, aplikasi dapat dibuka dari layar utama dalam mode standalone.</p>
      </div>`;
    if (typeof openCustomModal === 'function') {
      openCustomModal('Instal Aplikasi', 'Tambahkan SAKA Dirgantara ke layar utama.', html, true);
    } else {
      alert('Safari: tekan Share → Add to Home Screen → Add.');
    }
  }
}


/* =====================================================
   BOOT
===================================================== */

async function boot() {
  migrateResponsivePreference();
  setupGlobalUi();

  // Vercel tidak menjalankan templating Apps Script. Parameter public route
  // dibaca langsung dari URL; dataset dipertahankan sebagai fallback legacy.
  const query = new URLSearchParams(window.location.search);
  const initialCheckinCode =
    String(query.get('checkin') || (document.body && document.body.dataset.initialCheckinCode) || '').trim();
  const initialIzinCode =
    String(query.get('izin') || (document.body && document.body.dataset.initialIzinCode) || '').trim();

  if (initialIzinCode) {
    renderPublicIzin(initialIzinCode);
    return;
  }

  if (initialCheckinCode) {
    renderPublicCheckin(initialCheckinCode);
    return;
  }

  if (!state.token) {
    renderLogin();
    return;
  }

  try {
    await enterApp();
  } catch (error) {
    clearSession();
    renderLogin();
  }
}


/* =====================================================
   SERVER
===================================================== */

/* Activity status reflects callbacks, never a simulated server percentage. */
const activityState = { items: [], serial: 0, timer: null, hideTimer: null, cycle: null, pendingWrites: new Map() };

function activityLabel(method) {
  const labels = {
    login: 'Memeriksa akun', logout: 'Mengakhiri sesi',
    getDashboardData: 'Memuat dashboard', getModulesData: 'Memuat data modul',
    getRolePermissions: 'Memuat hak akses', saveRolePermissions: 'Menyimpan hak akses',
    checkSystemStructure: 'Memeriksa struktur sistem',
    updateSpreadsheetStructure: 'Memperbarui struktur sistem', safeResetSystem: 'Mereset sistem',
    getPublicCheckinData: 'Memuat formulir absensi', publicCheckin: 'Mengirim absensi',
    getAttendanceGeofenceSettings: 'Memuat lokasi absensi', saveAttendanceGeofenceSettings: 'Menyimpan titik absensi',
    memberAttendanceCheckin: 'Memverifikasi lokasi & absensi',
    getPublicIzinData: 'Memuat formulir izin', publicSubmitIzin: 'Mengirim pengajuan izin',
    transitionKegiatanStatus: 'Mengubah status kegiatan',
    importExcelDatabaseBatch: 'Memproses batch impor Excel',
    finishExcelDatabaseImport: 'Mencatat hasil impor',
    getKegiatanDokumentasiPreview: 'Memuat pratinjau dokumentasi',
    getIzinBuktiPreview: 'Memuat bukti izin', getPenilaianPdfPayload: 'Menyiapkan berkas PDF'
  };
  if (labels[method]) return labels[method];
  const actions = { get: 'Memuat', save: 'Menyimpan', delete: 'Menghapus', upload: 'Mengunggah',
    update: 'Memperbarui', generate: 'Membuat', verify: 'Memverifikasi', close: 'Menutup' };
  const match = String(method).match(/^(get|save|delete|upload|update|generate|verify|close)(.*)$/);
  if (!match) return 'Memproses permintaan';
  return actions[match[1]] + ' ' + match[2].replace(/([a-z])([A-Z])/g, '$1 $2').replace(/Pdf/g, 'PDF');
}

function ensureActivityPanel() {
  let bar = document.getElementById('activityProgress');
  if (bar) return bar;
  bar = document.createElement('div');
  bar.id = 'activityProgress';
  bar.className = 'activity-progress';
  bar.hidden = true;
  bar.innerHTML = '<div class="activity-progress-track" role="progressbar" aria-label="Sedang memproses"><span></span></div><span class="activity-sr" role="status" aria-live="polite"></span>';
  document.body.appendChild(bar);
  return bar;
}

function renderActivities() {
  const bar = ensureActivityPanel();
  const active = activityState.items;
  const label = active.length === 1 ? active[0].label : active.length + ' proses berjalan';
  bar.hidden = false;
  bar.className = 'activity-progress is-running';
  bar.querySelector('[role="progressbar"]').setAttribute('aria-label', label);
  bar.querySelector('[role="progressbar"]').setAttribute('aria-busy', 'true');
  bar.querySelector('[role="status"]').textContent = label;
  bar.title = label;
}

function startActivity(label) {
  clearTimeout(activityState.timer);
  clearTimeout(activityState.hideTimer);
  activityState.timer = null;
  if (!activityState.cycle) activityState.cycle = { label, status: 'success', detail: '', notified: 0 };
  const item = { id: ++activityState.serial, label };
  activityState.items.push(item);
  renderActivities();
  return item;
}

function finishActivity(item, status, detail = '') {
  if (!activityState.items.includes(item)) return;
  activityState.items = activityState.items.filter(row => row !== item);
  const rank = { success: 1, warning: 2, error: 3 };
  const cycle = activityState.cycle;
  if (rank[status] > rank[cycle.status]) {
    cycle.status = status;
    cycle.detail = detail;
    cycle.label = item.label;
  }
  if (activityState.items.length) { renderActivities(); return; }
  // Allow follow-up refreshes and existing feature-specific toasts to settle.
  activityState.timer = setTimeout(() => {
    activityState.timer = null;
    const bar = ensureActivityPanel();
    bar.className = 'activity-progress is-' + cycle.status;
    bar.querySelector('[role="progressbar"]').setAttribute('aria-busy', 'false');
    bar.querySelector('[role="status"]').textContent = cycle.status === 'error' ? 'Proses gagal' : 'Proses selesai';
    if (cycle.notified < rank[cycle.status]) {
      const message = cycle.status === 'error' ? (cycle.detail || 'Proses gagal. Silakan periksa kembali.')
        : cycle.status === 'warning' ? 'Proses selesai dengan catatan. Periksa hasilnya.'
        : cycle.label === 'Memuat dashboard' ? 'Dashboard berhasil dimuat.'
        : cycle.label === 'Memuat data modul' ? 'Data berhasil dimuat.'
        : cycle.label + ': selesai.';
      showToast(message, cycle.status);
    }
    activityState.cycle = null;
    activityState.hideTimer = setTimeout(() => { bar.hidden = true; }, 350);
  }, 220);
}

async function trackLocalActivity(label, work) {
  const item = startActivity(label);
  try {
    // Yield once so the browser can paint before local processing starts.
    await new Promise(resolve => setTimeout(resolve, 0));
    const result = await work();
    finishActivity(item, 'success');
    return result;
  } catch (error) {
    finishActivity(item, 'error', error.message || 'Proses lokal gagal.');
    throw error;
  }
}

function serverCall(method, ...args) {
  if (!method || typeof method !== 'string') {
    return Promise.reject(new Error('Metode server tidak valid.'));
  }

  const mutation = !/^(get|check)/.test(method);
  let key = null;
  let body = '';

  try {
    key = mutation ? JSON.stringify([method, args]) : null;
    body = JSON.stringify({ method, args });
  } catch (_) {
    return Promise.reject(new Error('Data permintaan tidak dapat dikirim.'));
  }

  // Vercel Functions memiliki batas payload. Fail fast agar pengguna mendapat
  // pesan yang jelas, bukan respons 413 generik dari platform.
  const payloadBytes = typeof TextEncoder !== 'undefined'
    ? new TextEncoder().encode(body).length
    : body.length;
  if (payloadBytes > 4 * 1024 * 1024) {
    return Promise.reject(new Error(
      'Ukuran permintaan terlalu besar untuk gateway Vercel. Kompres file atau gunakan file yang lebih kecil.'
    ));
  }

  if (key && activityState.pendingWrites.has(key)) {
    return activityState.pendingWrites.get(key);
  }

  const item = startActivity(activityLabel(method));

  const request = fetch('/api/rpc', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    cache: 'no-store',
    body
  }).then(async response => {
    const raw = await response.text();
    let envelope = {};

    try {
      envelope = raw ? JSON.parse(raw) : {};
    } catch (_) {
      throw new Error('Gateway mengembalikan respons yang tidak valid.');
    }

    if (!response.ok || envelope.success !== true) {
      const diagnostic = envelope.diagnosticCode ? ` [${envelope.diagnosticCode}]` : '';
      throw new Error((envelope.message || `Permintaan gagal (${response.status}).`) + diagnostic);
    }

    return envelope.data;
  }).then(result => {
    const partial = result && (result.success === false || Number(result.failed) > 0);
    finishActivity(item, partial ? 'warning' : 'success', partial ? 'Periksa laporan hasil proses.' : '');
    if (mutation && method !== 'markActionNotificationsRead') scheduleNotificationRefresh();
    return result;
  }, error => {
    finishActivity(item, 'error', error.message || 'Permintaan gagal.');
    throw error;
  }).finally(() => {
    if (key) activityState.pendingWrites.delete(key);
  });

  if (key) activityState.pendingWrites.set(key, request);
  return request;
}


/* =====================================================
   LOGIN
===================================================== */


function renderLogin() {
  const lastUsername = browserStorage.getItem('saka_last_username') || '';
  document.getElementById('appRoot').innerHTML = `
    <div class="auth-shell auth-shell-modern">

      <section class="auth-visual">
        <div class="visual-brand">
          <div class="brand-mark">
            <span class="material-symbols-rounded">flight</span>
          </div>

          <div class="brand-copy">
            <strong>SAKA</strong>
            <small>DIRGANTARA</small>
          </div>
        </div>

        <div class="visual-copy">
          <span class="eyebrow">Modern Organization Workspace</span>

          <h1>
            Kelola organisasi dengan
            <span>alur yang lebih cepat & rapi.</span>
          </h1>

          <p>
            Satu sistem untuk anggota, kegiatan, absensi, inventaris, surat, dan laporan.
            Dibuat agar lebih ringan, lebih mudah dipahami, dan nyaman dipakai dari desktop maupun mobile.
          </p>

          <div class="auth-feature-chips" aria-label="Fitur utama">
            <span>Absensi GPS</span>
            <span>Inventaris</span>
            <span>Keuangan</span>
            <span>PWA</span>
          </div>
        </div>

        <div class="auth-metrics" aria-label="Keunggulan sistem">
          <div class="auth-metric-card">
            <strong>Lebih cepat</strong>
            <small>Refresh data tanpa terasa memutus alur kerja utama.</small>
          </div>
          <div class="auth-metric-card">
            <strong>Lebih rapi</strong>
            <small>UI clean, fokus pada data yang paling sering dipakai.</small>
          </div>
          <div class="auth-metric-card">
            <strong>Lebih aman</strong>
            <small>Akses per role tetap terjaga untuk admin, pengurus, dan anggota.</small>
          </div>
        </div>

        <div class="auth-footer-note">
          SAKA Dirgantara Management System • Version ${UI_VERSION}
        </div>
      </section>

      <section class="auth-panel">
        <div class="auth-card auth-card-modern">

          <div class="mobile-logo visual-brand">
            <div class="brand-mark">
              <span class="material-symbols-rounded">flight</span>
            </div>

            <div class="brand-copy" style="color:#071827;">
              <strong>SAKA</strong>
              <small>DIRGANTARA</small>
            </div>
          </div>

          <div class="auth-card-head">
            <span class="auth-badge">Akses Sistem</span>
            <h2>Masuk ke dashboard</h2>
            <p>
              Gunakan akun yang telah didaftarkan administrator. Sistem akan mengingat username terakhir di perangkat ini.
            </p>
          </div>

          <form class="auth-form" onsubmit="handleLogin(event)">
            <div class="auth-field">
              <label for="loginUsername">USERNAME</label>
              <div class="input-shell">
                <span class="material-symbols-rounded auth-input-icon" aria-hidden="true">person</span>
                <input
                  class="auth-input"
                  id="loginUsername"
                  name="username"
                  autocomplete="username"
                  required
                  autofocus
                  value="${escapeHtml(lastUsername)}"
                  placeholder="Masukkan username"
                >
              </div>
            </div>

            <div class="auth-field">
              <div class="auth-field-top">
                <label for="loginPassword">PASSWORD</label>
                <small>Ketik sandi akun Anda</small>
              </div>
              <div class="input-shell input-shell-password">
                <span class="material-symbols-rounded auth-input-icon" aria-hidden="true">lock</span>
                <input
                  class="auth-input auth-input-password"
                  id="loginPassword"
                  name="password"
                  type="password"
                  autocomplete="current-password"
                  required
                  placeholder="Masukkan password"
                >
                <button
                  class="auth-password-toggle"
                  id="loginPasswordToggle"
                  type="button"
                  onclick="togglePasswordVisibility('loginPassword','loginPasswordToggle')"
                  aria-label="Tampilkan password"
                >
                  <span class="material-symbols-rounded" aria-hidden="true">visibility</span>
                </button>
              </div>
            </div>

            <div class="auth-helper-row">
              <div class="auth-helper-item">
                <span class="material-symbols-rounded" aria-hidden="true">schedule</span>
                Login cepat & responsif
              </div>
              <div class="auth-helper-item">
                <span class="material-symbols-rounded" aria-hidden="true">verified_user</span>
                Hak akses sesuai role
              </div>
            </div>

            <button
              id="loginButton"
              class="btn btn-primary auth-submit"
              type="submit"
            >
              <span class="material-symbols-rounded">login</span>
              Masuk ke Sistem
            </button>
          </form>

          <div class="login-note login-note-modern">
            <div>
              <strong>Setup awal</strong>
              <p>Jika database belum pernah dibuat, jalankan <strong>setupSystem_()</strong> satu kali dari Apps Script Editor.</p>
            </div>
            <span class="material-symbols-rounded" aria-hidden="true">rocket_launch</span>
          </div>

        </div>
      </section>
    </div>
  `;
}

function togglePasswordVisibility(inputId, toggleId) {
  const input = document.getElementById(inputId);
  const toggle = document.getElementById(toggleId);
  if (!input || !toggle) return;
  const icon = toggle.querySelector('.material-symbols-rounded');
  const isHidden = input.type === 'password';
  input.type = isHidden ? 'text' : 'password';
  toggle.setAttribute('aria-label', isHidden ? 'Sembunyikan password' : 'Tampilkan password');
  if (icon) icon.textContent = isHidden ? 'visibility_off' : 'visibility';
}


async function handleLogin(event) {
  event.preventDefault();

  const form = event.target;
  const button = document.getElementById('loginButton');
  const payload = Object.fromEntries(new FormData(form).entries());
  const username = String(payload.username || '').trim();

  if (!username || !String(payload.password || '').trim()) {
    showToast('Username dan password wajib diisi.', 'warning');
    return;
  }

  button.disabled = true;
  button.classList.add('is-loading');
  button.innerHTML = `
    <span class="material-symbols-rounded spin">progress_activity</span>
    Memeriksa akun...
  `;

  try {
    const result = await serverCall('login', payload);

    state.token = result.token;
    state.user = result.user;

    browserStorage.setItem('saka_v2_token', state.token);
    browserStorage.setItem('saka_last_username', username);

    await enterApp();

  } catch (error) {
    showToast(error.message, 'error');

  } finally {
    if (button) {
      button.disabled = false;
      button.classList.remove('is-loading');
      button.innerHTML = `
        <span class="material-symbols-rounded">login</span>
        Masuk ke Sistem
      `;
    }
  }
}

async function doLogout() {
  try {
    if (state.token) {
      await serverCall('logout', state.token);
    }
  } catch (_) {}

  clearSession();
  renderLogin();
}

function openBackendHealthCheck() {
  window.open('/api/health', '_blank', 'noopener,noreferrer');
}

function clearSession() {
  stopSystemLoading();
  state.token = '';
  state.user = null;
  state.permissions = {};
  state.userMemberOptions = [];
  state.userMemberOptionsLoaded = false;
  state.userMemberOptionsRequest = null;
  state.data = emptyAppData();
  state.modules = createModuleState();
  state.dashboardMeta = createDashboardState();
  state.penilaian = createPenilaianState();
  penilaianLoadRequest = null;
  state.view = 'dashboard';
  dataLoadGeneration++;
  moduleLoadRequests = {};
  dashboardLoadRequest = null;
  closeModal();
  browserStorage.removeItem('saka_v2_token');
}


/* =====================================================
   APP SHELL
===================================================== */

async function enterApp() {
  const token = state.token;
  const generation = ++dataLoadGeneration;
  state.view = 'dashboard';
  state.modules = createModuleState();
  state.dashboardMeta = createDashboardState();
  state.penilaian = createPenilaianState();
  moduleLoadRequests = {};
  dashboardLoadRequest = null;
  penilaianLoadRequest = null;
  state.data = emptyAppData();
  startSystemLoading();

  try {
    // Satu panggilan sekaligus memverifikasi sesi dan mengambil ringkasan.
    const data = await serverCall('getDashboardData', token);
    if (!isCurrentDataLoad(token, generation)) return;
    Object.assign(state.data, data);
    state.user = data.user;
    await loadRolePermissions();
    markDashboardClean();
    renderAppShell();
    renderView();
  } catch (error) {
    if (!isCurrentDataLoad(token, generation)) return;
    if (String(error.message).toLowerCase().includes('sesi')) {
      clearSession();
      renderLogin();
      return;
    }
    const backendDiagnostic = /\[GAS_|Apps Script|gateway|backend/i.test(String(error.message || ''));
    document.getElementById('appRoot').innerHTML = `
      <div class="setup-card backend-error-card">
        <div class="setup-icon backend-error-icon">
          <span class="material-symbols-rounded">cloud_off</span>
        </div>
        <span class="backend-error-badge">Koneksi Backend</span>
        <h2>Dashboard belum berhasil dimuat</h2>
        <p>${escapeHtml(error.message)}</p>
        <div class="backend-error-actions">
          <button class="btn btn-primary" type="button" onclick="enterApp()">
            <span class="material-symbols-rounded">refresh</span>Coba Lagi
          </button>
          ${backendDiagnostic ? `<button class="btn btn-light" type="button" onclick="openBackendHealthCheck()">
            <span class="material-symbols-rounded">monitor_heart</span>Tes Koneksi Backend
          </button>` : ''}
          <button class="btn btn-light" type="button" onclick="doLogout()">
            <span class="material-symbols-rounded">logout</span>Kembali ke Login
          </button>
        </div>
        ${backendDiagnostic ? `<p class="backend-error-help">Jika tes koneksi mengembalikan HTML atau meminta login Google, perbarui deployment Apps Script ke Web App production dan pastikan aksesnya dapat dipanggil tanpa login.</p>` : ''}
      </div>`;
  } finally {
    if (isCurrentDataLoad(token, generation)) stopSystemLoading();
  }
}

async function loadRolePermissions() {
  state.permissions = {};
  if (!state.user || String(state.user.Role).toUpperCase() !== 'PENGURUS') return;
  try {
    const payload = await serverCall('getRolePermissions', state.token);
    (payload.rows || []).forEach(row => {
      if (String(row.Role).toUpperCase() !== 'PENGURUS') return;
      state.permissions[String(row.Module)] = {
        view: String(row.CanView).toUpperCase() === 'TRUE',
        create: String(row.CanCreate).toUpperCase() === 'TRUE',
        edit: String(row.CanEdit).toUpperCase() === 'TRUE',
        delete: String(row.CanDelete).toUpperCase() === 'TRUE'
      };
    });
  } catch (_) {}
}

function renderAppShell() {
  const isAdmin = isRole('ADMIN');
  const isMember = isRole('ANGGOTA');
  const initials = getInitials(state.user && state.user.Nama);
  const collapsedClass = state.ui && state.ui.sidebarCollapsed ? ' sidebar-collapsed' : '';

  document.getElementById('appRoot').innerHTML = `
    <div class="app${collapsedClass}${isMember ? ' member-app' : ''}" id="appShell">

      <aside class="sidebar" id="sidebar" aria-label="Navigasi utama">
        <div class="sidebar-head">
          <button class="brand brand-button" type="button" onclick="switchViewByName('dashboard')" aria-label="Buka dashboard">
            <div class="brand-mark">
              <span class="material-symbols-rounded">flight</span>
            </div>
            <div class="brand-copy">
              <strong>SAKA</strong>
              <small>DIRGANTARA</small>
            </div>
          </button>

          <button class="sidebar-mobile-close" type="button" onclick="closeSidebar()" aria-label="Tutup navigasi">
            <span class="material-symbols-rounded">close</span>
          </button>
        </div>

        <div class="sidebar-nav-scroll">
          <div class="nav-section">
            <div class="nav-label">Utama</div>
            <nav class="nav">
              ${navButton('dashboard', 'space_dashboard', isMember ? 'Beranda' : 'Dashboard')}
              ${isMember ? navButton('member-profile', 'person', 'Profil Saya') : ''}
              ${isMember ? navButton('member-attendance', 'fact_check', 'Kehadiran Saya') : ''}
              ${isMember ? navButton('member-skk', 'workspace_premium', 'Progres SKK') : ''}
              ${!isMember && canViewModule('anggota') ? navButton('anggota', 'groups', 'Anggota') : ''}
              ${!isMember && canViewModule('kegiatan') ? navButton('kegiatan', 'calendar_month', 'Kegiatan') : ''}
              ${!isMember && canViewModule('absensi') ? navButton('absensi', 'fact_check', 'Absensi') : ''}
              ${!isMember && canViewModule('penilaian') ? navButton('penilaian', 'workspace_premium', 'Penilaian') : ''}
            </nav>
          </div>

          ${!isMember ? `<div class="nav-section">
            <div class="nav-label">Operasional</div>
            <nav class="nav">
              ${canViewModule('kas') ? navButton('kas', 'account_balance_wallet', 'Kas Organisasi') : ''}
              ${canViewModule('inventaris') ? navButton('inventaris', 'inventory_2', 'Inventaris') : ''}
              ${canViewModule('surat') ? navButton('surat', 'mail', 'Surat') : ''}
              ${canViewModule('pengurus') ? navButton('pengurus', 'account_tree', 'Struktur Pengurus') : ''}
              ${isAdmin ? navButton('users', 'admin_panel_settings', 'Pengguna') : ''}
            </nav>
          </div>` : ''}

          <div class="nav-section">
            <div class="nav-label">Aplikasi</div>
            <nav class="nav">
              ${navButton('settings', 'settings', 'Pengaturan')}
            </nav>
          </div>
        </div>

        <div class="sidebar-bottom">
          <button class="sidebar-collapse" type="button" onclick="toggleSidebarCompact()" title="Ciutkan sidebar" aria-label="Ciutkan sidebar">
            <span class="material-symbols-rounded">dock_to_right</span>
            <span>Ciutkan menu</span>
          </button>

          <div class="sidebar-user">
            <div class="user-avatar">${escapeHtml(initials)}</div>
            <div class="sidebar-user-copy">
              <strong>${escapeHtml(state.user.Nama)}</strong>
              <small>${escapeHtml(state.user.Role)}</small>
            </div>
            <button class="sidebar-logout" type="button" onclick="doLogout()" title="Keluar" aria-label="Keluar dari sistem">
              <span class="material-symbols-rounded">logout</span>
            </button>
          </div>
        </div>
      </aside>

      <button class="sidebar-scrim" id="sidebarScrim" type="button" onclick="closeSidebar()" aria-label="Tutup navigasi"></button>

      <main class="main">
        <header class="topbar">
          <div class="topbar-left">
            <button class="mobile-menu icon-button" type="button" onclick="toggleSidebar()" aria-label="Buka navigasi">
              <span class="material-symbols-rounded">menu</span>
            </button>

            <div class="topbar-title-wrap">
              <span class="topbar-kicker">SAKA Dirgantara</span>
              <h1 id="pageTitle">${isMember ? 'Beranda' : 'Dashboard'}</h1>
            </div>
          </div>

          <div class="topbar-right">
            ${!isMember ? `<button id="notificationBell" class="topbar-action icon-button notification-bell" type="button" onclick="openActionCenter()" title="Notification & Action Center" aria-label="Buka Notification & Action Center">
              <span class="material-symbols-rounded">notifications</span><span id="notificationBadge" class="notification-count" hidden></span>
            </button>` : ''}
            <button class="topbar-action icon-button" data-refresh-button="true" type="button" onclick="refreshData()" title="Muat ulang data" aria-label="Muat ulang data">
              <span class="material-symbols-rounded">refresh</span>
            </button>

            <div class="system-chip">
              <span class="system-dot"></span>
              <span>Sistem Aktif</span>
            </div>

            <div class="profile-chip" title="${escapeHtml(state.user.Role)}">
              <div class="profile-avatar">${escapeHtml(initials)}</div>
              <div class="profile-copy">
                <strong>${escapeHtml(state.user.Nama)}</strong>
                <small>${escapeHtml(state.user.Role)}</small>
              </div>
            </div>
          </div>
        </header>

        <section class="content" id="content">
          ${loadingHtml()}
        </section>
      </main>

      <nav class="mobile-bottom-nav" aria-label="Navigasi mobile">
        ${mobileNavButton('dashboard', 'space_dashboard', 'Beranda')}
        ${isMember ? mobileNavButton('member-attendance', 'location_on', 'Absen') : mobileNavButton('anggota', 'groups', 'Anggota')}
        ${isMember ? mobileNavButton('member-profile', 'person', 'Profil') : mobileNavButton('kegiatan', 'calendar_month', 'Kegiatan')}
        ${isMember ? mobileNavButton('member-skk', 'workspace_premium', 'SKK') : mobileNavButton('absensi', 'fact_check', 'Absensi')}
        ${isMember ? mobileNavButton('settings', 'settings', 'Setelan') : `<button class="mobile-nav-item mobile-more-item" type="button" data-mobile-more="true" onclick="toggleSidebar()">
          <span class="material-symbols-rounded">apps</span>
          <span>Lainnya</span>
        </button>`}
      </nav>

    </div>
  `;

  if (!isMember) scheduleNotificationRefresh();
  syncNavigationActive(state.view);
  requestAnimationFrame(() => {
    applyResponsiveMode();
    repairScrollLockState();
  });
}


function setRefreshButtonLoading(isLoading) {
  const button = document.querySelector('[data-refresh-button="true"]');
  if (!button) return;
  button.disabled = !!isLoading;
  button.classList.toggle('is-loading', !!isLoading);
  button.setAttribute('aria-busy', isLoading ? 'true' : 'false');
  const icon = button.querySelector('.material-symbols-rounded');
  if (icon) icon.textContent = isLoading ? 'autorenew' : 'refresh';
}

function navButton(view, icon, label) {
  return `
    <button
      class="nav-item ${view === state.view ? 'active' : ''}"
      data-view="${view}"
      type="button"
      title="${escapeHtml(label)}"
      onclick="switchView('${view}', this)"
    >
      <span class="nav-icon material-symbols-rounded">${icon}</span>
      <span class="nav-text">${escapeHtml(label)}</span>
    </button>
  `;
}

function mobileNavButton(view, icon, label) {
  return `
    <button
      class="mobile-nav-item ${view === state.view ? 'active' : ''}"
      data-view="${view}"
      type="button"
      onclick="switchView('${view}', this)"
    >
      <span class="material-symbols-rounded">${icon}</span>
      <span>${escapeHtml(label)}</span>
    </button>
  `;
}

function toggleSidebar() {
  const sidebar = document.getElementById('sidebar');
  if (!sidebar) return;

  if (isMobileLikeDevice()) {
    const opening = !sidebar.classList.contains('open');
    sidebar.classList.toggle('open', opening);
    document.body.classList.toggle('mobile-nav-open', opening);
    return;
  }

  toggleSidebarCompact();
}

function closeSidebar() {
  const sidebar = document.getElementById('sidebar');
  if (sidebar) sidebar.classList.remove('open');
  if (document.body) document.body.classList.remove('mobile-nav-open');
  repairScrollLockState();
}

function toggleSidebarCompact() {
  if (isMobileLikeDevice()) return;
  const app = document.getElementById('appShell');
  if (!app) return;
  const collapsed = !app.classList.contains('sidebar-collapsed');
  app.classList.toggle('sidebar-collapsed', collapsed);
  state.ui.sidebarCollapsed = collapsed;
  browserStorage.setItem('saka_sidebar_collapsed', collapsed ? '1' : '0');
}

function syncNavigationActive(view) {
  document.querySelectorAll('.nav-item, .mobile-nav-item[data-view]').forEach(item => {
    item.classList.toggle('active', item.dataset.view === view);
  });

  const more = document.querySelector('.mobile-more-item');
  if (more) {
    const primary = ['dashboard','anggota','kegiatan','absensi'];
    more.classList.toggle('active', primary.indexOf(view) === -1);
  }
}

function getInitials(name) {
  const parts = String(name || 'SAKA').trim().split(/\s+/).filter(Boolean);
  return parts.slice(0, 2).map(part => part.charAt(0).toUpperCase()).join('') || 'SA';
}


/* =====================================================
   DATA
===================================================== */

function emptyAppData() {
  return {
    anggota: [], kegiatan: [], absensi: [], kas: [], inventaris: [], kegiatanInventaris: [],
    surat: [], pengurus: [], users: [], dashboard: {}, memberPortal: null
  };
}

function createModuleState() {
  return MODULE_NAMES.reduce((out, name) => {
    out[name] = { loaded: false, dirty: true };
    return out;
  }, {});
}

function createDashboardState() {
  return { loaded: false, dirty: true, loadedAt: 0 };
}


function currentPenilaianPeriod() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

function createPenilaianState() {
  return {
    period: currentPenilaianPeriod(),
    loaded: false,
    dirty: true,
    data: null,
    loadedAt: 0,
    sharePayload: null
  };
}

function markPenilaianDirty() {
  if (!state.penilaian) state.penilaian = createPenilaianState();
  state.penilaian.dirty = true;
  state.penilaian.sharePayload = null;
}

function markPenilaianClean(data) {
  if (!state.penilaian) state.penilaian = createPenilaianState();
  state.penilaian.loaded = true;
  state.penilaian.dirty = false;
  state.penilaian.data = data || null;
  state.penilaian.loadedAt = Date.now();
}

function isPenilaianReady() {
  const meta = state.penilaian;
  return !!(meta && meta.loaded && !meta.dirty && meta.data && meta.data.period === meta.period);
}

async function ensurePenilaianLoaded(force = false) {
  if (!state.token) return false;
  if (!force && isPenilaianReady()) return true;
  if (penilaianLoadRequest) return penilaianLoadRequest;

  const token = state.token;
  const generation = dataLoadGeneration;
  const period = String(state.penilaian && state.penilaian.period || currentPenilaianPeriod());
  const request = serverCall('getPenilaianBulanan', token, period).then(data => {
    if (!isCurrentDataLoad(token, generation)) return false;
    if (!state.penilaian || state.penilaian.period !== period) return false;
    markPenilaianClean(data);
    return true;
  });

  penilaianLoadRequest = request;
  try {
    return await request;
  } finally {
    if (penilaianLoadRequest === request) penilaianLoadRequest = null;
  }
}

function markDashboardDirty() {
  if (!state.dashboardMeta) state.dashboardMeta = createDashboardState();
  state.dashboardMeta.dirty = true;
}

function markDashboardClean() {
  if (!state.dashboardMeta) state.dashboardMeta = createDashboardState();
  state.dashboardMeta.loaded = true;
  state.dashboardMeta.dirty = false;
  state.dashboardMeta.loadedAt = Date.now();
}

function isDashboardReady() {
  const meta = state.dashboardMeta;
  return !!(meta && meta.loaded && !meta.dirty);
}

function normalizeModuleNames(names) {
  return [...new Set((Array.isArray(names) ? names : [names])
    .map(name => String(name || '').trim())
    .filter(name => MODULE_NAMES.includes(name))
  )];
}

function getViewModules(view) {
  return normalizeModuleNames(VIEW_MODULE_DEPENDENCIES[view] || []);
}

function getFormModules(type) {
  return normalizeModuleNames(FORM_MODULE_DEPENDENCIES[type] || [type]);
}

function areModulesReady(names) {
  return normalizeModuleNames(names).every(name => {
    const meta = state.modules && state.modules[name];
    return !!(meta && meta.loaded && !meta.dirty);
  });
}

function markModuleDirty(name) {
  if (state.modules && state.modules[name]) state.modules[name].dirty = true;
}

function markAllModulesDirty() {
  MODULE_NAMES.forEach(markModuleDirty);
}

function markModuleClean(name) {
  if (!state.modules || !state.modules[name]) return;
  state.modules[name].loaded = true;
  state.modules[name].dirty = false;
}

function isCurrentDataLoad(token, generation) {
  return state.token === token && dataLoadGeneration === generation;
}

function handleDataLoadError(error) {
  if (String(error.message).toLowerCase().includes('sesi')) {
    clearSession();
    renderLogin();
    return;
  }
  showToast(error.message, 'error');
  renderError(error.message);
}

async function ensureModulesLoaded(names, force = false) {
  const denied = normalizeModuleNames(names).filter(name =>
    !canViewModule(name === 'kegiatanInventaris' ? 'inventaris' : name)
  );
  if (denied.length) {
    throw new Error('Fitur ini membutuhkan izin lihat modul: ' + denied.join(', ') + '. Hubungi administrator.');
  }
  const wanted = normalizeModuleNames(names);
  if (!wanted.length) return true;

  // Tunggu request modul yang sama jika masih berjalan agar navigasi/form
  // yang dipicu hampir bersamaan tidak membuat pembacaan Sheets ganda.
  const waiting = [];
  const candidates = [];

  wanted.forEach(name => {
    const meta = state.modules[name];
    if (!force && meta.loaded && !meta.dirty) return;

    if (!force && moduleLoadRequests[name]) {
      waiting.push(moduleLoadRequests[name]);
    } else {
      candidates.push(name);
    }
  });

  if (waiting.length) {
    const waitingResult = await Promise.all(waiting);
    if (waitingResult.some(result => result === false)) return false;
  }

  const missing = candidates.filter(name => {
    const meta = state.modules[name];
    return force || !meta.loaded || meta.dirty;
  });

  if (!missing.length) return true;

  const token = state.token;
  const generation = dataLoadGeneration;
  const request = serverCall('getModulesData', token, missing).then(result => {
    if (!isCurrentDataLoad(token, generation)) return false;

    const modules = result && result.modules ? result.modules : {};
    missing.forEach(name => {
      if (!Object.prototype.hasOwnProperty.call(modules, name)) {
        throw new Error('Server tidak mengembalikan data modul: ' + name);
      }
      state.data[name] = Array.isArray(modules[name]) ? modules[name] : [];
      markModuleClean(name);
    });

    if (result && result.user) state.user = result.user;
    if (result && result.version) state.data.version = result.version;
    return true;
  });

  missing.forEach(name => {
    moduleLoadRequests[name] = request;
  });

  try {
    return await request;
  } finally {
    missing.forEach(name => {
      if (moduleLoadRequests[name] === request) delete moduleLoadRequests[name];
    });
  }
}

async function ensureDashboardLoaded(force = false) {
  if (!state.token) return false;
  if (!force && isDashboardReady()) return true;
  if (dashboardLoadRequest) return dashboardLoadRequest;

  const token = state.token;
  const generation = dataLoadGeneration;
  const request = serverCall('getDashboardData', token, force === true).then(data => {
    if (!isCurrentDataLoad(token, generation)) return false;
    Object.assign(state.data, data || {});
    state.user = (data && data.user) || state.user;
    markDashboardClean();
    return true;
  });

  dashboardLoadRequest = request;
  try {
    return await request;
  } finally {
    if (dashboardLoadRequest === request) dashboardLoadRequest = null;
  }
}


async function refreshData() {
  scheduleNotificationRefresh();
  const view = state.view;
  setRefreshButtonLoading(true);

  // Maintenance sudah memiliki Cek Struktur / Update Struktur sendiri.
  // Tombol refresh global cukup menggambar ulang halaman tanpa request Dashboard.
  if (view === 'maintenance' || view === 'settings') {
    try {
      renderSettings();
    } finally {
      setRefreshButtonLoading(false);
    }
    return;
  }

  const token = state.token;
  const generation = ++dataLoadGeneration;
  moduleLoadRequests = {};
  dashboardLoadRequest = null;
  penilaianLoadRequest = null;
  markAllModulesDirty();
  markDashboardDirty();
  markPenilaianDirty();
  showLoading({ preserve: true, label: 'Menyegarkan data...' });

  try {
    if (view === 'dashboard' || isRole('ANGGOTA')) {
      if (!await ensureDashboardLoaded(true)) return;
    } else if (view === 'penilaian') {
      if (!await ensurePenilaianLoaded(true)) return;
    } else {
      if (!await ensureModulesLoaded(getViewModules(view), true)) return;
    }

    if (isCurrentDataLoad(token, generation) && state.view === view) renderView();
  } catch (error) {
    if (isCurrentDataLoad(token, generation) && state.view === view) handleDataLoadError(error);
  } finally {
    hideLoadingOverlay();
    setRefreshButtonLoading(false);
  }
}

/*
 * Setelah CRUD, gunakan baris yang dikembalikan server untuk memperbarui
 * state lokal. Modul tetap dianggap sudah mutakhir sehingga tidak perlu
 * meminta ulang seluruh sheet setelah satu transaksi.
 */
function applyLocalMutation(type, row, id, deleted = false) {
  const collection = state.data[type];
  if (!Array.isArray(collection)) return;

  const targetId = String(id || (row && row.ID) || '');
  const index = collection.findIndex(item => String(item.ID) === targetId);

  if (deleted) {
    if (index >= 0) collection.splice(index, 1);
    markModuleClean(type);
    return;
  }

  if (!row || !targetId) return;
  // saveUser() juga mengembalikan objek persistence untuk kebutuhan server.
  // Jangan pernah memasukkan PasswordHash ke state/browser.
  if (type === 'users') {
    row = {
      ID: row.ID,
      Username: row.Username,
      Nama: row.Nama,
      Role: row.Role,
      Status: row.Status,
      AnggotaID: row.AnggotaID || '',
      DibuatPada: row.DibuatPada,
      DiubahPada: row.DiubahPada
    };
  }
  if (index >= 0) {
    collection[index] = Object.assign({}, collection[index], row);
  } else {
    collection.unshift(row);
  }
  markModuleClean(type);

  // Kegiatan adalah sumber nama/tanggal yang ditampilkan pada Absensi.
  // Jika Absensi sudah pernah dimuat, sinkronkan snapshot lokalnya.
  if (type === 'kegiatan' && state.modules.absensi.loaded) {
    state.data.absensi = (state.data.absensi || []).map(item => {
      if (String(item.KegiatanID || '') !== targetId) return item;
      return Object.assign({}, item, {
        KegiatanNama: row.NamaKegiatan || item.KegiatanNama,
        Tanggal: row.Tanggal || item.Tanggal
      });
    });
  }
}

function finishLocalMutation(type, row, id, deleted = false) {
  applyLocalMutation(type, row, id, deleted);

  // Jangan invalidasi Dashboard untuk data yang tidak dipakai KPI Dashboard.
  if (DASHBOARD_DATA_MODULES.has(type)) markDashboardDirty();

  // Rekap keaktifan dihitung ulang hanya bila sumber penilaiannya berubah.
  if (PENILAIAN_SOURCE_MODULES.has(type)) markPenilaianDirty();

  // Jika Dashboard sedang terlihat (misalnya Quick Action), renderView() akan
  // mengambil ringkasan baru satu kali hanya bila data Dashboard memang berubah.
  if (state.view === type || (state.view === 'dashboard' && DASHBOARD_DATA_MODULES.has(type))) {
    renderView();
  }
}


/* =====================================================
   NAVIGATION
===================================================== */

function switchView(view, element) {
  clearTimeout(searchTimer);
  state.view = view;
  state.search = '';
  if (state.pagination && state.pagination[view]) {
    state.pagination[view].page = 1;
  }

  syncNavigationActive(view);

  const titles = {
    dashboard: 'Dashboard',
    anggota: 'Data Anggota',
    kegiatan: 'Kegiatan',
    absensi: 'Absensi',
    penilaian: 'Penilaian Keaktifan',
    kas: 'Kas Organisasi',
    inventaris: 'Inventaris',
    surat: 'Administrasi Surat',
    pengurus: 'Struktur Pengurus',
    users: 'Manajemen Pengguna',
    maintenance: 'Pengaturan',
    settings: 'Pengaturan',
    'member-profile': 'Profil Saya',
    'member-attendance': 'Kehadiran Saya',
    'member-skk': 'Progres SKK'
  };

  const pageTitle = document.getElementById('pageTitle');
  if (pageTitle) pageTitle.textContent = titles[view] || 'Dashboard';

  renderView();
  closeSidebar();
}

function switchViewByName(view) {
  switchView(view, null);
}

async function renderView() {
  const view = state.view;
  const token = state.token;
  const generation = dataLoadGeneration;
  const requiredModules = getViewModules(view);

  if (!canViewModule(view)) {
    state.view = 'dashboard';
    syncNavigationActive('dashboard');
    document.getElementById('content').innerHTML = `
      <div class="panel">
        <div class="panel-body">
          <h3>Akses Ditolak</h3>
          <p>Akun Anda tidak memiliki izin untuk membuka modul ini.</p>
        </div>
      </div>`;
    return;
  }

  if (view === 'dashboard' && !isDashboardReady()) {
    showLoading();
    try {
      if (!await ensureDashboardLoaded()) return;
      if (state.view !== view || !isCurrentDataLoad(token, generation)) return;
    } catch (error) {
      if (state.view === view && isCurrentDataLoad(token, generation)) handleDataLoadError(error);
      return;
    }
  }

  if (view === 'penilaian' && !isPenilaianReady()) {
    showLoading();
    try {
      if (!await ensurePenilaianLoaded()) return;
      if (state.view !== view || !isCurrentDataLoad(token, generation)) return;
    } catch (error) {
      if (state.view === view && isCurrentDataLoad(token, generation)) handleDataLoadError(error);
      return;
    }
  }

  if (requiredModules.length && !areModulesReady(requiredModules)) {
    showLoading();
    try {
      if (!await ensureModulesLoaded(requiredModules)) return;
      if (state.view !== view || !isCurrentDataLoad(token, generation)) return;
    } catch (error) {
      if (state.view === view && isCurrentDataLoad(token, generation)) handleDataLoadError(error);
      return;
    }
  }
  switch (state.view) {
    case 'member-profile':
      renderMemberProfile();
      break;
    case 'member-attendance':
      renderMemberAttendance();
      break;
    case 'member-skk':
      renderMemberSkk();
      break;
    case 'anggota':
      renderAnggota();
      break;
    case 'kegiatan':
      renderKegiatan();
      break;
    case 'absensi':
      renderAbsensi();
      break;
    case 'penilaian':
      renderPenilaian();
      break;
    case 'kas':
      renderKas();
      break;
    case 'inventaris':
      renderInventaris();
      break;
    case 'surat':
      renderSurat();
      break;
    case 'pengurus':
      renderPengurus();
      break;
    case 'users':
      renderUsers();
      break;
    case 'settings':
      renderSettings();
      break;
    case 'maintenance':
      state.view = 'settings';
      renderSettings();
      break;
    default:
      renderDashboard();
  }

  requestAnimationFrame(() => {
    enhanceResponsiveTables();
    syncNavigationActive(state.view);
  });
}


function renderMaintenance() {
  renderSettings();
}

function pwaInstallStatusText() {
  if (isPwaStandalone()) return 'Aplikasi sudah terpasang di perangkat ini.';
  if (deferredPwaInstallPrompt) return 'Perangkat mendukung instalasi langsung.';
  if (isIosDevice() && isSafariBrowser()) return 'Instal melalui menu Share Safari → Add to Home Screen.';
  return 'Opsi instalasi akan muncul saat browser mendukung PWA.';
}

function renderSettings() {
  const admin = isRole('ADMIN');
  const modeLabel = displayModeShortLabel();
  document.getElementById('content').innerHTML = `
    <div class="page-heading settings-heading">
      <div>
        <span class="dashboard-eyebrow">APLIKASI & SISTEM</span>
        <h2>Pengaturan</h2>
        <p>Atur pengalaman aplikasi${admin ? ', lokasi absensi anggota, dan pemeliharaan sistem' : ''} dari satu tempat.</p>
      </div>
    </div>

    <div class="settings-grid">
      <section class="panel settings-card">
        <div class="settings-card-icon"><span class="material-symbols-rounded">devices</span></div>
        <div class="settings-card-copy">
          <h3>Mode Tampilan</h3>
          <p>Sesuaikan kepadatan dan layout antarmuka dengan perangkat yang sedang digunakan.</p>
          <div class="settings-inline-status"><span>Mode aktif</span><strong>${escapeHtml(modeLabel)}</strong></div>
        </div>
        <button class="btn btn-light" type="button" onclick="openDisplayModeModal()"><span class="material-symbols-rounded">tune</span> Ubah Tampilan</button>
      </section>

      <section class="panel settings-card">
        <div class="settings-card-icon"><span class="material-symbols-rounded">install_mobile</span></div>
        <div class="settings-card-copy">
          <h3>Instal Aplikasi (PWA)</h3>
          <p>${escapeHtml(pwaInstallStatusText())}</p>
          <div class="settings-inline-status"><span>Status</span><strong>${isPwaStandalone() ? 'Terinstal' : 'Browser'}</strong></div>
        </div>
        <button id="pwaInstallSettingButton" class="btn btn-primary" type="button" onclick="installPwa()" hidden><span class="material-symbols-rounded">download</span> Instal Aplikasi</button>
      </section>
    </div>

    ${admin ? `
    <section class="panel settings-section geofence-settings-panel">
      <div class="panel-header">
        <div><span class="settings-section-kicker">ABSENSI ANGGOTA</span><h3>Area Absensi 2 KM</h3><p>Tentukan satu titik pusat. Anggota hanya dapat absen dari akun mereka jika GPS berada maksimal 2.000 meter dari titik ini.</p></div>
        <span class="geofence-radius-badge"><span class="material-symbols-rounded">radar</span> Radius tetap 2 km</span>
      </div>
      <div class="panel-body">
        <div id="geofenceSettingsState" class="geofence-settings-layout">
          <div class="geofence-loading">${loadingHtml(88)}</div>
        </div>
      </div>
    </section>

    <section class="panel settings-section">
      <div class="panel-header"><div><span class="settings-section-kicker">ADMINISTRASI SISTEM</span><h3>Pemeliharaan Database</h3><p>Periksa struktur sebelum update. Update juga memastikan trigger otomasi absensi/izin tetap aktif.</p></div></div>
      <div class="panel-body">
        <div class="maintenance-actions">
          <button class="btn btn-primary" onclick="checkMaintenance()"><span class="material-symbols-rounded">fact_check</span> Cek Struktur</button>
          <button class="btn btn-success" onclick="updateMaintenance()"><span class="material-symbols-rounded">upgrade</span> Update Struktur</button>
          <button class="btn btn-warning" onclick="resetMaintenance()"><span class="material-symbols-rounded">restart_alt</span> Reset Sesi</button>
        </div>
        <div id="maintenanceResult" class="maintenance-result"></div>
      </div>
    </section>

    <section class="panel excel-import-panel settings-section">
      <div class="panel-header">
        <div><span class="settings-section-kicker">DATA</span><h3>Import Database dari Excel</h3><p>Unggah .xlsx/.xls. Data lama tidak dihapus dan record yang cocok akan diperbarui.</p></div>
        <button class="btn btn-light" type="button" onclick="downloadExcelTemplate()"><span class="material-symbols-rounded">download</span> Unduh Template</button>
      </div>
      <div class="panel-body">
        <div class="excel-import-grid">
          <label class="excel-import-drop" for="excelDatabaseFile">
            <span class="material-symbols-rounded excel-import-icon">upload_file</span>
            <div><strong>Pilih file database Excel</strong><p>Maksimum 5 MB dan 2.000 baris. Proses dilakukan per batch agar stabil.</p></div>
            <input id="excelDatabaseFile" class="excel-file-input" type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel">
          </label>
          <div class="excel-import-notes"><strong>Mode impor aman</strong><ul><li>Data lama tidak dihapus.</li><li>ID/NTA/kode unik diperbarui bila cocok.</li><li>Data baru mendapat ID otomatis.</li><li>Absensi dicocokkan melalui ID, NTA, dan nama.</li></ul></div>
        </div>
        <div class="excel-import-actions"><button id="excelImportButton" class="btn btn-primary" type="button" onclick="importDatabaseExcel()"><span class="material-symbols-rounded">database_upload</span> Upload & Import</button></div>
        <div id="excelImportResult" class="maintenance-result"></div>
      </div>
    </section>` : ''}
  `;

  updatePwaInstallButton();
  if (admin) loadGeofenceSettings();
}

async function loadGeofenceSettings() {
  const host = document.getElementById('geofenceSettingsState');
  if (!host) return;
  try {
    const data = await serverCall('getAttendanceGeofenceSettings', state.token);
    if (!document.getElementById('geofenceSettingsState')) return;
    renderGeofenceSettingsForm(data || {});
  } catch (error) {
    host.innerHTML = `<div class="empty-state compact"><span class="material-symbols-rounded">location_off</span><strong>Pengaturan lokasi belum dapat dimuat</strong><p>${escapeHtml(error.message)}</p></div>`;
  }
}

function renderGeofenceSettingsForm(data) {
  const host = document.getElementById('geofenceSettingsState');
  if (!host) return;
  const configured = data.configured === true;
  const lat = configured ? Number(data.latitude).toFixed(6) : '';
  const lng = configured ? Number(data.longitude).toFixed(6) : '';
  host.innerHTML = `
    <div class="geofence-summary ${configured ? 'is-ready' : 'is-empty'}">
      <div class="geofence-summary-icon"><span class="material-symbols-rounded">${configured ? 'verified_user' : 'location_off'}</span></div>
      <div><small>STATUS AREA</small><strong>${configured ? 'Titik absensi aktif' : 'Belum dikonfigurasi'}</strong><p>${configured ? `${escapeHtml(data.label || 'Titik absensi')} • radius 2 km` : 'Anggota belum dapat menggunakan absensi mandiri sampai titik pusat disimpan.'}</p></div>
      ${configured && data.updatedAt ? `<span class="geofence-updated">Diperbarui ${escapeHtml(data.updatedAt)}</span>` : ''}
    </div>
    <div class="geofence-form-grid">
      <div class="form-group"><label>Nama / label lokasi</label><input id="geofenceLabel" class="form-control" maxlength="100" value="${escapeHtml(data.label || '')}" placeholder="Contoh: Lanud Muljono"></div>
      <div class="form-group"><label>Radius</label><div class="readonly-setting"><strong>2.000 meter</strong><span>Tetap / wajib</span></div></div>
      <div class="form-group"><label>Latitude</label><input id="geofenceLat" class="form-control" inputmode="decimal" value="${escapeHtml(lat)}" placeholder="-7.257472"></div>
      <div class="form-group"><label>Longitude</label><input id="geofenceLng" class="form-control" inputmode="decimal" value="${escapeHtml(lng)}" placeholder="112.752090"></div>
    </div>
    <div class="geofence-actions">
      <button class="btn btn-light" type="button" onclick="fillGeofenceFromCurrentLocation()"><span class="material-symbols-rounded">my_location</span> Gunakan Lokasi Saya</button>
      <button class="btn btn-primary" type="button" onclick="saveGeofenceSettings()"><span class="material-symbols-rounded">save</span> Simpan Titik Absensi</button>
      ${configured ? `<a class="btn btn-light" target="_blank" rel="noopener" href="https://www.google.com/maps?q=${encodeURIComponent(lat + ',' + lng)}"><span class="material-symbols-rounded">map</span> Lihat Titik</a>` : ''}
    </div>
    <div class="geofence-note"><span class="material-symbols-rounded">info</span><p>Absensi mandiri menggunakan GPS browser dan validasi jarak di backend. Lokasi dengan akurasi lebih buruk dari 500 meter akan ditolak agar pengecekan tidak terlalu kasar.</p></div>`;
}

function getBrowserLocation(options = {}) {
  return new Promise((resolve, reject) => {
    if (!window.isSecureContext) {
      reject(new Error('Pengecekan lokasi membutuhkan koneksi HTTPS.'));
      return;
    }
    if (!navigator.geolocation) {
      reject(new Error('Perangkat/browser ini tidak mendukung GPS/geolocation.'));
      return;
    }
    navigator.geolocation.getCurrentPosition(position => {
      resolve({
        latitude: Number(position.coords.latitude),
        longitude: Number(position.coords.longitude),
        accuracy: Number(position.coords.accuracy || 0),
        capturedAt: Number(position.timestamp || Date.now())
      });
    }, error => {
      const messages = {
        1: 'Izin lokasi ditolak. Aktifkan Location/GPS untuk aplikasi ini lalu coba lagi.',
        2: 'Lokasi belum dapat ditentukan. Pastikan GPS aktif dan sinyal lokasi tersedia.',
        3: 'Pencarian lokasi terlalu lama. Coba lagi di area dengan sinyal GPS lebih baik.'
      };
      reject(new Error(messages[error && error.code] || 'Gagal membaca lokasi perangkat.'));
    }, { enableHighAccuracy: true, timeout: Number(options.timeout || 15000), maximumAge: 0 });
  });
}

async function fillGeofenceFromCurrentLocation() {
  try {
    const location = await getBrowserLocation();
    const lat = document.getElementById('geofenceLat');
    const lng = document.getElementById('geofenceLng');
    if (lat) lat.value = location.latitude.toFixed(6);
    if (lng) lng.value = location.longitude.toFixed(6);
    showToast(`Lokasi terbaca • akurasi ±${Math.round(location.accuracy)} m`, location.accuracy <= 500 ? 'success' : 'warning');
  } catch (error) {
    showToast(error.message, 'error');
  }
}

async function saveGeofenceSettings() {
  const label = String(document.getElementById('geofenceLabel')?.value || '').trim();
  const latitude = Number(String(document.getElementById('geofenceLat')?.value || '').replace(',', '.'));
  const longitude = Number(String(document.getElementById('geofenceLng')?.value || '').replace(',', '.'));
  if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) {
    showToast('Latitude dan longitude wajib diisi dengan koordinat yang valid.', 'error');
    return;
  }
  try {
    const saved = await serverCall('saveAttendanceGeofenceSettings', state.token, { label, latitude, longitude });
    renderGeofenceSettingsForm(saved || {});
    showToast('Titik absensi 2 km berhasil disimpan.', 'success');
  } catch (error) {
    showToast(error.message, 'error');
  }
}

async function checkMaintenance() {
  const resultBox = document.getElementById('maintenanceResult');

  if (resultBox) {
    resultBox.innerHTML = loadingHtml(120);
  }

  try {
    const result = await serverCall('checkSystemStructure', state.token);

    if (resultBox) {
      resultBox.innerHTML = `
        <pre>${escapeHtml(JSON.stringify(result, null, 2))}</pre>
      `;
    }

    showToast('Pengecekan struktur selesai.', 'success');
  } catch (error) {
    if (resultBox) resultBox.innerHTML = errorBox(error.message);
    showToast(error.message, 'error');
  }
}

async function updateMaintenance() {
  const resultBox = document.getElementById('maintenanceResult');
  try {
    const check = await serverCall('checkSystemStructure', state.token);
    const extraColumns = check.extraColumns || [];
    const extraSheets = check.extraSheets || [];
    const details = [...extraSheets.map(name => 'Sheet: ' + name), ...extraColumns.map(item => item.sheet + ' → kolom: ' + item.header)].join('\n');
    let confirmCleanup = false;
    if (check.cleanupRequired) {
      confirmCleanup = await customConfirm('Ditemukan struktur ekstra yang akan dihapus:<br><br><pre class="dialog-details">' + escapeHtml(details) + '</pre><br>Lanjutkan update dan hapus elemen tersebut?', 'Konfirmasi Update Struktur', 'Hapus Struktur');
      if (!confirmCleanup) { showToast('Update dibatalkan; tidak ada struktur ekstra yang dihapus.', 'info'); return; }
    }
    const result = await serverCall('updateSpreadsheetStructure', state.token, confirmCleanup);
    showToast('Struktur spreadsheet berhasil diperbarui.', 'success');
    if (resultBox) resultBox.innerHTML = `<pre>${escapeHtml(JSON.stringify(result, null, 2))}</pre>`;
  } catch (error) { showToast(error.message, 'error'); if (resultBox) resultBox.innerHTML = errorBox(error.message); }
}

async function resetMaintenance() {
  if (!await customConfirm('Reset sesi akan mengeluarkan seluruh pengguna yang sedang login. Lanjutkan?', 'Reset Sesi', 'Reset Sesi')) {
    return;
  }

  try {
    await serverCall('safeResetSystem', state.token);
    clearSession();
    renderLogin();
    showToast('Seluruh sesi login telah direset.', 'success');
  } catch (error) {
    showToast(error.message, 'error');
  }
}


/* =====================================================
   EXCEL DATABASE IMPORT
===================================================== */

function ensureXlsxLibraryCore() {
  if (typeof XLSX !== 'undefined') return Promise.resolve();
  if (xlsxLoadPromise) return xlsxLoadPromise;
  const request = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
    script.async = true;
    let timeout;
    function finish(error) {
      clearTimeout(timeout);
      script.onload = null;
      script.onerror = null;
      if (error) {
        script.remove();
        reject(error);
      } else {
        resolve();
      }
    }
    script.onload = () => finish(typeof XLSX === 'undefined'
      ? new Error('Pembaca Excel tidak tersedia. Silakan coba lagi.') : null);
    script.onerror = () => finish(new Error('Gagal memuat pembaca Excel. Periksa koneksi lalu coba lagi.'));
    timeout = setTimeout(() => finish(new Error('Pemuatan pembaca Excel terlalu lama. Silakan coba lagi.')), 20000);
    document.head.appendChild(script);
  });
  xlsxLoadPromise = request.catch(error => {
    xlsxLoadPromise = null;
    throw error;
  });
  return xlsxLoadPromise;
}

async function downloadExcelTemplate() {
  const activity = startActivity('Membuat template Excel');
  try {
    await ensureXlsxLibrary();

    const workbook = XLSX.utils.book_new();
    const instructions = [
      ['TEMPLATE IMPORT DATABASE SAKA DIRGANTARA', ''],
      ['Versi', UI_VERSION],
      ['Petunjuk', 'Jangan mengubah nama sheet dan nama header.'],
      ['ID', 'Boleh dikosongkan untuk data baru. Sistem membuat ID otomatis.'],
      ['NTA Anggota', 'Pada input aplikasi NTA dibuat otomatis. Kolom NTA pada Excel dipertahankan untuk kompatibilitas database lama/import.'],
      ['Tanggal', 'Fleksibel: YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY, tanggal Excel, atau nama bulan. Sistem menormalkan otomatis.'],
      ['Absensi', 'Jika AnggotaID/KegiatanID tidak tersedia, isi NTA/NamaAnggota dan KegiatanNama.'],
      ['Users', 'Password wajib untuk user baru; kosongkan Password saat hanya memperbarui user lama.'],
      ['Krida', 'Nama resmi, singkatan, variasi ejaan, dan data Krida lama dapat diterima; nilai yang dikenali dinormalisasi otomatis.'],
      ['Keamanan', 'Import bersifat merge/update dan tidak menghapus data lama.']
    ];

    const infoSheet = XLSX.utils.aoa_to_sheet(instructions);
    infoSheet['!cols'] = [{wch: 22}, {wch: 82}];
    XLSX.utils.book_append_sheet(workbook, infoSheet, 'PETUNJUK');

    const structureRows = [['Sheet', 'Kolom']];
    Object.entries(EXCEL_TEMPLATE_COLUMNS).forEach(([name, columns]) => {
      columns.forEach(column => structureRows.push([name, column]));
    });
    const structureSheet = XLSX.utils.aoa_to_sheet(structureRows);
    structureSheet['!cols'] = [{wch: 18}, {wch: 28}];
    XLSX.utils.book_append_sheet(workbook, structureSheet, 'STRUKTUR');

    Object.entries(EXCEL_TEMPLATE_COLUMNS).forEach(([name, columns]) => {
      const sheet = XLSX.utils.aoa_to_sheet([columns]);
      sheet['!cols'] = columns.map(column => ({
        wch: Math.min(Math.max(column.length + 3, 14), 28)
      }));
      XLSX.utils.book_append_sheet(workbook, sheet, name);
    });

    XLSX.writeFile(
      workbook,
      'Template_Import_Database_SAKA_Dirgantara_v' + UI_VERSION + '.xlsx'
    );
    finishActivity(activity, 'success', 'Permintaan unduhan dikirim ke browser.');
    showToast('Template Excel siap. Permintaan unduhan telah dikirim ke browser.', 'success');
  } catch (error) {
    finishActivity(activity, 'error', error.message || 'Gagal membuat template Excel.');
    showToast(error.message || 'Gagal membuat template Excel.', 'error');
  }
}

async function importDatabaseExcel() {
  const fileInput = document.getElementById('excelDatabaseFile');
  const button = document.getElementById('excelImportButton');
  let resultBox = document.getElementById('excelImportResult');
  if (button && button.disabled) return;

  try {
    const file = fileInput && fileInput.files ? fileInput.files[0] : null;
    if (!file) throw new Error('Pilih file Excel terlebih dahulu.');
    if (!/\.(xlsx|xls)$/i.test(file.name)) throw new Error('Format file harus .xlsx atau .xls.');
    if (file.size > 5 * 1024 * 1024) throw new Error('Ukuran file maksimum 5 MB.');

    if (button) {
      button.disabled = true;
      button.innerHTML = `
        <span class="material-symbols-rounded">hourglass_top</span>
        Membaca Excel...
      `;
    }

    if (resultBox) resultBox.innerHTML = loadingHtml(130);

    await ensureXlsxLibrary();
    const parsed = await parseDatabaseExcel(file);
    const sheetNames = Object.keys(parsed.sheets);

    if (!sheetNames.length || parsed.totalRows === 0) {
      throw new Error('Tidak ditemukan data pada sheet yang didukung. Gunakan template Excel dari aplikasi.');
    }

    const summaryText = sheetNames
      .map(name => `${name}: ${parsed.sheets[name].length}`)
      .join(', ');

    const proceed = await customConfirm('File akan diimpor dengan mode GABUNG/UPDATE. Data lama tidak dihapus.<br><br>Total: ' + parsed.totalRows + ' baris<br>' + escapeHtml(summaryText) + '<br><br>Data akan diproses otomatis per batch untuk mencegah timeout. Lanjutkan?', 'Konfirmasi Import', 'Mulai Import');

    if (!proceed) {
      if (resultBox) resultBox.innerHTML = '';
      return;
    }

    const importOrder = [
      'Anggota','Kegiatan','Kas','Inventaris','Surat','Absensi','Pengurus','Users'
    ];
    const batchSize = 25;
    let context = { anggotaIdMap: {}, kegiatanIdMap: {} };
    const aggregateReport = {};
    let processed = 0;
    let imported = 0;
    let failed = 0;

    for (const sheetName of importOrder) {
      const rows = Array.isArray(parsed.sheets[sheetName]) ? parsed.sheets[sheetName] : [];
      if (!rows.length) continue;

      for (let start = 0; start < rows.length; start += batchSize) {
        const batch = rows.slice(start, start + batchSize);
        const end = Math.min(start + batch.length, rows.length);

        updateExcelImportProgress(
          processed,
          parsed.totalRows,
          sheetName,
          `Baris ${start + 1}-${end} dari ${rows.length}`
        );

        if (button) {
          button.innerHTML = `
            <span class="material-symbols-rounded">sync</span>
            ${processed} / ${parsed.totalRows}
          `;
        }

        const batchResult = await serverCall(
          'importExcelDatabaseBatch',
          state.token,
          { sheetName, rows: batch, context }
        );

        context = batchResult.context || context;
        mergeExcelImportBatchReport(aggregateReport, sheetName, batchResult.report);
        imported += Number(batchResult.imported || 0);
        failed += Number(batchResult.failed || 0);
        processed += batch.length;

        updateExcelImportProgress(
          processed,
          parsed.totalRows,
          sheetName,
          `${processed} dari ${parsed.totalRows} baris selesai`
        );
      }
    }

    try {
      await serverCall('finishExcelDatabaseImport', state.token, { imported, failed });
    } catch (_) {
      // Log kegagalan tidak boleh menggagalkan hasil import yang sudah tersimpan.
    }

    const result = {
      success: failed === 0,
      totalRows: parsed.totalRows,
      imported,
      failed,
      report: aggregateReport,
      message: failed
        ? 'Import selesai dengan beberapa baris gagal. Periksa laporan.'
        : 'Seluruh data Excel berhasil diimpor.'
    };

    await refreshData();
    resultBox = document.getElementById('excelImportResult');
    if (resultBox) resultBox.innerHTML = renderExcelImportReport(result);

    if (failed) {
      showToast(`Import selesai: ${imported} berhasil, ${failed} gagal.`, 'warning');
    } else {
      showToast(`${imported} data berhasil diimpor dari Excel.`, 'success');
    }
  } catch (error) {
    resultBox = document.getElementById('excelImportResult') || resultBox;
    if (resultBox) resultBox.innerHTML = errorBox(
      (error.message || 'Import Excel gagal.') +
      ' Data pada batch yang sudah selesai tetap tersimpan; file yang sama dapat diimpor ulang.'
    );
    showToast(error.message || 'Import Excel gagal.', 'error');
  } finally {
    if (button) {
      button.disabled = false;
      button.innerHTML = `
        <span class="material-symbols-rounded">database_upload</span>
        Upload & Import Excel
      `;
    }
  }
}

function mergeExcelImportBatchReport(target, sheetName, incoming) {
  if (!incoming) return;

  if (!target[sheetName]) {
    target[sheetName] = {
      sheet: sheetName,
      total: 0,
      created: 0,
      updated: 0,
      skipped: 0,
      failed: 0,
      normalized: 0,
      errors: [],
      warnings: []
    };
  }

  const out = target[sheetName];
  ['total','created','updated','skipped','failed','normalized'].forEach(key => {
    out[key] += Number(incoming[key] || 0);
  });

  (incoming.errors || []).forEach(item => {
    if (out.errors.length < 100) out.errors.push(item);
  });
  (incoming.warnings || []).forEach(item => {
    if (out.warnings.length < 120) out.warnings.push(item);
  });
}

function updateExcelImportProgress(done, total, sheetName, detail) {
  const resultBox = document.getElementById('excelImportResult');
  if (!resultBox) return;

  const safeTotal = Math.max(Number(total || 0), 1);
  const percent = Math.max(0, Math.min(100, Math.round((Number(done || 0) / safeTotal) * 100)));

  resultBox.innerHTML = `
    <div class="excel-import-summary">
      <div class="excel-import-summary-head">
        <strong>Mengimpor ${escapeHtml(sheetName || 'database')}...</strong>
        <span>${escapeHtml(done)} / ${escapeHtml(total)} (${percent}%)</span>
      </div>
      <div class="progress" style="margin-top:12px;">
        <div class="progress-bar" style="width:${percent}%"></div>
      </div>
      <p style="margin:10px 0 0;color:var(--muted);font-size:11px;">
        ${escapeHtml(detail || 'Memproses data...')}
      </p>
    </div>
  `;
}

function normalizeExcelImportToken(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function getExcelImportColumns(sheetName) {
  const columns = [...(EXCEL_TEMPLATE_COLUMNS[sheetName] || [])];
  if (sheetName === 'Kas' && !columns.includes('ImportKey')) {
    columns.push('ImportKey');
  }
  return columns;
}

function getExcelHeaderAliases(sheetName) {
  const aliases = {
    Anggota: {
      nomortandaanggota: 'NTA',
      noanggota: 'NTA',
      namaanggota: 'Nama',
      tanggalbergabung: 'TanggalGabung',
      tanggalpelantikan: 'TanggalPelantikan',
      tanggalpengesahan: 'TanggalPelantikan',
      tanggaldilantik: 'TanggalPelantikan',
      sekolah: 'SekolahInstansi',
      instansi: 'SekolahInstansi'
    },
    Kegiatan: {
      kegiatan: 'NamaKegiatan',
      namakegiatan: 'NamaKegiatan',
      pj: 'PenanggungJawab',
      penanggungjawabkegiatan: 'PenanggungJawab'
    },
    Absensi: {
      kegiatan: 'KegiatanNama',
      nama: 'NamaAnggota',
      anggota: 'NamaAnggota',
      status: 'StatusKehadiran',
      kehadiran: 'StatusKehadiran'
    },
    Kas: {
      jumlah: 'Nominal',
      nilai: 'Nominal'
    },
    Inventaris: {
      kode: 'KodeBarang',
      barang: 'NamaBarang',
      pj: 'PenanggungJawab'
    },
    Surat: {
      nomor: 'NomorSurat',
      file: 'LinkFile'
    },
    Pengurus: {
      namaanggota: 'Nama'
    },
    Users: {
      user: 'Username',
      namauser: 'Nama',
      passwordbaru: 'Password'
    }
  };

  return aliases[sheetName] || {};
}

function resolveExcelHeader(sheetName, rawHeader) {
  const token = normalizeExcelImportToken(rawHeader);
  if (!token) return '';

  const columns = getExcelImportColumns(sheetName);
  const direct = columns.find(column =>
    normalizeExcelImportToken(column) === token
  );
  if (direct) return direct;

  return getExcelHeaderAliases(sheetName)[token] || '';
}

function findExcelWorksheet(workbook, expectedName) {
  const aliases = {
    Anggota: ['Anggota', 'Data Anggota', 'Members'],
    Kegiatan: ['Kegiatan', 'Data Kegiatan', 'Activities'],
    Absensi: ['Absensi', 'Data Absensi', 'Attendance'],
    Kas: ['Kas', 'Keuangan', 'Cash'],
    Inventaris: ['Inventaris', 'Inventory'],
    Surat: ['Surat', 'Administrasi Surat', 'Mail'],
    Pengurus: ['Pengurus', 'Struktur Pengurus', 'Officers'],
    Users: ['Users', 'User', 'Pengguna']
  };

  const acceptedTokens = (aliases[expectedName] || [expectedName])
    .map(normalizeExcelImportToken);

  const actualName = (workbook.SheetNames || []).find(name =>
    acceptedTokens.includes(normalizeExcelImportToken(name))
  );

  return actualName ? workbook.Sheets[actualName] : null;
}

function hashExcelImportString(value) {
  const text = String(value || '');
  let h1 = 0x811c9dc5;
  let h2 = 0x9e3779b9;

  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    h1 = Math.imul(h1 ^ code, 16777619);
    h2 = Math.imul(h2 ^ (code + i + 1), 2246822519);
  }

  return (
    (h1 >>> 0).toString(16).padStart(8, '0') +
    (h2 >>> 0).toString(16).padStart(8, '0')
  );
}

function createExcelImportKey(sheetName, row, rowNumber) {
  const payload = Object.keys(row || {})
    .filter(key => key !== 'ID' && key !== 'ImportKey' && !key.startsWith('__'))
    .sort()
    .map(key => `${key}=${String(row[key] ?? '').trim()}`)
    .join('|');

  return `XLS-${String(sheetName || 'DATA').toUpperCase()}-${rowNumber}-${hashExcelImportString(payload)}`;
}

async function parseDatabaseExcelCore(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: 'array',
    cellDates: true,
    cellNF: false,
    cellText: false
  });

  const sheets = {};
  let totalRows = 0;

  Object.keys(EXCEL_TEMPLATE_COLUMNS).forEach(sheetName => {
    const worksheet = findExcelWorksheet(workbook, sheetName);
    if (!worksheet) return;

    const rawRows = XLSX.utils.sheet_to_json(worksheet, {
      defval: '',
      raw: true
    });

    const rows = rawRows
      .map((row, index) => {
        const rowNumber = index + 2;
        const normalized = { __rowNumber: rowNumber };

        Object.keys(row || {}).forEach(key => {
          const canonicalKey = resolveExcelHeader(sheetName, key);
          if (!canonicalKey) return;
          normalized[canonicalKey] = normalizeExcelValue(row[key]);
        });

        // Kunci retry internal untuk mencegah transaksi Kas tanpa ID
        // terduplikasi ketika file yang sama diimpor ulang setelah batch terputus.
        if (sheetName === 'Kas' && !String(normalized.ID || '').trim()) {
          normalized.ImportKey = normalized.ImportKey ||
            createExcelImportKey(sheetName, normalized, rowNumber);
        }

        return normalized;
      })
      .filter(row => Object.keys(row).some(key => {
        if (key === '__rowNumber' || key === 'ImportKey') return false;
        const value = row[key];
        return value !== '' && value !== null && value !== undefined;
      }));

    if (rows.length) {
      sheets[sheetName] = rows;
      totalRows += rows.length;
    }
  });

  if (totalRows > 2000) {
    throw new Error(
      `File berisi ${totalRows} baris. Maksimum 2.000 baris per sekali import.`
    );
  }

  return { sheets, totalRows };
}

function normalizeExcelValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, '0');
    const day = String(value.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  if (typeof value === 'string') return value.trim();
  return value;
}

function renderExcelImportReport(result) {
  const report = result && result.report ? result.report : {};
  const cards = Object.keys(report).map(name => {
    const item = report[name];
    const errors = (item.errors || []).length
      ? `
        <details class="excel-import-errors">
          <summary>${item.failed} baris gagal - lihat detail</summary>
          <ul>
            ${(item.errors || []).map(error => `
              <li>Baris ${escapeHtml(error.row)}: ${escapeHtml(error.message)}</li>
            `).join('')}
          </ul>
        </details>
      `
      : '';

    const warnings = (item.warnings || []).length
      ? `
        <details class="excel-import-warnings">
          <summary>${item.normalized || 0} baris dinormalisasi otomatis - lihat detail</summary>
          <ul>
            ${(item.warnings || []).map(warning => `
              <li>Baris ${escapeHtml(warning.row)}: ${escapeHtml(warning.message)}</li>
            `).join('')}
          </ul>
        </details>
      `
      : '';

    return `
      <div class="excel-report-card">
        <strong>${escapeHtml(name)}</strong>
        <div class="excel-report-stats">
          <span>Baru: ${escapeHtml(item.created)}</span>
          <span>Update: ${escapeHtml(item.updated)}</span>
          <span>Lewati: ${escapeHtml(item.skipped)}</span>
          <span>Normalisasi: ${escapeHtml(item.normalized || 0)}</span>
          <span class="${item.failed ? 'is-error' : ''}">Gagal: ${escapeHtml(item.failed)}</span>
        </div>
        ${warnings}
        ${errors}
      </div>
    `;
  }).join('');

  return `
    <div class="excel-import-summary">
      <div class="excel-import-summary-head">
        <strong>${escapeHtml(result.message || 'Import selesai.')}</strong>
        <span>${escapeHtml(result.imported || 0)} berhasil / ${escapeHtml(result.failed || 0)} gagal</span>
      </div>
      <div class="excel-report-grid">${cards}</div>
    </div>
  `;
}


/* =====================================================
   DASHBOARD
===================================================== */


function getMemberPortalData() {
  const portal = state.data && state.data.memberPortal;
  if (!portal || !portal.member) return null;
  return portal;
}

function memberPortalUnavailable() {
  const content = document.getElementById('content');
  if (!content) return;
  content.innerHTML = `
    <div class="panel">
      <div class="panel-body">
        ${errorBox('Data portal anggota belum tersedia. Muat ulang halaman atau hubungi administrator bila masalah berlanjut.')}
      </div>
    </div>`;
}

function memberProgress(value) {
  const percent = Math.max(0, Math.min(100, Number(value || 0)));
  return `<div class="member-progress" aria-label="${percent}%"><span style="width:${percent}%"></span></div>`;
}

function memberInfoItem(icon, label, value) {
  return `
    <div class="member-info-item">
      <span class="material-symbols-rounded">${icon}</span>
      <div><small>${escapeHtml(label)}</small><strong>${escapeHtml(value || '-')}</strong></div>
    </div>`;
}

function memberSelfCheckinCard(portal) {
  const config = portal.selfCheckin || {};
  const activities = Array.isArray(config.activities) ? config.activities : [];
  if (!config.geofenceConfigured) {
    return `<section class="member-checkin-card is-disabled"><div class="member-checkin-icon"><span class="material-symbols-rounded">location_off</span></div><div class="member-checkin-copy"><span class="dashboard-eyebrow">ABSENSI MANDIRI</span><h3>Lokasi absensi belum tersedia</h3><p>Admin belum menentukan titik pusat absensi. Hubungi pengurus jika kegiatan sedang berlangsung.</p></div></section>`;
  }
  if (!activities.length) {
    return `<section class="member-checkin-card is-idle"><div class="member-checkin-icon"><span class="material-symbols-rounded">event_available</span></div><div class="member-checkin-copy"><span class="dashboard-eyebrow">ABSENSI MANDIRI</span><h3>Belum ada kegiatan untuk check-in</h3><p>Saat kegiatan berstatus <b>Berjalan</b> pada hari ini, tombol absensi akan muncul di sini. Area wajib maksimal ${Math.round(Number(config.radiusMeters || 2000)/1000)} km dari ${escapeHtml(config.label || 'titik admin')}.</p></div></section>`;
  }
  return `<section class="member-checkin-card is-ready">
    <div class="member-checkin-head"><div><span class="dashboard-eyebrow">ABSENSI MANDIRI • GPS WAJIB</span><h3>Check-in kegiatan</h3><p>Lokasi Anda diverifikasi maksimal ${Number(config.radiusMeters || 2000).toLocaleString('id-ID')} m dari ${escapeHtml(config.label || 'titik absensi')}.</p></div><span class="member-checkin-radius"><span class="material-symbols-rounded">radar</span> 2 KM</span></div>
    <div class="member-checkin-list">${activities.map(item => `
      <div class="member-checkin-activity">
        <div><strong>${escapeHtml(item.NamaKegiatan || 'Kegiatan')}</strong><span>${escapeHtml(formatLongDate(item.Tanggal))} • ${escapeHtml(item.Lokasi || 'Lokasi kegiatan')}</span></div>
        ${item.allowCheckin === false && item.attendanceStatus ? `<div class="member-checkin-done">${statusBadge(item.attendanceStatus)}<small>Sudah tercatat</small></div>` : `<div class="member-checkin-action">${String(item.attendanceStatus || '').toLowerCase()==='alpa' && String(item.attendanceMethod || '').toLowerCase()==='otomatis' ? '<small class="member-auto-alpa-note">Alpa otomatis dapat diganti saat check-in valid</small>' : ''}<button class="btn btn-primary member-checkin-button" type="button" onclick="memberCheckin('${escapeHtml(String(item.ID || ''))}', this)"><span class="material-symbols-rounded">my_location</span> Absen Sekarang</button></div>`}
      </div>`).join('')}</div>
    <div id="memberLocationFeedback" class="member-location-feedback" hidden></div>
    <div class="member-checkin-privacy"><span class="material-symbols-rounded">shield</span><span>GPS dipakai untuk validasi jarak. Sistem hanya mencatat metode check-in dan jarak, bukan menyimpan koordinat GPS Anda di tabel absensi.</span></div>
  </section>`;
}

async function memberCheckin(kegiatanId, button) {
  if (!kegiatanId || !isRole('ANGGOTA')) return;
  const original = button ? button.innerHTML : '';
  if (button) { button.disabled = true; button.innerHTML = '<span class="material-symbols-rounded spin">progress_activity</span> Cek GPS...'; }
  const feedback = document.getElementById('memberLocationFeedback');
  try {
    const location = await getBrowserLocation({ timeout: 18000 });
    if (feedback) {
      feedback.hidden = false;
      feedback.className = 'member-location-feedback is-checking';
      feedback.innerHTML = `<span class="material-symbols-rounded">my_location</span><div><strong>Lokasi ditemukan</strong><span>Akurasi GPS ±${Math.round(location.accuracy)} m. Memverifikasi jarak ke titik admin...</span></div>`;
    }
    const result = await serverCall('memberAttendanceCheckin', state.token, kegiatanId, location);
    if (feedback) {
      feedback.className = 'member-location-feedback is-success';
      feedback.innerHTML = `<span class="material-symbols-rounded">verified</span><div><strong>Absensi berhasil</strong><span>Jarak ${Number(result.distanceMeters || 0).toLocaleString('id-ID')} m dari ${escapeHtml(result.geofenceLabel || 'titik absensi')} • akurasi ±${Math.round(location.accuracy)} m.</span></div>`;
    }
    showToast(result.alreadyCheckedIn ? 'Kehadiran Anda sudah tercatat.' : 'Absensi berhasil dicatat.', 'success');
    await refreshData();
  } catch (error) {
    if (feedback) {
      feedback.hidden = false;
      feedback.className = 'member-location-feedback is-error';
      feedback.innerHTML = `<span class="material-symbols-rounded">location_off</span><div><strong>Check-in ditolak</strong><span>${escapeHtml(error.message)}</span></div>`;
    }
    showToast(error.message, 'error');
  } finally {
    if (button && document.body.contains(button)) { button.disabled = false; button.innerHTML = original; }
  }
}

function renderMemberDashboard() {
  const portal = getMemberPortalData();
  if (!portal) return memberPortalUnavailable();
  const member = portal.member || {};
  const attendance = portal.attendance || [];
  const stats = portal.attendanceStats || {};
  const upcoming = portal.upcoming || [];
  const skk = portal.skk || {};
  const skkStats = skk.stats || {};
  const relevantTotal = Number(skkStats.relevantItems || skkStats.totalItems || 0);
  const relevantDone = Number(skkStats.relevantItems ? skkStats.relevantCompletedItems : skkStats.completedItems || 0);
  const skkPercent = Number(skkStats.percent || 0);

  const agendaHtml = upcoming.length ? upcoming.slice(0, 4).map(item => `
    <article class="member-agenda-item">
      <div class="agenda-date-box"><strong>${escapeHtml(formatDayNumber(item.Tanggal))}</strong><span>${escapeHtml(formatShortMonth(item.Tanggal))}</span></div>
      <div class="member-agenda-copy">
        <strong>${escapeHtml(item.NamaKegiatan || 'Kegiatan')}</strong>
        <span>${escapeHtml(item.Lokasi || 'Lokasi belum ditentukan')}</span>
        <small>${escapeHtml(formatLongDate(item.Tanggal))}${item.PenanggungJawab ? ` • ${escapeHtml(item.PenanggungJawab)}` : ''}</small>
      </div>
      ${statusBadge(item.Status)}
    </article>`).join('') : emptyDashboardState('event_busy','Belum ada agenda mendatang.');

  const attendanceHtml = attendance.length ? attendance.slice(0, 5).map(item => `
    <div class="member-history-row">
      <div class="member-history-icon"><span class="material-symbols-rounded">${String(item.StatusKehadiran).toLowerCase()==='hadir'?'check_circle':'event_note'}</span></div>
      <div class="member-history-copy"><strong>${escapeHtml(item.KegiatanNama || 'Kegiatan')}</strong><span>${escapeHtml(formatLongDate(item.Tanggal))}${item.Lokasi ? ` • ${escapeHtml(item.Lokasi)}` : ''}</span></div>
      ${statusBadge(item.StatusKehadiran)}
    </div>`).join('') : emptyDashboardState('fact_check','Belum ada riwayat kehadiran.');

  document.getElementById('content').innerHTML = `
    <div class="page-heading dashboard-heading member-welcome-heading">
      <div>
        <span class="dashboard-date">${escapeHtml(formatDashboardToday())}</span>
        <h2>Halo, ${escapeHtml(member.Nama || state.user.Nama)}</h2>
        <p>${escapeHtml(member.NTA || 'NTA belum tersedia')} • ${escapeHtml(member.Krida || 'Krida belum ditentukan')} • ${escapeHtml(member.Status || '-')}</p>
      </div>
      <div class="member-status-card">${statusBadge(member.Status)}</div>
    </div>

    ${memberSelfCheckinCard(portal)}

    <div class="stats-grid dashboard-kpi-grid member-kpi-grid">
      ${statCard('percent','green','Kehadiran',`${Number(stats.rate || 0)}%`,`${Number(stats.Hadir || 0)} hadir dari ${Number(stats.total || 0)} pertemuan`,"switchViewByName('member-attendance')")}
      ${statCard('check_circle','blue','Total Hadir',Number(stats.Hadir || 0),`${Number(stats.Izin || 0)} izin • ${Number(stats.Sakit || 0)} sakit` ,"switchViewByName('member-attendance')")}
      ${statCard('workspace_premium','gold','Progres SKK',skk.available ? `${skkPercent}%` : '-',skk.available ? `${relevantDone} dari ${relevantTotal} butir sesuai Krida` : (skk.message || 'Belum tersedia'),"switchViewByName('member-skk')")}
      ${statCard('event','cyan','Agenda Mendatang',upcoming.length,upcoming.length ? `Terdekat ${formatDate(upcoming[0].Tanggal)}` : 'Belum ada agenda')}
    </div>

    <div class="dashboard-section-grid primary member-dashboard-grid">
      <section class="panel dashboard-panel">
        <div class="panel-header"><div><h3>Agenda Mendatang</h3><p>Kegiatan yang dapat Anda persiapkan berikutnya</p></div><span class="dashboard-panel-total">${upcoming.length} agenda</span></div>
        <div class="panel-body"><div class="member-agenda-list">${agendaHtml}</div></div>
      </section>

      <section class="panel dashboard-panel">
        <div class="panel-header"><div><h3>Progress Saya</h3><p>Ringkasan kehadiran dan kecakapan</p></div></div>
        <div class="panel-body">
          <div class="member-progress-card">
            <div class="member-progress-head"><div><small>KEHADIRAN</small><strong>${Number(stats.rate || 0)}%</strong></div><span>${Number(stats.total || 0)} catatan</span></div>
            ${memberProgress(stats.rate)}
          </div>
          <div class="member-progress-card">
            <div class="member-progress-head"><div><small>SKK ${member.Krida ? '• ' + escapeHtml(member.Krida) : ''}</small><strong>${skk.available ? skkPercent + '%' : '-'}</strong></div><span>${skk.available ? `${relevantDone}/${relevantTotal} butir` : 'Belum tersedia'}</span></div>
            ${memberProgress(skk.available ? skkPercent : 0)}
          </div>
          <button class="dashboard-link-button" type="button" onclick="switchViewByName('member-skk')">Lihat detail progres SKK <span class="material-symbols-rounded">arrow_forward</span></button>
        </div>
      </section>
    </div>

    <section class="panel dashboard-panel member-recent-panel">
      <div class="panel-header"><div><h3>Riwayat Kehadiran Terbaru</h3><p>Lima catatan terakhir pada kegiatan SAKA Dirgantara</p></div></div>
      <div class="panel-body"><div class="member-history-list">${attendanceHtml}</div><button class="dashboard-link-button" type="button" onclick="switchViewByName('member-attendance')">Lihat semua riwayat <span class="material-symbols-rounded">arrow_forward</span></button></div>
    </section>`;
}

function renderMemberProfile() {
  const portal = getMemberPortalData();
  if (!portal) return memberPortalUnavailable();
  const m = portal.member || {};
  document.getElementById('content').innerHTML = `
    <div class="page-heading"><div><h2>Profil Saya</h2><p>Data keanggotaan yang tercatat pada sistem. Hubungi pengurus jika ada data yang perlu diperbarui.</p></div></div>
    <section class="panel member-profile-card">
      <div class="member-profile-hero">
        <div class="member-profile-avatar">${escapeHtml(getInitials(m.Nama))}</div>
        <div><span class="dashboard-eyebrow">ANGGOTA SAKA DIRGANTARA</span><h3>${escapeHtml(m.Nama || '-')}</h3><p>${escapeHtml(m.NTA || 'NTA belum tersedia')}</p></div>
        <div>${statusBadge(m.Status)}</div>
      </div>
      <div class="panel-body member-info-grid">
        ${memberInfoItem('badge','NTA',m.NTA)}
        ${memberInfoItem('groups','Krida',m.Krida)}
        ${memberInfoItem('military_tech','Jabatan',m.Jabatan)}
        ${memberInfoItem('wc','Jenis Kelamin',m.JenisKelamin)}
        ${memberInfoItem('cake','Tempat, Tanggal Lahir',[m.TempatLahir, m.TanggalLahir ? formatDate(m.TanggalLahir) : ''].filter(Boolean).join(', '))}
        ${memberInfoItem('school','Sekolah / Instansi',m.SekolahInstansi)}
        ${memberInfoItem('call','Telepon',m.Telepon)}
        ${memberInfoItem('home','Alamat',m.Alamat)}
        ${memberInfoItem('event','Tanggal Gabung',m.TanggalGabung ? formatLongDate(m.TanggalGabung) : '-')}
        ${memberInfoItem('verified','Tanggal Pelantikan',m.TanggalPelantikan ? formatLongDate(m.TanggalPelantikan) : '-')}
      </div>
    </section>`;
}

function renderMemberAttendance() {
  const portal = getMemberPortalData();
  if (!portal) return memberPortalUnavailable();
  const rows = portal.attendance || [];
  const stats = portal.attendanceStats || {};
  const body = rows.length ? rows.map((item,index) => `
    <tr>
      <td>${index + 1}</td>
      <td><span class="table-name">${escapeHtml(item.KegiatanNama || 'Kegiatan')}</span><small class="member-table-meta">${escapeHtml(item.Lokasi || 'Lokasi -')}</small></td>
      <td>${escapeHtml(formatLongDate(item.Tanggal))}</td>
      <td>${statusBadge(item.StatusKehadiran)}</td>
      <td>${escapeHtml(item.Catatan || '-')}</td>
      <td>${escapeHtml(item.Metode || '-')}</td>
    </tr>`).join('') : emptyTableRow(6,'Belum ada riwayat kehadiran.');

  document.getElementById('content').innerHTML = `
    <div class="page-heading"><div><h2>Kehadiran Saya</h2><p>Check-in kegiatan dan pantau seluruh riwayat kehadiran dari akun Anda.</p></div></div>
    ${memberSelfCheckinCard(portal)}
    <div class="stats-grid dashboard-kpi-grid member-kpi-grid">
      ${statCard('percent','green','Persentase Hadir',`${Number(stats.rate || 0)}%`,`${Number(stats.total || 0)} pertemuan tercatat`)}
      ${statCard('check_circle','blue','Hadir',Number(stats.Hadir || 0))}
      ${statCard('event_available','cyan','Izin / Sakit',Number(stats.Izin || 0) + Number(stats.Sakit || 0),`${Number(stats.Izin || 0)} izin • ${Number(stats.Sakit || 0)} sakit`)}
      ${statCard('person_off','red','Alpa',Number(stats.Alpa || 0))}
    </div>
    <section class="panel">
      <div class="panel-header"><div><h3>Riwayat Lengkap</h3><p>${rows.length} catatan kehadiran</p></div></div>
      <div class="table-wrap"><table class="data-table"><thead><tr><th>No</th><th>Kegiatan</th><th>Tanggal</th><th>Status</th><th>Catatan</th><th>Metode</th></tr></thead><tbody>${body}</tbody></table></div>
    </section>`;
}

function renderMemberSkk() {
  const portal = getMemberPortalData();
  if (!portal) return memberPortalUnavailable();
  const member = portal.member || {};
  const skk = portal.skk || {};
  if (!skk.available) {
    document.getElementById('content').innerHTML = `<div class="page-heading"><div><h2>Progres SKK</h2><p>Checklist kecakapan khusus pribadi.</p></div></div><div class="panel"><div class="panel-body">${errorBox(skk.message || 'Data SKK belum tersedia.')}</div></div>`;
    return;
  }

  const completedRefs = new Set((skk.completion || []).map(item => String(item.ReferensiID || '')));
  const memberKrida = String(member.Krida || '').trim().toLowerCase();
  let catalog = (skk.catalog || []).filter(item => !memberKrida || String(item.Krida || '').trim().toLowerCase() === memberKrida);
  if (!catalog.length) catalog = skk.catalog || [];
  const groups = {};
  catalog.forEach(item => {
    const key = String(item.KelompokSKK || 'SKK');
    (groups[key] ||= []).push(item);
  });
  const stats = skk.stats || {};
  const total = catalog.length;
  const done = catalog.filter(item => completedRefs.has(String(item.ID || ''))).length;
  const percent = total ? Math.round(done / total * 100) : 0;
  const groupHtml = Object.keys(groups).map(group => {
    const items = groups[group];
    const completed = items.filter(item => completedRefs.has(String(item.ID || ''))).length;
    return `<section class="member-skk-group"><div class="member-skk-group-head"><div><h3>${escapeHtml(group)}</h3><p>${completed} dari ${items.length} butir selesai</p></div><strong>${items.length ? Math.round(completed/items.length*100) : 0}%</strong></div><div class="member-skk-list">${items.map(item => {
      const checked = completedRefs.has(String(item.ID || ''));
      return `<div class="member-skk-item ${checked ? 'is-complete' : ''}"><span class="material-symbols-rounded">${checked ? 'check_circle' : 'radio_button_unchecked'}</span><div><strong>${escapeHtml(item.Kode || 'SKK')}</strong><p>${escapeHtml(item.Butir || '-')}</p></div></div>`;
    }).join('')}</div></section>`;
  }).join('');

  document.getElementById('content').innerHTML = `
    <div class="page-heading"><div><h2>Progres SKK</h2><p>${escapeHtml(member.Krida || 'Semua Krida')} • progres checklist yang telah diverifikasi pengurus.</p></div></div>
    <section class="panel member-skk-summary"><div class="panel-body"><div class="member-skk-summary-head"><div><span class="dashboard-eyebrow">PROGRES KECAKAPAN</span><h3>${done} / ${total} butir</h3><p>${Number(stats.totalEarnedPoints || 0)} poin SKK tercatat</p></div><strong>${percent}%</strong></div>${memberProgress(percent)}</div></section>
    <div class="member-skk-groups">${groupHtml || emptyDashboardState('workspace_premium','Belum ada master SKK untuk Krida Anda.')}</div>`;
}

function renderDashboard() {
  if (isRole('ANGGOTA')) {
    renderMemberDashboard();
    return;
  }
  const d = state.data.dashboard || {};
  const upcoming = Array.isArray(d.kegiatanTerbaru) ? d.kegiatanTerbaru : [];
  const lastMeeting = d.pertemuanTerakhir || null;
  const trend = Array.isArray(d.trenKehadiran) ? d.trenKehadiran : [];
  const krida = Array.isArray(d.komposisiKrida) ? d.komposisiKrida : [];
  const transactions = Array.isArray(d.transaksiTerakhir) ? d.transaksiTerakhir : [];
  const attention = d.perhatian || {};

  const activityMeta = [
    d.kegiatanSelesaiBulanIni ? `${d.kegiatanSelesaiBulanIni} selesai` : '',
    d.kegiatanBerjalanBulanIni ? `${d.kegiatanBerjalanBulanIni} berjalan` : '',
    d.kegiatanRencanaBulanIni ? `${d.kegiatanRencanaBulanIni} rencana` : '',
    d.kegiatanLiburBulanIni ? `${d.kegiatanLiburBulanIni} libur` : '',
    d.kegiatanDibatalkanBulanIni ? `${d.kegiatanDibatalkanBulanIni} dibatalkan` : ''
  ].filter(Boolean).join(' • ') || 'Belum ada kegiatan bulan ini';

  const attendanceMeta = dashboardAttendanceDeltaText(
    d.tingkatKehadiranBulanIni,
    d.tingkatKehadiranBulanLalu,
    d.perubahanKehadiran
  );

  const inventoryMeta = d.inventarisPerluTindakLanjut
    ? `${d.inventarisPerluTindakLanjut} unit rusak berat/hilang`
    : `${d.inventarisKondisi?.baik || 0} unit kondisi baik`;

  const inventorySummaryHtml = `
    <div class="inventory-dashboard-summary">
      <div class="inventory-summary-total">
        <span class="material-symbols-rounded">inventory_2</span>
        <div><strong>${escapeHtml(d.totalInventaris || 0)}</strong><span>Total unit tercatat</span></div>
      </div>
      <div class="inventory-summary-grid">
        <div><strong>${escapeHtml(d.inventarisKondisi?.baik || 0)}</strong><span>Baik</span></div>
        <div><strong>${escapeHtml((d.inventarisKondisi?.rusakRingan || 0) + (d.inventarisKondisi?.rusakBerat || 0))}</strong><span>Perlu cek</span></div>
        <div><strong>${escapeHtml(d.inventarisKondisi?.hilang || 0)}</strong><span>Hilang</span></div>
        <div><strong>${escapeHtml(d.inventarisUnitDipakaiBulanIni || 0)}</strong><span>Dipakai bulan ini</span></div>
      </div>
    </div>
    <button class="dashboard-link-button" type="button" onclick="switchViewByName('inventaris')">
      Kelola inventaris & riwayat kegiatan
      <span class="material-symbols-rounded">arrow_forward</span>
    </button>
  `;

  const letterMeta = d.suratBulanIni
    ? `${d.suratMasukBulanIni || 0} masuk • ${d.suratKeluarBulanIni || 0} keluar bulan ini`
    : 'Belum ada surat bulan ini';

  const upcomingHtml = upcoming.length
    ? upcoming.map(item => `
      <button class="agenda-item" type="button" onclick="switchViewByName('kegiatan')">
        <div class="agenda-date-box">
          <strong>${escapeHtml(formatDayNumber(item.Tanggal))}</strong>
          <span>${escapeHtml(formatShortMonth(item.Tanggal))}</span>
        </div>
        <div class="agenda-copy">
          <strong>${escapeHtml(item.NamaKegiatan || '-')}</strong>
          <span>${escapeHtml(item.Lokasi || 'Lokasi belum ditentukan')}</span>
          <small>${escapeHtml(daysUntilText(item.Tanggal))}</small>
        </div>
        <div>${statusBadge(item.Status)}</div>
      </button>
    `).join('')
    : emptyDashboardState('event_busy', 'Belum ada agenda mendatang.');

  const lastMeetingHtml = lastMeeting
    ? `
      <div class="last-meeting-head">
        <div>
          <span class="dashboard-eyebrow">${lastMeeting.isLibur ? 'PERTEMUAN LIBUR' : 'PERTEMUAN TERAKHIR'}</span>
          <h3>${escapeHtml(lastMeeting.nama || 'Pertemuan')}</h3>
          <p>${escapeHtml(formatLongDate(lastMeeting.tanggal))}${lastMeeting.lokasi ? ` • ${escapeHtml(lastMeeting.lokasi)}` : ''}</p>
        </div>
        ${lastMeeting.isLibur
          ? `<span class="badge gray">Libur</span>`
          : `<div class="meeting-rate"><strong>${escapeHtml(lastMeeting.rate ?? 0)}%</strong><span>Kehadiran</span></div>`
        }
      </div>
      ${lastMeeting.isLibur ? `
        <div class="dashboard-notice neutral">
          <span class="material-symbols-rounded">event_busy</span>
          <div><strong>Tidak ada latihan</strong><span>Pertemuan ini dicatat sebagai hari libur.</span></div>
        </div>
      ` : `
        <div class="meeting-stats-grid">
          ${meetingMiniStat('check_circle', 'Hadir', lastMeeting.hadir || 0, 'green')}
          ${meetingMiniStat('event_available', 'Izin', lastMeeting.izin || 0, 'blue')}
          ${meetingMiniStat('sick', 'Sakit', lastMeeting.sakit || 0, 'gold')}
          ${meetingMiniStat('person_off', 'Alpa', lastMeeting.alpa || 0, 'red')}
        </div>
      `}
      <button class="dashboard-link-button" type="button" onclick="switchViewByName('absensi')">
        Lihat detail absensi
        <span class="material-symbols-rounded">arrow_forward</span>
      </button>
    `
    : emptyDashboardState('fact_check', 'Belum ada riwayat pertemuan/absensi.');

  const trendHtml = trend.length
    ? `<div class="attendance-trend-list">${trend.map(item => {
        const rate = item.rate === null || item.rate === undefined ? 0 : Number(item.rate);
        return `
          <div class="attendance-trend-row">
            <div class="trend-label">
              <strong>${escapeHtml(formatDate(item.tanggal))}</strong>
              <span>${escapeHtml(item.nama || 'Pertemuan')}</span>
            </div>
            <div class="trend-track">
              <div class="trend-fill ${item.isLibur ? 'is-libur' : ''}" style="width:${item.isLibur ? 100 : Math.max(0, Math.min(rate, 100))}%"></div>
            </div>
            <strong class="trend-value">${item.isLibur ? 'Libur' : `${rate}%`}</strong>
          </div>
        `;
      }).join('')}</div>`
    : emptyDashboardState('monitoring', 'Belum cukup data untuk menampilkan tren.');

  const transactionHtml = transactions.length
    ? `<div class="transaction-list">${transactions.map(item => `
        <div class="transaction-item">
          <div class="transaction-icon ${String(item.Jenis || '').toLowerCase() === 'pemasukan' ? 'in' : 'out'}">
            <span class="material-symbols-rounded">${String(item.Jenis || '').toLowerCase() === 'pemasukan' ? 'south_west' : 'north_east'}</span>
          </div>
          <div class="transaction-copy">
            <strong>${escapeHtml(item.Kategori || item.Keterangan || item.Jenis || '-')}</strong>
            <span>${escapeHtml(formatDate(item.Tanggal))}</span>
          </div>
          <strong class="transaction-value ${String(item.Jenis || '').toLowerCase() === 'pemasukan' ? 'in' : 'out'}">
            ${String(item.Jenis || '').toLowerCase() === 'pemasukan' ? '+' : '-'}${escapeHtml(formatRupiah(Math.abs(Number(item.Nominal || 0))))}
          </strong>
        </div>
      `).join('')}</div>`
    : emptyDashboardState('payments', 'Belum ada transaksi kas.');

  const activeMembers = Number(d.anggotaAktif || 0);
  const kridaHtml = krida.length
    ? `<div class="krida-list">${krida.map(item => {
        const percent = activeMembers ? Math.round((Number(item.count || 0) / activeMembers) * 100) : 0;
        return `
          <div class="krida-row">
            <div class="krida-row-head">
              <span>${escapeHtml(item.name || '-')}</span>
              <strong>${escapeHtml(item.count || 0)}</strong>
            </div>
            <div class="krida-track"><div class="krida-fill" style="width:${Math.max(0, Math.min(percent, 100))}%"></div></div>
          </div>
        `;
      }).join('')}</div>`
    : emptyDashboardState('groups', 'Belum ada komposisi Krida yang dapat ditampilkan.');

  const attentionItems = [
    {
      count: Number(attention.anggotaCalon || 0),
      icon: 'person_add',
      title: 'Calon anggota belum terlantik',
      view: 'anggota',
      severity: 'warning'
    },
    {
      count: Number(attention.anggotaTanpaNta || 0),
      icon: 'badge',
      title: 'Anggota aktif belum memiliki NTA',
      view: 'anggota',
      severity: 'warning'
    },
    {
      count: Number(attention.inventarisBermasalah || 0),
      icon: 'inventory_2',
      title: 'Unit inventaris rusak berat / hilang',
      view: 'inventaris',
      severity: 'danger'
    },
    {
      count: Number(attention.kegiatanBelumLengkap || 0),
      icon: 'event_note',
      title: 'Agenda mendatang belum lengkap lokasi / PJ',
      view: 'kegiatan',
      severity: 'warning'
    },
    {
      count: Number(attention.anggotaNonaktif || 0),
      icon: 'person_off',
      title: 'Anggota berstatus nonaktif',
      view: 'anggota',
      severity: 'neutral'
    }
  ].filter(item => item.count > 0);

  const attentionHtml = attentionItems.length
    ? `<div class="attention-list">${attentionItems.map(item => `
        <button class="attention-item ${item.severity}" type="button" onclick="switchViewByName('${item.view}')">
          <span class="material-symbols-rounded">${item.icon}</span>
          <div><strong>${escapeHtml(item.count)}</strong><span>${escapeHtml(item.title)}</span></div>
          <span class="material-symbols-rounded attention-arrow">chevron_right</span>
        </button>
      `).join('')}</div>`
    : `
      <div class="dashboard-notice success">
        <span class="material-symbols-rounded">verified</span>
        <div><strong>Tidak ada isu prioritas</strong><span>Data utama organisasi dalam kondisi baik.</span></div>
      </div>
    `;

  const quickActions = [
    canPermission('anggota','create') ? dashboardQuickAction('person_add', 'Tambah Anggota', "openForm('anggota')") : '',
    canPermission('kegiatan','create') ? dashboardQuickAction('event', 'Tambah Kegiatan', "openForm('kegiatan')") : '',
    canPermission('absensi','create') ? dashboardQuickAction('playlist_add_check', 'Absensi Kegiatan', "openBatchAbsensi()") : '',
    canPermission('kas','create') ? dashboardQuickAction('payments', 'Transaksi Kas', "openForm('kas')") : ''
  ].filter(Boolean).join('');

  document.getElementById('content').innerHTML = `
    <div class="page-heading dashboard-heading">
      <div>
        <span class="dashboard-date">${escapeHtml(formatDashboardToday())}</span>
        <h2>Selamat Datang, ${escapeHtml(state.user.Nama)}</h2>
        <p>Ringkasan operasional organisasi SAKA Dirgantara bulan berjalan.</p>
      </div>

      <div class="page-actions">
        ${state.data.databaseUrl ? `
          <a href="${escapeHtml(state.data.databaseUrl)}" target="_blank" rel="noopener noreferrer" class="btn btn-light">
            <span class="material-symbols-rounded">database</span> Database
          </a>
        ` : ''}
      </div>
    </div>

    ${quickActions ? `<div class="dashboard-quick-actions">${quickActions}</div>` : ''}

    <div class="stats-grid dashboard-kpi-grid">
      ${statCard('groups','blue','Anggota Aktif',d.anggotaAktif || 0,`${d.anggotaCalon || 0} calon • ${d.anggotaBaruBulanIni || 0} baru bulan ini`,"switchViewByName('anggota')")}
      ${statCard('event','cyan','Kegiatan Bulan Ini',d.kegiatanBulanIni || 0,activityMeta,"switchViewByName('kegiatan')")}
      ${statCard('fact_check','green','Kehadiran Bulan Ini',(d.tingkatKehadiranBulanIni || 0) + '%',attendanceMeta,"switchViewByName('absensi')")}
      ${statCard('account_balance_wallet','gold','Saldo Kas',formatRupiah(d.saldo || 0),dashboardNetText(d.netBulanIni),"switchViewByName('kas')")}
      ${statCard('inventory_2','purple','Unit Inventaris',d.totalInventaris || 0,inventoryMeta,"switchViewByName('inventaris')")}
      ${statCard('mail','red','Administrasi Surat',d.suratBulanIni || 0,letterMeta,"switchViewByName('surat')")}
    </div>

    <div class="dashboard-section-grid primary">
      <div class="panel dashboard-panel">
        <div class="panel-header">
          <div><h3>Pertemuan Terakhir</h3><p>Rekap kehadiran pada pertemuan terakhir yang tercatat</p></div>
        </div>
        <div class="panel-body">${lastMeetingHtml}</div>
      </div>

      <div class="panel dashboard-panel">
        <div class="panel-header">
          <div><h3>Agenda Berikutnya</h3><p>Kegiatan terdekat yang masih aktif</p></div>
          <button class="btn btn-light btn-compact" type="button" onclick="switchViewByName('kegiatan')">Lihat Semua</button>
        </div>
        <div class="agenda-list">${upcomingHtml}</div>
      </div>
    </div>

    <div class="dashboard-section-grid secondary">
      <div class="panel dashboard-panel">
        <div class="panel-header">
          <div><h3>Tren Kehadiran</h3><p>Lima pertemuan terakhir</p></div>
        </div>
        <div class="panel-body">${trendHtml}</div>
      </div>

      <div class="panel dashboard-panel">
        <div class="panel-header">
          <div><h3>Keuangan Bulan Ini</h3><p>Arus kas bulan berjalan dan transaksi terbaru</p></div>
          <button class="btn btn-light btn-compact" type="button" onclick="switchViewByName('kas')">Kas</button>
        </div>
        <div class="panel-body">
          <div class="monthly-finance-grid">
            <div><span>Pemasukan</span><strong class="is-positive">${escapeHtml(formatRupiah(d.pemasukanBulanIni || 0))}</strong></div>
            <div><span>Pengeluaran</span><strong class="is-negative">${escapeHtml(formatRupiah(d.pengeluaranBulanIni || 0))}</strong></div>
            <div><span>Net Bulan Ini</span><strong class="${Number(d.netBulanIni || 0) >= 0 ? 'is-positive' : 'is-negative'}">${escapeHtml(formatRupiah(d.netBulanIni || 0))}</strong></div>
          </div>
          <div class="dashboard-subtitle">Transaksi terbaru</div>
          ${transactionHtml}
        </div>
      </div>
    </div>

    <div class="dashboard-section-grid secondary">
      <div class="panel dashboard-panel">
        <div class="panel-header">
          <div><h3>Ringkasan Inventaris</h3><p>Kondisi aset dan pemakaian untuk kegiatan</p></div>
          <span class="dashboard-panel-total">${escapeHtml(d.inventarisKegiatanBulanIni || 0)} riwayat bulan ini</span>
        </div>
        <div class="panel-body">${inventorySummaryHtml}</div>
      </div>

      <div class="panel dashboard-panel">
        <div class="panel-header">
          <div><h3>Perlu Perhatian</h3><p>Data yang sebaiknya segera ditindaklanjuti</p></div>
        </div>
        <div class="panel-body">${attentionHtml}</div>
      </div>
    </div>
  `;
}

function dashboardQuickAction(icon, label, action) {
  return `
    <button class="dashboard-quick-action" type="button" onclick="${action}">
      <span class="material-symbols-rounded">${icon}</span>
      <span>${escapeHtml(label)}</span>
    </button>
  `;
}

function meetingMiniStat(icon, label, value, color) {
  return `
    <div class="meeting-mini-stat ${color}">
      <span class="material-symbols-rounded">${icon}</span>
      <div><strong>${escapeHtml(value)}</strong><span>${escapeHtml(label)}</span></div>
    </div>
  `;
}

function emptyDashboardState(icon, message) {
  return `
    <div class="dashboard-empty">
      <span class="material-symbols-rounded">${icon}</span>
      <p>${escapeHtml(message)}</p>
    </div>
  `;
}

function dashboardAttendanceDeltaText(current, previous, delta) {
  if (previous === null || previous === undefined || delta === null || delta === undefined) {
    return 'Belum ada pembanding bulan lalu';
  }
  const value = Number(delta || 0);
  if (value > 0) return `Naik ${value}% dari bulan lalu`;
  if (value < 0) return `Turun ${Math.abs(value)}% dari bulan lalu`;
  return 'Sama dengan bulan lalu';
}

function dashboardNetText(value) {
  const amount = Number(value || 0);
  if (amount > 0) return `Surplus ${formatRupiah(amount)} bulan ini`;
  if (amount < 0) return `Defisit ${formatRupiah(Math.abs(amount))} bulan ini`;
  return 'Arus kas bulan ini seimbang';
}

function formatDashboardToday() {
  try {
    return new Intl.DateTimeFormat('id-ID', {
      weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
    }).format(new Date());
  } catch (_) {
    return '';
  }
}

function parseLocalYmd(value) {
  const parts = String(value || '').substring(0, 10).split('-').map(Number);
  if (parts.length !== 3 || parts.some(Number.isNaN)) return null;
  const date = new Date(parts[0], parts[1] - 1, parts[2]);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatLongDate(value) {
  const date = parseLocalYmd(value);
  if (!date) return formatDate(value);
  try {
    return new Intl.DateTimeFormat('id-ID', {
      day: 'numeric', month: 'long', year: 'numeric'
    }).format(date);
  } catch (_) {
    return formatDate(value);
  }
}

function formatDayNumber(value) {
  const date = parseLocalYmd(value);
  return date ? String(date.getDate()).padStart(2, '0') : '-';
}

function formatShortMonth(value) {
  const date = parseLocalYmd(value);
  if (!date) return '-';
  try {
    return new Intl.DateTimeFormat('id-ID', { month: 'short' }).format(date).replace('.', '');
  } catch (_) {
    return '-';
  }
}

function daysUntilText(value) {
  const target = parseLocalYmd(value);
  if (!target) return '';
  const today = new Date();
  const current = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const diff = Math.round((target.getTime() - current.getTime()) / 86400000);
  if (diff === 0) return 'Hari ini';
  if (diff === 1) return 'Besok';
  if (diff > 1) return `${diff} hari lagi`;
  if (diff === -1) return 'Kemarin';
  return `${Math.abs(diff)} hari lalu`;
}


/* =====================================================
   ANGGOTA
===================================================== */

function renderAnggota() {
  const data = sortAnggotaForList(filterData(
    state.data.anggota,
    ['Nama','NTA','Krida','Jabatan','Status','TanggalGabung']
  ));

  const canManage = canPermission('anggota','edit') || canPermission('anggota','create') || canPermission('anggota','delete');
  const pager = paginateData('anggota', data);
  const activeCount = state.data.anggota.filter(item =>
    String(item.Status || '').trim().toLowerCase() === 'aktif'
  ).length;
  const candidateCount = state.data.anggota.filter(item =>
    String(item.Status || '').trim().toLowerCase() === 'calon anggota'
  ).length;
  const inactiveCount = state.data.anggota.filter(item => {
    const status = String(item.Status || '').trim().toLowerCase().replace(/\s+/g, '');
    return status === 'nonaktif';
  }).length;

  const rows = pager.items.length
    ? pager.items.map(item => `
      <tr>
        <td>
          <span class="table-name">${escapeHtml(item.Nama)}</span>
        </td>
        <td>${escapeHtml(item.NTA || '-')}</td>
        <td>${escapeHtml(item.Krida || '-')}</td>
        <td>${escapeHtml(item.Jabatan || '-')}</td>
        <td>${statusBadge(item.Status)}</td>

        <td>
          <div class="row-actions">

            <button
              class="btn-icon gold"
              title="Kartu anggota"
              onclick="showMemberCard('${escapeJs(item.ID)}')"
            >
              <span class="material-symbols-rounded">badge</span>
            </button>

            ${canManage ? `
              <button ${canPermission('anggota', 'edit') ? '' : 'disabled'}
                class="btn-icon"
                title="Edit"
                onclick="openForm('anggota','${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">edit</span>
              </button>

              <button ${canPermission('anggota', 'delete') ? '' : 'disabled'}
                class="btn-icon delete"
                title="Hapus"
                onclick="deleteItem('anggota','${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">delete</span>
              </button>
            ` : ''}

          </div>
        </td>
      </tr>
    `).join('')
    : emptyTableRow(6, 'Belum ada data anggota.');

  renderManagementPage({
    title: 'Data Anggota',
    description: `${state.data.anggota.length} anggota terdaftar • ${candidateCount} calon • ${activeCount} aktif • ${inactiveCount} nonaktif. NTA dibuat otomatis oleh sistem.`,
    type: 'anggota',
    addLabel: 'Tambah Anggota',
    canAdd: canManage,
    table: `
      <table class="data-table">
        <thead>
          <tr>
            <th>Nama</th>
            <th>NTA</th>
            <th>Krida</th>
            <th>Jabatan</th>
            <th>Status</th>
            <th>Aksi</th>
          </tr>
        </thead>

        <tbody>${rows}</tbody>
      </table>
    `,
    footer: renderPagination('anggota', pager)
  });
}

function showMemberCard(id) {
  const item = getItem('anggota', id);

  openCustomModal(
    'Kartu Anggota',
    'Pratinjau identitas anggota.',
    `
      <div class="member-card">

        <div class="card-top">
          <div class="card-org">
            <div class="brand-mark">
              <span class="material-symbols-rounded">flight</span>
            </div>

            <div>
              <strong>SAKA DIRGANTARA</strong>
              <small>KARTU ANGGOTA</small>
            </div>
          </div>

          <div class="card-type">MEMBER ID</div>
        </div>

        <h2>${escapeHtml(item.Nama || '-')}</h2>
        <div class="nta">NTA: ${escapeHtml(item.NTA || '-')}</div>

        <div class="card-bottom">
          <div>
            <small>Krida</small>
            <strong>${escapeHtml(item.Krida || '-')}</strong>
          </div>

          <div>
            <small>Jabatan</small>
            <strong>${escapeHtml(item.Jabatan || 'Anggota')}</strong>
          </div>

          <div>
            <small>Status</small>
            <strong>${escapeHtml(item.Status || '-')}</strong>
          </div>
        </div>

      </div>
    `
  );
}


/* =====================================================
   KEGIATAN
===================================================== */

function renderKegiatan() {
  const data = sortLatestFirst(filterData(
    state.data.kegiatan,
    [
      'NamaKegiatan',
      'Jenis',
      'Lokasi',
      'PenanggungJawab',
      'Status'
    ]
  ));

  const canManage = canPermission('kegiatan','edit') || canPermission('kegiatan','create') || canPermission('kegiatan','delete');
  const canQr = isAnyRole(['ADMIN','PENGURUS']);
  const pager = paginateData('kegiatan', data);

  const rows = pager.items.length
    ? pager.items.map(item => `
      <tr>
        <td>
          <span class="table-name">
            ${escapeHtml(item.NamaKegiatan)}
          </span>
        </td>

        <td>${formatDate(item.Tanggal)}</td>
        <td>${escapeHtml(item.Jenis || '-')}</td>
        <td>${escapeHtml(item.Lokasi || '-')}</td>
        <td>${statusBadge(item.Status)}</td>

        <td>
          <div class="row-actions">

            ${canManage && String(item.Status || '').trim() === 'Rencana' ? `
              <button
                class="btn-icon gold"
                title="Mulai kegiatan dan buka absensi"
                onclick="startKegiatan('${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">play_circle</span>
              </button>
            ` : ''}

            ${canQr && String(item.Status || '').trim() === 'Berjalan' ? `
              <button
                class="btn-icon gold"
                title="QR Absensi"
                onclick="showQr('${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">qr_code_2</span>
              </button>
            ` : ''}

            ${canQr && ['Rencana','Berjalan'].includes(String(item.Status || '').trim()) ? `
              <button
                class="btn-icon izin-link-button"
                title="Link Izin Anggota"
                onclick="showIzinLink('${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">approval</span>
              </button>
            ` : ''}

            ${canQr && String(item.Status || '').trim() === 'Berjalan' ? `
              <button
                class="btn-icon"
                title="Absensi massal"
                onclick="openBatchAbsensi('${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">playlist_add_check</span>
              </button>
            ` : ''}

            ${canManage && String(item.Status || '').trim() === 'Berjalan' ? `
              <button
                class="btn-icon success"
                title="Selesaikan kegiatan"
                onclick="finishKegiatan('${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">task_alt</span>
              </button>
            ` : ''}

            ${canManage && ['Rencana','Berjalan'].includes(String(item.Status || '').trim()) ? `
              <button
                class="btn-icon delete"
                title="Batalkan kegiatan"
                onclick="openCancelKegiatanModal('${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">block</span>
              </button>
            ` : ''}

            <button
              class="btn-icon"
              title="Detail & Laporan Kegiatan"
              onclick="openKegiatanDetail('${escapeJs(item.ID)}')"
            >
              <span class="material-symbols-rounded">description</span>
            </button>

            ${canManage && ['Rencana','Berjalan'].includes(String(item.Status || '').trim()) ? `
              <button ${canPermission('kegiatan', 'edit') ? '' : 'disabled'}
                class="btn-icon"
                title="Edit metadata kegiatan"
                onclick="openForm('kegiatan','${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">edit</span>
              </button>
            ` : ''}

            ${canManage && String(item.Status || '').trim() === 'Rencana' ? `
              <button ${canPermission('kegiatan', 'delete') ? '' : 'disabled'}
                class="btn-icon delete"
                title="Hapus kegiatan Rencana"
                onclick="deleteItem('kegiatan','${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">delete</span>
              </button>
            ` : ''}

          </div>
        </td>
      </tr>
    `).join('')
    : emptyTableRow(6, 'Belum ada kegiatan.');

  renderManagementPage({
    title: 'Manajemen Kegiatan',
    description: 'Rencana → Berjalan → Selesai. Link izin aktif otomatis H-3 sampai Hari H pukul 07.00. Saat kegiatan selesai, absensi yang belum tercatat difinalkan otomatis. Verifikasi izin maksimal pukul 16.00 Hari H.',
    type: 'kegiatan',
    addLabel: 'Tambah Kegiatan',
    canAdd: canManage,
    table: `
      <table class="data-table">
        <thead>
          <tr>
            <th>Kegiatan</th>
            <th>Tanggal</th>
            <th>Jenis</th>
            <th>Lokasi</th>
            <th>Status</th>
            <th>Aksi</th>
          </tr>
        </thead>

        <tbody>${rows}</tbody>
      </table>
    `,
    footer: renderPagination('kegiatan', pager)
  });
}

async function showIzinLink(id) {
  openCustomModal(
    'Link Izin Anggota',
    'Memeriksa jadwal pengajuan izin...',
    loadingHtml(170),
    true
  );

  try {
    const result = await serverCall('getIzinLink', state.token, id);
    const activeClass = result.canSubmit ? 'is-active' : 'is-inactive';
    document.getElementById('modalSubtitle').textContent =
      `${result.kegiatan} • ${formatDate(result.tanggal)}`;
    document.getElementById('modalBody').innerHTML = `
      <div class="izin-link-card">
        <div class="izin-window-status ${activeClass}">
          <span class="material-symbols-rounded">${result.canSubmit ? 'event_available' : 'schedule'}</span>
          <div>
            <strong>${escapeHtml(result.state || '-')}</strong>
            <p>${escapeHtml(result.message || '')}</p>
          </div>
        </div>

        <div class="izin-window-grid">
          <div><span>Mulai otomatis</span><strong>${escapeHtml(result.openAt || '-')}</strong></div>
          <div><span>Berakhir otomatis</span><strong>${escapeHtml(result.autoCloseAt || '-')}</strong></div>
        </div>

        ${result.manualClosedAt ? `
          <div class="izin-manual-close-note">
            Ditutup ${escapeHtml(result.manualClosedAt)} oleh ${escapeHtml(result.manualClosedBy || 'Pengurus')}.
          </div>
        ` : ''}

        <label class="izin-link-label">LINK PENGAJUAN IZIN</label>
        <div class="qr-link izin-url-box">${escapeHtml(result.url)}</div>

        <div class="izin-link-actions">
          <button class="btn btn-primary" type="button" onclick="copyText('${escapeJs(result.url)}')">
            <span class="material-symbols-rounded">content_copy</span> Salin Link
          </button>
          <button class="btn btn-light" type="button" onclick="window.open('${escapeJs(result.url)}','_blank','noopener')">
            <span class="material-symbols-rounded">open_in_new</span> Buka Link
          </button>
          ${result.canSubmit && !result.manualClosedAt ? `
            <button class="btn btn-danger" type="button" onclick="closeKegiatanIzin('${escapeJs(id)}')">
              <span class="material-symbols-rounded">lock</span> Tutup Izin Sekarang
            </button>
          ` : ''}
        </div>

        <div class="login-note">
          Pengajuan hanya dapat dikirim mulai H-3 pukul 00.00 sampai Hari H pukul 07.00, kecuali ditutup lebih awal oleh Pengurus/Admin. Verifikasi oleh Pengurus maksimal pukul 16.00 Hari H.
        </div>
      </div>`;
  } catch (error) {
    const body = document.getElementById('modalBody');
    if (body) body.innerHTML = errorBox(error.message);
  }
}

async function closeKegiatanIzin(id) {
  if (!await customConfirm('Tutup pengajuan izin untuk kegiatan ini sekarang? Setelah ditutup, anggota tidak dapat mengirim izin baru.', 'Tutup Pengajuan Izin', 'Tutup Pengajuan')) return;
  try {
    const result = await serverCall('closeIzinKegiatan', state.token, id);
    applyLocalMutation('kegiatan', result, id, false);
    if (state.view === 'kegiatan') renderKegiatan();
    showToast('Pengajuan izin kegiatan telah ditutup.', 'success');
    await showIzinLink(id);
  } catch (error) {
    showToast(error.message, 'error');
  }
}

async function showQr(id) {
  openCustomModal(
    'QR Absensi',
    'Mempersiapkan QR Code...',
    loadingHtml(160),
    true
  );

  try {
    const result = await serverCall(
      'getAttendanceLink',
      state.token,
      id
    );

    document.getElementById('modalSubtitle').textContent =
      `${result.kegiatan} • ${formatDate(result.tanggal)}`;

    document.getElementById('modalBody').innerHTML = `
      <div class="qr-card">
        <img src="${escapeHtml(result.qrUrl)}" alt="QR Absensi">

        <h3>Scan untuk Check-in</h3>

        <p>
          Anggota mengisi NTA setelah membuka tautan QR.
        </p>

        <div class="qr-link">
          ${escapeHtml(result.url)}
        </div>

        <div style="margin-top:14px;">
          <button
            class="btn btn-primary"
            onclick="copyText('${escapeJs(result.url)}')"
          >
            <span class="material-symbols-rounded">content_copy</span>
            Salin Link
          </button>
        </div>
      </div>
    `;

  } catch (error) {
    document.getElementById('modalBody').innerHTML =
      errorBox(error.message);
  }
}

async function startKegiatan(id) {
  const item = getItem('kegiatan', id);
  if (!item || !item.ID) {
    showToast('Kegiatan tidak ditemukan pada data saat ini.', 'error');
    return;
  }

  if (!await customConfirm(`Mulai “${escapeHtml(item.NamaKegiatan || 'kegiatan ini')}” dan buka absensi sekarang?`, 'Mulai Kegiatan', 'Mulai')) {
    return;
  }

  openCustomModal(
    'Memulai kegiatan',
    'Mengubah status Rencana menjadi Berjalan dan menyiapkan QR Code...',
    loadingHtml(180),
    true
  );

  try {
    const result = await serverCall(
      'transitionKegiatanStatus',
      state.token,
      id,
      'Berjalan',
      ''
    );
    closeModal();
    finishLocalMutation('kegiatan', result, id);
    showToast('Kegiatan dimulai. Absensi dan QR sekarang aktif.', 'success');
    await showQr(id);
  } catch (error) {
    const body = document.getElementById('modalBody');
    if (body) body.innerHTML = errorBox(error.message);
  }
}

async function finishKegiatan(id) {
  const item = getItem('kegiatan', id);
  if (!item || !item.ID) {
    showToast('Kegiatan tidak ditemukan pada data saat ini.', 'error');
    return;
  }

  if (!await customConfirm(`Selesaikan “${escapeHtml(item.NamaKegiatan || 'kegiatan ini')}”? Setelah selesai, QR dan input absensi baru akan ditutup. Anggota Aktif/Calon Anggota yang belum memiliki catatan absensi akan otomatis difinalkan. Izin yang masih menunggu verifikasi tetap dapat mengubah Alpa otomatis menjadi Izin/Sakit sampai pukul 16.00 Hari H.`, 'Selesaikan Kegiatan', 'Selesaikan')) {
    return;
  }

  try {
    const result = await serverCall(
      'transitionKegiatanStatus',
      state.token,
      id,
      'Selesai',
      ''
    );
    finishLocalMutation('kegiatan', result, id);
    markModuleDirty('absensi');
    const auto = result && result.AutoAbsensi ? result.AutoAbsensi : null;
    const summary = auto ? ` ${Number(auto.created || 0)} catatan absensi dibuat otomatis (${Number(auto.izin || 0)} Izin, ${Number(auto.sakit || 0)} Sakit, ${Number(auto.alpa || 0)} Alpa).` : '';
    showToast('Kegiatan selesai. QR dan check-in telah ditutup.' + summary, 'success');
  } catch (error) {
    showToast(error.message, 'error');
  }
}

function openCancelKegiatanModal(id) {
  const item = getItem('kegiatan', id);
  if (!item || !item.ID) {
    showToast('Kegiatan tidak ditemukan pada data saat ini.', 'error');
    return;
  }

  openCustomModal(
    'Batalkan Kegiatan',
    `${item.NamaKegiatan || 'Kegiatan'} • ${formatDate(item.Tanggal)}`,
    `
      <form class="form-grid" onsubmit="submitCancelKegiatan(event,'${escapeJs(id)}')">
        <div class="kegiatan-cancel-warning full">
          <span class="material-symbols-rounded">warning</span>
          <div>
            <strong>Status Dibatalkan bersifat final</strong>
            <span>QR/check-in akan ditutup dan kegiatan tidak dapat dikembalikan ke Rencana atau Berjalan.</span>
          </div>
        </div>
        <div class="form-group full">
          <label>Alasan Pembatalan</label>
          <textarea class="form-control" name="alasan" maxlength="500" required placeholder="Tuliskan alasan pembatalan..."></textarea>
        </div>
        <div class="form-actions full">
          <button type="button" class="btn btn-light" onclick="closeModal()">Kembali</button>
          <button type="submit" class="btn btn-danger" id="cancelKegiatanButton">
            <span class="material-symbols-rounded">block</span>
            Batalkan Kegiatan
          </button>
        </div>
      </form>
    `,
    true
  );
}

async function submitCancelKegiatan(event, id) {
  event.preventDefault();
  const form = event.target;
  const button = document.getElementById('cancelKegiatanButton');
  const alasan = String(new FormData(form).get('alasan') || '').trim();
  if (!alasan) {
    showToast('Alasan pembatalan wajib diisi.', 'warning');
    return;
  }

  if (button) {
    button.disabled = true;
    button.textContent = 'Membatalkan...';
  }

  try {
    const result = await serverCall(
      'transitionKegiatanStatus',
      state.token,
      id,
      'Dibatalkan',
      alasan
    );
    closeModal();
    finishLocalMutation('kegiatan', result, id);
    showToast('Kegiatan dibatalkan dan check-in ditutup.', 'success');
  } catch (error) {
    showToast(error.message, 'error');
    if (button) {
      button.disabled = false;
      button.innerHTML = '<span class="material-symbols-rounded">block</span> Batalkan Kegiatan';
    }
  }
}


/* =====================================================
   ABSENSI
===================================================== */

function renderAbsensi() {
  const data = sortLatestFirst(filterData(
    state.data.absensi,
    [
      'KegiatanNama',
      'NamaAnggota',
      'StatusKehadiran',
      'Metode',
      'Catatan'
    ]
  ));

  const canManage = canPermission('absensi','edit') || canPermission('absensi','create') || canPermission('absensi','delete');
  const mode = getPaginationState('absensi').mode || 'rows';
  let visibleItems = [];
  let pager;
  let meetingBanner = '';

  if (mode === 'meeting') {
    const groups = groupAbsensiByMeeting(data);
    pager = paginateMeetingGroups(groups);
    const group = pager.currentGroup;
    visibleItems = group ? group.items : [];

    if (group) {
      meetingBanner = `
        <div class="meeting-page-banner">
          <div>
            <span class="meeting-page-label">Pertemuan ${pager.page} dari ${pager.totalPages}</span>
            <strong>${escapeHtml(group.name || 'Kegiatan')}</strong>
          </div>
          <div class="meeting-page-meta">
            <span class="material-symbols-rounded">calendar_month</span>
            ${formatDate(group.date)}
            <span class="meeting-page-count">${group.items.length} catatan absensi</span>
          </div>
        </div>
      `;
    }
  } else {
    pager = paginateData('absensi', data);
    visibleItems = pager.items;
  }

  const rows = visibleItems.length
    ? visibleItems.map(item => `
      <tr>
        <td>${formatDate(item.Tanggal)}</td>

        <td>
          ${escapeHtml(item.KegiatanNama || '-')}
        </td>

        <td>
          <span class="table-name">
            ${escapeHtml(item.NamaAnggota || '-')}
          </span>
        </td>

        <td>${attendanceBadge(item.StatusKehadiran)}</td>

        <td>
          ${item.Metode === 'QR'
            ? `<span class="badge purple">QR</span>`
            : `<span class="badge gray">Manual</span>`
          }
        </td>

        <td>${escapeHtml(item.Catatan || '-')}</td>

        <td>
          ${canManage ? `
            <div class="row-actions">
              <button ${canPermission('absensi', 'edit') ? '' : 'disabled'}
                class="btn-icon"
                title="Edit absensi"
                onclick="openForm('absensi','${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">edit</span>
              </button>

              <button ${canPermission('absensi', 'delete') ? '' : 'disabled'}
                class="btn-icon delete"
                title="Hapus absensi"
                onclick="deleteItem('absensi','${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">delete</span>
              </button>
            </div>
          ` : '-'}
        </td>
      </tr>
    `).join('')
    : emptyTableRow(7, 'Belum ada data absensi.');

  renderManagementPage({
    title: 'Absensi Anggota',
    description: mode === 'meeting'
      ? 'Tampilan per pertemuan. Pertemuan terbaru selalu berada di halaman pertama.'
      : 'Tanggal terbaru dan input terbaru tampil paling atas.',
    type: 'absensi',
    addLabel: 'Input Absensi',
    canAdd: false,
    toolbarExtra: `${canManage ? `
      <button class="btn btn-primary batch-attendance-launch" type="button" onclick="openBatchAbsensi()">
        <span class="material-symbols-rounded">playlist_add_check</span>
        Absensi Massal
      </button>
      <button ${canPermission('absensi', 'create') ? '' : 'disabled'} class="btn btn-light" type="button" onclick="openForm('absensi')" title="Untuk satu anggota, koreksi khusus, atau mencatat Libur">
        <span class="material-symbols-rounded">person_check</span>
        Input Satuan / Libur
      </button>
    ` : ''}${renderAbsensiViewControl()}`,
    table: `
      ${meetingBanner}
      <table class="data-table">
        <thead>
          <tr>
            <th>Tanggal</th>
            <th>Kegiatan</th>
            <th>Anggota</th>
            <th>Kehadiran</th>
            <th>Metode</th>
            <th>Catatan</th>
            <th>Aksi</th>
          </tr>
        </thead>

        <tbody>${rows}</tbody>
      </table>
    `,
    footer: mode === 'meeting'
      ? renderMeetingPagination(pager)
      : renderPagination('absensi', pager)
  });
}


/* =====================================================
   BATCH ABSENSI
===================================================== */

async function openBatchAbsensi(kegiatanId = '') {
  if (!isAnyRole(['ADMIN','PENGURUS'])) {
    showToast('Anda tidak memiliki akses untuk mengelola absensi.', 'error');
    return;
  }

  openCustomModal(
    'Absensi Massal',
    'Memuat kegiatan dan daftar anggota...',
    loadingHtml(220),
    false
  );

  const modal = document.getElementById('modal');
  if (modal) modal.classList.add('batch-attendance-modal');
  const generation = modalLoadGeneration;

  try {
    if (!await ensureModulesLoaded(['absensi','kegiatan','anggota'])) return;
    if (generation !== modalLoadGeneration) return;

    const running = getRunningKegiatanOptions();
    if (!running.length) {
      document.getElementById('modalSubtitle').textContent = 'Tidak ada kegiatan aktif';
      document.getElementById('modalBody').innerHTML = emptyStateHtml(
        'event_busy',
        'Belum ada kegiatan berstatus Berjalan. Mulai kegiatan terlebih dahulu sebelum menggunakan absensi massal.'
      );
      return;
    }

    const requested = String(kegiatanId || '').trim();
    const selected = running.some(item => String(item.ID) === requested)
      ? requested
      : String(running[0].ID);

    renderBatchAbsensiModal(selected);
  } catch (error) {
    if (generation !== modalLoadGeneration) return;
    const body = document.getElementById('modalBody');
    if (body) body.innerHTML = errorBox(error.message);
  }
}

function getRunningKegiatanOptions() {
  return [...(state.data.kegiatan || [])]
    .filter(item => String(item.Status || '').trim().toLowerCase() === 'berjalan')
    .sort((a, b) => String(b.Tanggal || '').localeCompare(String(a.Tanggal || '')));
}

function getBatchAttendanceMembers() {
  return [...(state.data.anggota || [])]
    .filter(item => ['aktif','calon anggota'].includes(String(item.Status || '').trim().toLowerCase()))
    .sort((a, b) => String(a.Nama || '').localeCompare(String(b.Nama || ''), 'id'));
}

async function renderBatchAbsensiModal(kegiatanId) {
  const izinGeneration = ++batchIzinLoadGeneration;
  const kegiatan = getItem('kegiatan', kegiatanId);
  const running = getRunningKegiatanOptions();
  const members = getBatchAttendanceMembers();
  const body = document.getElementById('modalBody');

  if (!body || !kegiatan || String(kegiatan.Status || '').trim() !== 'Berjalan') {
    if (body) body.innerHTML = errorBox('Kegiatan tidak tersedia untuk absensi massal.');
    return;
  }

  body.innerHTML = loadingHtml(160);
  let izinRows = [];
  try {
    izinRows = await serverCall('getBatchIzinVerifikasi', state.token, kegiatanId);
    if (izinGeneration !== batchIzinLoadGeneration) return;
  } catch (error) {
    body.innerHTML = errorBox(error.message);
    return;
  }

  const activityRows = (state.data.absensi || []).filter(item =>
    String(item.KegiatanID || '') === String(kegiatanId || '')
  );
  const liburRow = activityRows.find(item => String(item.StatusKehadiran || '') === 'Libur');

  const izinByMember = {};
  (izinRows || []).forEach(item => {
    const memberId = String(item.AnggotaID || '');
    if (memberId && !izinByMember[memberId]) izinByMember[memberId] = item;
  });

  const existingByMember = {};
  activityRows.forEach(item => {
    if (String(item.StatusKehadiran || '') === 'Libur') return;
    const memberId = String(item.AnggotaID || '');
    if (memberId && !existingByMember[memberId]) existingByMember[memberId] = item;
  });

  const activityOptions = running.map(item => `
    <option value="${escapeHtml(item.ID)}" ${String(item.ID) === String(kegiatanId) ? 'selected' : ''}>
      ${escapeHtml(item.NamaKegiatan)} — ${escapeHtml(formatDate(item.Tanggal))}
    </option>
  `).join('');

  document.getElementById('modalSubtitle').textContent =
    `${kegiatan.NamaKegiatan} • ${formatDate(kegiatan.Tanggal)} • ${members.length} anggota tersedia`;

  if (liburRow) {
    body.innerHTML = `
      <div class="batch-attendance-shell">
        <div class="batch-attendance-selector">
          <label>Kegiatan Berjalan</label>
          <select class="form-control" onchange="renderBatchAbsensiModal(this.value)">${activityOptions}</select>
        </div>
        <div class="batch-attendance-blocked">
          <span class="material-symbols-rounded">event_busy</span>
          <div>
            <strong>Kegiatan sudah ditandai Libur</strong>
            <p>Hapus catatan Libur pada menu Absensi terlebih dahulu sebelum mencatat kehadiran anggota.</p>
          </div>
        </div>
      </div>`;
    return;
  }

  const rows = members.length ? members.map((member, index) => {
    const existing = existingByMember[String(member.ID)] || {};
    const onlineIzin = izinByMember[String(member.ID)] || null;
    const hasExisting = Boolean(existing && existing.ID);
    const verifiedStatus = batchIzinAttendanceStatus(onlineIzin);
    const status = String(existing.StatusKehadiran || verifiedStatus);
    const notePrefix = verifiedStatus === 'Alpa'
      ? 'Pengajuan izin ditolak'
      : (verifiedStatus === 'Sakit' ? 'Sakit terverifikasi' : 'Izin terverifikasi');
    const verificationNote = verifiedStatus === 'Alpa'
      ? String(onlineIzin?.CatatanVerifikasi || onlineIzin?.Alasan || '')
      : String(onlineIzin?.Alasan || '');
    const note = String(existing.Catatan || (verifiedStatus ? `${notePrefix}${verificationNote ? ': ' + verificationNote : ''}` : ''));
    return `
      <tr class="batch-attendance-row" data-member-id="${escapeHtml(member.ID)}" data-online-izin="${onlineIzin ? '1' : '0'}" data-existing-absensi="${hasExisting ? '1' : '0'}">
        <td class="batch-attendance-number">${index + 1}</td>
        <td>
          <span class="table-name">${escapeHtml(member.Nama || '-')}</span>
          <small class="batch-attendance-member-meta">${escapeHtml(member.NTA || '-')} • ${escapeHtml(member.Status || '-')}${onlineIzin && !hasExisting ? ` • <strong data-batch-izin-label>${escapeHtml(verifiedStatus ? (verifiedStatus === 'Alpa' ? 'Izin ditolak → Alpa' : verifiedStatus + ' terverifikasi') : 'Menunggu verifikasi izin')}</strong>` : ''}</small>
        </td>
        <td>
          <select class="form-control batch-attendance-status" data-batch-status onchange="updateBatchAbsensiSummary()">
            <option value="">Belum dicatat</option>
            ${['Hadir','Izin','Sakit','Alpa'].map(value =>
              `<option value="${value}" ${status === value ? 'selected' : ''}>${value}</option>`
            ).join('')}
          </select>
        </td>
        <td>
          <input
            class="form-control batch-attendance-note"
            data-batch-note
            maxlength="500"
            value="${escapeHtml(note)}"
            placeholder="Opsional"
          >
        </td>
      </tr>`;
  }).join('') : emptyTableRow(4, 'Tidak ada anggota Aktif/Calon Anggota yang dapat diabsen.');

  body.innerHTML = `
    <form class="batch-attendance-shell" id="batchAbsensiForm" onsubmit="submitBatchAbsensi(event,'${escapeJs(kegiatanId)}')">
      <div class="batch-attendance-topbar">
        <div class="batch-attendance-selector">
          <label>Kegiatan Berjalan</label>
          <select class="form-control" onchange="renderBatchAbsensiModal(this.value)">${activityOptions}</select>
        </div>
        <div class="batch-attendance-actions">
          <button type="button" class="btn btn-light" onclick="setBatchAbsensiAllStatus('Hadir')">
            <span class="material-symbols-rounded">done_all</span>
            Semua Hadir
          </button>
          <button type="button" class="btn btn-light" onclick="setBatchAbsensiAllStatus('')">
            <span class="material-symbols-rounded">ink_eraser</span>
            Kosongkan Baru
          </button>
        </div>
      </div>

      <div class="batch-attendance-info">
        <span class="material-symbols-rounded">info</span>
        <p>Status tersimpan tetap diprioritaskan. Pengajuan <strong>Disetujui</strong> memprefill Izin/Sakit, sedangkan pengajuan <strong>Ditolak</strong> memprefill Alpa. Saat kegiatan diselesaikan, anggota yang belum memiliki catatan akan difinalkan otomatis. Pengajuan yang masih menunggu dapat mengubah Alpa otomatis menjadi Izin/Sakit jika disetujui sebelum pukul 16.00 Hari H. Setelah pukul 16.00, pengajuan yang belum diverifikasi otomatis ditolak dan status tetap/menjadi Alpa.</p>
      </div>

      <section class="batch-izin-panel" aria-labelledby="batchIzinTitle">
        <h3 id="batchIzinTitle">Verifikasi Izin / Sakit</h3>
        <p>Periksa bukti, lalu setujui atau tolak setiap pengajuan. Batas verifikasi pukul 16.00 Hari H. Keputusan langsung disimpan dan pada kegiatan yang sudah selesai, Alpa otomatis dapat diperbarui menjadi Izin/Sakit sebelum batas waktu tersebut.</p>
        <div class="batch-izin-list">${(izinRows || []).map(renderBatchIzinCard).join('') || '<p>Belum ada pengajuan izin/sakit pada kegiatan ini.</p>'}</div>
      </section>
      <div class="batch-attendance-summary" id="batchAttendanceSummary"></div>

      <div class="table-wrap batch-attendance-table-wrap">
        <table class="data-table batch-attendance-table">
          <thead>
            <tr><th>No</th><th>Anggota</th><th>Status</th><th>Catatan</th></tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>

      <div class="form-actions batch-attendance-footer">
        <button type="button" class="btn btn-light" onclick="closeModal()">Batal</button>
        <button type="submit" class="btn btn-primary" id="saveBatchAbsensiButton" ${members.length ? '' : 'disabled'}>
          <span class="material-symbols-rounded">save</span>
          Simpan Absensi Massal
        </button>
      </div>
    </form>`;

  updateBatchAbsensiSummary();
}

function batchIzinAttendanceStatus(item) {
  if (!item) return '';
  if (item.Status === 'Ditolak') return 'Alpa';
  if (item.Status !== 'Disetujui') return '';
  return item.JenisPengajuan === 'Sakit' ? 'Sakit' : 'Izin';
}

function renderBatchIzinCard(item) {
  const pending = !['Disetujui', 'Ditolak'].includes(String(item.Status || ''));
  return `<article class="batch-izin-card" data-izin-id="${escapeHtml(item.ID)}" data-member-id="${escapeHtml(item.AnggotaID)}" data-jenis="${escapeHtml(item.JenisPengajuan || 'Izin')}">
    <div class="batch-izin-heading"><strong>${escapeHtml(item.NamaAnggota || '-')} · ${escapeHtml(item.NTA || '-')}</strong><span data-izin-badge>${izinVerificationBadge(item.Status)}</span></div>
    <p><strong>${escapeHtml(item.JenisPengajuan || 'Izin')} · ${escapeHtml(item.JenisBukti || '-')}</strong></p>
    <p data-izin-alasan>${escapeHtml(item.Alasan || '')}</p>
    <button type="button" class="btn btn-light" onclick="toggleBatchIzinProof(this)">Lihat / Tutup Bukti</button>
    <div data-izin-proof class="batch-izin-proof" hidden></div>
    <p data-izin-result role="status">${escapeHtml(item.CatatanVerifikasi || '')}</p>
    ${pending && canPermission('absensi', 'edit') ? `<div data-izin-controls>
      <label>Catatan verifikasi (wajib jika ditolak)<textarea class="form-control" data-izin-note maxlength="500" rows="2" placeholder="Tuliskan alasan jika pengajuan ditolak"></textarea></label>
      <div class="batch-izin-buttons"><button type="button" class="btn btn-primary" onclick="verifyBatchIzin(this,'Disetujui')">Setujui</button><button type="button" class="btn btn-light" onclick="verifyBatchIzin(this,'Ditolak')">Tolak</button></div>
    </div>` : (pending ? '<p>Hak edit Absensi diperlukan untuk memverifikasi.</p>' : '')}
  </article>`;
}

async function toggleBatchIzinProof(button) {
  const card = button.closest('[data-izin-id]');
  const box = card.querySelector('[data-izin-proof]');
  if (!box.hidden) { box.hidden = true; return; }
  box.hidden = false;
  if (box.dataset.loaded === '1' || box.dataset.loading === '1') return;
  box.dataset.loading = '1';
  box.textContent = 'Memuat bukti...';
  try {
    const proof = await serverCall('getIzinBuktiPreview', state.token, card.dataset.izinId);
    if (!card.isConnected) return;
    if (!/^data:(application\/pdf|image\/(jpeg|png|webp));base64,/i.test(String(proof.dataUrl || ''))) throw new Error('Format bukti tidak didukung.');
    const url = escapeHtml(proof.dataUrl);
    const name = escapeHtml(proof.name || 'Bukti pengajuan');
    box.innerHTML = proof.mimeType === 'application/pdf'
      ? `<iframe src="${url}" title="${name}"></iframe><a class="btn btn-light" href="${url}" download="${name}">Unduh PDF jika pratinjau tidak tampil</a>`
      : `<img src="${url}" alt="${name}">`;
    box.dataset.loaded = '1';
  } catch (error) { box.textContent = error.message || 'Gagal memuat bukti. Tutup lalu buka lagi untuk mencoba.'; }
  finally { delete box.dataset.loading; }
}

async function verifyBatchIzin(button, decision) {
  const card = button.closest('[data-izin-id]');
  const form = card.closest('form');
  if (form.dataset.verifying === '1' || form.dataset.saving === '1') return;
  if (!canPermission('absensi', 'edit')) { showToast('Hak edit Absensi diperlukan.', 'error'); return; }
  if (!['Disetujui', 'Ditolak'].includes(decision)) return;
  const note = card.querySelector('[data-izin-note]').value.trim();
  const message = card.querySelector('[data-izin-result]');
  if (decision === 'Ditolak' && !note) {
    message.textContent = 'Alasan penolakan wajib diisi.';
    card.querySelector('[data-izin-note]').focus();
    return;
  }
  // Keep all unsaved attendance values in place while verification is saved.
  const controls = [...form.querySelectorAll('input,select,textarea,button')];
  const disabled = controls.map(control => control.disabled);
  form.dataset.verifying = '1';
  controls.forEach(control => { control.disabled = true; });
  message.textContent = 'Menyimpan keputusan...';
  try {
    const saved = await serverCall('verifyIzinKegiatan', state.token, card.dataset.izinId, decision, note);
    if (!saved || saved.Status !== decision) throw new Error('Server belum mengonfirmasi keputusan verifikasi.');
    markModuleDirty('absensi');
    if (!card.isConnected) return;
    card.querySelector('[data-izin-badge]').innerHTML = izinVerificationBadge(saved.Status);
    card.querySelector('[data-izin-controls]').remove();
    const status = batchIzinAttendanceStatus({Status: saved.Status, JenisPengajuan: card.dataset.jenis});
    const row = [...form.querySelectorAll('.batch-attendance-row')].find(el => el.dataset.memberId === card.dataset.memberId);
    if (row && row.dataset.existingAbsensi !== '1') {
      row.querySelector('[data-batch-status]').value = status;
      row.querySelector('[data-batch-note]').value = (status === 'Alpa' ? 'Pengajuan izin ditolak: ' + note : status + ' terverifikasi: ' + card.querySelector('[data-izin-alasan]').textContent).slice(0, 500);
      const label = row.querySelector('[data-batch-izin-label]');
      if (label) label.textContent = status === 'Alpa' ? 'Izin ditolak → Alpa' : status + ' terverifikasi';
    }
    message.textContent = saved.Status + (note ? ': ' + note : '') + (row?.dataset.existingAbsensi === '1' ? '. Absensi yang sudah tersimpan tetap dipertahankan.' : '. Tekan Simpan Absensi Massal untuk menyimpan kehadiran.');
    updateBatchAbsensiSummary();
  } catch (error) { message.textContent = error.message || 'Verifikasi gagal. Silakan coba lagi.'; }
  finally {
    delete form.dataset.verifying;
    controls.forEach((control, index) => { control.disabled = disabled[index]; });
  }
}

function setBatchAbsensiAllStatus(status) {
  document.querySelectorAll('.batch-attendance-row [data-batch-status]').forEach(select => {
    const row = select.closest('.batch-attendance-row');
    if (status === 'Hadir' && row && row.dataset.onlineIzin === '1' && row.dataset.existingAbsensi !== '1') {
      return;
    }
    // Tombol Kosongkan hanya mengosongkan baris yang belum memiliki record tersimpan
    // dan tidak menghapus prefill Izin Online.
    if (!status) {
      if (row && row.dataset.onlineIzin === '1' && row.dataset.existingAbsensi !== '1') return;
      const memberId = row ? String(row.dataset.memberId || '') : '';
      const kegiatanSelect = document.querySelector('.batch-attendance-selector select');
      const kegiatanId = kegiatanSelect ? String(kegiatanSelect.value || '') : '';
      const existing = (state.data.absensi || []).some(item =>
        String(item.KegiatanID || '') === kegiatanId &&
        String(item.AnggotaID || '') === memberId &&
        String(item.StatusKehadiran || '') !== 'Libur'
      );
      if (existing) return;
    }
    select.value = status;
  });
  updateBatchAbsensiSummary();
}

function updateBatchAbsensiSummary() {
  const box = document.getElementById('batchAttendanceSummary');
  if (!box) return;
  const counts = { Hadir: 0, Izin: 0, Sakit: 0, Alpa: 0, kosong: 0 };
  document.querySelectorAll('.batch-attendance-row [data-batch-status]').forEach(select => {
    const value = String(select.value || '');
    if (Object.prototype.hasOwnProperty.call(counts, value)) counts[value]++;
    else counts.kosong++;
  });
  const total = counts.Hadir + counts.Izin + counts.Sakit + counts.Alpa;
  box.innerHTML = `
    <span><strong>${total}</strong> akan disimpan</span>
    <span>Hadir <strong>${counts.Hadir}</strong></span>
    <span>Izin <strong>${counts.Izin}</strong></span>
    <span>Sakit <strong>${counts.Sakit}</strong></span>
    <span>Alpa <strong>${counts.Alpa}</strong></span>
    <span>Belum dicatat <strong>${counts.kosong}</strong></span>`;
}

async function submitBatchAbsensi(event, kegiatanId) {
  event.preventDefault();
  const form = event.target;
  if (form.dataset.verifying === '1' || form.dataset.saving === '1') return;
  const button = document.getElementById('saveBatchAbsensiButton');
  if (button && button.disabled) return;

  const rows = [...document.querySelectorAll('.batch-attendance-row')]
    .map(row => {
      const status = row.querySelector('[data-batch-status]');
      const note = row.querySelector('[data-batch-note]');
      return {
        AnggotaID: String(row.dataset.memberId || ''),
        StatusKehadiran: status ? String(status.value || '') : '',
        Catatan: note ? String(note.value || '').trim() : ''
      };
    })
    .filter(item => item.AnggotaID && item.StatusKehadiran);

  if (!rows.length) {
    showToast('Pilih minimal satu status kehadiran.', 'warning');
    return;
  }

  if (rows.length > 500) {
    showToast('Maksimum 500 anggota per proses absensi massal.', 'error');
    return;
  }

  form.dataset.saving = '1';
  if (button) {
    button.disabled = true;
    button.innerHTML = '<span class="material-symbols-rounded">hourglass_top</span> Menyimpan...';
  }

  try {
    const result = await serverCall('saveAbsensiBatch', state.token, {
      KegiatanID: kegiatanId,
      rows
    });

    (result.rows || []).forEach(item => {
      applyLocalMutation('absensi', item, item.ID, false);
    });
    markModuleClean('absensi');
    markDashboardDirty();
    markPenilaianDirty();

    closeModal();
    if (state.view === 'absensi') renderAbsensi();

    const created = Number(result.created || 0);
    const updated = Number(result.updated || 0);
    const unchanged = Number(result.unchanged || 0);
    showToast(
      `Absensi massal tersimpan: ${created} baru, ${updated} diperbarui${unchanged ? `, ${unchanged} tidak berubah` : ''}.`,
      'success'
    );
  } catch (error) {
    showToast(error.message, 'error');
    if (button) {
      button.disabled = false;
      button.innerHTML = '<span class="material-symbols-rounded">save</span> Simpan Absensi Massal';
    }
  } finally {
    delete form.dataset.saving;
  }
}


/* =====================================================
   PENILAIAN KEAKTIFAN BULANAN (V3.4.0)
===================================================== */

function setPenilaianPeriod(value) {
  const period = String(value || '').trim();
  if (!/^\d{4}-\d{2}$/.test(period)) return;
  if (!state.penilaian) state.penilaian = createPenilaianState();
  if (state.penilaian.period === period && state.penilaian.loaded) return;
  state.penilaian.period = period;
  state.penilaian.loaded = false;
  state.penilaian.dirty = true;
  state.penilaian.data = null;
  state.penilaian.sharePayload = null;
  penilaianLoadRequest = null;
  const pageState = getPaginationState('penilaian');
  pageState.page = 1;
  renderView();
}

function assessmentScoreText(value) {
  return value === null || value === undefined || value === '' ? '-' : Number(value).toFixed(1);
}

function assessmentPredicateBadge(value) {
  const label = String(value || 'Belum Dinilai');
  const key = label.toLowerCase();
  const cls = key === 'sangat aktif'
    ? 'assessment-excellent'
    : key === 'aktif'
      ? 'assessment-active'
      : key === 'cukup aktif'
        ? 'assessment-fair'
        : key === 'perlu penguatan'
          ? 'assessment-support'
          : 'assessment-empty';
  return `<span class="assessment-predicate ${cls}">${escapeHtml(label)}</span>`;
}

function renderPenilaian() {
  const data = state.penilaian && state.penilaian.data;
  if (!data) {
    renderError('Data penilaian belum tersedia.');
    return;
  }

  const canManage = canPermission('penilaian','edit') || canPermission('penilaian','create');
  const isAdmin = isRole('ADMIN');
  const summary = data.summary || {};
  const filtered = (data.rows || []).filter(item => {
    if (!state.search) return true;
    const q = String(state.search || '').toLowerCase();
    return [item.Nama,item.NTA,item.Krida,item.Status,item.predicate,item.review]
      .some(value => String(value || '').toLowerCase().includes(q));
  });
  const pager = paginateData('penilaian', filtered);
  const report = data.report || null;
  const counts = summary.predicateCounts || {};
  const manualComponents = (data.components || []).filter(item =>
    String(item.TipeSumber || '').toUpperCase() === 'MANUAL' &&
    !isSkkComponentClient(item) &&
    String(item.Status || '').toLowerCase() === 'aktif'
  );

  const rows = pager.items.length ? pager.items.map(item => {
    const a = item.attendance || {};
    return `
      <tr>
        <td>
          <span class="table-name">${escapeHtml(item.Nama || '-')}</span>
          <small class="assessment-member-meta">${escapeHtml(item.NTA || '-')} • ${escapeHtml(item.Krida || '-')}</small>
        </td>
        <td><span class="assessment-attendance-mini"><b>${a.Hadir || 0}</b>/<span>${a.Izin || 0}</span>/<span>${a.Sakit || 0}</span>/<span>${a.Alpa || 0}</span></span></td>
        <td>${escapeHtml(assessmentScoreText(a.points))}<small class="assessment-max"> / ${escapeHtml(assessmentScoreText(a.maxPoints))}</small></td>
        <td><strong class="assessment-skk-points">${escapeHtml(assessmentScoreText(item.skkPoints))}</strong></td>
        <td>${escapeHtml(assessmentScoreText(item.otherManualPoints))}</td>
        <td><strong>${escapeHtml(assessmentScoreText(item.totalPoints))}</strong></td>
        <td><span class="assessment-index">${escapeHtml(assessmentScoreText(item.score))}</span><small>/100</small></td>
        <td>${assessmentPredicateBadge(item.predicate)}</td>
        <td>
          <button class="btn-icon" type="button" title="Lihat review" onclick="openPenilaianMemberDetail('${escapeJs(item.ID)}')">
            <span class="material-symbols-rounded">analytics</span>
          </button>
        </td>
      </tr>`;
  }).join('') : emptyTableRow(9, 'Belum ada anggota yang sesuai dengan pencarian.');

  document.getElementById('content').innerHTML = `
    <div class="page-heading management-heading assessment-heading">
      <div>
        <span class="assessment-eyebrow">REKAP BULANAN</span>
        <h2>Penilaian Keaktifan</h2>
        <p>Absensi dihitung otomatis. SKK yang sudah disetor dan diparaf dicentang satu kali untuk mendapatkan poin tambahan.</p>
      </div>
      <div class="page-actions assessment-page-actions">
        ${report && report.PDFUrl ? `<a class="btn btn-light" target="_blank" rel="noopener" href="${escapeHtml(report.PDFUrl)}"><span class="material-symbols-rounded">picture_as_pdf</span> PDF Terakhir</a>` : ''}
        <button class="btn btn-light" type="button" onclick="generatePenilaianPdf()"><span class="material-symbols-rounded">description</span> Generate PDF</button>
        ${report && report.PDFFileID ? `<button class="btn btn-light" type="button" onclick="sharePenilaianPdf()"><span class="material-symbols-rounded">share</span> Bagikan PDF</button><button class="btn btn-primary" type="button" onclick="sharePenilaianWhatsApp()"><span class="material-symbols-rounded">chat</span> Bagikan ke WhatsApp</button>` : ''}
      </div>
    </div>

    <section class="assessment-control-panel">
      <label class="assessment-period-control">
        <span class="material-symbols-rounded">calendar_month</span>
        <span>Periode</span>
        <input type="month" value="${escapeHtml(data.period || state.penilaian.period)}" onchange="setPenilaianPeriod(this.value)">
      </label>
      <div class="assessment-control-actions">
        ${canManage ? `<button class="btn btn-primary" type="button" onclick="openSkkChecklistPicker()"><span class="material-symbols-rounded">checklist</span> Checklist SKK</button>` : ''}
        ${canManage ? `<button class="btn btn-light" type="button" onclick="openPenilaianEntryForm()" ${manualComponents.length ? '' : 'title="Belum ada komponen manual tambahan"'}><span class="material-symbols-rounded">add_circle</span> Poin Lain</button>` : ''}
        ${isAdmin ? `<button class="btn btn-light" type="button" onclick="openKomponenPenilaianManager()"><span class="material-symbols-rounded">tune</span> Atur Komponen</button>` : ''}
      </div>
    </section>

    <div class="assessment-summary-grid">
      <article class="assessment-summary-card"><span>Rata-rata Indeks</span><strong>${escapeHtml(assessmentScoreText(summary.averageScore))}</strong><small>dari 100</small></article>
      <article class="assessment-summary-card"><span>Anggota Dinilai</span><strong>${Number(summary.assessedMembers || 0)}</strong><small>dari ${Number(summary.totalMembers || 0)} anggota</small></article>
      <article class="assessment-summary-card"><span>Pertemuan</span><strong>${Number(data.meetingCount || 0)}</strong><small>pertemuan absensi</small></article>
      <article class="assessment-summary-card"><span>Total Poin</span><strong>${escapeHtml(assessmentScoreText(summary.totalPoints))}</strong><small>${Number(summary.skkCompleted || 0)} checklist SKK bulan ini</small></article>
    </div>

    <section class="assessment-review-card">
      <div class="assessment-review-icon"><span class="material-symbols-rounded">insights</span></div>
      <div>
        <span class="assessment-eyebrow">REVIEW KESELURUHAN</span>
        <p>${escapeHtml(summary.review || 'Belum ada review.')}</p>
        <div class="assessment-review-counts">
          <span><b>${Number(counts['Sangat Aktif'] || 0)}</b> Sangat Aktif</span>
          <span><b>${Number(counts['Aktif'] || 0)}</b> Aktif</span>
          <span><b>${Number(counts['Cukup Aktif'] || 0)}</b> Cukup Aktif</span>
          <span><b>${Number(counts['Perlu Penguatan'] || 0)}</b> Perlu Penguatan</span>
        </div>
      </div>
    </section>

    ${(Number(summary.activeWeightTotal || 0) !== 100) ? `
      <div class="assessment-weight-note">
        <span class="material-symbols-rounded">info</span>
        Total bobot komponen reguler saat ini ${escapeHtml(assessmentScoreText(summary.activeWeightTotal))}. Komponen bonus tidak masuk penormalan; bonus maksimum saat ini +${escapeHtml(assessmentScoreText(summary.bonusMaxTotal || 0))} poin indeks.
      </div>` : ''}

    <div class="panel assessment-table-panel">
      <div class="table-toolbar">
        <div class="search-box">
          <span class="material-symbols-rounded">search</span>
          <input placeholder="Cari anggota / NTA / predikat..." value="${escapeHtml(state.search)}" oninput="setSearch(this.value)">
        </div>
        <div class="table-toolbar-right">
          <span class="assessment-rule-chip">H/I/S/A: ${Number(data.attendanceRules?.Hadir || 0)}/${Number(data.attendanceRules?.Izin || 0)}/${Number(data.attendanceRules?.Sakit || 0)}/${Number(data.attendanceRules?.Alpa || 0)} poin</span>
          <span class="table-version">Version ${escapeHtml(state.data.version || UI_VERSION)}</span>
        </div>
      </div>
      <div class="table-wrap responsive-table-wrap">
        <table class="data-table assessment-table">
          <thead><tr><th>Anggota</th><th>H/I/S/A</th><th>Poin Absensi</th><th>SKK</th><th>Poin Lain</th><th>Total</th><th>Indeks</th><th>Predikat</th><th>Detail</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
      ${renderPagination('penilaian', pager)}
    </div>
  `;
}

function getAssessmentComponentById(id) {
  const data = state.penilaian && state.penilaian.data;
  return ((data && data.components) || []).find(item => String(item.ID) === String(id)) || null;
}


function parseAssessmentGenericRulesClient(item) {
  try {
    const parsed = JSON.parse(String(item && item.AturanJSON || '{}'));
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (_) {
    return {};
  }
}

function isSkkComponentClient(item) {
  if (!item) return false;
  if (String(item.Kode || '').trim().toUpperCase() === 'SKK') return true;
  const rules = parseAssessmentGenericRulesClient(item);
  return String(rules.Kind || '').toUpperCase() === 'SKK_CHECKLIST';
}

function parseAssessmentRulesClient(item) {
  try {
    const parsed = JSON.parse(String(item && item.AturanJSON || '{}'));
    return {
      Hadir: Number(parsed.Hadir ?? 10),
      Izin: Number(parsed.Izin ?? 5),
      Sakit: Number(parsed.Sakit ?? 5),
      Alpa: Number(parsed.Alpa ?? 0)
    };
  } catch (_) {
    return { Hadir: 10, Izin: 5, Sakit: 5, Alpa: 0 };
  }
}

function openKomponenPenilaianManager() {
  if (!isRole('ADMIN')) {
    showToast('Pengaturan komponen hanya tersedia untuk Administrator.', 'error');
    return;
  }
  const data = state.penilaian && state.penilaian.data;
  const components = (data && data.components) || [];
  const rows = components.length ? components.map(item => `
    <tr>
      <td><span class="table-name">${escapeHtml(item.NamaKomponen || '-')}</span><small class="assessment-member-meta">${escapeHtml(item.Kode || '-')}</small></td>
      <td>${escapeHtml(isSkkComponentClient(item) ? 'CHECKLIST SKK' : String(item.TipeSumber || '-'))}</td>
      <td>${isSkkComponentClient(item) ? `+${escapeHtml(assessmentScoreText(item.Bobot))} indeks` : `${escapeHtml(assessmentScoreText(item.Bobot))}%`}</td>
      <td>${escapeHtml(assessmentScoreText(item.PoinMaksimum))}</td>
      <td>${statusBadge(item.Status)}</td>
      <td><div class="row-actions">
        <button class="btn-icon" type="button" title="Edit" onclick="openKomponenPenilaianForm('${escapeJs(item.ID)}')"><span class="material-symbols-rounded">edit</span></button>
        ${String(item.TipeSumber || '').toUpperCase() !== 'ABSENSI' && !isSkkComponentClient(item) ? `<button class="btn-icon delete" type="button" title="Hapus" onclick="deleteKomponenPenilaianClient('${escapeJs(item.ID)}')"><span class="material-symbols-rounded">delete</span></button>` : ''}
      </div></td>
    </tr>`).join('') : emptyTableRow(6, 'Belum ada komponen penilaian.');

  openCustomModal('Komponen Penilaian', 'Atur bobot, poin, dan sumber penilaian. Komponen manual dapat digunakan untuk SKK atau indikator lain.', `
    <div class="assessment-component-toolbar">
      <div><strong>Bobot reguler: ${escapeHtml(assessmentScoreText(data?.summary?.activeWeightTotal || 0))}% • Bonus maks.: +${escapeHtml(assessmentScoreText(data?.summary?.bonusMaxTotal || 0))}</strong><small>SKK diperlakukan sebagai bonus sehingga tidak menurunkan nilai dasar.</small></div>
      <button class="btn btn-primary" type="button" onclick="openKomponenPenilaianForm()"><span class="material-symbols-rounded">add</span> Tambah Komponen</button>
    </div>
    <div class="table-wrap"><table class="data-table"><thead><tr><th>Komponen</th><th>Sumber</th><th>Bobot</th><th>Poin Maks.</th><th>Status</th><th>Aksi</th></tr></thead><tbody>${rows}</tbody></table></div>
  `);
}

function openKomponenPenilaianForm(id = '') {
  if (!isRole('ADMIN')) return;
  const item = id ? getAssessmentComponentById(id) : null;
  const isAttendance = !!(item && String(item.TipeSumber || '').toUpperCase() === 'ABSENSI');
  const isSkk = !!(item && isSkkComponentClient(item));
  const rules = parseAssessmentRulesClient(item || {});
  const genericRules = parseAssessmentGenericRulesClient(item || {});
  const type = isAttendance ? 'ABSENSI' : 'MANUAL';
  const defaultName = isAttendance ? 'Kehadiran' : (isSkk ? 'SKK / Kecakapan Khusus' : '');
  const defaultCode = isAttendance ? 'ABSENSI' : (isSkk ? 'SKK' : '');
  const maxDefault = isAttendance ? 10 : (isSkk ? 25 : 20);
  const weightDefault = isAttendance ? 100 : (isSkk ? 10 : 20);

  openCustomModal(id ? 'Edit Komponen Penilaian' : 'Tambah Komponen Penilaian', isAttendance
    ? 'Atur poin Absensi dan bobotnya dalam Indeks Keaktifan.'
    : isSkk
      ? 'SKK berasal dari checklist butir yang sudah disetor dan diparaf. Bobot SKK berfungsi sebagai bonus maksimum ke indeks.'
      : 'Komponen manual dapat digunakan untuk tugas, prestasi, atau indikator lain.', `
    <form class="form-grid assessment-component-form" onsubmit="submitKomponenPenilaian(event,'${escapeJs(id)}')">
      <input type="hidden" name="TipeSumber" value="${type}">
      <label class="field"><span>Nama Komponen *</span><input class="form-control" name="NamaKomponen" required maxlength="80" value="${escapeHtml(item?.NamaKomponen || defaultName)}"></label>
      <label class="field"><span>Kode *</span><input class="form-control" name="Kode" required maxlength="30" ${(isAttendance || isSkk) ? 'readonly' : ''} value="${escapeHtml(item?.Kode || defaultCode)}" placeholder="Contoh: PRESTASI"></label>
      <label class="field"><span>${isSkk ? 'Bonus Maksimum ke Indeks' : 'Bobot'} *</span><input class="form-control" type="number" min="1" max="100" step="1" name="Bobot" required value="${escapeHtml(item?.Bobot ?? weightDefault)}"><small>${isSkk ? 'Contoh +10: SKK dapat menambah indeks maksimum 10 poin dan tidak menurunkan nilai dasar.' : 'Bobot relatif dalam Indeks 0–100.'}</small></label>
      <label class="field"><span>${isAttendance ? 'Poin Maksimum / Pertemuan' : 'Poin Maksimum / Bulan'} *</span><input class="form-control" type="number" min="1" max="10000" step="0.1" name="PoinMaksimum" required value="${escapeHtml(item?.PoinMaksimum ?? maxDefault)}"></label>
      ${isSkk ? `<label class="field"><span>Poin per Butir SKK *</span><input class="form-control" type="number" min="0.1" max="10000" step="0.1" name="PoinPerButir" required value="${escapeHtml(genericRules.PoinPerButir ?? 5)}"><small>Satu checklist yang baru diverifikasi mendapatkan poin ini satu kali.</small></label>` : ''}
      <label class="field"><span>Urutan</span><input class="form-control" type="number" min="1" max="999" name="Urutan" value="${escapeHtml(item?.Urutan ?? (isAttendance ? 1 : (isSkk ? 2 : 10)))}"></label>
      ${isAttendance
        ? `<label class="field"><span>Status</span><input class="form-control" value="Aktif" readonly><input type="hidden" name="Status" value="Aktif"><small>Absensi adalah komponen dasar dan selalu aktif.</small></label>`
        : `<label class="field"><span>Status *</span><select class="form-control" name="Status" required><option ${String(item?.Status || 'Aktif') === 'Aktif' ? 'selected' : ''}>Aktif</option><option ${String(item?.Status || '') === 'Nonaktif' ? 'selected' : ''}>Nonaktif</option></select></label>`}
      ${isAttendance ? `
        <div class="assessment-rule-fields field-span-2">
          <strong>Poin Status Absensi</strong>
          <div class="assessment-rule-grid">
            <label><span>Hadir</span><input class="form-control" type="number" min="0" step="0.1" name="PoinHadir" value="${rules.Hadir}" required></label>
            <label><span>Izin</span><input class="form-control" type="number" min="0" step="0.1" name="PoinIzin" value="${rules.Izin}" required></label>
            <label><span>Sakit</span><input class="form-control" type="number" min="0" step="0.1" name="PoinSakit" value="${rules.Sakit}" required></label>
            <label><span>Alpa</span><input class="form-control" type="number" min="0" step="0.1" name="PoinAlpa" value="${rules.Alpa}" required></label>
          </div>
        </div>` : ''}
      <label class="field field-span-2"><span>Keterangan</span><textarea class="form-control" name="Keterangan" rows="3" maxlength="500">${escapeHtml(item?.Keterangan || '')}</textarea></label>
      <div class="form-actions field-span-2"><button type="button" class="btn btn-light" onclick="openKomponenPenilaianManager()">Kembali</button><button type="submit" class="btn btn-primary"><span class="material-symbols-rounded">save</span> Simpan Komponen</button></div>
    </form>
  `);
}

async function submitKomponenPenilaian(event, id = '') {
  event.preventDefault();
  const form = event.target;
  const payload = Object.fromEntries(new FormData(form).entries());
  if (id) payload.ID = id;
  const button = form.querySelector('button[type="submit"]');
  if (button) button.disabled = true;
  try {
    await serverCall('saveKomponenPenilaian', state.token, payload);
    markPenilaianDirty();
    penilaianLoadRequest = null;
    await ensurePenilaianLoaded(true);
    showToast('Komponen penilaian berhasil disimpan.', 'success');
    openKomponenPenilaianManager();
    if (state.view === 'penilaian') renderPenilaian();
  } catch (error) {
    showToast(error.message, 'error');
    if (button) button.disabled = false;
  }
}

async function deleteKomponenPenilaianClient(id) {
  const item = getAssessmentComponentById(id);
  if (!item) return;
  if (!await customConfirm(`Hapus komponen “${escapeHtml(item.NamaKomponen || 'ini')}”?`, 'Hapus Komponen', 'Hapus')) return;
  try {
    await serverCall('deleteKomponenPenilaian', state.token, id);
    markPenilaianDirty();
    await ensurePenilaianLoaded(true);
    showToast('Komponen penilaian dihapus.', 'success');
    openKomponenPenilaianManager();
    if (state.view === 'penilaian') renderPenilaian();
  } catch (error) {
    showToast(error.message, 'error');
  }
}

function assessmentDefaultDate(period) {
  const today = new Date();
  const ymd = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
  return ymd.substring(0,7) === period ? ymd : `${period}-01`;
}

function openPenilaianEntryForm(memberId = '', entryId = '') {
  if (!isAnyRole(['ADMIN','PENGURUS'])) return;
  const data = state.penilaian && state.penilaian.data;
  if (!data) return;
  const entry = entryId ? (data.entries || []).find(item => String(item.ID) === String(entryId)) : null;
  const allManual = (data.components || []).filter(item =>
    String(item.TipeSumber || '').toUpperCase() === 'MANUAL' && !isSkkComponentClient(item)
  );
  const activeManual = allManual.filter(item => String(item.Status || '').toLowerCase() === 'aktif');
  const historicalComponent = entry
    ? allManual.find(item => String(item.ID) === String(entry.KomponenID || ''))
    : null;
  const manual = activeManual.slice();
  if (historicalComponent && !manual.some(item => String(item.ID) === String(historicalComponent.ID))) {
    manual.push(historicalComponent);
  }
  if (!manual.length) {
    showToast('Belum ada komponen manual aktif. Admin dapat menambahkan komponen seperti SKK melalui Atur Komponen.', 'error');
    return;
  }
  const selectedMember = entry?.AnggotaID || memberId || '';
  const selectedComponent = entry?.KomponenID || manual[0].ID;
  const selected = manual.find(item => String(item.ID) === String(selectedComponent)) || manual[0];
  const period = data.period || state.penilaian.period;
  const memberOptions = (data.rows || []).map(item => `<option value="${escapeHtml(item.ID)}" ${String(item.ID) === String(selectedMember) ? 'selected' : ''}>${escapeHtml(item.Nama || '-')} — ${escapeHtml(item.NTA || '-')}</option>`).join('');
  const componentOptions = manual.map(item => {
    const inactive = String(item.Status || '').toLowerCase() !== 'aktif';
    const suffix = inactive ? ' • Nonaktif (histori)' : '';
    return `<option value="${escapeHtml(item.ID)}" data-max="${escapeHtml(item.PoinMaksimum)}" ${String(item.ID) === String(selectedComponent) ? 'selected' : ''}>${escapeHtml(item.NamaKomponen || '-')}${escapeHtml(suffix)} (maks. ${escapeHtml(assessmentScoreText(item.PoinMaksimum))}/bulan)</option>`;
  }).join('');

  openCustomModal(entry ? 'Edit Poin Individu' : 'Tambah Poin Individu', 'Gunakan untuk tugas, prestasi, atau komponen tambahan lain. SKK dikelola melalui Checklist SKK.', `
    <form class="form-grid" onsubmit="submitPenilaianEntry(event,'${escapeJs(entryId)}')">
      <label class="field field-span-2"><span>Anggota *</span><select class="form-control" name="AnggotaID" required><option value="">Pilih anggota...</option>${memberOptions}</select></label>
      <label class="field"><span>Komponen *</span><select class="form-control" name="KomponenID" onchange="syncPenilaianEntryComponent(this)" required>${componentOptions}</select><small id="penilaianEntryMaxHelp">Maksimum ${escapeHtml(assessmentScoreText(selected?.PoinMaksimum))} poin/bulan.</small></label>
      <label class="field"><span>Tanggal *</span><input class="form-control" type="date" name="Tanggal" required value="${escapeHtml(entry?.Tanggal || assessmentDefaultDate(period))}"></label>
      <label class="field"><span>Poin *</span><input class="form-control" type="number" min="0" step="0.1" max="${escapeHtml(selected?.PoinMaksimum || 10000)}" name="Poin" required value="${escapeHtml(entry?.Poin ?? '')}"></label>
      <label class="field field-span-2"><span>Catatan</span><textarea class="form-control" name="Catatan" rows="3" maxlength="1000" placeholder="Contoh: Prestasi lomba / kontribusi kegiatan">${escapeHtml(entry?.Catatan || '')}</textarea></label>
      <div class="form-actions field-span-2"><button type="button" class="btn btn-light" onclick="closeModal()">Batal</button><button type="submit" class="btn btn-primary"><span class="material-symbols-rounded">save</span> Simpan Poin</button></div>
    </form>
  `);
}

function syncPenilaianEntryComponent(select) {
  const option = select && select.options ? select.options[select.selectedIndex] : null;
  const max = option ? Number(option.dataset.max || 0) : 0;
  const form = select ? select.closest('form') : null;
  const input = form ? form.querySelector('[name="Poin"]') : null;
  if (input && max > 0) input.max = String(max);
  const help = document.getElementById('penilaianEntryMaxHelp');
  if (help) help.textContent = `Maksimum ${assessmentScoreText(max)} poin/bulan.`;
}

async function submitPenilaianEntry(event, id = '') {
  event.preventDefault();
  const form = event.target;
  const payload = Object.fromEntries(new FormData(form).entries());
  if (id) payload.ID = id;
  const period = String(payload.Tanggal || '').substring(0,7);
  if (period !== String(state.penilaian.period || '')) {
    showToast(`Tanggal harus berada pada periode ${state.penilaian.period}.`, 'error');
    return;
  }
  const button = form.querySelector('button[type="submit"]');
  if (button) button.disabled = true;
  try {
    await serverCall('savePenilaianAnggota', state.token, payload);
    closeModal();
    markPenilaianDirty();
    await ensurePenilaianLoaded(true);
    if (state.view === 'penilaian') renderPenilaian();
    showToast('Poin individu berhasil disimpan.', 'success');
  } catch (error) {
    showToast(error.message, 'error');
    if (button) button.disabled = false;
  }
}

async function deletePenilaianEntry(id, memberId = '') {
  if (!await customConfirm('Hapus catatan poin ini?', 'Hapus Catatan', 'Hapus')) return;
  try {
    await serverCall('deletePenilaianAnggota', state.token, id);
    markPenilaianDirty();
    await ensurePenilaianLoaded(true);
    if (memberId) openPenilaianMemberDetail(memberId);
    else closeModal();
    if (state.view === 'penilaian') renderPenilaian();
    showToast('Catatan poin dihapus.', 'success');
  } catch (error) {
    showToast(error.message, 'error');
  }
}

function openPenilaianMemberDetail(memberId) {
  const data = state.penilaian && state.penilaian.data;
  const item = ((data && data.rows) || []).find(row => String(row.ID) === String(memberId));
  if (!item) return;
  const entries = (data.entries || []).filter(entry => String(entry.AnggotaID) === String(memberId));
  const skkEntries = entries.filter(entry => String(entry.KomponenKode || '').toUpperCase() === 'SKK');
  const otherEntries = entries.filter(entry => String(entry.KomponenKode || '').toUpperCase() !== 'SKK');
  const canManage = canPermission('penilaian','edit') || canPermission('penilaian','create');
  const componentRows = (item.components || []).map(comp => `
    <tr><td>${escapeHtml(comp.name || '-')}</td><td>${escapeHtml(comp.source || '-')}</td><td>${escapeHtml(assessmentScoreText(comp.points))} / ${escapeHtml(assessmentScoreText(comp.maxPoints))}</td><td>${escapeHtml(assessmentScoreText(comp.index))}</td><td>${comp.mode === 'BONUS' ? `+${escapeHtml(assessmentScoreText(comp.bonusContribution || 0))} / +${escapeHtml(assessmentScoreText(comp.weight))}` : `${escapeHtml(assessmentScoreText(comp.weight))}%`}</td></tr>`).join('');
  const entryRows = otherEntries.length ? otherEntries.map(entry => `
    <tr><td>${escapeHtml(formatDate(entry.Tanggal))}</td><td>${escapeHtml(entry.KomponenNama || '-')}</td><td><strong>${escapeHtml(assessmentScoreText(entry.Poin))}</strong></td><td>${escapeHtml(entry.Catatan || '-')}</td><td>${canManage ? `<div class="row-actions"><button class="btn-icon" onclick="openPenilaianEntryForm('${escapeJs(memberId)}','${escapeJs(entry.ID)}')"><span class="material-symbols-rounded">edit</span></button><button class="btn-icon delete" onclick="deletePenilaianEntry('${escapeJs(entry.ID)}','${escapeJs(memberId)}')"><span class="material-symbols-rounded">delete</span></button></div>` : '-'}</td></tr>`).join('') : emptyTableRow(5, 'Belum ada poin tambahan lain pada periode ini.');

  openCustomModal('Review Keaktifan Anggota', `${item.Nama || '-'} • ${data.periodLabel || ''}`, `
    <div class="assessment-member-overview">
      <div><span>Indeks Keaktifan</span><strong>${escapeHtml(assessmentScoreText(item.score))}<small>/100</small></strong>${assessmentPredicateBadge(item.predicate)}</div>
      <div class="assessment-member-review"><span class="assessment-eyebrow">REVIEW</span><p>${escapeHtml(item.review || '-')}</p></div>
    </div>
    <div class="assessment-detail-stats">
      <span>Hadir <b>${Number(item.attendance?.Hadir || 0)}</b></span><span>Izin <b>${Number(item.attendance?.Izin || 0)}</b></span><span>Sakit <b>${Number(item.attendance?.Sakit || 0)}</b></span><span>Alpa <b>${Number(item.attendance?.Alpa || 0)}</b></span><span>Belum tercatat <b>${Number(item.attendance?.missing || 0)}</b></span>
    </div>
    <h3 class="assessment-modal-title">Skor per Komponen</h3>
    <div class="table-wrap"><table class="data-table"><thead><tr><th>Komponen</th><th>Sumber</th><th>Poin</th><th>Indeks</th><th>Bobot / Bonus</th></tr></thead><tbody>${componentRows || emptyTableRow(5,'Belum ada komponen aktif.')}</tbody></table></div>
    <div class="assessment-skk-summary"><div><span class="assessment-eyebrow">SKK BULAN INI</span><strong>${skkEntries.length} butir • ${escapeHtml(assessmentScoreText(item.skkPoints))} poin</strong></div>${canManage ? `<button class="btn btn-primary btn-compact" onclick="openSkkChecklist('${escapeJs(memberId)}')"><span class="material-symbols-rounded">checklist</span> Checklist SKK</button>` : ''}</div>
    <div class="assessment-modal-section-head"><h3>Poin Tambahan Lain</h3>${canManage ? `<button class="btn btn-light btn-compact" onclick="openPenilaianEntryForm('${escapeJs(memberId)}')"><span class="material-symbols-rounded">add</span> Tambah</button>` : ''}</div>
    <div class="table-wrap"><table class="data-table"><thead><tr><th>Tanggal</th><th>Komponen</th><th>Poin</th><th>Catatan</th><th>Aksi</th></tr></thead><tbody>${entryRows}</tbody></table></div>
  `);
}


function openSkkChecklistPicker() {
  if (!isAnyRole(['ADMIN','PENGURUS'])) return;
  const data = state.penilaian && state.penilaian.data;
  const members = (data && data.rows) || [];
  if (!members.length) {
    showToast('Belum ada anggota yang dapat dipilih.', 'error');
    return;
  }
  const options = members.map(item => `<option value="${escapeHtml(item.ID)}">${escapeHtml(item.Nama || '-')} — ${escapeHtml(item.NTA || '-')} • ${escapeHtml(item.Krida || '-')}</option>`).join('');
  openCustomModal('Checklist SKK Anggota', 'Pilih anggota. Setelah dibuka, centang hanya butir SKK yang fisiknya sudah disetor dan sudah diparaf/ditandatangani.', `
    <form class="form-grid" onsubmit="event.preventDefault(); openSkkChecklist(this.AnggotaID.value)">
      <label class="field field-span-2"><span>Anggota *</span><select class="form-control" name="AnggotaID" required><option value="">Pilih anggota...</option>${options}</select></label>
      <div class="form-actions field-span-2"><button type="button" class="btn btn-light" onclick="closeModal()">Batal</button><button type="submit" class="btn btn-primary"><span class="material-symbols-rounded">checklist</span> Buka Checklist</button></div>
    </form>
  `);
}

function skkKridaMatch(memberKrida, catalogKrida) {
  const a = String(memberKrida || '').toLowerCase();
  const b = String(catalogKrida || '').toLowerCase();
  if (!a || !b) return false;
  if (a.includes('olahraga') && b.includes('olahraga')) return true;
  if (a.includes('pengetahuan') && b.includes('pengetahuan')) return true;
  if (a.includes('jasa') && b.includes('jasa')) return true;
  return a === b;
}

async function openSkkChecklist(memberId) {
  if (!memberId || !isAnyRole(['ADMIN','PENGURUS'])) return;
  openCustomModal('Checklist SKK', 'Memuat master SKK dan riwayat anggota...', loadingHtml(180), false);
  const modal = document.getElementById('modal');
  if (modal) {
    modal.classList.remove('small');
    modal.classList.add('skk-checklist-modal');
  }
  try {
    const payload = await serverCall('getSkkChecklist', state.token, memberId);
    renderSkkChecklist(payload);
  } catch (error) {
    const body = document.getElementById('modalBody');
    if (body) body.innerHTML = errorBox(error.message);
  }
}

function renderSkkChecklist(payload) {
  const member = payload.member || {};
  const component = payload.component || {};
  const catalog = Array.isArray(payload.catalog) ? payload.catalog : [];
  const completion = Array.isArray(payload.completion) ? payload.completion : [];
  const completedMap = new Map(completion.map(item => [String(item.ReferensiID || ''), item]));
  const validCatalog = catalog.filter(item =>
    String(item.ID || '').trim() &&
    String(item.Krida || '').trim() &&
    String(item.KelompokSKK || '').trim() &&
    String(item.Butir || '').trim()
  );
  const kridas = [...new Set(validCatalog.map(item => String(item.Krida || '').trim()).filter(Boolean))];
  const preferred = kridas.find(k => skkKridaMatch(member.Krida, k)) || '';
  const period = String(state.penilaian?.period || currentPenilaianPeriod());
  const defaultDate = assessmentDefaultDate(period);

  const modal = document.getElementById('modal');
  if (modal) {
    modal.classList.remove('small');
    modal.classList.add('skk-checklist-modal');
  }

  const body = document.getElementById('modalBody');
  if (!body) return;

  if (!validCatalog.length) {
    body.innerHTML = `
      <div class="skk-empty-master">
        <span class="material-symbols-rounded">warning</span>
        <h3>Butir SKK belum tersedia</h3>
        <p>MasterSKK tidak berisi butir yang valid. Jalankan <b>System Maintenance → Update Struktur</b>, lalu buka kembali Checklist SKK.</p>
        <button class="btn btn-primary" type="button" onclick="closeModal(); switchViewByName('maintenance')">
          <span class="material-symbols-rounded">build</span> Buka System Maintenance
        </button>
      </div>`;
    return;
  }

  const grouped = {};
  validCatalog.forEach(item => {
    const krida = String(item.Krida || 'Lainnya').trim();
    const group = String(item.KelompokSKK || 'SKK').trim();
    if (!grouped[krida]) grouped[krida] = {};
    if (!grouped[krida][group]) grouped[krida][group] = [];
    grouped[krida][group].push(item);
  });

  const sections = kridas.map(krida => {
    const groupHtml = Object.keys(grouped[krida] || {}).map(group => {
      const items = grouped[krida][group] || [];
      const done = items.filter(item => completedMap.has(String(item.ID))).length;
      const itemHtml = items.map(item => {
        const existing = completedMap.get(String(item.ID));
        const checked = !!existing;
        const meta = checked
          ? `${formatDate(existing.Tanggal)} • ${assessmentScoreText(existing.Poin)} poin • ${escapeHtml(existing.Petugas || 'Terverifikasi')}`
          : `Belum diverifikasi • +${assessmentScoreText(component.PoinPerButir || 0)} poin saat dicentang`;
        return `
          <label class="skk-check-item"
                 data-skk-search="${escapeHtml((String(item.Kode || '') + ' ' + String(item.KelompokSKK || '') + ' ' + String(item.Butir || '')).toLowerCase())}"
                 data-skk-krida="${escapeHtml(krida)}"
                 data-skk-group-name="${escapeHtml(group)}">
            <input type="checkbox" data-skk-id="${escapeHtml(item.ID)}" data-original="${checked ? '1' : '0'}" ${checked ? 'checked' : ''} onchange="updateSkkChecklistCounters()">
            <span class="skk-check-box"><span class="material-symbols-rounded">check</span></span>
            <span class="skk-check-copy"><strong>${escapeHtml(item.Butir || '-')}</strong><small>${escapeHtml(item.Kode || '-')} • ${meta}</small></span>
          </label>`;
      }).join('');

      return `
        <div class="skk-group" data-skk-group data-skk-krida-group="${escapeHtml(krida)}" data-skk-group-name="${escapeHtml(group)}">
          <div class="skk-group-head">
            <span>${escapeHtml(group)}</span>
            <b>${done}/${items.length}</b>
          </div>
          <div class="skk-group-items">${itemHtml}</div>
        </div>`;
    }).join('');

    return `
      <section class="skk-krida-section" data-skk-section="${escapeHtml(krida)}">
        <div class="skk-krida-title">
          <span class="material-symbols-rounded">flight_takeoff</span>
          <h3>${escapeHtml(krida)}</h3>
        </div>
        ${groupHtml}
      </section>`;
  }).join('');

  body.innerHTML = `
    <div class="skk-checklist-shell" data-member-id="${escapeHtml(member.ID || '')}">
      <div class="skk-member-card">
        <div>
          <span class="assessment-eyebrow">ANGGOTA</span>
          <h3>${escapeHtml(member.Nama || '-')}</h3>
          <p>${escapeHtml(member.NTA || '-')} • ${escapeHtml(member.Krida || '-')}</p>
        </div>
        <div class="skk-total-progress">
          <strong id="skkCompletedCount">${Number(payload.stats?.completedItems || 0)}</strong>
          <span>/ ${Number(payload.stats?.totalItems || validCatalog.length)} butir selesai</span>
        </div>
      </div>

      <div class="skk-score-note">
        <span class="material-symbols-rounded">workspace_premium</span>
        <div>
          <strong>+${assessmentScoreText(component.PoinPerButir || 0)} poin per butir terverifikasi</strong>
          <p>Maksimum yang masuk skor bulanan ${assessmentScoreText(component.PoinMaksimum || 0)} poin; bonus indeks maksimum +${assessmentScoreText(component.BonusMaksimum || 0)}. Checklist tetap tersimpan walaupun poin bulan itu sudah mencapai batas.</p>
        </div>
      </div>

      <div class="skk-toolbar">
        <label>
          <span>Tanggal paraf / verifikasi</span>
          <input class="form-control" type="date" id="skkChecklistDate" value="${escapeHtml(defaultDate)}">
        </label>
        <label>
          <span>Krida</span>
          <select class="form-control" id="skkKridaFilter" onchange="filterSkkChecklist()">
            <option value="">Semua Krida (${validCatalog.length} butir)</option>
            ${kridas.map(k => {
              const count = validCatalog.filter(item => String(item.Krida || '').trim() === k).length;
              return `<option value="${escapeHtml(k)}">${escapeHtml(k)} (${count})</option>`;
            }).join('')}
          </select>
        </label>
        <label class="skk-search-field">
          <span>Cari Butir</span>
          <input class="form-control" id="skkSearchInput" placeholder="Navigasi, meteorologi, SAR..." oninput="filterSkkChecklist()">
        </label>
      </div>

      ${preferred ? `
        <div class="skk-preferred-bar">
          <span>Krida anggota terdeteksi: <b>${escapeHtml(preferred)}</b></span>
          <button class="btn btn-light btn-compact" type="button" onclick="applyMemberKridaFilter('${escapeJs(preferred)}')">
            Tampilkan Krida Anggota
          </button>
          <button class="btn btn-light btn-compact" type="button" onclick="showAllSkkItems()">
            Tampilkan Semua
          </button>
        </div>` : `
        <div class="skk-preferred-bar">
          <span>Semua master SKK ditampilkan. Pilih Krida untuk mempersempit daftar.</span>
          <button class="btn btn-light btn-compact" type="button" onclick="showAllSkkItems()">Tampilkan Semua</button>
        </div>`}

      <div id="skkVisibleInfo" class="skk-visible-info">${validCatalog.length} butir ditampilkan</div>
      <div id="skkNoResults" class="skk-no-results" style="display:none">
        <span class="material-symbols-rounded">search_off</span>
        <strong>Tidak ada butir yang cocok</strong>
        <small>Ubah filter Krida atau kata pencarian.</small>
      </div>

      <div class="skk-checklist-list">${sections}</div>

      <div class="skk-savebar">
        <div>
          <strong id="skkChangedCount">0 perubahan</strong>
          <small>Centang hanya setelah buku/form SKK sudah disetor dan diparaf.</small>
        </div>
        <div class="form-actions">
          <button class="btn btn-light" type="button" onclick="closeModal()">Tutup</button>
          <button class="btn btn-primary" type="button" onclick="saveSkkChecklistClient()">
            <span class="material-symbols-rounded">save</span> Simpan Checklist
          </button>
        </div>
      </div>
    </div>`;

  showAllSkkItems();
  updateSkkChecklistCounters();
}

function applyMemberKridaFilter(krida) {
  const select = document.getElementById('skkKridaFilter');
  if (select) select.value = String(krida || '');
  filterSkkChecklist();
}

function showAllSkkItems() {
  const select = document.getElementById('skkKridaFilter');
  const search = document.getElementById('skkSearchInput');
  if (select) select.value = '';
  if (search) search.value = '';
  filterSkkChecklist();
}

function filterSkkChecklist() {
  const q = String(document.getElementById('skkSearchInput')?.value || '').trim().toLowerCase();
  const krida = String(document.getElementById('skkKridaFilter')?.value || '');
  let visibleCount = 0;

  document.querySelectorAll('.skk-check-item').forEach(item => {
    const matchesText = !q || String(item.dataset.skkSearch || '').includes(q);
    const matchesKrida = !krida || String(item.dataset.skkKrida || '') === krida;
    const visible = matchesText && matchesKrida;
    item.style.display = visible ? '' : 'none';
    if (visible) visibleCount += 1;
  });

  document.querySelectorAll('[data-skk-group]').forEach(group => {
    const hasVisible = [...group.querySelectorAll('.skk-check-item')].some(item => item.style.display !== 'none');
    group.style.display = hasVisible ? '' : 'none';
  });

  document.querySelectorAll('[data-skk-section]').forEach(section => {
    const hasVisible = [...section.querySelectorAll('.skk-check-item')].some(item => item.style.display !== 'none');
    section.style.display = hasVisible ? '' : 'none';
  });

  const info = document.getElementById('skkVisibleInfo');
  const empty = document.getElementById('skkNoResults');
  const list = document.querySelector('.skk-checklist-list');
  if (info) info.textContent = `${visibleCount} butir ditampilkan`;
  if (empty) empty.style.display = visibleCount ? 'none' : 'flex';
  if (list) list.style.display = visibleCount ? 'flex' : 'none';
}

function updateSkkChecklistCounters() {
  const boxes = [...document.querySelectorAll('[data-skk-id]')];
  const checked = boxes.filter(box => box.checked).length;
  const changed = boxes.filter(box => (box.checked ? '1' : '0') !== String(box.dataset.original || '0')).length;
  const completed = document.getElementById('skkCompletedCount');
  const changedEl = document.getElementById('skkChangedCount');
  if (completed) completed.textContent = String(checked);
  if (changedEl) changedEl.textContent = `${changed} perubahan`;
}

async function saveSkkChecklistClient() {
  const shell = document.querySelector('.skk-checklist-shell');
  const memberId = shell && shell.dataset.memberId;
  const date = document.getElementById('skkChecklistDate')?.value;
  if (!memberId || !date) {
    showToast('Anggota dan tanggal verifikasi wajib tersedia.', 'error');
    return;
  }
  const boxes = [...document.querySelectorAll('[data-skk-id]')];
  const changedBoxes = boxes.filter(box => (box.checked ? '1' : '0') !== String(box.dataset.original || '0'));
  if (!changedBoxes.length) {
    showToast('Tidak ada perubahan checklist.', 'error');
    return;
  }
  const removed = changedBoxes.filter(box => !box.checked).length;
  if (removed && !await customConfirm(`${removed} checklist yang sebelumnya sudah terverifikasi akan dibatalkan dan poin historisnya dihapus. Lanjutkan?`, 'Batalkan Checklist', 'Batalkan')) return;

  const changes = changedBoxes.map(box => ({ MasterSKKID: box.dataset.skkId, Checked: !!box.checked }));
  const button = document.querySelector('.skk-savebar .btn-primary');
  if (button) button.disabled = true;
  try {
    const result = await serverCall('saveSkkChecklist', state.token, memberId, date, changes);
    closeModal();
    markPenilaianDirty();
    penilaianLoadRequest = null;
    await ensurePenilaianLoaded(true);
    if (state.view === 'penilaian') renderPenilaian();
    showToast(`Checklist SKK disimpan: +${Number(result.created || 0)} butir, -${Number(result.removed || 0)} koreksi.`, 'success');
  } catch (error) {
    showToast(error.message, 'error');
    if (button) button.disabled = false;
  }
}

async function generatePenilaianPdf() {
  const period = state.penilaian && state.penilaian.period;
  if (!period) return;
  const button = document.querySelector('.assessment-page-actions button[onclick="generatePenilaianPdf()"]');
  if (button) { button.disabled = true; button.innerHTML = '<span class="material-symbols-rounded spin">progress_activity</span> Membuat PDF...'; }
  openCustomModal('Membuat PDF Penilaian', 'Mohon tunggu sampai proses selesai.', `<div class="pdf-progress-dialog"><div class="pdf-progress-icon"><span class="material-symbols-rounded spin">picture_as_pdf</span></div><h3 id="pdfProgressTitle">Menyiapkan data penilaian</h3><p id="pdfProgressText">Mengumpulkan rekap anggota, absensi, dan komponen penilaian...</p><div class="pdf-progress-track"><span></span></div><small>Proses berjalan di Google Drive. Jangan tutup halaman ini.</small></div>`, true);
  const progressTimer = setTimeout(() => { const t=document.getElementById('pdfProgressTitle'); const x=document.getElementById('pdfProgressText'); if(t) t.textContent='Menyusun dokumen PDF'; if(x) x.textContent='Membuat tabel penilaian dan menyimpan file ke Google Drive...'; }, 1200);
  try {
    const result = await serverCall('generatePenilaianBulananPdf', state.token, period);
    clearTimeout(progressTimer);
    markPenilaianDirty(); penilaianLoadRequest = null;
    await ensurePenilaianLoaded(true);
    if (state.penilaian && state.penilaian.data) { state.penilaian.data.report = result.report || Object.assign({}, state.penilaian.data.report || {}, { PDFFileID: result.fileId, PDFUrl: result.url, name: result.name }); }
    if (state.penilaian) state.penilaian.sharePayload = null;
    closeModal();
    if (state.view === 'penilaian') renderPenilaian();
    showToast('PDF berhasil dibuat dan disimpan di Google Drive.', 'success');
    showPdfSuccessDialog(result, period);
    serverCall('getPenilaianPdfPayload', state.token, result.fileId).then(payload => { if (state.penilaian && state.penilaian.period === period) state.penilaian.sharePayload = payload; }).catch(() => {});
  } catch (error) {
    clearTimeout(progressTimer); closeModal(); showToast(error.message || 'PDF gagal dibuat.', 'error');
  } finally {
    if (button) { button.disabled = false; button.innerHTML = '<span class="material-symbols-rounded">description</span> Generate PDF'; }
  }
}

function showPdfSuccessDialog(result, period) {
  const url = result && result.url ? result.url : '';
  openCustomModal('PDF Berhasil Dibuat', 'File sudah tersimpan di Google Drive.', `<div class="pdf-success-dialog"><div class="pdf-success-icon"><span class="material-symbols-rounded">task_alt</span></div><h3>PDF siap digunakan</h3><p>${escapeHtml(result.name || ('Skor Keaktifan ' + period + '.pdf'))}</p><div class="pdf-success-actions"><a class="btn btn-light" target="_blank" rel="noopener" href="${escapeHtml(url)}"><span class="material-symbols-rounded">open_in_new</span> Buka PDF</a><button type="button" class="btn btn-primary" onclick="closeModal(); sharePenilaianWhatsApp()"><span class="material-symbols-rounded">chat</span> Ke WhatsApp</button></div></div>`, true);
}

function assessmentBase64ToBlob(payload) {
  const binary = atob(String(payload.base64 || ''));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new Blob([bytes], { type: payload.mimeType || 'application/pdf' });
}

function downloadAssessmentPayload(payload) {
  const blob = assessmentBase64ToBlob(payload);
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = payload.name || 'Skor-Keaktifan.pdf';
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

async function getAssessmentSharePayload() {
  const data = state.penilaian && state.penilaian.data;
  const report = data && data.report;
  if (!report || !report.PDFFileID) {
    throw new Error('Generate PDF terlebih dahulu sebelum dibagikan.');
  }
  if (state.penilaian.sharePayload) return state.penilaian.sharePayload;
  const payload = await serverCall('getPenilaianPdfPayload', state.token, report.PDFFileID);
  state.penilaian.sharePayload = payload;
  return payload;
}

async function sharePenilaianWhatsApp() {
  const data = state.penilaian && state.penilaian.data;
  const report = data && data.report;
  if (!report || !report.PDFFileID) { showToast('Generate PDF terlebih dahulu.', 'error'); return; }
  const defaultMessage = `Assalamu’alaikum/Salam sejahtera.\n\nBerikut kami sampaikan laporan penilaian keaktifan anggota SAKA Dirgantara periode ${data.periodLabel}.\n\nRata-rata indeks: ${assessmentScoreText(data.summary?.averageScore || 0)}/100\nAnggota dinilai: ${data.summary?.assessedMembers || 0} dari ${data.summary?.totalMembers || 0}\n\nPDF laporan dapat dibuka melalui tautan Google Drive berikut:\n${report.PDFUrl || ''}\n\nMohon diperiksa. Terima kasih.`;
  const message = await customTextPrompt('Tinjau dan edit template pemberitahuan sebelum membuka WhatsApp.', 'Template WhatsApp', defaultMessage);
  if (!message || !message.trim()) return;
  window.open('https://wa.me/?text=' + encodeURIComponent(message.trim()), '_blank', 'noopener');
  showToast('WhatsApp dibuka dengan template pemberitahuan.', 'success');
}

function customTextPrompt(message, title, initial) {
  return new Promise(resolve => {
    if (pendingCustomDialog) pendingCustomDialog.resolve(false);
    pendingCustomDialog = { resolve };
    openCustomModal(title || 'Template Pesan', 'Pesan dapat diedit sebelum dibagikan.', `<div class="custom-dialog-content"><div class="custom-dialog-message">${escapeHtml(message)}</div><textarea id="customDialogInput" class="form-control" rows="10" maxlength="4000">${escapeHtml(initial || '')}</textarea><div class="custom-dialog-actions"><button type="button" class="btn btn-light" onclick="resolveCustomDialog('')">Batal</button><button type="button" class="btn btn-primary" onclick="resolveCustomDialog(document.getElementById('customDialogInput')?.value || '')"><span class="material-symbols-rounded">chat</span> Buka WhatsApp</button></div></div>`, false);
  });
}

async function sharePenilaianPdf() {
  const data = state.penilaian && state.penilaian.data;
  if (!data?.report?.PDFFileID) {
    showToast('Generate PDF terlebih dahulu.', 'error');
    return;
  }
  try {
    const payload = await getAssessmentSharePayload();
    const blob = assessmentBase64ToBlob(payload);
    const file = new File([blob], payload.name || `Skor Keaktifan ${data.period}.pdf`, { type: 'application/pdf' });
    const shareData = {
      title: `Skor Keaktifan SAKA Dirgantara - ${data.periodLabel}`,
      text: `Rekap skor keaktifan anggota SAKA Dirgantara periode ${data.periodLabel}.`,
      files: [file]
    };
    if (navigator.share && (!navigator.canShare || navigator.canShare({ files: [file] }))) {
      await navigator.share(shareData);
      return;
    }
    downloadAssessmentPayload(payload);
    const message = `Rekap skor keaktifan SAKA Dirgantara periode ${data.periodLabel}. PDF sudah diunduh; silakan lampirkan file tersebut ke grup.`;
    window.open(`https://wa.me/?text=${encodeURIComponent(message)}`, '_blank', 'noopener');
    showToast('PDF diunduh. WhatsApp dibuka untuk memilih grup dan melampirkan file.', 'success');
  } catch (error) {
    if (String(error && error.name || '') === 'AbortError') return;
    showToast(error.message || 'PDF belum dapat dibagikan.', 'error');
  }
}

async function downloadPenilaianPdf() {
  try {
    const payload = await getAssessmentSharePayload();
    downloadAssessmentPayload(payload);
  } catch (error) {
    showToast(error.message, 'error');
  }
}

/* =====================================================
   KAS
===================================================== */

function renderKas() {
  const period = cashPeriodRows();
  const data = filterData(
    period.rows,
    ['Jenis','Kategori','Keterangan','Petugas','NoBukti']
  );

  data.sort((a,b) => String(b.Tanggal).localeCompare(String(a.Tanggal)));
  const canManage = canPermission('kas','edit') || canPermission('kas','create') || canPermission('kas','delete');
  const kegiatanById = Object.fromEntries((state.data.kegiatan || []).map(item => [String(item.ID), item]));

  const rows = data.length
    ? data.map(item => {
      const kegiatan = kegiatanById[String(item.KegiatanID || '')] || null;
      return `
      <tr>
        <td>${formatDate(item.Tanggal)}</td>
        <td>${cashBadge(item.Jenis)}</td>
        <td>${escapeHtml(item.NoBukti || '-')}</td>
        <td>${escapeHtml(item.Kategori || '-')}</td>
        <td>${escapeHtml(item.Keterangan || '-')}</td>
        <td>${escapeHtml(kegiatan ? kegiatan.NamaKegiatan : '-')}</td>
        <td><strong>${formatRupiah(item.Nominal)}</strong></td>
        <td>${escapeHtml(item.Petugas || '-')}</td>
        <td>${canManage ? `<div class="row-actions"><button ${canPermission('kas', 'edit') ? '' : 'disabled'} class="btn-icon" onclick="openForm('kas','${escapeJs(item.ID)}')"><span class="material-symbols-rounded">edit</span></button><button ${canPermission('kas', 'delete') ? '' : 'disabled'} class="btn-icon delete" onclick="deleteItem('kas','${escapeJs(item.ID)}')"><span class="material-symbols-rounded">delete</span></button></div>` : '-'}</td>
      </tr>`;
    }).join('')
    : emptyTableRow(9, 'Belum ada transaksi kas.');

  renderManagementPage({
    title: 'Kas Organisasi',
    description: 'Pemasukan, pengeluaran, saldo, dan transaksi yang terkait dengan kegiatan.',
    type: 'kas',
    addLabel: 'Tambah Transaksi',
    canAdd: canManage,
    summary: cashSummaryHtml(period),
    toolbarExtra: periodFilterHtml('kas'),
    table: `
      <table class="data-table">
        <thead><tr><th>Tanggal</th><th>Jenis</th><th>No Bukti</th><th>Kategori</th><th>Keterangan</th><th>Kegiatan</th><th>Nominal</th><th>Petugas</th><th>Aksi</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>`
  });
}


/* =====================================================
   INVENTARIS
===================================================== */

function renderInventaris() {
  const usageByInventory = {};
  (state.data.kegiatanInventaris || []).forEach(item => {
    const key = String(item.InventarisID || '');
    if (key) usageByInventory[key] = (usageByInventory[key] || 0) + 1;
  });

  const data = filterData(
    state.data.inventaris,
    [
      'KodeBarang',
      'NamaBarang',
      'Kategori',
      'Kondisi',
      'Lokasi',
      'PenanggungJawab'
    ]
  );

  const canEdit = canPermission('inventaris','edit') || canPermission('inventaris','create');
  const canDelete = canPermission('inventaris','delete');

  const rows = data.length
    ? data.map(item => `
      <tr>
        <td>${escapeHtml(item.KodeBarang || '-')}</td>

        <td>
          <span class="table-name">
            ${escapeHtml(item.NamaBarang)}
          </span>
        </td>

        <td>${escapeHtml(item.Kategori || '-')}</td>
        <td>${escapeHtml(item.Jumlah || 0)}</td>
        <td>${conditionBadge(item.Kondisi)}</td>
        <td>${escapeHtml(item.Lokasi || '-')}</td>
        <td>${escapeHtml(item.PenanggungJawab || '-')}</td>
        <td><span class="inventory-usage-count">${escapeHtml(usageByInventory[String(item.ID)] || 0)} kegiatan</span><br><small>${Number(item.SedangDipakai || 0)} dipakai · ${Number(item.Tersedia || 0)} tersedia</small></td>

        <td>
          <div class="row-actions">
            <button class="btn btn-light btn-compact" onclick="openInventoryQuick('${escapeJs(item.ID)}')">Pakai / Kembali</button>
            ${canEdit ? `
              <button ${canPermission('inventaris', 'edit') ? '' : 'disabled'}
                class="btn-icon"
                onclick="openForm('inventaris','${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">edit</span>
              </button>
            ` : ''}

            ${canDelete ? `
              <button ${canPermission('inventaris', 'delete') ? '' : 'disabled'}
                class="btn-icon delete"
                onclick="deleteItem('inventaris','${escapeJs(item.ID)}')"
              >
                <span class="material-symbols-rounded">delete</span>
              </button>
            ` : ''}
          </div>
        </td>
      </tr>
    `).join('')
    : emptyTableRow(9, 'Belum ada data inventaris.');

  renderManagementPage({
    title: 'Inventaris Organisasi',
    description: 'Pencatatan aset dan perlengkapan SAKA Dirgantara.',
    type: 'inventaris',
    addLabel: 'Tambah Inventaris',
    canAdd: canEdit,
    summary: inventorySummaryHtml(),
    toolbarExtra: '<span class="inventory-page-note">Klik Pakai / Kembali untuk mengelola barang langsung dari sini.</span>',
    table: `
      <table class="data-table">
        <thead>
          <tr>
            <th>Kode</th>
            <th>Barang</th>
            <th>Kategori</th>
            <th>Jumlah</th>
            <th>Kondisi</th>
            <th>Lokasi</th>
            <th>PJ</th>
            <th>Riwayat</th>
            <th>Aksi</th>
          </tr>
        </thead>

        <tbody>${rows}</tbody>
      </table>
    `
  });
}


/* =====================================================
   SURAT
===================================================== */

function renderSurat() {
  const data = filterData(
    state.data.surat,
    [
      'NomorSurat',
      'Jenis',
      'Perihal',
      'AsalTujuan',
      'Status'
    ]
  );

  data.sort((a,b) =>
    String(b.Tanggal).localeCompare(String(a.Tanggal))
  );

  const canEdit = canPermission('surat','edit') || canPermission('surat','create');
  const canDelete = canPermission('surat','delete');

  const renderRows = group => group.length
    ? group.map(item => {
      const linkFile = safeHttpUrl(item.LinkFile);

      return `
        <tr>
          <td>${formatDate(item.Tanggal)}</td>

          <td>
            ${escapeHtml(item.NomorSurat || '-')}
          </td>

          <td>${mailBadge(item.Jenis)}</td>

          <td>
            <span class="table-name">
              ${escapeHtml(item.Perihal || '-')}
            </span>
          </td>

          <td>${escapeHtml(item.AsalTujuan || '-')}</td>
          <td>${statusBadge(item.Status)}</td>

          <td>
            ${linkFile
              ? `<a class="btn-icon" href="${escapeHtml(linkFile)}" target="_blank" rel="noopener noreferrer" title="Buka berkas">
                  <span class="material-symbols-rounded">open_in_new</span>
                </a>`
              : '-'
            }
          </td>

          <td>
            <div class="row-actions">
              ${canEdit ? `
                <button ${canPermission('surat', 'edit') ? '' : 'disabled'}
                  class="btn-icon"
                  onclick="openForm('surat','${escapeJs(item.ID)}')"
                >
                  <span class="material-symbols-rounded">edit</span>
                </button>
              ` : ''}

              ${canDelete ? `
                <button ${canPermission('surat', 'delete') ? '' : 'disabled'}
                  class="btn-icon delete"
                  onclick="deleteItem('surat','${escapeJs(item.ID)}')"
                >
                  <span class="material-symbols-rounded">delete</span>
                </button>
              ` : ''}
            </div>
          </td>
        </tr>
      `;
    }).join('')
    : emptyTableRow(8, 'Belum ada arsip surat.');

  renderManagementPage({
    title: 'Administrasi Surat',
    description: 'Arsip surat masuk, surat keluar, dan dokumen organisasi.',
    type: 'surat',
    addLabel: 'Tambah Surat',
    canAdd: canEdit,
    summary: summaryCardsHtml([
      ['Surat masuk', data.filter(row => row.Jenis === 'Surat Masuk').length],
      ['Surat keluar', data.filter(row => row.Jenis === 'Surat Keluar').length]
    ]),
    toolbarExtra: `<label>Jenis surat <select class="form-control" onchange="state.letterKind=this.value;renderSurat()">${['Semua','Surat Masuk','Surat Keluar'].map(value => `<option ${value === (state.letterKind || 'Semua') ? 'selected' : ''}>${value}</option>`).join('')}</select></label>`,
    table: ['Surat Masuk','Surat Keluar'].filter(kind => !state.letterKind || state.letterKind === 'Semua' || state.letterKind === kind).map(kind => `
      <h3 class="register-table-title">${kind}</h3>
      <table class="data-table">
        <thead><tr><th>Tanggal</th><th>Nomor</th><th>Jenis</th><th>Perihal</th><th>${kind === 'Surat Masuk' ? 'Asal' : 'Tujuan'}</th><th>Status</th><th>File</th><th>Aksi</th></tr></thead>
        <tbody>${renderRows(data.filter(row => row.Jenis === kind))}</tbody>
      </table>`).join('')
  });
}

/* PENGURUS */

function renderPengurus() {
  const canManage = canPermission('pengurus','edit') || canPermission('pengurus','create') || canPermission('pengurus','delete');

  const data = [...state.data.pengurus]
    .filter(item => {
      if (!state.search) return true;

      const haystack = [
        item.Nama,
        item.Jabatan,
        item.Bidang,
        item.Periode
      ].join(' ').toLowerCase();

      return haystack.includes(state.search);
    })
    .sort((a,b) =>
      Number(a.Urutan || 999) - Number(b.Urutan || 999)
    );

  const cards = data.length
    ? data.map(item => `
      <div class="org-card">
        <span class="org-order">#${escapeHtml(item.Urutan || '-')}</span>

        <div class="org-avatar">
          <span class="material-symbols-rounded">person</span>
        </div>

        <h3>${escapeHtml(item.Nama || '-')}</h3>

        <div class="org-role">
          ${escapeHtml(item.Jabatan || '-')}
        </div>

        <div class="org-meta">
          Bidang: ${escapeHtml(item.Bidang || '-')}<br>
          Periode: ${escapeHtml(item.Periode || '-')}<br>
          Status: ${escapeHtml(item.Status || '-')}
        </div>

        ${canManage ? `
          <div class="row-actions" style="margin-top:13px;">
            <button ${canPermission('pengurus', 'edit') ? '' : 'disabled'}
              class="btn-icon"
              onclick="openForm('pengurus','${escapeJs(item.ID)}')"
            >
              <span class="material-symbols-rounded">edit</span>
            </button>

            <button ${canPermission('pengurus', 'delete') ? '' : 'disabled'}
              class="btn-icon delete"
              onclick="deleteItem('pengurus','${escapeJs(item.ID)}')"
            >
              <span class="material-symbols-rounded">delete</span>
            </button>
          </div>
        ` : ''}
      </div>
    `).join('')
    : `
      <div style="grid-column:1/-1;">
        ${emptyStateHtml('account_tree','Belum ada struktur pengurus.')}
      </div>
    `;

  document.getElementById('content').innerHTML = `
    <div class="page-heading">
      <div>
        <h2>Struktur Pengurus</h2>
        <p>
          Susunan kepengurusan dan pembagian bidang organisasi.
        </p>
      </div>

      <div class="page-actions">
        ${canManage ? `
          <button ${canPermission('pengurus', 'create') ? '' : 'disabled'}
            class="btn btn-primary"
            onclick="openForm('pengurus')"
          >
            <span class="material-symbols-rounded">add</span>
            Tambah Pengurus
          </button>
        ` : ''}
      </div>
    </div>

    <div class="panel">
      <div class="table-toolbar">
        <div class="search-box">
          <span class="material-symbols-rounded">search</span>

          <input
            placeholder="Cari pengurus..."
            value="${escapeHtml(state.search)}"
            oninput="setSearch(this.value)"
          >
        </div>
      </div>

      <div class="panel-body">
        <div class="org-grid">
          ${cards}
        </div>
      </div>
    </div>
  `;
}


/* =====================================================
   USERS
===================================================== */

async function openRolePermissionSettings() {
  try {
    const payload = await serverCall('getRolePermissions', state.token);
    const rows = payload.rows || [];
    const modules = payload.modules || [];
    const current = Object.fromEntries(rows.filter(r => r.Role === 'PENGURUS').map(r => [r.Module, r]));
    const label = {anggota:'Anggota',kegiatan:'Kegiatan',absensi:'Absensi',kas:'Kas Organisasi',inventaris:'Inventaris',surat:'Surat',pengurus:'Struktur Pengurus',penilaian:'Penilaian'};
    const body = `<form id="rolePermissionForm"><div class="permission-intro">Tentukan fitur yang dapat diakses oleh akun <strong>PENGURUS</strong>. Perubahan berlaku untuk semua akun Pengurus.</div><div class="table-wrap"><table class="data-table permission-table"><thead><tr><th>Modul</th><th>Lihat</th><th>Tambah</th><th>Edit</th><th>Hapus</th></tr></thead><tbody>${modules.map(m => { const r=current[m]||{}; return `<tr><td><strong>${label[m]||m}</strong></td>${['CanView','CanCreate','CanEdit','CanDelete'].map(k => `<td><input type="checkbox" name="${m}_${k}" ${String(r[k]).toUpperCase()==='TRUE'?'checked':''}></td>`).join('')}</tr>`; }).join('')}</tbody></table></div><div class="custom-dialog-actions"><button type="button" class="btn btn-light" onclick="closeModal()">Batal</button><button type="submit" class="btn btn-primary">Simpan Hak Akses</button></div></form>`;
    openCustomModal('Konfigurasi Hak Akses Pengurus', 'Atur kemampuan lihat, tambah, edit, dan hapus.', body, false);
    document.getElementById('rolePermissionForm').onsubmit = async function(event) { event.preventDefault(); const form=event.target; const out=modules.map(m => ({ID:(current[m]||{}).ID||'',Role:'PENGURUS',Module:m,CanView:!!form.elements[m+'_CanView']?.checked,CanCreate:!!form.elements[m+'_CanCreate']?.checked,CanEdit:!!form.elements[m+'_CanEdit']?.checked,CanDelete:!!form.elements[m+'_CanDelete']?.checked})); try { await serverCall('saveRolePermissions', state.token, out); closeModal(); showToast('Hak akses Pengurus berhasil disimpan.', 'success'); } catch(e) { showToast(e.message,'error'); } };
  } catch (error) { showToast(error.message || 'Gagal memuat hak akses.', 'error'); }
}


async function ensureUserMemberOptionsLoaded(force = false) {
  if (!state.token || !isRole('ADMIN')) return [];
  if (!force && state.userMemberOptionsLoaded) return state.userMemberOptions || [];
  if (!force && state.userMemberOptionsRequest) return state.userMemberOptionsRequest;

  const token = state.token;
  const request = serverCall('getUserMemberOptions', token).then(payload => {
    if (token !== state.token) return [];
    state.userMemberOptions = Array.isArray(payload && payload.rows) ? payload.rows : [];
    state.userMemberOptionsLoaded = true;
    return state.userMemberOptions;
  });

  state.userMemberOptionsRequest = request;
  try {
    return await request;
  } finally {
    if (state.userMemberOptionsRequest === request) state.userMemberOptionsRequest = null;
  }
}

function warmUserMemberOptions() {
  if (state.userMemberOptionsLoaded || state.userMemberOptionsRequest || !isRole('ADMIN')) return;
  ensureUserMemberOptionsLoaded().then(() => {
    if (state.view === 'users') renderUsers();
  }).catch(() => {
    // Daftar pengguna tetap dapat dipakai walau label anggota gagal dimuat.
  });
}


function resolveUserMemberLabel(anggotaId) {
  const id = String(anggotaId || '').trim();
  if (!id) return '-';

  const member = (state.userMemberOptions || []).find(item => String(item.ID || '') === id) ||
    (state.data.anggota || []).find(item => String(item.ID || '') === id);

  if (member) return `${member.Nama || '-'}${member.NTA ? ' • ' + member.NTA : ''}`;
  return state.userMemberOptionsLoaded ? 'Data anggota tidak ditemukan' : 'Anggota tertaut';
}


function renderUsers() {
  if (!isRole('ADMIN')) {
    switchViewByName('dashboard');
    return;
  }

  const data = filterData(
    state.data.users,
    ['Username','Nama','Role','Status','AnggotaID']
  );

  const rows = data.length
    ? data.map(item => `
      <tr>
        <td>
          <span class="table-name">
            ${escapeHtml(item.Nama)}
          </span>
        </td>

        <td>${escapeHtml(item.Username)}</td>
        <td><span class="badge gold">${escapeHtml(item.Role)}</span></td>
        <td>${escapeHtml(resolveUserMemberLabel(item.AnggotaID))}</td>
        <td>${statusBadge(item.Status)}</td>

        <td>
          <div class="row-actions">
            <button ${canPermission('users', 'edit') ? '' : 'disabled'}
              class="btn-icon"
              onclick="openForm('users','${escapeJs(item.ID)}')"
            >
              <span class="material-symbols-rounded">edit</span>
            </button>

            <button ${canPermission('users', 'delete') ? '' : 'disabled'}
              class="btn-icon delete"
              onclick="deleteItem('users','${escapeJs(item.ID)}')"
            >
              <span class="material-symbols-rounded">delete</span>
            </button>
          </div>
        </td>
      </tr>
    `).join('')
    : emptyTableRow(6, 'Belum ada pengguna.');

  renderManagementPage({
    title: 'Manajemen Pengguna',
    description: 'Kelola akun administrator, pengurus, dan portal anggota.',
    type: 'users',
    addLabel: 'Tambah Pengguna',
    canAdd: true,
    toolbarExtra: '<button class="btn btn-light" type="button" onclick="openRolePermissionSettings()"><span class="material-symbols-rounded">admin_panel_settings</span> Atur Hak Akses Pengurus</button>',
    table: `
      <table class="data-table">
        <thead>
          <tr>
            <th>Nama</th>
            <th>Username</th>
            <th>Role</th>
            <th>Anggota Tertaut</th>
            <th>Status</th>
            <th>Aksi</th>
          </tr>
        </thead>

        <tbody>${rows}</tbody>
      </table>
    `
  });

  warmUserMemberOptions();
}


/* =====================================================
   MANAGEMENT PAGE
===================================================== */

function renderManagementPage(config) {
  document.getElementById('content').innerHTML = `
    <div class="page-heading management-heading">
      <div>
        <h2>${config.title}</h2>
        <p>${config.description}</p>
      </div>

      <div class="page-actions">
        ${canPermission(config.type, 'create') ? `
          <button
            class="btn btn-primary"
            onclick="openForm('${config.type}')"
          >
            <span class="material-symbols-rounded">add</span>
            ${config.addLabel}
          </button>
        ` : ''}
      </div>
    </div>


    ${config.summary || ''}
    <div class="panel">

      <div class="table-toolbar">
        <div class="search-box">
          <span class="material-symbols-rounded">search</span>

          <input
            placeholder="Cari data..."
            value="${escapeHtml(state.search)}"
            oninput="setSearch(this.value)"
          >
        </div>

        <div class="table-toolbar-right">
          ${config.toolbarExtra || ''}
          <span class="table-version">
            Version ${escapeHtml(state.data.version || UI_VERSION)}
          </span>
        </div>
      </div>

      <div class="table-wrap responsive-table-wrap">
        ${config.table}
      </div>

      ${config.footer || ''}

    </div>
  `;
}


/* =====================================================
   SEARCH
===================================================== */

function setSearch(value) {
  state.search = String(value || '').toLowerCase();
  const pagination = state.pagination && state.pagination[state.view];
  if (pagination) pagination.page = 1;
  clearTimeout(searchTimer);

  searchTimer = setTimeout(() => {
    renderView();

    requestAnimationFrame(() => {
      const input = document.querySelector('.search-box input');
      if (!input) return;

      input.focus();
      const end = input.value.length;
      if (typeof input.setSelectionRange === 'function') {
        input.setSelectionRange(end, end);
      }
    });
  }, 180);
}

function filterData(data, fields) {
  if (!state.search) {
    return [...data];
  }

  return data.filter(item =>
    fields.some(field =>
      String(item[field] || '')
        .toLowerCase()
        .includes(state.search)
    )
  );
}


/* =====================================================
   SORTING / PAGINATION
===================================================== */

function sortAnggotaForList(data) {
  const statusPriority = status => {
    const key = String(status || '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '');

    if (key === 'calonanggota') return 0;
    if (key === 'aktif') return 1;
    if (key === 'alumni') return 2;
    if (key === 'nonaktif') return 3;
    return 2;
  };

  return [...(data || [])].sort((a, b) => {
    const byStatus = statusPriority(a && a.Status) - statusPriority(b && b.Status);
    if (byStatus !== 0) return byStatus;

    // Untuk daftar anggota, "terbaru" mengutamakan tanggal bergabung.
    // Jika tanggal bergabung sama/kosong, gunakan waktu input sebagai pembeda.
    const joinA = String(a && a.TanggalGabung || '');
    const joinB = String(b && b.TanggalGabung || '');
    const byJoin = joinB.localeCompare(joinA);
    if (byJoin !== 0) return byJoin;

    const createdA = String(a && a.DibuatPada || '');
    const createdB = String(b && b.DibuatPada || '');
    const byCreated = createdB.localeCompare(createdA);
    if (byCreated !== 0) return byCreated;

    const updatedA = String(a && a.DiubahPada || '');
    const updatedB = String(b && b.DiubahPada || '');
    const byUpdated = updatedB.localeCompare(updatedA);
    if (byUpdated !== 0) return byUpdated;

    return String(a && a.Nama || '').localeCompare(
      String(b && b.Nama || ''),
      'id',
      { sensitivity: 'base' }
    );
  });
}

function sortLatestFirst(data) {
  return [...(data || [])].sort((a, b) => {
    const dateA = String(a && a.Tanggal || '');
    const dateB = String(b && b.Tanggal || '');
    const byDate = dateB.localeCompare(dateA);
    if (byDate !== 0) return byDate;

    const createdA = String(a && a.DibuatPada || '');
    const createdB = String(b && b.DibuatPada || '');
    const byCreated = createdB.localeCompare(createdA);
    if (byCreated !== 0) return byCreated;

    const updatedA = String(a && a.DiubahPada || '');
    const updatedB = String(b && b.DiubahPada || '');
    const byUpdated = updatedB.localeCompare(updatedA);
    if (byUpdated !== 0) return byUpdated;

    return String(b && b.ID || '').localeCompare(String(a && a.ID || ''));
  });
}

function getPaginationState(view) {
  if (!state.pagination) state.pagination = {};
  if (!state.pagination[view]) {
    state.pagination[view] = { page: 1, pageSize: 10 };
  }
  return state.pagination[view];
}

function paginateData(view, data) {
  const pageState = getPaginationState(view);
  const pageSize = Math.max(1, Number(pageState.pageSize) || 10);
  const totalItems = (data || []).length;
  const totalPages = Math.max(1, Math.ceil(totalItems / pageSize));
  const page = Math.min(Math.max(1, Number(pageState.page) || 1), totalPages);
  pageState.page = page;
  pageState.pageSize = pageSize;

  const startIndex = (page - 1) * pageSize;
  const items = (data || []).slice(startIndex, startIndex + pageSize);

  return {
    page,
    pageSize,
    totalItems,
    totalPages,
    startIndex,
    endIndex: totalItems ? Math.min(startIndex + items.length, totalItems) : 0,
    items
  };
}

function setPage(view, page) {
  const pageState = getPaginationState(view);
  pageState.page = Math.max(1, Number(page) || 1);
  renderView();
  scrollTableIntoView();
}

function setPageSize(view, value) {
  const size = Number(value);
  if (![10, 25, 50].includes(size)) return;
  const pageState = getPaginationState(view);
  pageState.pageSize = size;
  pageState.page = 1;
  renderView();
}

function renderPagination(view, pager) {
  const pageSizes = [10, 25, 50];
  const from = pager.totalItems ? pager.startIndex + 1 : 0;
  const to = pager.endIndex;

  return `
    <div class="pagination-bar">
      <div class="pagination-summary">
        Menampilkan <strong>${from}-${to}</strong> dari <strong>${pager.totalItems}</strong> data
      </div>

      <div class="pagination-controls">
        <label class="page-size-control">
          <span>Per halaman</span>
          <select onchange="setPageSize('${escapeJs(view)}', this.value)">
            ${pageSizes.map(size => `
              <option value="${size}" ${size === pager.pageSize ? 'selected' : ''}>${size}</option>
            `).join('')}
          </select>
        </label>

        <button
          class="pagination-btn"
          type="button"
          onclick="setPage('${escapeJs(view)}', ${pager.page - 1})"
          ${pager.page <= 1 ? 'disabled' : ''}
          aria-label="Halaman sebelumnya"
        >
          <span class="material-symbols-rounded">chevron_left</span>
        </button>

        <span class="pagination-page">Halaman ${pager.page} / ${pager.totalPages}</span>

        <button
          class="pagination-btn"
          type="button"
          onclick="setPage('${escapeJs(view)}', ${pager.page + 1})"
          ${pager.page >= pager.totalPages ? 'disabled' : ''}
          aria-label="Halaman berikutnya"
        >
          <span class="material-symbols-rounded">chevron_right</span>
        </button>
      </div>
    </div>
  `;
}

function setAbsensiPaginationMode(mode) {
  if (!['rows', 'meeting'].includes(mode)) return;
  const pageState = getPaginationState('absensi');
  pageState.mode = mode;
  pageState.page = 1;
  renderAbsensi();
}

function renderAbsensiViewControl() {
  const mode = getPaginationState('absensi').mode || 'rows';
  return `
    <label class="absensi-view-control">
      <span class="material-symbols-rounded">view_agenda</span>
      <span>Tampilan</span>
      <select onchange="setAbsensiPaginationMode(this.value)">
        <option value="rows" ${mode === 'rows' ? 'selected' : ''}>Baris</option>
        <option value="meeting" ${mode === 'meeting' ? 'selected' : ''}>Per Pertemuan</option>
      </select>
    </label>
  `;
}

function groupAbsensiByMeeting(data) {
  const groups = new Map();

  (data || []).forEach(item => {
    const activityId = String(item.KegiatanID || '').trim();
    const fallback = `${String(item.Tanggal || '')}|${String(item.KegiatanNama || '')}`;
    const key = activityId || fallback;

    if (!groups.has(key)) {
      groups.set(key, {
        key,
        id: activityId,
        date: String(item.Tanggal || ''),
        name: String(item.KegiatanNama || 'Kegiatan'),
        latestCreated: String(item.DibuatPada || ''),
        items: []
      });
    }

    const group = groups.get(key);
    group.items.push(item);
    if (String(item.DibuatPada || '') > group.latestCreated) {
      group.latestCreated = String(item.DibuatPada || '');
    }
  });

  return Array.from(groups.values())
    .map(group => {
      group.items = sortLatestFirst(group.items);
      return group;
    })
    .sort((a, b) => {
      const byDate = String(b.date || '').localeCompare(String(a.date || ''));
      if (byDate !== 0) return byDate;
      return String(b.latestCreated || '').localeCompare(String(a.latestCreated || ''));
    });
}

function paginateMeetingGroups(groups) {
  const pageState = getPaginationState('absensi');
  const totalItems = groups.length;
  const totalPages = Math.max(1, totalItems);
  const page = Math.min(Math.max(1, Number(pageState.page) || 1), totalPages);
  pageState.page = page;

  return {
    page,
    totalItems,
    totalPages,
    currentGroup: totalItems ? groups[page - 1] : null
  };
}

function renderMeetingPagination(pager) {
  const group = pager.currentGroup;
  return `
    <div class="pagination-bar meeting-pagination">
      <div class="pagination-summary">
        ${group
          ? `Pertemuan <strong>${pager.page}</strong> dari <strong>${pager.totalItems}</strong>`
          : 'Belum ada pertemuan'}
      </div>

      <div class="pagination-controls">
        <span class="page-size-static">1 pertemuan / halaman</span>

        <button
          class="pagination-btn"
          type="button"
          onclick="setPage('absensi', ${pager.page - 1})"
          ${pager.page <= 1 ? 'disabled' : ''}
          aria-label="Pertemuan sebelumnya"
        >
          <span class="material-symbols-rounded">chevron_left</span>
        </button>

        <span class="pagination-page">${pager.page} / ${pager.totalPages}</span>

        <button
          class="pagination-btn"
          type="button"
          onclick="setPage('absensi', ${pager.page + 1})"
          ${pager.page >= pager.totalPages ? 'disabled' : ''}
          aria-label="Pertemuan berikutnya"
        >
          <span class="material-symbols-rounded">chevron_right</span>
        </button>
      </div>
    </div>
  `;
}

function scrollTableIntoView() {
  requestAnimationFrame(() => {
    const panel = document.querySelector('.content .panel');
    if (!panel) return;

    const scroller = getPrimaryScrollContainer();
    if (scroller && scroller !== document.scrollingElement && scroller !== document.documentElement && scroller !== document.body) {
      const scrollerRect = scroller.getBoundingClientRect();
      const panelRect = panel.getBoundingClientRect();
      const top = scroller.scrollTop + panelRect.top - scrollerRect.top - 72;
      scroller.scrollTo({ top: Math.max(0, top), behavior: 'smooth' });
      return;
    }

    const top = panel.getBoundingClientRect().top + window.scrollY - 90;
    window.scrollTo({ top: Math.max(0, top), behavior: 'smooth' });
  });
}

/* =====================================================
   FORMS
===================================================== */

async function openForm(type, id = '') {
  if (!canPermission(type, id ? 'edit' : 'create')) {
    showToast('Anda tidak memiliki izin untuk tindakan ini.', 'error');
    return;
  }
  modalLoadGeneration++;
  const requiredModules = getFormModules(type);

  if (!areModulesReady(requiredModules)) {
    const token = state.token;
    const generation = dataLoadGeneration;
    openCustomModal('Memuat formulir', 'Menyiapkan data pilihan...', loadingHtml(180));
    const modalGeneration = modalLoadGeneration;
    try {
      if (!await ensureModulesLoaded(requiredModules)) return;
      if (modalGeneration !== modalLoadGeneration || !isCurrentDataLoad(token, generation)) return;
    } catch (error) {
      if (modalGeneration !== modalLoadGeneration || !isCurrentDataLoad(token, generation)) return;
      closeModal();
      handleDataLoadError(error);
      return;
    }
  }

if (type === 'users' && !state.userMemberOptionsLoaded) {
  const token = state.token;
  const generation = dataLoadGeneration;
  const modalGeneration = modalLoadGeneration;
  if (!document.getElementById('modalBackdrop').classList.contains('show')) {
    openCustomModal('Memuat formulir', 'Menyiapkan daftar anggota...', loadingHtml(150));
  } else {
    document.getElementById('modalSubtitle').textContent = 'Menyiapkan daftar anggota...';
    document.getElementById('modalBody').innerHTML = loadingHtml(150);
  }
  try {
    await ensureUserMemberOptionsLoaded();
    if (modalGeneration !== modalLoadGeneration || !isCurrentDataLoad(token, generation)) return;
  } catch (error) {
    if (modalGeneration !== modalLoadGeneration || !isCurrentDataLoad(token, generation)) return;
    closeModal();
    showToast(error.message || 'Gagal memuat daftar anggota.', 'error');
    return;
  }
}

  if (type === 'inventaris') {
    const generation = modalLoadGeneration;
    const token = state.token;
    try {
      state.inventoryQuickOptions = await serverCall('getInventoryQuickOptions', token);
      if (generation !== modalLoadGeneration || token !== state.token) return;
    } catch (error) { showToast(error.message, 'error'); return; }
  }
  const item = id ? getItem(type, id) : {};
  const definition = getFormDefinition(type, item);

  document.getElementById('modal').classList.remove('small');
  document.getElementById('modalTitle').textContent =
    id ? 'Edit Data' : definition.title;

  document.getElementById('modalSubtitle').textContent =
    definition.subtitle;

  document.getElementById('modalBody').innerHTML = `
    <form
      class="form-grid"
      id="dataForm"
      onsubmit="submitForm(event,'${escapeJs(type)}','${escapeJs(id)}')"
    >

      ${definition.notice || ''}

      ${definition.fields
        .map(field => createField(field, item))
        .join('')
      }

      <div class="form-actions">
        <button
          type="button"
          class="btn btn-light"
          onclick="closeModal()"
        >
          Batal
        </button>

        <button
          type="submit"
          class="btn btn-primary"
          id="saveButton"
        >
          <span class="material-symbols-rounded">save</span>
          Simpan Data
        </button>
      </div>

    </form>
  `;

  const backdrop = document.getElementById('modalBackdrop');
  backdrop.classList.add('show');
  backdrop.setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');

  if (type === 'surat') syncLetterForm();
  if (type === 'anggota') {
    const statusSelect = document.querySelector('#dataForm [name="Status"]');
    syncAnggotaFormStatus(statusSelect ? statusSelect.value : 'Calon Anggota');
  }
}

function getFormDefinition(type, item = {}) {
  if (type === 'anggota') {
    const memberStatusField = selectField(
      'Status',
      'Status Keanggotaan',
      ['Calon Anggota','Aktif','Nonaktif','Alumni'],
      true
    );
    memberStatusField.defaultValue = 'Calon Anggota';
    memberStatusField.onchange = 'syncAnggotaFormStatus(this.value)';

    const inaugurationField = field(
      'TanggalPelantikan',
      'Tanggal Pelantikan / Pengesahan',
      'date'
    );
    inaugurationField.wrapperClass = 'member-inauguration-field';
    inaugurationField.help = 'Wajib diisi ketika anggota sudah terlantik dan status diubah menjadi Aktif.';

    return {
      title: 'Tambah Anggota',
      subtitle: 'Masukkan identitas anggota. NTA dan tanggal pendaftaran dibuat otomatis oleh sistem.',
      notice: `
        <div class="member-nta-auto-note">
          <span class="material-symbols-rounded">badge</span>
          <div>
            <strong>NTA dibuat otomatis</strong>
            <span>Calon anggota: MULCA00142026. Setelah terlantik: MUL00192026. Nomor urut 001 tetap dipertahankan.</span>
          </div>
        </div>
      `,
      fields: [
        field('Nama','Nama Lengkap','text',true),

        selectField(
          'JenisKelamin',
          'Jenis Kelamin',
          ['Laki-laki','Perempuan']
        ),

        field('TempatLahir','Tempat Lahir'),
        field('TanggalLahir','Tanggal Lahir','date'),
        field('Telepon','Nomor Telepon'),
        field('SekolahInstansi','Sekolah / Instansi'),

        selectField(
          'Krida',
          'Krida',
          [
            'Krida Olahraga Dirgantara',
            'Krida Pengetahuan Dirgantara',
            'Krida Jasa Kedirgantaraan'
          ]
        ),

        field('Jabatan','Jabatan'),
        memberStatusField,
        inaugurationField,
        textareaField('Alamat','Alamat',true)
      ]
    };
  }

  if (type === 'kegiatan') {
    const officerSeen = new Set();
    const officerOptions = (state.data.pengurus || [])
      .filter(entry => String(entry.Status || '').trim().toLowerCase() === 'aktif')
      .sort((a, b) => Number(a.Urutan || 999) - Number(b.Urutan || 999))
      .map(entry => {
        const name = String(entry.Nama || '').trim();
        if (!name || officerSeen.has(name.toLowerCase())) return null;
        officerSeen.add(name.toLowerCase());
        const meta = [entry.Jabatan, entry.Bidang].filter(Boolean).join(' • ');
        return { value: name, label: meta || 'Pengurus aktif' };
      })
      .filter(Boolean);

    const currentStatus = String(item.Status || 'Rencana').trim() || 'Rencana';
    const stateMessage = item && item.ID
      ? `Status saat ini: <strong>${escapeHtml(currentStatus)}</strong>. Status hanya berubah melalui tombol aksi pada daftar Kegiatan.`
      : 'Kegiatan baru otomatis dibuat dengan status <strong>Rencana</strong>.';

    return {
      title: 'Tambah Kegiatan',
      subtitle: 'Tambahkan atau perbarui metadata kegiatan. Status dikendalikan oleh state machine.',
      notice: `
        <div class="kegiatan-state-note">
          <span class="material-symbols-rounded">conversion_path</span>
          <div>
            <strong>Alur status terkunci</strong>
            <span>${stateMessage} Alur: Rencana → Berjalan → Selesai, atau Dibatalkan.</span>
          </div>
        </div>
      `,
      fields: [
        field('NamaKegiatan','Nama Kegiatan','text',true),
        field('Tanggal','Tanggal','date',true),

        selectField(
          'Jenis',
          'Jenis Kegiatan',
          [
            'Latihan',
            'Rapat',
            'Pendidikan',
            'Bakti Sosial',
            'Kunjungan',
            'Upacara',
            'Lainnya'
          ]
        ),

        field('Lokasi','Lokasi'),
        datalistField(
          'PenanggungJawab',
          'Penanggung Jawab',
          officerOptions,
          false,
          'Pilih nama pengurus aktif dari daftar, atau ketik nama instruktur/pihak luar secara manual.'
        ),
        textareaField('Keterangan','Keterangan',true)
      ]
    };
  }

  if (type === 'absensi') {
    // Gunakan filter yang sama dengan Absensi Massal agar aturan anggota/kegiatan
    // tidak memiliki dua implementasi berbeda.
    const activityOptions = getRunningKegiatanOptions().map(item => ({
      value: item.ID,
      label: `${item.NamaKegiatan} - ${formatDate(item.Tanggal)}`
    }));

    const memberOptions = getBatchAttendanceMembers().map(item => ({
      value: item.ID,
      label: `${item.Nama}${String(item.Status || '').trim().toLowerCase() === 'calon anggota' ? ' (Calon)' : ''}`
    }));

    return {
      title: item && item.ID ? 'Edit Absensi' : 'Input Absensi Satuan / Libur',
      subtitle: item && item.ID
        ? 'Gunakan form ini untuk koreksi satu catatan absensi.'
        : 'Gunakan untuk satu anggota atau mencatat Libur. Untuk absensi rutin gunakan Absensi Massal.',
      fields: [
        customSelectField(
          'KegiatanID',
          'Kegiatan',
          activityOptions,
          true
        ),

        customSelectField(
          'AnggotaID',
          'Anggota (kosongkan jika Libur)',
          memberOptions,
          false
        ),

        customSelectField(
          'StatusKehadiran',
          'Status Kehadiran',
          [
            { value: 'Hadir', label: 'Hadir' },
            { value: 'Izin', label: 'Izin' },
            { value: 'Sakit', label: 'Sakit' },
            { value: 'Alpa', label: 'Alpa' },
            { value: 'Libur', label: 'Libur — Tidak ada latihan' }
          ],
          true
        ),

        textareaField('Catatan','Catatan / Keterangan',true)
      ]
    };
  }

  if (type === 'kas') {
    const activityOptions = [...(state.data.kegiatan || [])]
      .sort((a, b) => String(b.Tanggal || '').localeCompare(String(a.Tanggal || '')))
      .map(item => ({
        value: item.ID,
        label: `${item.NamaKegiatan} - ${formatDate(item.Tanggal)}`
      }));

    return {
      title: 'Tambah Transaksi',
      subtitle: 'Catat transaksi kas organisasi. Hubungkan transaksi ke kegiatan agar otomatis masuk laporan PDF kegiatan.',
      fields: [
        field('Tanggal','Tanggal','date',true),

        selectField(
          'Jenis',
          'Jenis Transaksi',
          ['Pemasukan','Pengeluaran'],
          true
        ),

        field('Kategori','Kategori'),
        field('Nominal','Nominal','number',true),
        field('NoBukti','Nomor Bukti'),
        customSelectField('KegiatanID','Terkait Kegiatan (opsional)',activityOptions,false),
        field('Petugas','Petugas'),
        textareaField('Keterangan','Keterangan',true)
      ]
    };
  }

  if (type === 'inventaris') {
    return {
      title: 'Tambah Inventaris',
      subtitle: 'Catat barang dan perlengkapan organisasi.',
      fields: [
        {...field('KodeBarang','Kode Barang'), help:'Kosongkan untuk kode otomatis INV-SADIRGA/KATEGORI/0001.'},
        field('NamaBarang','Nama Barang','text',true),
        selectField('Kategori','Kategori', INVENTORY_CATEGORIES, true),
        field('Jumlah','Jumlah','number',true),

        selectField(
          'Kondisi',
          'Kondisi',
          ['Baik','Rusak Ringan','Rusak Berat','Hilang']
        ),

        field('Lokasi','Lokasi'),
        customSelectField('PenanggungJawab','Penanggung Jawab', (state.inventoryQuickOptions?.pengurus || []).map(row => ({value:row.Nama,label:row.Nama + ' — ' + row.Jabatan})), true),
        textareaField('Keterangan','Keterangan',true)
      ]
    };
  }

  if (type === 'surat') {
    return {
      title: 'Tambah Surat',
      subtitle: 'Arsipkan surat masuk atau surat keluar.',
      fields: [
        {...selectField('Penomoran','Penomoran surat keluar',['Otomatis','Manual'],true), defaultValue:item.ID ? 'Manual' : 'Otomatis', onchange:'syncLetterForm()'},
        {...selectField('FormatSurat','Format surat keluar',['PANPEL-SADIRGA','SADIRGA-SDA','SK/MUSAKA'],true), defaultValue:letterFormatFromNumber(item.NomorSurat), onchange:'syncLetterForm()', help:'PANPEL: panitia kegiatan. SADIRGA-SDA: surat resmi instansi. SK/MUSAKA: keputusan musyawarah.'},
        {...field('NomorSurat','Nomor Surat'), help:'Surat masuk mengikuti nomor pengirim. Nomor otomatis diterbitkan saat disimpan; urutan terpisah per format dan bulan.'},
        {...field('Tanggal','Tanggal','date',true), onchange:'syncLetterForm()'},

        {...selectField(
          'Jenis',
          'Jenis Surat',
          ['Surat Masuk','Surat Keluar'],
          true
        ), onchange:'syncLetterForm()'},

        field('Perihal','Perihal','text',true),
        field('AsalTujuan','Asal / Tujuan'),

        selectField(
          'Status',
          'Status',
          ['Draft','Diproses','Selesai','Diarsipkan']
        ),

        field('LinkFile','Link File / Google Drive','url'),
        textareaField('Keterangan','Keterangan',true)
      ]
    };
  }

  if (type === 'pengurus') {
    const memberOptions = state.data.anggota.map(item => ({
      value: item.ID,
      label: item.Nama
    }));

    return {
      title: 'Tambah Pengurus',
      subtitle: 'Susun struktur kepengurusan organisasi.',
      fields: [
        customSelectField(
          'AnggotaID',
          'Anggota',
          memberOptions,
          true
        ),

        field('Jabatan','Jabatan','text',true),
        field('Bidang','Bidang / Divisi'),
        field('Periode','Periode'),
        field('Urutan','Urutan Struktur','number'),

        selectField(
          'Status',
          'Status',
          ['Aktif','Nonaktif']
        )
      ]
    };
  }

  if (type === 'users') {
    return {
      title: item && item.ID ? 'Edit Pengguna' : 'Tambah Pengguna',
      subtitle: 'Buat akun administrator, pengurus, atau portal anggota. Untuk role ANGGOTA, tautkan akun ke satu data anggota.',
      fields: [
        {...field('Nama','Nama Pengguna'), help:'Untuk role ANGGOTA, nama otomatis mengikuti data anggota yang ditautkan.'},
        field('Username','Username','text',true),

        selectField(
          'Role',
          'Role',
          ['ADMIN','PENGURUS','ANGGOTA'],
          true
        ),

        customSelectField(
          'AnggotaID',
          'Tautkan Anggota (wajib untuk ANGGOTA)',
          ((state.userMemberOptionsLoaded ? state.userMemberOptions : state.data.anggota) || []).map(member => ({
            value: member.ID,
            label: `${member.Nama} — ${member.NTA || 'NTA belum tersedia'}`
          })),
          false
        ),

        selectField(
          'Status',
          'Status',
          ['Aktif','Nonaktif']
        ),

        field(
          'Password',
          'Password Baru',
          'password'
        )
      ]
    };
  }

  throw new Error('Jenis form tidak dikenal: ' + String(type));
}

function field(name, label, type = 'text', required = false) {
  return {
    name,
    label,
    type,
    required,
    full: false
  };
}

function textareaField(name, label, full = false) {
  return {
    name,
    label,
    type: 'textarea',
    full
  };
}

function selectField(
  name,
  label,
  options,
  required = false
) {
  return {
    name,
    label,
    type: 'select',
    options: options.map(value => ({
      value,
      label: value
    })),
    required
  };
}

function customSelectField(
  name,
  label,
  options,
  required = false
) {
  return {
    name,
    label,
    type: 'select',
    options,
    required
  };
}

function datalistField(
  name,
  label,
  options,
  required = false,
  help = ''
) {
  return {
    name,
    label,
    type: 'datalist',
    options: Array.isArray(options) ? options : [],
    required,
    help
  };
}

function createField(field, item) {
  const itemHasValue = item && Object.prototype.hasOwnProperty.call(item, field.name);
  const value = itemHasValue ? (item[field.name] ?? '') : (field.defaultValue ?? '');
  const fullClass = field.full ? ' full' : '';
  const wrapperClass = field.wrapperClass ? ` ${field.wrapperClass}` : '';
  const helpHtml = field.help ? `<small class="form-help">${escapeHtml(field.help)}</small>` : '';
  const onchangeAttr = field.onchange ? ` onchange="${field.onchange}"` : '';

  if (field.type === 'datalist') {
    const listId = 'list-' + String(field.name || 'field').replace(/[^a-zA-Z0-9_-]/g, '-');
    return `
      <div class="form-group${fullClass}${wrapperClass}">
        <label>${field.label}</label>
        <input
          class="form-control"
          type="text"
          name="${field.name}"
          value="${escapeHtml(value)}"
          list="${escapeHtml(listId)}"
          autocomplete="off"
          ${field.required ? 'required' : ''}
          ${onchangeAttr}
        >
        <datalist id="${escapeHtml(listId)}">
          ${(field.options || []).map(option => `
            <option value="${escapeHtml(option.value)}" label="${escapeHtml(option.label || '')}"></option>
          `).join('')}
        </datalist>
        ${helpHtml}
      </div>
    `;
  }

  if (field.type === 'textarea') {
    return `
      <div class="form-group${fullClass}${wrapperClass}">
        <label>${field.label}</label>

        <textarea
          class="form-control"
          name="${field.name}"
        >${escapeHtml(value)}</textarea>
        ${helpHtml}
      </div>
    `;
  }

  if (field.type === 'select') {
    const hasCurrentOption = field.options.some(option =>
      String(option.value) === String(value)
    );
    const customCurrentOption = value && !hasCurrentOption
      ? `<option value="${escapeHtml(value)}" selected>${escapeHtml(value)} (data tersimpan)</option>`
      : '';

    return `
      <div class="form-group${fullClass}${wrapperClass}">
        <label>${field.label}</label>

        <select
          class="form-control"
          name="${field.name}"
          ${field.required ? 'required' : ''}
          ${onchangeAttr}
        >
          <option value="">Pilih...</option>
          ${customCurrentOption}

          ${field.options.map(option => `
            <option
              value="${escapeHtml(option.value)}"
              ${String(option.value) === String(value) ? 'selected' : ''}
            >
              ${escapeHtml(option.label)}
            </option>
          `).join('')}
        </select>
        ${helpHtml}
      </div>
    `;
  }

  return `
    <div class="form-group${fullClass}${wrapperClass}">
      <label>${field.label}</label>

      <input
        class="form-control"
        type="${field.type}"
        name="${field.name}"
        value="${escapeHtml(value)}"
        ${field.required ? 'required' : ''}
        ${field.type === 'password' ? 'autocomplete="new-password" minlength="8"' : ''}
        ${field.type === 'number' ? 'min="0" step="1"' : ''}
        ${onchangeAttr}
      >
      ${helpHtml}
    </div>
  `;
}


function syncAnggotaFormStatus(statusValue) {
  const wrapper = document.querySelector('#dataForm .member-inauguration-field');
  if (!wrapper) return;

  const input = wrapper.querySelector('[name="TanggalPelantikan"]');
  if (!input) return;

  const status = String(statusValue || '').trim();
  const hasStoredDate = Boolean(String(input.value || '').trim());
  const shouldShow = status === 'Aktif' || status === 'Alumni' || hasStoredDate;

  wrapper.hidden = !shouldShow;
  input.required = status === 'Aktif';
}


/* =====================================================
   SUBMIT / DELETE
===================================================== */

async function submitForm(event, type, id) {
  event.preventDefault();

  const form = event.target;
  const button = document.getElementById('saveButton');

  const data = Object.fromEntries(
    new FormData(form).entries()
  );

  if (type === 'absensi' && data.StatusKehadiran === 'Libur' && !String(data.Catatan || '').trim()) {
    data.Catatan = 'Tidak ada latihan';
  }

  if (type === 'anggota' && data.Status === 'Aktif' && !String(data.TanggalPelantikan || '').trim()) {
    showToast('Tanggal Pelantikan wajib diisi untuk anggota yang sudah terlantik.', 'warning');
    const inaugurationInput = form.querySelector('[name="TanggalPelantikan"]');
    if (inaugurationInput) inaugurationInput.focus();
    return;
  }

  if (id) {
    data.ID = id;
  }

  const methods = {
    anggota: 'saveAnggota',
    kegiatan: 'saveKegiatan',
    absensi: 'saveAbsensi',
    kas: 'saveKas',
    inventaris: 'saveInventaris',
    surat: 'saveSurat',
    pengurus: 'savePengurus',
    users: 'saveUser'
  };

  const method = methods[type];
  if (!method) {
    showToast('Jenis data tidak dikenali.', 'error');
    return;
  }

  button.disabled = true;
  button.textContent = 'Menyimpan...';

  try {
    const result = await serverCall(
      method,
      state.token,
      data
    );

    closeModal();

    if (type === 'kegiatan' && id) {
      const synced = Number(result && result.SyncAbsensi || 0);
      showToast(
        synced
          ? `Kegiatan berhasil diperbarui. ${synced} data absensi terkait ikut disinkronkan.`
          : 'Kegiatan berhasil diperbarui. Tidak ada data absensi terkait yang perlu diubah.'
      );
    } else if (type === 'anggota' && result && result.NTA) {
      showToast(`Data anggota berhasil disimpan. NTA: ${result.NTA}`, 'success');
    } else {
      showToast('Data berhasil disimpan.');
    }

    finishLocalMutation(type, result, id);

  } catch (error) {
    showToast(error.message, 'error');

  } finally {
    if (button) {
      button.disabled = false;
      button.innerHTML = `
        <span class="material-symbols-rounded">save</span>
        Simpan Data
      `;
    }
  }
}

async function deleteItem(type, id) {
  if (!canPermission(type, 'delete')) {
    showToast('Anda tidak memiliki izin untuk tindakan ini.', 'error');
    return;
  }
  if (!await customConfirm('Hapus data ini? Tindakan ini tidak dapat dibatalkan.', 'Hapus Data', 'Hapus')) {
    return;
  }

  const methods = {
    anggota: 'deleteAnggota',
    kegiatan: 'deleteKegiatan',
    absensi: 'deleteAbsensi',
    kas: 'deleteKas',
    inventaris: 'deleteInventaris',
    surat: 'deleteSurat',
    pengurus: 'deletePengurus',
    users: 'deleteUser'
  };

  const method = methods[type];
  if (!method) {
    showToast('Jenis data tidak dikenali.', 'error');
    return;
  }

  try {
    await serverCall(
      method,
      state.token,
      id
    );

    showToast('Data berhasil dihapus.');
    finishLocalMutation(type, null, id, true);

  } catch (error) {
    showToast(error.message, 'error');
  }
}


/* =====================================================
   PUBLIC IZIN ANGGOTA
===================================================== */

async function renderPublicIzin(code) {
  document.documentElement.classList.add('public-izin-page');
  repairScrollLockState();
  document.getElementById('appRoot').innerHTML = `
    <div class="auth-shell izin-public-shell">
      <section class="auth-visual">
        <div class="visual-brand">
          <div class="brand-mark"><span class="material-symbols-rounded">approval</span></div>
          <div class="brand-copy"><strong>IZIN ANGGOTA</strong><small>SAKA DIRGANTARA</small></div>
        </div>
        <div class="visual-copy">
          <span class="eyebrow">Permission Request</span>
          <h1>Ajukan izin kegiatan <span>secara tertib.</span></h1>
          <p>Masukkan NTA, pilih jenis pengajuan, jelaskan alasan, dan lampirkan bukti pendukung. Pengajuan harus diverifikasi Pengurus sebelum digunakan pada Absensi.</p>
        </div>
      </section>

      <section class="auth-panel">
        <div class="auth-card izin-public-card">
          <div id="publicIzinContent">${loadingHtml(200)}</div>
        </div>
      </section>
    </div>`;

  try {
    const kegiatan = await serverCall('getPublicIzinData', code);
    const formHtml = kegiatan.canSubmit ? `
      <form id="publicIzinForm" onsubmit="submitPublicIzin(event,'${escapeJs(code)}')">
        <div class="auth-field">
          <label>NOMOR TANDA ANGGOTA (NTA)</label>
          <input class="auth-input" name="nta" required autocomplete="off" placeholder="Masukkan NTA Anda">
        </div>
        <div class="auth-field">
          <label>JENIS PENGAJUAN</label>
          <select class="auth-input" name="jenisPengajuan" id="publicIzinType" required onchange="syncPublicIzinEvidenceOptions(this.value)">
            <option value="Izin">Izin</option>
            <option value="Sakit">Sakit</option>
          </select>
        </div>
        <div class="auth-field">
          <label>ALASAN</label>
          <textarea class="auth-input izin-public-textarea" name="alasan" required maxlength="500" placeholder="Tuliskan alasan secara singkat dan jelas"></textarea>
        </div>
        <div class="auth-field">
          <label>JENIS BUKTI</label>
          <select class="auth-input" name="jenisBukti" id="publicIzinEvidenceType" required>
            <option value="Surat Izin">Surat Izin</option>
          </select>
          <small class="izin-file-help" id="publicIzinEvidenceHelp">Untuk izin, lampirkan surat izin. Untuk sakit, bukti dapat berupa surat dokter atau bukti obat.</small>
        </div>
        <div class="auth-field">
          <label>BUKTI PENDUKUNG</label>
          <input class="auth-input izin-file-input" id="publicIzinFile" name="bukti" type="file" required accept="application/pdf,image/jpeg,image/png,image/webp">
          <small class="izin-file-help">PDF/JPG/PNG/WEBP. Maksimum 3 MB setelah diproses.</small>
        </div>
        <div id="publicIzinError" class="public-izin-error" role="alert" hidden></div>
        <button class="btn btn-primary auth-submit" type="submit" id="publicIzinButton">
          <span class="material-symbols-rounded">send</span> Kirim Pengajuan Izin
        </button>
      </form>` : `
      <div class="izin-public-closed">
        <span class="material-symbols-rounded">lock_clock</span>
        <strong>${escapeHtml(kegiatan.state || 'Tidak tersedia')}</strong>
        <p>${escapeHtml(kegiatan.message || '')}</p>
      </div>`;

    document.getElementById('publicIzinContent').innerHTML = `
      <div class="mobile-logo visual-brand">
        <div class="brand-mark"><span class="material-symbols-rounded">approval</span></div>
        <div class="brand-copy" style="color:#071827;"><strong>IZIN ANGGOTA</strong><small>SAKA DIRGANTARA</small></div>
      </div>
      <h2>${escapeHtml(kegiatan.NamaKegiatan)}</h2>
      <p>${formatDate(kegiatan.Tanggal)} • ${escapeHtml(kegiatan.Lokasi || 'Lokasi belum ditentukan')}</p>
      <div class="izin-public-window">
        <span><strong>Mulai:</strong> ${escapeHtml(kegiatan.openAt || '-')}</span>
        <span><strong>Batas:</strong> ${escapeHtml(kegiatan.autoCloseAt || '-')}</span>
      </div>
      ${formHtml}
      <div class="login-note">Pengajuan yang sudah dikirim dapat diperbarui selama periode izin masih aktif. Bukti lama akan diganti oleh bukti terbaru.</div>`;
  } catch (error) {
    document.getElementById('publicIzinContent').innerHTML = errorBox(error.message);
  }
}

function syncPublicIzinEvidenceOptions(type) {
  const select = document.getElementById('publicIzinEvidenceType');
  const help = document.getElementById('publicIzinEvidenceHelp');
  if (!select) return;
  const isSick = String(type || '') === 'Sakit';
  select.innerHTML = isSick
    ? '<option value="Surat Dokter">Surat Dokter</option><option value="Bukti Obat">Bukti Obat</option>'
    : '<option value="Surat Izin">Surat Izin</option>';
  if (help) {
    help.textContent = isSick
      ? 'Untuk sakit, bukti dapat berupa surat dokter atau foto/nota/kemasan obat yang relevan.'
      : 'Untuk izin, lampirkan surat izin.';
  }
}

async function submitPublicIzin(event, code) {
  event.preventDefault();
  const form = event.target;
  const button = document.getElementById('publicIzinButton');
  if (form.dataset.submitting === '1') return;
  if (!form.reportValidity()) return;
  const errorBox = document.getElementById('publicIzinError');
  if (errorBox) { errorBox.hidden = true; errorBox.textContent = ''; }
  const data = new FormData(form);
  const file = document.getElementById('publicIzinFile')?.files?.[0];

  if (!file) {
    showToast('Bukti pendukung wajib dilampirkan.', 'warning');
    return;
  }

  form.dataset.submitting = '1';
  form.setAttribute('aria-busy', 'true');
  if (button) {
    button.disabled = true;
    button.innerHTML = '<span class="material-symbols-rounded">hourglass_top</span> Memproses...';
  }

  try {
    const prepared = await prepareIzinEvidenceFile(file);
    if (button) button.textContent = 'Mengirim pengajuan...';
    const result = await serverCall('publicSubmitIzin', code, {
      nta: String(data.get('nta') || '').trim(),
      jenisPengajuan: String(data.get('jenisPengajuan') || '').trim(),
      alasan: String(data.get('alasan') || '').trim(),
      jenisBukti: String(data.get('jenisBukti') || '').trim(),
      name: prepared.name,
      mimeType: prepared.mimeType,
      base64: prepared.base64
    });

    if (!result || result.success !== true) {
      throw new Error(result?.message || 'Server belum mengonfirmasi penyimpanan pengajuan.');
    }
    document.getElementById('publicIzinContent').innerHTML = `
      <div class="izin-submit-success" role="status" tabindex="-1">
        <div class="setup-icon" style="background:linear-gradient(135deg,#14a777,#12b6ca);">
          <span class="material-symbols-rounded">task_alt</span>
        </div>
        <h2>${result.updated ? 'Pengajuan Diperbarui' : 'Pengajuan Terkirim'}</h2>
        <p>Pengajuan <strong>${escapeHtml(result.nama)}</strong> untuk <strong>${escapeHtml(result.kegiatan)}</strong> berhasil disimpan dan menunggu verifikasi Pengurus.</p>
        <div class="login-note">Pengajuan belum menjadi status Absensi. Setelah Pengurus menyetujui, sistem akan memprefill Izin atau Sakit pada Absensi Massal sesuai jenis pengajuan.</div>
      </div>`;
    const success = document.querySelector('.izin-submit-success');
    if (success) {
      success.focus({ preventScroll: true });
      success.scrollIntoView({ block: 'start', behavior: 'auto' });
    }
  } catch (error) {
    const message = error?.message || 'Pengajuan gagal dikirim. Silakan coba lagi.';
    if (errorBox) {
      errorBox.textContent = message;
      errorBox.hidden = false;
      errorBox.scrollIntoView({ block: 'nearest', behavior: 'auto' });
    }
    showToast(message, 'error');
    if (button) {
      button.disabled = false;
      button.innerHTML = '<span class="material-symbols-rounded">send</span> Kirim Pengajuan Izin';
    }
  } finally {
    delete form.dataset.submitting;
    form.removeAttribute('aria-busy');
  }
}

async function prepareIzinEvidenceFileCore(file) {
  const mime = String(file?.type || '').toLowerCase();
  if (['image/jpeg','image/png','image/webp'].includes(mime)) {
    const prepared = await prepareKegiatanDocumentationImage(file);
    return prepared;
  }
  if (mime !== 'application/pdf') {
    throw new Error('Bukti pendukung harus berupa PDF, JPG, PNG, atau WEBP.');
  }
  if (file.size > 3 * 1024 * 1024) {
    throw new Error('PDF bukti pendukung maksimum 3 MB.');
  }
  return {
    name: String(file.name || 'surat-izin.pdf').replace(/[^a-zA-Z0-9 _.-]+/g, '').slice(0, 90) || 'surat-izin.pdf',
    mimeType: 'application/pdf',
    base64: await blobToDocumentationBase64(file)
  };
}

/* =====================================================
   PUBLIC QR CHECK-IN
===================================================== */

async function renderPublicCheckin(code) {
  document.getElementById('appRoot').innerHTML = `
    <div class="auth-shell">

      <section class="auth-visual">
        <div class="visual-brand">
          <div class="brand-mark">
            <span class="material-symbols-rounded">qr_code_2</span>
          </div>

          <div class="brand-copy">
            <strong>ABSENSI</strong>
            <small>SAKA DIRGANTARA</small>
          </div>
        </div>

        <div class="visual-copy">
          <span class="eyebrow">QR Attendance</span>

          <h1>
            Check-in kegiatan
            <span>lebih cepat.</span>
          </h1>

          <p>
            Scan QR, masukkan NTA, kemudian sistem akan mencatat
            kehadiran secara otomatis.
          </p>
        </div>
      </section>


      <section class="auth-panel">
        <div class="auth-card">

          <div id="publicCheckinContent">
            ${loadingHtml(200)}
          </div>

        </div>
      </section>

    </div>
  `;

  try {
    const kegiatan = await serverCall(
      'getPublicCheckinData',
      code
    );

    document.getElementById('publicCheckinContent').innerHTML = `
      <div class="mobile-logo visual-brand">
        <div class="brand-mark">
          <span class="material-symbols-rounded">qr_code_2</span>
        </div>

        <div class="brand-copy" style="color:#071827;">
          <strong>ABSENSI</strong>
          <small>SAKA DIRGANTARA</small>
        </div>
      </div>

      <h2>${escapeHtml(kegiatan.NamaKegiatan)}</h2>

      <p>
        ${formatDate(kegiatan.Tanggal)}
        •
        ${escapeHtml(kegiatan.Lokasi || 'Lokasi belum ditentukan')}
      </p>

      <form onsubmit="submitPublicCheckin(event,'${escapeJs(code)}')">

        <div class="auth-field">
          <label>NOMOR TANDA ANGGOTA (NTA)</label>

          <input
            class="auth-input"
            name="nta"
            required
            autocomplete="off"
            placeholder="Masukkan NTA Anda"
          >
        </div>

        <button
          class="btn btn-primary auth-submit"
          type="submit"
          id="publicCheckinButton"
        >
          <span class="material-symbols-rounded">how_to_reg</span>
          Check-in Sekarang
        </button>
      </form>

      <div class="login-note">
        Kehadiran hanya dapat dicatat satu kali untuk setiap anggota
        pada kegiatan yang sama.
      </div>
    `;

  } catch (error) {
    document.getElementById('publicCheckinContent').innerHTML =
      errorBox(error.message);
  }
}

async function submitPublicCheckin(event, code) {
  event.preventDefault();

  const form = event.target;
  const nta = new FormData(form).get('nta');
  const button = document.getElementById('publicCheckinButton');

  button.disabled = true;
  button.textContent = 'Memproses...';

  try {
    const result = await serverCall(
      'publicCheckin',
      code,
      nta
    );

    document.getElementById('publicCheckinContent').innerHTML = `
      <div style="text-align:center;">
        <div
          class="setup-icon"
          style="background:linear-gradient(135deg,#14a777,#12b6ca);"
        >
          <span class="material-symbols-rounded">check_circle</span>
        </div>

        <h2 style="margin-top:20px;">Check-in Berhasil</h2>

        <p>
          Kehadiran <strong>${escapeHtml(result.nama)}</strong>
          pada kegiatan
          <strong>${escapeHtml(result.kegiatan)}</strong>
          berhasil dicatat.
        </p>
      </div>
    `;

  } catch (error) {
    showToast(error.message, 'error');

    button.disabled = false;
    button.innerHTML = `
      <span class="material-symbols-rounded">how_to_reg</span>
      Check-in Sekarang
    `;
  }
}


/* =====================================================
   MODAL
===================================================== */

function customConfirm(message, title = 'Konfirmasi', confirmLabel = 'Lanjutkan') {
  return new Promise(resolve => {
    if (pendingCustomDialog) pendingCustomDialog.resolve(false);
    pendingCustomDialog = { resolve };
    openCustomModal(title, 'Periksa tindakan ini sebelum melanjutkan.', `<div class="custom-dialog-content"><div class="custom-dialog-message">${message}</div><div class="custom-dialog-actions"><button type="button" class="btn btn-light" onclick="resolveCustomDialog(false)">Batal</button><button type="button" class="btn btn-primary" onclick="resolveCustomDialog(true)">${escapeHtml(confirmLabel)}</button></div></div>`, true);
  });
}
function customPrompt(message, title = 'Input Diperlukan') {
  return new Promise(resolve => {
    if (pendingCustomDialog) pendingCustomDialog.resolve(false);
    pendingCustomDialog = { resolve };
    openCustomModal(title, 'Lengkapi informasi untuk melanjutkan.', `<div class="custom-dialog-content"><div class="custom-dialog-message">${escapeHtml(message)}</div><textarea id="customDialogInput" class="form-control" rows="4" maxlength="500" placeholder="Tuliskan keterangan..."></textarea><div class="custom-dialog-actions"><button type="button" class="btn btn-light" onclick="resolveCustomDialog('')">Batal</button><button type="button" class="btn btn-primary" onclick="resolveCustomDialog(document.getElementById('customDialogInput')?.value || '')">Simpan</button></div></div>`, true);
  });
}
function resolveCustomDialog(value) {
  const dialog = pendingCustomDialog; pendingCustomDialog = null; closeModal(); if (dialog) dialog.resolve(value);
}

function openCustomModal(
  title,
  subtitle,
  body,
  small = false
) {
  const modal = document.getElementById('modal');

  modal.classList.toggle('small', !!small);

  modalLoadGeneration++;
  document.getElementById('modalTitle').textContent = title;
  document.getElementById('modalSubtitle').textContent = subtitle;
  document.getElementById('modalBody').innerHTML = body;

  const backdrop = document.getElementById('modalBackdrop');
  backdrop.classList.add('show');
  backdrop.setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');
}

function closeModal() {
  if (pendingCustomDialog) { const dialog = pendingCustomDialog; pendingCustomDialog = null; dialog.resolve(false); }
  modalLoadGeneration++;
  const backdrop = document.getElementById('modalBackdrop');
  const modal = document.getElementById('modal');

  if (backdrop) {
    backdrop.classList.remove('show');
    backdrop.setAttribute('aria-hidden', 'true');
  }

  if (modal) {
    modal.classList.remove('small');
    modal.classList.remove('batch-attendance-modal');
    modal.classList.remove('skk-checklist-modal');
  }

  if (document.body) document.body.classList.remove('modal-open');
  repairScrollLockState();
}


/* =====================================================
   HELPERS
===================================================== */

function getItem(type, id) {
  return (state.data[type] || []).find(item =>
    String(item.ID) === String(id)
  ) || {};
}

function isRole(role) {
  return (
    state.user &&
    String(state.user.Role).toUpperCase() ===
    String(role).toUpperCase()
  );
}

function isAnyRole(roles) {
  return roles.some(role => isRole(role));
}

function canPermission(module, action) {
  if (isRole('ADMIN')) return true;
  if (!isRole('PENGURUS')) return false;
  const p = state.permissions[module] || {};
  return p.view === true && p[action] === true;
}

function canViewModule(module) {
  if (module === 'dashboard' || module === 'settings') return true;
  if (module === 'maintenance') return isRole('ADMIN');
  if (isRole('ANGGOTA')) {
    return ['member-profile','member-attendance','member-skk'].includes(module);
  }
  if (module === 'users') return isRole('ADMIN');
  return canPermission(module, 'view');
}

function statCard(icon, color, label, value, meta = '', action = '') {
  const clickable = !!action;
  return `
    <div
      class="stat-card ${clickable ? 'stat-card-clickable' : ''}"
      ${clickable ? `onclick="${action}" role="button" tabindex="0" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();${action}}"` : ''}
    >
      <div class="stat-icon ${color}">
        <span class="material-symbols-rounded">${icon}</span>
      </div>

      <div class="stat-card-copy">
        <small>${escapeHtml(label)}</small>
        <h3>${escapeHtml(value)}</h3>
        ${meta ? `<p>${escapeHtml(meta)}</p>` : ''}
      </div>
    </div>
  `;
}

function formatRupiah(value) {
  return new Intl.NumberFormat(
    'id-ID',
    {
      style: 'currency',
      currency: 'IDR',
      maximumFractionDigits: 0
    }
  ).format(Number(value || 0));
}

function formatDate(value) {
  if (!value) {
    return '-';
  }

  const parts = String(value)
    .substring(0,10)
    .split('-');

  if (parts.length !== 3) {
    return escapeHtml(value);
  }

  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

function statusBadge(status) {
  const value = String(status || '-');
  const lower = value.toLowerCase();

  let color = 'gray';

  if (
    ['aktif','selesai','diarsipkan','hadir'].includes(lower)
  ) {
    color = 'green';

  } else if (
    ['rencana','berjalan','diproses','izin'].includes(lower)
  ) {
    color = 'blue';

  } else if (
    ['nonaktif','dibatalkan','alpa'].includes(lower)
  ) {
    color = 'red';

  } else if (
    ['calon anggota'].includes(lower)
  ) {
    color = 'purple';

  } else if (
    ['alumni','draft','sakit'].includes(lower)
  ) {
    color = 'gold';
  }

  return `<span class="badge ${color}">${escapeHtml(value)}</span>`;
}

function attendanceBadge(status) {
  const map = {
    Hadir: 'green',
    Izin: 'blue',
    Sakit: 'gold',
    Alpa: 'red',
    Libur: 'gray'
  };

  return `
    <span class="badge ${map[status] || 'gray'}">
      ${escapeHtml(status || '-')}
    </span>
  `;
}

function cashBadge(type) {
  return `
    <span class="badge ${type === 'Pemasukan' ? 'green' : 'red'}">
      ${escapeHtml(type || '-')}
    </span>
  `;
}

function conditionBadge(value) {
  const map = {
    'Baik': 'green',
    'Rusak Ringan': 'gold',
    'Rusak Berat': 'red',
    'Hilang': 'red'
  };

  return `
    <span class="badge ${map[value] || 'gray'}">
      ${escapeHtml(value || '-')}
    </span>
  `;
}

function mailBadge(value) {
  return `
    <span class="badge ${
      value === 'Surat Masuk'
        ? 'blue'
        : 'gold'
    }">
      ${escapeHtml(value || '-')}
    </span>
  `;
}

function emptyTableRow(colspan, message) {
  return `
    <tr>
      <td colspan="${colspan}">
        ${emptyStateHtml('inventory_2', message)}
      </td>
    </tr>
  `;
}

function emptyStateHtml(icon, message) {
  return `
    <div class="empty-state">
      <span class="material-symbols-rounded">${icon}</span>
      <h3>Data masih kosong</h3>
      <p>${escapeHtml(message)}</p>
    </div>
  `;
}

function systemLoadingHtml() {
  return `<section class="system-loading" role="status" aria-live="polite" aria-busy="true" aria-labelledby="systemLoadingTitle">
      <div class="system-loading-card">
        <div class="system-loading-emblem" aria-hidden="true">
          <svg viewBox="0 0 24 24" width="36" height="36" fill="currentColor"><path d="M21 16v-2l-8-5V3.5a1.5 1.5 0 0 0-3 0V9l-8 5v2l8-2.5V19l-2 1.5V22l3.5-1 3.5 1v-1.5L13 19v-5.5z"/></svg>
        </div>
        <p class="system-loading-brand">SAKA DIRGANTARA</p>
        <h1 id="systemLoadingTitle">Memuat sistem<span class="loading-dots" aria-hidden="true"><i></i><i></i><i></i></span></h1>
        <p class="system-loading-description">Menyiapkan dashboard organisasi Anda.</p>
        <div class="system-loading-track" aria-hidden="true"><span></span></div>
        <p class="system-loading-status" id="systemLoadingStatus">Menunggu respons sistem...</p>
        <small id="systemLoadingElapsed" aria-live="off">Waktu tunggu: 0 detik</small>
      </div>
    </section>`;
}

function stopSystemLoading() {
  if (systemLoadingTimer !== null) clearInterval(systemLoadingTimer);
  systemLoadingTimer = null;
}

function startSystemLoading() {
  stopSystemLoading();
  document.getElementById('appRoot').innerHTML = systemLoadingHtml();
  const startedAt = Date.now();
  systemLoadingTimer = setInterval(() => {
    const elapsedLabel = document.getElementById('systemLoadingElapsed');
    if (!elapsedLabel) {
      stopSystemLoading();
      return;
    }
    const seconds = Math.floor((Date.now() - startedAt) / 1000);
    elapsedLabel.textContent = 'Waktu tunggu: ' + seconds + ' detik';
    if (seconds >= 15) {
      const status = document.getElementById('systemLoadingStatus');
      if (status && status.textContent !== 'Pemuatan lebih lama dari biasanya. Masih menunggu respons...') {
        status.textContent = 'Pemuatan lebih lama dari biasanya. Masih menunggu respons...';
      }
    }
  }, 1000);
}

function loadingHtml(height = 300) {
  return `
    <div class="loading" style="min-height:${height}px;" role="status" aria-live="polite" aria-busy="true">
      <span class="loading-dots" aria-hidden="true"><i></i><i></i><i></i></span>
      <span class="loading-label">Memuat data...</span>
    </div>
  `;
}


function showLoading(options = {}) {
  const content = document.getElementById('content');
  if (!content) return;

  const preserve = !!options.preserve;
  const label = escapeHtml(options.label || 'Memuat data...');

  if (preserve && content.children.length) {
    content.classList.add('content-is-refreshing');
    const existing = content.querySelector('.content-refresh-overlay');
    if (existing) {
      const labelNode = existing.querySelector('.loading-label');
      if (labelNode) labelNode.textContent = options.label || 'Memuat data...';
      return;
    }

    content.insertAdjacentHTML('beforeend', `
      <div class="content-refresh-overlay" role="status" aria-live="polite" aria-busy="true">
        <div class="content-refresh-card">
          <span class="material-symbols-rounded spin" aria-hidden="true">progress_activity</span>
          <div>
            <strong>Menyegarkan tampilan</strong>
            <span class="loading-label">${label}</span>
          </div>
        </div>
      </div>
    `);
    return;
  }

  content.innerHTML = loadingHtml();
}

function hideLoadingOverlay() {
  const content = document.getElementById('content');
  if (!content) return;
  content.classList.remove('content-is-refreshing');
  const overlay = content.querySelector('.content-refresh-overlay');
  if (overlay) overlay.remove();
}

function renderError(message) {
  const content = document.getElementById('content');

  if (content) {
    content.innerHTML = `
      <div class="setup-card">
        <div
          class="setup-icon"
          style="background:#c34651;"
        >
          <span class="material-symbols-rounded">error</span>
        </div>

        <h2>Terjadi Kesalahan</h2>
        <p>${escapeHtml(message)}</p>

        <button
          class="btn btn-primary"
          onclick="refreshData()"
        >
          Coba Lagi
        </button>
      </div>
    `;
  }
}

function errorBox(message) {
  return `
    <div class="empty-state">
      <span
        class="material-symbols-rounded"
        style="color:#d8505b;"
      >
        error
      </span>

      <h3>Terjadi Kesalahan</h3>
      <p>${escapeHtml(message)}</p>
    </div>
  `;
}

function showToast(message, type = 'success') {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const allowed = ['success','error','warning','info'];
  const normalizedType = allowed.includes(type) ? type : 'info';
  // Existing completion notices take precedence over the generic fallback.
  if (activityState.cycle) {
    activityState.cycle.notified = Math.max(activityState.cycle.notified, ({ success:1, warning:2, error:3 }[normalizedType] || 0));
  }
  const icons = {
    success: 'check_circle',
    error: 'error',
    warning: 'warning',
    info: 'info'
  };

  const toast = document.createElement('div');
  toast.className = `toast ${normalizedType}`;
  toast.setAttribute('role', normalizedType === 'error' ? 'alert' : 'status');
  toast.setAttribute(
    'aria-live',
    normalizedType === 'error' ? 'assertive' : 'polite'
  );
  toast.innerHTML = `
    <span class="material-symbols-rounded toast-icon">${icons[normalizedType]}</span>
    <div class="toast-message">${escapeHtml(message)}</div>
    <button type="button" class="toast-close" aria-label="Tutup notifikasi">×</button>
  `;

  const closeButton = toast.querySelector('.toast-close');
  if (closeButton) {
    closeButton.addEventListener('click', () => toast.remove());
  }

  while (container.children.length >= 3) container.firstElementChild.remove();
  container.appendChild(toast);

  setTimeout(() => {
    if (toast.parentElement) toast.remove();
  }, 4500);
}

async function copyText(text) {
  const value = String(text ?? '');

  try {
    if (navigator.clipboard && window.isSecureContext) {
      await navigator.clipboard.writeText(value);
    } else {
      const textarea = document.createElement('textarea');
      textarea.value = value;
      textarea.setAttribute('readonly', '');
      textarea.style.position = 'fixed';
      textarea.style.opacity = '0';
      document.body.appendChild(textarea);
      textarea.select();

      const copied = document.execCommand('copy');
      textarea.remove();

      if (!copied) throw new Error('Copy gagal');
    }

    showToast('Link berhasil disalin.');
  } catch (_) {
    showToast('Gagal menyalin link.', 'error');
  }
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function safeHttpUrl(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return '';

  try {
    const parsed = new URL(raw);
    return ['http:', 'https:'].includes(parsed.protocol) ? parsed.href : '';
  } catch (_) {
    return '';
  }
}

function escapeJs(value) {
  return String(value ?? '')
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}


/* =====================================================
   DETAIL & LAPORAN KEGIATAN
===================================================== */

function izinVerificationBadge(status) {
  const value = String(status || 'Menunggu Verifikasi');
  if (value === 'Disetujui') return '<span class="badge green">Disetujui</span>';
  if (value === 'Ditolak') return '<span class="badge red">Ditolak</span>';
  return '<span class="badge gold">Menunggu Verifikasi</span>';
}

async function verifyOnlineIzin(izinId, decision, kegiatanId) {
  const approve = decision === 'Disetujui';
  let note = '';
  if (approve) {
    if (!await customConfirm('Setujui pengajuan ini? Setelah disetujui, status Izin/Sakit akan dipakai sebagai prefill pada Absensi Massal.', 'Setujui Pengajuan', 'Setujui')) return;
  } else {
    note = await customPrompt('Tuliskan alasan penolakan pengajuan:', 'Alasan Penolakan');
    if (!note.trim()) {
      showToast('Alasan penolakan wajib diisi.', 'warning');
      return;
    }
  }
  try {
    await serverCall('verifyIzinKegiatan', state.token, izinId, decision, note.trim());
    showToast(
      approve
        ? 'Pengajuan disetujui. Absensi Massal akan memprefill Izin/Sakit.'
        : 'Pengajuan ditolak. Absensi Massal akan memprefill Alpa.',
      'success'
    );
    markModuleDirty('absensi');
    await openKegiatanDetail(kegiatanId);
  } catch (error) {
    showToast(error.message, 'error');
  }
}

async function openKegiatanDetail(kegiatanId) {
  const kegiatan = getItem('kegiatan', kegiatanId);

  openCustomModal(
    'Detail & Laporan Kegiatan',
    kegiatan.NamaKegiatan
      ? `${kegiatan.NamaKegiatan} • ${formatDate(kegiatan.Tanggal)}`
      : 'Kelola laporan, kehadiran, keuangan, dan dokumentasi kegiatan.',
    loadingHtml(260)
  );

  try {
    // Detail kegiatan hanya membutuhkan katalog inventaris sebagai data pilihan.
    // Katalog dimuat saat detail benar-benar dibuka, bukan saat membuka menu Kegiatan.
    if (!await ensureModulesLoaded(['inventaris'])) return;

    const workspace = await serverCall('getKegiatanReportData', state.token, kegiatanId);

    const report = workspace.laporan || {};
    const attendance = workspace.kehadiran || { rows: [], summary: {} };
    const finance = workspace.keuangan || { rows: [] };
    const officers = Array.isArray(workspace.pengurus) ? workspace.pengurus : [];
    const docs = Array.isArray(workspace.dokumentasi) ? workspace.dokumentasi : [];
    const izinRows = Array.isArray(workspace.izin) ? workspace.izin : [];
    const activityInventory = Array.isArray(workspace.inventarisKegiatan)
      ? workspace.inventarisKegiatan
      : [];
    const inventoryCatalog = Array.isArray(workspace.katalogInventaris)
      ? workspace.katalogInventaris
      : (Array.isArray(state.data.inventaris) ? state.data.inventaris : []);

    const officerOptions = officers.map(item => {
      const text = [item.Nama, item.Jabatan].filter(Boolean).join(' • ');
      return `<option value="${escapeHtml(item.Nama || '')}">${escapeHtml(text)}</option>`;
    }).join('');

    const hadirRows = (attendance.rows || []).filter(item => String(item.Status || '').toLowerCase() === 'hadir');
    const rows = hadirRows.length
      ? hadirRows.map((item, index) => `
          <tr>
            <td>${index + 1}</td>
            <td><span class="table-name">${escapeHtml(item.Nama || '-')}</span></td>
            <td>${escapeHtml(item.NTA || '-')}</td>
            <td>${escapeHtml(item.JenisKelamin || '-')}</td>
            <td>${escapeHtml(item.Waktu || '-')}</td>
          </tr>
        `).join('')
      : emptyTableRow(5, 'Belum ada peserta hadir.');

    const izinRowsHtml = izinRows.length
      ? izinRows.map((item, index) => `
          <tr>
            <td>${index + 1}</td>
            <td><span class="table-name">${escapeHtml(item.NamaAnggota || '-')}</span><small class="batch-attendance-member-meta">${escapeHtml(item.NTA || '-')}</small></td>
            <td><span class="badge ${String(item.JenisPengajuan || 'Izin') === 'Sakit' ? 'red' : 'blue'}">${escapeHtml(item.JenisPengajuan || 'Izin')}</span><small class="batch-attendance-member-meta">${escapeHtml(item.JenisBukti || '-')}</small></td>
            <td>${escapeHtml(item.Alasan || '-')}</td>
            <td>${izinVerificationBadge(item.Status)}${item.CatatanVerifikasi ? `<small class="batch-attendance-member-meta">${escapeHtml(item.CatatanVerifikasi)}</small>` : ''}${item.DiverifikasiOleh ? `<small class="batch-attendance-member-meta">oleh ${escapeHtml(item.DiverifikasiOleh)}</small>` : ''}</td>
            <td>${escapeHtml(item.DikirimPada || '-')}</td>
            <td><button type="button" class="btn btn-light btn-small" onclick="viewIzinEvidence('${escapeJs(item.ID)}')"><span class="material-symbols-rounded">attach_file</span> Bukti</button></td>
            <td>
              <div class="row-actions">
                <button type="button" class="btn-icon" title="Setujui" onclick="verifyOnlineIzin('${escapeJs(item.ID)}','Disetujui','${escapeJs(kegiatanId)}')"><span class="material-symbols-rounded">check_circle</span></button>
                <button type="button" class="btn-icon delete" title="Tolak" onclick="verifyOnlineIzin('${escapeJs(item.ID)}','Ditolak','${escapeJs(kegiatanId)}')"><span class="material-symbols-rounded">cancel</span></button>
              </div>
            </td>
          </tr>
        `).join('')
      : emptyTableRow(8, 'Belum ada pengajuan izin/sakit online.');

    const schedule = Array.isArray(report.SusunanKegiatan) ? report.SusunanKegiatan : [];
    const scheduleHtml = schedule.length
      ? schedule.map((item, index) => renderKegiatanReportScheduleRow(item, index)).join('')
      : renderKegiatanReportScheduleRow({}, 0);

    const financeRows = (finance.rows || []).length
      ? finance.rows.map((item, index) => `
          <tr>
            <td>${index + 1}</td>
            <td>${escapeHtml(formatDate(item.Tanggal))}</td>
            <td>${escapeHtml(item.NoBukti || '-')}</td>
            <td>${escapeHtml(item.Uraian || '-')}</td>
            <td>${escapeHtml(formatRupiah(item.Penerimaan || 0))}</td>
            <td>${escapeHtml(formatRupiah(item.Pengeluaran || 0))}</td>
            <td>${escapeHtml(formatRupiah(item.Saldo || 0))}</td>
          </tr>
        `).join('')
      : emptyTableRow(7, 'Belum ada transaksi Kas yang dihubungkan ke kegiatan ini.');

    const inventoryOptions = inventoryCatalog
      .slice()
      .sort((a, b) => String(a.NamaBarang || '').localeCompare(String(b.NamaBarang || ''), 'id'))
      .map(item => `<option value="${escapeHtml(item.ID || '')}">${escapeHtml([item.KodeBarang, item.NamaBarang].filter(Boolean).join(' — ') || 'Inventaris')}</option>`)
      .join('');

    const inventoryRowsHtml = activityInventory.length
      ? activityInventory.map(item => {
          const awalPhoto = String(item.FotoAwalID || '');
          const afterPhoto = String(item.FotoSetelahID || '');
          return `
            <article class="activity-inventory-card">
              <div class="activity-inventory-card-head">
                <div>
                  <span class="kegiatan-detail-eyebrow">${escapeHtml(item.KodeBarang || 'INVENTARIS')}</span>
                  <h4>${escapeHtml(item.NamaBarang || '-')}</h4>
                </div>
                <div class="row-actions">${item.StatusPemakaian !== 'Selesai' ? `<button type="button" class="btn-icon" title="Edit / catat pengembalian" aria-label="Edit / catat pengembalian" onclick="editKegiatanInventory('${escapeJs(item.ID)}','${escapeJs(kegiatanId)}')"><span class="material-symbols-rounded">edit</span></button>` : ''}<button type="button" class="btn-icon delete" title="Hapus riwayat pemakaian" aria-label="Hapus riwayat pemakaian" onclick="deleteKegiatanInventory('${escapeJs(item.ID)}','${escapeJs(kegiatanId)}')"><span class="material-symbols-rounded">delete</span></button></div>
              </div>
              <div class="activity-inventory-compare">
                <div class="inventory-quantity-block">
                  <span>Awal</span><strong>${escapeHtml(item.JumlahAwal || 0)} unit</strong><small>${escapeHtml(item.KondisiAwal || 'Baik')}</small>
                </div>
                <div class="inventory-quantity-block is-used">
                  <span>Dipakai</span><strong>${escapeHtml(item.JumlahDipakai || 0)} unit</strong><small>selama kegiatan</small>
                </div>
                <div class="inventory-quantity-block">
                  <span>Setelah</span><strong>${escapeHtml(item.JumlahSetelah || 0)} unit</strong><small>${escapeHtml(item.KondisiSetelah || 'Baik')}</small>
                </div>
              </div>
              <div class="activity-inventory-photos">
                ${inventoryPhotoBox(item.ID, kegiatanId, 'Awal', awalPhoto, 'Foto kondisi awal')}
                ${inventoryPhotoBox(item.ID, kegiatanId, 'Setelah', afterPhoto, 'Foto kondisi setelah kegiatan')}
              </div>
              ${item.Keterangan ? `<p class="activity-inventory-note">${escapeHtml(item.Keterangan)}</p>` : ''}
            </article>
          `;
        }).join('')
      : `<div class="inventory-empty-state"><span class="material-symbols-rounded">inventory_2</span><strong>Belum ada inventaris yang dicatat untuk kegiatan ini.</strong><p>Tambahkan barang yang dipakai, lalu isi jumlah serta kondisi awal dan setelah kegiatan.</p></div>`;

    const documentationHtml = docs.length
      ? docs.map(item => `
          <article class="documentation-card" id="documentationCard-${escapeHtml(item.ID)}">
            <div class="documentation-image-wrap">
              <div class="documentation-image-loading" id="documentationLoading-${escapeHtml(item.ID)}">
                <div class="spinner"></div>
                <span>Memuat foto...</span>
              </div>
              <img id="documentationImage-${escapeHtml(item.ID)}" alt="Dokumentasi ${escapeHtml(kegiatan.NamaKegiatan || 'kegiatan')}" loading="lazy">
            </div>
            <div class="documentation-card-body documentation-card-report-body">
              <div class="documentation-card-copy">
                <strong>${escapeHtml(item.NamaFile || 'Dokumentasi')}</strong>
                <span>${escapeHtml(item.Uploader || '-')} • ${escapeHtml(item.DibuatPada || '-')}</span>
              </div>
              <div class="documentation-report-meta">
                <label>
                  <span>Urutan</span>
                  <input id="documentationOrder-${escapeHtml(item.ID)}" class="documentation-order-input" type="number" min="1" max="999" value="${escapeHtml(item.Urutan || 1)}">
                </label>
                <label class="documentation-caption-label">
                  <span>Keterangan lampiran</span>
                  <input id="documentationCaption-${escapeHtml(item.ID)}" class="documentation-caption-input" type="text" maxlength="500" value="${escapeHtml(item.Keterangan || '')}" placeholder="Contoh: Penyampaian materi navigasi udara">
                </label>
                <button type="button" class="btn btn-light btn-compact" onclick="saveKegiatanDocumentationMetadata('${escapeJs(item.ID)}')">
                  <span class="material-symbols-rounded">save</span> Simpan
                </button>
                <button type="button" class="btn-icon delete" title="Hapus dokumentasi" aria-label="Hapus dokumentasi" onclick="deleteKegiatanDocumentation('${escapeJs(item.ID)}','${escapeJs(kegiatanId)}')">
                  <span class="material-symbols-rounded">delete</span>
                </button>
              </div>
            </div>
          </article>
        `).join('')
      : `
          <div class="documentation-empty">
            <span class="material-symbols-rounded">photo_library</span>
            <strong>Belum ada dokumentasi</strong>
            <p>Tambahkan foto agar otomatis masuk Lampiran II pada PDF laporan.</p>
          </div>
        `;

    const body = document.getElementById('modalBody');
    if (body) {
      body.innerHTML = `
        <div class="kegiatan-report-shell">
          <section class="report-overview-card">
            <div>
              <span class="kegiatan-detail-eyebrow">LAPORAN KEGIATAN</span>
              <h3>${escapeHtml(kegiatan.NamaKegiatan || '-')}</h3>
              <p>${escapeHtml(formatDate(kegiatan.Tanggal))} • ${escapeHtml(kegiatan.Lokasi || 'Lokasi belum ditentukan')}</p>
            </div>
            <div class="report-overview-actions">
              ${report.PDFUrl ? `<a class="btn btn-light" target="_blank" rel="noopener" href="${escapeHtml(report.PDFUrl)}"><span class="material-symbols-rounded">picture_as_pdf</span> PDF Terakhir</a>` : ''}
              <button type="button" class="btn btn-primary" onclick="saveAndGenerateKegiatanReport('${escapeJs(kegiatanId)}')">
                <span class="material-symbols-rounded">picture_as_pdf</span> Generate PDF
              </button>
            </div>
          </section>

          <form id="kegiatanReportForm" class="kegiatan-report-form" onsubmit="saveKegiatanReport(event,'${escapeJs(kegiatanId)}')">
            <section class="kegiatan-detail-section report-section-card">
              <div class="kegiatan-detail-section-head">
                <div><span class="kegiatan-detail-eyebrow">A</span><h3>Waktu & Tempat</h3><p>Tanggal dan lokasi mengikuti data Kegiatan.</p></div>
              </div>
              <div class="report-form-grid">
                <label class="report-field"><span>Jam Mulai</span><input class="form-control" type="time" name="JamMulai" value="${escapeHtml(report.JamMulai || '')}"></label>
                <label class="report-field"><span>Jam Selesai</span><input class="form-control" type="time" name="JamSelesai" value="${escapeHtml(report.JamSelesai || '')}"></label>
                <label class="report-field"><span>Cuaca</span><select class="form-control" name="Cuaca">${['Cerah','Berawan','Mendung','Hujan','Hujan Lebat','Lainnya'].map(v => `<option ${String(report.Cuaca || 'Cerah') === v ? 'selected' : ''}>${v}</option>`).join('')}</select></label>
                <label class="report-field"><span>Pamong / Instruktur Hadir</span><input class="form-control" type="number" min="0" max="99" name="JumlahPamongInstruktur" value="${escapeHtml(report.JumlahPamongInstruktur ?? 0)}"></label>
              </div>
            </section>

            <section class="kegiatan-detail-section report-section-card">
              <div class="kegiatan-detail-section-head"><div><span class="kegiatan-detail-eyebrow">B</span><h3>Ringkasan Kehadiran</h3><p>Angka anggota dihitung otomatis dari Absensi.</p></div></div>
              <div class="report-attendance-summary">
                ${reportSummaryBox('how_to_reg','Hadir',attendance.summary?.hadir || 0)}
                ${reportSummaryBox('male','Putra',attendance.summary?.putra || 0)}
                ${reportSummaryBox('female','Putri',attendance.summary?.putri || 0)}
                ${reportSummaryBox('event_available','Izin',attendance.summary?.izin || 0)}
                ${reportSummaryBox('sick','Sakit',attendance.summary?.sakit || 0)}
                ${reportSummaryBox('person_off','Alpa',attendance.summary?.alpa || 0)}
              </div>
            </section>

            <section class="kegiatan-detail-section report-section-card">
              <div class="kegiatan-detail-section-head"><div><span class="kegiatan-detail-eyebrow">C</span><h3>Materi Latihan</h3></div></div>
              <div class="report-form-grid">
                <label class="report-field"><span>Fokus Krida</span><select class="form-control" name="FokusKrida">${['Krida Pengetahuan Kedirgantaraan','Krida Olahraga Dirgantara','Krida Jasa Kedirgantaraan','Umum / Gabungan','Lainnya'].map(v => `<option ${String(report.FokusKrida || '') === v ? 'selected' : ''}>${v}</option>`).join('')}</select></label>
                <label class="report-field report-field-full"><span>Topik Materi</span><input class="form-control" name="TopikMateri" maxlength="500" value="${escapeHtml(report.TopikMateri || '')}" placeholder="Contoh: Pengenalan Dasar Aeromodelling"></label>
                <label class="report-field report-field-full"><span>Instruktur / Pemateri</span><input class="form-control" name="Instruktur" list="reportOfficerList" maxlength="300" value="${escapeHtml(report.Instruktur || '')}" placeholder="Pilih pengurus atau ketik pihak luar"><datalist id="reportOfficerList">${officerOptions}</datalist></label>
              </div>
            </section>

            <section class="kegiatan-detail-section report-section-card">
              <div class="kegiatan-detail-section-head report-schedule-head">
                <div><span class="kegiatan-detail-eyebrow">D</span><h3>Susunan Kegiatan</h3><p>Maksimal 30 agenda.</p></div>
                <button class="btn btn-light btn-compact" type="button" onclick="addKegiatanReportScheduleRow()"><span class="material-symbols-rounded">add</span> Tambah Agenda</button>
              </div>
              <div id="reportScheduleRows" class="report-schedule-list">${scheduleHtml}</div>
            </section>

            <section class="kegiatan-detail-section report-section-card">
              <div class="kegiatan-detail-section-head"><div><span class="kegiatan-detail-eyebrow">E</span><h3>Evaluasi & Catatan</h3></div></div>
              <div class="report-form-grid">
                <label class="report-field report-field-full"><span>Pencapaian</span><textarea class="form-control" name="Pencapaian" rows="3">${escapeHtml(report.Pencapaian || '')}</textarea></label>
                <label class="report-field report-field-full"><span>Kendala</span><textarea class="form-control" name="Kendala" rows="3">${escapeHtml(report.Kendala || '')}</textarea></label>
                <label class="report-field report-field-full"><span>Tindak Lanjut / Saran</span><textarea class="form-control" name="TindakLanjut" rows="3">${escapeHtml(report.TindakLanjut || '')}</textarea></label>
              </div>
            </section>

            <section class="kegiatan-detail-section report-section-card">
              <div class="kegiatan-detail-section-head"><div><span class="kegiatan-detail-eyebrow">F</span><h3>Laporan Keuangan</h3><p>Transaksi berasal dari Kas yang dihubungkan ke kegiatan ini.</p></div></div>
              <div class="report-finance-summary">
                <span>Saldo awal <strong>${escapeHtml(formatRupiah(finance.saldoAwal || 0))}</strong></span>
                <span>Penerimaan <strong class="is-positive">${escapeHtml(formatRupiah(finance.totalPenerimaan || 0))}</strong></span>
                <span>Pengeluaran <strong class="is-negative">${escapeHtml(formatRupiah(finance.totalPengeluaran || 0))}</strong></span>
                <span>Saldo akhir <strong>${escapeHtml(formatRupiah(finance.saldoAkhir || 0))}</strong></span>
              </div>
              <div class="table-wrap report-finance-table"><table class="data-table"><thead><tr><th>No</th><th>Tanggal</th><th>No Bukti</th><th>Uraian</th><th>Penerimaan</th><th>Pengeluaran</th><th>Saldo</th></tr></thead><tbody>${financeRows}</tbody></table></div>
            </section>

            <section class="kegiatan-detail-section report-section-card">
              <div class="kegiatan-detail-section-head"><div><span class="kegiatan-detail-eyebrow">PENGESAHAN</span><h3>Tempat, Tanggal & Penandatangan</h3></div></div>
              <div class="report-form-grid">
                <label class="report-field"><span>Tempat Laporan</span><input class="form-control" name="TempatLaporan" value="${escapeHtml(report.TempatLaporan || 'Sidoarjo')}"></label>
                <label class="report-field"><span>Tanggal Laporan</span><input class="form-control" type="date" name="TanggalLaporan" value="${escapeHtml(report.TanggalLaporan || '')}"></label>
                <label class="report-field"><span>Penandatangan 1</span><input class="form-control" name="Penandatangan1Nama" list="reportOfficerList" value="${escapeHtml(report.Penandatangan1Nama || '')}"></label>
                <label class="report-field"><span>Jabatan Penandatangan 1</span><input class="form-control" name="Penandatangan1Jabatan" value="${escapeHtml(report.Penandatangan1Jabatan || '')}"></label>
                <label class="report-field"><span>Penandatangan 2</span><input class="form-control" name="Penandatangan2Nama" list="reportOfficerList" value="${escapeHtml(report.Penandatangan2Nama || '')}"></label>
                <label class="report-field"><span>Jabatan Penandatangan 2</span><input class="form-control" name="Penandatangan2Jabatan" value="${escapeHtml(report.Penandatangan2Jabatan || '')}"></label>
              </div>
              <div class="report-save-actions"><button type="submit" class="btn btn-light"><span class="material-symbols-rounded">save</span> Simpan Data Laporan</button><button type="button" class="btn btn-primary" onclick="saveAndGenerateKegiatanReport('${escapeJs(kegiatanId)}')"><span class="material-symbols-rounded">picture_as_pdf</span> Simpan & Generate PDF</button></div>
            </section>
          </form>

          <section class="kegiatan-detail-section izin-detail-section">
            <div class="kegiatan-detail-section-head"><div><span class="kegiatan-detail-eyebrow">IZIN / SAKIT ONLINE</span><h3>Verifikasi Pengajuan Anggota</h3><p>${escapeHtml(izinRows.length)} pengajuan tersimpan. Hanya pengajuan berstatus Disetujui yang masuk sebagai prefill Absensi Massal.</p></div></div>
            <div class="table-wrap"><table class="data-table"><thead><tr><th>No</th><th>Anggota</th><th>Jenis / Bukti</th><th>Alasan</th><th>Verifikasi</th><th>Dikirim</th><th>Bukti</th><th>Aksi</th></tr></thead><tbody>${izinRowsHtml}</tbody></table></div>
          </section>

          <section class="kegiatan-detail-section">
            <div class="kegiatan-detail-section-head"><div><span class="kegiatan-detail-eyebrow">LAMPIRAN I</span><h3>Daftar Peserta Hadir</h3><p>${escapeHtml(hadirRows.length)} peserta berstatus hadir.</p></div></div>
            <div class="table-wrap kegiatan-attendance-table"><table class="data-table"><thead><tr><th>No</th><th>Nama</th><th>NTA</th><th>JK</th><th>Waktu Check-in</th></tr></thead><tbody>${rows}</tbody></table></div>
          </section>

          <section class="kegiatan-detail-section activity-inventory-section">
            <div class="kegiatan-detail-section-head">
              <div><span class="kegiatan-detail-eyebrow">INVENTARIS</span><h3>Pemakaian Inventaris Kegiatan</h3><p>Catat barang yang digunakan, jumlah awal, jumlah dipakai, jumlah setelah kegiatan, dan bukti foto kondisinya.</p></div>
            </div>
            <form class="activity-inventory-form" id="activityInventoryForm" onsubmit="saveKegiatanInventory(event,'${escapeJs(kegiatanId)}')">
              <input type="hidden" name="ID" value="">
              <input type="hidden" name="Versi" value="">
              <div class="inventory-form-grid">
                <label class="report-field inventory-field-full"><span>Inventaris</span><select class="form-control" name="InventarisID" required><option value="">Pilih barang inventaris...</option>${inventoryOptions}</select></label>
                <label class="report-field"><span>Jumlah Awal</span><input class="form-control" type="number" name="JumlahAwal" min="0" step="1" required></label>
                <label class="report-field"><span>Jumlah Dipakai</span><input class="form-control" type="number" name="JumlahDipakai" min="0" step="1" required></label>
                <label class="report-field"><span>Jumlah Kembali / Setelah</span><input class="form-control" type="number" name="JumlahSetelah" min="0" step="1"></label>
                <label class="report-field"><span>Status</span><select class="form-control" name="StatusPemakaian"><option value="Dipakai">Dipakai</option><option value="Selesai">Selesai / dikembalikan</option></select></label>
                <label class="report-field"><span>Jumlah Rusak Saat Kembali</span><input class="form-control" type="number" name="JumlahRusakKembali" min="0" step="1" value="0"></label>
                <label class="report-field"><span>Kondisi Awal</span><select class="form-control" name="KondisiAwal"><option>Baik</option><option>Rusak Ringan</option><option>Rusak Berat</option><option>Hilang</option></select></label>
                <label class="report-field"><span>Kondisi Setelah Kegiatan</span><select class="form-control" name="KondisiSetelah"><option>Baik</option><option>Rusak Ringan</option><option>Rusak Berat</option><option>Hilang</option></select></label>
                <label class="report-field inventory-field-full"><span>Keterangan</span><input class="form-control" type="text" name="Keterangan" maxlength="500" placeholder="Contoh: 2 unit mengalami goresan ringan"></label>
              </div>
              <div class="inventory-form-actions"><button type="submit" class="btn btn-primary"><span class="material-symbols-rounded">add_box</span> Tambahkan Pemakaian</button></div>
            </form>
            <div class="activity-inventory-list">${inventoryRowsHtml}</div>
          </section>

          <section class="kegiatan-detail-section documentation-section">
            <div class="kegiatan-detail-section-head documentation-head"><div><span class="kegiatan-detail-eyebrow">LAMPIRAN II</span><h3>Dokumentasi Kegiatan</h3><p>${docs.length} foto tersimpan. Keterangan di bawah foto akan ikut masuk PDF.</p></div></div>
            <div class="documentation-upload-box">
              <div class="documentation-upload-input"><span class="material-symbols-rounded">add_photo_alternate</span><div><strong>Tambah dokumentasi</strong><span>JPG, PNG, atau WEBP. Maksimal 5 foto per batch.</span></div><input id="kegiatanDocumentationInput" type="file" accept="image/jpeg,image/png,image/webp" multiple onchange="handleKegiatanDocumentationSelection(this)"></div>
              <div class="documentation-upload-actions"><span id="kegiatanDocumentationSelectionText">Belum ada foto dipilih.</span><button id="kegiatanDocumentationUploadButton" type="button" class="btn btn-primary" disabled onclick="uploadKegiatanDocumentation('${escapeJs(kegiatanId)}')"><span class="material-symbols-rounded">cloud_upload</span> Unggah Foto</button></div>
              <div id="kegiatanDocumentationProgress" class="documentation-progress"></div>
            </div>
            <div class="documentation-gallery" id="documentationGallery">${documentationHtml}</div>
          </section>
        </div>
      `;
    }

    const inventoryPhotoItems = activityInventory.reduce((items, item) => {
      [item.FotoAwalID, item.FotoSetelahID].forEach(id => {
        if (id) items.push({ ID: id });
      });
      return items;
    }, []);
    if (docs.length || inventoryPhotoItems.length) {
      requestAnimationFrame(() => loadKegiatanDocumentationPreviews([...docs, ...inventoryPhotoItems]));
    }
  } catch (error) {
    const body = document.getElementById('modalBody');
    if (body) body.innerHTML = errorBox(error.message);
    showToast(error.message, 'error');
  }
}

async function viewIzinEvidence(izinId) {
  const previewWindow = window.open('', '_blank');
  if (previewWindow) {
    previewWindow.document.write('<!doctype html><html><head><title>Memuat bukti pengajuan...</title></head><body style="font-family:Arial,sans-serif;padding:24px">Memuat bukti pendukung...</body></html>');
  } else {
    openCustomModal('Bukti Pengajuan', 'Memuat bukti yang tersimpan...', loadingHtml(180), true);
  }

  try {
    const result = await serverCall('getIzinBuktiPreview', state.token, izinId);
    const mime = String(result.mimeType || '').toLowerCase();
    if (previewWindow) {
      const safeName = escapeHtml(result.name || 'Bukti pengajuan');
      const content = mime === 'application/pdf'
        ? `<iframe src="${escapeHtml(result.dataUrl)}" style="position:fixed;inset:0;width:100%;height:100%;border:0" title="${safeName}"></iframe>`
        : `<div style="min-height:100vh;display:flex;align-items:center;justify-content:center;background:#eef2f5"><img src="${escapeHtml(result.dataUrl)}" alt="${safeName}" style="max-width:100%;max-height:100vh;object-fit:contain"></div>`;
      previewWindow.document.open();
      previewWindow.document.write(`<!doctype html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${safeName}</title><style>html,body{margin:0;min-height:100%;background:#eef2f5}</style></head><body>${content}</body></html>`);
      previewWindow.document.close();
      return;
    }

    const body = document.getElementById('modalBody');
    const subtitle = document.getElementById('modalSubtitle');
    if (subtitle) subtitle.textContent = result.name || 'Bukti pengajuan';
    if (!body) return;
    body.innerHTML = mime === 'application/pdf'
      ? `<div class="izin-proof-view"><iframe class="izin-proof-frame" src="${escapeHtml(result.dataUrl)}" title="${escapeHtml(result.name || 'Bukti pengajuan')}"></iframe></div>`
      : `<div class="izin-proof-view"><img class="izin-proof-image" src="${escapeHtml(result.dataUrl)}" alt="${escapeHtml(result.name || 'Bukti pengajuan')}"></div>`;
  } catch (error) {
    if (previewWindow) {
      previewWindow.document.open();
      previewWindow.document.write(`<!doctype html><html><body style="font-family:Arial,sans-serif;padding:24px;color:#b4232f">${escapeHtml(error.message || 'Gagal memuat bukti.')}</body></html>`);
      previewWindow.document.close();
    } else {
      const body = document.getElementById('modalBody');
      if (body) body.innerHTML = errorBox(error.message);
    }
  }
}

function reportSummaryBox(icon, label, value) {
  return `<div class="report-summary-box"><span class="material-symbols-rounded">${icon}</span><div><strong>${escapeHtml(value)}</strong><span>${escapeHtml(label)}</span></div></div>`;
}

function inventoryPhotoBox(recordId, kegiatanId, phase, photoId, label) {
  const inputId = `inventoryPhoto-${escapeHtml(recordId)}-${escapeHtml(phase)}`;
  return `
    <div class="inventory-photo-box">
      <div class="inventory-photo-preview">
        ${photoId ? `
          <div class="documentation-image-loading" id="documentationLoading-${escapeHtml(photoId)}"><div class="spinner"></div><span>Memuat foto...</span></div>
          <img id="documentationImage-${escapeHtml(photoId)}" alt="${escapeHtml(label)}" loading="lazy">
        ` : `<span class="material-symbols-rounded">add_a_photo</span><small>Belum ada foto</small>`}
      </div>
      <div class="inventory-photo-copy"><strong>${escapeHtml(label)}</strong><span>Unggah untuk mengganti foto</span></div>
      <label class="btn btn-light btn-compact inventory-photo-upload" for="${inputId}"><span class="material-symbols-rounded">upload</span> Pilih Foto</label>
      <input id="${inputId}" class="inventory-photo-input" type="file" accept="image/jpeg,image/png,image/webp" onchange="handleKegiatanInventoryPhoto(this,'${escapeJs(kegiatanId)}','${escapeJs(recordId)}','${escapeJs(phase)}')">
    </div>
  `;
}

function renderKegiatanReportScheduleRow(item = {}, index = 0) {
  return `
    <div class="report-schedule-row">
      <input class="form-control report-schedule-start" type="time" value="${escapeHtml(item.mulai || '')}" aria-label="Jam mulai">
      <span class="report-schedule-dash">-</span>
      <input class="form-control report-schedule-end" type="time" value="${escapeHtml(item.selesai || '')}" aria-label="Jam selesai">
      <input class="form-control report-schedule-text" type="text" maxlength="500" value="${escapeHtml(item.kegiatan || '')}" placeholder="Uraian kegiatan">
      <button type="button" class="btn-icon delete" onclick="removeKegiatanReportScheduleRow(this)" title="Hapus agenda"><span class="material-symbols-rounded">delete</span></button>
    </div>
  `;
}

function addKegiatanReportScheduleRow() {
  const box = document.getElementById('reportScheduleRows');
  if (!box) return;
  if (box.querySelectorAll('.report-schedule-row').length >= 30) {
    showToast('Maksimal 30 susunan kegiatan.', 'warning');
    return;
  }
  box.insertAdjacentHTML('beforeend', renderKegiatanReportScheduleRow());
}

function removeKegiatanReportScheduleRow(button) {
  const row = button && button.closest ? button.closest('.report-schedule-row') : null;
  const box = document.getElementById('reportScheduleRows');
  if (row) row.remove();
  if (box && !box.querySelector('.report-schedule-row')) box.innerHTML = renderKegiatanReportScheduleRow();
}

function collectKegiatanReportPayload() {
  const form = document.getElementById('kegiatanReportForm');
  if (!form) throw new Error('Form laporan kegiatan belum tersedia.');
  const payload = Object.fromEntries(new FormData(form).entries());
  payload.SusunanKegiatan = Array.from(document.querySelectorAll('#reportScheduleRows .report-schedule-row')).map(row => ({
    mulai: row.querySelector('.report-schedule-start')?.value || '',
    selesai: row.querySelector('.report-schedule-end')?.value || '',
    kegiatan: row.querySelector('.report-schedule-text')?.value || ''
  })).filter(item => item.mulai || item.selesai || item.kegiatan);
  return payload;
}

async function saveKegiatanReport(event, kegiatanId) {
  if (event) event.preventDefault();
  try {
    const payload = collectKegiatanReportPayload();
    await serverCall('saveLaporanKegiatan', state.token, kegiatanId, payload);
    showToast('Data laporan kegiatan berhasil disimpan.', 'success');
  } catch (error) {
    showToast(error.message || 'Gagal menyimpan laporan kegiatan.', 'error');
  }
}

async function saveAndGenerateKegiatanReport(kegiatanId) {
  const buttons = Array.from(document.querySelectorAll('button[onclick*="saveAndGenerateKegiatanReport"]'));
  buttons.forEach(btn => btn.disabled = true);
  try {
    const payload = collectKegiatanReportPayload();
    await serverCall('saveLaporanKegiatan', state.token, kegiatanId, payload);
    showToast('Menyusun PDF laporan beserta lampiran...', 'info');
    const result = await serverCall('generateKegiatanReportPdf', state.token, kegiatanId);
    showToast(result.message || 'PDF laporan berhasil dibuat.', 'success');
    if (result.url) window.open(result.url, '_blank', 'noopener');
    await openKegiatanDetail(kegiatanId);
  } catch (error) {
    showToast(error.message || 'Gagal membuat PDF laporan kegiatan.', 'error');
  } finally {
    buttons.forEach(btn => btn.disabled = false);
  }
}

function editKegiatanInventory(recordId, kegiatanId) {
  const item = (state.data.kegiatanInventaris || []).find(row => String(row.ID) === String(recordId));
  const form = document.getElementById('activityInventoryForm');
  if (!item || !form) return;
  Object.entries({ID:item.ID, Versi:item.Versi || '', InventarisID:item.InventarisID, JumlahAwal:item.JumlahAwal, JumlahDipakai:item.JumlahDipakai, JumlahSetelah:item.JumlahSetelah || '', KondisiAwal:item.KondisiAwal || 'Baik', KondisiSetelah:item.KondisiSetelah || 'Baik', JumlahRusakKembali:item.JumlahRusakKembali || 0, StatusPemakaian:item.StatusPemakaian || 'Dipakai', Keterangan:item.Keterangan || ''}).forEach(([name,value]) => { const el=form.elements[name]; if(el) el.value=value; });
  form.scrollIntoView({behavior:'smooth', block:'center'});
}

async function saveKegiatanInventory(event, kegiatanId) {
  event.preventDefault();
  const form = event.target;
  const payload = Object.fromEntries(new FormData(form).entries());
  const button = form.querySelector('button[type="submit"]');
  if (button) button.disabled = true;

  try {
    await serverCall('saveKegiatanInventaris', state.token, kegiatanId, payload);
    markModuleDirty('kegiatanInventaris');
    markDashboardDirty();
    showToast('Pemakaian inventaris berhasil dicatat.', 'success');
    await openKegiatanDetail(kegiatanId);
    if (state.view === 'dashboard') renderView();
  } catch (error) {
    showToast(error.message || 'Gagal mencatat pemakaian inventaris.', 'error');
    if (button) button.disabled = false;
  }
}

async function deleteKegiatanInventory(recordId, kegiatanId) {
  if (!await customConfirm('Hapus riwayat pemakaian ini beserta foto kondisi terkait?', 'Hapus Riwayat Inventaris', 'Hapus')) return;
  try {
    await serverCall('deleteKegiatanInventaris', state.token, recordId);
    markModuleDirty('kegiatanInventaris');
    markDashboardDirty();
    showToast('Riwayat pemakaian inventaris dihapus.', 'success');
    await openKegiatanDetail(kegiatanId);
    if (state.view === 'dashboard') renderView();
  } catch (error) {
    showToast(error.message || 'Gagal menghapus riwayat inventaris.', 'error');
  }
}

async function handleKegiatanInventoryPhoto(input, kegiatanId, recordId, phase) {
  const file = input && input.files && input.files[0];
  if (!file) return;
  try {
    const payload = await prepareKegiatanDocumentationImage(file);
    await serverCall('uploadKegiatanInventarisPhoto', state.token, kegiatanId, recordId, phase, payload);
    markModuleDirty('kegiatanInventaris');
    showToast('Foto kondisi ' + (phase === 'Awal' ? 'awal' : 'setelah kegiatan') + ' berhasil disimpan.', 'success');
    await openKegiatanDetail(kegiatanId);
  } catch (error) {
    showToast(error.message || 'Gagal menyimpan foto kondisi inventaris.', 'error');
    input.value = '';
  }
}

async function saveKegiatanDocumentationMetadata(documentationId) {
  const caption = document.getElementById('documentationCaption-' + documentationId);
  const order = document.getElementById('documentationOrder-' + documentationId);
  try {
    await serverCall('updateKegiatanDokumentasiMetadata', state.token, documentationId, {
      Keterangan: caption ? caption.value : '',
      Urutan: order ? order.value : 1
    });
    showToast('Keterangan dokumentasi disimpan.', 'success');
  } catch (error) {
    showToast(error.message || 'Gagal menyimpan keterangan dokumentasi.', 'error');
  }
}

function handleKegiatanDocumentationSelection(input) {
  const text = document.getElementById('kegiatanDocumentationSelectionText');
  const button = document.getElementById('kegiatanDocumentationUploadButton');
  const files = Array.from(input && input.files ? input.files : []);

  if (files.length > 5) {
    input.value = '';
    if (text) text.textContent = 'Maksimal 5 foto per sekali unggah.';
    if (button) button.disabled = true;
    showToast('Pilih maksimal 5 foto per sekali unggah. Setelah selesai Anda dapat menambah foto lagi.', 'warning');
    return;
  }

  const invalid = files.find(file => !/^image\/(jpeg|png|webp)$/i.test(String(file.type || '')));
  if (invalid) {
    input.value = '';
    if (text) text.textContent = 'Format foto tidak didukung.';
    if (button) button.disabled = true;
    showToast('Dokumentasi hanya mendukung JPG, PNG, atau WEBP.', 'error');
    return;
  }

  if (text) {
    text.textContent = files.length
      ? `${files.length} foto siap diunggah.`
      : 'Belum ada foto dipilih.';
  }
  if (button) button.disabled = files.length === 0;
}

async function uploadKegiatanDocumentation(kegiatanId) {
  const input = document.getElementById('kegiatanDocumentationInput');
  const button = document.getElementById('kegiatanDocumentationUploadButton');
  const progress = document.getElementById('kegiatanDocumentationProgress');
  const files = Array.from(input && input.files ? input.files : []);

  if (!files.length) {
    showToast('Pilih foto dokumentasi terlebih dahulu.', 'warning');
    return;
  }

  if (files.length > 5) {
    showToast('Maksimal 5 foto per sekali unggah.', 'warning');
    return;
  }

  if (button) button.disabled = true;

  try {
    for (let index = 0; index < files.length; index++) {
      const file = files[index];
      if (progress) {
        progress.innerHTML = `
          <div class="documentation-progress-line">
            <span>Menyiapkan foto ${index + 1} dari ${files.length}</span>
            <strong>${escapeHtml(file.name)}</strong>
          </div>
          <div class="progress"><div class="progress-bar" style="width:${Math.round((index / files.length) * 100)}%"></div></div>
        `;
      }

      const payload = await prepareKegiatanDocumentationImage(file);

      if (progress) {
        progress.innerHTML = `
          <div class="documentation-progress-line">
            <span>Mengunggah foto ${index + 1} dari ${files.length}</span>
            <strong>${escapeHtml(file.name)}</strong>
          </div>
          <div class="progress"><div class="progress-bar" style="width:${Math.round((index / files.length) * 100)}%"></div></div>
        `;
      }

      await serverCall('uploadKegiatanDokumentasi', state.token, kegiatanId, payload);
    }

    if (progress) {
      progress.innerHTML = `
        <div class="documentation-progress-line success">
          <span>Selesai</span>
          <strong>${files.length} foto berhasil diunggah.</strong>
        </div>
        <div class="progress"><div class="progress-bar" style="width:100%"></div></div>
      `;
    }

    showToast(`${files.length} foto dokumentasi berhasil disimpan.`, 'success');
    await openKegiatanDetail(kegiatanId);
  } catch (error) {
    showToast(error.message || 'Gagal mengunggah dokumentasi.', 'error');
    if (progress) progress.innerHTML = errorBox(error.message || 'Gagal mengunggah dokumentasi.');
    if (button) button.disabled = false;
  }
}

async function prepareKegiatanDocumentationImageCore(file) {
  if (!file || !/^image\/(jpeg|png|webp)$/i.test(String(file.type || ''))) {
    throw new Error('Format foto tidak didukung. Gunakan JPG, PNG, atau WEBP.');
  }

  if (file.size > 15 * 1024 * 1024) {
    throw new Error(`Foto ${file.name} terlalu besar. Maksimum file asli 15 MB.`);
  }

  const image = await loadImageFileForDocumentation(file);
  const maxDimension = 1600;
  const width = Number(image.naturalWidth || image.width || 0);
  const height = Number(image.naturalHeight || image.height || 0);

  if (!width || !height) {
    throw new Error(`Foto ${file.name} tidak dapat dibaca.`);
  }

  const scale = Math.min(1, maxDimension / Math.max(width, height));
  const targetWidth = Math.max(1, Math.round(width * scale));
  const targetHeight = Math.max(1, Math.round(height * scale));
  const canvas = document.createElement('canvas');
  canvas.width = targetWidth;
  canvas.height = targetHeight;

  const context = canvas.getContext('2d', { alpha: false });
  if (!context) throw new Error('Browser tidak dapat memproses foto.');
  context.fillStyle = '#ffffff';
  context.fillRect(0, 0, targetWidth, targetHeight);
  context.drawImage(image, 0, 0, targetWidth, targetHeight);

  let blob = await canvasToDocumentationBlob(canvas, .82);
  if (blob.size > 1800 * 1024) blob = await canvasToDocumentationBlob(canvas, .70);
  if (blob.size > 2300 * 1024) {
    throw new Error(`Foto ${file.name} masih terlalu besar setelah kompresi. Coba foto dengan resolusi lebih kecil.`);
  }

  const base64 = await blobToDocumentationBase64(blob);
  const stem = String(file.name || 'dokumentasi')
    .replace(/\.[^.]+$/, '')
    .replace(/[^a-zA-Z0-9 _.-]+/g, '')
    .trim()
    .slice(0, 70) || 'dokumentasi';

  return {
    name: stem + '.jpg',
    mimeType: 'image/jpeg',
    base64
  };
}

function loadImageFileForDocumentation(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const image = new Image();
    image.onload = () => {
      URL.revokeObjectURL(url);
      resolve(image);
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error(`Foto ${file.name} tidak dapat dibaca oleh browser.`));
    };
    image.src = url;
  });
}

function canvasToDocumentationBlob(canvas, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob(blob => {
      if (blob) resolve(blob);
      else reject(new Error('Gagal mengompresi foto dokumentasi.'));
    }, 'image/jpeg', quality);
  });
}

function blobToDocumentationBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const value = String(reader.result || '');
      const comma = value.indexOf(',');
      resolve(comma >= 0 ? value.slice(comma + 1) : value);
    };
    reader.onerror = () => reject(new Error('Gagal membaca foto dokumentasi.'));
    reader.readAsDataURL(blob);
  });
}

async function loadKegiatanDocumentationPreviews(items) {
  const queue = [...(items || [])];
  const workerCount = Math.min(3, queue.length);

  async function worker() {
    while (queue.length) {
      const item = queue.shift();
      if (!item) continue;
      const image = document.getElementById('documentationImage-' + item.ID);
      const loading = document.getElementById('documentationLoading-' + item.ID);
      if (!image) continue;

      try {
        const result = await serverCall('getKegiatanDokumentasiPreview', state.token, item.ID);
        if (!document.getElementById('documentationImage-' + item.ID)) continue;
        image.src = result.dataUrl;
        image.classList.add('loaded');
        if (loading) loading.remove();
      } catch (_) {
        if (loading) {
          loading.innerHTML = `
            <span class="material-symbols-rounded">broken_image</span>
            <span>Foto gagal dimuat</span>
          `;
        }
      }
    }
  }

  await Promise.all(Array.from({ length: workerCount }, () => worker()));
}

async function deleteKegiatanDocumentation(documentationId, kegiatanId) {
  if (!await customConfirm('Hapus foto dokumentasi ini? File di Google Drive juga akan dipindahkan ke Sampah.', 'Hapus Dokumentasi', 'Hapus')) {
    return;
  }

  try {
    await serverCall('deleteKegiatanDokumentasi', state.token, documentationId);
    showToast('Dokumentasi berhasil dihapus.', 'success');
    await openKegiatanDetail(kegiatanId);
  } catch (error) {
    showToast(error.message || 'Gagal menghapus dokumentasi.', 'error');
  }
}

function enhanceResponsiveTables() {
  document.querySelectorAll('.responsive-table-wrap .data-table').forEach(table => {
    const headers = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent.trim());
    if (!headers.length) return;

    table.querySelectorAll('tbody tr').forEach(row => {
      const cells = Array.from(row.children).filter(cell => cell.tagName === 'TD');
      cells.forEach((cell, index) => {
        if (cell.hasAttribute('colspan')) return;
        if (!cell.dataset.label) cell.dataset.label = headers[index] || 'Data';
      });
    });
  });
}

function normalizeDisplayModePreference(value) {
  const normalized = String(value || '').trim().toLowerCase();
  return ['auto','mobile','desktop'].includes(normalized) ? normalized : 'auto';
}

function getDisplayModePreference() {
  const stored = state && state.ui
    ? state.ui.displayMode
    : browserStorage.getItem('saka_display_mode');
  return normalizeDisplayModePreference(stored);
}

function displayModeShortLabel(mode = getDisplayModePreference()) {
  const labels = { auto: 'AUTO', mobile: 'MOBILE', desktop: 'DESKTOP' };
  return labels[normalizeDisplayModePreference(mode)] || 'AUTO';
}

function getResponsiveMetrics() {
  const visualWidth = window.visualViewport
    ? Number(window.visualViewport.width || 0)
    : 0;

  const clientWidth = Number(
    document.documentElement &&
    document.documentElement.clientWidth || 0
  );

  const innerWidth = Number(window.innerWidth || 0);

  // Prioritaskan lebar area yang benar-benar terlihat pengguna.
  // Jangan gunakan Math.max() karena pada HP tertentu nilai innerWidth
  // dapat menyerupai viewport desktop (900-1000 px).
  const viewportWidth =
    visualWidth ||
    clientWidth ||
    innerWidth ||
    1024;

  const uaDataMobile = navigator.userAgentData
    ? navigator.userAgentData.mobile
    : null;

  const ua = String(navigator.userAgent || '');
  const touchPoints = Number(navigator.maxTouchPoints || 0);
  const coarsePointer = Boolean(
    window.matchMedia &&
    window.matchMedia('(pointer: coarse)').matches
  );

  const screenW = Number(window.screen && window.screen.width || 0);
  const screenH = Number(window.screen && window.screen.height || 0);
  const shortScreen = screenW && screenH
    ? Math.min(screenW, screenH)
    : 0;

  const mobileUa = /Android|webOS|iPhone|iPod|BlackBerry|IEMobile|Opera Mini|Mobile/i.test(ua);

  return {
    viewportWidth,
    uaDataMobile,
    ua,
    touchPoints,
    coarsePointer,
    screenW,
    screenH,
    shortScreen,
    mobileUa
  };
}

function isProbablePhysicalPhone(metrics = getResponsiveMetrics()) {
  const physicalPhoneByScreen = (
    metrics.shortScreen > 0 &&
    metrics.shortScreen <= 600 &&
    (metrics.touchPoints > 0 || metrics.coarsePointer)
  );

  return Boolean(
    physicalPhoneByScreen ||
    metrics.uaDataMobile === true ||
    metrics.mobileUa
  );
}

function migrateResponsivePreference() {
  const migrationKey = 'saka_display_mode_migrated_3_1_3';

  if (browserStorage.getItem(migrationKey) === '1') {
    return;
  }

  const stored = normalizeDisplayModePreference(
    browserStorage.getItem('saka_display_mode')
  );

  // Versi lama dapat menyimpan Desktop saat HP salah terdeteksi.
  // Reset satu kali ke Auto hanya pada perangkat yang sangat mungkin HP.
  if (stored === 'desktop' && isProbablePhysicalPhone()) {
    browserStorage.setItem('saka_display_mode', 'auto');

    if (state && state.ui) {
      state.ui.displayMode = 'auto';
    }
  }

  browserStorage.setItem(migrationKey, '1');
}

function detectResponsiveMode() {
  const metrics = getResponsiveMetrics();
  const viewportWidth = metrics.viewportWidth;

  // Prioritas tertinggi: karakteristik perangkat menunjukkan HP fisik.
  // Ini tetap bekerja ketika Chrome Android "Desktop site" melaporkan
  // viewport sekitar 900-1000 px.
  if (isProbablePhysicalPhone(metrics)) {
    if (viewportWidth <= 960 || metrics.shortScreen <= 600) {
      return 'mobile';
    }

    return 'tablet';
  }

  if (viewportWidth <= 720) {
    return 'mobile';
  }

  // Perangkat touch/coarse pointer berukuran menengah diperlakukan tablet.
  if (
    viewportWidth <= 1100 ||
    (metrics.coarsePointer && metrics.touchPoints > 0)
  ) {
    return 'tablet';
  }

  return 'desktop';
}

function getResponsiveMode() {
  const preference = getDisplayModePreference();
  if (preference === 'mobile' || preference === 'desktop') return preference;
  return detectResponsiveMode();
}

function isMobileLikeDevice() {
  return getResponsiveMode() !== 'desktop';
}

function openDisplayModeModal() {
  const preference = getDisplayModePreference();
  const detected = detectResponsiveMode();
  const active = getResponsiveMode();

  openCustomModal(
    'Mode Tampilan',
    'Pilih tampilan yang paling nyaman untuk perangkat ini.',
    `
      <div class="display-mode-panel">
        <div class="display-mode-current">
          <span class="material-symbols-rounded">devices</span>
          <div>
            <small>MODE AKTIF</small>
            <strong>${escapeHtml(active === 'desktop' ? 'Desktop' : active === 'mobile' ? 'Mobile' : 'Tablet')}</strong>
            <p>${preference === 'auto'
              ? `Otomatis terdeteksi sebagai ${escapeHtml(detected)}.`
              : `Pilihan manual tersimpan: ${escapeHtml(preference)}.`}
            </p>
          </div>
        </div>

        <div class="display-mode-options">
          ${displayModeOptionHtml(
            'auto',
            'devices',
            'Otomatis',
            'Sistem memilih Mobile, Tablet, atau Desktop sesuai perangkat.',
            preference
          )}
          ${displayModeOptionHtml(
            'mobile',
            'smartphone',
            'Mobile',
            'Tampilan HP yang nyaman, teks dan tombol dibuat lebih besar serta mudah dibaca.',
            preference
          )}
          ${displayModeOptionHtml(
            'desktop',
            'desktop_windows',
            'Desktop',
            'Paksa layout desktop penuh, termasuk saat aplikasi dibuka dari HP.',
            preference
          )}
        </div>

        <div class="display-mode-note">
          <span class="material-symbols-rounded">info</span>
          <span>Mode Desktop di layar HP menggunakan kanvas desktop minimal 1024 px. Geser horizontal atau gunakan zoom browser bila diperlukan.</span>
        </div>
      </div>
    `,
    true
  );
}

function displayModeOptionHtml(mode, icon, title, description, preference) {
  const selected = normalizeDisplayModePreference(preference) === mode;
  return `
    <button
      class="display-mode-option ${selected ? 'selected' : ''}"
      type="button"
      onclick="setDisplayMode('${mode}')"
      aria-pressed="${selected ? 'true' : 'false'}"
    >
      <span class="display-mode-option-icon material-symbols-rounded">${icon}</span>
      <span class="display-mode-option-copy">
        <strong>${escapeHtml(title)}</strong>
        <small>${escapeHtml(description)}</small>
      </span>
      <span class="display-mode-check material-symbols-rounded">${selected ? 'check_circle' : 'radio_button_unchecked'}</span>
    </button>
  `;
}

function setDisplayMode(mode) {
  const preference = normalizeDisplayModePreference(mode);
  state.ui.displayMode = preference;
  browserStorage.setItem('saka_display_mode', preference);

  const applied = applyResponsiveMode();
  closeModal();
  closeSidebar();

  const mini = document.getElementById('displayModeMini');
  if (mini) mini.textContent = displayModeShortLabel(preference);

  requestAnimationFrame(() => {
    enhanceResponsiveTables();
    scrollPrimaryTo(0, 'auto');
  });

  const label = preference === 'auto'
    ? `Otomatis (${applied})`
    : preference === 'mobile'
      ? 'Mobile'
      : 'Desktop';
  showToast(`Mode tampilan: ${label}.`, 'success');
}

function getPrimaryScrollContainer() {
  const main = document.querySelector('#appShell .main');
  const mode = getResponsiveMode();

  // Mobile dan desktop memakai .main sebagai scroll container deterministik.
  // Ini menghindari perilaku wheel yang tidak konsisten pada document/iframe Apps Script.
  if (main && (mode === 'mobile' || mode === 'desktop')) return main;

  return document.scrollingElement || document.documentElement || document.body;
}

function scrollPrimaryTo(top = 0, behavior = 'auto') {
  const value = Math.max(0, Number(top) || 0);
  const scroller = getPrimaryScrollContainer();

  if (scroller && scroller !== document.scrollingElement && scroller !== document.documentElement && scroller !== document.body) {
    scroller.scrollTo({ top: value, behavior });
    return;
  }

  window.scrollTo({ top: value, behavior });
}

function repairScrollLockState() {
  const body = document.body;
  if (!body) return;

  const backdrop = document.getElementById('modalBackdrop');
  const modalVisible = Boolean(
    backdrop &&
    backdrop.classList.contains('show') &&
    backdrop.getAttribute('aria-hidden') !== 'true'
  );
  if (!modalVisible) body.classList.remove('modal-open');

  const sidebar = document.getElementById('sidebar');
  const sidebarVisible = Boolean(sidebar && sidebar.classList.contains('open'));
  if (!sidebarVisible) body.classList.remove('mobile-nav-open');
}


/* =====================================================
   V3.1.6 DESKTOP SCROLL RESTORE
   Menghapus sisa scroll-lock mobile/modal ketika mode aktif Desktop.
   Tidak membuka background ketika modal desktop benar-benar sedang tampil.
===================================================== */
function restoreDesktopScrollState() {
  if (getResponsiveMode() !== 'desktop') return;

  const root = document.documentElement;
  const body = document.body;
  if (!root || !body) return;

  const backdrop = document.getElementById('modalBackdrop');
  const modalVisible = Boolean(
    backdrop &&
    backdrop.classList.contains('show') &&
    backdrop.getAttribute('aria-hidden') !== 'true'
  );

  // Drawer mobile tidak pernah menjadi scroll-lock yang valid di desktop.
  body.classList.remove('mobile-nav-open');

  // Hanya lepaskan modal lock bila modal memang tidak terlihat.
  if (!modalVisible) {
    body.classList.remove('modal-open');
  }

  // Bersihkan kemungkinan inline style sisa browser/state lama.
  [root, body].forEach(element => {
    element.style.removeProperty('height');
    element.style.removeProperty('max-height');
    element.style.removeProperty('overflow');
    element.style.removeProperty('overflow-y');
    element.style.removeProperty('touch-action');
  });

  const app = document.getElementById('appShell');
  const main = app && app.querySelector('.main');

  [app, main].filter(Boolean).forEach(element => {
    element.style.removeProperty('height');
    element.style.removeProperty('max-height');
    element.style.removeProperty('overflow');
    element.style.removeProperty('overflow-y');
    element.style.removeProperty('touch-action');
  });
}


/* =====================================================
   V3.1.7 DESKTOP MOUSE WHEEL FALLBACK
   Desktop menggunakan .main sebagai scroll container utama.
   Fallback ini memastikan roda mouse/trackpad tetap menggerakkan konten
   pada Web App Apps Script, tanpa mengganggu elemen yang punya scroll sendiri.
===================================================== */
function canScrollElementVertically(element, deltaY) {
  if (!element || !deltaY) return false;

  const maxScrollTop = Math.max(0, element.scrollHeight - element.clientHeight);
  if (maxScrollTop <= 1) return false;

  if (deltaY < 0) return element.scrollTop > 0;
  if (deltaY > 0) return element.scrollTop < maxScrollTop - 1;
  return false;
}

function findNestedVerticalScroller(target, boundary, deltaY) {
  let element = target instanceof Element ? target : null;

  while (element && element !== boundary && element !== document.body) {
    const style = window.getComputedStyle(element);
    const overflowY = style.overflowY;

    if (
      (overflowY === 'auto' || overflowY === 'scroll') &&
      canScrollElementVertically(element, deltaY)
    ) {
      return element;
    }

    element = element.parentElement;
  }

  return null;
}

function setupDesktopWheelScroll() {
  if (window.__sakaDesktopWheelBound) return;
  window.__sakaDesktopWheelBound = true;

  document.addEventListener('wheel', event => {
    if (getResponsiveMode() !== 'desktop') return;
    if (event.defaultPrevented || event.ctrlKey) return;

    const body = document.body;
    if (!body || body.classList.contains('modal-open')) return;

    const main = document.querySelector('#appShell .main');
    if (!main || !main.contains(event.target)) return;

    const deltaY = Number(event.deltaY || 0);
    if (!deltaY) return;

    // Biarkan textarea/pre/modal atau elemen lain yang memang masih bisa
    // scroll vertikal menangani wheel-nya sendiri.
    if (findNestedVerticalScroller(event.target, main, deltaY)) return;

    if (!canScrollElementVertically(main, deltaY)) return;

    event.preventDefault();
    main.scrollTop += deltaY;
  }, { passive: false });
}

function applyResponsiveMode() {
  const mode = getResponsiveMode();
  const root = document.documentElement;
  const body = document.body;

  ['saka-mobile','saka-tablet','saka-desktop'].forEach(className => {
    root.classList.remove(className);
    if (body) body.classList.remove(className);
  });

  const preference = getDisplayModePreference();

  root.classList.add('saka-' + mode);
  root.dataset.sakaMode = mode;
  root.dataset.sakaDisplayPreference = preference;
  if (body) {
    body.classList.add('saka-' + mode);
    body.dataset.sakaMode = mode;
    body.dataset.sakaDisplayPreference = preference;
  }

  const app = document.getElementById('appShell');
  if (app && mode !== 'desktop') {
    // Preferensi collapse adalah fitur desktop dan tidak boleh mengganggu mobile/tablet.
    app.classList.remove('sidebar-collapsed');
  } else if (app && mode === 'desktop' && state.ui.sidebarCollapsed) {
    app.classList.add('sidebar-collapsed');
  }

  if (mode === 'desktop') {
    closeSidebar();
    restoreDesktopScrollState();
  } else {
    repairScrollLockState();
  }

  return mode;
}

function handleResponsiveShell() {
  applyResponsiveMode();
}

function setupGlobalUi() {
  applyResponsiveMode();
  repairScrollLockState();
  setupDesktopWheelScroll();

  const backdrop = document.getElementById('modalBackdrop');

  if (backdrop && !backdrop.dataset.bound) {
    backdrop.dataset.bound = 'true';
    backdrop.addEventListener('click', event => {
      if (event.target === backdrop) closeModal();
    });
  }

  document.addEventListener('keydown', event => {
    if (event.key === 'Escape') {
      const visibleBackdrop = document.querySelector('.modal-backdrop.show');
      if (visibleBackdrop) closeModal();
      if (document.body.classList.contains('mobile-nav-open')) closeSidebar();
    }
  });

  if (!window.__sakaResponsiveBound) {
    window.__sakaResponsiveBound = true;
    window.addEventListener('resize', handleResponsiveShell, { passive: true });
    window.addEventListener('orientationchange', () => {
      setTimeout(handleResponsiveShell, 80);
    }, { passive: true });

    if (window.visualViewport) {
      window.visualViewport.addEventListener('resize', handleResponsiveShell, { passive: true });
    }

    window.addEventListener('pageshow', () => {
      applyResponsiveMode();
      restoreDesktopScrollState();
    }, { passive: true });
  }
}


function ensureXlsxLibrary(...args) {
  return trackLocalActivity('Memuat pembaca Excel', () => ensureXlsxLibraryCore(...args));
}

function parseDatabaseExcel(...args) {
  return trackLocalActivity('Membaca dan memvalidasi Excel', () => parseDatabaseExcelCore(...args));
}

function prepareKegiatanDocumentationImage(...args) {
  return trackLocalActivity('Menyiapkan dan mengompresi foto', () => prepareKegiatanDocumentationImageCore(...args));
}

function prepareIzinEvidenceFile(...args) {
  return trackLocalActivity('Menyiapkan bukti izin', () => prepareIzinEvidenceFileCore(...args));
}
/* Ringkasan modul dan transaksi inventaris langsung. */
const INVENTORY_CATEGORIES = ['Perkemahan','Perlengkapan Upacara','Penerbangan','Elektronik','Administrasi','Kesehatan / P3K','Olahraga','Perkakas','Konsumsi / Dapur','Lainnya'];

function summaryCardsHtml(cards) {
  return `<div class="module-summary">${cards.map(([label,value]) => `<article class="module-summary-card"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></article>`).join('')}</div>`;
}

function localDateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;
}

function cashPeriodRows(now = new Date()) {
  const months = Number(state.cashMonths === undefined ? 1 : state.cashMonths);
  const today = localDateKey(now);
  const start = months ? localDateKey(new Date(now.getFullYear(), now.getMonth()-months+1, 1)) : '';
  const throughToday = (state.data.kas || []).filter(row => String(row.Tanggal || '').slice(0,10) <= today);
  return {start, today, rows:throughToday.filter(row => String(row.Tanggal || '').slice(0,10) >= start), all:throughToday};
}

function cashSummaryHtml(period) {
  const sum = (rows, kind) => rows.filter(row => row.Jenis === kind).reduce((total,row) => total + Number(row.Nominal || 0),0);
  return summaryCardsHtml([
    ['Saldo saat ini', formatRupiah(sum(period.all,'Pemasukan') - sum(period.all,'Pengeluaran'))],
    ['Pemasukan periode', formatRupiah(sum(period.rows,'Pemasukan'))],
    ['Pengeluaran periode', formatRupiah(sum(period.rows,'Pengeluaran'))],
    ['Selisih periode', formatRupiah(sum(period.rows,'Pemasukan') - sum(period.rows,'Pengeluaran'))]
  ]) + `<p class="module-period-note">Periode: ${period.start ? formatDate(period.start) : 'Awal pencatatan'} – ${formatDate(period.today)}. Ringkasan tidak berubah oleh pencarian teks.</p>`;
}

function periodFilterHtml() {
  return `<label>Periode kas <select class="form-control" onchange="state.cashMonths=Number(this.value);renderKas()">${[[1,'Bulan ini'],[3,'3 bulan terakhir'],[6,'6 bulan terakhir'],[0,'Seluruh periode']].map(([value,label]) => `<option value="${value}" ${value === (state.cashMonths === undefined ? 1 : state.cashMonths) ? 'selected' : ''}>${label}</option>`).join('')}</select></label>`;
}

function inventorySummaryHtml() {
  const data = state.data.inventaris || [];
  return summaryCardsHtml([
    ['Jenis barang',data.length],
    ['Unit fisik',data.reduce((n,row) => n + Number(row.Jumlah || 0),0)],
    ['Sedang dipakai',data.reduce((n,row) => n + Number(row.SedangDipakai || 0),0)],
    ['Tersedia',data.reduce((n,row) => n + Number(row.Tersedia || 0),0)]
  ]);
}

function letterFormatFromNumber(number) {
  return ['PANPEL-SADIRGA','SADIRGA-SDA','SK/MUSAKA'].find(format => String(number || '').includes('/'+format+'/')) || 'SADIRGA-SDA';
}

function syncLetterForm() {
  const form = document.getElementById('dataForm');
  if (!form || !form.elements.FormatSurat) return;
  const fields = form.elements;
  const outgoing = fields.Jenis.value === 'Surat Keluar';
  const auto = outgoing && fields.Penomoran.value === 'Otomatis';
  ['Penomoran','FormatSurat'].forEach(name => {
    fields[name].closest('.form-group').hidden = !outgoing;
    fields[name].disabled = !outgoing;
    fields[name].required = outgoing;
  });
  fields.FormatSurat.disabled = !auto;
  fields.FormatSurat.required = auto;
  fields.NomorSurat.readOnly = auto;
  fields.NomorSurat.required = !auto;
  if (auto) {
    if (fields.NomorSurat.value) fields.NomorSurat.dataset.manual = fields.NomorSurat.value;
    fields.NomorSurat.value = '';
    const parts = fields.Tanggal.value.split('-');
    const roman = ['I','II','III','IV','V','VI','VII','VIII','IX','X','XI','XII'][Number(parts[1])-1];
    fields.NomorSurat.placeholder = `Otomatis: NN/${fields.FormatSurat.value}/${roman || 'BULAN'}/${parts[0] || 'TAHUN'}`;
  } else {
    if (!fields.NomorSurat.value) fields.NomorSurat.value = fields.NomorSurat.dataset.manual || '';
    fields.NomorSurat.placeholder = outgoing ? 'Masukkan nomor surat' : 'Nomor asli dari pengirim';
  }
}

async function openInventoryQuick(inventoryId, recordId = '') {
  const token = state.token;
  openCustomModal('Pemakaian & Pengembalian', 'Memuat stok terkini...', loadingHtml(160));
  const generation = modalLoadGeneration;
  try {
    const options = await serverCall('getInventoryQuickOptions', token);
    if (!await ensureModulesLoaded(['inventaris','kegiatanInventaris'], true)) return;
    if (generation !== modalLoadGeneration || token !== state.token) return;
    state.inventoryQuickOptions = options;
    const inventory = (state.data.inventaris || []).find(row => String(row.ID) === String(inventoryId));
    if (!inventory) throw new Error('Barang tidak ditemukan.');
    const records = (state.data.kegiatanInventaris || []).filter(row => String(row.InventarisID) === String(inventoryId));
    const record = recordId ? records.find(row => String(row.ID) === String(recordId)) : null;
    if (recordId && (!record || record.StatusPemakaian !== 'Dipakai')) throw new Error('Catatan sudah berubah. Buka kembali menu inventaris.');
    const activities = Object.fromEntries(options.kegiatan.map(row => [String(row.ID),row]));
    const permitted = canPermission('inventaris',record ? 'edit' : 'create');
    const fields = record ? [
      field('JumlahSetelah','Jumlah kembali (termasuk yang rusak)','number',true),
      field('JumlahRusakKembali','Dari jumlah kembali, berapa unit rusak?','number',true),
      selectField('KondisiSetelah','Kondisi kembali',['Baik','Rusak Ringan','Rusak Berat','Hilang'],true),
      textareaField('Keterangan','Catatan kehilangan / kerusakan',true)
    ] : [
      customSelectField('KegiatanID','Dipakai untuk kegiatan',options.kegiatan.filter(row => ['Rencana','Berjalan'].includes(row.Status)).map(row => ({value:row.ID,label:row.NamaKegiatan+' — '+formatDate(row.Tanggal)})),true),
      field('JumlahDipakai','Jumlah dipakai','number',true),
      textareaField('Keterangan','Catatan pemakaian',true)
    ];
    const values = record ? {JumlahSetelah:record.JumlahDipakai,JumlahRusakKembali:0,KondisiSetelah:'Baik',Keterangan:record.Keterangan || ''} : {JumlahDipakai:1};
    const history = records.map(row => `<tr><td>${escapeHtml(activities[String(row.KegiatanID)]?.NamaKegiatan || row.KegiatanID)}</td><td>${escapeHtml(row.StatusPemakaian || 'Legacy')}</td><td>${Number(row.JumlahDipakai || 0)}</td><td>${row.JumlahSetelah === '' ? '—' : escapeHtml(row.JumlahSetelah ?? '—')}</td><td>${Number(row.JumlahRusakKembali || 0)}</td><td>${row.StatusPemakaian === 'Dipakai' && canPermission('inventaris','edit') ? `<button class="btn btn-light btn-compact" onclick="openInventoryQuick('${escapeJs(inventoryId)}','${escapeJs(row.ID)}')">Kembalikan</button>` : '—'}</td></tr>`).join('');
    openCustomModal(inventory.NamaBarang, `${inventory.KodeBarang || ''} · ${Number(inventory.Tersedia || 0)} tersedia · ${Number(inventory.SedangDipakai || 0)} dipakai`, `
      ${permitted ? `<h3>${record ? 'Pengembalian: '+escapeHtml(activities[String(record.KegiatanID)]?.NamaKegiatan || record.KegiatanID) : 'Catat pemakaian'}</h3>
      <p>${record ? 'Dipakai '+Number(record.JumlahDipakai)+' unit. Selisih jumlah dicatat sebagai hilang; barang rusak tidak masuk stok tersedia. Pengembalian ini menutup catatan.' : 'Pilih kegiatan Rencana / Berjalan. Stok tersedia diperiksa kembali saat menyimpan.'}</p>
      <form class="form-grid" onsubmit="submitInventoryQuick(event,'${escapeJs(inventoryId)}','${escapeJs(recordId)}')">
        ${fields.map(def => createField(def,values)).join('')}
        <div class="form-actions"><button type="button" class="btn btn-light" onclick="${record ? `openInventoryQuick('${escapeJs(inventoryId)}')` : 'closeModal()'}">${record ? 'Batal pengembalian' : 'Tutup'}</button><button type="submit" class="btn btn-primary">${record ? 'Simpan pengembalian' : 'Simpan pemakaian'}</button></div>
      </form>` : '<p>Pilih catatan aktif di bawah untuk pengembalian sesuai hak akses Anda.</p>'}
      <h3>Riwayat pemakaian</h3><div class="table-wrap"><table class="data-table"><thead><tr><th>Kegiatan</th><th>Status</th><th>Dipakai</th><th>Kembali</th><th>Rusak</th><th>Aksi</th></tr></thead><tbody>${history || emptyTableRow(6,'Belum ada pemakaian.')}</tbody></table></div>`);
  } catch (error) { showToast(error.message, 'error'); }
}

async function submitInventoryQuick(event, inventoryId, recordId) {
  event.preventDefault();
  const button = event.target.querySelector('button[type="submit"]');
  button.disabled = true;
  const data = Object.fromEntries(new FormData(event.target).entries());
  const record = (state.data.kegiatanInventaris || []).find(row => String(row.ID) === String(recordId));
  const kegiatanId = record ? record.KegiatanID : data.KegiatanID;
  const payload = Object.assign({},data,{InventarisID:inventoryId,StatusPemakaian:record ? 'Selesai' : 'Dipakai'});
  if (record) Object.assign(payload,{ID:record.ID,Versi:record.Versi,JumlahDipakai:record.JumlahDipakai,KondisiAwal:record.KondisiAwal});
  try {
    await serverCall('saveKegiatanInventaris',state.token,kegiatanId,payload);
    markModuleDirty('inventaris');
    markModuleDirty('kegiatanInventaris');
    markDashboardDirty();
    showToast(record ? 'Pengembalian tersimpan. Stok diperbarui.' : 'Pemakaian tersimpan. Stok direservasi.','success');
    await openInventoryQuick(inventoryId);
    if (state.view === 'inventaris') renderInventaris();
  } catch (error) { showToast(error.message,'error'); button.disabled = false; }
}


/* Notification & Action Center */
function scheduleNotificationRefresh() {
  clearTimeout(notificationRefreshTimer);
  notificationRefreshTimer = setTimeout(() => { loadActionNotifications().catch(() => {}); },700);
}

async function loadActionNotifications() {
  const token = state.token;
  if (!token || !document.getElementById('notificationBell')) return;
  const version = ++notificationFetchVersion;
  try {
    const result = await serverCall('getActionNotifications',token);
    if (token !== state.token || version !== notificationFetchVersion) return;
    state.actionNotifications = result;
    const badge = document.getElementById('notificationBadge');
    const bell = document.getElementById('notificationBell');
    if (badge) { badge.hidden = !result.unread; badge.textContent = result.unread>99?'99+':String(result.unread); }
    if (bell) { bell.title = 'Notification & Action Center'; bell.setAttribute('aria-label',`${result.unread} notifikasi belum dibaca. Buka Notification & Action Center`); }
    renderActionCenterBody();
    return result;
  } catch (error) {
    if (token === state.token && version === notificationFetchVersion) {
      const bell = document.getElementById('notificationBell');
      if (bell) bell.title = 'Notifikasi gagal dimuat. Klik untuk mencoba kembali.';
      const host = document.getElementById('actionCenterContent');
      if (host) host.innerHTML = `<p role="alert">${escapeHtml(error.message || 'Notifikasi gagal dimuat.')}</p><button class="btn btn-primary" onclick="openActionCenter()">Coba lagi</button>`;
    }
    throw error;
  }
}

async function openActionCenter() {
  state.actionCenterFilter = state.actionCenterFilter || 'all';
  openCustomModal('Notification & Action Center','Daftar tindak lanjut berdasarkan kondisi data saat ini.', '<div id="actionCenterContent">'+loadingHtml(160)+'</div>');
  try { await loadActionNotifications(); } catch (_) {}
}

function setActionCenterFilter(value) {
  state.actionCenterFilter = value;
  renderActionCenterBody();
}

function renderActionCenterBody() {
  const host = document.getElementById('actionCenterContent');
  const data = state.actionNotifications;
  if (!host || !data) return;
  const filter = state.actionCenterFilter || 'all';
  const visible = data.items.filter(item => filter==='all' || filter==='unread'&&!item.read || filter==='high'&&item.priority==='high' || filter===item.module);
  const labels = {kegiatan:'Kegiatan',inventaris:'Inventaris',surat:'Surat',absensi:'Izin / Absensi',kas:'Kas'};
  const modules = [...new Set(data.items.map(item=>item.module))];
  const choices = [['all','Semua'],['unread','Belum dibaca'],['high','Prioritas tinggi'],...modules.map(module=>[module,labels[module]])];
  host.innerHTML = `
    <div class="action-center-summary"><strong>${data.unread} belum dibaca</strong><span>${data.total} tindak lanjut aktif</span><button type="button" class="btn btn-light btn-compact" onclick="openActionCenter()">Muat ulang</button></div>
    <div class="action-center-toolbar"><label>Filter <select class="form-control" onchange="setActionCenterFilter(this.value)">${choices.map(([value,label])=>`<option value="${value}" ${filter===value?'selected':''}>${label}</option>`).join('')}</select></label><button type="button" class="btn btn-light btn-compact" onclick="readActionNotifications()" ${data.unread?'':'disabled'}>Tandai semua dibaca</button></div>
    <p class="module-period-note">Dibaca tidak berarti selesai. Item hilang dari daftar setelah kondisi sumber diselesaikan dan data dimuat ulang.${data.total>200?' Menampilkan 200 item teratas; jumlah belum dibaca berlaku untuk item yang ditampilkan.':''}</p>
    <div class="action-center-list">${visible.length?visible.map(item=>`<article class="action-center-item ${item.read?'is-read':'is-unread'}"><div class="action-center-item-heading"><span class="action-priority ${item.priority==='high'?'is-high':''}">${item.priority==='high'?'Prioritas tinggi':labels[item.module]}</span><span>${item.read?'Dibaca':'Baru'}</span></div><h3>${escapeHtml(item.title)}</h3><p>${escapeHtml(item.detail)}</p><div class="row-actions"><button type="button" class="btn btn-primary btn-compact" onclick="performNotificationAction('${escapeJs(item.id)}')">${escapeHtml(item.label)}</button>${item.read?'':`<button type="button" class="btn btn-light btn-compact" onclick="readActionNotifications('${escapeJs(item.id)}')">Tandai dibaca</button>`}</div></article>`).join(''):'<div class="action-center-empty"><span class="material-symbols-rounded">task_alt</span><h3>Tidak ada notifikasi pada filter ini</h3><p>Gunakan Semua untuk melihat tindak lanjut lainnya.</p></div>'}</div>`;
}

async function readActionNotifications(id = '') {
  const data = state.actionNotifications;
  if (!data) return;
  const ids = data.items.filter(item => !item.read && (!id || item.id===id)).map(item=>item.id);
  if (!ids.length) return;
  const host = document.getElementById('actionCenterContent');
  if (host) host.querySelectorAll('button').forEach(button=>button.disabled=true);
  try {
    await serverCall('markActionNotificationsRead',state.token,ids);
    await loadActionNotifications();
  } catch (error) { showToast(error.message,'error'); renderActionCenterBody(); }
}

async function performNotificationAction(id) {
  const item = (state.actionNotifications?.items || []).find(row=>row.id===id);
  if (!item || !canPermission(item.module,'view')) return;
  closeModal();
  try {
    if (item.action==='inventory-return') {
      await openInventoryQuick(item.targetId,canPermission('inventaris','edit')?item.recordId:'');
    } else if (item.action==='attendance') {
      await openNotificationPermission(item.recordId);
    } else if (item.action==='edit' && canPermission(item.module,'edit')) {
      if (!await ensureModulesLoaded([item.module],true)) return;
      await openForm(item.module,item.targetId);
    } else {
      switchView(item.module);
    }
  } catch (error) { showToast(error.message,'error'); }
}


async function openNotificationPermission(id) {
  const token = state.token;
  openCustomModal('Verifikasi Pengajuan Izin','Memuat pengajuan...',loadingHtml(160));
  const generation = modalLoadGeneration;
  try {
    const row = await serverCall('getActionPermissionDetail',token,id);
    if (generation !== modalLoadGeneration || token !== state.token) return;
    const canVerify = canPermission('absensi','edit') && ['Menunggu Verifikasi','Terkirim'].includes(row.Status);
    openCustomModal('Verifikasi Pengajuan Izin',row.NamaAnggota+' — '+row.KegiatanNama,`
      <p><strong>${escapeHtml(row.JenisPengajuan || 'Izin')}</strong> · ${escapeHtml(row.Status)}</p>
      <p>${escapeHtml(row.Alasan || 'Tanpa alasan')}</p>
      ${row.BuktiNamaFile?`<button class="btn btn-light" onclick="viewIzinEvidence('${escapeJs(id)}')">Lihat bukti pengajuan</button>`:''}
      ${canVerify?`<form class="form-grid" onsubmit="submitNotificationPermission(event,'${escapeJs(id)}')">
        ${createField(selectField('Keputusan','Keputusan',['Disetujui','Ditolak'],true),{})}
        ${createField({...textareaField('Catatan','Catatan (wajib untuk penolakan)',true),help:'Persetujuan menjadi prefill Izin/Sakit pada Absensi Massal.'},{})}
        <div class="form-actions"><button type="submit" class="btn btn-primary">Simpan keputusan</button></div></form>`:'<p>Pengajuan sudah diproses atau akun Anda hanya memiliki akses baca.</p>'}
      <button type="button" class="btn btn-light" onclick="openActionCenter()">Kembali ke notifikasi</button>`);
  } catch (error) { showToast(error.message,'error'); }
}

async function submitNotificationPermission(event,id) {
  event.preventDefault();
  const form = event.target;
  const button = form.querySelector('button[type="submit"]');
  const data = Object.fromEntries(new FormData(form).entries());
  if (data.Keputusan==='Ditolak' && !data.Catatan.trim()) { showToast('Alasan penolakan wajib diisi.','warning'); return; }
  button.disabled = true;
  try {
    await serverCall('verifyActionPermission',state.token,id,data.Keputusan,data.Catatan);
    markModuleDirty('absensi');
    showToast('Keputusan verifikasi tersimpan.','success');
    await openActionCenter();
  } catch (error) { showToast(error.message,'error'); button.disabled=false; }
}
