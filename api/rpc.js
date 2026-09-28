const ALLOWED_METHODS = new Set([
  "checkSystemStructure",
  "closeIzinKegiatan",
  "deleteAbsensi",
  "deleteAnggota",
  "deleteInventaris",
  "deleteKas",
  "deleteKegiatan",
  "deleteKegiatanDokumentasi",
  "deleteKegiatanInventaris",
  "deleteKomponenPenilaian",
  "deletePengurus",
  "deletePenilaianAnggota",
  "deleteSurat",
  "deleteUser",
  "finishExcelDatabaseImport",
  "generateKegiatanReportPdf",
  "generatePenilaianBulananPdf",
  "getActionNotifications",
  "getActionPermissionDetail",
  "getAttendanceLink",
  "getAttendanceGeofenceSettings",
  "getBatchIzinVerifikasi",
  "getDashboardData",
  "getInventoryQuickOptions",
  "getIzinBuktiPreview",
  "getIzinLink",
  "getKegiatanDokumentasiPreview",
  "getKegiatanReportData",
  "getModulesData",
  "getUserMemberOptions",
  "getPenilaianBulanan",
  "getPenilaianPdfPayload",
  "getPublicCheckinData",
  "getPublicIzinData",
  "getRolePermissions",
  "getSkkChecklist",
  "importExcelDatabaseBatch",
  "login",
  "logout",
  "markActionNotificationsRead",
  "memberAttendanceCheckin",
  "publicCheckin",
  "publicSubmitIzin",
  "safeResetSystem",
  "saveAbsensi",
  "saveAbsensiBatch",
  "saveAttendanceGeofenceSettings",
  "saveAnggota",
  "saveInventaris",
  "saveKas",
  "saveKegiatan",
  "saveKegiatanInventaris",
  "saveKomponenPenilaian",
  "saveLaporanKegiatan",
  "savePengurus",
  "savePenilaianAnggota",
  "saveRolePermissions",
  "saveSkkChecklist",
  "saveSurat",
  "saveUser",
  "transitionKegiatanStatus",
  "updateKegiatanDokumentasiMetadata",
  "updateSpreadsheetStructure",
  "uploadKegiatanDokumentasi",
  "uploadKegiatanInventarisPhoto",
  "verifyActionPermission",
  "verifyIzinKegiatan"
]);

function normalizeBody(body) {
  if (body && typeof body === 'object') return body;
  if (typeof body === 'string' && body.trim()) {
    try { return JSON.parse(body); } catch (_) { return {}; }
  }
  return {};
}

function validateGasUrl(gasUrl) {
  let parsed;
  try {
    parsed = new URL(gasUrl);
  } catch (_) {
    return {
      code: 'GAS_URL_INVALID',
      message: 'GAS_API_URL tidak valid. Gunakan URL Web App Apps Script yang berakhir /exec.'
    };
  }

  if (parsed.protocol !== 'https:' || parsed.hostname !== 'script.google.com') {
    return {
      code: 'GAS_URL_INVALID_HOST',
      message: 'GAS_API_URL harus berupa URL resmi Web App Apps Script di script.google.com.'
    };
  }

  if (/\/dev\/?$/i.test(parsed.pathname)) {
    return {
      code: 'GAS_URL_TEST_DEPLOYMENT',
      message: 'GAS_API_URL masih menggunakan URL /dev (Test deployment). Gunakan URL Web App production yang berakhir /exec.'
    };
  }

  if (!/\/macros\/s\/[^/]+\/exec\/?$/i.test(parsed.pathname)) {
    return {
      code: 'GAS_URL_NOT_EXEC',
      message: 'GAS_API_URL bukan URL deployment Web App yang valid. Salin Web app URL dari Deploy → Manage deployments; URL harus berakhir /exec.'
    };
  }

  return null;
}

function upstreamHost(url) {
  try { return new URL(url).hostname; } catch (_) { return ''; }
}

function diagnoseNonJson({ raw, status, contentType, finalUrl }) {
  const html = String(raw || '').toLowerCase();
  const host = upstreamHost(finalUrl);

  if (host === 'accounts.google.com' || /serviceLogin|accounts\.google\.com|sign in - google accounts|masuk - akun google/i.test(raw || '')) {
    return {
      code: 'GAS_LOGIN_REQUIRED',
      message: 'Apps Script meminta login Google. Deploy Web App sebagai “Execute as: Me” dan beri akses anonim/publik (“Anyone” / ANYONE_ANONYMOUS), lalu gunakan URL /exec.'
    };
  }

  if (/script function not found\s*:\s*dopost/i.test(raw || '')) {
    return {
      code: 'GAS_DOPost_MISSING',
      message: 'Deployment Apps Script yang aktif belum memiliki doPost(). Update deployment Web App ke versi source terbaru, lalu coba lagi.'
    };
  }

  if (/authorization is required|you do not have permission|access denied|permission denied/i.test(raw || '') || status === 401 || status === 403) {
    return {
      code: 'GAS_ACCESS_DENIED',
      message: 'Akses ke Web App Apps Script ditolak. Pastikan deployment dapat dipanggil tanpa login dan akun pemilik sudah memberi otorisasi layanan Google yang dibutuhkan.'
    };
  }

  if (/page not found|requested file does not exist|file you have requested does not exist|bad request/i.test(raw || '') || status === 404) {
    return {
      code: 'GAS_DEPLOYMENT_NOT_FOUND',
      message: 'Deployment Apps Script tidak ditemukan atau URL sudah tidak aktif. Salin ulang Web app URL dari Manage deployments dan gunakan URL /exec.'
    };
  }

  if (String(contentType || '').toLowerCase().includes('text/html') || html.includes('<html') || html.includes('<!doctype')) {
    return {
      code: 'GAS_HTML_RESPONSE',
      message: 'Apps Script mengembalikan halaman HTML, bukan JSON. Biasanya deployment belum diperbarui, bukan Web App production, atau pengaturan akses masih meminta login Google.'
    };
  }

  return {
    code: 'GAS_NON_JSON',
    message: 'Backend Apps Script mengembalikan respons non-JSON. Periksa deployment Web App dan URL GAS_API_URL.'
  };
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, max-age=0');

  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ success: false, message: 'Method not allowed.' });
  }

  const gasUrl = String(process.env.GAS_API_URL || '').trim();
  const secret = String(process.env.GAS_API_SECRET || '').trim();

  if (!gasUrl || !secret) {
    return res.status(500).json({
      success: false,
      diagnosticCode: 'GATEWAY_CONFIG_MISSING',
      message: 'Konfigurasi gateway belum lengkap. Set GAS_API_URL dan GAS_API_SECRET di Vercel, lalu redeploy.'
    });
  }

  const urlIssue = validateGasUrl(gasUrl);
  if (urlIssue) {
    return res.status(500).json({ success: false, diagnosticCode: urlIssue.code, message: urlIssue.message });
  }

  const body = normalizeBody(req.body);
  const method = String(body.method || '').trim();
  const args = Array.isArray(body.args) ? body.args : [];

  if (!ALLOWED_METHODS.has(method)) {
    return res.status(400).json({ success: false, message: `Metode API tidak diizinkan: ${method || '-'}` });
  }

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 55000);

  try {
    const upstream = await fetch(gasUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Accept': 'application/json'
      },
      redirect: 'follow',
      cache: 'no-store',
      signal: controller.signal,
      body: JSON.stringify({ secret, method, args })
    });

    const raw = await upstream.text();
    const normalizedRaw = raw ? raw.replace(/^\uFEFF/, '').trim() : '';
    let payload;
    try {
      payload = normalizedRaw ? JSON.parse(normalizedRaw) : null;
    } catch (_) {
      const diagnosis = diagnoseNonJson({
        raw,
        status: upstream.status,
        contentType: upstream.headers.get('content-type') || '',
        finalUrl: upstream.url
      });
      return res.status(502).json({
        success: false,
        diagnosticCode: diagnosis.code,
        message: diagnosis.message,
        upstreamStatus: upstream.status,
        upstreamContentType: upstream.headers.get('content-type') || '',
        upstreamHost: upstreamHost(upstream.url)
      });
    }

    if (!upstream.ok) {
      return res.status(502).json({
        success: false,
        message: payload?.message || `Backend Apps Script gagal (${upstream.status}).`,
        upstreamStatus: upstream.status
      });
    }

    if (!payload || payload.success !== true) {
      return res.status(400).json({
        success: false,
        message: payload?.message || 'Backend menolak permintaan.'
      });
    }

    return res.status(200).json(payload);
  } catch (error) {
    const timedOut = error && error.name === 'AbortError';
    return res.status(502).json({
      success: false,
      diagnosticCode: timedOut ? 'GAS_TIMEOUT' : 'GAS_FETCH_FAILED',
      message: timedOut
        ? 'Backend Apps Script tidak merespons dalam 55 detik.'
        : (error instanceof Error ? error.message : 'Gagal menghubungi backend Apps Script.')
    });
  } finally {
    clearTimeout(timeout);
  }
}
