function validateGasUrl(gasUrl) {
  let parsed;
  try { parsed = new URL(gasUrl); } catch (_) {
    return { code: 'GAS_URL_INVALID', message: 'GAS_API_URL tidak valid.' };
  }
  if (parsed.protocol !== 'https:' || parsed.hostname !== 'script.google.com') {
    return { code: 'GAS_URL_INVALID_HOST', message: 'GAS_API_URL harus menggunakan host script.google.com.' };
  }
  if (/\/dev\/?$/i.test(parsed.pathname)) {
    return { code: 'GAS_URL_TEST_DEPLOYMENT', message: 'GAS_API_URL memakai /dev. Gunakan URL production /exec.' };
  }
  if (!/\/macros\/s\/[^/]+\/exec\/?$/i.test(parsed.pathname)) {
    return { code: 'GAS_URL_NOT_EXEC', message: 'GAS_API_URL harus berupa Web App URL yang berakhir /exec.' };
  }
  return null;
}

function hostOf(url) {
  try { return new URL(url).hostname; } catch (_) { return ''; }
}

function classifyHtml(raw, status, finalUrl) {
  const host = hostOf(finalUrl);
  if (host === 'accounts.google.com' || /serviceLogin|accounts\.google\.com|sign in - google accounts|masuk - akun google/i.test(raw || '')) {
    return { code: 'GAS_LOGIN_REQUIRED', message: 'Web App meminta login Google. Ubah akses deployment agar dapat dipanggil anonim/publik.' };
  }
  if (/script function not found\s*:\s*dopost/i.test(raw || '')) {
    return { code: 'GAS_DOPost_MISSING', message: 'Deployment aktif belum memakai backend terbaru yang memiliki doPost().' };
  }
  if (status === 401 || status === 403 || /authorization is required|you do not have permission|access denied/i.test(raw || '')) {
    return { code: 'GAS_ACCESS_DENIED', message: 'Akses Web App Apps Script ditolak.' };
  }
  if (status === 404 || /page not found|requested file does not exist|file you have requested does not exist/i.test(raw || '')) {
    return { code: 'GAS_DEPLOYMENT_NOT_FOUND', message: 'URL deployment Apps Script tidak ditemukan/tidak aktif.' };
  }
  return { code: 'GAS_HEALTH_NOT_JSON', message: 'Health check menerima HTML. Deployment kemungkinan belum diperbarui ke backend terbaru (v3.8.1).' };
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  if (req.method !== 'GET') {
    res.setHeader('Allow', 'GET');
    return res.status(405).json({ success: false, message: 'Method not allowed.' });
  }

  const gasUrl = String(process.env.GAS_API_URL || '').trim();
  const hasSecret = Boolean(String(process.env.GAS_API_SECRET || '').trim());
  if (!gasUrl || !hasSecret) {
    return res.status(500).json({
      success: false,
      diagnosticCode: 'GATEWAY_CONFIG_MISSING',
      message: 'GAS_API_URL atau GAS_API_SECRET belum tersedia pada environment Vercel.'
    });
  }

  const issue = validateGasUrl(gasUrl);
  if (issue) return res.status(500).json({ success: false, diagnosticCode: issue.code, message: issue.message });

  try {
    const healthUrl = gasUrl + (gasUrl.includes('?') ? '&' : '?') + 'health=1';
    const upstream = await fetch(healthUrl, {
      method: 'GET',
      headers: { 'Accept': 'application/json' },
      redirect: 'follow',
      cache: 'no-store'
    });
    const raw = await upstream.text();
    const normalized = raw ? raw.replace(/^\uFEFF/, '').trim() : '';
    let payload = null;
    try { payload = normalized ? JSON.parse(normalized) : null; } catch (_) {}

    if (payload && payload.success === true) {
      return res.status(200).json({
        success: true,
        gateway: 'ok',
        appsScript: payload.data || {},
        upstreamStatus: upstream.status,
        upstreamHost: hostOf(upstream.url)
      });
    }

    if (payload) {
      return res.status(502).json({
        success: false,
        diagnosticCode: 'GAS_HEALTH_REJECTED',
        message: payload.message || 'Apps Script menolak health check.',
        upstreamStatus: upstream.status
      });
    }

    const diagnosis = classifyHtml(raw, upstream.status, upstream.url);
    return res.status(502).json({
      success: false,
      diagnosticCode: diagnosis.code,
      message: diagnosis.message,
      upstreamStatus: upstream.status,
      upstreamContentType: upstream.headers.get('content-type') || '',
      upstreamHost: hostOf(upstream.url)
    });
  } catch (error) {
    return res.status(502).json({
      success: false,
      diagnosticCode: 'GAS_HEALTH_FETCH_FAILED',
      message: error instanceof Error ? error.message : 'Health check Apps Script gagal.'
    });
  }
}
