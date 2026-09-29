const SS = SpreadsheetApp.getActiveSpreadsheet();

const SHEET_PROFILE = 'PROFILE';
const SHEET_LINKS = 'LINKS';

const SHEET_ADMIN_USERS = 'ADMIN_USERS';
const SHEET_CLICK_LOG = 'CLICK_LOG';

// ======================================================
// KSL BIO LINK VER 1.3 - SMART CAMPAIGN
// ======================================================
const SHEET_PAGE_VIEW_LOG = 'PAGE_VIEW_LOG';


function doGet(e) {
  const start = Date.now();

  ensureAppReadyOnce_();

  const page = e && e.parameter && e.parameter.page
    ? String(e.parameter.page).toLowerCase()
    : 'home';

  console.log('DOGET page:', page, 'time:', Date.now() - start, 'ms');

  const template = HtmlService.createTemplateFromFile('Index');
  template.page = page;

  return template
    .evaluate()
    .setTitle('KSL Bio Link')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Jalankan function ini sekali untuk membuat header sheet otomatis.
 */
function setupBioLinkApp() {
  ensureAppReady_();

  return {
    success: true,
    message: 'Setup selesai. Struktur aplikasi berhasil diperiksa tanpa menghapus data existing.'
  };
}

function setupProfileSheet_() {
  // NON-DESTRUCTIVE SETUP
  // Tidak pernah clear sheet PROFILE.
  ensureProfileSheet_();

  return {
    success: true,
    message: 'PROFILE siap tanpa menghapus data existing.'
  };
}

function setupLinksSheet_() {
  let sh = SS.getSheetByName(SHEET_LINKS);
  if (!sh) sh = SS.insertSheet(SHEET_LINKS);

  sh.clear();

  const headers = [
    'ID_LINK',
    'TITLE',
    'URL',
    'ICON',
    'SORT_ORDER',
    'STATUS',
    'CREATED_AT',
    'UPDATED_AT'
  ];

  const now = formatDateTime_(new Date());

  const rows = [
    [
      'LNK-001',
      'Playlist Kas RT',
      'https://youtube.com/playlist?list=PLLyqNH15aMcDlIwhVa3A9lRCAg3_rINot&si=2U7gLGmx2RBQsRHE',
      '▶️',
      1,
      'ACTIVE',
      now,
      now
    ],
    [
      'LNK-002',
      'Instagram KSL',
      'https://instagram.com/username',
      '📸',
      2,
      'ACTIVE',
      now,
      now
    ],
    [
      'LNK-003',
      'Dukung di Trakteer',
      'https://trakteer.id/username',
      '☕',
      3,
      'ACTIVE',
      now,
      now
    ]
  ];

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);

  sh.setFrozenRows(1);
}

/**
 * Data untuk admin.
 */
function getAdminData(token) {
  requireAdmin_(token);

  const links = getLinks_()
    .sort((a, b) => Number(a.SORT_ORDER || 0) - Number(b.SORT_ORDER || 0));

  const clickMap = getClickStatsMap_();

  const viewStats = getPageViewStats_();

  const linksWithStats = links.map(link => {
    const id = String(link.ID_LINK || '');
    link.TOTAL_CLICK = clickMap[id] || 0;
    return link;
  });

  return {
    success: true,
    profile: getProfile_(),
    links: linksWithStats,
    summary: {

      totalLinks:
        links.length,

      activeLinks:
        links.filter(
          l =>
            String(
              l.STATUS || ''
            ).toUpperCase() ===
            'ACTIVE'
        ).length,

      liveLinks:
        links.filter(
          isLinkPubliclyVisible_
        ).length,

      scheduledLinks:
        links.filter(
          l =>
            getLinkScheduleStatus_(l)
            === 'SCHEDULED'
        ).length,

      featuredLinks:
        links.filter(
          l =>
            String(
              l.FEATURED || ''
            ).toUpperCase()
            === 'YES'
        ).length,

      totalClicks:
        Object.keys(clickMap)
          .reduce(
            (sum, key) =>
              sum +
              Number(
                clickMap[key] || 0
              ),
            0
          ),

      totalViews:
        viewStats.totalViews,

      viewsToday:
        viewStats.viewsToday,

      views7Days:
        viewStats.views7Days,

      views30Days:
        viewStats.views30Days,

      ctr:
        viewStats.totalViews
          ? Number(
              (
                Object.keys(clickMap)
                  .reduce(
                    (sum, key) =>
                      sum +
                      Number(
                        clickMap[key] || 0
                      ),
                    0
                  )
                /
                viewStats.totalViews
                * 100
              ).toFixed(1)
            )
          : 0
    },
    analytics: getClickAnalytics_(30)
  };
}

function getProfile_() {

  const sh =
    SS.getSheetByName(
      SHEET_PROFILE
    );

  if (!sh) return {};


  const values =
    sh
      .getDataRange()
      .getValues();

  const data = {};


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const key =
      String(
        values[i][0] || ''
      )
        .trim();

    if (!key) continue;


    /*
     * FIRST OCCURRENCE WINS.
     * Mencegah stale duplicate di bawah
     * menimpa value canonical.
     */
    if (
      Object.prototype
        .hasOwnProperty.call(
          data,
          key
        )
    ) {
      continue;
    }


    data[key] =
      normalizeSheetValue_(
        values[i][1]
      );
  }


  return data;
}

function getLinks_() {
  const sh = SS.getSheetByName(SHEET_LINKS);
  if (!sh) return [];

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h || '').trim());

  return values.slice(1)
    .filter(row => row.join('').trim() !== '')
    .map(row => {
      const obj = {};

      headers.forEach((h, i) => {
        obj[h] = normalizeSheetValue_(row[i]);
      });

      return obj;
    });
}

function saveProfile(payload, token) {
  requireAdmin_(token);

  payload = payload || {};

  const sh = SS.getSheetByName(SHEET_PROFILE);

  if (!sh) {
    throw new Error('Sheet PROFILE belum tersedia.');
  }

  /*
  * VER 1.4
  * Pastikan managed key PROFILE tidak duplicate
  * sebelum proses save.
  */
  dedupeProfileKeys_(sh);

  const allowedKeys = [
    'APP_TITLE',
    'NAME',
    'BIO',
    'AVATAR_URL',

    // EVOLUTION 1 - QR BRANDING
    'QR_LOGO_ENABLED',

    // EVOLUTION 2 - SPLIT VIEW
    'PUBLIC_LAYOUT_MODE',

    'THEME_COLOR',
    'BG_STYLE',
    'BUTTON_STYLE',
    'WHATSAPP',
    'INSTAGRAM',
    'YOUTUBE',
    'TRAKTEER'
  ];

  /*
   * NON-DESTRUCTIVE SAVE
   *
   * - Tidak clear sheet
   * - Existing key hanya update VALUE
   * - Key baru append
   * - Custom key tidak disentuh
   */

  const values = sh.getDataRange().getValues();

  const keyRowMap = {};

    for (let r = 1; r < values.length; r++) {
      const key =
        String(values[r][0] || '')
          .trim()
          .toUpperCase();

      if (key && !keyRowMap[key]) {
        keyRowMap[key] = r + 1;
      }
    }

  allowedKeys.forEach(function(key) {

    if (
      !Object.prototype.hasOwnProperty.call(
        payload,
        key
      )
    ) {
      return;
    }

    const normalizedKey =
      String(key).trim().toUpperCase();

    const value =
      payload[key] === null ||
      payload[key] === undefined
        ? ''
        : payload[key];

    if (keyRowMap[normalizedKey]) {

      sh
        .getRange(
          keyRowMap[normalizedKey],
          2
        )
        .setValue(value);

    } else {

      sh.appendRow([
        key,
        value
      ]);

      keyRowMap[normalizedKey] =
        sh.getLastRow();
    }
  });

  sh.setFrozenRows(1);

  clearPublicCache_();

  return {
    success: true,
    message: 'Profil berhasil disimpan tanpa menghapus data existing.'
  };
}

function saveLink(payload, token) {
  requireAdmin_(token);
  payload = payload || {};

  const title = String(payload.TITLE || '').trim();
  const url = String(payload.URL || '').trim();

  if (!title) {
    throw new Error('Judul link wajib diisi.');
  }

  if (!url) {
    throw new Error('URL wajib diisi.');
  }

  if (!/^https?:\/\//i.test(url)) {
    throw new Error('URL harus diawali http:// atau https://');
  }

  const sh = SS.getSheetByName(SHEET_LINKS);

  if (!sh) {
    throw new Error('Sheet LINKS belum tersedia.');
  }

  // Pastikan header VER 1.3 tersedia
  ensureLinksSheet_();

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());

  const idIndex = headers.indexOf('ID_LINK');

  const idLink = String(payload.ID_LINK || '').trim();
  const now = formatDateTime_(new Date());

  const startAt = normalizeScheduleInput_(payload.START_AT);
  const endAt = normalizeScheduleInput_(payload.END_AT);

  if (startAt && endAt) {
    const startDate = parseScheduleDate_(startAt);
    const endDate = parseScheduleDate_(endAt);

    if (startDate && endDate && endDate < startDate) {
      throw new Error('Tanggal selesai tidak boleh lebih awal dari tanggal mulai.');
    }
  }

  const rowData = {
    ID_LINK: idLink || generateId_('LNK'),
    TITLE: title,
    URL: url,
    ICON: String(payload.ICON || '🔗').trim(),
    SORT_ORDER: Number(payload.SORT_ORDER || 999),
    STATUS: String(payload.STATUS || 'ACTIVE').toUpperCase(),

    // VER 1.3
    CTA_TYPE: String(payload.CTA_TYPE || 'LINK').toUpperCase(),
    CAMPAIGN: String(payload.CAMPAIGN || '').trim(),
    FEATURED:
      String(payload.FEATURED || 'NO').toUpperCase() === 'YES'
        ? 'YES'
        : 'NO',

    START_AT: startAt,
    END_AT: endAt,

    CREATED_AT: now,
    UPDATED_AT: now
  };

  if (idLink) {
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][idIndex]) === idLink) {

        const createdAtIndex = headers.indexOf('CREATED_AT');

        if (createdAtIndex >= 0) {
          rowData.CREATED_AT =
            values[r][createdAtIndex] || now;
        }

        const newRow = headers.map(header => {
          return rowData[header] !== undefined
            ? rowData[header]
            : '';
        });

        sh
          .getRange(r + 1, 1, 1, headers.length)
          .setValues([newRow]);

        clearPublicCache_();

        return {
          success: true,
          message: 'Smart Link berhasil diperbarui.'
        };
      }
    }
  }

  const newRow = headers.map(header => {
    return rowData[header] !== undefined
      ? rowData[header]
      : '';
  });

  sh.appendRow(newRow);

  clearPublicCache_();

  return {
    success: true,
    message: 'Smart Link berhasil ditambahkan.'
  };
}

function normalizeScheduleInput_(value) {
  const text = String(value || '').trim();

  if (!text) return '';

  return text
    .replace('T', ' ')
    .substring(0, 16);
}


function parseScheduleDate_(value) {
  if (!value) return null;

  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }

  const text = String(value || '')
    .trim()
    .replace('T', ' ')
    .substring(0, 16);

  if (!text) return null;

  try {
    return Utilities.parseDate(
      text,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm'
    );
  } catch (err) {
    return null;
  }
}


function getLinkScheduleStatus_(link) {
  const status =
    String(link.STATUS || '').toUpperCase();

  if (status !== 'ACTIVE') {
    return 'INACTIVE';
  }

  const now = new Date();

  const startAt =
    parseScheduleDate_(link.START_AT);

  const endAt =
    parseScheduleDate_(link.END_AT);

  if (startAt && now < startAt) {
    return 'SCHEDULED';
  }

  if (endAt && now > endAt) {
    return 'ENDED';
  }

  return 'LIVE';
}


function isLinkPubliclyVisible_(link) {
  return getLinkScheduleStatus_(link) === 'LIVE';
}

function deleteLink(idLink, token) {
  requireAdmin_(token);
  idLink = String(idLink || '').trim();
  if (!idLink) throw new Error('ID link tidak valid.');

  const sh = SS.getSheetByName(SHEET_LINKS);
  if (!sh) throw new Error('Sheet LINKS belum tersedia.');

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idIndex = headers.indexOf('ID_LINK');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIndex]) === idLink) {
      sh.deleteRow(r + 1);

      clearPublicCache_();

      return {
        success: true,
        message: 'Link berhasil dihapus.'
      };
    }
  }

  throw new Error('Link tidak ditemukan.');
}

function toggleLinkStatus(idLink, token) {
  requireAdmin_(token);
  idLink = String(idLink || '').trim();
  if (!idLink) throw new Error('ID link tidak valid.');

  const sh = SS.getSheetByName(SHEET_LINKS);
  if (!sh) throw new Error('Sheet LINKS belum tersedia.');

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());

  const idIndex = headers.indexOf('ID_LINK');
  const statusIndex = headers.indexOf('STATUS');
  const updatedIndex = headers.indexOf('UPDATED_AT');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIndex]) === idLink) {
      const current = String(values[r][statusIndex] || '').toUpperCase();
      const next = current === 'ACTIVE' ? 'INACTIVE' : 'ACTIVE';

      sh.getRange(r + 1, statusIndex + 1).setValue(next);

      if (updatedIndex >= 0) {
        sh.getRange(r + 1, updatedIndex + 1).setValue(formatDateTime_(new Date()));
      }

      clearPublicCache_();

      return {
        success: true,
        message: 'Status link berhasil diubah.',
        status: next
      };
    }
  }

  throw new Error('Link tidak ditemukan.');
}

function generateId_(prefix) {
  return prefix + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

function formatDateTime_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
}

function setupBioLinkAppV2() {
  ensureAppReady_();

  return {
    success: true,
    message: 'Setup otomatis selesai. Semua sheet yang dibutuhkan sudah tersedia.'
  };
}

function setupAdminUsersSheet_() {
  let sh = SS.getSheetByName(SHEET_ADMIN_USERS);
  if (!sh) sh = SS.insertSheet(SHEET_ADMIN_USERS);

  sh.clear();

  const now = formatDateTime_(new Date());

  const data = [
    ['EMAIL', 'PASSWORD', 'NAME', 'ROLE', 'STATUS', 'CREATED_AT'],
    ['admin@ksl.local', 'admin123', 'Admin KSL', 'ADMIN', 'ACTIVE', now]
  ];

  sh.getRange(1, 1, data.length, data[0].length).setValues(data);
  sh.setFrozenRows(1);
}

function setupClickLogSheet_() {
  let sh = SS.getSheetByName(SHEET_CLICK_LOG);
  if (!sh) sh = SS.insertSheet(SHEET_CLICK_LOG);

  sh.clear();

  const headers = [
    'LOG_ID',
    'ID_LINK',
    'TITLE',
    'URL',
    'CLICKED_AT',
    'USER_AGENT'
  ];

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
}

/**
 * Login admin sederhana.
 * Catatan:
 * Untuk produksi, password sebaiknya di-hash.
 */
function adminLogin(payload) {
  ensureAppReady_();

  payload = payload || {};

  const email = String(payload.email || '').trim().toLowerCase();
  const password = String(payload.password || '').trim();

  if (!email) throw new Error('Email wajib diisi.');
  if (!password) throw new Error('Password wajib diisi.');

  const sh = SS.getSheetByName(SHEET_ADMIN_USERS);
  if (!sh) throw new Error('Sheet ADMIN_USERS belum tersedia.');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    sh.appendRow([
      'admin@ksl.local',
      'admin123',
      'Admin KSL',
      'ADMIN',
      'ACTIVE',
      formatDateTime_(new Date())
    ]);
  }

  const newValues = sh.getDataRange().getValues();
  const headers = newValues[0].map(h => String(h).trim());

  const emailIndex = headers.indexOf('EMAIL');
  const passIndex = headers.indexOf('PASSWORD');
  const nameIndex = headers.indexOf('NAME');
  const roleIndex = headers.indexOf('ROLE');
  const statusIndex = headers.indexOf('STATUS');

  if (
    emailIndex < 0 ||
    passIndex < 0 ||
    nameIndex < 0 ||
    roleIndex < 0 ||
    statusIndex < 0
  ) {
    throw new Error('Header ADMIN_USERS belum lengkap.');
  }

  for (let i = 1; i < newValues.length; i++) {
    const rowEmail = String(newValues[i][emailIndex] || '').trim().toLowerCase();
    const rowPassword = String(newValues[i][passIndex] || '').trim();
    const rowStatus = String(newValues[i][statusIndex] || '').trim().toUpperCase();

    if (rowEmail === email && rowPassword === password && rowStatus === 'ACTIVE') {
      const token = generateAdminToken_(email);

      CacheService.getScriptCache().put(
        'ADMIN_TOKEN_' + token,
        JSON.stringify({
          email,
          name: newValues[i][nameIndex] || 'Admin',
          role: newValues[i][roleIndex] || 'ADMIN',
          loginAt: formatDateTime_(new Date())
        }),
        21600
      );

      return {
        success: true,
        message: 'Login berhasil.',
        token,
        user: {
          email,
          name: newValues[i][nameIndex] || 'Admin',
          role: newValues[i][roleIndex] || 'ADMIN'
        }
      };
    }
  }

  throw new Error('Email atau password salah.');
}

function adminMe(token) {
  const user = getAdminUserByToken_(token);

  return {
    success: true,
    user
  };
}

function adminLogout(token) {
  token = String(token || '').trim();

  if (token) {
    CacheService.getScriptCache().remove('ADMIN_TOKEN_' + token);
  }

  return {
    success: true,
    message: 'Logout berhasil.'
  };
}

function getAdminUserByToken_(token) {
  token = String(token || '').trim();
  if (!token) throw new Error('Token tidak tersedia.');

  const raw = CacheService.getScriptCache().get('ADMIN_TOKEN_' + token);
  if (!raw) throw new Error('Sesi login sudah habis. Silakan login ulang.');

  return JSON.parse(raw);
}

function requireAdmin_(token) {
  const user = getAdminUserByToken_(token);

  if (String(user.role || '').toUpperCase() !== 'ADMIN') {
    throw new Error('Akses ditolak.');
  }

  return user;
}

function generateAdminToken_(email) {
  return Utilities.base64EncodeWebSafe(
    email + '|' + new Date().getTime() + '|' + Utilities.getUuid()
  );
}

function recordLinkClick(payload) {
  payload = payload || {};

  const idLink = String(payload.ID_LINK || '').trim();
  const title = String(payload.TITLE || '').trim();
  const url = String(payload.URL || '').trim();
  const userAgent = String(payload.USER_AGENT || '').slice(0, 500);

  if (!idLink) {
    return {
      success: false,
      message: 'ID link kosong.'
    };
  }

  const sh = SS.getSheetByName(SHEET_CLICK_LOG);
  if (!sh) {
    return {
      success: false,
      message: 'Sheet CLICK_LOG belum tersedia.'
    };
  }

  sh.appendRow([
    generateId_('CLK'),
    idLink,
    title,
    url,
    formatDateTime_(new Date()),
    userAgent
  ]);

  return {
    success: true,
    message: 'Klik tercatat.'
  };
}

function getClickStatsMap_() {
  const sh = SS.getSheetByName(SHEET_CLICK_LOG);
  if (!sh) return {};

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return {};

  const headers = values[0].map(h => String(h).trim());
  const idIndex = headers.indexOf('ID_LINK');

  const map = {};

  for (let i = 1; i < values.length; i++) {
    const id = String(values[i][idIndex] || '').trim();
    if (!id) continue;

    map[id] = (map[id] || 0) + 1;
  }

  return map;
}


/**
 * VER 1.2 - Analytics klik berdasarkan CLICK_LOG.
 * days: 7, 30, 90, atau 0 untuk seluruh data.
 */
function getClickAnalytics(days, token) {
  requireAdmin_(token);
  return {
    success: true,
    analytics: getClickAnalytics_(days)
  };
}

function getClickAnalytics_(days) {
  days = Number(days);
  if (![0, 7, 30, 90].includes(days)) days = 30;

  const sh = SS.getSheetByName(SHEET_CLICK_LOG);
  const timezone = Session.getScriptTimeZone();
  const now = new Date();
  const todayKey = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');

  const empty = {
    rangeDays: days,
    totalInRange: 0,
    clicksToday: 0,
    clicks7Days: 0,
    clicks30Days: 0,
    averagePerDay: 0,
    peakDate: '',
    peakClicks: 0,
    dailySeries: [],
    topLinks: [],
    devices: [
      { name: 'Mobile', value: 0 },
      { name: 'Desktop', value: 0 },
      { name: 'Tablet', value: 0 },
      { name: 'Lainnya', value: 0 }
    ]
  };

  if (!sh || sh.getLastRow() < 2) {
    empty.dailySeries = buildEmptyDailySeries_(days || 30, now, timezone);
    return empty;
  }

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const idIndex = headers.indexOf('ID_LINK');
  const titleIndex = headers.indexOf('TITLE');
  const dateIndex = headers.indexOf('CLICKED_AT');
  const uaIndex = headers.indexOf('USER_AGENT');

  if (dateIndex < 0) return empty;

  const cutoff = days > 0
    ? new Date(now.getTime() - (days - 1) * 86400000)
    : null;

  if (cutoff) cutoff.setHours(0, 0, 0, 0);

  const cutoff7 = new Date(now.getTime() - 6 * 86400000);
  cutoff7.setHours(0, 0, 0, 0);
  const cutoff30 = new Date(now.getTime() - 29 * 86400000);
  cutoff30.setHours(0, 0, 0, 0);

  const dailyMap = {};
  const linkMap = {};
  const deviceMap = { Mobile: 0, Desktop: 0, Tablet: 0, Lainnya: 0 };
  let totalInRange = 0;
  let clicksToday = 0;
  let clicks7Days = 0;
  let clicks30Days = 0;
  let earliestDate = null;

  for (let i = 1; i < values.length; i++) {
    const clickedAt = parseClickDate_(values[i][dateIndex], timezone);
    if (!clickedAt) continue;

    const dateKey = Utilities.formatDate(clickedAt, timezone, 'yyyy-MM-dd');
    if (dateKey === todayKey) clicksToday++;
    if (clickedAt >= cutoff7) clicks7Days++;
    if (clickedAt >= cutoff30) clicks30Days++;

    if (cutoff && clickedAt < cutoff) continue;

    totalInRange++;
    if (!earliestDate || clickedAt < earliestDate) earliestDate = clickedAt;
    dailyMap[dateKey] = (dailyMap[dateKey] || 0) + 1;

    const id = idIndex >= 0 ? String(values[i][idIndex] || '').trim() : '';
    const title = titleIndex >= 0 ? String(values[i][titleIndex] || '').trim() : '';
    const linkKey = id || title || 'UNKNOWN';
    if (!linkMap[linkKey]) {
      linkMap[linkKey] = { id: id, title: title || id || 'Tanpa Judul', clicks: 0 };
    }
    linkMap[linkKey].clicks++;

    const ua = uaIndex >= 0 ? String(values[i][uaIndex] || '') : '';
    deviceMap[classifyDevice_(ua)]++;
  }

  const seriesDays = days > 0
    ? days
    : Math.max(1, earliestDate ? Math.ceil((now - earliestDate) / 86400000) + 1 : 30);

  const dailySeries = [];
  for (let offset = seriesDays - 1; offset >= 0; offset--) {
    const d = new Date(now.getTime() - offset * 86400000);
    const key = Utilities.formatDate(d, timezone, 'yyyy-MM-dd');
    dailySeries.push({
      date: key,
      label: Utilities.formatDate(d, timezone, 'dd MMM'),
      clicks: dailyMap[key] || 0
    });
  }

  let peakDate = '';
  let peakClicks = 0;
  dailySeries.forEach(item => {
    if (item.clicks > peakClicks) {
      peakClicks = item.clicks;
      peakDate = item.date;
    }
  });

  const topLinks = Object.keys(linkMap)
    .map(key => linkMap[key])
    .sort((a, b) => b.clicks - a.clicks || a.title.localeCompare(b.title))
    .slice(0, 5);

  return {
    rangeDays: days,
    totalInRange: totalInRange,
    clicksToday: clicksToday,
    clicks7Days: clicks7Days,
    clicks30Days: clicks30Days,
    averagePerDay: Number((totalInRange / Math.max(1, seriesDays)).toFixed(1)),
    peakDate: peakDate,
    peakClicks: peakClicks,
    dailySeries: dailySeries,
    topLinks: topLinks,
    devices: Object.keys(deviceMap).map(name => ({ name: name, value: deviceMap[name] }))
  };
}

function parseClickDate_(value, timezone) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  const text = String(value || '').trim();
  if (!text) return null;

  try {
    return Utilities.parseDate(text, timezone, 'yyyy-MM-dd HH:mm:ss');
  } catch (err) {
    const parsed = new Date(text);
    return isNaN(parsed.getTime()) ? null : parsed;
  }
}

function classifyDevice_(userAgent) {
  const ua = String(userAgent || '').toLowerCase();
  if (!ua) return 'Lainnya';
  if (/ipad|tablet|playbook|silk/.test(ua)) return 'Tablet';
  if (/mobile|android|iphone|ipod|windows phone/.test(ua)) return 'Mobile';
  if (/windows|macintosh|linux|cros/.test(ua)) return 'Desktop';
  return 'Lainnya';
}

function buildEmptyDailySeries_(days, now, timezone) {
  const result = [];
  for (let offset = days - 1; offset >= 0; offset--) {
    const d = new Date(now.getTime() - offset * 86400000);
    result.push({
      date: Utilities.formatDate(d, timezone, 'yyyy-MM-dd'),
      label: Utilities.formatDate(d, timezone, 'dd MMM'),
      clicks: 0
    });
  }
  return result;
}

function ensureAppReady_() {
  ensureProfileSheet_();
  ensureLinksSheet_();
  ensureAdminUsersSheet_();
  ensureClickLogSheet_();

  // VER 1.3
  ensurePageViewLogSheet_();
}

function dedupeProfileKeys_(sh) {
  if (!sh || sh.getLastRow() < 2) return;

  const managedKeys = new Set([
    'APP_TITLE',
    'NAME',
    'BIO',
    'AVATAR_URL',

    'QR_LOGO_ENABLED',
    'PUBLIC_LAYOUT_MODE',

    'THEME_COLOR',
    'BG_STYLE',
    'BUTTON_STYLE',

    'WHATSAPP',
    'INSTAGRAM',
    'YOUTUBE',
    'TRAKTEER'
  ]);

  const values = sh.getDataRange().getValues();

  const firstRowMap = {};
  const duplicateRows = [];

  for (let r = 1; r < values.length; r++) {

    const key = String(
      values[r][0] || ''
    )
      .trim()
      .toUpperCase();

    if (!key) continue;

    /*
     * Hanya key resmi aplikasi yang dibersihkan.
     * Custom key milik user tidak disentuh.
     */
    if (!managedKeys.has(key)) {
      continue;
    }

    if (!firstRowMap[key]) {

      /*
       * Simpan occurrence pertama sebagai canonical row.
       */
      firstRowMap[key] = r + 1;

    } else {

      /*
       * Row berikutnya adalah duplicate.
       */
      duplicateRows.push(r + 1);
    }
  }

  /*
   * WAJIB delete dari bawah agar nomor row
   * tidak bergeser saat proses berjalan.
   */
  duplicateRows
    .sort((a, b) => b - a)
    .forEach(function(rowNumber) {
      sh.deleteRow(rowNumber);
    });
}

function ensureProfileSheet_() {

  let sh =
    SS.getSheetByName(
      SHEET_PROFILE
    );

  if (!sh) {

    sh =
      SS.insertSheet(
        SHEET_PROFILE
      );

    const data = [
      ['KEY', 'VALUE'],

      [
        'APP_TITLE',
        'KSL Bio Link'
      ],

      [
        'NAME',
        'Kucing Storia Labs'
      ],

      [
        'BIO',
        'Web App, Google Sheets, Apps Script, dan Source Code'
      ],

      [
        'AVATAR_URL',
        'https://placehold.co/300x300?text=KSL'
      ],

      [
        'QR_LOGO_ENABLED',
        'YES'
      ],

      [
        'PUBLIC_LAYOUT_MODE',
        'STANDARD'
      ],

      [
        'THEME_COLOR',
        '#2563eb'
      ],

      [
        'BG_STYLE',
        'SOFT'
      ],

      [
        'BUTTON_STYLE',
        'CARD'
      ],

      [
        'WHATSAPP',
        'https://wa.me/6281234567890'
      ],

      [
        'INSTAGRAM',
        'https://instagram.com/username'
      ],

      [
        'YOUTUBE',
        'https://youtube.com/@username'
      ],

      [
        'TRAKTEER',
        'https://trakteer.id/username'
      ]
    ];

    sh
      .getRange(
        1,
        1,
        data.length,
        data[0].length
      )
      .setValues(data);

    sh.setFrozenRows(1);

    return;
  }


  /* =========================================
     PASTIKAN HEADER
  ========================================= */

  const firstCell =
    String(
      sh
        .getRange(1, 1)
        .getValue() || ''
    )
      .trim()
      .toUpperCase();

  if (firstCell !== 'KEY') {

    sh
      .getRange(
        1,
        1,
        1,
        2
      )
      .setValues([
        ['KEY', 'VALUE']
      ]);

    sh.setFrozenRows(1);
  }


  /* =========================================
     FIX VER 1.4
     BERSIHKAN DUPLICATE MANAGED KEY
  ========================================= */

  dedupeProfileKeys_(sh);


  /* =========================================
     MIGRATION NON-DESTRUCTIVE
  ========================================= */

  const requiredKeys = {

    APP_TITLE:
      'KSL Bio Link',

    NAME:
      'Kucing Storia Labs',

    BIO:
      'Web App, Google Sheets, Apps Script, dan Source Code',

    AVATAR_URL:
      'https://placehold.co/300x300?text=KSL',

    QR_LOGO_ENABLED:
      'YES',

    PUBLIC_LAYOUT_MODE:
      'STANDARD',

    THEME_COLOR:
      '#2563eb',

    BG_STYLE:
      'SOFT',

    BUTTON_STYLE:
      'CARD',

    WHATSAPP:
      'https://wa.me/6281234567890',

    INSTAGRAM:
      'https://instagram.com/username',

    YOUTUBE:
      'https://youtube.com/@username',

    TRAKTEER:
      'https://trakteer.id/username'
  };


  const values =
    sh
      .getDataRange()
      .getValues();

  const existingKeys = {};


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const key =
      String(
        values[i][0] || ''
      )
        .trim()
        .toUpperCase();

    if (key) {
      existingKeys[key] = true;
    }
  }


  Object
    .keys(requiredKeys)
    .forEach(function(key) {

      if (!existingKeys[key]) {

        sh.appendRow([
          key,
          requiredKeys[key]
        ]);
      }
    });


  sh.setFrozenRows(1);
}

function ensureLinksSheet_() {
  let sh = SS.getSheetByName(SHEET_LINKS);

  if (!sh) {
    sh = SS.insertSheet(SHEET_LINKS);
  }

  // KOLOM LAMA TETAP PADA POSISI SEMULA
  const headers = [
    'ID_LINK',
    'TITLE',
    'URL',
    'ICON',
    'SORT_ORDER',
    'STATUS',
    'CREATED_AT',
    'UPDATED_AT',

    // =====================================
    // VER 1.3 - SELALU TAMBAH DI PALING KANAN
    // =====================================
    'CTA_TYPE',
    'CAMPAIGN',
    'FEATURED',
    'START_AT',
    'END_AT'
  ];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);

    const now = formatDateTime_(new Date());

    const rows = [
      [
        'LNK-001',
        'Playlist Kas RT',
        'https://youtube.com/',
        '▶️',
        1,
        'ACTIVE',
        now,
        now,

        // VER 1.3
        'WATCH',
        'Konten Utama',
        'NO',
        '',
        ''
      ],
      [
        'LNK-002',
        'Instagram KSL',
        'https://instagram.com/username',
        '📸',
        2,
        'ACTIVE',
        now,
        now,

        // VER 1.3
        'SOCIAL',
        'Social Media',
        'NO',
        '',
        ''
      ],
      [
        'LNK-003',
        'Dukung di Trakteer',
        'https://trakteer.id/username',
        '☕',
        3,
        'ACTIVE',
        now,
        now,

        // VER 1.3
        'BUY',
        'Support KSL',
        'YES',
        '',
        ''
      ]
    ];

    sh.getRange(
      2,
      1,
      rows.length,
      headers.length
    ).setValues(rows);

    sh.setFrozenRows(1);
    return;
  }

  // Existing sheet:
  // hanya menambahkan header yang belum ada
  // ke kolom paling kanan.
  ensureHeaders_(sh, headers);
}

function ensurePageViewLogSheet_() {
  let sh = SS.getSheetByName(SHEET_PAGE_VIEW_LOG);

  if (!sh) {
    sh = SS.insertSheet(SHEET_PAGE_VIEW_LOG);
  }

  const headers = [
    'VIEW_ID',
    'VIEWED_AT',
    'USER_AGENT'
  ];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return;
  }

  ensureHeaders_(sh, headers);
}

function ensureAdminUsersSheet_() {
  let sh = SS.getSheetByName(SHEET_ADMIN_USERS);
  if (!sh) {
    sh = SS.insertSheet(SHEET_ADMIN_USERS);
  }

  const headers = [
    'EMAIL',
    'PASSWORD',
    'NAME',
    'ROLE',
    'STATUS',
    'CREATED_AT'
  ];

  if (sh.getLastRow() === 0) {
    const now = formatDateTime_(new Date());

    const data = [
      headers,
      ['admin@ksl.local', 'admin123', 'Admin KSL', 'ADMIN', 'ACTIVE', now]
    ];

    sh.getRange(1, 1, data.length, headers.length).setValues(data);
    sh.setFrozenRows(1);
    return;
  }

  ensureHeaders_(sh, headers);

  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    sh.appendRow([
      'admin@ksl.local',
      'admin123',
      'Admin KSL',
      'ADMIN',
      'ACTIVE',
      formatDateTime_(new Date())
    ]);
  }
}

function ensureClickLogSheet_() {
  let sh = SS.getSheetByName(SHEET_CLICK_LOG);
  if (!sh) {
    sh = SS.insertSheet(SHEET_CLICK_LOG);
  }

  const headers = [
    'LOG_ID',
    'ID_LINK',
    'TITLE',
    'URL',
    'CLICKED_AT',
    'USER_AGENT'
  ];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return;
  }

  ensureHeaders_(sh, headers);
}

function ensureHeaders_(sh, requiredHeaders) {
  const lastColumn = Math.max(sh.getLastColumn(), requiredHeaders.length);

  const currentHeaders = sh
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(h => String(h || '').trim());

  const headerMap = {};
  currentHeaders.forEach((h, i) => {
    if (h) headerMap[h] = i + 1;
  });

  requiredHeaders.forEach(header => {
    if (!headerMap[header]) {
      const newCol = sh.getLastColumn() + 1;
      sh.getRange(1, newCol).setValue(header);
    }
  });

  sh.setFrozenRows(1);
}

function getPublicData() {
  const start = Date.now();

  const cache =
    CacheService.getScriptCache();

  const cacheKey =
    'KSL_BIO_PUBLIC_DATA_V13';

  const cached =
    cache.get(cacheKey);

  if (cached) {
    const parsed =
      JSON.parse(cached);

    parsed.debug =
      parsed.debug || {};

    parsed.debug.fromCache = true;
    parsed.debug.total =
      Date.now() - start;

    return parsed;
  }

  const timing = {};

  try {

    let t = Date.now();

    const profile =
      getProfile_();

    timing.profile =
      Date.now() - t;

    t = Date.now();

    const rawLinks =
      getLinks_();

    timing.readLinks =
      Date.now() - t;

    t = Date.now();

    const links = rawLinks

      // VER 1.3
      .filter(isLinkPubliclyVisible_)

      .sort((a, b) => {

        // FEATURED selalu ke atas
        const featuredA =
          String(a.FEATURED || '').toUpperCase() === 'YES'
            ? 1 : 0;

        const featuredB =
          String(b.FEATURED || '').toUpperCase() === 'YES'
            ? 1 : 0;

        if (featuredA !== featuredB) {
          return featuredB - featuredA;
        }

        return (
          Number(a.SORT_ORDER || 0) -
          Number(b.SORT_ORDER || 0)
        );
      })

      .map(item => ({
        ID_LINK:
          String(item.ID_LINK || ''),

        TITLE:
          String(item.TITLE || ''),

        URL:
          String(item.URL || ''),

        ICON:
          String(item.ICON || '🔗'),

        SORT_ORDER:
          Number(item.SORT_ORDER || 0),

        STATUS:
          String(item.STATUS || ''),

        CTA_TYPE:
          String(item.CTA_TYPE || 'LINK'),

        CAMPAIGN:
          String(item.CAMPAIGN || ''),

        FEATURED:
          String(item.FEATURED || 'NO'),

        START_AT:
          String(item.START_AT || ''),

        END_AT:
          String(item.END_AT || '')
      }));

    timing.filterSortMap =
      Date.now() - t;

    timing.total =
      Date.now() - start;

    const result = {
      success: true,
      profile,
      links,

      debug: {
        fromCache: false,
        timing,
        totalLinks:
          rawLinks.length,

        activeLinks:
          links.length
      }
    };

    // Jadwal perlu lebih responsif,
    // cache cukup 30 detik
    cache.put(
      cacheKey,
      JSON.stringify(result),
      30
    );

    return result;

  } catch (err) {

    return {
      success: false,
      message:
        err.message ||
        'Server error getPublicData',

      stack:
        err.stack || '',

      debug: {
        fromCache: false,
        timing,
        total:
          Date.now() - start
      }
    };
  }
}

function clearPublicCache_() {

  const cache =
    CacheService.getScriptCache();

  cache.remove(
    'KSL_BIO_PUBLIC_DATA_V1'
  );

  cache.remove(
    'KSL_BIO_PUBLIC_DATA_V13'
  );
}

function recordPageView(payload) {
  payload = payload || {};

  ensurePageViewLogSheet_();

  const sh =
    SS.getSheetByName(
      SHEET_PAGE_VIEW_LOG
    );

  if (!sh) {
    return {
      success: false
    };
  }

  const userAgent =
    String(
      payload.USER_AGENT || ''
    ).slice(0, 500);

  sh.appendRow([
    generateId_('VIEW'),
    formatDateTime_(new Date()),
    userAgent
  ]);

  return {
    success: true
  };
}

function getPageViewStats_() {

  const sh =
    SS.getSheetByName(
      SHEET_PAGE_VIEW_LOG
    );

  if (
    !sh ||
    sh.getLastRow() < 2
  ) {
    return {
      totalViews: 0,
      viewsToday: 0,
      views7Days: 0,
      views30Days: 0
    };
  }

  const values =
    sh.getDataRange().getValues();

  const headers =
    values[0].map(
      h => String(h || '').trim()
    );

  const dateIndex =
    headers.indexOf('VIEWED_AT');

  if (dateIndex < 0) {
    return {
      totalViews: 0,
      viewsToday: 0,
      views7Days: 0,
      views30Days: 0
    };
  }

  const timezone =
    Session.getScriptTimeZone();

  const now =
    new Date();

  const todayKey =
    Utilities.formatDate(
      now,
      timezone,
      'yyyy-MM-dd'
    );

  const cutoff7 =
    new Date(
      now.getTime() -
      (6 * 86400000)
    );

  cutoff7.setHours(
    0, 0, 0, 0
  );

  const cutoff30 =
    new Date(
      now.getTime() -
      (29 * 86400000)
    );

  cutoff30.setHours(
    0, 0, 0, 0
  );

  let totalViews = 0;
  let viewsToday = 0;
  let views7Days = 0;
  let views30Days = 0;

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const viewedAt =
      parseClickDate_(
        values[i][dateIndex],
        timezone
      );

    if (!viewedAt) continue;

    totalViews++;

    const key =
      Utilities.formatDate(
        viewedAt,
        timezone,
        'yyyy-MM-dd'
      );

    if (key === todayKey) {
      viewsToday++;
    }

    if (viewedAt >= cutoff7) {
      views7Days++;
    }

    if (viewedAt >= cutoff30) {
      views30Days++;
    }
  }

  return {
    totalViews,
    viewsToday,
    views7Days,
    views30Days
  };
}

function debugPublicData() {
  const out = {
    success: true,
    step: '',
    sheets: {},
    profile: null,
    linksCount: 0,
    activeLinksCount: 0,
    error: null
  };

  try {
    out.step = 'CHECK_CONST';
    out.consts = {
      SHEET_PROFILE: typeof SHEET_PROFILE !== 'undefined' ? SHEET_PROFILE : 'UNDEFINED',
      SHEET_LINKS: typeof SHEET_LINKS !== 'undefined' ? SHEET_LINKS : 'UNDEFINED',
      SHEET_ADMIN_USERS: typeof SHEET_ADMIN_USERS !== 'undefined' ? SHEET_ADMIN_USERS : 'UNDEFINED',
      SHEET_CLICK_LOG: typeof SHEET_CLICK_LOG !== 'undefined' ? SHEET_CLICK_LOG : 'UNDEFINED'
    };

    out.step = 'CHECK_SHEETS_BEFORE_ENSURE';
    out.sheets.before = SS.getSheets().map(s => s.getName());

    out.step = 'RUN_ENSURE_APP_READY';
    ensureAppReady_();

    out.step = 'CHECK_SHEETS_AFTER_ENSURE';
    out.sheets.after = SS.getSheets().map(s => s.getName());

    out.step = 'READ_PROFILE';
    out.profile = getProfile_();

    out.step = 'READ_LINKS';
    const links = getLinks_();
    out.linksCount = links.length;
    out.activeLinksCount = links.filter(item => String(item.STATUS || '').toUpperCase() === 'ACTIVE').length;
    out.sampleLinks = links.slice(0, 3);

    out.step = 'DONE';
    return out;

  } catch (err) {
    out.success = false;
    out.error = {
      message: err.message,
      stack: err.stack
    };
    return out;
  }
}

function normalizeSheetValue_(value) {
  if (value instanceof Date) {
    return formatDateTime_(value);
  }

  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'number') {
    return value;
  }

  if (typeof value === 'boolean') {
    return value;
  }

  return String(value);
}

function ensureAppReadyOnce_() {
  const props = PropertiesService.getScriptProperties();

  const ready = props.getProperty('BIO_LINK_APP_READY_V13');

  if (ready === 'YES') {
    return;
  }

  ensureAppReady_();

  props.setProperty(
    'BIO_LINK_APP_READY_V13',
    'YES'
  );
}

function resetAppReadyFlag() {
  PropertiesService.getScriptProperties().deleteProperty('BIO_LINK_APP_READY');

  return {
    success: true,
    message: 'Flag setup direset. Buka ulang web app untuk setup otomatis.'
  };
}

function resetAdminPassword(payload, token) {
  const user = requireAdmin_(token);

  payload = payload || {};

  const oldPassword = String(payload.oldPassword || '').trim();
  const newPassword = String(payload.newPassword || '').trim();

  if (!oldPassword) throw new Error('Password lama wajib diisi.');
  if (!newPassword) throw new Error('Password baru wajib diisi.');
  if (newPassword.length < 6) throw new Error('Password baru minimal 6 karakter.');

  const sh = SS.getSheetByName(SHEET_ADMIN_USERS);
  if (!sh) throw new Error('Sheet ADMIN_USERS belum tersedia.');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('Data admin belum tersedia.');

  const headers = values[0].map(h => String(h || '').trim());

  const emailIndex = headers.indexOf('EMAIL');
  const passIndex = headers.indexOf('PASSWORD');
  const statusIndex = headers.indexOf('STATUS');

  if (emailIndex < 0 || passIndex < 0 || statusIndex < 0) {
    throw new Error('Header ADMIN_USERS belum lengkap.');
  }

  const loginEmail = String(user.email || '').trim().toLowerCase();

  for (let r = 1; r < values.length; r++) {
    const rowEmail = String(values[r][emailIndex] || '').trim().toLowerCase();
    const rowPassword = String(values[r][passIndex] || '').trim();
    const rowStatus = String(values[r][statusIndex] || '').trim().toUpperCase();

    if (rowEmail === loginEmail && rowStatus === 'ACTIVE') {
      if (rowPassword !== oldPassword) {
        throw new Error('Password lama tidak sesuai.');
      }

      sh.getRange(r + 1, passIndex + 1).setValue(newPassword);

      CacheService.getScriptCache().remove('ADMIN_TOKEN_' + token);

      return {
        success: true,
        message: 'Password berhasil diubah. Silakan login ulang.'
      };
    }
  }

  throw new Error('User admin aktif tidak ditemukan.');
}

function uploadAvatarToDrive(payload, token) {
  requireAdmin_(token);

  payload = payload || {};

  const fileName =
    String(payload.fileName || 'avatar.png').trim();

  const mimeType =
    String(payload.mimeType || 'image/png').trim();

  const base64 =
    String(payload.base64 || '').trim();

  if (!base64) {
    throw new Error('Data gambar kosong.');
  }

  if (!mimeType.startsWith('image/')) {
    throw new Error('File harus berupa gambar.');
  }

  const match =
    base64.match(/^data:(.+);base64,(.+)$/);

  if (!match) {
    throw new Error('Format gambar tidak valid.');
  }

  const realMimeType =
    match[1];

  const data =
    match[2];

  if (!realMimeType.startsWith('image/')) {
    throw new Error('File harus berupa gambar.');
  }

  const bytes =
    Utilities.base64Decode(data);

  if (bytes.length > 2 * 1024 * 1024) {
    throw new Error('Ukuran gambar maksimal 2MB.');
  }

  const folder =
    getOrCreateBioLinkUploadFolder_();

  const safeName =
    fileName
      .replace(/[^\w.\- ]+/g, '')
      .replace(/\s+/g, '-');

  const finalName =
    'avatar-' +
    formatFileTime_(new Date()) +
    '-' +
    safeName;

  const blob =
    Utilities.newBlob(
      bytes,
      realMimeType,
      finalName
    );

  const file =
    folder.createFile(blob);

  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  const url =
    'https://drive.google.com/thumbnail?id=' +
    file.getId() +
    '&sz=w600';

  upsertProfileValue_(
    'AVATAR_URL',
    url
  );

  console.log(
    'AVATAR SAVED TO PROFILE:',
    url
  );

  clearPublicCache_();

 return {
    success: true,
    message: 'Foto berhasil diupload dan AVATAR_URL tersimpan ke PROFILE.',
    fileId: file.getId(),
    url: url
  };
}


function getQrLogoDataUrl(avatarUrl, token) {
  requireAdmin_(token);

  const url =
    String(avatarUrl || '').trim();

  if (!url) {
    return {
      success: false,
      status: 'EMPTY',
      message: 'Logo belum tersedia.'
    };
  }

  /*
   * Jika sudah data URL,
   * langsung gunakan.
   */
  if (/^data:image\//i.test(url)) {
    return {
      success: true,
      status: 'READY',
      dataUrl: url
    };
  }

  if (!/^https?:\/\//i.test(url)) {
    return {
      success: false,
      status: 'INVALID',
      message: 'URL logo tidak valid.'
    };
  }

  try {

    const response =
      UrlFetchApp.fetch(
        url,
        {
          muteHttpExceptions: true,
          followRedirects: true
        }
      );

    const responseCode =
      response.getResponseCode();

    if (
      responseCode < 200 ||
      responseCode >= 300
    ) {
      return {
        success: false,
        status: 'FETCH_FAILED',
        message: 'Logo gagal dimuat.'
      };
    }

    const blob =
      response.getBlob();

    const contentType =
      String(
        blob.getContentType() || ''
      ).toLowerCase();

    const bytes =
      blob.getBytes();

    if (
      !contentType.startsWith('image/')
    ) {
      return {
        success: false,
        status: 'INVALID_TYPE',
        message: 'File logo bukan gambar.'
      };
    }

    if (
      bytes.length >
      2 * 1024 * 1024
    ) {
      return {
        success: false,
        status: 'TOO_LARGE',
        message: 'Logo maksimal 2MB.'
      };
    }

    const base64 =
      Utilities.base64Encode(bytes);

    return {
      success: true,
      status: 'READY',

      dataUrl:
        'data:' +
        contentType +
        ';base64,' +
        base64
    };

  } catch (err) {

    console.error(
      'GET QR LOGO ERROR:',
      err
    );

    return {
      success: false,
      status: 'FETCH_FAILED',
      message: 'Logo gagal dimuat.'
    };
  }
}

function getOrCreateBioLinkUploadFolder_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty('BIO_LINK_UPLOAD_FOLDER_ID');

  if (existingId) {
    try {
      return DriveApp.getFolderById(existingId);
    } catch (err) {
      props.deleteProperty('BIO_LINK_UPLOAD_FOLDER_ID');
    }
  }

  const folderName = 'KSL Bio Link Uploads';
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    const folder = folders.next();
    props.setProperty('BIO_LINK_UPLOAD_FOLDER_ID', folder.getId());
    return folder;
  }

  const folder = DriveApp.createFolder(folderName);
  props.setProperty('BIO_LINK_UPLOAD_FOLDER_ID', folder.getId());

  return folder;
}

function formatFileTime_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyyMMdd-HHmmss'
  );
}

function upsertProfileValue_(key, value) {
  const sh = SS.getSheetByName(SHEET_PROFILE);

  if (!sh) {
    throw new Error('Sheet PROFILE belum tersedia.');
  }

  const normalizedKey =
    String(key || '')
      .trim()
      .toUpperCase();

  if (!normalizedKey) {
    throw new Error('Profile key tidak valid.');
  }

  const values =
    sh.getDataRange().getValues();

  let foundRow = 0;

  for (let r = 1; r < values.length; r++) {

    const existingKey =
      String(values[r][0] || '')
        .trim()
        .toUpperCase();

    if (existingKey === normalizedKey) {
      foundRow = r + 1;
      break;
    }
  }

  if (foundRow) {

    /*
     * Update VALUE saja.
     * Data PROFILE lain tidak disentuh.
     */
    sh
      .getRange(foundRow, 2)
      .setValue(value);

  } else {

    /*
     * Key belum ada → append.
     */
    sh.appendRow([
      key,
      value
    ]);
  }

  clearPublicCache_();

  return value;
}
