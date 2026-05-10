const SCHEDULE_URL = 'https://idolmaster-official.jp/schedule';
const CMS_API_BASE_URL = 'https://cmsapi-frontend.idolmaster-official.jp/sitern/api/';
const CMS_TOKEN_ENDPOINT = 'cmsbase/Token/get';
const CMS_ARTICLE_LIST_ENDPOINT = 'idolmaster/Article/list';
const SHEET_NAME = '年間スケジュール';
const LOOKAHEAD_MONTHS = 3;
const EVENT_DURATION_MINUTES = 60;
const CALENDAR_NAME_PREFIX = 'P予定表';
const TIME_ZONE = 'Asia/Tokyo';
const CALENDAR_OPERATION_INTERVAL_MS = 1200;
const CALENDAR_RATE_LIMIT_RETRY_MS = 10000;
const CALENDAR_RATE_LIMIT_MAX_RETRIES = 3;

const HEADERS = [
  'イベントID',
  '状態',
  '日付',
  '開始時刻',
  '終了時刻',
  'タイトル',
  'ブランド',
  '場所',
  'URL',
  '取得日時',
  '同期日時',
  'カレンダーイベントID',
  '同期ハッシュ'
];

const BRANDS = [
  { name: 'THE IDOL M@STER', patterns: ['THE IDOLM@STER', 'THE IDOL M@STER', '765', 'ASOBINOTES', 'Dearly Stars', 'ディアリースターズ'] },
  { name: 'シンデレラガールズ', patterns: ['シンデレラガールズ', 'CINDERELLA'] },
  { name: 'ミリオンライブ！', patterns: ['ミリオンライブ', 'MILLION'] },
  { name: 'SideM', patterns: ['SideM', '315'] },
  { name: 'シャイニーカラーズ', patterns: ['シャイニーカラーズ', 'SHINY COLORS', 'シャニマス'] },
  { name: '学園アイドルマスター', patterns: ['学園アイドルマスター', '学マス', 'Gakuen'] },
  { name: 'その他', patterns: [] }
];

const BRAND_CODE_NAMES = {
  IDOLMASTER: 'THE IDOL M@STER',
  CINDERELLAGIRLS: 'シンデレラガールズ',
  MILLIONLIVE: 'ミリオンライブ！',
  SIDEM: 'SideM',
  SHINYCOLORS: 'シャイニーカラーズ',
  GAKUEN: '学園アイドルマスター',
  OTHER: 'その他'
};

/**
 * 公式スケジュールを取得し、スプレッドシートへ反映します。
 * Apps Script のトリガーにはこの関数、または updateAndSyncProducerSchedule を指定してください。
 */
function updateProducerSchedule() {
  const events = fetchOfficialScheduleEvents();
  const sheet = getScheduleSheet_();
  upsertEventsToSheet_(sheet, events);
}

/**
 * シート上のイベントをブランド別カレンダーへ同期します。
 */
function syncScheduleToCalendars() {
  const sheet = getScheduleSheet_();
  const records = readScheduleRecords_(sheet);
  const syncedAt = new Date();

  records.forEach(function(record) {
    if (!isWithinTargetRange_(record.date)) {
      return;
    }

    const syncHash = buildSyncHash_(record);
    if (shouldSkipCalendarSync_(record, syncHash)) {
      return;
    }

    const eventIdsByBrand = syncRecordToBrandCalendars_(record);
    sheet.getRange(record.rowIndex, HEADERS.indexOf('同期日時') + 1).setValue(syncedAt);
    sheet.getRange(record.rowIndex, HEADERS.indexOf('カレンダーイベントID') + 1).setValue(JSON.stringify(eventIdsByBrand));
    sheet.getRange(record.rowIndex, HEADERS.indexOf('同期ハッシュ') + 1).setValue(syncHash);
  });
}

/**
 * 取得からカレンダー同期までを一括実行します。
 */
function updateAndSyncProducerSchedule() {
  updateProducerSchedule();
  syncScheduleToCalendars();
}

/**
 * 公式サイトからイベント候補を取得します。
 * サイト側の HTML/JSON 構造変更に備え、複数の抽出方法を順に試します。
 */
function fetchOfficialScheduleEvents() {
  const fetchedAt = new Date();
  const apiEvents = fetchOfficialScheduleEventsFromApi_(fetchedAt);

  if (apiEvents.length > 0) {
    return uniqueEvents_(apiEvents).filter(function(event) {
      return isWithinTargetRange_(event.date);
    }).sort(compareEvents_);
  }

  const html = UrlFetchApp.fetch(SCHEDULE_URL, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'User-Agent': 'Producer-Schedule Google Apps Script' }
  }).getContentText('UTF-8');

  const events = parseScheduleEvents_(html, fetchedAt);
  return uniqueEvents_(events).filter(function(event) {
    return isWithinTargetRange_(event.date);
  }).sort(compareEvents_);
}

/**
 * 公式サイトが利用している CMS API からスケジュールを取得します。
 */
function fetchOfficialScheduleEventsFromApi_(fetchedAt) {
  try {
    const token = fetchCmsToken_();
    const range = getTargetRange_();
    const response = fetchCmsArticleList_(token, range.start, range.end);
    const articles = response && response.data && response.data.article_list ? response.data.article_list : [];

    return articles.map(function(article) {
      return normalizeOfficialApiEvent_(article, fetchedAt);
    }).filter(Boolean);
  } catch (error) {
    console.warn('公式 CMS API からの取得に失敗しました。HTML 解析へフォールバックします: ' + error.message);
    return [];
  }
}

/**
 * CMS API 呼び出しに必要な一時トークンを取得します。
 */
function fetchCmsToken_() {
  const response = fetchJson_(CMS_API_BASE_URL + CMS_TOKEN_ENDPOINT);
  const token = response && response.data && response.data.token;

  if (!token) {
    throw new Error('CMS トークンが取得できませんでした。');
  }

  return token;
}

/**
 * 公式 API の記事一覧をスケジュールカテゴリ・対象期間で取得します。
 */
function fetchCmsArticleList_(token, startDate, endDate) {
  const payload = {
    category: ['SCHEDULE'],
    target_start_date: toUnixSeconds_(startDate),
    target_end_date: toUnixSeconds_(endDate)
  };
  const params = {
    site: 'jp',
    ip: 'idolmaster',
    token: token,
    sort: 'asc',
    limit: 200,
    data: JSON.stringify(payload)
  };

  return fetchJson_(CMS_API_BASE_URL + CMS_ARTICLE_LIST_ENDPOINT + '?' + toQueryString_(params));
}

/**
 * 公式 API のスケジュール記事をシート用イベントへ変換します。
 */
function normalizeOfficialApiEvent_(article, fetchedAt) {
  if (!article || article.delflg === '1' || article.publish_status !== 'publish') {
    return null;
  }

  const title = cleanTitle_(article.title);
  const start = unixSecondsToDate_(article.event_startdate);
  const end = unixSecondsToDate_(article.event_enddate);

  if (!title || !start) {
    return null;
  }

  const brands = normalizeApiBrands_(article.brand);
  const hasTime = hasExplicitTime_(start, article.event_dspdate);
  const url = article.event_url || article.url || SCHEDULE_URL;
  const event = buildEvent_(title, start, hasTime, article.event_place || '', absolutizeUrl_(url), fetchedAt, brands);

  if (end && hasTime && end.getTime() > start.getTime()) {
    event.end = end;
  }

  event.id = String(article._id || event.id);
  return event;
}

/**
 * HTML 内の JSON データと本文テキストからイベント情報を抽出します。
 */
function parseScheduleEvents_(html, fetchedAt) {
  const events = [];
  events.push.apply(events, parseNextDataEvents_(html, fetchedAt));
  events.push.apply(events, parseJsonLdEvents_(html, fetchedAt));
  events.push.apply(events, parseTextFallbackEvents_(html, fetchedAt));
  return events;
}

/**
 * Next.js 系ページの __NEXT_DATA__ からイベントらしいオブジェクトを探します。
 */
function parseNextDataEvents_(html, fetchedAt) {
  const match = html.match(/<script[^>]+id=["']__NEXT_DATA__["'][^>]*>([\s\S]*?)<\/script>/i);
  if (!match) {
    return [];
  }

  try {
    const data = JSON.parse(decodeHtmlEntities_(match[1]));
    const candidates = [];
    collectObjects_(data, candidates);
    return candidates.map(function(item) {
      return normalizeEventObject_(item, fetchedAt);
    }).filter(Boolean);
  } catch (error) {
    console.warn('NEXT_DATA の解析に失敗しました: ' + error.message);
    return [];
  }
}

/**
 * JSON-LD の Event 構造からイベントを抽出します。
 */
function parseJsonLdEvents_(html, fetchedAt) {
  const results = [];
  const scripts = html.match(/<script[^>]+type=["']application\/ld\+json["'][^>]*>[\s\S]*?<\/script>/ig) || [];

  scripts.forEach(function(script) {
    const jsonText = script.replace(/^<script[^>]*>/i, '').replace(/<\/script>$/i, '');
    try {
      const data = JSON.parse(decodeHtmlEntities_(jsonText));
      const items = Array.isArray(data) ? data : [data];
      items.forEach(function(item) {
        const event = normalizeEventObject_(item, fetchedAt);
        if (event) {
          results.push(event);
        }
      });
    } catch (error) {
      console.warn('JSON-LD の解析に失敗しました: ' + error.message);
    }
  });

  return results;
}

/**
 * 構造化データが取れない場合の簡易フォールバックです。
 * 日付に続く短いテキストをイベント名として扱います。
 */
function parseTextFallbackEvents_(html, fetchedAt) {
  const text = decodeHtmlEntities_(html)
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, '\n')
    .replace(/\s+/g, ' ')
    .trim();

  const results = [];
  const pattern = /((20\d{2})[./年-](\d{1,2})[./月-](\d{1,2})日?(?:\s*[（(][^)）]+[)）])?(?:\s+(\d{1,2}):(\d{2}))?)\s+(.{4,120}?)(?=20\d{2}[./年-]\d{1,2}[./月-]\d{1,2}|$)/g;
  let match;

  while ((match = pattern.exec(text)) !== null) {
    const year = Number(match[2]);
    const month = Number(match[3]);
    const day = Number(match[4]);
    const hour = match[5] ? Number(match[5]) : null;
    const minute = match[6] ? Number(match[6]) : 0;
    const title = cleanTitle_(match[7]);

    if (!title) {
      continue;
    }

    results.push(buildEvent_(title, new Date(year, month - 1, day, hour || 0, minute, 0), hour !== null, '', SCHEDULE_URL, fetchedAt));
  }

  return results;
}

/**
 * 任意のオブジェクトをイベント形式へ正規化します。
 */
function normalizeEventObject_(item, fetchedAt) {
  if (!item || typeof item !== 'object') {
    return null;
  }

  const title = cleanTitle_(item.title || item.name || item.eventName || item.subject || item.heading);
  const dateValue = item.startDate || item.date || item.start_at || item.startAt || item.eventDate || item.publishedAt;
  const url = item.url || item.link || item.href || SCHEDULE_URL;
  const location = extractLocation_(item);

  if (!title || !dateValue) {
    return null;
  }

  const parsed = parseDateValue_(String(dateValue));
  if (!parsed) {
    return null;
  }

  return buildEvent_(title, parsed.date, parsed.hasTime, location, absolutizeUrl_(url), fetchedAt);
}

/**
 * シートへイベントを追加または更新し、取得範囲内で消えた既存イベントを中止扱いにします。
 */
function upsertEventsToSheet_(sheet, events) {
  const fetchedAt = new Date();
  const existing = readScheduleRecords_(sheet);
  const existingById = {};
  existing.forEach(function(record) {
    existingById[record.id] = record;
  });

  const seenIds = {};

  events.forEach(function(event) {
    seenIds[event.id] = true;
    const rowValues = eventToRow_(event, fetchedAt);
    const existingRecord = existingById[event.id];

    if (existingRecord) {
      rowValues[HEADERS.indexOf('同期日時')] = existingRecord.syncedAt || '';
      rowValues[HEADERS.indexOf('カレンダーイベントID')] = existingRecord.calendarEventIds || '';
      rowValues[HEADERS.indexOf('同期ハッシュ')] = existingRecord.syncHash || '';
      sheet.getRange(existingRecord.rowIndex, 1, 1, HEADERS.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
  });

  existing.forEach(function(record) {
    if (isWithinTargetRange_(record.date) && !seenIds[record.id] && record.status !== 'CANCELED') {
      sheet.getRange(record.rowIndex, HEADERS.indexOf('状態') + 1).setValue('CANCELED');
      sheet.getRange(record.rowIndex, HEADERS.indexOf('取得日時') + 1).setValue(fetchedAt);
    }
  });

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).sort([
      { column: HEADERS.indexOf('日付') + 1, ascending: true },
      { column: HEADERS.indexOf('開始時刻') + 1, ascending: true },
      { column: HEADERS.indexOf('タイトル') + 1, ascending: true }
    ]);
  }
}

/**
 * 1 件のイベントを関連ブランドごとのカレンダーに同期します。
 */
function syncRecordToBrandCalendars_(record) {
  const currentIds = safeJsonParse_(record.calendarEventIds, {});
  const nextIds = {};

  record.brands.forEach(function(brand) {
    const calendar = getOrCreateCalendar_(brand);
    const previousId = currentIds[brand];
    const calendarEvent = previousId ? findCalendarEvent_(calendar, previousId) : null;
    const title = record.status === 'CANCELED' ? '【中止】' + record.title : record.title;
    const description = buildCalendarDescription_(record);

    let event = calendarEvent;
    if (!event) {
      event = createCalendarEvent_(calendar, record, title, description);
    } else {
      updateCalendarEvent_(event, record, title, description);
    }

    nextIds[brand] = event.getId();
    waitForCalendarQuota_();
  });

  return nextIds;
}

/**
 * 前回同期時と内容が同じ場合は、CalendarApp の更新処理を省略します。
 */
function shouldSkipCalendarSync_(record, syncHash) {
  if (!record.syncHash || record.syncHash !== syncHash) {
    return false;
  }

  const currentIds = safeJsonParse_(record.calendarEventIds, {});
  return record.brands.every(function(brand) {
    return !!currentIds[brand];
  });
}

/**
 * カレンダー予定に反映する項目だけから差分判定用ハッシュを作ります。
 */
function buildSyncHash_(record) {
  const values = [
    record.status,
    record.title,
    Utilities.formatDate(record.date, TIME_ZONE, 'yyyy-MM-dd'),
    record.hasTime ? Utilities.formatDate(record.start, TIME_ZONE, 'HH:mm') : '',
    record.hasTime ? Utilities.formatDate(record.end, TIME_ZONE, 'HH:mm') : '',
    record.brands.join(','),
    record.location || '',
    record.url || ''
  ];

  return createHash_(values.join('|'));
}

/**
 * カレンダー予定を新規作成します。
 */
function createCalendarEvent_(calendar, record, title, description) {
  if (record.hasTime) {
    return runCalendarOperation_(function() {
      return calendar.createEvent(title, record.start, record.end, {
        location: record.location,
        description: description
      });
    });
  }

  return runCalendarOperation_(function() {
    return calendar.createAllDayEvent(title, record.date, {
      location: record.location,
      description: description
    });
  });
}

/**
 * 既存のカレンダー予定を最新情報で更新します。
 */
function updateCalendarEvent_(event, record, title, description) {
  runCalendarOperation_(function() {
    event.setTitle(title);
    event.setLocation(record.location || '');
    event.setDescription(description);

    if (record.hasTime) {
      event.setTime(record.start, record.end);
    } else {
      event.setAllDayDate(record.date);
    }
  });
}

function buildCalendarDescription_(record) {
  return [
    '公式スケジュールから同期しました。',
    '状態: ' + record.status,
    'ブランド: ' + record.brands.join(', '),
    record.url ? 'URL: ' + record.url : ''
  ].filter(Boolean).join('\n');
}

function getScheduleSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    if (currentHeaders.join() !== HEADERS.join()) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }
  }

  sheet.setFrozenRows(1);
  return sheet;
}

function readScheduleRecords_(sheet) {
  if (sheet.getLastRow() < 2) {
    return [];
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  return rows.map(function(row, index) {
    const date = normalizeSheetDate_(row[HEADERS.indexOf('日付')]);
    if (!date) {
      return null;
    }

    const startText = String(row[HEADERS.indexOf('開始時刻')] || '');
    const hasTime = !!startText;
    const start = hasTime ? combineDateAndTime_(date, startText) : new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const endText = String(row[HEADERS.indexOf('終了時刻')] || '');
    const end = hasTime && endText ? combineDateAndTime_(date, endText) : new Date(start.getTime() + EVENT_DURATION_MINUTES * 60000);

    return {
      rowIndex: index + 2,
      id: String(row[HEADERS.indexOf('イベントID')] || ''),
      status: String(row[HEADERS.indexOf('状態')] || 'ACTIVE'),
      date: date,
      start: start,
      end: end,
      hasTime: hasTime,
      title: String(row[HEADERS.indexOf('タイトル')] || ''),
      brands: splitBrands_(row[HEADERS.indexOf('ブランド')]),
      location: String(row[HEADERS.indexOf('場所')] || ''),
      url: String(row[HEADERS.indexOf('URL')] || ''),
      syncedAt: row[HEADERS.indexOf('同期日時')] || '',
      calendarEventIds: String(row[HEADERS.indexOf('カレンダーイベントID')] || ''),
      syncHash: String(row[HEADERS.indexOf('同期ハッシュ')] || '')
    };
  }).filter(function(record) {
    return record && record.id && record.title && record.date;
  });
}

function eventToRow_(event, fetchedAt) {
  return [
    event.id,
    event.status,
    event.date,
    event.hasTime ? formatTime_(event.start) : '',
    event.hasTime ? formatTime_(event.end) : '',
    event.title,
    event.brands.join(', '),
    event.location,
    event.url,
    fetchedAt,
    '',
    '',
    ''
  ];
}

function buildEvent_(title, start, hasTime, location, url, fetchedAt, detectedBrands) {
  const date = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const end = hasTime ? new Date(start.getTime() + EVENT_DURATION_MINUTES * 60000) : new Date(date);
  const brands = detectedBrands && detectedBrands.length ? detectedBrands : detectBrands_(title + ' ' + location + ' ' + url);

  return {
    id: buildEventId_(date, title, brands, url),
    status: isCanceledTitle_(title) ? 'CANCELED' : 'ACTIVE',
    date: date,
    start: hasTime ? start : date,
    end: end,
    hasTime: hasTime,
    title: title,
    brands: brands,
    location: location || '',
    url: url || SCHEDULE_URL,
    fetchedAt: fetchedAt
  };
}

function detectBrands_(text) {
  const normalized = String(text || '').toLowerCase();
  const matched = BRANDS.filter(function(brand) {
    return brand.patterns.some(function(pattern) {
      return normalized.indexOf(pattern.toLowerCase()) !== -1;
    });
  }).map(function(brand) {
    return brand.name;
  });

  return matched.length ? matched : ['その他'];
}

function normalizeApiBrands_(brands) {
  if (!Array.isArray(brands) || brands.length === 0) {
    return ['その他'];
  }

  const names = brands.map(function(brand) {
    return BRAND_CODE_NAMES[brand.code] || brand.name || 'その他';
  }).filter(Boolean);

  return names.length ? uniqueValues_(names) : ['その他'];
}

function getOrCreateCalendar_(brand) {
  const name = CALENDAR_NAME_PREFIX + '（' + brand + '）';
  const calendars = CalendarApp.getCalendarsByName(name);

  if (calendars.length) {
    return calendars[0];
  }

  return runCalendarOperation_(function() {
    return CalendarApp.createCalendar(name);
  });
}

/**
 * Calendar API の短時間連続実行による制限を避けるため、操作間隔を空けます。
 */
function waitForCalendarQuota_() {
  Utilities.sleep(CALENDAR_OPERATION_INTERVAL_MS);
}

/**
 * Calendar API 操作を一定間隔で実行し、短時間制限に当たった場合だけ再試行します。
 */
function runCalendarOperation_(operation) {
  let lastError = null;

  for (let attempt = 0; attempt <= CALENDAR_RATE_LIMIT_MAX_RETRIES; attempt++) {
    try {
      const result = operation();
      waitForCalendarQuota_();
      return result;
    } catch (error) {
      lastError = error;
      if (!isCalendarRateLimitError_(error) || attempt === CALENDAR_RATE_LIMIT_MAX_RETRIES) {
        throw error;
      }
      Utilities.sleep(CALENDAR_RATE_LIMIT_RETRY_MS * (attempt + 1));
    }
  }

  throw lastError;
}

function isCalendarRateLimitError_(error) {
  return /too many calendars|too many calendar events|try again later|Service invoked too many times/i.test(String(error && error.message || error));
}

function findCalendarEvent_(calendar, eventId) {
  try {
    return calendar.getEventById(eventId);
  } catch (error) {
    return null;
  }
}

function collectObjects_(value, output) {
  if (Array.isArray(value)) {
    value.forEach(function(item) {
      collectObjects_(item, output);
    });
    return;
  }

  if (!value || typeof value !== 'object') {
    return;
  }

  if (looksLikeEventObject_(value)) {
    output.push(value);
  }

  Object.keys(value).forEach(function(key) {
    collectObjects_(value[key], output);
  });
}

function looksLikeEventObject_(item) {
  const hasTitle = !!(item.title || item.name || item.eventName || item.subject || item.heading);
  const hasDate = !!(item.startDate || item.date || item.start_at || item.startAt || item.eventDate || item.publishedAt);
  return hasTitle && hasDate;
}

function extractLocation_(item) {
  if (!item.location) {
    return item.place || item.venue || '';
  }

  if (typeof item.location === 'string') {
    return item.location;
  }

  return item.location.name || item.location.address || '';
}

function parseDateValue_(value) {
  const normalized = value.replace(/[年月]/g, '/').replace(/日/g, '').replace(/[：]/g, ':');
  const isoDate = normalized.match(/(20\d{2})[-/](\d{1,2})[-/](\d{1,2})(?:[T\s]+(\d{1,2}):(\d{2}))?/);

  if (!isoDate) {
    return null;
  }

  const year = Number(isoDate[1]);
  const month = Number(isoDate[2]);
  const day = Number(isoDate[3]);
  const hour = isoDate[4] ? Number(isoDate[4]) : 0;
  const minute = isoDate[5] ? Number(isoDate[5]) : 0;

  return {
    date: new Date(year, month - 1, day, hour, minute, 0),
    hasTime: !!isoDate[4]
  };
}

function isWithinTargetRange_(date) {
  if (!date) {
    return false;
  }

  const range = getTargetRange_();
  const start = range.start;
  const end = range.end;
  const target = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  return target >= start && target < end;
}

function getTargetRange_() {
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const end = new Date(start.getFullYear(), start.getMonth() + LOOKAHEAD_MONTHS, start.getDate() + 1);
  return { start: start, end: end };
}

function uniqueEvents_(events) {
  const seen = {};
  return events.filter(function(event) {
    if (!event || seen[event.id]) {
      return false;
    }
    seen[event.id] = true;
    return true;
  });
}

function uniqueValues_(values) {
  const seen = {};
  return values.filter(function(value) {
    if (seen[value]) {
      return false;
    }
    seen[value] = true;
    return true;
  });
}

function compareEvents_(a, b) {
  return a.start.getTime() - b.start.getTime() || a.title.localeCompare(b.title, 'ja');
}

function buildEventId_(date, title, brands, url) {
  const source = [
    Utilities.formatDate(date, TIME_ZONE, 'yyyy-MM-dd'),
    title,
    brands.join('|'),
    url || ''
  ].join('|');
  return createHash_(source).slice(0, 24);
}

function createHash_(source) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, source, Utilities.Charset.UTF_8);
  return digest.map(function(byte) {
    const value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('');
}

function cleanTitle_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .replace(/^[-｜|:：\s]+/, '')
    .trim();
}

function isCanceledTitle_(title) {
  return /中止|延期|キャンセル|CANCEL/i.test(title);
}

function formatTime_(date) {
  return Utilities.formatDate(date, TIME_ZONE, 'HH:mm');
}

function normalizeSheetDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const parsed = parseDateValue_(String(value));
  return parsed ? new Date(parsed.date.getFullYear(), parsed.date.getMonth(), parsed.date.getDate()) : null;
}

function combineDateAndTime_(date, timeText) {
  const match = String(timeText).match(/(\d{1,2}):(\d{2})/);
  if (!match) {
    return new Date(date);
  }

  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), Number(match[1]), Number(match[2]), 0);
}

function splitBrands_(value) {
  const brands = String(value || '').split(',').map(function(brand) {
    return brand.trim();
  }).filter(Boolean);
  return brands.length ? brands : ['その他'];
}

function fetchJson_(url) {
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'User-Agent': 'Producer-Schedule Google Apps Script' }
  });
  const statusCode = response.getResponseCode();
  const body = response.getContentText('UTF-8');

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('HTTP ' + statusCode + ': ' + body.slice(0, 200));
  }

  return JSON.parse(body);
}

function toQueryString_(params) {
  return Object.keys(params).map(function(key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');
}

function toUnixSeconds_(date) {
  return Math.floor(date.getTime() / 1000);
}

function unixSecondsToDate_(value) {
  const seconds = Number(value);
  if (!seconds) {
    return null;
  }

  return new Date(seconds * 1000);
}

function hasExplicitTime_(date, displayDate) {
  if (date.getHours() !== 0 || date.getMinutes() !== 0) {
    return true;
  }

  return /\d{1,2}:\d{2}/.test(String(displayDate || ''));
}

function safeJsonParse_(value, fallback) {
  try {
    return value ? JSON.parse(value) : fallback;
  } catch (error) {
    return fallback;
  }
}

function absolutizeUrl_(url) {
  const value = String(url || '');
  if (!value) {
    return SCHEDULE_URL;
  }
  if (/^https?:\/\//i.test(value)) {
    return value;
  }
  if (value.charAt(0) === '/') {
    return 'https://idolmaster-official.jp' + value;
  }
  return SCHEDULE_URL;
}

function decodeHtmlEntities_(text) {
  return String(text || '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ');
}
