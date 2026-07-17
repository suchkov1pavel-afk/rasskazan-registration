/**
 * Сервер регистрации «Рассказань».
 *
 * 1. Вставьте ID Google Таблицы и папки Google Диска ниже.
 * 2. Один раз запустите setupSheet() или setupSheets().
 * 3. Разверните проект как веб-приложение с доступом «Все».
 */
const SPREADSHEET_ID = 'PASTE_GOOGLE_SHEET_ID_HERE';
const UPLOAD_FOLDER_ID = 'PASTE_GOOGLE_DRIVE_FOLDER_ID_HERE';

const REGISTRATION_SHEET = 'Регистрация';
const MAX_SCREENSHOT_BYTES = 6 * 1024 * 1024;
const PRICE_INCREASE_AT = new Date('2026-07-25T00:00:00+04:00');
const PRICE_INCREASE_AMOUNT = 300;
const REGISTRATION_CLOSE_AT = new Date('2026-07-28T00:00:00+04:00');

const PLAN_PRICES = Object.freeze({
  '29 июля — 2 августа': 2500,
  '30 июля — 2 августа': 1900,
  '31 июля — 2 августа': 1300,
  '1–2 августа': 900,
});

const FAMILY_TIERS = Object.freeze({
  standard: { label: 'Обычный тариф', baseCap: Number.POSITIVE_INFINITY },
  second: { label: 'Второй член многодетной семьи', baseCap: 2000 },
  thirdPlus: { label: 'Третий или последующий член многодетной семьи', baseCap: 1500 },
});

const HEADERS = Object.freeze([
  'ФИО',
  'Возраст',
  'Номер телефона',
  'Город',
  'Церковь',
  'Даты пребывания',
  'Льгота',
  'Тариф',
  'Оплата',
  'Палатка',
  'Спальник',
]);

function doGet() {
  return jsonOutput({
    ok: true,
    service: 'Рассказань — регистрация',
    message: 'Сервис работает',
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (new Date().getTime() >= REGISTRATION_CLOSE_AT.getTime()) {
      throw new Error('Регистрация завершена 28 июля в 00:00.');
    }

    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const normalized = validateAndNormalize(payload);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ensureRegistrationSheet(spreadsheet);
    const screenshotUrl = saveScreenshot(normalized);

    const rows = normalized.participants.map(participant => [
      safeCell(participant.name),
      participant.age,
      safeCell(normalized.phone),
      safeCell(normalized.city),
      safeCell(normalized.church),
      safeCell(normalized.plan),
      safeCell(participant.familyLabel),
      safeCell(formatRubles(participant.price)),
      safeCell(normalized.paymentMethod),
      safeCell(participant.tent),
      safeCell(participant.sleepingBag),
    ]);

    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);
    formatBodyRows(sheet, startRow, rows.length);

    const note = [
      `Заявка: ${normalized.applicationId}`,
      `Создана: ${Utilities.formatDate(new Date(), 'Europe/Samara', 'dd.MM.yyyy HH:mm:ss')}`,
      normalized.comment ? `Комментарий: ${normalized.comment}` : '',
      normalized.sourceUrl ? `Страница: ${normalized.sourceUrl}` : '',
    ].filter(Boolean).join('\n');
    sheet.getRange(startRow, 1, rows.length, 1).setNote(note);

    const tariffNotes = normalized.participants.map(() => [
      [
        `Способ оплаты: ${normalized.paymentMethod}`,
        screenshotUrl ? `Подтверждение: ${screenshotUrl}` : '',
        normalized.priceIncrease ? `Повышение после 25 июля: +${formatRubles(normalized.priceIncrease)}` : 'Цена до 25 июля',
      ].filter(Boolean).join('\n')
    ]);
    sheet.getRange(startRow, 9, rows.length, 1).setNotes(tariffNotes);

    SpreadsheetApp.flush();
    return jsonOutput({
      ok: true,
      applicationId: normalized.applicationId,
      total: normalized.total,
      rowsAdded: rows.length,
    });
  } catch (error) {
    console.error(error);
    return jsonOutput({
      ok: false,
      error: error && error.message ? error.message : String(error),
    });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** Запустите один раз после вставки ID таблицы. */
function setupSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  spreadsheet.setSpreadsheetLocale('ru_RU');
  spreadsheet.setSpreadsheetTimeZone('Europe/Samara');
  const sheet = ensureRegistrationSheet(spreadsheet);
  formatRegistrationSheet(sheet);
  hideLegacySheets(spreadsheet);
  SpreadsheetApp.flush();
}

/** Оставлено для совместимости с предыдущей инструкцией. */
function setupSheets() {
  setupSheet();
}

function ensureRegistrationSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(REGISTRATION_SHEET);
  if (!sheet) sheet = spreadsheet.insertSheet(REGISTRATION_SHEET, 0);

  const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
  migrateLegacyLayoutIfNeeded(sheet, currentHeaders);
  const refreshedHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
  const needsHeaders = refreshedHeaders.join('|') !== HEADERS.join('|');
  if (needsHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  formatRegistrationSheet(sheet);
  return sheet;
}

function migrateLegacyLayoutIfNeeded(sheet, currentHeaders) {
  const previousHeaders = [
    'ФИО','Возраст','Номер телефона','Город','Церковь',
    'Даты пребывания','Льгота','Тариф','Палатка','Спальник'
  ];
  const oldestHeaders = [
    'ФИО','Возраст','Номер телефона','Город','Церковь',
    'Даты пребывания','Палатка','Спальник','Оплата'
  ];

  const previousMatch =
    currentHeaders.slice(0, previousHeaders.length).join('|') === previousHeaders.join('|');
  const oldestMatch =
    currentHeaders.slice(0, oldestHeaders.length).join('|') === oldestHeaders.join('|');

  if ((!previousMatch && !oldestMatch) || sheet.getLastRow() < 2) return;

  let migrated = [];

  if (previousMatch) {
    const oldRows = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, previousHeaders.length)
      .getDisplayValues();

    migrated = oldRows.map(row => {
      const parsed = splitTariffAndPayment(row[7]);
      return [
        row[0], row[1], row[2], row[3], row[4], row[5],
        row[6] || 'Обычный тариф',
        parsed.tariff,
        parsed.payment,
        row[8],
        row[9],
      ];
    });
  } else {
    const oldRows = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, oldestHeaders.length)
      .getDisplayValues();

    migrated = oldRows.map(row => {
      const paymentLines = String(row[8] || '').split(/\n+/).filter(Boolean);
      const parsed = splitTariffAndPayment(paymentLines[0] || '');
      const benefit = paymentLines[1] || 'Обычный тариф';
      return [
        row[0], row[1], row[2], row[3], row[4], row[5],
        benefit,
        parsed.tariff,
        parsed.payment,
        row[6],
        row[7],
      ];
    });
  }

  sheet
    .getRange(
      1,
      1,
      sheet.getLastRow(),
      Math.max(previousHeaders.length, oldestHeaders.length, HEADERS.length)
    )
    .clearContent();

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  if (migrated.length) {
    sheet.getRange(2, 1, migrated.length, HEADERS.length).setValues(migrated);
  }
}

function splitTariffAndPayment(value) {
  const text = String(value || '').trim();
  const paymentMatch = text.match(/(Перевод|Наличными)/i);
  const tariffMatch = text.match(/[\d\s]+(?:₽|руб\.?)/i);

  return {
    tariff: tariffMatch ? tariffMatch[0].replace(/\s+/g, ' ').trim() : text.replace(/[·|]\s*(Перевод|Наличными).*/i, '').trim(),
    payment: paymentMatch
      ? (paymentMatch[1].toLowerCase() === 'перевод' ? 'Перевод' : 'Наличными')
      : '',
  };
}

function formatRegistrationSheet(sheet) {
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(false);
  sheet.setTabColor('#173129');

  const widths = [280, 90, 175, 140, 190, 185, 230, 135, 155, 105, 105];
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));

  const header = sheet.getRange(1, 1, 1, HEADERS.length);
  header
    .setValues([HEADERS])
    .setFontFamily('Arial')
    .setFontSize(12)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#173129')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(1, 48);

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const body = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
  body
    .setFontFamily('Arial')
    .setFontSize(12)
    .setFontColor('#202124')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.getRange(2, 2, lastRow - 1, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 8, lastRow - 1, 4).setHorizontalAlignment('center');

  if (sheet.getFilter()) sheet.getFilter().remove();
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(1, 1, sheet.getLastRow(), HEADERS.length).createFilter();
  }
}

function formatBodyRows(sheet, startRow, rowCount) {
  const range = sheet.getRange(startRow, 1, rowCount, HEADERS.length);
  range
    .setFontFamily('Arial')
    .setFontSize(12)
    .setFontColor('#202124')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.getRange(startRow, 2, rowCount, 1).setHorizontalAlignment('center');
  sheet.getRange(startRow, 8, rowCount, 4).setHorizontalAlignment('center');
  sheet.setRowHeights(startRow, rowCount, 58);

  for (let i = 0; i < rowCount; i++) {
    const row = startRow + i;
    sheet.getRange(row, 1, 1, HEADERS.length)
      .setBackground(row % 2 === 0 ? '#f7faf8' : '#ffffff');
  }

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, sheet.getLastRow(), HEADERS.length).createFilter();
}

function hideLegacySheets(spreadsheet) {
  ['Заявки', 'Участники'].forEach(name => {
    const sheet = spreadsheet.getSheetByName(name);
    if (sheet && !sheet.isSheetHidden()) sheet.hideSheet();
  });

  spreadsheet.getSheets().forEach(sheet => {
    if (sheet.getName() === REGISTRATION_SHEET) return;
    if (/^Лист\d*$|^Sheet\d*$/i.test(sheet.getName()) && sheet.getLastRow() === 0) {
      if (!sheet.isSheetHidden()) sheet.hideSheet();
    }
  });
}

function validateAndNormalize(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Пустая заявка.');
  if (payload.personalDataConsent !== true) throw new Error('Необходимо согласие на обработку персональных данных.');

  const plan = cleanText(payload.plan, 80);
  if (!Object.prototype.hasOwnProperty.call(PLAN_PRICES, plan)) {
    throw new Error('Неизвестные даты пребывания.');
  }

  const rawParticipants = Array.isArray(payload.participants) ? payload.participants : [];
  if (rawParticipants.length < 1 || rawParticipants.length > 30) {
    throw new Error('Некорректное количество участников.');
  }

  const priceIncrease = new Date().getTime() >= PRICE_INCREASE_AT.getTime() ? PRICE_INCREASE_AMOUNT : 0;
  const planPrice = PLAN_PRICES[plan] + priceIncrease;
  const participants = rawParticipants.map((participant, index) => {
    const name = cleanText(participant.name, 150);
    const age = Number(participant.age);
    const familyTier = cleanFamilyTier(participant.familyTier);

    if (name.length < 3) throw new Error(`Проверьте ФИО участника ${index + 1}.`);
    if (!Number.isInteger(age) || age < 13 || age > 99) {
      throw new Error(`Проверьте возраст участника ${index + 1}.`);
    }

    const tier = FAMILY_TIERS[familyTier];
    return {
      name,
      age,
      tent: cleanYesNo(participant.tent),
      sleepingBag: cleanYesNo(participant.sleepingBag),
      familyTier,
      familyLabel: tier.label,
      price: Math.min(
        planPrice,
        Number.isFinite(tier.baseCap) ? tier.baseCap + priceIncrease : Number.POSITIVE_INFINITY
      ),
    };
  });

  const phone = cleanText(payload.phone, 40);
  if ((phone.match(/\d/g) || []).length < 10) throw new Error('Проверьте номер телефона.');

  const paymentMethod = cleanText(payload.paymentMethod, 30);
  if (!['Перевод', 'Наличными'].includes(paymentMethod)) {
    throw new Error('Некорректный способ оплаты.');
  }

  const applicationId = cleanText(payload.applicationId, 60) || createApplicationId();
  const calculatedTotal = participants.reduce((sum, participant) => sum + participant.price, 0);

  return {
    applicationId,
    phone,
    city: cleanText(payload.city, 100),
    church: cleanText(payload.church, 150),
    plan,
    planPrice,
    priceIncrease,
    paymentMethod,
    total: calculatedTotal,
    participants,
    comment: cleanText(payload.comment, 1000),
    sourceUrl: cleanText(payload.sourceUrl, 500),
    screenshot: normalizeScreenshot(payload.screenshot, paymentMethod),
  };
}


function normalizeScreenshot(screenshot, paymentMethod) {
  if (!screenshot) {
    if (paymentMethod === 'Перевод') throw new Error('При переводе нужен скриншот оплаты.');
    return null;
  }

  const name = cleanText(screenshot.name, 160);
  const mimeType = cleanText(screenshot.mimeType, 80);
  const base64 = String(screenshot.base64 || '').replace(/\s/g, '');
  if (!/^image\/(jpeg|png|webp|heic|heif)$/i.test(mimeType)) {
    throw new Error('Подтверждение должно быть изображением.');
  }
  if (!base64) throw new Error('Пустое изображение подтверждения.');
  return { name, mimeType, base64 };
}

function saveScreenshot(normalized) {
  if (!normalized.screenshot) return '';
  const bytes = Utilities.base64Decode(normalized.screenshot.base64);
  if (bytes.length > MAX_SCREENSHOT_BYTES) {
    throw new Error('Скриншот слишком большой. Максимум — 6 МБ.');
  }

  const safeName = normalized.screenshot.name.replace(/[^a-zA-Zа-яА-Я0-9._-]+/g, '_');
  const fileName = `${normalized.applicationId}_${safeName || 'payment.jpg'}`;
  const blob = Utilities.newBlob(bytes, normalized.screenshot.mimeType, fileName);
  const file = DriveApp.getFolderById(UPLOAD_FOLDER_ID).createFile(blob);
  return file.getUrl();
}

function cleanFamilyTier(value) {
  const tier = String(value || 'standard');
  return Object.prototype.hasOwnProperty.call(FAMILY_TIERS, tier) ? tier : 'standard';
}

function cleanYesNo(value) {
  return String(value) === 'Да' ? 'Да' : 'Нет';
}

function cleanText(value, maxLength) {
  return String(value == null ? '' : value)
    .replace(/[\u0000-\u001F\u007F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, maxLength);
}

function safeCell(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

function formatRubles(value) {
  return `${Number(value).toLocaleString('ru-RU')} ₽`;
}

function createApplicationId() {
  return `RZ-${Utilities.formatDate(new Date(), 'Europe/Samara', 'yyyyMMdd-HHmmss')}-${Math.floor(1000 + Math.random() * 9000)}`;
}

function jsonOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
