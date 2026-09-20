/**
 * ============================================================
 * FINAL Code.gs
 * Система аудита / чек-лист
 * ============================================================
 *
 * Функции:
 *
 * 1. Читает исходную Google Таблицу.
 * 2. Разбивает пункты на разделы по слову "РАЗДЕЛ".
 * 3. Для каждого раздела предусматривается ответственный.
 * 4. Варианты ответа:
 *      - Да
 *      - Нет
 *      - Значительные нарушения
 *      - Незначительные недочёты
 *      - Пропущено
 * 5. Для:
 *      - Нет
 *      - Значительные нарушения
 *      - Незначительные недочёты
 *    требуется комментарий.
 * 6. Поддерживает несколько фотографий на один пункт.
 * 7. Каждая фотография создаётся отдельным файлом.
 * 8. После отправки создаётся отдельная Google Таблица:
 *
 *    Отчёт_по_аудиту_ОБЪЕКТ_ДАТА
 *
 * 9. В реестре:
 *
 *    Дата/время
 *    Дата проведения
 *    Объект
 *    Раздел
 *    Ответственный
 *    Пункт
 *    Ответ
 *    Комментарий
 *    Фото
 *
 * 10. Первоначальная настройка не создаёт файлы/папки
 *     на Google Drive, что существенно снижает вероятность
 *     ошибки "Exceeded maximum execution time".
 *
 * ============================================================
 */


/* ============================================================
 * НАСТРОЙКИ
 * ============================================================
 *
 * ВАЖНО:
 * Здесь укажите точное название вкладки исходного чек-листа.
 */

const CFG = {

  SOURCE_SHEET_NAME: 'Исходный_чек-лист',

  /*
   * Строки, начинающиеся с этого слова,
   * считаются заголовками разделов.
   */
  SECTION_PREFIX: 'РАЗДЕЛ:',

  /*
   * Служебные листы.
   */
  QUESTIONS_SHEET_NAME: '_AUDIT_QUESTIONS',
  SETTINGS_SHEET_NAME: '_AUDIT_SETTINGS',

  /*
   * Папки Google Drive.
   */
  REPORT_FOLDER_NAME: 'Отчёты по аудитам',
  PHOTO_FOLDER_NAME: 'Фото',

  /*
   * Куда сохранять РЕЗУЛЬТАТЫ аудита.
   *
   * 'SCRIPT_ACCOUNT' — в Drive аккаунта, под которым выполняется Apps Script.
   * 'SECOND_ACCOUNT' — в указанную ниже папку второго аккаунта.
   *
   * Для SECOND_ACCOUNT второй аккаунт должен предоставить аккаунту Apps Script
   * доступ РЕДАКТОРА к указанной папке.
   */
  RESULT_STORAGE_MODE: 'SECOND_ACCOUNT',
  SECOND_RESULTS_ROOT_FOLDER_ID: '1JKvLdF_M4Cqvi5JcoVykcJql-HPBxxkt',

  /* Название служебной подпапки для пакетной загрузки фотографий. */
  UPLOAD_RESULTS_FOLDER_NAME: 'Фото_аудитов_выгрузка_v2',

  /*
   * Название листа в отдельном файле отчёта.
   */
  RESPONSE_SHEET_NAME: 'Реестр ответов',

  /* Лист итоговой детализации по шаблону. */
  RESULT_SHEET_NAME: 'Результат',

  /*
   * Варианты ответа.
   *
   * "Пропущено" возвращён.
   */
  ANSWERS: [
    'Да',
    'Нет',
    'Значительные нарушения',
    'Незначительные недочёты',
    'Пропущено'
  ],

  /*
   * Для этих ответов требуется описание нарушения.
   */
  NEGATIVE_ANSWERS: [
    'Нет',
    'Значительные нарушения',
    'Незначительные недочёты'
  ],

  /*
   * Максимальное количество фотографий
   * на один пункт.
   */
  MAX_PHOTOS: 20,

  /*
   * Ограничения чтения исходной таблицы.
   *
   * Они специально ограничены для защиты
   * от превышения времени выполнения.
   */
  MAX_ROWS: 5000,
  MAX_COLS: 30

};


/* ============================================================
 * WEB APP
 * ============================================================ */

/**
 * Запуск Web App.
 *
 * Требуется файл Index.html в этом же проекте.
 */
function doGet(e) {

  // PWA: при первом запуске с GitHub Pages данные формы
  // запрашиваются через JSONP. Обычный запуск Apps Script
  // по-прежнему открывает Index.html.
  const action = e && e.parameter ? String(e.parameter.action || '') : '';
  const callback = e && e.parameter ? String(e.parameter.callback || '') : '';

  if (action === 'getAppData' && callback) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return ContentService
        .createTextOutput('throw new Error("Недопустимый callback.");')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    try {
      const data = getAppData();
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify(data) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    } catch (err) {
      const message = String(err && err.message ? err.message : err)
        .replace(/\\/g, '\\\\')
        .replace(/"/g, '\\"')
        .replace(/\r?\n/g, '\\n');
      return ContentService
        .createTextOutput(callback + '(null);')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
  }

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Аудит / проверка')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/* ============================================================
 * МЕНЮ GOOGLE ТАБЛИЦЫ
 * ============================================================ */

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu('Аудит')

    .addItem(
      '1. Диагностика исходной таблицы',
      'diagnoseChecklistApp'
    )

    .addItem(
      '2. Подготовить приложение',
      'setupChecklistApp'
    )

    .addItem(
      '3. Показать URL приложения',
      'showWebAppUrl'
    )

    .addToUi();

}


/* ============================================================
 * ПОДСЧЁТ РАЗДЕЛОВ
 * ============================================================ */

/**
 * Возвращает количество уникальных разделов.
 *
 * Эта функция обязательна для diagnoseChecklistApp()
 * и setupChecklistApp().
 */
function countSections_(questions) {

  const sections = {};

  (questions || []).forEach(function(q) {

    const section =
      String(q.section || '').trim();

    if (section) {
      sections[section] = true;
    }

  });

  return Object.keys(sections).length;

}


/* ============================================================
 * ПОЛУЧЕНИЕ ИСХОДНОГО ЛИСТА
 * ============================================================ */

/**
 * Получает только тот лист, который указан
 * в CFG.SOURCE_SHEET_NAME.
 *
 * Никакого перебора всех листов.
 */
function getSourceSheet_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'Не удалось определить исходную Google Таблицу.'
    );
  }

  /*
   * Новый чек-лист имеет табличную структуру:
   *
   * ПУНКТ | РАЗДЕЛ | Вес
   *
   * Поэтому больше не требуется искать лист
   * по фиксированному названию "Исходный_чек-лист".
   *
   * Ищем лист по заголовкам колонок.
   */
  const sheets = ss.getSheets();

  for (let i = 0; i < sheets.length; i++) {

    const sheet = sheets[i];

    if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 2) {
      continue;
    }

    const headers =
      sheet
        .getRange(
          1,
          1,
          1,
          Math.min(sheet.getLastColumn(), 3)
        )
        .getDisplayValues()[0]
        .map(function(value) {
          return String(value || '').trim().toUpperCase();
        });

    /*
     * Новый формат:
     * A = ПУНКТ
     * B = РАЗДЕЛ
     * C = Вес
     */
    if (
      headers[0] === 'ПУНКТ' &&
      headers[1] === 'РАЗДЕЛ'
    ) {
      return sheet;
    }

  }

  throw new Error(
    'Не найден лист нового чек-листа.\n\n' +
    'Ожидаются заголовки в первой строке:\n' +
    'ПУНКТ | РАЗДЕЛ | Вес'
  );

}

/* ============================================================
 * РАЗБОР ИСХОДНОЙ ТАБЛИЦЫ
 * ============================================================ */

/**
 * Разбирает исходную таблицу.
 *
 * Логика:
 *
 * РАЗДЕЛ 1
 *     пункт
 *     пункт
 *     пункт
 *
 * РАЗДЕЛ 2
 *     пункт
 *     пункт
 *
 * и т.д.
 */
function parseSourceSheet_(sheet) {

  const lastRow =
    Math.min(
      Math.max(
        sheet.getLastRow(),
        1
      ),
      CFG.MAX_ROWS
    );

  const lastCol =
    Math.min(
      Math.max(
        sheet.getLastColumn(),
        1
      ),
      CFG.MAX_COLS
    );

  if (lastRow < 2 || lastCol < 2) {
    return [];
  }

  /*
   * Новый формат чек-листа:
   *
   * A — ПУНКТ
   * B — РАЗДЕЛ
   * C — Вес
   *
   * Первая строка — заголовки.
   */
  const values =
    sheet
      .getRange(
        1,
        1,
        lastRow,
        Math.min(lastCol, 4)
      )
      .getDisplayValues();

  const headers =
    values[0].map(function(value) {
      return String(value || '').trim().toUpperCase();
    });

  const pointCol = headers.indexOf('ПУНКТ');
  const sectionCol = headers.indexOf('РАЗДЕЛ');
  const weightCol = headers.indexOf('ВЕС');
  const answerCountCol = headers.indexOf('КОЛИЧЕСТВО ВАРИАНТОВ ОТВЕТА');

  if (pointCol < 0 || sectionCol < 0) {
    throw new Error(
      'Неверная структура чек-листа.\n\n' +
      'Не найдены обязательные колонки:\n' +
      'ПУНКТ и РАЗДЕЛ.'
    );
  }

  const result = [];
  let number = 0;

  /*
   * Каждая строка нового чек-листа = один пункт.
   *
   * Пример:
   * ПУНКТ | РАЗДЕЛ | Вес
   * Тротуар у входа чистый | Внутренняя служебная территория! | 20
   */
  for (let r = 1; r < values.length; r++) {

    const row = values[r];

    const point =
      String(
        row[pointCol] || ''
      ).trim();

    const section =
      String(
        row[sectionCol] || ''
      ).trim();

    /*
     * Пустые строки пропускаем.
     */
    if (!point || !section) {
      continue;
    }

    /*
     * В новой таблице вес хранится в отдельной колонке.
     * Пока он только извлекается и сохраняется в объект вопроса.
     */
    const weight =
      weightCol >= 0
        ? String(
            row[weightCol] || ''
          ).trim()
        : '';

    /*
     * В 4-м столбце указывается количество
     * вариантов ответа для конкретного пункта.
     */
    const answerCount =
      answerCountCol >= 0
        ? Number(
            String(row[answerCountCol] || '').replace(',', '.').trim()
          )
        : 5;

    number++;

    result.push({

      id:
        'Q' + number,

      section:
        section,

      point:
        point,

      weight:
        weight,

      answerCount:
        answerCount === 3 ? 3 : 5,

      sourceRow:
        r + 1

    });

  }

  return result;

}



/* ============================================================
 * ДИАГНОСТИКА
 * ============================================================ */

/**
 * Быстрая диагностика.
 *
 * Эта функция:
 *
 * - не создаёт отчёт;
 * - не создаёт папки;
 * - не работает с фотографиями;
 * - не перебирает Drive.
 */
function diagnoseChecklistApp() {

  const started =
    Date.now();

  const source =
    getSourceSheet_();

  const questions =
    parseSourceSheet_(
      source
    );

  const result = {

    spreadsheet:
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getName(),

    sourceSheet:
      source.getName(),

    rows:
      source.getLastRow(),

    columns:
      source.getLastColumn(),

    sections:
      countSections_(
        questions
      ),

    questions:
      questions.length,

    milliseconds:
      Date.now() - started

  };

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  try {

    SpreadsheetApp
      .getUi()
      .alert(

        'Диагностика завершена\n\n' +

        'Исходный лист: ' +
        result.sourceSheet +

        '\nСтрок: ' +
        result.rows +

        '\nКолонок: ' +
        result.columns +

        '\nРазделов: ' +
        result.sections +

        '\nПунктов: ' +
        result.questions +

        '\nВремя выполнения: ' +
        result.milliseconds +
        ' мс'

      );

  } catch (e) {

    console.log(e);

  }

  return result;

}


/* ============================================================
 * ПЕРВОНАЧАЛЬНАЯ НАСТРОЙКА
 * ============================================================ */

/**
 * Основная функция первоначальной подготовки.
 *
 * ВАЖНО:
 *
 * Здесь специально НЕТ DriveApp.
 * Поэтому функция работает значительно быстрее
 * предыдущей версии.
 */
function setupChecklistApp() {

  const lock =
    LockService.getScriptLock();

  if (
    !lock.tryLock(5000)
  ) {

    throw new Error(
      'Другой запуск скрипта уже выполняется.\n' +
      'Подождите несколько секунд.'
    );

  }

  try {

    const started =
      Date.now();

    const ss =
      SpreadsheetApp
        .getActiveSpreadsheet();

    if (!ss) {

      throw new Error(
        'Не удалось определить Google Таблицу.'
      );

    }

    const source =
      getSourceSheet_();

    const questions =
      parseSourceSheet_(
        source
      );

    if (!questions.length) {

      throw new Error(

        'В исходном листе "' +
        source.getName() +
        '" не найдено ни одного пункта.\n\n' +

        'Проверьте:\n' +
        '1. правильное название листа;\n' +
        '2. наличие строк, начинающихся с "' +
        CFG.SECTION_PREFIX +
        '".'

      );

    }

    /*
     * Записываем вопросы
     * в служебный лист.
     */
    writeQuestionsFast_(
      ss,
      questions
    );

    /*
     * Записываем настройки.
     */
    writeSettingsFast_(
      ss,
      source
    );

    SpreadsheetApp.flush();

    const message =

      'Приложение подготовлено.\n\n' +

      'Исходный лист: ' +
      source.getName() +

      '\nРазделов: ' +
      countSections_(
        questions
      ) +

      '\nПунктов: ' +
      questions.length +

      '\nВремя выполнения: ' +
      (
        Date.now() -
        started
      ) +
      ' мс';

    try {

      SpreadsheetApp
        .getUi()
        .alert(
          message
        );

    } catch (e) {

      console.log(
        message
      );

    }

    return message;

  } finally {

    lock.releaseLock();

  }

}


/* ============================================================
 * СЛУЖЕБНЫЙ ЛИСТ ВОПРОСОВ
 * ============================================================ */

function writeQuestionsFast_(
  ss,
  questions
) {

  let sheet =
    ss.getSheetByName(
      CFG.QUESTIONS_SHEET_NAME
    );

  if (!sheet) {

    sheet =
      ss.insertSheet(
        CFG.QUESTIONS_SHEET_NAME
      );

  }

  const oldRows =
    sheet.getLastRow();

  const oldCols =
    sheet.getLastColumn();

  /*
   * Не используем sheet.clear(),
   * чтобы не выполнять лишние операции.
   */
  if (
    oldRows > 0 &&
    oldCols > 0
  ) {

    sheet
      .getRange(
        1,
        1,
        oldRows,
        oldCols
      )
      .clearContent();

  }

  sheet
    .getRange(
      1,
      1,
      1,
      5
    )
    .setValues([[
      'ID',
      'Раздел',
      'Пункт',
      'Строка исходного листа',
      'Количество вариантов ответа'
    ]]);

  const data =
    questions.map(
      function(q) {

        return [
          q.id,
          q.section,
          q.point,
          q.sourceRow,
          q.answerCount || 5
        ];

      }
    );

  sheet
    .getRange(
      2,
      1,
      data.length,
      5
    )
    .setValues(
      data
    );

  /*
   * Скрываем служебный лист.
   */
  sheet.hideSheet();

}


/* ============================================================
 * СЛУЖЕБНЫЙ ЛИСТ НАСТРОЕК
 * ============================================================ */

function writeSettingsFast_(
  ss,
  source
) {

  let sheet =
    ss.getSheetByName(
      CFG.SETTINGS_SHEET_NAME
    );

  if (!sheet) {

    sheet =
      ss.insertSheet(
        CFG.SETTINGS_SHEET_NAME
      );

  }

  const oldRows =
    sheet.getLastRow();

  const oldCols =
    sheet.getLastColumn();

  if (
    oldRows > 0 &&
    oldCols > 0
  ) {

    sheet
      .getRange(
        1,
        1,
        oldRows,
        oldCols
      )
      .clearContent();

  }

  sheet
    .getRange(
      1,
      1,
      4,
      2
    )
    .setValues([

      [
        'Параметр',
        'Значение'
      ],

      [
        'Источник',
        source.getName()
      ],

      [
        'Объект',
        source.getName()
      ],

      [
        'Дата обновления',
        new Date()
      ]

    ]);

  sheet.hideSheet();

}


/* ============================================================
 * ДАННЫЕ ДЛЯ INDEX.HTML
 * ============================================================ */

/**
 * Эта функция вызывается Index.html.
 */
function getAppData() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  /*
   * Каждый раз перечитываем исходный чек-лист.
   * Это важно: количество вариантов ответа находится
   * непосредственно в 4-м столбце исходного листа и
   * не должно зависеть от старого кэша _AUDIT_QUESTIONS.
   */
  const source =
    getSourceSheet_();

  let questions =
    parseSourceSheet_(
      source
    );

  if (!questions.length) {

    throw new Error(
      'В исходном листе нет пунктов.'
    );

  }

  /* Обновляем служебный лист актуальными данными. */
  writeQuestionsFast_(
    ss,
    questions
  );

  const settings =
    readSettings_(
      ss
    );

  /*
   * Формируем разделы.
   */
  const sectionMap = {};

  questions.forEach(
    function(q) {

      if (
        !sectionMap[
          q.section
        ]
      ) {

        sectionMap[
          q.section
        ] = {

          name:
            q.section,

          /*
           * Ответственный вводится
           * непосредственно в форме.
           */
          responsible:
            '',

          items:
            []

        };

      }

      sectionMap[
        q.section
      ]
      .items
      .push({

        id:
          q.id,

        point:
          q.point,

        answers:
          Number(q.answerCount) === 3
            ? ['Да', 'Нет', 'Пропущено']
            : CFG.ANSWERS

      });

    }
  );

  return {

    title:
      'Аудит — ' +
      (
        settings.object ||
        ''
      ),

    objectNameDefault:
      settings.object ||
      '',

    dateDefault:
      Utilities.formatDate(

        new Date(),

        Session.getScriptTimeZone(),

        'yyyy-MM-dd'

      ),

    answers:
      CFG.ANSWERS,

    negativeAnswers:
      CFG.NEGATIVE_ANSWERS,

    maxPhotos:
      CFG.MAX_PHOTOS,

    sections:
      Object.keys(
        sectionMap
      )
      .map(
        function(key) {

          return sectionMap[
            key
          ];

        }
      )

  };

}


/* ============================================================
 * ЧТЕНИЕ ВОПРОСОВ
 * ============================================================ */

function readQuestions_(ss) {

  const sheet =
    ss.getSheetByName(
      CFG.QUESTIONS_SHEET_NAME
    );

  if (
    !sheet ||
    sheet.getLastRow() < 2
  ) {

    return [];

  }

  const count =
    sheet.getLastRow() - 1;

  return sheet

    .getRange(
      2,
      1,
      count,
      5
    )

    .getDisplayValues()

    .filter(
      function(row) {

        return (
          row[0] &&
          row[2]
        );

      }
    )

    .map(
      function(row) {

        return {

          id:
            row[0],

          section:
            row[1],

          point:
            row[2],

          sourceRow:
            row[3],

          answerCount:
            Number(row[4]) === 3 ? 3 : 5

        };

      }
    );

}


/* ============================================================
 * ЧТЕНИЕ НАСТРОЕК
 * ============================================================ */

function readSettings_(ss) {

  const sheet =
    ss.getSheetByName(
      CFG.SETTINGS_SHEET_NAME
    );

  if (
    !sheet ||
    sheet.getLastRow() < 2
  ) {

    return {};

  }

  const values =
    sheet
      .getRange(
        1,
        1,
        sheet.getLastRow(),
        2
      )
      .getDisplayValues();

  const obj = {};

  values
    .slice(1)
    .forEach(
      function(row) {

        obj[
          row[0]
        ] =
          row[1];

      }
    );

  return {

    object:
      obj['Объект'] ||
      ''

  };

}


/* ============================================================
 * ОТПРАВКА АУДИТА
 * ============================================================ */

/**
 * Эта функция вызывается только после
 * нажатия "Отправить проверку".
 *
 * Именно здесь создаются:
 *
 * - отдельный файл отчёта;
 * - папка фотографий;
 * - файлы фотографий.
 */

/* ============================================================
 * ЧЕРНОВИК АУДИТА — СЕРВЕРНОЕ ХРАНЕНИЕ
 * ============================================================ */

const AUDIT_DRAFT_FOLDER_NAME = 'Аудит_Черновики';
const AUDIT_DRAFT_PHOTO_FOLDER_NAME = 'Фото_черновиков';

function getAuditDraftFolder_() {
  const folders = DriveApp.getFoldersByName(AUDIT_DRAFT_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(AUDIT_DRAFT_FOLDER_NAME);
}

function getAuditDraftPhotoRoot_() {
  const folders = DriveApp.getFoldersByName(AUDIT_DRAFT_PHOTO_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(AUDIT_DRAFT_PHOTO_FOLDER_NAME);
}

function getAuditDraftJsonFile_() {
  const folder = getAuditDraftFolder_();
  const files = folder.getFilesByName('current.json');
  if (files.hasNext()) return files.next();
  return folder.createFile('current.json', '{}', MimeType.PLAIN_TEXT);
}

function loadAuditDraft() {
  const file = getAuditDraftJsonFile_();

  try {
    const text = file.getBlob().getDataAsString('UTF-8');
    const draft = JSON.parse(text || '{}');

    if (!draft || !draft.savedAt) {
      return { exists: false };
    }

    return {
      exists: true,
      draft: draft
    };
  } catch (e) {
    return { exists: false };
  }
}

function saveAuditDraft(payload) {
  payload = payload || {};

  const draft = {
    version: 1,
    savedAt: new Date().toISOString(),
    objectName: String(payload.objectName || ''),
    auditDate: String(payload.auditDate || ''),
    sections: payload.sections || []
  };

  const file = getAuditDraftJsonFile_();
  file.setContent(JSON.stringify(draft));

  return {
    ok: true,
    savedAt: draft.savedAt
  };
}

function saveAuditDraftPhoto(photo) {
  if (!photo || !photo.data) {
    throw new Error('Фотография для черновика не получена.');
  }

  const root = getAuditDraftPhotoRoot_();
  const draftFolder = getOrCreateSubfolder_(root, 'current');

  const safeQuestionId = sanitizeFileName_(photo.questionId || 'question');
  const safeName = sanitizeFileName_(photo.name || 'photo.jpg');

  // Удаляем старый файл с тем же draftPhotoId, если он передан.
  if (photo.replaceId) {
    const oldFiles = draftFolder.getFilesByName('photo_' + photo.replaceId + '_' + safeName);
    while (oldFiles.hasNext()) oldFiles.next().setTrashed(true);
  }

  const photoId = Utilities.getUuid();
  const bytes = Utilities.base64Decode(String(photo.data));
  const blob = Utilities.newBlob(
    bytes,
    photo.mimeType || 'image/jpeg',
    safeName
  );

  const file = draftFolder.createFile(blob);
  file.setName('photo_' + photoId + '_' + safeQuestionId + '_' + safeName);

  return {
    ok: true,
    photoId: photoId,
    fileId: file.getId(),
    name: safeName,
    mimeType: file.getMimeType()
  };
}

function loadAuditDraftPhoto(photoId) {
  if (!photoId) throw new Error('Не указан идентификатор фотографии.');

  const root = getAuditDraftPhotoRoot_();
  const folder = root.getFoldersByName('current');

  if (!folder.hasNext()) throw new Error('Папка черновика не найдена.');

  const files = folder.next().getFiles();

  while (files.hasNext()) {
    const file = files.next();

    if (file.getName().indexOf('photo_' + photoId + '_') === 0) {
      const blob = file.getBlob();
      return {
        ok: true,
        photoId: photoId,
        name: file.getName().replace(
          new RegExp('^photo_' + photoId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '_[^_]+_'),
          ''
        ),
        mimeType: blob.getContentType(),
        data: Utilities.base64Encode(blob.getBytes())
      };
    }
  }

  throw new Error('Фотография черновика не найдена.');
}

function deleteAuditDraftPhoto(photoId) {
  if (!photoId) return { ok: true };

  const root = getAuditDraftPhotoRoot_();
  const folder = root.getFoldersByName('current');

  if (!folder.hasNext()) return { ok: true };

  const files = folder.next().getFiles();

  while (files.hasNext()) {
    const file = files.next();

    if (file.getName().indexOf('photo_' + photoId + '_') === 0) {
      file.setTrashed(true);
    }
  }

  return { ok: true };
}

function removeAuditDraft() {
  const folder = getAuditDraftFolder_();
  const files = folder.getFilesByName('current.json');

  while (files.hasNext()) {
    files.next().setTrashed(true);
  }

  const photoRoot = getAuditDraftPhotoRoot_();
  const photoFolders = photoRoot.getFoldersByName('current');

  while (photoFolders.hasNext()) {
    photoFolders.next().setTrashed(true);
  }

  return { ok: true };
}


function htmlEscapeV2_(value) {
  return String(value || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : (e && e.parameter ? String(e.parameter.payload || '') : '');
    if (!raw) throw new Error('Данные проверки не получены.');
    let payload;
    try { payload = JSON.parse(raw); }
    catch (jsonError) { payload = JSON.parse(String(e && e.parameter && e.parameter.payload || 'null')); }
    if (!payload) throw new Error('Данные проверки не получены.');

    const action = String(payload.action || '').trim();
    if (action === 'uploadPhotoBatch') {
      return ContentService.createTextOutput(JSON.stringify(uploadAuditPhotoBatchV2_(payload))).setMimeType(ContentService.MimeType.JSON);
    }
    if (action === 'uploadPhoto') {
      return ContentService.createTextOutput(JSON.stringify(uploadAuditPhotoV2_(payload))).setMimeType(ContentService.MimeType.JSON);
    }
    if (action === 'uploadFinalize') {
      const result = finalizeAuditUploadV2_(payload);
      const url = ScriptApp.getService().getUrl();
      const auditId=htmlEscapeV2_(result.auditId);
      return HtmlService.createHtmlOutput(
        '<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><style>body{font-family:Arial,sans-serif;padding:28px;font-size:18px;text-align:center;line-height:1.5} .box{max-width:560px;margin:30px auto} .bar{height:14px;background:#e5e7eb;border-radius:8px;overflow:hidden;margin:18px 0}.fill{height:100%;width:8%;background:#4f46e5;transition:width .3s}.muted{color:#666;font-size:15px}</style></head>' +
        '<body><div class="box"><h2>Проверка успешно сохранена</h2>' +
        '<p><b>Отчёт создан</b></p><div class="bar"><div id="fill" class="fill"></div></div>' +
        '<p id="status">Фотографии загружены. Формируем приложение PDF…</p>' +
        '<p class="muted" id="detail">Не закрывайте эту страницу.</p>' +
        '<p><a id="report" style="display:inline-block;padding:14px 24px;background:#4f46e5;color:#fff;text-decoration:none;border-radius:10px" href="' + htmlEscapeV2_(result.reportXlsxUrl || result.reportUrl) + '" target="_blank">Скачать отчёт Excel</a></p>' +
        '<p id="appWrap" style="display:none"><a id="app" style="display:inline-block;padding:14px 24px;background:#16803c;color:#fff;text-decoration:none;border-radius:10px" href="#" target="_blank">Скачать приложение PDF</a></p>' +
        '<p id="draftWrap" style="display:none"><a id="draft" style="display:inline-block;padding:14px 24px;background:#7c3aed;color:#fff;text-decoration:none;border-radius:10px" href="#" target="_blank">Скачать черновик формы</a></p>' +
        '<p><a id="folder" style="display:inline-block;padding:14px 24px;background:#374151;color:#fff;text-decoration:none;border-radius:10px" href="' + htmlEscapeV2_(result.auditFolderUrl || '') + '" target="_blank">Открыть папку аудита на Google Диске</a></p>' +
        '<p><a style="display:inline-block;padding:14px 24px;background:#6b7280;color:#fff;text-decoration:none;border-radius:10px" href="' + url + '">Вернуться к аудиту</a></p>' +
        '</div><script>' +
        'var id="' + auditId + '";' +
        'function poll(){google.script.run.withSuccessHandler(function(s){var f=document.getElementById("fill"),st=document.getElementById("status"),d=document.getElementById("detail");if(s.status==="processing"){var pct=s.total?Math.max(8,Math.round(s.current/s.total*92)):12;f.style.width=pct+"%";st.textContent="Формируем приложение PDF…";d.textContent="Пунктов обработано: "+s.current+" / "+s.total;setTimeout(poll,1500);}else if(s.status==="done"){f.style.width="100%";st.textContent="Готово. Excel-отчёт и приложение PDF созданы.";d.textContent="Фотографии: ' + String(result.uploadedPhotos) + '";document.getElementById("report").href=s.reportXlsxUrl||document.getElementById("report").href;document.getElementById("app").href=s.appendixPdfUrl;document.getElementById("draft").href=s.draftUrl||"#";document.getElementById("appWrap").style.display="block";document.getElementById("draftWrap").style.display="block";document.getElementById("folder").href=s.auditFolderUrl||document.getElementById("folder").href;}else if(s.status==="error"){f.style.width="100%";st.textContent="Отчёт создан, но приложение PDF не удалось сформировать.";d.textContent=s.error||"Неизвестная ошибка.";}else{setTimeout(poll,1500);}}).getAuditAppendixStatusV2(id);}poll();</script></body></html>'
      ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const result = submitAudit(payload);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(error && error.message ? error.message : error)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function submitAudit(
  payload
) {

  if (!payload) {

    throw new Error(
      'Данные проверки не получены.'
    );

  }

  const objectName =
    String(
      payload.objectName ||
      ''
    ).trim();

  const auditDate =
    String(
      payload.auditDate ||
      ''
    ).trim();

  if (!objectName) {

    throw new Error(
      'Укажите наименование объекта.'
    );

  }

  if (!auditDate) {

    throw new Error(
      'Укажите дату проведения.'
    );

  }

  /*
   * Время выгрузки фиксируем один раз для всей операции.
   * Благодаря этому повторная выгрузка одного и того же аудита
   * создаёт отдельный файл отчёта и отдельное Приложение.
   */
  const uploadAt = new Date();
  const uploadStamp =
    Utilities.formatDate(
      uploadAt,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd_HH-mm-ss-SSS'
    );

  /*
   * Создаём новый отдельный файл отчёта для КАЖДОЙ выгрузки.
   */
  const report =
    getOrCreateReport_(
      objectName,
      auditDate,
      uploadStamp
    );

  /*
   * Создаём папку фотографий.
   */
  const photoRoot =
    getOrCreateFolder_(
      CFG.PHOTO_FOLDER_NAME
    );

  const photoFolder =
    getOrCreateSubfolder_(
      photoRoot,

      'архив фото_' +
      sanitizeFileName_(auditDate) +
      '_' +
      sanitizeFileName_(objectName)

    );

  /*
   * Загружаем актуальные веса из исходного чек-листа.
   * Это позволяет считать баллы непосредственно по 3-му столбцу
   * исходного чек-листа, не полагаясь на данные браузера.
   */
  const sourceQuestions =
    parseSourceSheet_(
      getSourceSheet_()
    );

  const questionMap = {};

  sourceQuestions.forEach(function(q) {
    questionMap[q.id] = q;
  });

  const rows = [];

  const now =
    new Date();

  /*
   * Перебираем разделы.
   */
  (
    payload.sections ||
    []
  )
  .forEach(
    function(section) {

      /*
       * Ответственный относится
       * ко всему разделу.
       */
      const responsible =
        String(
          section.responsible ||
          ''
        ).trim();

      /*
       * Перебираем пункты.
       */
      (
        section.items ||
        []
      )
      .forEach(
        function(item) {

          const answer =
            String(
              item.answer ||
              ''
            ).trim();

          const comment =
            String(
              item.comment ||
              ''
            ).trim();

          /*
           * Пустой пункт не сохраняем.
           */
          if (!answer) {
            return;
          }

          /*
           * Для нарушений комментарий обязателен.
           */
          if (

            CFG
              .NEGATIVE_ANSWERS
              .indexOf(
                answer
              ) >= 0 &&

            !comment

          ) {

            throw new Error(

              'Для пункта "' +
              item.point +
              '" необходимо описание нарушения.'

            );

          }

          /*
           * Список ссылок на фотографии.
           */
          const photoLinks = [];

          const photos =
            item.photos ||
            [];

          /*
           * ВАЖНО:
           *
           * Каждый элемент photos создаётся
           * отдельным File.
           *
           * Поэтому:
           *
           * Фото 1
           * Фото 2
           * Фото 3
           *
           * не перезаписывают друг друга.
           */
          photos
            .slice(
              0,
              CFG.MAX_PHOTOS
            )
            .forEach(
              function(photo, index) {

                if (
                  !photo ||
                  !photo.data
                ) {

                  return;

                }

                const blob =
                  Utilities.newBlob(

                    Utilities.base64Decode(
                      photo.data
                    ),

                    photo.mimeType ||
                    'image/jpeg',

                    photo.name ||
                    (
                      'Фото_' +
                      (index + 1) +
                      '.jpg'
                    )

                  );

                const file =
                  photoFolder
                    .createFile(
                      blob
                    );

                photoLinks.push(
                  file.getUrl()
                );

              }
            );

          /*
           * Одна строка = один вопрос.
           */
          rows.push([

            now,

            auditDate,

            objectName,

            section.name,

            responsible,

            item.point,

            answer,

            comment,

            photoLinks.join(
              '\n'
            )

          ]);

        }
      );

    }
  );

  if (!rows.length) {

    throw new Error(
      'Нет заполненных пунктов.'
    );

  }

  /*
   * Записываем ответы одним setValues().
   *
   * Это быстрее, чем записывать
   * каждую строку отдельно.
   */
  const sheet =
    report.sheet;

  const startRow =
    Math.max(
      sheet.getLastRow() + 1,
      2
    );

  sheet
    .getRange(
      startRow,
      1,
      rows.length,
      9
    )
    .setValues(
      rows
    );

  /*
   * Форматирование только шапки,
   * а не всей таблицы.
   */
  sheet.setFrozenRows(1);

  sheet
    .getRange(
      1,
      1,
      1,
      9
    )
    .setFontWeight(
      'bold'
    )
    .setWrap(
      true
    );

  sheet.setColumnWidth(
    9,
    320
  );

  /*
   * Создаём/обновляем вкладку «Результат» по приложенному шаблону.
   * В ней по каждому ответственному считаются количества ответов
   * и набранные баллы по весам исходного чек-листа.
   */
  createAuditResultSheet_(
    report.file.getId(),
    payload,
    questionMap
  );
  SpreadsheetApp.flush();
  const reportPdf = createAuditReportPdfV2_(report.file,payload.objectName,payload.auditDate,uploadStamp,getResultsReportFolder_());
  const appendixPdf = createAuditApplicationDocx_(payload, uploadStamp);

  return {
    ok: true,
    savedRows: rows.length,
    appendixDocxName: appendixPdf.name,
    appendixDocxUrl: appendixPdf.url,
    appendixPdfName: appendixPdf.name,
    appendixPdfUrl: appendixPdf.url,
    reportName: report.name,
    reportUrl: report.file.getUrl(),
    reportPdfName: reportPdf.name,
    reportPdfUrl: reportPdf.downloadUrl
  };

}


/* FIX: submitAudit() closing brace restored before getOrCreateReport_(). */
/* ============================================================
 * СОЗДАНИЕ ОТДЕЛЬНОГО ФАЙЛА ОТЧЁТА
 * ============================================================ */

function getOrCreateReport_(
  objectName,
  auditDate,
  uploadStamp,
  storageMode
) {

  /*
   * Для КАЖДОЙ выгрузки создаём новый файл.
   * В имени учитываем объект, дату аудита и точное время выгрузки.
   * Это исключает повторное накопление одинаковых ответов
   * в одном и том же файле при повторной выгрузке.
   */
  const name =
    'Отчёт_по_аудиту_' +
    sanitizeFileName_(objectName) +
    '_' +
    sanitizeFileName_(auditDate) +
    '_' +
    sanitizeFileName_(uploadStamp || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss-SSS')) + '_' + Utilities.getUuid().slice(0, 8);

  const folder = getResultsReportFolder_(storageMode);

  /*
   * Всегда создаём НОВУЮ Google Таблицу.
   */
  const created =
    SpreadsheetApp.create(
      name
    );

  const file =
    DriveApp.getFileById(
      created.getId()
    );

  file.moveTo(folder);

  const reportSS =
    SpreadsheetApp.openById(
      file.getId()
    );

  let sheet =
    reportSS.getSheetByName(
      CFG.RESPONSE_SHEET_NAME
    );

  if (!sheet) {
    sheet =
      reportSS.getSheets()[0];
    sheet.setName(
      CFG.RESPONSE_SHEET_NAME
    );
  }

  if (sheet.getLastRow() === 0) {
    sheet
      .getRange(1, 1, 1, 9)
      .setValues([[
        'Дата/время',
        'Дата проведения',
        'Наименование объекта',
        'Раздел',
        'Ответственный',
        'Пункт',
        'Ответ',
        'Комментарий',
        'Фото'
      ]]);
  }

  return {
    file: file,
    sheet: sheet,
    name: name
  };
}


/* ============================================================
 * GOOGLE DRIVE — ПАПКИ
 * ============================================================ */

function getResultsRootFolder_(storageMode) {
  const mode = String(storageMode || CFG.RESULT_STORAGE_MODE || 'SCRIPT_ACCOUNT').toUpperCase();

  if (mode === 'SECOND_ACCOUNT') {
    const id = String(CFG.SECOND_RESULTS_ROOT_FOLDER_ID || '').trim();
    if (!id) {
      throw new Error('Не указан SECOND_RESULTS_ROOT_FOLDER_ID.');
    }
    try {
      return DriveApp.getFolderById(id);
    } catch (e) {
      throw new Error(
        'Нет доступа к папке второго аккаунта. Откройте доступ к папке для аккаунта, под которым выполняется Apps Script, с правом «Редактор». ' +
        String(e && e.message ? e.message : e)
      );
    }
  }

  return DriveApp.getRootFolder();
}

function getResultsReportFolder_(storageMode) {
  const root = getResultsRootFolder_(storageMode);
  return getOrCreateSubfolder_(root, CFG.REPORT_FOLDER_NAME);
}

function getAuditFolderV3_(auditId, objectName, auditDate, storageMode, startedAt) {
  const root = getResultsReportFolder_(storageMode);
  let dt = startedAt ? new Date(startedAt) : null;
  if (!dt || isNaN(dt.getTime())) dt = new Date();
  const stamp = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd.MM.yyyy_HH-mm-ss');
  const base = stamp + '_' + sanitizeFileName_(objectName || 'объект');
  const suffix = sanitizeFileName_(auditId || Utilities.getUuid()).slice(-12);
  let folders = root.getFoldersByName(base);
  if (folders.hasNext()) {
    const existing = folders.next();
    try {
      if (String(existing.getDescription() || '') === String(auditId || '')) return existing;
    } catch (e) {}
    const unique = base + '_' + suffix;
    folders = root.getFoldersByName(unique);
    if (folders.hasNext()) return folders.next();
    const folder = root.createFolder(unique);
    try { folder.setDescription(String(auditId || '')); } catch (e) {}
    return folder;
  }
  const folder = root.createFolder(base);
  try { folder.setDescription(String(auditId || '')); } catch (e) {}
  return folder;
}

function getResultsPhotoFolder_(storageMode) {
  const root = getResultsRootFolder_(storageMode);
  return getOrCreateSubfolder_(root, CFG.PHOTO_FOLDER_NAME);
}

function getOrCreateFolder_(
  name
) {

  const folders =
    DriveApp.getFoldersByName(
      name
    );

  if (
    folders.hasNext()
  ) {

    return folders.next();

  }

  return DriveApp.createFolder(
    name
  );

}


/**
 * Создаёт подпапку внутри указанной папки.
 */
function getOrCreateSubfolder_(
  parent,
  name
) {

  const folders =
    parent.getFoldersByName(
      name
    );

  if (
    folders.hasNext()
  ) {

    return folders.next();

  }

  return parent.createFolder(
    name
  );

}


/* ============================================================
 * ОЧИСТКА ИМЕНИ ФАЙЛА
 * ============================================================ */

function sanitizeFileName_(
  name
) {

  return String(name)

    .replace(
      /[\\\/:*?"<>|#%{}~&]/g,
      '_'
    )

    .replace(
      /\s+/g,
      ' '
    )

    .trim()

    .substring(
      0,
      150
    );

}


/* ============================================================
 * URL WEB APP
 * ============================================================ */

function showWebAppUrl() {

  const url =
    ScriptApp
      .getService()
      .getUrl();

  try {

    SpreadsheetApp
      .getUi()
      .alert(

        url

          ? 'URL приложения:\n\n' +
            url

          : 'Приложение ещё не опубликовано как Web App.'

      );

  } catch (e) {

    console.log(
      url
    );

  }

}



/**
 * Создаёт отдельное PDF-приложение по нарушениям.
 * PDF сохраняется В ТУ ЖЕ ПАПКУ, что и основной отчёт:
 * CFG.REPORT_FOLDER_NAME.
 */

/* ============================================================
 * ИТОГОВАЯ ДЕТАЛИЗАЦИЯ
 * ============================================================ */

/**
 * Создаёт вкладку «Результат» по структуре приложенного файла.
 *
 * Расчёт:
 *   Да                     = вес пункта
 *   Нет                    = 0
 *   Незначительные недочёты = 60
 *   Значительные нарушения = 30
 *   Пропущено              = не участвует в расчёте
 *
 * Количество ответов считается отдельно по каждому ответственному.
 */
function createAuditResultSheet_(
  reportFileId,
  payload,
  questionMap
) {

  const reportSS =
    SpreadsheetApp.openById(
      reportFileId
    );

  let resultSheet =
    reportSS.getSheetByName(
      CFG.RESULT_SHEET_NAME
    );

  if (!resultSheet) {
    resultSheet =
      reportSS.insertSheet(
        CFG.RESULT_SHEET_NAME
      );
  }

  resultSheet.clear();

  const headers = [
    '',
    'Ответственный',
    'Да',
    'Незначительные недочёты',
    'Значительные нарушения',
    'Нет',
    'Пропущено',
    'Набрано баллов',
    'Всего возможных балов',
    'Доля ответственности от всего аудита, %',
    '% выполнения от своего объема',
    '% от общего объема'
  ];

  resultSheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues([headers]);

  const byResponsible = {};

  /*
   * Один ответственный = один раздел в текущей форме.
   */
  (payload.sections || []).forEach(function(section) {

    // Исключённые пользователем разделы в результат не попадают.
    if (section && section.skipped) {
      return;
    }

    const responsible =
      String(
        section.responsible || ''
      ).trim() ||
      'Не указан';

    if (!byResponsible[responsible]) {
      byResponsible[responsible] = {
        responsible: responsible,
        yes: 0,
        minor: 0,
        major: 0,
        no: 0,
        skipped: 0,
        score: 0,
        possible: 0,
        participated: 0
      };
    }

    (section.questions || section.items || []).forEach(function(item) {

      const answer =
        String(
          item.answer || ''
        ).trim();

      if (!answer) {
        return;
      }

      /*
       * Ищем вес по ID вопроса.
       */
      const q =
        questionMap[
          String(item.id || '')
        ];

      let weight = 0;

      if (q) {
        weight =
          Number(
            String(q.weight || '')
              .replace(',', '.')
              .trim()
          );

        if (!isFinite(weight)) {
          weight = 0;
        }
      }

      const stat =
        byResponsible[responsible];

      /*
       * Количество каждого типа ответа.
       */
      if (answer === 'Да') {
        stat.yes++;
      } else if (
        answer === 'Незначительные недочёты'
      ) {
        stat.minor++;
      } else if (
        answer === 'Значительные нарушения'
      ) {
        stat.major++;
      } else if (
        answer === 'Нет'
      ) {
        stat.no++;
      } else if (
        answer === 'Пропущено'
      ) {
        stat.skipped++;
      }

      /*
       * Пропущено не участвует в подсчёте.
       */
      if (answer === 'Пропущено') {
        return;
      }

      stat.participated++;
      stat.possible += weight;

      if (answer === 'Да') {
        stat.score += weight;
      } else if (
        answer === 'Незначительные недочёты'
      ) {
        stat.score += 60;
      } else if (
        answer === 'Значительные нарушения'
      ) {
        stat.score += 30;
      } else if (
        answer === 'Нет'
      ) {
        stat.score += 0;
      }

    });

  });

  const stats =
    Object.keys(byResponsible)
      .map(function(key) {
        return byResponsible[key];
      });

  /*
   * Сначала строки ответственных.
   */
  const output = [];

  stats.forEach(function(stat, index) {

    output.push([
      index + 1,
      stat.responsible,
      stat.yes,
      stat.minor,
      stat.major,
      stat.no,
      stat.skipped,
      stat.score,
      stat.possible,
      '',
      '',
      ''
    ]);

  });

  /*
   * Строка ИТОГО.
   */
  output.push([
    '',
    'ВСЕГО:',
    stats.reduce(function(sum, x) { return sum + x.yes; }, 0),
    stats.reduce(function(sum, x) { return sum + x.minor; }, 0),
    stats.reduce(function(sum, x) { return sum + x.major; }, 0),
    stats.reduce(function(sum, x) { return sum + x.no; }, 0),
    stats.reduce(function(sum, x) { return sum + x.skipped; }, 0),
    stats.reduce(function(sum, x) { return sum + x.score; }, 0),
    stats.reduce(function(sum, x) { return sum + x.possible; }, 0),
    '',
    '',
    ''
  ]);

  if (output.length) {
    resultSheet
      .getRange(
        2,
        1,
        output.length,
        headers.length
      )
      .setValues(output);
  }

  const totalPossible =
    stats.reduce(function(sum, x) {
      return sum + x.possible;
    }, 0);

  const totalScore =
    stats.reduce(function(sum, x) {
      return sum + x.score;
    }, 0);

  /*
   * Формулы:
   * J — доля ответственности = возможные баллы ответственного /
   *                           возможные баллы всего аудита.
   * K — выполнение от своего объёма = набрано / возможные.
   * L — вклад в общий результат = набрано / возможные всего.
   */
  stats.forEach(function(stat, index) {

    const row = index + 2;

    resultSheet
      .getRange(row, 10)
      .setValue(
        totalPossible
          ? stat.possible / totalPossible
          : 0
      );

    resultSheet
      .getRange(row, 11)
      .setValue(
        stat.possible
          ? stat.score / stat.possible
          : 0
      );

    resultSheet
      .getRange(row, 12)
      .setValue(
        totalPossible
          ? stat.score / totalPossible
          : 0
      );

  });

  const totalRow =
    stats.length + 2;

  resultSheet
    .getRange(totalRow, 10)
    .setValue(
      totalPossible
        ? 1
        : 0
    );

  resultSheet
    .getRange(totalRow, 11)
    .setValue(
      totalPossible
        ? totalScore / totalPossible
        : 0
    );

  resultSheet
    .getRange(totalRow, 12)
    .setValue(
      totalPossible
        ? totalScore / totalPossible
        : 0
    );

  /*
   * Оформление по приложенному шаблону.
   */
  resultSheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setFontWeight('bold')
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  if (output.length) {
    resultSheet
      .getRange(
        2,
        1,
        output.length,
        headers.length
      )
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }

  /*
   * Проценты.
   */
  if (stats.length) {
    resultSheet
      .getRange(
        2,
        10,
        stats.length + 1,
        3
      )
      .setNumberFormat('0%');
  }

  /*
   * Округление процентов первой таблицы до целых с гарантией
   * итоговой суммы 100%.
   *
   * Используем метод наибольших остатков:
   * сначала берём целую часть каждого процента,
   * затем распределяем недостающие процентные пункты
   * по самым большим дробным остаткам.
   */
  if (stats.length) {
    const rawShares = stats.map(function(stat) {
      return totalPossible
        ? stat.possible / totalPossible * 100
        : 0;
    });

    const roundedShares = rawShares.map(function(value) {
      return Math.floor(value);
    });

    let remaining =
      100 -
      roundedShares.reduce(function(sum, value) {
        return sum + value;
      }, 0);

    const order = rawShares
      .map(function(value, index) {
        return {
          index: index,
          remainder: value - Math.floor(value)
        };
      })
      .sort(function(a, b) {
        return b.remainder - a.remainder;
      });

    for (let i = 0; i < remaining && i < order.length; i++) {
      roundedShares[order[i].index]++;
    }

    roundedShares.forEach(function(value, index) {
      resultSheet
        .getRange(index + 2, 10)
        .setValue(value / 100);
    });

    resultSheet
      .getRange(
        2,
        10,
        stats.length,
        1
      )
      .setNumberFormat('0%');
  }

  /*
   * Размеры колонок максимально близки к приложенному шаблону.
   */
  const widths = [
    35, 180, 75, 170, 170, 85, 100,
    110, 160, 180, 180, 150
  ];

  widths.forEach(function(width, index) {
    resultSheet.setColumnWidth(
      index + 1,
      width
    );
  });

  resultSheet.setFrozenRows(1);
  resultSheet.getRange(1, 1, totalRow, headers.length)
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true
    );

  /*
   * Закрепляем итоговую строку визуально.
   */
  resultSheet
    .getRange(
      totalRow,
      1,
      1,
      headers.length
    )
    .setFontWeight('bold');

  /*
   * Перенос длинных заголовков.
   */
  resultSheet
    .getRange(
      1,
      1,
      totalRow,
      headers.length
    )
    .setWrap(true);

  /*
   * ============================================================
   * ВТОРАЯ ТАБЛИЦА — ВЫПОЛНЕНИЕ ПО РАЗДЕЛАМ
   * ============================================================
   */
  const sectionSummary = {};

  (payload.sections || []).forEach(function(section) {
    // Во второй таблице также показываем только включённые разделы.
    if (section && section.skipped) {
      return;
    }

    const sectionName =
      String(section.name || '').trim() || 'Без раздела';

    const responsible =
      String(section.responsible || '').trim() || 'Не указан';

    if (!sectionSummary[sectionName]) {
      sectionSummary[sectionName] = {
        responsible: responsible,
        score: 0,
        possible: 0
      };
    } else if (
      sectionSummary[sectionName].responsible === 'Не указан' &&
      responsible !== 'Не указан'
    ) {
      sectionSummary[sectionName].responsible = responsible;
    }

    (section.questions || section.items || []).forEach(function(item) {
      const answer =
        String(item.answer || '').trim();

      if (!answer || answer === 'Пропущено') {
        return;
      }

      const q =
        questionMap[String(item.id || '')];

      let weight = 0;

      if (q) {
        weight =
          Number(
            String(q.weight || '')
              .replace(',', '.')
              .trim()
          );

        if (!isFinite(weight)) {
          weight = 0;
        }
      }

      const stat =
        sectionSummary[sectionName];

      stat.possible += weight;

      if (answer === 'Да') {
        stat.score += weight;
      } else if (answer === 'Незначительные недочёты') {
        stat.score += 60;
      } else if (answer === 'Значительные нарушения') {
        stat.score += 30;
      }
    });
  });

  const sectionRows =
    Object.keys(sectionSummary).map(function(sectionName) {
      const stat =
        sectionSummary[sectionName];

      return [
        sectionName,
        stat.responsible,
        stat.possible
          ? stat.score / stat.possible
          : 0
      ];
    });

  const secondTableStartRow =
    resultSheet.getLastRow() + 3;

  resultSheet
    .getRange(
      secondTableStartRow,
      1,
      1,
      3
    )
    .setValues([[
      'Наименование раздела',
      'Ответственный за раздел',
      'Процент выполнения'
    ]])
    .setFontWeight('bold')
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  if (sectionRows.length) {
    resultSheet
      .getRange(
        secondTableStartRow + 1,
        1,
        sectionRows.length,
        3
      )
      .setValues(sectionRows)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Процент каждого раздела — его собственное выполнение,
    // а не доля в сумме разделов. Значения уже записаны как score / possible.
    resultSheet
      .getRange(
        secondTableStartRow + 1,
        3,
        sectionRows.length,
        1
      )
      .setNumberFormat('0%');
  }

  resultSheet
    .getRange(
      secondTableStartRow,
      1,
      sectionRows.length + 1,
      3
    )
    .setBorder(true, true, true, true, true, true);

  /*
   * Ширина колонок второй таблицы.
   * Названия разделов и фамилии ответственных
   * должны отображаться полностью/без сильного сжатия.
   */
  resultSheet.setColumnWidth(1, 300);
  resultSheet.setColumnWidth(2, 240);
  resultSheet.setColumnWidth(3, 170);

  return resultSheet;
}



function createAuditApplicationDocx_(payload, uploadStamp) {
  const objectName = String(payload.objectName || 'объект').trim() || 'объект';
  const safeObjectName = sanitizeFileName_(objectName);
  const auditDate = sanitizeFileName_(payload.auditDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  const stamp = sanitizeFileName_(uploadStamp || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss-SSS'));
  const docName = 'Приложение_' + safeObjectName + '_' + auditDate + '_' + stamp;

  const results = [];
  let answered = 0;
  let negativeCount = 0;
  let skippedSections = 0;
  let totalPhotos = 0;

  (payload.sections || []).forEach(function(section) {
    if (section.skipped) {
      skippedSections++;
      return;
    }

    (section.items || []).forEach(function(item) {
      const answer = String(item.answer || '').trim();
      if (!answer) return;

      const isNegative = CFG.NEGATIVE_ANSWERS.indexOf(answer) >= 0;
      const photos = Array.isArray(item.photos) ? item.photos.filter(function(photo) {
        return !!(photo && photo.data);
      }).slice(0, CFG.MAX_PHOTOS) : [];

      answered++;
      if (isNegative) negativeCount++;
      totalPhotos += photos.length;

      results.push({
        section: String(section.name || ''),
        responsible: String(section.responsible || ''),
        point: String(item.point || ''),
        answer: answer,
        comment: String(item.comment || ''),
        photos: photos,
        negative: isNegative
      });
    });
  });

  const doc = DocumentApp.create(docName);
  const body = doc.getBody();
  body.clear();

  // A4 landscape: все пропорции приложения рассчитаны под широкую страницу.
  body.setPageWidth(841.89);
  body.setPageHeight(595.28);
  body.setMarginTop(24);
  body.setMarginBottom(24);
  body.setMarginLeft(26);
  body.setMarginRight(26);

  const title = body.appendParagraph('ПРИЛОЖЕНИЕ К ОТЧЁТУ');
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);
  title.setFontSize(20).setBold(true).setSpacingAfter(1);

  const sub = body.appendParagraph('Результаты аудита');
  sub.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  sub.setBold(true).setFontSize(13).setSpacingAfter(6);

  const info = body.appendTable([
    ['Объект', objectName, 'Дата аудита', String(payload.auditDate || '')],
    ['Пунктов с ответами', String(answered), 'Пунктов с недостатками', String(negativeCount)],
    ['Фотографий', String(totalPhotos), 'Пропущенных разделов', String(skippedSections)]
  ]);
  info.setBorderWidth(1);
  for (let r = 0; r < 3; r++) {
    const c0 = info.getCell(r, 0);
    const c2 = info.getCell(r, 2);
    c0.setBackgroundColor('#E8EEF7');
    c2.setBackgroundColor('#E8EEF7');
    c0.getChild(0).asParagraph().editAsText().setBold(false);
    c2.getChild(0).asParagraph().editAsText().setBold(false);
  }

  body.appendParagraph('').setSpacingAfter(2);
  const heading = body.appendParagraph('РЕЗУЛЬТАТЫ ПРОВЕРКИ');
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  heading.setFontSize(15).setSpacingAfter(4);

  if (!results.length) {
    body.appendParagraph('Заполненных пунктов нет.');
  } else {
    results.forEach(function(item, index) {
      const hasPhotos = item.photos.length > 0;
      // Нарушение с фото начинается на новой странице и весь блок хранится в одной таблице.
      if (item.negative && hasPhotos) body.appendPageBreak();
      appendAuditResultBlockV2_(body, item, index + 1);
    });
  }

  doc.saveAndClose();

  const sourceFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = sourceFile.getAs(MimeType.PDF).setName(docName + '.pdf');
  const folder = getResultsReportFolder_();
  const pdfFile = folder.createFile(pdfBlob);
  try { sourceFile.setTrashed(true); } catch (e) {}

  return {
    name: pdfFile.getName(),
    url: makeDriveDownloadUrlV2_(pdfFile.getId()),
    id: pdfFile.getId(),
    items: results.length,
    negativeItems: negativeCount,
    photos: totalPhotos
  };
}

function appendAuditResultBlockV2_(body, item, index) {
  // Один табличный блок на пункт: текст и все фотографии являются частью одной строки.
  const card = body.appendTable([['', '']]);
  card.setBorderWidth(1);

  const textCell = card.getCell(0, 0);
  const photoCell = card.getCell(0, 1);
  const bg = item.negative ? '#FCE8E6' : '#F8F9FA';
  textCell.setBackgroundColor(bg);
  photoCell.setBackgroundColor(item.negative ? '#FFF4F2' : '#FFFFFF');

  const p = textCell.getChild(0).asParagraph();
  p.appendText('ПУНКТ ' + index).setBold(true).setFontSize(11);

  const sec = textCell.appendParagraph(item.section || '');
  sec.setBold(true).setFontSize(10);

  const resp = textCell.appendParagraph('Ответственный: ' + (item.responsible || 'не указан'));
  resp.setFontSize(9);

  const point = textCell.appendParagraph(item.point || '');
  point.setBold(true).setFontSize(11);

  const ans = textCell.appendParagraph('Ответ: ' + item.answer);
  ans.setBold(true).setFontSize(12);
  if (item.negative) ans.setForegroundColor('#C5221F');
  else if (item.answer === 'Да') ans.setForegroundColor('#137333');
  else if (item.answer === 'Пропущено') ans.setForegroundColor('#B06000');

  const ct = textCell.appendParagraph('Комментарий:');
  ct.setBold(true).setFontSize(9);
  const comment = textCell.appendParagraph(item.comment || '—');
  comment.setFontSize(9);
  if (item.negative) comment.setForegroundColor('#8B0000');

  const phTitle = photoCell.getChild(0).asParagraph();
  phTitle.appendText(item.photos.length ? 'ФОТО: ' + item.photos.length : 'ФОТО: нет').setBold(true).setFontSize(9);

  if (!item.photos.length) return;

  // До четырёх изображений в ряд; уменьшенный размер позволяет держать весь блок пункта на одной странице.
  const cols = Math.min(4, item.photos.length);
  const grid = photoCell.appendTable();
  grid.setBorderWidth(0);
  let row = null;
  item.photos.forEach(function(photo, photoIndex) {
    if (photoIndex % cols === 0) row = grid.appendTableRow();
    const cell = row.appendTableCell();
    try {
      const bytes = Utilities.base64Decode(String(photo.data || ''));
      const blob = Utilities.newBlob(bytes, photo.mimeType || 'image/jpeg', photo.name || ('Фото_' + (photoIndex + 1) + '.jpg'));
      const img = cell.appendImage(blob);
      let w = img.getWidth();
      let h = img.getHeight();
      const maxW = 118;
      const maxH = 118;
      if (w > maxW) { h = Math.round(h * maxW / w); w = maxW; }
      if (h > maxH) { w = Math.round(w * maxH / h); h = maxH; }
      img.setWidth(w).setHeight(h);
      const cap = cell.appendParagraph('Фото ' + (photoIndex + 1));
      cap.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      cap.setFontSize(7);
    } catch (e) {
      const err = cell.appendParagraph('Фото ' + (photoIndex + 1) + ': ошибка');
      err.setFontSize(7).setForegroundColor('#C5221F');
    }
  });
}

function createViolationsPdf_(payload, uploadStamp) {
  const objectName = sanitizeFileName_(payload.objectName || 'объект');
  const auditDate = sanitizeFileName_(
    payload.auditDate ||
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  );

  const pdfName =
    'Приложение к отчету по аудиту_' +
    objectName + '_' +
    auditDate + '_' +
    sanitizeFileName_(uploadStamp || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss-SSS'));

  const violations = [];

  (payload.sections || []).forEach(function(section) {
    (section.items || []).forEach(function(item) {
      const answer = String(item.answer || '').trim();

      if (CFG.NEGATIVE_ANSWERS.indexOf(answer) === -1) return;

      violations.push({
        section: String(section.name || ''),
        responsible: String(section.responsible || ''),
        point: String(item.point || ''),
        answer: answer,
        comment: String(item.comment || ''),
        photos: Array.isArray(item.photos) ? item.photos : []
      });
    });
  });

  const doc = DocumentApp.create(pdfName);
  const body = doc.getBody();
  body.clear();

  body.appendParagraph('ПРИЛОЖЕНИЕ К ОТЧЁТУ ПО АУДИТУ')
    .setHeading(DocumentApp.ParagraphHeading.TITLE);

  body.appendParagraph('Объект: ' + String(payload.objectName || ''));
  body.appendParagraph('Дата проведения: ' + String(payload.auditDate || ''));
  body.appendParagraph('Количество нарушений: ' + violations.length);
  body.appendHorizontalRule();

  if (!violations.length) {
    body.appendParagraph('По результатам проверки нарушений не выявлено.');
  } else {
    violations.forEach(function(v, index) {
      body.appendParagraph('НАРУШЕНИЕ № ' + (index + 1))
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);

      body.appendParagraph('Раздел: ' + v.section);
      body.appendParagraph('Ответственный: ' + (v.responsible || 'не указан'));
      body.appendParagraph('Пункт: ' + v.point);
      body.appendParagraph('Ответ: ' + v.answer);

      body.appendParagraph('Описание нарушения:').setBold(true);
      body.appendParagraph(v.comment || 'Описание не указано.');

      if (v.photos.length) {
        body.appendParagraph('Фотографии:').setBold(true);

        v.photos.forEach(function(photo, photoIndex) {
          try {
            const bytes = Utilities.base64Decode(String(photo.data || ''));
            const blob = Utilities.newBlob(
              bytes,
              photo.mimeType || 'image/jpeg',
              photo.name || ('Фото_' + (photoIndex + 1) + '.jpg')
            );

            body.appendParagraph(
              'Фото ' + (photoIndex + 1) + ': ' + (photo.name || '')
            );

            const image = body.appendImage(blob);

            const maxWidth = 480;
            const maxHeight = 620;
            let width = image.getWidth();
            let height = image.getHeight();

            if (width > maxWidth) {
              height = Math.round(height * maxWidth / width);
              width = maxWidth;
            }

            if (height > maxHeight) {
              width = Math.round(width * maxHeight / height);
              height = maxHeight;
            }

            image.setWidth(width);
            image.setHeight(height);

          } catch (e) {
            body.appendParagraph(
              'Не удалось добавить фото "' +
              String(photo.name || '') + '": ' +
              String(e.message || e)
            );
          }
        });
      } else {
        body.appendParagraph('Фотографии: отсутствуют.');
      }

      body.appendHorizontalRule();
    });
  }

  doc.saveAndClose();

  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs(MimeType.PDF)
    .setName(pdfName + '.pdf');

  // ВАЖНО: та же папка, что используется основным отчётом.
  const reportFolder = getResultsReportFolder_();
  const pdfFile = reportFolder.createFile(pdfBlob);

  try {
    docFile.setTrashed(true);
  } catch (e) {}

  return {
    name: pdfFile.getName(),
    url: pdfFile.getUrl(),
    id: pdfFile.getId(),
    violations: violations.length
  };
}


/* ============================================================
 * V2: ПОЭТАПНАЯ ВЫГРУЗКА ФОТОГРАФИЙ
 * ============================================================ */
function getUploadRootV2_(storageMode) {
  return getOrCreateSubfolder_(
    getResultsRootFolder_(storageMode),
    CFG.UPLOAD_RESULTS_FOLDER_NAME
  );
}

function getUploadSessionFolderV2_(auditId, objectName, auditDate, storageMode, startedAt) {
  const auditFolder = getAuditFolderV3_(auditId, objectName, auditDate, storageMode, startedAt);
  // Фотографии храним в отдельной подпапке внутри папки конкретного аудита,
  // чтобы Excel, Приложение PDF и черновик всегда оставались сверху.
  return getOrCreateSubfolder_(auditFolder, 'Фото');
}


function uploadAuditPhotoBatchV2_(payload) {
  const auditId=String(payload.auditId||'').trim();
  if(!auditId) throw new Error('Не указан auditId.');
  const photos=Array.isArray(payload.photos)?payload.photos:[];
  if(!photos.length) throw new Error('Не передан пакет фотографий.');
  if(photos.length>3) throw new Error('Размер пакета не должен превышать 3 фотографии.');
  const folder=getUploadSessionFolderV2_(auditId,payload.objectName,payload.auditDate,payload.storageMode,payload.startedAt);
  const existing={};
  const files=folder.getFiles();
  while(files.hasNext()){
    const f=files.next();
    const m=f.getName().match(/^photo_(P_[^_]+)_/);
    if(m) existing[m[1]]=f;
  }
  const uploaded=[];
  photos.forEach(function(p){
    const photoId=String(p.photoId||'').trim();
    if(!photoId) throw new Error('В пакете есть фотография без photoId.');
    if(!p.data) throw new Error('Не передано содержимое фотографии '+photoId+'.');
    if(existing[photoId]){uploaded.push({photoId:photoId,alreadyExists:true,fileId:existing[photoId].getId(),url:existing[photoId].getUrl()});return;}
    const prefix='photo_'+sanitizeFileName_(photoId)+'_';
    const fileName=prefix+'раздел_'+sanitizeFileName_(p.sectionIndex||'0')+'_пункт_'+sanitizeFileName_(p.questionId||'0')+'_фото_'+sanitizeFileName_(p.order||'1')+'_'+sanitizeFileName_(p.name||'photo.jpg');
    const bytes=Utilities.base64Decode(String(p.data));
    const file=folder.createFile(Utilities.newBlob(bytes,p.mimeType||'image/jpeg',fileName));
    existing[photoId]=file;
    uploaded.push({photoId:photoId,alreadyExists:false,fileId:file.getId(),url:file.getUrl()});
  });
  return {ok:true,auditId:auditId,count:uploaded.length,uploaded:uploaded};
}

function uploadAuditPhotoV2_(payload) {
  const auditId=String(payload.auditId||'').trim();
  const photoId=String(payload.photoId||'').trim();
  if(!auditId) throw new Error('Не указан auditId.');
  if(!photoId) throw new Error('Не указан photoId.');
  if(!payload.data) throw new Error('Не передано содержимое фотографии.');

  const folder=getUploadSessionFolderV2_(auditId,payload.objectName,payload.auditDate,payload.storageMode,payload.startedAt);
  const prefix='photo_'+sanitizeFileName_(photoId)+'_';
  const files=folder.getFiles();
  while(files.hasNext()){
    const f=files.next();
    if(f.getName().indexOf(prefix)===0) return {ok:true,photoId:photoId,alreadyExists:true,fileId:f.getId(),url:f.getUrl()};
  }

  const fileName=prefix +
    'раздел_' + sanitizeFileName_(payload.sectionIndex || '0') +
    '_пункт_' + sanitizeFileName_(payload.questionId || '0') +
    '_фото_' + sanitizeFileName_(payload.order || '1') +
    '_' + sanitizeFileName_(payload.name || 'photo.jpg');
  const bytes=Utilities.base64Decode(String(payload.data));
  const blob=Utilities.newBlob(bytes,payload.mimeType || 'image/jpeg',fileName);
  const file=folder.createFile(blob);
  return {ok:true,photoId:photoId,alreadyExists:false,fileId:file.getId(),url:file.getUrl()};
}

function indexUploadFilesV2_(folder) {
  const map={};
  // Начиная с v4 фотографии находятся в подпапке "Фото".
  // Оставляем fallback на саму папку для совместимости со старыми загрузками.
  let sourceFolder = folder;
  const photoFolders = folder.getFoldersByName('Фото');
  if (photoFolders.hasNext()) sourceFolder = photoFolders.next();
  const files=sourceFolder.getFiles();
  while(files.hasNext()){
    const file=files.next();
    const m=file.getName().match(/^photo_(P_[^_]+)_/);
    if(m) map[m[1]]={id:file.getId(),url:file.getUrl()};
  }
  return map;
}

function finalizeAuditUploadV2_(payload) {
  const objectName=String(payload.objectName||'').trim();
  const auditDate=String(payload.auditDate||'').trim();
  const auditId=String(payload.auditId||'').trim();
  if(!objectName) throw new Error('Укажите наименование объекта.');
  if(!auditDate) throw new Error('Укажите дату проведения.');
  if(!auditId) throw new Error('Не указан auditId.');

  const folder=getAuditFolderV3_(auditId,objectName,auditDate,payload.storageMode,payload.startedAt);
  const fileMap=indexUploadFilesV2_(folder);
  const expected=Array.isArray(payload.expectedPhotos)?payload.expectedPhotos:[];
  const missing=expected.filter(function(p){return !p || !p.photoId || !fileMap[String(p.photoId)];});
  if(missing.length){
    throw new Error('Не все фотографии загружены. Отсутствует: '+missing.slice(0,12).map(function(p){return p.photoId;}).join(', ')+(missing.length>12?' и ещё '+(missing.length-12):''));
  }

  const uploadAt=new Date();
  const uploadStamp=Utilities.formatDate(uploadAt,Session.getScriptTimeZone(),'yyyy-MM-dd_HH-mm-ss-SSS');
  const report=getOrCreateReport_(objectName,auditDate,uploadStamp,payload.storageMode);
  const sourceQuestions=parseSourceSheet_(getSourceSheet_());
  const questionMap={}; sourceQuestions.forEach(function(q){questionMap[q.id]=q;});

  // Аварийная/локальная версия может прислать снимок чек-листа,
  // прочитанный из Excel. Если он валиден, используем его веса
  // для расчёта результата. Внешний вид и структура отчёта не меняются.
  const clientSections = payload && payload.appData && Array.isArray(payload.appData.sections)
    ? payload.appData.sections : [];
  clientSections.forEach(function(section){
    (section.items || []).forEach(function(item){
      const id=String(item.id||'').trim();
      if(!id) return;
      const weight=String(item.weight==null?'':item.weight).trim();
      if(!questionMap[id]) questionMap[id]={id:id,point:String(item.point||''),weight:weight};
      else if(weight !== '') questionMap[id].weight=weight;
    });
  });

  const rows=[];

  (payload.sections||[]).forEach(function(section){
    if(section.skipped) return;
    const responsible=String(section.responsible||'').trim();
    (section.questions||[]).forEach(function(item){
      const answer=String(item.answer||'').trim();
      const comment=String(item.comment||'').trim();
      if(!answer) return;
      if(CFG.NEGATIVE_ANSWERS.indexOf(answer)>=0 && !comment) throw new Error('Для пункта "'+String(item.point||'')+'" необходимо описание нарушения.');
      const photoLinks=(item.photos||[]).filter(function(p){return p && p.photoId && fileMap[String(p.photoId)];}).sort(function(a,b){return (a.order||0)-(b.order||0);}).map(function(p){return fileMap[String(p.photoId)].url;});
      rows.push([uploadAt,auditDate,objectName,String(section.name||''),responsible,String(item.point||''),answer,comment,photoLinks.join('\n')]);
    });
  });
  if(!rows.length) throw new Error('Нет заполненных пунктов.');

  const sheet=report.sheet; const startRow=Math.max(sheet.getLastRow()+1,2);
  sheet.getRange(startRow,1,rows.length,9).setValues(rows);
  sheet.setFrozenRows(1); sheet.getRange(1,1,1,9).setFontWeight('bold').setWrap(true); sheet.setColumnWidth(9,320);

  // Создаём только Excel-отчёт. Отдельный PDF основного отчёта больше не нужен.
  createAuditResultSheet_(report.file.getId(),payload,questionMap);
  SpreadsheetApp.flush();
  const auditFolder=getAuditFolderV3_(auditId,objectName,auditDate,payload.storageMode,payload.startedAt);
  const reportXlsx=createAuditReportXlsxFileV3_(report.file,report.name,auditFolder);
  const draftFile=createAuditDraftFileV3_(payload,fileMap,auditFolder,uploadStamp);
  try { report.file.setTrashed(true); } catch (e) {}

  const statusKey='AUDIT_APPENDIX_V2_'+auditId;
  PropertiesService.getScriptProperties().setProperty(statusKey,JSON.stringify({
    status:'queued',auditId:auditId,reportName:reportXlsx.name,reportUrl:reportXlsx.downloadUrl,
    reportXlsxName:reportXlsx.name,reportXlsxUrl:reportXlsx.downloadUrl,draftName:draftFile.name,draftUrl:draftFile.downloadUrl,auditFolderUrl:auditFolder.getUrl(),
    uploadedPhotos:expected.length,updatedAt:new Date().toISOString()
  }));

  // Сохраняем задание в staging-папке, чтобы фоновый триггер не зависел от браузера.
  const jobName='appendix_job_'+sanitizeFileName_(auditId)+'.json';
  const oldJobs=folder.getFilesByName(jobName); while(oldJobs.hasNext()) oldJobs.next().setTrashed(true);
  folder.createFile(jobName,JSON.stringify({payload:payload,fileMap:fileMap,uploadStamp:uploadStamp,reportName:reportXlsx.name,reportUrl:reportXlsx.downloadUrl,reportXlsxName:reportXlsx.name,reportXlsxUrl:reportXlsx.downloadUrl,draftName:draftFile.name,draftUrl:draftFile.downloadUrl,auditFolderUrl:auditFolder.getUrl()}),MimeType.PLAIN_TEXT);
  // PDF-приложение формируем сразу в рамках этой же выгрузки.
  // Это убирает зависимость от time-based триггеров и лимита триггеров.
  processAuditAppendixV2_();

  const finalStatus=getAuditAppendixStatusV2(auditId);
  return {ok:true,savedRows:rows.length,reportName:reportXlsx.name,reportUrl:reportXlsx.downloadUrl,reportXlsxName:reportXlsx.name,reportXlsxUrl:reportXlsx.downloadUrl,draftName:draftFile.name,draftUrl:draftFile.downloadUrl,auditFolderUrl:auditFolder.getUrl(),appendixPdfName:finalStatus.appendixPdfName||'',appendixPdfUrl:finalStatus.appendixPdfUrl||'',uploadedPhotos:expected.length,appendixStatus:finalStatus.status||'done',auditId:auditId};
}


function scheduleAuditAppendixTriggerV2_(){
  // Для всех очередей достаточно одного одноразового триггера:
  // обработчик сам просматривает все задания appendix_job_*.json.
  // Перед созданием удаляем дубликаты этого же обработчика, чтобы
  // повторные выгрузки не приводили к превышению лимита триггеров.
  const triggers=ScriptApp.getProjectTriggers();
  let kept=false;
  triggers.forEach(function(trigger){
    if(trigger.getHandlerFunction()!=='processAuditAppendixV2_') return;
    if(!kept){
      kept=true;
    }else{
      try{ScriptApp.deleteTrigger(trigger);}catch(e){}
    }
  });
  if(!kept){
    ScriptApp.newTrigger('processAuditAppendixV2_').timeBased().after(1000).create();
  }
}

function processAuditAppendixV2_(){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1000)) return;
  try{
    const roots=[];
    try { roots.push(getResultsReportFolder_('SCRIPT_ACCOUNT')); } catch(e) {}
    try { roots.push(getResultsReportFolder_('SECOND_ACCOUNT')); } catch(e) {}
    roots.forEach(function(root){
      const folders=root.getFolders();
      while(folders.hasNext()){
        const folder=folders.next();
      const jobs=folder.getFiles();
      while(jobs.hasNext()){
        const job=jobs.next();
        if(job.getName().indexOf('appendix_job_')!==0) continue;
        let data;
        try{data=JSON.parse(job.getBlob().getDataAsString());}catch(e){continue;}
        const payload=data.payload||{}; const auditId=String(payload.auditId||'').trim();
        if(!auditId) continue;
        const key='AUDIT_APPENDIX_V2_'+auditId;
        try{
          PropertiesService.getScriptProperties().setProperty(key,JSON.stringify({status:'processing',stage:'pdf',current:0,total:0,auditId:auditId,reportName:data.reportName,reportUrl:data.reportUrl,reportXlsxName:data.reportXlsxName||'',reportXlsxUrl:data.reportXlsxUrl||'',draftName:data.draftName||'',draftUrl:data.draftUrl||'',auditFolderUrl:data.auditFolderUrl||'',updatedAt:new Date().toISOString()}));
          const result=createAuditApplicationPdfFromDriveV2_(payload,data.fileMap||{},data.uploadStamp||'',auditId);
          PropertiesService.getScriptProperties().setProperty(key,JSON.stringify({status:'done',stage:'done',current:result.items,total:result.items,auditId:auditId,reportName:data.reportName,reportUrl:data.reportUrl,reportXlsxName:data.reportXlsxName||'',reportXlsxUrl:data.reportXlsxUrl||'',appendixPdfName:result.name,appendixPdfUrl:result.downloadUrl,auditFolderUrl:result.folderUrl||data.auditFolderUrl||'',packageName:result.packageName||'',packageUrl:result.packageUrl||'',uploadedPhotos:(Array.isArray(payload.expectedPhotos)?payload.expectedPhotos.length:0),updatedAt:new Date().toISOString()}));
          job.setTrashed(true);
        }catch(err){
          PropertiesService.getScriptProperties().setProperty(key,JSON.stringify({status:'error',stage:'pdf',auditId:auditId,reportName:data.reportName,reportUrl:data.reportUrl,error:String(err&&err.message?err.message:err),updatedAt:new Date().toISOString()}));
          job.setTrashed(true);
        }
        }
      }
    });
  }finally{lock.releaseLock();}
}

function getAuditAppendixStatusV2(auditId){
  const id=String(auditId||'').trim();
  if(!id) return {status:'error',error:'Не указан auditId.'};
  const raw=PropertiesService.getScriptProperties().getProperty('AUDIT_APPENDIX_V2_'+id);
  if(!raw) return {status:'unknown',auditId:id};
  try{return JSON.parse(raw);}catch(e){return {status:'error',auditId:id,error:String(e)};}
}

function createAuditApplicationPdfFromDriveV2_(payload,fileMap,uploadStamp,auditId){
  /*
   * Формат приложения приведён к образцу пользователя:
   * A4 portrait, титульный блок с реквизитами, затем разделы.
   * В каждом разделе: название + процент, ответственный,
   * таблица "Пункт / Заметка / Результат".
   * Фотографии находятся в колонке "Заметка".
   */
  const objectName=String(payload.objectName||'объект').trim()||'объект';
  const safeObjectName=sanitizeFileName_(objectName);
  const auditDate=sanitizeFileName_(payload.auditDate||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'));
  const stamp=sanitizeFileName_(uploadStamp||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd_HH-mm-ss-SSS'));
  const docName='Приложение_'+safeObjectName+'_'+auditDate+'_'+stamp;

  const checklistName = String(
    (payload.appData && (payload.appData.title || payload.appData.checklistTitle)) ||
    'Аудит'
  ).trim() || 'Аудит';
  const checkerName = String(
    payload.checkerName || payload.auditorName || payload.inspectorName || ''
  ).trim();
  const startedAt = String(
    payload.startedAt || payload.auditStartedAt || ''
  ).trim();
  const finishedAt = String(
    payload.finishedAt || payload.auditFinishedAt || ''
  ).trim();

  const sections=[];
  let totalAnswered=0;
  let totalPhotos=0;
  let totalScore=0;
  let totalPossible=0;

  (payload.sections||[]).forEach(function(section){
    if(section && section.skipped) return;
    const rows=[];
    let sectionScore=0;
    let sectionPossible=0;

    (section.items || section.questions || []).forEach(function(item){
      const answer=String(item.answer||'').trim();
      if(!answer) return;

      const qId=String(item.id||'').trim();
      const appItems=(payload.appData && Array.isArray(payload.appData.sections)) ? payload.appData.sections : [];
      let weight=0;
      for(let si=0;si<appItems.length;si++){
        const its=appItems[si].items||[];
        for(let qi=0;qi<its.length;qi++){
          if(String(its[qi].id||'').trim()===qId){
            weight=Number(String(its[qi].weight==null?'':its[qi].weight).replace(',','.'))||0;
            si=appItems.length; break;
          }
        }
      }

      if(answer!=='Пропущено'){
        sectionPossible += weight;
        if(answer==='Да') sectionScore += weight;
        else if(answer==='Незначительные недочёты') sectionScore += 60;
        else if(answer==='Значительные нарушения') sectionScore += 30;
      }

      const photos=(item.photos||[]).filter(function(photo){
        return photo && photo.photoId && fileMap[String(photo.photoId)];
      }).slice(0,CFG.MAX_PHOTOS);

      totalAnswered++;
      totalPhotos += photos.length;
      rows.push({
        point:String(item.point||''),
        comment:String(item.comment||''),
        answer:answer,
        photos:photos
      });
    });

    const percent=sectionPossible ? (sectionScore/sectionPossible*100) : 0;
    totalScore += sectionScore;
    totalPossible += sectionPossible;
    sections.push({
      name:String(section.name||''),
      responsible:String(section.responsible||'').trim(),
      rows:rows,
      percent:percent
    });
  });

  const overallPercent=totalPossible ? (totalScore/totalPossible*100) : 0;

  const doc=DocumentApp.create(docName);
  const body=doc.getBody();
  body.clear();
  body.setPageWidth(595.28);  // A4 portrait
  body.setPageHeight(841.89);
  body.setMarginTop(28);
  body.setMarginBottom(24);
  body.setMarginLeft(28);
  body.setMarginRight(28);

  // ===== Верхняя часть как в образце =====
  const title=body.appendParagraph('Результаты чек-листа');
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.setBold(true).setFontSize(14).setSpacingBefore(0).setSpacingAfter(7);

  const meta=body.appendTable([
    ['Название чек-листа:', checklistName],
    ['Имя проверяющего:', checkerName],
    ['Проверяемый объект:', objectName],
    ['Время начала чек-листа:', formatReportDateTime_(startedAt)],
    ['Время завершения чек-листа:', formatReportDateTime_(finishedAt || uploadStamp)]
  ]);
  meta.setBorderWidth(0);
  meta.setColumnWidth(0,175);
  meta.setColumnWidth(1,335);
  for(let r=0;r<5;r++){
    const a=meta.getCell(r,0).getChild(0).asParagraph();
    const b=meta.getCell(r,1).getChild(0).asParagraph();
    a.setFontSize(8).setBold(true).setSpacingBefore(0).setSpacingAfter(1);
    b.setFontSize(8).setSpacingBefore(0).setSpacingAfter(1);
  }

  body.appendParagraph('').setSpacingBefore(0).setSpacingAfter(1);
  const resultTitle=body.appendParagraph('Результаты чек-листа');
  resultTitle.setBold(true).setFontSize(10).setSpacingBefore(0).setSpacingAfter(3);

  sections.forEach(function(section){
    appendAuditReportSectionV3_(body,section,fileMap);
  });

  const final=body.appendParagraph('Чек-лист пройден на '+formatPercentRu_(overallPercent));
  final.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  final.setBold(true).setFontSize(9).setSpacingBefore(5).setSpacingAfter(0);

  doc.saveAndClose();

  const source=DriveApp.getFileById(doc.getId());
  const folder=getAuditFolderV3_(auditId,objectName,auditDate,payload.storageMode,payload.startedAt);
  const pdfName=docName+'.pdf';
  const pdfBlob=source.getAs(MimeType.PDF).setName(pdfName);
  const file=folder.createFile(pdfBlob);

  const packageFile=null;

  try{source.setTrashed(true);}catch(e){}

  return {
    name:file.getName(),
    url:file.getUrl(),
    downloadUrl:makeDriveDownloadUrlV2_(file.getId()),
    id:file.getId(),
    items:totalAnswered,
    negativeItems:0,
    photos:totalPhotos,
    folderUrl:folder.getUrl(),
    packageName:'',
    packageUrl:''
  };
}

function formatReportDateTime_(value){
  const s=String(value||'').trim();
  if(!s) return '';
  const d=new Date(s);
  if(!isNaN(d.getTime())){
    return Utilities.formatDate(d,Session.getScriptTimeZone(),'d MMMM yyyy г. HH:mm');
  }
  return s;
}

function formatPercentRu_(value){
  const n=Number(value)||0;
  let text=n.toFixed(2).replace('.',',');
  text=text.replace(/,00$/,'').replace(/(,\d)0$/,'$1');
  return text+'%';
}

function appendAuditReportSectionV3_(body,section,fileMap){
  const header=body.appendParagraph(
    String(section.name||'Раздел')+' ( '+formatPercentRu_(section.percent)+' )'
  );
  header.setBold(true).setFontSize(10).setSpacingBefore(4).setSpacingAfter(1);

  const resp=body.appendParagraph('Проверяемый сотрудник: '+(section.responsible||'Не указан'));
  resp.setFontSize(7).setSpacingBefore(0).setSpacingAfter(2);

  const table=body.appendTable();
  table.setBorderWidth(0.5);
  table.setColumnWidth(0,290);
  table.setColumnWidth(1,150);
  table.setColumnWidth(2,70);

  const head=table.appendTableRow();
  ['Пункт','Заметка','Результат'].forEach(function(text,i){
    const c=head.appendTableCell(text);
    c.setBackgroundColor('#EEEEEE');
    c.setPaddingTop(2); c.setPaddingBottom(2); c.setPaddingLeft(3); c.setPaddingRight(3);
    const p=c.getChild(0).asParagraph();
    p.setFontSize(7).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  });

  section.rows.forEach(function(item,rowIndex){
    const row=table.appendTableRow();
    const pointCell=row.appendTableCell();
    const noteCell=row.appendTableCell();
    const resultCell=row.appendTableCell();

    [pointCell,noteCell,resultCell].forEach(function(c){
      c.setPaddingTop(2); c.setPaddingBottom(2); c.setPaddingLeft(3); c.setPaddingRight(3);
      if(rowIndex % 2 === 0) c.setBackgroundColor('#F3F3F3');
    });

    const pp=pointCell.getChild(0).asParagraph();
    pp.appendText(item.point||'').setFontSize(7);
    pp.setSpacingBefore(0).setSpacingAfter(0);

    const cp=noteCell.getChild(0).asParagraph();
    cp.setSpacingBefore(0).setSpacingAfter(0);
    if(item.comment){
      cp.appendText(item.comment).setFontSize(7);
    }

    // Фотографии находятся именно в колонке "Заметка", как в образце.
    if(item.photos.length){
      let grid=null;
      for(let i=0;i<item.photos.length;i+=2){
        grid=noteCell.appendTable();
        grid.setBorderWidth(0);
        grid.setColumnWidth(0,72);
        grid.setColumnWidth(1,72);
        const photoRow=grid.appendTableRow();
        for(let j=0;j<2;j++){
          const idx=i+j;
          const cell=photoRow.appendTableCell();
          cell.setPaddingTop(1); cell.setPaddingBottom(1); cell.setPaddingLeft(1); cell.setPaddingRight(1);
          if(idx>=item.photos.length) continue;
          try{
            const ref=fileMap[String(item.photos[idx].photoId)];
            const f=DriveApp.getFileById(ref.id);
            const img=cell.appendImage(f.getBlob());
            let w=img.getWidth(), h=img.getHeight();
            const maxW=70, maxH=70;
            if(w>maxW){h=Math.round(h*maxW/w);w=maxW;}
            if(h>maxH){w=Math.round(w*maxH/h);h=maxH;}
            img.setWidth(w); img.setHeight(h);
          }catch(e){
            cell.appendParagraph('Фото: ошибка').setFontSize(5).setForegroundColor('#C5221F');
          }
        }
      }
    }

    const rp=resultCell.getChild(0).asParagraph();
    rp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    rp.setSpacingBefore(0).setSpacingAfter(0);
    rp.appendText(item.answer||'').setFontSize(7).setBold(false);
    if(item.answer==='Да') rp.setForegroundColor('#34A853');
    else if(item.answer==='Нет') rp.setForegroundColor('#EA4335');
    else if(item.answer==='Пропущено') rp.setForegroundColor('#4EA3D8');
    else if(item.answer==='Значительные нарушения') rp.setForegroundColor('#EA4335');
    else if(item.answer==='Незначительные недочёты') rp.setForegroundColor('#F29900');
  });

  body.appendParagraph('').setSpacingBefore(0).setSpacingAfter(1);
}

function makeSpreadsheetXlsxDownloadUrlV2_(spreadsheetId){
  return 'https://docs.google.com/spreadsheets/d/' +
    encodeURIComponent(String(spreadsheetId)) +
    '/export?format=xlsx';
}

// Оставлено для совместимости со старыми вызовами. Файл на сервере не создаётся.
function createAuditReportXlsxV2_(reportFile,objectName,auditDate,uploadStamp,auditFolder){
  return {
    name: reportFile.getName() + '.xlsx',
    id: reportFile.getId(),
    url: makeSpreadsheetXlsxDownloadUrlV2_(reportFile.getId()),
    downloadUrl: makeSpreadsheetXlsxDownloadUrlV2_(reportFile.getId())
  };
}

function createAuditReportXlsxFileV3_(reportFile,reportName,auditFolder){
  const url=makeSpreadsheetXlsxDownloadUrlV2_(reportFile.getId());
  const response=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});
  const code=response.getResponseCode();
  if(code<200 || code>=300) throw new Error('Не удалось сформировать Excel-файл. Код: '+code+' '+response.getContentText().slice(0,300));
  const file=auditFolder.createFile(response.getBlob().setName(String(reportName||'Отчёт_по_аудиту')+'.xlsx'));
  return {name:file.getName(),id:file.getId(),url:file.getUrl(),downloadUrl:makeDriveDownloadUrlV2_(file.getId())};
}

function createAuditDraftFileV3_(payload,fileMap,auditFolder,uploadStamp){
  const safeObject=sanitizeFileName_(payload.objectName||'объект');
  const safeDate=sanitizeFileName_(payload.auditDate||'без_даты');
  const stamp=sanitizeFileName_(uploadStamp||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd_HH-mm-ss-SSS'));
  const audit=JSON.parse(JSON.stringify(payload));
  audit.savedAt=new Date().toISOString();
  const photos=[];
  (payload.sections||[]).forEach(function(section){
    (section.questions||section.items||[]).forEach(function(item){
      (item.photos||[]).forEach(function(photo){
        const ref=fileMap[String(photo.photoId||'')];
        photos.push({photoId:photo.photoId||'',sectionIndex:String(section.index||photo.sectionIndex||''),questionId:String(item.id||photo.questionId||''),name:photo.name||'photo.jpg',mimeType:photo.mimeType||'image/jpeg',order:photo.order||1,driveFileId:ref?ref.id:'',driveUrl:ref?ref.url:''});
      });
    });
  });
  const pack={format:'audit-pwa-draft-server',version:1,exportedAt:new Date().toISOString(),audit:audit,photos:photos};
  const name='Черновик_формы_'+safeObject+'_'+safeDate+'_'+stamp+'.audit';
  const file=auditFolder.createFile(name,JSON.stringify(pack,null,2),MimeType.PLAIN_TEXT);
  return {name:file.getName(),id:file.getId(),url:file.getUrl(),downloadUrl:makeDriveDownloadUrlV2_(file.getId())};
}

function createAuditReportPdfV2_(reportFile,objectName,auditDate,uploadStamp,auditFolder){
  const safeObjectName=sanitizeFileName_(objectName||'объект');
  const safeDate=sanitizeFileName_(auditDate||'без_даты');
  const stamp=sanitizeFileName_(uploadStamp||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd_HH-mm-ss-SSS'));
  const pdfName='Отчёт_по_аудиту_'+safeObjectName+'_'+safeDate+'_'+stamp+'.pdf';
  const blob=reportFile.getAs(MimeType.PDF).setName(pdfName);
  const file=auditFolder.createFile(blob);
  return {name:file.getName(),id:file.getId(),url:file.getUrl(),downloadUrl:makeDriveDownloadUrlV2_(file.getId())};
}

function makeDriveDownloadUrlV2_(fileId){
  return 'https://drive.google.com/uc?export=download&id='+encodeURIComponent(String(fileId));
}

function appendAuditResultBlockFromDriveV2_(body,item,index,fileMap){
  const hasComment=!!String(item.comment||'').trim();
  const hasPhotos=Array.isArray(item.photos)&&item.photos.length>0;

  // Компактный верхний блок: граница ответа максимально вправо.
  const head=body.appendTable([['','']]);
  head.setBorderWidth(1);
  const left=head.getCell(0,0), right=head.getCell(0,1);
  // Узкая колонка ответа, чтобы "Ответ: Да" помещалось в одну строку.
  head.setColumnWidth(0, 445);
  head.setColumnWidth(1, 98);

  const sec=left.getChild(0).asParagraph();
  sec.setSpacingBefore(0).setSpacingAfter(0);
  sec.appendText('РАЗДЕЛ: '+(item.section||'')).setFontSize(11).setBold(false);

  const point=left.appendParagraph('№ '+index+' — '+(item.point||''));
  point.setFontSize(11).setSpacingBefore(0).setSpacingAfter(0).setBold(false);

  if(item.responsible){
    const resp=left.appendParagraph('Ответственный: '+item.responsible);
    resp.setFontSize(8).setSpacingBefore(0).setSpacingAfter(0).setBold(false);
  }

  const ans=right.getChild(0).asParagraph();
  ans.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  ans.setSpacingBefore(0).setSpacingAfter(0);
  ans.appendText('Ответ: '+(item.answer||''));
  ans.setBold(false).setFontSize(12);
  // Цвета Да/Нет сохраняем.
  if(item.answer==='Да') ans.setForegroundColor('#137333');
  else if(item.answer==='Нет') ans.setForegroundColor('#C5221F');
  else if(item.answer==='Пропущено') ans.setForegroundColor('#B06000');

  // Комментарий — только при наличии текста, без зазора после блока пункта.
  if(hasComment){
    const commentTable=body.appendTable([['']]);
    commentTable.setBorderWidth(1);
    const cc=commentTable.getCell(0,0);
    const cl=cc.getChild(0).asParagraph();
    cl.setSpacingBefore(0).setSpacingAfter(0);
    cl.appendText('Комментарий:').setFontSize(10).setBold(false);
    const cp=cc.appendParagraph(item.comment);
    cp.setFontSize(9).setSpacingBefore(0).setSpacingAfter(0).setBold(false);
    if(item.negative) cp.setForegroundColor('#8B0000');
  }

  // Фотографии — только если они есть. Без заголовка, подписей и границ.
  // Три фотографии в один ряд.
  if(hasPhotos){
    for(let i=0;i<item.photos.length;i+=3){
      const photoTable=body.appendTable([['','','']]);
      photoTable.setBorderWidth(0);
      photoTable.setColumnWidth(0, 181);
      photoTable.setColumnWidth(1, 181);
      photoTable.setColumnWidth(2, 181);

      for(let j=0;j<3;j++){
        const idx=i+j;
        const cell=photoTable.getCell(0,j);
        if(idx>=item.photos.length) continue;
        const photo=item.photos[idx];
        try{
          const ref=fileMap[String(photo.photoId)];
          const f=DriveApp.getFileById(ref.id);
          const img=cell.appendImage(f.getBlob());
          let w=img.getWidth(), h=img.getHeight();
          const maxW=175, maxH=195;
          if(w>maxW){h=Math.round(h*maxW/w);w=maxW;}
          if(h>maxH){w=Math.round(w*maxH/h);h=maxH;}
          img.setWidth(w).setHeight(h);
        }catch(e){
          const err=cell.getChild(0).asParagraph();
          err.setSpacingBefore(0).setSpacingAfter(0);
          err.appendText('Ошибка фото').setFontSize(7).setForegroundColor('#C5221F').setBold(false);
        }
      }
    }
  }

  // Минимальный интервал перед следующим пунктом.
  body.appendParagraph('').setSpacingBefore(0).setSpacingAfter(0);
}
