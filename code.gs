// Глобальные переменные для кеширования
let spreadsheet;
let sheet;
let referenceSheet;
let balancesSheet; // Новая переменная

const SPREADSHEET_ID = '1DxcZcfdNqLC5kVSZutkcT1mWBYKpQBtvMkSHwJMfD1w';
const DATA_SHEET_NAME = 'Выдачи';
const REF_SHEET_NAME = 'Справочник';
const BALANCES_SHEET_NAME = 'Остатки'; // Новая константа

/**
 * Инициализация доступа к таблице
 */
function initializeSheets() {
  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  if (!sheet) {
    sheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(DATA_SHEET_NAME);
      sheet.appendRow(['Время', 'Старший смены', 'ID сотрудника', 'Наименование учёта', 'Количество']);
    }
  }
  if (!referenceSheet) {
    referenceSheet = spreadsheet.getSheetByName(REF_SHEET_NAME);
    if (!referenceSheet) {
      throw new Error(`Лист "${REF_SHEET_NAME}" не найден.`);
    }
  }
  if (!balancesSheet) {
    balancesSheet = spreadsheet.getSheetByName(BALANCES_SHEET_NAME);
    if (!balancesSheet) {
      throw new Error(`Лист "${BALANCES_SHEET_NAME}" не найден.`);
    }
  }
}

/**
 * Проверка ID старшего смены
 */
function isShiftManager(id) {
  const validIds = {
    "634137": "Кадеев Денис Юрьевич",
    "998356": "Петрашкевич Виталий Валерьевич",
    "1228349": "Маховиченко Владислав Валерьевич",
    "936430": "Семеняко Иоанн Андреевич",
    "819898": "Тамкович Андрей Юрьевич",
    "958269": "Шлетгауэр Вадим Юрьевич",
    "634151": "Ерофеев Максим Александрович"
  };
  const name = validIds[id];
  return name ? { valid: true, name: name } : { valid: false };
}

/**
 * Загрузка справочника материалов и данных об остатках
 */
function getMaterialsWithBalances() {
  try {
    initializeSheets();
    
    // Загрузка данных из Справочника
    const materialsData = referenceSheet.getDataRange().getValues();
    const headers = materialsData[0];

    const catCol = headers.indexOf('Категория (отображение)');
    const subCatCol = headers.indexOf('Подкатегория (отображение)');
    const unitCol = headers.indexOf('Ед. изм.');
    const typeCol = headers.indexOf('Тип');
    const nameCol = headers.indexOf('Наименование учёта');

    if ([catCol, subCatCol, unitCol, typeCol, nameCol].some(idx => idx === -1)) {
      throw new Error('Некорректные заголовки в листе "Справочник"');
    }

    const materials = {};
    const units = {};

    for (let i = 1; i < materialsData.length; i++) {
      const row = materialsData[i];
      const category = row[catCol];
      const subcategory = row[subCatCol];
      const unit = row[unitCol];
      const type = row[typeCol];
      const accountName = row[nameCol];

      if (!category || !type || !accountName) continue;

      if (type === 'category') {
        materials[category] = {
          value: accountName,
          display: category,
          isMaterial: false,
          hasBalance: false, 
          subcategories: []
        };
        units[accountName] = unit;
      } else if (type === 'subcategory') {
        if (!materials[category]) {
          materials[category] = {
            value: category,
            display: category,
            subcategories: [],
            isMaterial: false,
            hasBalance: false
          };
        }
        materials[category].subcategories.push({
          value: accountName,
          display: subcategory,
          hasBalance: false
        });
        units[accountName] = unit;
      } else if (type === 'material') {
        materials[category] = {
          value: accountName,
          display: category,
          isMaterial: true,
          hasBalance: false
        };
        units[accountName] = unit;
      }
    }
    
    // Загрузка данных из Остатков
    const balancesRange = balancesSheet.getRange('A2:B' + balancesSheet.getLastRow()).getValues();
    const balances = {};
    for (const [name, balance] of balancesRange) {
      if (name) {
        balances[name.trim()] = balance;
      }
    }

    // Обновление материалов данными об остатках
    for (const categoryKey in materials) {
      const category = materials[categoryKey];
      if (category.isMaterial) {
        const balance = balances[category.value] || 0;
        category.hasBalance = balance > 0;
      } else {
        let hasBalanceInSubcategories = false;
        category.subcategories.forEach(sub => {
          const balance = balances[sub.value] || 0;
          sub.hasBalance = balance > 0;
          if (sub.hasBalance) {
            hasBalanceInSubcategories = true;
          }
        });
        category.hasBalance = hasBalanceInSubcategories;
      }
    }

    return { materials, units };
  } catch (e) {
    console.error('Ошибка при загрузке материалов и остатков:', e);
    throw new Error('Не удалось загрузить справочник и остатки.');
  }
}

/**
 * Запись данных в таблицу
 */
function appendDataToSheet(data) {
  const errorId = 'ERR_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
  try {
    initializeSheets();
    sheet.appendRow([
      data.timestamp,
      data.shiftManager,
      data.employeeId,
      data.accountName,
      data.quantity
    ]);
    return { success: true };
  } catch (error) {
    console.error(`Ошибка [${errorId}]:`, error);
    return { success: false, errorId, message: `Ошибка сервера [${errorId}]` };
  }
}

/**
 * Обработчик GET-запроса
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Учет выдачи материалов')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}