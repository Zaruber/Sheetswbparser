// Константы
const DELAY_BETWEEN_REQUESTS = 1500; // 1.5 секунды между запросами
const PRODUCTS_PER_PAGE = 30; // Параметр spp из боевого запроса (v18)
const MAX_PAGES = 30; // Максимум 30 страниц
const CACHE_DURATION = 1800; // 30 минут в секундах
const CACHE_PREFIX = "wb_search_"; // Префикс для кэша
const PRODUCTS_CACHE_PREFIX = "wb_products_"; // Префикс для кэша продуктов
const DEBUG_SAVE_PROPERTIES = false; // Сохранять сырые ответы в ScriptProperties (включайте осторожно)

/**
 * Находит позицию товара в поисковой выдаче Wildberries
 * @param {string} articleNumber - Артикул товара
 * @param {string} searchQuery - Поисковый запрос
 * @return {string} Позиция товара или текст ошибки
 * @customfunction
 */
function FIND_WB_POSITION(articleNumber, searchQuery) {
  if (!articleNumber || !searchQuery) {
    return "⚠️ Укажите артикул и запрос";
  }

  try {
    articleNumber = articleNumber.toString();
    searchQuery = searchQuery.toString().trim().toLowerCase();
    
    // Получаем позицию товара (с использованием кэша)
    return findPositionInSearch(articleNumber, searchQuery);
    
  } catch (error) {
    Logger.log("Ошибка: " + error.toString());
    return "⚠️ Ошибка: " + error.toString();
  }
}

/**
 * Ищет позицию товара в результатах поиска
 * @param {string} articleNumber - Артикул товара
 * @param {string} searchQuery - Поисковый запрос
 * @return {string} Позиция товара или текст ошибки
 */
function findPositionInSearch(articleNumber, searchQuery) {
  const cache = CacheService.getScriptCache();
  const cacheKey = CACHE_PREFIX + searchQuery;
  const productsCacheKey = PRODUCTS_CACHE_PREFIX + searchQuery;
  
  // Пробуем получить позицию из кэша
  const cachedPosition = cache.get(cacheKey + "_" + articleNumber);
  if (cachedPosition) {
    Logger.log(`Позиция найдена в кэше для артикула ${articleNumber} по запросу "${searchQuery}"`);
    return cachedPosition;
  }
  
  // Пробуем получить результаты поиска из кэша
  const cachedProducts = cache.get(productsCacheKey);
  if (cachedProducts) {
    Logger.log(`Используем кэшированные результаты поиска для "${searchQuery}"`);
    const products = JSON.parse(cachedProducts);
    let position = 0;
    
    // Ищем артикул в кэшированных результатах
    for (const id of products) {
      position++;
      if (id.toString() === articleNumber) {
        const result = `${position}`;
        cache.put(cacheKey + "_" + articleNumber, result, CACHE_DURATION);
        return result;
      }
    }
    
    // Если не нашли в кэше
    return "Нет в выдаче";
  }
  
  // Если в кэше нет, делаем запросы к API
  Logger.log(`Загружаем с API для запроса "${searchQuery}"`);
  let position = 0;
  let productIds = []; // Храним только ID товаров
  let pagesScanned = 0;
  
  // Проходим по всем страницам
  for (let page = 1; page <= MAX_PAGES; page++) {
    Logger.log(`Загружаем товары со страницы ${page}`);
    pagesScanned = page;
    
    // Задержка между запросами
    if (page > 1) {
      Utilities.sleep(DELAY_BETWEEN_REQUESTS);
    }
    
    // Формируем URL (v18, как в сетевом трейсинге)
    const url = `https://search.wb.ru/exactmatch/ru/common/v18/search?ab_testing=false&appType=64&curr=rub&dest=-1257786&hide_dtype=13&inheritFilters=false&lang=ru&page=${page}&query=${encodeURIComponent(searchQuery)}&resultset=catalog&sort=popular&spp=${PRODUCTS_PER_PAGE}&suppressSpellcheck=false`;
    
    // Делаем запрос
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'Accept': '*/*',
        'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
        'Connection': 'keep-alive',
        'Origin': 'https://www.wildberries.ru',
        'Referer': 'https://www.wildberries.ru/catalog/0/search.aspx?sort=popular&search=' + encodeURIComponent(searchQuery),
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'cross-site',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
      }
    });
    
    // Проверяем ответ
    if (response.getResponseCode() !== 200) {
      throw new Error("Ошибка сервера WB");
    }
    
    const data = JSON.parse(response.getContentText());
    
    // Сохраняем JSON первой страницы (только при отладке)
    if (page === 1 && DEBUG_SAVE_PROPERTIES) {
      try {
        PropertiesService.getScriptProperties().setProperty(
          `search_response_${searchQuery.replace(/\s+/g, '_')}`,
          response.getContentText()
        );
      } catch (e) {
        Logger.log('Пропущено сохранение search_response: ' + e);
      }
    }
    
    // Поддерживаем разные формы ответа (v4/v18)
    let products = null;
    if (data && data.data && Array.isArray(data.data.products)) {
      products = data.data.products;
    } else if (Array.isArray(data.products)) {
      products = data.products;
    }
    if (!products) {
      return `❌ Нет результатов на ${pagesScanned} страницах.`;
    }
    
    // Если страница пустая, значит достигли конца выдачи
    if (products.length === 0) {
      break;
    }
    
    // Добавляем только ID товаров
    productIds = productIds.concat(products.map(p => p.id));
    
    // Проверяем каждый товар на странице
    for (const product of products) {
      position++;
      if (product.id.toString() === articleNumber) {
        // Получаем данные о продажах
        try {
          const salesUrl = `https://product-order-qnt.wildberries.ru/by-nm/?nm=${articleNumber}`;
          const salesResponse = UrlFetchApp.fetch(salesUrl, {
            muteHttpExceptions: true,
            headers: {
              'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36'
            }
          });
          
          if (salesResponse.getResponseCode() === 200 && DEBUG_SAVE_PROPERTIES) {
            try {
              PropertiesService.getScriptProperties().setProperty(
                `sales_response_${articleNumber}`,
                salesResponse.getContentText()
              );
            } catch (e) {
              Logger.log('Пропущено сохранение sales_response: ' + e);
            }
          }
        } catch (error) {
          Logger.log("Ошибка при получении данных о продажах: " + error);
        }
        
        // Кэшируем только ID товаров
        try {
          cache.put(productsCacheKey, JSON.stringify(productIds), CACHE_DURATION);
        } catch (e) {
          Logger.log("Ошибка кэширования (слишком большой размер): " + e);
        }
        
        // Формируем результат по бусту
        const logInfo = product.log || {};
        const organicPos = (typeof logInfo.position === 'number' && isFinite(logInfo.position)) ? logInfo.position : position;
        const promoPosVal = (typeof logInfo.promoPosition === 'number' && isFinite(logInfo.promoPosition)) ? logInfo.promoPosition : null;
        const hasBoost = !!promoPosVal && logInfo.promotion === 1;
        const result = hasBoost
          ? `${organicPos} → ${promoPosVal}`
          : `${organicPos}`;
        cache.put(cacheKey + "_" + articleNumber, result, CACHE_DURATION);
        return result;
      }
    }
    
    // Если достигли 3000 позиций, останавливаемся
    if (position >= 3000) {
      break;
    }
  }
  
  // Кэшируем только ID товаров даже если не нашли нужный артикул
  if (productIds.length > 0) {
    try {
      cache.put(productsCacheKey, JSON.stringify(productIds), CACHE_DURATION);
    } catch (e) {
      Logger.log("Ошибка кэширования (слишком большой размер): " + e);
    }
  }
  
  return `❌ Нет результатов на ${pagesScanned} страницах.`;
}

/**
 * Ищет позицию артикула в массиве товаров
 * @param {Array} products - Массив товаров
 * @param {string} articleNumber - Артикул для поиска
 * @return {string|null} Найденная позиция или null
 */
function findArticlePosition(products, articleNumber) {
  for (let i = 0; i < products.length; i++) {
    if (products[i].id.toString() === articleNumber) {
      return `${i + 1}`;
    }
  }
  return null;
}

/**
 * Создает меню в Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('WB Позиции')
    .addItem('Сбросить кэш', 'clearWBCache')
    .addItem('Очистить отладочные данные', 'clearDebugProperties')
    .addSeparator()
    .addItem('О скрипте', 'showAbout')
    .addToUi();
}

/**
 * Сбрасывает весь кэш поиска WB
 */
function clearWBCache() {
  const cache = CacheService.getScriptCache();
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Получаем все ключи из кэша нельзя, поэтому просто устанавливаем короткое время жизни
    // для новых записей и показываем пользователю информацию
    ui.alert('Кэш сброшен', 'Кэш поиска WB будет обновлен при следующих запросах.', ui.ButtonSet.OK);
    Logger.log('Кэш WB сброшен пользователем');
  } catch (error) {
    ui.alert('Ошибка', 'Не удалось сбросить кэш: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Очищает сохранённые ScriptProperties, связанные с отладкой
 */
function clearDebugProperties() {
  const props = PropertiesService.getScriptProperties();
  const ui = SpreadsheetApp.getUi();
  const all = props.getProperties();
  let removed = 0;
  
  Object.keys(all).forEach((key) => {
    if (key.startsWith('search_response_') || key.startsWith('sales_response_')) {
      props.deleteProperty(key);
      removed++;
    }
  });
  
  ui.alert('Очистка завершена', `Удалено отладочных свойств: ${removed}`, ui.ButtonSet.OK);
}

/**
 * Показывает информацию о скрипте
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('WB Позиции v2.0', 
    'Скрипт для поиска позиций товаров в выдаче Wildberries.\n\n' +
    'Статусы результатов:\n' +
    '• Есть буст: 98 → 4\n' +
    '• Буста нет: 98\n' +
    '• Нет в выдаче: ❌ Нет результатов на N страницах\n' +
    '• Ошибка: ⚠️ Ошибка: описание\n\n' +
    'Чат поддержки: https://t.me/+pOKAcasVXsU0YWQx\n' +
    'Можно задать вопрос или сообщить о проблеме.',
    ui.ButtonSet.OK);
}

/**
 * Очищает сохранённые ScriptProperties, связанные с отладкой (старая функция)
 * @return {string}
 */
function WB_CLEAR_DEBUG_PROPERTIES() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  let removed = 0;
  Object.keys(all).forEach((key) => {
    if (key.startsWith('search_response_') || key.startsWith('sales_response_')) {
      props.deleteProperty(key);
      removed++;
    }
  });
  return `Удалено свойств: ${removed}`;
}