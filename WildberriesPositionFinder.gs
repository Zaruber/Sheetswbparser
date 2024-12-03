// Константы
const DELAY_BETWEEN_REQUESTS = 1500; // 1.5 секунды между запросами
const PRODUCTS_PER_PAGE = 24; // Количество товаров на странице как в Python версии
const MAX_PAGES = 30; // Максимум 30 страниц
const CACHE_DURATION = 1800; // 30 минут в секундах
const CACHE_PREFIX = "wb_search_"; // Префикс для кэша
const PRODUCTS_CACHE_PREFIX = "wb_products_"; // Префикс для кэша продуктов

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
    return "❌ " + error.toString();
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
        const result = `🎯 ${position}`;
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
  
  // Проходим по всем страницам
  for (let page = 1; page <= MAX_PAGES; page++) {
    Logger.log(`Загружаем товары со страницы ${page}`);
    
    // Задержка между запросами
    if (page > 1) {
      Utilities.sleep(DELAY_BETWEEN_REQUESTS);
    }
    
    // Формируем URL
    const url = `https://search.wb.ru/exactmatch/ru/common/v4/search?appType=1&curr=rub&dest=-455460&page=${page}&query=${encodeURIComponent(searchQuery)}&resultset=catalog&sort=popular&spp=${PRODUCTS_PER_PAGE}&suppressSpellcheck=false`;
    
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
    
    // Сохраняем JSON первой страницы
    if (page === 1) {
      PropertiesService.getScriptProperties().setProperty(
        `search_response_${searchQuery.replace(/\s+/g, '_')}`,
        response.getContentText()
      );
    }
    
    if (!data.data || !data.data.products) {
      return "❌ Нет результатов";
    }
    
    const products = data.data.products;
    
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
              'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
          });
          
          if (salesResponse.getResponseCode() === 200) {
            PropertiesService.getScriptProperties().setProperty(
              `sales_response_${articleNumber}`,
              salesResponse.getContentText()
            );
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
        
        const result = `🎯 ${position}`;
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
  
  return "Нет в выдаче";
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
      return `🎯 ${i + 1}`;
    }
  }
  return null;
}
