#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
wb-parser-playwright — парсер каталога Wildberries через реальный браузер Chromium.

Зачем Playwright, а не requests
-------------------------------
Wildberries закрывает поисковое API (``search.wb.ru/exactmatch/.../search``)
антиботом JVO1 VM: обычный HTTP-клиент получает HTTP 403 с пустым телом.
Браузер же сам выполняет challenge-JS, считает fingerprint и ставит cookie
с токеном ``u`` — поэтому запрос из контекста страницы проходит.

Как здесь устроен обход
-----------------------
1. Chromium запускается в «живом» режиме (headed по умолчанию, ``--headless`` — флаг).
2. На страницу навешивается stealth (1.x или 2.x) + доп. патчи fingerprint
   (navigator.webdriver, WebGL, canvas, chrome.runtime, permissions, languages).
3. Перед навигацией выполняется «человеческое» поведение: движения мыши,
   плавный скролл, ввод по символам, случайные паузы ``random.uniform(1.5, 4.0)``.
4. Данные товаров НЕ скроллятся по DOM как основной путь — они перехватываются
   из сетевого ответа ``search.wb.ru/.../search`` прямо в браузере
   (``page.on('response')``). Это надёжнее: не зависит от вёрстки карточек.
5. DOM-парсинг карточек остаётся как резервный путь.

ГДЕ ИСКАТЬ НОВЫЕ СЕЛЕКТОРЫ / ЭНДПОИНТ, ЕСЛИ WB ПОМЕНЯЕТ РАЗМЕТКУ
------------------------------------------------------------------
* Поисковая строка: константа ``SEARCH_INPUT_SELECTORS``.
  Быстрый способ найти новый селектор: открыть сайт, F12 -> Ctrl+Shift+C ->
  кликнуть по полю поиска -> скопировать CSS-селектор.
* Карточки товаров (резервный путь): константа ``CARD_SELECTORS``.
* Структура JSON: функция ``extract_products()``.
* Признаки блокировки в DOM: константа ``BLOCK_MARKERS``.
* Эндпоинт поиска жёстко НЕ зашит: фильтр Network -> Fetch/XHR -> строка
  с ``search.wb.ru`` попадёт в перехватчик автоматически.
"""

from __future__ import annotations

import argparse
import logging
import os
import random
import re
import sys
import time
from dataclasses import dataclass
from logging.handlers import RotatingFileHandler
from typing import Any, Dict, List, Optional
from urllib.parse import quote_plus, urlparse

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    from playwright.sync_api import (
        Error as PlaywrightError,
        TimeoutError as PlaywrightTimeoutError,
        sync_playwright,
    )
except ImportError:  # pragma: no cover - зависит от окружения
    print(
        "Не найден Playwright. Установите зависимости:\n"
        "    pip install -r requirements.txt\n"
        "    playwright install chromium",
        file=sys.stderr,
    )
    raise

# --- playwright-stealth: поддерживаем и 2.x, и 1.x ----------------------- #
# 2.x API: Stealth().use_sync(sync_playwright()) — обёртка контекста,
#          Stealth().apply_stealth_sync(page)   — разовое применение.
# 1.x API: модуль-функция stealth_sync(page).
try:
    from playwright_stealth import Stealth  # playwright-stealth >= 2.0

    STEALTH_API = "2.x"
except ImportError:  # pragma: no cover - старые версии
    Stealth = None  # type: ignore[assignment]
    STEALTH_API = "none"

try:  # playwright-stealth 1.x
    from playwright_stealth import stealth_sync as _legacy_stealth_sync  # type: ignore

    if STEALTH_API == "none":
        STEALTH_API = "1.x"
except ImportError:
    _legacy_stealth_sync = None  # type: ignore[assignment]

if STEALTH_API == "none":
    print(
        "Не найден playwright-stealth. Установите: pip install -r requirements.txt\n"
        "Парсер продолжит работу на встроенных патчах fingerprint, "
        "но вероятность блокировки выше.",
        file=sys.stderr,
    )


# --------------------------------------------------------------------------- #
# КОНСТАНТЫ
# --------------------------------------------------------------------------- #

BASE_URL = "https://www.wildberries.ru/"
SEARCH_PAGE_URL = "https://www.wildberries.ru/catalog/search/?search={query}&page={page}"

# Регион (Москва / СПб) — используется как ориентир и для возможного URL-API.
DEST_MOSCOW = -1257786
DEST_SPB = -102281

# --- Селекторы (см. докстринг: где искать новые) --------------------------- #
SEARCH_INPUT_SELECTORS = [
    "input.search-bar__input",
    "input#search-bar__input",
    ".search-bar input[type='text']",
    "header input[type='search']",
    "input[placeholder*='Поиск']",
    "input[aria-label*='Поиск']",
]

SUBMIT_SELECTORS = [
    ".search-bar__btn-search",
    "button.search-bar__submit",
    "button[type='submit']",
]

CARD_SELECTORS = [
    "div.product-card",
    "a.product-card__link",
    "article.product-card",
]

# Экран JVO1 «Проверяем браузер»: WB грузит behavior-tracker и ждёт, пока
# страница «поживёт» (движения мыши, скролл). Пока этого нет — выдачи не будет.
CHALLENGE_MARKERS = [
    "Проверяем браузер",
    "Проверка браузера",
    "Проверяем вас",
    "Подождите, идёт проверка",
]

# Признаки того, что нас заблокировали (смотрим в тексте страницы).
BLOCK_MARKERS = [
    # Антибот / проверка «вы человек»
    "Доступ ограничен",
    "Вы точно человек",
    "Подтвердите, что запросы отправляли вы",
    "Подтвердите, что вы не робот",
    "jvo",
    "antibot",
    "challenge",
    # Экран WB для недоверенного IP (VPN / дата-центр / мобильный оператор)
    "нужно выключить VPN",
    "не смогли загрузить страницу",
    "используйте QR-код",
    "Попробовать снова",
]

# Реальные User-Agent'ы последних версий Chrome.
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
]

VIEWPORT = {"width": 1920, "height": 1080}

EXCEL_HEADERS = [
    "Название",
    "Цена со скидкой, ₽",
    "Цена без скидки, ₽",
    "Скидка, %",
    "Рейтинг",
    "Отзывы",
    "ID товара",
    "Ссылка",
    "Продавец",
    "Бренд",
]
EXCEL_WIDTHS = [62, 20, 20, 12, 10, 10, 14, 52, 28, 20]

STEALTH_INIT_SCRIPT = """
(() => {
  // 1. navigator.webdriver -> undefined
  try {
    Object.defineProperty(Navigator.prototype, 'webdriver', {
      get: () => undefined, configurable: true,
    });
  } catch (e) {}

  // 2. Языки браузера
  try {
    Object.defineProperty(navigator, 'languages', {
      get: () => ['ru-RU', 'ru', 'en-US', 'en'], configurable: true,
    });
  } catch (e) {}

  // 3. WebGL: правдоподобные параметры вместо llvmpipe/SwiftShader
  try {
    const patch = (proto) => {
      const getParam = proto.getParameter;
      proto.getParameter = function (param) {
        if (param === 37445) return 'Intel Inc.';
        if (param === 37446) return 'Intel Iris OpenGL Engine';
        return getParam.call(this, param);
      };
    };
    patch(window.WebGLRenderingContext.prototype);
    if (window.WebGL2RenderingContext) patch(window.WebGL2RenderingContext.prototype);
  } catch (e) {}

  // 4. Canvas fingerprint: лёгкий, но воспроизводимый шум
  try {
    const toDataURL = HTMLCanvasElement.prototype.toDataURL;
    const getImageData = CanvasRenderingContext2D.prototype.getImageData;
    HTMLCanvasElement.prototype.toDataURL = function (...args) {
      const ctx = this.getContext('2d');
      if (ctx) {
        const data = ctx.getImageData(0, 0, 1, 1);
        data.data[0] = (data.data[0] + 1) % 256;
        ctx.putImageData(data, 0, 0);
      }
      return toDataURL.apply(this, args);
    };
    CanvasRenderingContext2D.prototype.getImageData = function (...args) {
      const result = getImageData.apply(this, args);
      if (result && result.data && result.data.length) {
        result.data[0] = (result.data[0] + 1) % 256;
      }
      return result;
    };
  } catch (e) {}

  // 5. chrome.runtime есть только у настоящего браузера
  try {
    if (!window.chrome) window.chrome = {};
    if (!window.chrome.runtime) window.chrome.runtime = {};
  } catch (e) {}

  // 6. Permissions: Notification всегда 'default', как в Chrome
  try {
    const origQuery = navigator.permissions && navigator.permissions.query;
    if (origQuery) {
      navigator.permissions.query = (params) =>
        params && params.name === 'notifications'
          ? Promise.resolve({ state: Notification.permission, onchange: null })
          : origQuery.call(navigator.permissions, params);
    }
  } catch (e) {}
})();
"""

LOGGER_NAME = "wb_parser_pw"


# --------------------------------------------------------------------------- #
# ЛОГИРОВАНИЕ
# --------------------------------------------------------------------------- #

def _force_utf8_console() -> None:
    """Переключает вывод консоли на UTF-8, иначе в cmd будут «кракозябры»."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def setup_logging(verbose: bool = False, log_file: str = "parser.log") -> logging.Logger:
    _force_utf8_console()

    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()
    logger.propagate = False

    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%H:%M:%S")

    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.DEBUG if verbose else logging.INFO)
    console.setFormatter(fmt)
    logger.addHandler(console)

    try:
        file_handler = RotatingFileHandler(
            log_file, maxBytes=2_000_000, backupCount=3, encoding="utf-8"
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(fmt)
        logger.addHandler(file_handler)
    except OSError:
        pass

    return logger


# --------------------------------------------------------------------------- #
# МОДЕЛЬ ТОВАРА
# --------------------------------------------------------------------------- #

@dataclass
class Product:
    name: str
    sale_price: Optional[float]
    price: Optional[float]
    discount: Optional[int]
    rating: Optional[float]
    feedbacks: Optional[int]
    product_id: Optional[int]
    url: str
    supplier: str
    brand: str

    def as_row(self) -> List[Any]:
        return [
            self.name, self.sale_price, self.price, self.discount, self.rating,
            self.feedbacks, self.product_id, self.url, self.supplier, self.brand,
        ]


def _to_float(value: Any) -> Optional[float]:
    try:
        if value is None or value == "":
            return None
        return round(float(value), 2)
    except (TypeError, ValueError):
        return None


def _to_int(value: Any) -> Optional[int]:
    try:
        if value is None or value == "":
            return None
        return int(float(value))
    except (TypeError, ValueError):
        return None


def _to_str(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split())


def _kopecks_to_rubles(value: Any) -> Optional[float]:
    """WB отдаёт цены в копейках: 7299000 -> 72990.00 ₽."""
    amount = _to_float(value)
    return None if amount is None else round(amount / 100.0, 2)


def _is_meaningful_card(card: Dict[str, Any]) -> bool:
    """
    Отсеивает мусорные «карточки» из DOM-фолбэка.

    В разметке выдачи много служебных блоков, у которых innerText состоит
    из скидки («−64%») или срока доставки («Послезавтра») — это не товар.
    Настоящее название — достаточно длинная строка с буквами.
    """
    if not _to_int(card.get("id")):
        return False
    name = _to_str(card.get("name"))
    if len(name) < 10:
        return False
    return sum(ch.isalpha() for ch in name) >= 6


# --- Разбор текстовой карточки (DOM-фолбэк) ------------------------------- #

_RE_PRICE = re.compile(r"(\d[\d\s\u00a0]*)\s*\u20bd")
_RE_DISCOUNT = re.compile(r"[-−](\d{1,3})\s*%")
# Рейтинг идёт перед счётчиком отзывов: «4,5 · 58 оценок» или «5 · 1 оценка».
# При малом числе отзывов WB пишет ЦЕЛОЕ число («5»), поэтому запятая не нужна.
# Порядок важен: _RE_RATING применяется ПЕРВЫМ и съедает «5 · 5 оценок» целиком,
# иначе _RE_FEEDBACKS успеет убрать «5 оценок» и рейтинг потеряется.
_RE_RATING = re.compile(
    r"(\d(?:[.,]\d)?)\s*[·•*]\s*\d[\d\s\u00a0]*\s*(?:оцен\w*|отзыв\w*)",
    re.IGNORECASE,
)
_RE_FEEDBACKS = re.compile(r"(\d[\d\s\u00a0]*)\s*(?:оцен\w*|отзыв\w*)", re.IGNORECASE)
_RE_DATE = re.compile(r"(послезавтра|завтра|сегодня|\d{1,2}\s+[а-яё]{3,})", re.IGNORECASE)
# Бренд и название WB разделяет косой чертой: «Apple / Смартфон iPhone 15».
_RE_BRAND_NAME = re.compile(r"([^/\n]{2,40}?)\s*/\s*(.+)$")
# Служебные плашки WB — не часть бренда или названия.
_BADGES = re.compile(
    r"(ХОРОШАЯ ЦЕНА РАСПРОДАЖА|РАСПРОДАЖА|Осталось мало|BE ORIGINAL|с WB Кошельком)",
    re.IGNORECASE,
)


def _clean_number(raw: str) -> Optional[float]:
    """«5 848» / «5\u00a0848» -> 5848.0"""
    digits = re.sub(r"[^\d,.]", "", raw).replace(",", ".")
    try:
        return float(digits)
    except ValueError:
        return None


def parse_card_text(card: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    Разбирает текст карточки товара в dict полей WB.

    Реальный вид текста (innerText карточки):
        «−70%5 848 ₽20 373 ₽с WB КошелькомIKON/Шины зимние R15 ...4,8· 93 оценкиЗавтра»

    Отсюда берём: id, цены, скидку, бренд, название, рейтинг, отзывы.

    ВАЖНО: это запасной путь. Основной — перехват JSON из сети (точные данные).
    Здесь регулярки описывают текущую вёрстку; если WB её поменяет,
    правьте именно этот блок (смотрите HTML карточки через DevTools).
    """
    product_id = _to_int(card.get("id"))
    if not product_id:
        return None

    text = _to_str(card.get("text"))
    if len(text) < 15:
        return None

    # Цены: первая — со скидкой, вторая — без. «₽» разделяет значения.
    prices = [_clean_number(m) for m in _RE_PRICE.findall(text)]
    prices = [p for p in prices if p and p > 0]
    sale_price = prices[0] if prices else None
    base_price = prices[1] if len(prices) > 1 else None

    # Скидка: «−70%»
    discount = _to_int(_RE_DISCOUNT.search(text).group(1)) if _RE_DISCOUNT.search(text) else None

    # --- Отзывы: «93 оценки», «5 618 отзывов» --------------------------- #
    # Считаем ДО разбора рейтинга: дальше текст будет чиститься.
    feedbacks = None
    m_fb = _RE_FEEDBACKS.search(text)
    if m_fb:
        feedbacks = _to_int(_clean_number(m_fb.group(1)))

    # --- Рейтинг: «4,5 · 58 оценок» или «5 · 1 оценка» ------------------ #
    # Отсекаем мусор вида «55 · 93 оценки» (размер шин) — рейтинг не может быть > 5.
    rating = None
    m = _RE_RATING.search(text)
    if m:
        candidate = _to_float(m.group(1).replace(",", "."))
        if candidate is not None and 0 <= candidate <= 5:
            rating = candidate
        else:
            # Это НЕ рейтинг. Убираем всю конструкцию, чтобы число
            # не попало в название товара.
            text = text.replace(m.group(0), " ")

    # --- Бренд и название ------------------------------------------------ #
    # Приоритет: aria-label ссылки (там WB пишет «Бренд Название»),
    # затем data-атрибут, затем разбор текста по «Бренд / Название».
    brand, name = "", ""

    cleaned = _BADGES.sub(" ", text)
    # Порядок обязателен: сначала убираем целые конструкции «рейтинг · отзывы»,
    # затем оставшийся счётчик отзывов, даты, скидку и цены.
    cleaned = _RE_RATING.sub(" ", cleaned)
    cleaned = _RE_FEEDBACKS.sub(" ", cleaned)
    cleaned = _RE_DATE.sub(" ", cleaned)
    cleaned = _RE_DISCOUNT.sub(" ", cleaned)
    cleaned = _RE_PRICE.sub(" ", cleaned)
    cleaned = _to_str(cleaned)

    m = _RE_BRAND_NAME.search(cleaned)
    if m:
        brand, name = m.group(1), m.group(2)
    else:
        name = cleaned

    # Убираем мусор по краям: «55·», «·», служебные хвосты.
    name = _to_str(name).strip(" ·•*-—·")
    brand = _to_str(brand).strip(" ·•*-—/")

    # Если «бренд» получился слишком длинным — это часть названия.
    if brand and len(brand) > 25:
        name = f"{brand} {name}".strip()
        brand = ""

    # Продавец: в новой выдаче не показывается, оставляем пустым
    supplier = ""

    return {
        "id": product_id,
        "name": _to_str(name),
        "brand": _to_str(brand),
        "supplier": supplier,
        "salePriceU": None if sale_price is None else int(sale_price * 100),
        "priceU": None if base_price is None else int(base_price * 100),
        "sale": discount,
        "rating": None if rating is None else int(rating * 10),  # rating в WB = x10
        "feedbacks": feedbacks,
    }


def parse_product(raw: Dict[str, Any]) -> Optional[Product]:
    """Сырой dict товара из ответа WB -> Product. None, если это не товар."""
    if not isinstance(raw, dict):
        return None

    product_id = _to_int(raw.get("id"))
    name = _to_str(raw.get("name"))
    if not product_id or not name:
        return None

    sale_price = _kopecks_to_rubles(raw.get("salePriceU"))
    price = _kopecks_to_rubles(raw.get("priceU"))

    # У части карточек цены лежат в prices{} — берём оттуда тем же способом.
    prices = raw.get("prices") if isinstance(raw.get("prices"), dict) else {}
    if sale_price is None:
        sale_price = _kopecks_to_rubles(prices.get("product"))
    if price is None:
        price = _kopecks_to_rubles(prices.get("basic"))

    discount = _to_int(raw.get("sale"))
    if discount is None and price and sale_price and price > 0:
        discount = int(round((1 - sale_price / price) * 100))

    # WB отдаёт рейтинг как целое x10 (48 -> 4.8). Приводим к шкале 0–5.
    rating = _to_float(raw.get("rating"))
    if rating is not None and rating > 10:
        rating = round(rating / 10.0, 1)

    return Product(
        name=name,
        sale_price=sale_price,
        price=price,
        discount=discount,
        rating=rating,
        feedbacks=_to_int(raw.get("feedbacks")),
        product_id=product_id,
        url=f"https://www.wildberries.ru/catalog/{product_id}/detail.aspx",
        supplier=_to_str(raw.get("supplier")),
        brand=_to_str(raw.get("brand")),
    )


def extract_products(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Достаёт список товаров из JSON разных версий ответа WB.

    Если структура изменится — правьте ИМЕННО ЭТОТ МЕТОД.
    Актуальную форму ответа: F12 -> Network -> Fetch/XHR ->
    найти запрос ``search.wb.ru/.../search`` -> Response -> скопировать JSON.
    """
    if not isinstance(payload, dict):
        return []

    data = payload.get("data")
    if isinstance(data, list):
        return data
    if not isinstance(data, dict):
        return []

    for key in ("products", "teasers", "goods", "items"):
        items = data.get(key)
        if isinstance(items, list):
            return items
    return []


# --------------------------------------------------------------------------- #
# ПАРСЕР
# --------------------------------------------------------------------------- #

class WildberriesParser:
    """Синхронный Playwright-парсер каталога WB."""

    def __init__(
        self,
        headless: bool = False,
        proxy: Optional[str] = None,
        delay_range: tuple = (1.5, 4.0),
        retries: int = 3,
        timeout: int = 45_000,
        slow_mo: int = 0,
        dest: int = DEST_MOSCOW,
        logger: Optional[logging.Logger] = None,
        artifacts_dir: str = "output",
    ) -> None:
        self.headless = headless
        self.proxy_arg = proxy
        self.delay_range = delay_range
        self.retries = max(1, retries)
        self.timeout = timeout
        self.slow_mo = slow_mo
        self.dest = dest
        self.log = logger or logging.getLogger(LOGGER_NAME)
        self.artifacts_dir = artifacts_dir

        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

        self._pending: List[Any] = []      # перехваченные ответы search.wb.ru
        self.blocked = False                # WB отвечает 403/498 (антибот, IP в бане)
        self._blocked_reasons: List[str] = []

    # ------------------------------------------------------------------ #
    # Прокси
    # ------------------------------------------------------------------ #

    def _parse_proxy(self, raw: Optional[str]) -> Optional[Dict[str, str]]:
        """
        Разбирает прокси ``http://user:pass@host:port`` в формат Playwright:
            {"server": "http://host:port", "username": "user", "password": "pass"}
        Без логина/пароля — просто {"server": ...}.
        """
        if not raw:
            return None

        parsed = urlparse(raw if "://" in raw else f"http://{raw}")
        if not parsed.hostname:
            self.log.error("Не удалось разобрать прокси: %s", raw)
            return None

        server = f"{parsed.scheme or 'http'}://{parsed.hostname}"
        if parsed.port:
            server += f":{parsed.port}"

        result = {"server": server}
        if parsed.username:
            result["username"] = parsed.username
            result["password"] = parsed.password or ""
        self.log.info("Прокси: %s (auth: %s)", server, "да" if parsed.username else "нет")
        return result

    # ------------------------------------------------------------------ #
    # Запуск браузера
    # ------------------------------------------------------------------ #

    def _start_playwright(self) -> Any:
        """
        Запускает sync_playwright. Для playwright-stealth 2.x контекст
        оборачивается через Stealth().use_sync(...), чтобы патчи применялись
        ко всем создаваемым страницам.
        """
        if STEALTH_API == "2.x":
            stealth = Stealth(
                # Согласуем маскировку с UA/locale, которые ставим сами.
                navigator_languages_override=("ru-RU", "ru"),
                navigator_platform_override="Win32",
                webgl_vendor_override="Intel Inc.",
                webgl_renderer_override="Intel Iris OpenGL Engine",
            )
            self.log.debug("Stealth: use_sync + apply_stealth_sync")
            return stealth.use_sync(sync_playwright()).start()

        manager = sync_playwright()
        if STEALTH_API == "1.x" and hasattr(Stealth, "hook_playwright_context"):
            try:
                Stealth().hook_playwright_context(manager)  # type: ignore[operator]
            except Exception as exc:
                self.log.debug("hook_playwright_context не применился: %s", exc)
        return manager.start()

    def start(self) -> None:
        """Поднимает Chromium, контекст и страницу со всеми патчами."""
        proxy = self._parse_proxy(self.proxy_arg)

        self.log.info("=" * 60)
        self.log.info("Запуск Chromium (headless=%s, playwright-stealth=%s)",
                      self.headless, STEALTH_API)

        self.playwright = self._start_playwright()
        try:
            self.browser = self.playwright.chromium.launch(
                headless=self.headless,
                proxy=proxy,
                slow_mo=self.slow_mo,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--disable-features=IsolateOrigins,site-per-process",
                    "--disable-dev-shm-usage",
                    "--no-sandbox",
                    "--lang=ru-RU",
                ],
            )
        except PlaywrightError as exc:
            self.log.error("Не удалось запустить Chromium: %s", exc)
            self.log.error("Браузер не установлен? Выполните: playwright install chromium")
            self.stop()
            raise

        self.context = self.browser.new_context(
            viewport=VIEWPORT,
            user_agent=random.choice(USER_AGENTS),
            locale="ru-RU",
            timezone_id="Europe/Moscow",
            geolocation={"latitude": 55.7558, "longitude": 37.6173},  # Москва
            permissions=["geolocation"],
            java_script_enabled=True,
            ignore_https_errors=False,
        )
        self.context.set_default_timeout(self.timeout)
        self.context.set_default_navigation_timeout(self.timeout)

        # Патчи fingerprint применяются ДО любой навигации.
        self.context.add_init_script(STEALTH_INIT_SCRIPT)

        self.page = self.context.new_page()
        self._apply_stealth()

        # Перехват сетевых ответов к поисковому API.
        self.page.on("response", self._on_response)

    def _apply_stealth(self) -> None:
        """Применяет stealth к странице + ручные патчи поверх."""
        if STEALTH_API == "2.x":
            try:
                Stealth().apply_stealth_sync(self.page)
                self.log.debug("apply_stealth_sync применён")
            except Exception as exc:
                self.log.warning("apply_stealth_sync не применился (%s)", exc)
        elif STEALTH_API == "1.x" and _legacy_stealth_sync is not None:
            try:
                _legacy_stealth_sync(self.page)
                self.log.debug("stealth_sync (1.x) применён")
            except Exception as exc:
                self.log.warning("stealth_sync не применился (%s)", exc)

        # Дополнительно убираем самый явный след автоматизации.
        try:
            self.page.add_init_script(
                "Object.defineProperty(navigator,'webdriver',{get:()=>undefined});"
            )
        except PlaywrightError:
            pass

    def _on_response(self, response: Any) -> None:
        """
        Слушатель сети. Складываем ответы к ``search.wb.ru/.../search``,
        чтобы разобрать их ПОСЛЕ навигации: внутри обработчика нельзя
        вызывать блокирующие API вроде ``response.json()``.
        """
        try:
            url = response.url
            if "search.wb.ru" in url and "/search" in url:
                self.log.debug("Перехвачен поисковый ответ: %s", url[:120])
                self._pending.append(response)

            status = response.status
            if status in (403, 429, 498, 503):
                self._blocked_reasons.append(f"HTTP {status} <- {url[:140]}")
        except (PlaywrightError, AttributeError):
            pass

    def stop(self) -> None:
        """Корректное закрытие браузера."""
        for obj in (self.page, self.context, self.browser):
            try:
                if obj:
                    obj.close()
            except Exception:
                pass
        try:
            if self.playwright:
                self.playwright.stop()
        except Exception:
            pass
        self.page = self.context = self.browser = self.playwright = None

    # ------------------------------------------------------------------ #
    # Человеческое поведение
    # ------------------------------------------------------------------ #

    def human_pause(self, low: float = 0.4, high: float = 1.6) -> None:
        time.sleep(random.uniform(low, high))

    def human_mouse_move(self, steps: int = 14) -> None:
        """Плавно ведём мышь по случайной траектории."""
        if not self.page:
            return
        x = random.uniform(VIEWPORT["width"] * 0.15, VIEWPORT["width"] * 0.85)
        y = random.uniform(VIEWPORT["height"] * 0.15, VIEWPORT["height"] * 0.85)
        try:
            self.page.mouse.move(x, y, steps=steps)
        except PlaywrightError:
            pass

    def human_scroll(self, rounds: int = 4) -> None:
        """Прокручивает страницу колесом мыши, имитируя чтение выдачи."""
        if not self.page:
            return
        for _ in range(rounds):
            try:
                self.page.mouse.wheel(0, random.uniform(350, 900))
            except PlaywrightError:
                return
            self.human_pause(0.25, 0.9)
        try:
            self.page.mouse.wheel(0, -random.uniform(150, 400))
        except PlaywrightError:
            pass

    # ------------------------------------------------------------------ #
    # Диагностика
    # ------------------------------------------------------------------ #

    def save_artifacts(self, tag: str) -> None:
        """Сохраняет скриншот и HTML — по ним видно, что показал WB."""
        if not self.page:
            return
        try:
            os.makedirs(self.artifacts_dir, exist_ok=True)
            stamp = time.strftime("%Y%m%d_%H%M%S")
            safe_tag = re.sub(r"[^a-zA-Z0-9_]+", "_", tag)[:40]

            png = os.path.join(self.artifacts_dir, f"{stamp}_{safe_tag}.png")
            html = os.path.join(self.artifacts_dir, f"{stamp}_{safe_tag}.html")
            self.page.screenshot(path=png, full_page=False)
            with open(html, "w", encoding="utf-8") as fh:
                fh.write(self.page.content())
            self.log.warning("Сохранена диагностика: %s / %s", png, html)
        except (PlaywrightError, OSError) as exc:
            self.log.debug("Не удалось сохранить артефакты: %s", exc)

    def wait_for_wb_ready(self, timeout_ms: int = 60_000) -> bool:
        """
        Ждёт, пока пройдёт JVO1-челлендж («Проверяем браузер»).

        Пока идёт проверка, мы «живём»: двигаем мышь и слегка скроллим —
        behavior-tracker WB именно это и собирает. Без активности экран
        не исчезает и выдача не приходит.

        Возвращает True, если страница выглядит готовой.
        """
        deadline = time.time() + timeout_ms / 1000.0
        started = time.time()

        while time.time() < deadline:
            try:
                body = (self.page.inner_text("body") or "")
            except (PlaywrightError, AttributeError):
                return True

            if not any(marker.lower() in body.lower() for marker in CHALLENGE_MARKERS):
                # Экрана челленджа нет.
                self.log.debug("Челлендж пройден за %.1f с", time.time() - started)
                return True

            # Живём: водим мышью по экрану.
            self.human_mouse_move(steps=random.randint(6, 18))
            if random.random() < 0.3:
                try:
                    self.page.mouse.wheel(0, random.uniform(-200, 250))
                except PlaywrightError:
                    pass
            time.sleep(random.uniform(0.6, 1.4))

        self.log.warning("Челлендж JVO1 не пройден за %d с", timeout_ms // 1000)
        return False

    def _looks_blocked(self) -> bool:
        """Проверяем признаки блокировки в тексте страницы."""
        try:
            body = (self.page.inner_text("body") or "").lower()
        except (PlaywrightError, AttributeError):
            return False

        for marker in BLOCK_MARKERS:
            if marker.lower() in body:
                self.log.warning("Найден маркер блокировки в DOM: «%s»", marker)
                return True

        # Экран проверки браузера — это не бан, а незавершённый челлендж.
        for marker in CHALLENGE_MARKERS:
            if marker.lower() in body:
                self.log.warning("Страница всё ещё на экране челленджа: «%s»", marker)
                return True

        # Пустая выдача + короткая страница = почти наверняка отказ WB.
        if len(body.strip()) < 400:
            self.log.warning("Страница почти пустая (%d симв.) — вероятен отказ WB", len(body))
            return True
        return False

    def _log_wb_cookies(self) -> None:
        """Показывает, какие cookie WB поставил (токен JVO1 = 'u')."""
        try:
            cookies = self.context.cookies()
            names = [c["name"] for c in cookies if "wb" in c["name"] or c["name"] == "u"]
            self.log.debug("Cookie WB после загрузки: %s", ", ".join(names) or "нет")
        except (PlaywrightError, AttributeError):
            pass

    # ------------------------------------------------------------------ #
    # Навигация
    # ------------------------------------------------------------------ #

    def open_home(self) -> None:
        """
        Открывает главную и имитирует поведение человека.
        Бросает PlaywrightError, если все попытки провалились.
        """
        self.log.info("Открываю %s", BASE_URL)

        last_error: Optional[BaseException] = None
        for attempt in range(1, self.retries + 1):
            try:
                self.page.goto(BASE_URL, wait_until="domcontentloaded", timeout=30_000)
                self.page.wait_for_timeout(2000)
                self.wait_for_wb_ready()               # экран «Проверяем браузер»
                self.human_scroll(rounds=3)
                self._log_wb_cookies()
                return
            except (PlaywrightTimeoutError, PlaywrightError) as exc:
                last_error = exc
                first_line = str(exc).splitlines()[0] if str(exc) else exc.__class__.__name__
                self.log.warning("Главная не открылась (попытка %d/%d): %s",
                                 attempt, self.retries, first_line[:140])
                if attempt < self.retries:
                    time.sleep(5)

        if not self.headless:
            # Почти всегда headed-режим падает на machines без дисплея (SSH/CI).
            self.log.error(
                "Headed-режим не смог открыть страницу. Чаще всего причина — "
                "отсутствие интерактивного дисплея (SSH, CI, контейнер), "
                "а не отсутствие браузера. Запустите с флагом --headless."
            )
        raise last_error if last_error else PlaywrightError("navigation failed")

    def submit_search_via_ui(self, query: str) -> bool:
        """Вводит запрос в поисковую строку. False — если не нашли поле."""
        for selector in SEARCH_INPUT_SELECTORS:
            try:
                box = self.page.locator(selector).first
                if box.count() == 0:
                    continue
                box.wait_for(state="visible", timeout=8000)
                self.human_mouse_move()
                box.click()
                box.fill("")
                box.type(query, delay=random.randint(60, 140))  # печать по символам
                self.human_pause(0.4, 1.2)

                for submit in SUBMIT_SELECTORS:
                    btn = self.page.locator(submit).first
                    if btn.count() > 0:
                        try:
                            btn.click(timeout=3000)
                            return True
                        except PlaywrightError:
                            pass
                box.press("Enter")     # запасной путь — отправка по Enter
                return True
            except (PlaywrightError, AttributeError):
                continue

        self.log.warning("Поисковая строка не найдена ни по одному селектору")
        return False

    def goto_search(self, query: str, page_num: int) -> None:
        """
        Открывает страницу выдачи. Сначала пробуем живой путь (ввод + Enter),
        при неудаче — прямую навигацию по URL.
        """
        self._pending.clear()
        self.human_pause(0.5, 1.5)

        if page_num == 1:
            if not self.submit_search_via_ui(query):
                url = SEARCH_PAGE_URL.format(query=quote_plus(query), page=1)
                self.log.info("Перехожу по прямой ссылке: %s", url)
                self.page.goto(url, wait_until="domcontentloaded")
        else:
            # Пагинация прямым URL — стабильнее, чем клик по «Далее».
            url = SEARCH_PAGE_URL.format(query=quote_plus(query), page=page_num)
            self.log.info("Страница %s -> %s", page_num, url)
            self.page.goto(url, wait_until="domcontentloaded")

        self.page.wait_for_timeout(3000)
        self.wait_for_wb_ready()           # JVO1 VM на странице выдачи
        self.human_scroll(rounds=5)        # «дочитываем» выдачу, грузим карточки
        self.page.wait_for_timeout(2000)

    # ------------------------------------------------------------------ #
    # Сбор данных
    # ------------------------------------------------------------------ #

    def _drain_network_products(self) -> List[Dict[str, Any]]:
        """Разбирает перехваченные ответы search.wb.ru."""
        collected: List[Dict[str, Any]] = []
        responses, self._pending = self._pending, []

        for response in responses:
            try:
                payload = response.json()
            except (PlaywrightError, ValueError) as exc:
                self.log.debug("Ответ не JSON: %s", exc)
                continue

            items = extract_products(payload)
            if items:
                self.log.info("Из перехвата сети: %d товаров", len(items))
                collected.extend(items)
        return collected

    def _scrape_dom_fallback(self) -> List[Dict[str, Any]]:
        """
        Резервный путь: если перехват сети ничего не дал, разбираем карточки
        прямо из DOM.

        WB рендерит карточку примерно так:
            −70% | 5 848 ₽ | 20 373 ₽ | с WB Кошельком | IKON/Шины зимние ... |
            4,8 · 93 оценки | Завтра
        Из этой строки восстанавливаем цену, скидку, бренд, название,
        рейтинг, число отзывов и продавца.
        """
        for selector in CARD_SELECTORS:
            try:
                cards = self.page.locator(selector)
                if cards.count() == 0:
                    continue

                # id + полный текст карточки (data-nm — название в разметке WB)
                raw = cards.evaluate_all(
                    """
                    (nodes) => nodes.map((n) => {
                        const link = n.matches('a[href*="/catalog/"]')
                            ? n : n.querySelector('a[href*="/catalog/"]');
                        const m = link ? (link.getAttribute('href') || '').match(/\\/catalog\\/(\\d+)/) : null;
                        return {
                            id: m ? m[1] : null,
                            text: (n.innerText || '').replace(/\\s+/g, ' ').trim(),
                            dataNm: n.getAttribute('data-nm') || '',
                        };
                    })
                    """
                )

                items = [parse_card_text(r) for r in raw]
                items = [i for i in items if i]
                if items:
                    self.log.info("DOM-фолбэк (%s): %d карточек", selector, len(items))
                    return items
            except (PlaywrightError, AttributeError) as exc:
                self.log.debug("Селектор %s не сработал: %s", selector, exc)
        return []

    def collect_page(self) -> List[Dict[str, Any]]:
        """Достаёт сырые товары страницы: сначала сеть, при неудаче — DOM."""
        items = self._drain_network_products()
        if items:
            return items

        self.log.warning("Сетевой ответ не дал товаров, пробую DOM-парсинг")
        items = self._scrape_dom_fallback()
        if items:
            return items

        if self._looks_blocked():
            self.blocked = True
            self.save_artifacts("blocked")
            self.log.error(
                "WB не отдал товары: страница пустая или содержит маркер блокировки. "
                "Скриншот и HTML сохранены в папке %s/", self.artifacts_dir
            )
        return []

    # ------------------------------------------------------------------ #
    # Основной метод
    # ------------------------------------------------------------------ #

    def parse(self, query: str, pages: int = 1) -> List[Product]:
        """Обходит страницы выдачи и возвращает уникальные товары."""
        self.log.info("Wildberries (Playwright) | запрос: «%s» | страниц: %s",
                      query, pages)

        collected: Dict[int, Product] = {}

        for page_num in range(1, pages + 1):
            attempt_ok = False

            for attempt in range(1, self.retries + 1):
                try:
                    self.log.info("--- Страница %s (попытка %d/%d) ---",
                                  page_num, attempt, self.retries)
                    self.goto_search(query, page_num)
                    raw_items = self.collect_page()

                    if raw_items:
                        for raw in raw_items:
                            product = parse_product(raw)
                            if product and product.product_id not in collected:
                                collected[product.product_id] = product
                        self.log.info("Страница %s: уникальных товаров всего %d",
                                      page_num, len(collected))
                        attempt_ok = True
                        break

                    if self.blocked:
                        self.log.error("Блокировка: повторять бессмысленно")
                        break

                    self.log.warning("Страница %s: товаров нет", page_num)

                except KeyboardInterrupt:
                    self.log.warning("Прервано пользователем")
                    raise
                except (PlaywrightError, PlaywrightTimeoutError) as exc:
                    self.log.warning("Ошибка на странице %s (попытка %d/%d): %s",
                                     page_num, attempt, self.retries, exc)
                    self.save_artifacts(f"error_p{page_num}_a{attempt}")

                if attempt < self.retries:
                    time.sleep(5)          # ретрай с паузой 5 секунд
                    self.human_pause(1.0, 2.5)

            if not attempt_ok:
                self.log.error("Страница %s пропущена", page_num)

            if self.blocked:
                self.log.error("Останавливаюсь: доступ закрыт")
                break

            if page_num < pages:
                time.sleep(random.uniform(*self.delay_range))

        products = list(collected.values())
        if self._blocked_reasons:
            self.log.info("Замеченные HTTP-коды блокировки: %s",
                          "; ".join(sorted(set(self._blocked_reasons))))
        self.log.info("Итого уникальных товаров: %d", len(products))
        return products


# --------------------------------------------------------------------------- #
# EXCEL
# --------------------------------------------------------------------------- #

def resolve_output_path(raw_path: str) -> str:
    """Голое имя файла -> ./output/, путь с папками -> как есть."""
    if os.path.dirname(raw_path):
        return raw_path
    base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "output", os.path.basename(raw_path))


def save_to_excel(products: List[Product], output: str, logger: logging.Logger) -> str:
    path = resolve_output_path(output)
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Товары"

    ws.append(EXCEL_HEADERS)
    fill = PatternFill("solid", fgColor="8A2BE2")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for product in products:
        ws.append(product.as_row())

    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(EXCEL_HEADERS))}{ws.max_row}"

    for idx, width in enumerate(EXCEL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[0].alignment = Alignment(vertical="top", wrap_text=True)
        row[1].number_format = "# ##0.00"
        row[2].number_format = "# ##0.00"
        row[4].number_format = "0.0"

    wb.save(path)
    logger.info("Результат сохранён: %s (%d строк)", path, len(products))
    return path


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="parser.py",
        description="Парсер Wildberries на Playwright (обход JVO1 VM) -> Excel.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--query", "-q", required=True, help="Поисковый запрос, например 'iphone 15'")
    p.add_argument("--output", "-o", default="wb_result.xlsx",
                   help="Имя файла или путь. Голое имя попадёт в ./output/")
    p.add_argument("--pages", "-p", type=int, default=1,
                   help="Сколько страниц выдачи обойти (на странице ~100 товаров)")
    p.add_argument("--proxy", default=None,
                   help="Прокси в формате http://user:pass@host:port")
    p.add_argument("--headless", action="store_true",
                   help="Запуск без окна браузера (по умолчанию с окном)")
    p.add_argument("--verbose", "-v", action="store_true", help="Подробные логи (DEBUG)")
    # Дополнительно (не обязательны, но полезны):
    p.add_argument("--delay-min", type=float, default=1.5, help="Мин. пауза между страницами, сек")
    p.add_argument("--delay-max", type=float, default=4.0, help="Макс. пауза между страницами, сек")
    p.add_argument("--retries", type=int, default=3, help="Попыток на страницу при ошибке")
    p.add_argument("--timeout", type=int, default=45_000, help="Таймаут ожиданий, мс")
    p.add_argument("--slow-mo", type=int, default=0,
                   help="Замедление действий Playwright, мс (для отладки)")
    p.add_argument("--dest", type=int, default=DEST_MOSCOW,
                   help="Код региона WB (Москва -1257786, СПб -102281)")
    return p


def main(argv: Optional[List[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)
    log = setup_logging(verbose=args.verbose)

    if args.pages < 1:
        log.error("--pages должен быть >= 1")
        return 2
    if args.delay_max < args.delay_min:
        log.error("--delay-max должен быть >= --delay-min")
        return 2

    parser = WildberriesParser(
        headless=args.headless,
        proxy=args.proxy,
        delay_range=(args.delay_min, args.delay_max),
        retries=args.retries,
        timeout=args.timeout,
        slow_mo=args.slow_mo,
        dest=args.dest,
        logger=log,
    )

    try:
        parser.start()
        parser.open_home()
        products = parser.parse(args.query, args.pages)
    except KeyboardInterrupt:
        log.warning("Прервано пользователем (Ctrl+C)")
        return 130
    except PlaywrightError as exc:
        log.error("Playwright: %s", str(exc).splitlines()[0] if str(exc) else exc)
        log.error("Подсказки:")
        log.error("  1) Браузер не скачан — выполните: playwright install chromium")
        log.error("  2) Нет интерактивного дисплея — добавьте флаг --headless")
        log.error("  3) Нет сети до wildberries.ru или сработал блок IP (403/498)")
        return 4
    finally:
        parser.stop()

    if not products:
        if parser.blocked:
            log.error("Итог: сбор не выполнен — WB заблокировал доступ (403/498).")
            log.error("Скриншот и HTML лежат в папке output/ — посмотрите, что показал сайт.")
            return 3
        log.error("Ничего не найдено. Проверьте запрос и запустите с --verbose")
        return 1

    try:
        save_to_excel(products, args.output, log)
    except PermissionError:
        log.error("Файл занят Excel — закройте его и повторите.")
        return 1
    except OSError as exc:
        log.error("Не удалось сохранить файл: %s", exc)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
