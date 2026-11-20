"""
ToolRent («Монтаж») — Flask + Microsoft Access (pyodbc).

Запуск:  python server.py   →  http://127.0.0.1:5000

Этап 2: каталог читает категории, бренды и инструменты из database.accdb
(только чтение, ReadOnly=1). Если Access недоступен (нет pyodbc, драйвера или
файла), сайт НЕ падает: показывает демонстрационные данные из CATEGORIES/TOOLS
ниже, а причина печатается в консоль. Демо-данные остаются как резерв.
Регистрация, авторизация и бронирование через БД пока не реализованы.
JavaScript в проекте не используется.
"""
import os
import platform
import struct
import sys
from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, send_from_directory, abort, g, Response)

try:
    import pyodbc
except ImportError:          # без pyodbc работает демо-режим
    pyodbc = None

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
IMG_DIR = os.path.join(BASE_DIR, "img_tool")
DB_PATH = os.path.join(BASE_DIR, "database.accdb")

# static_folder=None: папки static нет, CSS и картинки отдаём явными маршрутами.
app = Flask(__name__, static_folder=None,
            template_folder=os.path.join(BASE_DIR, "templates"))
app.secret_key = "dev-only-change-before-deploy"

# ---------------------------------------------------------------------------
# ДЕМО-ДАННЫЕ (резерв на случай недоступного Access; удалить, когда подключение
# будет проверено). Формат прежний.
# ---------------------------------------------------------------------------
CATEGORIES = [
    {"slug": "perforators", "name": "Перфораторы", "img": "perforator.svg"},
    {"slug": "drills",      "name": "Дрели и шуруповёрты", "img": "drill.svg"},
    {"slug": "grinders",    "name": "Шлифовка и резка", "img": "grinder.svg"},
    {"slug": "power",       "name": "Электростанции", "img": "generator.svg"},
    {"slug": "measure",     "name": "Измерительные", "img": "laser.svg"},
    {"slug": "scaffold",    "name": "Леса и подмости", "img": "scaffold.svg"},
]

TOOLS = [
    {"id": 1, "brand": "Makita", "name": "Перфоратор SDS-plus HR2470", "cat": "perforators",
     "price": 900, "deposit": 8000, "qty": 6, "img": "perforator.svg",
     "desc": "Универсальный перфоратор для сверления бетона и кирпича, а также лёгкого долбления. Подходит для монтажных работ и ремонта квартир.",
     "specs": [("Тип патрона", "SDS-plus"), ("Мощность", "780 Вт"), ("Энергия удара", "2,4 Дж"), ("Режимы", "сверление, ударное сверление, долбление")]},
    {"id": 2, "brand": "Bosch", "name": "Перфоратор SDS-max GBH 8-45", "cat": "perforators",
     "price": 2100, "deposit": 20000, "qty": 2, "img": "perforator.svg",
     "desc": "Тяжёлый перфоратор для демонтажа и сверления больших отверстий в железобетоне. Работа на строительных объектах.",
     "specs": [("Тип патрона", "SDS-max"), ("Мощность", "1500 Вт"), ("Энергия удара", "12,5 Дж"), ("Вес", "~8 кг")]},
    {"id": 3, "brand": "DeWalt", "name": "Аккумуляторная дрель-шуруповёрт DCD791", "cat": "drills",
     "price": 700, "deposit": 6000, "qty": 8, "img": "drill.svg",
     "desc": "Компактный бесщёточный шуруповёрт с двумя аккумуляторами. Для сборки конструкций, монтажа и бытового ремонта.",
     "specs": [("Напряжение", "18 В"), ("Патрон", "быстрозажимной 13 мм"), ("Комплект", "2 аккумулятора, зарядное устройство"), ("Скорости", "2")]},
    {"id": 4, "brand": "Metabo", "name": "Углошлифовальная машина W 24-230", "cat": "grinders",
     "price": 1100, "deposit": 9000, "qty": 4, "img": "grinder.svg",
     "desc": "Мощная болгарка под диск 230 мм для резки металла, камня и зачистки поверхностей.",
     "specs": [("Диск", "230 мм"), ("Мощность", "2400 Вт"), ("Защита", "кожух, защита от перезапуска"), ("Вес", "~5 кг")]},
    {"id": 5, "brand": "Hilti", "name": "Алмазная пила для штроб DCH 300", "cat": "grinders",
     "price": 2800, "deposit": 25000, "qty": 1, "img": "grinder.svg",
     "desc": "Штроборез для аккуратных каналов в бетоне и кирпиче. Подключается к строительному пылесосу.",
     "specs": [("Диск", "2 × 125 мм"), ("Глубина реза", "до 30 мм"), ("Пылеудаление", "да"), ("Питание", "220 В")]},
    {"id": 6, "brand": "Hyundai", "name": "Бензиновый генератор HHY 7000FE", "cat": "power",
     "price": 3200, "deposit": 30000, "qty": 3, "img": "generator.svg",
     "desc": "Автономное электроснабжение стройплощадки: освещение, электроинструмент, насосы.",
     "specs": [("Мощность", "5,0 кВт"), ("Пуск", "электростартер"), ("Выходы", "220 В / 12 В"), ("Топливо", "бензин")]},
    {"id": 7, "brand": "Bosch", "name": "Лазерный нивелир GCL 2-50", "cat": "measure",
     "price": 1300, "deposit": 12000, "qty": 5, "img": "laser.svg",
     "desc": "Линейный лазерный нивелир для разметки стен, потолков и укладки плитки. В комплекте штатив.",
     "specs": [("Лучи", "2 линии"), ("Радиус работы", "до 20 м"), ("Самовыравнивание", "да"), ("Комплект", "штатив, кейс")]},
    {"id": 8, "brand": "Krause", "name": "Строительные леса, секция 2×2 м", "cat": "scaffold",
     "price": 600, "deposit": 5000, "qty": 12, "img": "scaffold.svg",
     "desc": "Модульные леса для фасадных и отделочных работ. Собираются без инструмента.",
     "specs": [("Размер секции", "2 × 2 м"), ("Материал", "сталь"), ("Нагрузка", "по паспорту изделия"), ("Сборка", "без инструмента")]},
]

DURATIONS = [1, 3, 7, 14, 30]


# ---------------------------------------------------------------------------
# Подключение к Access (только чтение)
# ---------------------------------------------------------------------------
ACCESS_DRIVER = "Microsoft Access Driver (*.mdb, *.accdb)"
DB_CONN_STR = f"DRIVER={{{ACCESS_DRIVER}}};DBQ={DB_PATH};ReadOnly=1;"
USE_ACCESS = pyodbc is not None and os.path.isfile(DB_PATH)

_logged_errors = set()


def log_db_error(exc):
    msg = str(exc)
    if msg not in _logged_errors:          # не засоряем консоль одинаковыми ошибками
        _logged_errors.add(msg)
        print("[ToolRent] Access недоступен, показываю демо-данные. Причина:", msg)


def query(sql, params=()):
    """SELECT с параметрами (?) → список словарей. Соединение открывается и закрывается на каждый запрос."""
    conn = pyodbc.connect(DB_CONN_STR)
    try:
        cur = conn.cursor()
        cur.execute(sql, list(params))
        names = [c[0] for c in cur.description]
        return [dict(zip(names, row)) for row in cur.fetchall()]
    finally:
        conn.close()


def use_db(db_func, demo_func):
    """Данные из Access; при любой ошибке — демо-данные (источник запоминается в g.source)."""
    if USE_ACCESS:
        try:
            result = db_func()
            if g.get("source") != "demo":
                g.source = "access"
            return result
        except Exception as e:
            log_db_error(e)
    g.source = "demo"
    return demo_func()


# ---------------------------------------------------------------------------
# Картинки и текст
# ---------------------------------------------------------------------------
CATEGORY_IMG_BY_ID = {}      # при желании: {id_category: "файл.svg"}
CATEGORY_IMG_WORDS = [        # иначе — по началу слов в названии категории
    (("перфор",), "perforator.svg"),
    (("дрел", "шуруп"), "drill.svg"),
    (("шлиф", "резк", "болгар", "пил"), "grinder.svg"),
    (("генер", "электростан"), "generator.svg"),
    (("нивел", "лазер", "измер"), "laser.svg"),
    (("леса", "подмост"), "scaffold.svg"),
]
DEFAULT_IMG = "perforator.svg"


def category_img(cat_id, name):
    if cat_id in CATEGORY_IMG_BY_ID:
        return CATEGORY_IMG_BY_ID[cat_id]
    words = (name or "").casefold().split()
    for stems, img in CATEGORY_IMG_WORDS:
        if any(w.startswith(s) for w in words for s in stems):
            return img
    return DEFAULT_IMG


def clean_text(v):
    """Текст из Access: None → "", переносы строк приводим к \\n, структуру не трогаем."""
    if v is None:
        return ""
    return str(v).replace("\r\n", "\n").replace("\r", "\n").strip("\n")


def tool_img(photo, cat_id):
    """Фото из img_tool по ИмяФото; нет файла — заглушка (иллюстрация категории), без JavaScript."""
    name = os.path.basename(str(photo or "").strip())
    if name and os.path.isfile(os.path.join(IMG_DIR, name)):
        return name
    cat = next((c for c in get_categories() if c["id"] == cat_id), None)
    return cat["img"] if cat else DEFAULT_IMG


# ---------------------------------------------------------------------------
# Источник «Access»
# ---------------------------------------------------------------------------
TOOL_SELECT = ("t.[id_tool] AS t_id, t.[Название] AS t_name, t.[Описание] AS t_desc, "
               "t.[Характеристики] AS t_specs, t.[ЦенаЗаСутки] AS t_price, t.[Залог] AS t_deposit, "
               "t.[КоличествоНаСкладе] AS t_qty, t.[ИмяФото] AS t_photo, t.[id_category] AS t_cat, "
               "t.[id_brand] AS t_brand_id, b.[Название] AS b_name")
TOOL_FROM = "[Инструмент] AS t LEFT JOIN [Бренд] AS b ON t.[id_brand] = b.[id_brand]"
TOOL_SORT = {                      # белый список: в SQL попадают только эти строки
    "": "t.[id_tool]",
    "id_desc": "t.[id_tool] DESC",
    "price_asc": "t.[ЦенаЗаСутки], t.[id_tool]",
    "price_desc": "t.[ЦенаЗаСутки] DESC, t.[id_tool]",
    "name": "t.[Название], t.[id_tool]",
}


def tool_from_row(r):
    return {
        "id": r["t_id"],
        "brand": clean_text(r["b_name"]),
        "brand_id": r["t_brand_id"],
        "name": clean_text(r["t_name"]),
        "cat": r["t_cat"],
        "price": r["t_price"] if r["t_price"] is not None else 0,
        "deposit": r["t_deposit"] if r["t_deposit"] is not None else 0,
        "qty": int(r["t_qty"] or 0),
        "img": tool_img(r["t_photo"], r["t_cat"]),
        "desc": clean_text(r["t_desc"]),
        "specs": clean_text(r["t_specs"]),
    }


# ---------------------------------------------------------------------------
# Источник «демо» (те же ключи, что у данных из Access)
# ---------------------------------------------------------------------------
def demo_categories():
    return [{"id": i, "name": c["name"], "img": c["img"]} for i, c in enumerate(CATEGORIES, 1)]


def demo_tools():
    cat_id = {c["slug"]: i for i, c in enumerate(CATEGORIES, 1)}
    brand_id = {b: i for i, b in enumerate(sorted({t["brand"] for t in TOOLS}), 1)}
    return [{"id": t["id"], "brand": t["brand"], "brand_id": brand_id[t["brand"]], "name": t["name"],
             "cat": cat_id[t["cat"]], "price": t["price"], "deposit": t["deposit"], "qty": t["qty"],
             "img": t["img"], "desc": t["desc"],
             "specs": "\n".join(f"{k}: {v}" for k, v in t["specs"])} for t in TOOLS]


# ---------------------------------------------------------------------------
# Функции данных, которые используют маршруты
# ---------------------------------------------------------------------------
def get_categories():
    if "cats" not in g:
        def from_access():
            rows = query("SELECT [id_category] AS c_id, [Название] AS c_name FROM [Категория] ORDER BY [Название]")
            return [{"id": r["c_id"], "name": clean_text(r["c_name"]),
                     "img": category_img(r["c_id"], r["c_name"])} for r in rows]
        g.cats = use_db(from_access, demo_categories)
    return g.cats


def get_category(cat_id):
    return next((c for c in get_categories() if c["id"] == cat_id), None)


def list_brands():
    def from_access():
        rows = query("SELECT DISTINCT b.[id_brand] AS b_id, b.[Название] AS b_name "
                     "FROM [Бренд] AS b INNER JOIN [Инструмент] AS t ON t.[id_brand] = b.[id_brand] "
                     "WHERE t.[Активен] = True ORDER BY b.[Название]")
        return [{"id": r["b_id"], "name": clean_text(r["b_name"])} for r in rows]

    def from_demo():
        seen = {t["brand_id"]: t["brand"] for t in demo_tools()}
        return [{"id": i, "name": n} for i, n in sorted(seen.items(), key=lambda x: x[1])]
    return use_db(from_access, from_demo)


def list_tools(cat=0, brand=0, avail=False, sort="", top=0):
    def from_access():
        where, params = ["t.[Активен] = True"], []
        if cat:
            where.append("t.[id_category] = ?")
            params.append(cat)
        if brand:
            where.append("t.[id_brand] = ?")
            params.append(brand)
        if avail:
            where.append("t.[КоличествоНаСкладе] > 0")
        top_sql = f"TOP {int(top)} " if top else ""   # TOP в Access параметром не задаётся
        sql = (f"SELECT {top_sql}{TOOL_SELECT} FROM {TOOL_FROM} "
               f"WHERE {' AND '.join(where)} ORDER BY {TOOL_SORT.get(sort, TOOL_SORT[''])}")
        return [tool_from_row(r) for r in query(sql, params)]

    def from_demo():
        res = [t for t in demo_tools()
               if (not cat or t["cat"] == cat) and (not brand or t["brand_id"] == brand)
               and (not avail or t["qty"] > 0)]
        keys = {"": lambda t: t["id"], "id_desc": lambda t: -t["id"], "price_asc": lambda t: t["price"],
                "price_desc": lambda t: -t["price"], "name": lambda t: t["name"]}
        res.sort(key=keys.get(sort, keys[""]))
        return res[:top] if top else res
    return use_db(from_access, from_demo)


def get_tool(tool_id):
    def from_access():
        rows = query(f"SELECT {TOOL_SELECT} FROM {TOOL_FROM} "
                     "WHERE t.[id_tool] = ? AND t.[Активен] = True", [tool_id])
        return tool_from_row(rows[0]) if rows else None

    def from_demo():
        return next((t for t in demo_tools() if t["id"] == tool_id), None)
    return use_db(from_access, from_demo)


def min_price():
    def from_access():
        rows = query("SELECT MIN(t.[ЦенаЗаСутки]) AS m FROM [Инструмент] AS t WHERE t.[Активен] = True")
        return rows[0]["m"] if rows else None
    return use_db(from_access, lambda: min(t["price"] for t in TOOLS))


def to_int(value):
    try:
        return max(int(value), 0)
    except (TypeError, ValueError):
        return 0


@app.template_filter("rub")
def rub(value):
    return f"{int(value or 0):,}".replace(",", "\u00a0") + "\u00a0₽"


# ---------------------------------------------------------------------------
# Общий контекст: режим B2C/B2B и боковая панель заявки (хранится в cookie-сессии)
# ---------------------------------------------------------------------------
def cart_items():
    get_categories()                                   # определяет источник данных (access/demo)
    if session.get("cart") and session.get("cart_src") != g.source:
        session["cart"] = {}                           # id из другого источника — не смешиваем
    items = []
    for tid, days in session.get("cart", {}).items():
        tool = get_tool(int(tid))
        if tool:
            items.append({"tool": tool, "days": days, "sum": tool["price"] * days})
    return items


def safe_next(default="index"):
    nxt = request.values.get("next", "")
    if not nxt.startswith("/") or nxt.startswith("//"):
        nxt = url_for(default)
    return nxt


@app.context_processor
def inject_globals():
    items = cart_items()
    return dict(
        mode=session.get("mode", "b2c"),
        categories=get_categories(),
        data_source=g.get("source", "demo"),
        cart_items=items,
        cart_count=len(items),
        cart_rent=sum(i["sum"] for i in items),
        cart_deposit=sum(i["tool"]["deposit"] for i in items),
        current=request.full_path.rstrip("?"),
    )


# ---------------------------------------------------------------------------
# Файлы: style.css лежит в корне, картинки — в img_tool
# ---------------------------------------------------------------------------
@app.route("/style.css")
def css():
    return send_from_directory(BASE_DIR, "style.css", mimetype="text/css")


@app.route("/img/<path:name>")
def img(name):
    return send_from_directory(IMG_DIR, name)


# ---------------------------------------------------------------------------
# Страницы каталога
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html", popular=list_tools(sort="id_desc", top=4),
                           min_price=min_price())


@app.route("/catalog")
def catalog():
    q = request.args.get("q", "").strip()
    cat = to_int(request.args.get("cat"))
    brand = to_int(request.args.get("brand"))
    avail = request.args.get("avail") == "1"
    sort = request.args.get("sort", "")

    result = list_tools(cat=cat, brand=brand, avail=avail, sort=sort)
    if q:   # поиск в Python: регистр кириллицы считается одинаково на любом драйвере
        ql = q.casefold()
        result = [t for t in result if ql in t["name"].casefold() or ql in t["brand"].casefold()]

    return render_template("catalog.html", tools=result, brands=list_brands(),
                           q=q, cat=cat, brand=brand, avail=avail, sort=sort,
                           active_cat=get_category(cat))


@app.route("/tool/<int:tool_id>")
def tool(tool_id):
    t = get_tool(tool_id)
    if not t:
        abort(404)
    similar = [x for x in list_tools(cat=t["cat"]) if x["id"] != t["id"]][:3]
    return render_template("tool.html", t=t, durations=DURATIONS,
                           category=get_category(t["cat"]), similar=similar)


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        flash("Вход пока не подключён: это визуальный макет. Авторизация появится на следующем этапе.")
        return redirect(url_for("login"))
    return render_template("login.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        flash("Регистрация пока не подключена: данные не сохраняются.")
        return redirect(url_for("register"))
    return render_template("register.html")


@app.route("/profile")
def profile():
    return render_template("profile.html")


@app.route("/booking", methods=["GET", "POST"])
def booking():
    if request.method == "POST":
        flash("Оформление пока не подключено: заявка нигде не сохраняется.")
        return redirect(url_for("booking"))
    return render_template("booking.html")


@app.route("/b2b")
def b2b():
    return render_template("b2b.html")


@app.route("/b2b/upload", methods=["POST"])
def b2b_upload():
    f = request.files.get("estimate")
    name = f.filename if f and f.filename else "файл не выбран"
    # Макет: файл НЕ читается и НЕ сохраняется.
    flash(f"Макет формы: «{name}» не анализируется и не сохраняется.")
    return redirect(url_for("b2b") + "#ai")


# ---------------------------------------------------------------------------
# Режим клиента и боковая панель заявки (без JavaScript: POST/GET + redirect)
# ---------------------------------------------------------------------------
@app.route("/mode/<m>")
def set_mode(m):
    if m in ("b2c", "b2b"):
        session["mode"] = m
    return redirect(safe_next())


@app.route("/cart/add", methods=["POST"])
def cart_add():
    try:
        tid = int(request.form.get("tool_id", 0))
        days = int(request.form.get("days", 1))
    except ValueError:
        abort(400)
    if not get_tool(tid) or days not in DURATIONS:
        abort(400)
    cart = session.get("cart", {})
    if session.get("cart_src") != g.source:   # заявка из другого источника данных не смешивается
        cart = {}
    cart[str(tid)] = days
    session["cart"] = cart
    session["cart_src"] = g.source
    return redirect(safe_next() + "#cart")  # якорь открывает панель, страница остаётся на месте


@app.route("/cart/remove/<int:tool_id>", methods=["POST"])
def cart_remove(tool_id):
    cart = session.get("cart", {})
    cart.pop(str(tool_id), None)
    session["cart"] = cart
    return redirect(safe_next() + "#cart")


@app.route("/cart/clear", methods=["POST"])
def cart_clear():
    session["cart"] = {}
    return redirect(safe_next() + "#cart")




@app.errorhandler(404)
def not_found(_):
    return render_template("catalog.html", tools=[], brands=[], q="", cat=0, brand=0,
                           avail=False, sort="", active_cat=None, not_found=True), 404


# ===== ВРЕМЕННО: /db-check — диагностика подключения. Удалить весь блок до строки «конец /db-check» =====
# Показывает только структуру (драйверы, названия колонок и типы, количество строк).
# Содержимое таблиц, в том числе пользовательских, НЕ выводится. Доступна только с localhost.
DB_EXPECTED = {
    "Пользователь": ["id_user", "Логин", "Пароль", "ФИО", "Email", "Телефон", "Адрес", "Роль"],
    "B2B_Реквизиты": ["id_b2b", "id_user", "НазваниеКомпании", "ИНН", "КПП", "ЮрАдрес"],
    "Инструмент": ["id_tool", "id_brand", "Название", "Описание", "Характеристики", "ЦенаЗаСутки", "Залог",
                   "КоличествоНаСкладе", "ИмяФото", "Активен", "id_category"],
    "Бренд": ["id_brand", "Название"],
    "Категория": ["id_category", "Название"],
    "Бронирование": ["id_rent", "id_user", "ДатаСоздания", "ДатаНачала", "ДатаОкончания", "Статус",
                     "СуммаАренды", "СуммаЗалога", "СпособДоставки"],
    "ПозицияБронирования": ["id_item", "id_rent", "id_tool", "Количество", "ЦенаНаМоментБрони"],
    "БлокировкаИнструмента": ["id_block", "id_tool", "ДатаНачала", "ДатаОкончания", "Причина"],
    "ИсторияСтатусов": ["id_log", "id_rent", "СтарыйСтатус", "НовыйСтатус", "КтоИзменил_id_user", "Когда",
                        "Комментарий"],
}


@app.route("/db-check")
def db_check():
    if request.remote_addr not in ("127.0.0.1", "::1"):
        abort(404)
    L = []
    add = L.append
    add("ToolRent: диагностика Access (временная страница, содержимое таблиц не показывается)")
    add(f"Python {sys.version.split()[0]}, {struct.calcsize('P') * 8}-bit, {platform.platform()}")
    add(f"Файл базы: {DB_PATH}")
    add(f"Файл найден: {'да, %d КБ' % (os.path.getsize(DB_PATH) // 1024) if os.path.isfile(DB_PATH) else 'НЕТ'}")
    if pyodbc is None:
        add("pyodbc: НЕ УСТАНОВЛЕН → выполните: pip install pyodbc")
        return Response("\n".join(L), mimetype="text/plain; charset=utf-8")
    add(f"pyodbc {pyodbc.version}")
    drivers = pyodbc.drivers()
    add("Драйверы Access: " + (", ".join(d for d in drivers if "Access" in d) or "НЕ НАЙДЕНЫ"))
    if ACCESS_DRIVER not in drivers:
        add(f"Нужен драйвер «{ACCESS_DRIVER}». Установите Microsoft Access Database Engine "
            "той же разрядности, что и Python (см. строку с разрядностью выше).")
    try:
        conn = pyodbc.connect(DB_CONN_STR)
    except Exception as e:
        add(f"\nПОДКЛЮЧЕНИЕ НЕ УДАЛОСЬ: {e}")
        return Response("\n".join(L), mimetype="text/plain; charset=utf-8")
    add("\nПодключение: OK (режим только чтение)\n")
    cur = conn.cursor()
    for table, expected in DB_EXPECTED.items():
        try:
            n = cur.execute(f"SELECT COUNT(*) FROM [{table}]").fetchone()[0]
            cols = {r.column_name: f"{r.type_name}({r.column_size})" for r in cur.columns(table=table)}
            add(f"[{table}] строк: {n}")
            add("   колонки: " + ", ".join(f"{k}:{v}" for k, v in cols.items()))
            miss = [c for c in expected if c not in cols]
            extra = [c for c in cols if c not in expected]
            if miss:
                add("   !!! НЕТ ожидаемых колонок: " + ", ".join(miss))
            if extra:
                add("   (есть сверх списка: " + ", ".join(extra) + ")")
        except Exception as e:
            add(f"[{table}] ОШИБКА: {e}")
    add("\nПроверка каталога:")
    try:
        rows = cur.execute("SELECT [id_tool], [ИмяФото] FROM [Инструмент] WHERE [Активен] = True").fetchall()
        add(f"   активных инструментов: {len(rows)}")
        nophoto = [r[0] for r in rows if not (r[1] or "").strip()]
        nofile = [(r[0], r[1]) for r in rows if (r[1] or "").strip()
                  and not os.path.isfile(os.path.join(IMG_DIR, os.path.basename(r[1].strip())))]
        add(f"   без ИмяФото (будет заглушка): {len(nophoto)}  id: {nophoto[:10]}")
        add(f"   файла нет в img_tool (будет заглушка): {len(nofile)}  {nofile[:10]}")
        no_cat = cur.execute("SELECT COUNT(*) FROM [Инструмент] AS t LEFT JOIN [Категория] AS c "
                             "ON t.[id_category] = c.[id_category] WHERE c.[id_category] IS NULL").fetchone()[0]
        no_brand = cur.execute("SELECT COUNT(*) FROM [Инструмент] AS t LEFT JOIN [Бренд] AS b "
                               "ON t.[id_brand] = b.[id_brand] WHERE b.[id_brand] IS NULL").fetchone()[0]
        add(f"   инструментов без категории: {no_cat}, без бренда: {no_brand}")
    except Exception as e:
        add(f"   ОШИБКА: {e}")
    conn.close()
    try:
        add(f"   запрос каталога (как на сайте): {len(query(f'SELECT {TOOL_SELECT} FROM {TOOL_FROM} WHERE t.[Активен] = True'))} строк")
    except Exception as e:
        add(f"   запрос каталога: ОШИБКА: {e}")
    return Response("\n".join(L), mimetype="text/plain; charset=utf-8")
# ===== конец /db-check =====


if __name__ == "__main__":
    app.run(debug=True)
