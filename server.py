import os
from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, send_from_directory, abort)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# static_folder=None: папки static нет, CSS и картинки отдаём явными маршрутами.
app = Flask(__name__, static_folder=None,
            template_folder=os.path.join(BASE_DIR, "templates"))
app.secret_key = "dev-only-change-before-deploy"

# ---------------------------------------------------------------------------
# ДЕМО-ДАННЫЕ (позже заменяются запросами к Access: Категория, Бренд, Инструмент)
# Поля повторяют будущую схему: id_tool, Бренд, Название, ЦенаЗаСутки, Залог,
# Количество, ИмяФото, id_category, Описание, Характеристики.
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


def get_tool(tool_id):
    return next((t for t in TOOLS if t["id"] == tool_id), None)


def get_category(slug):
    return next((c for c in CATEGORIES if c["slug"] == slug), None)


@app.template_filter("rub")
def rub(value):
    return f"{int(value):,}".replace(",", "\u00a0") + "\u00a0₽"


# ---------------------------------------------------------------------------
# Общий контекст: режим B2C/B2B и боковая панель заявки (хранится в cookie-сессии)
# ---------------------------------------------------------------------------
def cart_items():
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
        categories=CATEGORIES,
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
    return send_from_directory(os.path.join(BASE_DIR, "img_tool"), name)


# ---------------------------------------------------------------------------
# Страницы
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html", popular=TOOLS[:4])


@app.route("/catalog")
def catalog():
    q = request.args.get("q", "").strip()
    cat = request.args.get("cat", "")
    brand = request.args.get("brand", "")
    avail = request.args.get("avail") == "1"
    sort = request.args.get("sort", "")

    result = TOOLS
    if q:
        ql = q.lower()
        result = [t for t in result if ql in t["name"].lower() or ql in t["brand"].lower()]
    if cat:
        result = [t for t in result if t["cat"] == cat]
    if brand:
        result = [t for t in result if t["brand"] == brand]
    if avail:
        result = [t for t in result if t["qty"] > 0]
    if sort == "price_asc":
        result = sorted(result, key=lambda t: t["price"])
    elif sort == "price_desc":
        result = sorted(result, key=lambda t: -t["price"])
    elif sort == "name":
        result = sorted(result, key=lambda t: t["name"])

    brands = sorted({t["brand"] for t in TOOLS})
    return render_template("catalog.html", tools=result, brands=brands,
                           q=q, cat=cat, brand=brand, avail=avail, sort=sort,
                           active_cat=get_category(cat))


@app.route("/tool/<int:tool_id>")
def tool(tool_id):
    t = get_tool(tool_id)
    if not t:
        abort(404)
    similar = [x for x in TOOLS if x["cat"] == t["cat"] and x["id"] != t["id"]][:3]
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
    cart[str(tid)] = days
    session["cart"] = cart
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
    return render_template("catalog.html", tools=[], brands=[], q="", cat="", brand="",
                           avail=False, sort="", active_cat=None, not_found=True), 404


if __name__ == "__main__":
    app.run(debug=True)
