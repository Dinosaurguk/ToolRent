"""
ToolRent — сайт аренды строительного инструмента.
Python + Flask + Microsoft Access (pyodbc). JavaScript не используется.

Запуск:   python server.py                 →  http://127.0.0.1:5000
Админ:    python server.py make-admin ЛОГИН   (назначить единственного администратора)

Настройки берутся из переменных окружения (см. README):
  TOOLRENT_SECRET_KEY       ключ подписи сессий
  TOOLRENT_BASE_URL         адрес сайта для ссылок в письмах
  TOOLRENT_SMTP_HOST / _PORT / _USER / _PASSWORD / _FROM / _TLS   почта
  TOOLRENT_POINT_RUB_RATE   сколько рублей скидки даёт 1 балл (пусто — списание баллов выключено)
"""
import hashlib
import hmac
import os
import platform
import queue
import re
import secrets
import smtplib
import struct
import sys
import threading
import time
from datetime import date, datetime, timedelta
from decimal import ROUND_DOWN, ROUND_HALF_UP, Decimal, InvalidOperation
from email.message import EmailMessage
from functools import wraps

from flask import (Flask, Response, abort, flash, g, redirect, render_template,
                   request, send_from_directory, session, url_for)
from markupsafe import Markup, escape
from werkzeug.security import check_password_hash, generate_password_hash

try:
    import pyodbc
except ImportError:
    pyodbc = None

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
IMG_DIR = os.path.join(BASE_DIR, "img_tool")
DB_PATH = os.path.join(BASE_DIR, "database.accdb")

# ---------------------------------------------------------------------------
# Настройки (единое место для коэффициентов и параметров)
# ---------------------------------------------------------------------------
def _env(name, default=None):
    v = os.environ.get(name)
    return v if v not in (None, "") else default


POINTS_PERCENT = Decimal("5")            # начисление баллов: % от стоимости аренды (залог не входит)
SAVINGS_COEF = Decimal("2.325")          # коэффициент «Оценочной выгоды от аренды» (меняется только здесь)
try:
    POINT_RATE = Decimal(_env("TOOLRENT_POINT_RUB_RATE")) if _env("TOOLRENT_POINT_RUB_RATE") else None
except InvalidOperation:
    POINT_RATE = None                    # курс не задан или задан неверно — списание баллов выключено
if POINT_RATE is not None and POINT_RATE <= 0:
    POINT_RATE = None

DURATIONS = [1, 3, 7, 14, 30]
MAX_RENT_DAYS = 365
MAX_CART_LINES = 30
TOKEN_HOURS = {"email": 24, "reset": 1}
MAIL_RETRY_SECONDS = (60, 300, 1800)     # повторные попытки отправки письма

ROLE_USER, ROLE_B2B, ROLE_ADMIN = "user", "b2b", "admin"

# Статусы позиций заявки и заявки целиком (хранятся в базе как текст)
P_NEW, P_OK, P_NO, P_OUT, P_BACK, P_CANCEL = (
    "На рассмотрении", "Одобрена", "Отклонена", "Выдана", "Возвращена", "Отменена")
O_NEW, O_WORK, O_RENT, O_DONE, O_NO, O_CANCEL = (
    "На рассмотрении", "В обработке", "В аренде", "Завершена", "Отклонена", "Отменена")
RESERVING = (P_OK, P_OUT)                # статусы, при которых инструмент занят
DEAD = (P_NO, P_CANCEL)                  # позиции, которые не входят в стоимость

app = Flask(__name__, static_folder=None, template_folder=os.path.join(BASE_DIR, "templates"))
_secret = _env("TOOLRENT_SECRET_KEY")
if not _secret:
    _secret = secrets.token_hex(32)
    print("[ToolRent] TOOLRENT_SECRET_KEY не задан: создан временный ключ, "
          "после перезапуска сайта все пользователи выйдут из аккаунтов.")
app.secret_key = _secret
app.config.update(MAX_CONTENT_LENGTH=8 * 1024 * 1024, PERMANENT_SESSION_LIFETIME=timedelta(days=30),
                  SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax")

# ---------------------------------------------------------------------------
# Работа с Access (pyodbc). Все запросы параметризованы (?).
# ---------------------------------------------------------------------------
ACCESS_DRIVER = "Microsoft Access Driver (*.mdb, *.accdb)"
CONN_STR = f"DRIVER={{{ACCESS_DRIVER}}};DBQ={DB_PATH};"


class DBUnavailable(Exception):
    pass


def connect():
    if pyodbc is None:
        raise DBUnavailable("не установлен пакет pyodbc (pip install pyodbc)")
    if not os.path.isfile(DB_PATH):
        raise DBUnavailable(f"не найден файл базы {DB_PATH}")
    try:
        return pyodbc.connect(CONN_STR, autocommit=False)
    except pyodbc.Error as e:
        raise DBUnavailable(str(e))


# Названия колонок Access → ключи словарей в Python (в шаблонах удобнее латиница)
COL = {
    "Логин": "login", "Пароль": "pwd", "ФИО": "fio", "Телефон": "phone", "Email": "email", "Адрес": "address",
    "Роль": "role", "Баллы": "points", "EmailПодтвержден": "email_ok", "Активен": "active",
    "ПричинаБлокировки": "block_reason", "ДатаБлокировки": "block_date", "ДатаРегистрации": "reg_date",
    "НазваниеКомпании": "company", "ИНН": "inn", "КПП": "kpp", "ЮрАдрес": "legal_address",
    "ТипОрганизации": "org_type", "Название": "name", "Описание": "descr", "Характеристики": "specs",
    "ЦенаЗаСутки": "price", "Залог": "deposit", "КоличествоНаСкладе": "stock", "ИмяФото": "photo",
    "ДатаСоздания": "created", "Статус": "status", "СуммаАренды": "rent_sum", "СуммаЗалога": "deposit_sum",
    "СпособДоставки": "delivery", "СуммаДоСкидки": "sum_before", "СуммаСкидки": "discount",
    "Количество": "qty", "ЦенаНаМоментБрони": "price_at", "ДатаНачала": "start", "ДатаОкончания": "end",
    "СтарыйСтатус": "old_status", "НовыйСтатус": "new_status", "КтоИзменил_id_user": "by_user",
    "Когда": "at", "Комментарий": "comment", "Причина": "reason", "ТипОперации": "op_type", "Дата": "dt",
    "Код": "code", "ТипСкидки": "disc_type", "РазмерСкидки": "amount", "МинимальнаяСумма": "min_sum",
    "ЛимитИспользований": "limit", "ДатаВыдачи": "issued", "Использован": "used",
    "ДатаИспользования": "used_at", "Заголовок": "title", "Текст": "body", "Тип": "kind",
    "Прочитано": "is_read", "ТокенХеш": "token_hash", "СрокДействия": "expires", "ТипДействия": "action",
}


def _lat(row):
    return {COL.get(k, k): v for k, v in row.items()}


def _run(cur, sql, params):
    params = list(params)
    if pyodbc is not None and any(isinstance(p, str) and len(p) > 255 for p in params):
        # длинный текст (поля «Длинный текст») драйвер Access принимает только как LONGVARCHAR
        cur.setinputsizes([(pyodbc.SQL_WLONGVARCHAR, len(p), 0) if isinstance(p, str) and len(p) > 255 else None
                           for p in params])
    cur.execute(sql, params)


class Tx:
    """Одно соединение и одна транзакция: при ошибке всё откатывается."""

    def __enter__(self):
        self.conn = connect()
        self.cur = self.conn.cursor()
        return self

    def __exit__(self, et, ev, tb):
        try:
            if et is None:
                self.conn.commit()
            else:
                self.conn.rollback()
        finally:
            self.conn.close()
        return False

    def q(self, sql, params=()):
        _run(self.cur, sql, params)
        names = [c[0] for c in self.cur.description]
        return [_lat(dict(zip(names, r))) for r in self.cur.fetchall()]

    def one(self, sql, params=()):
        rows = self.q(sql, params)
        return rows[0] if rows else None

    def x(self, sql, params=()):
        _run(self.cur, sql, params)
        return self.cur.rowcount

    def insert(self, table, vals):
        cols = list(vals)
        sql = (f"INSERT INTO [{table}] ({', '.join('[' + c + ']' for c in cols)}) "
               f"VALUES ({', '.join(['?'] * len(cols))})")
        _run(self.cur, sql, [vals[c] for c in cols])
        self.cur.execute("SELECT @@IDENTITY")
        return int(self.cur.fetchone()[0])

    def update(self, table, key, key_val, vals):
        sets = ", ".join(f"[{c}] = ?" for c in vals)
        return self.x(f"UPDATE [{table}] SET {sets} WHERE [{key}] = ?", list(vals.values()) + [key_val])


def q(sql, params=()):
    with Tx() as t:
        return t.q(sql, params)


def q1(sql, params=()):
    rows = q(sql, params)
    return rows[0] if rows else None


# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------
def now():
    return datetime.now().replace(microsecond=0)


def today():
    return datetime.combine(date.today(), datetime.min.time())


def D(x):
    return Decimal(str(x)) if x is not None else Decimal(0)


def money(x):
    return D(x).quantize(Decimal("0.01"), ROUND_HALF_UP)


def parse_date(s):
    try:
        return datetime.strptime((s or "").strip(), "%Y-%m-%d")
    except ValueError:
        return None


def as_dt(v):
    """Дата из базы (datetime или строка) → datetime."""
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime.combine(v, datetime.min.time())
    if isinstance(v, str):
        for f in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(v[:19], f)
            except ValueError:
                pass
    return None


def rent_days(start, end):
    return (as_dt(end).date() - as_dt(start).date()).days + 1


def clean_text(v):
    if v is None:
        return ""
    return str(v).replace("\r\n", "\n").replace("\r", "\n").strip("\n")


@app.template_filter("rub")
def rub(value):
    v = money(value)
    s = f"{int(v):,}".replace(",", "\u00a0") if v == v.to_integral() else f"{v:,.2f}".replace(",", "\u00a0")
    return s + "\u00a0₽"


@app.template_filter("d")
def fmt_d(value):
    v = as_dt(value)
    return v.strftime("%d.%m.%Y") if v else "—"


@app.template_filter("dt")
def fmt_dt(value):
    v = as_dt(value)
    return v.strftime("%d.%m.%Y %H:%M") if v else "—"


@app.template_filter("iso")
def fmt_iso(value):
    v = as_dt(value)
    return v.strftime("%Y-%m-%d") if v else ""


def field(name, maxlen=255, required=False, label=None):
    """Текстовое поле формы: обрезает пробелы, проверяет длину."""
    v = (request.form.get(name) or "").strip()
    if required and not v:
        raise ValueError(f"Заполните поле «{label or name}».")
    if len(v) > maxlen:
        raise ValueError(f"Поле «{label or name}» слишком длинное (максимум {maxlen}).")
    return v


def to_int(value, default=0, lo=0, hi=10**9):
    try:
        return min(max(int(value), lo), hi)
    except (TypeError, ValueError):
        return default


def to_money(value, label):
    try:
        v = Decimal(str(value).replace(",", ".").replace(" ", "").replace("\u00a0", ""))
    except InvalidOperation:
        raise ValueError(f"Поле «{label}» должно быть числом.")
    if v < 0 or v > Decimal("100000000"):
        raise ValueError(f"Поле «{label}» вне допустимых пределов.")
    return money(v)


def safe_next(default="index"):
    nxt = request.values.get("next", "")
    if not nxt.startswith("/") or nxt.startswith("//") or "\\" in nxt:
        nxt = url_for(default)
    return nxt


# ---------------------------------------------------------------------------
# Безопасность: CSRF, токены, пароли
# ---------------------------------------------------------------------------
def csrf_token():
    if "_csrf" not in session:
        session["_csrf"] = secrets.token_hex(16)
    return session["_csrf"]


app.jinja_env.globals["csrf_field"] = lambda: Markup(
    f'<input type="hidden" name="csrf" value="{escape(csrf_token())}">')


def hash_password(pwd):
    return generate_password_hash(pwd, method="pbkdf2:sha256")     # ~100 символов, помещается в поле на 255


def token_digest(raw):
    return hashlib.sha256(raw.encode()).hexdigest()


def make_token(t, uid, kind):
    """Одноразовая ссылка: в базе хранится только хеш токена. Прежние токены этого типа гасятся."""
    t.x("UPDATE [ТокеныАккаунта] SET [Использован] = ? WHERE [id_user] = ? AND [Тип] = ? AND [Использован] = ?",
        [True, uid, kind, False])
    raw = secrets.token_urlsafe(32)
    t.insert("ТокеныАккаунта", {"id_user": uid, "ТокенХеш": token_digest(raw), "Тип": kind,
                                "ДатаСоздания": now(), "СрокДействия": now() + timedelta(hours=TOKEN_HOURS[kind]),
                                "Использован": False})
    return raw


def find_token(t, raw, kind):
    row = t.one("SELECT * FROM [ТокеныАккаунта] WHERE [ТокенХеш] = ? AND [Тип] = ? AND [Использован] = ?",
                [token_digest(raw), kind, False])
    if row and as_dt(row["expires"]) and as_dt(row["expires"]) >= now():
        return row
    return None


_fails = {}


def too_many_failures(key):
    stamps = [s for s in _fails.get(key, []) if s > time.time() - 600]
    _fails[key] = stamps
    return len(stamps) >= 5


def note_failure(key):
    _fails.setdefault(key, []).append(time.time())


# ---------------------------------------------------------------------------
# Почта и уведомления (письма уходят в фоне, с повторными попытками)
# ---------------------------------------------------------------------------
_mail_q = queue.Queue()
_mail_thread = None


def base_url():
    return (_env("TOOLRENT_BASE_URL") or (request.url_root if request else "http://127.0.0.1:5000/")).rstrip("/")


def _smtp_send(job):
    host = _env("TOOLRENT_SMTP_HOST")
    port = int(_env("TOOLRENT_SMTP_PORT", "587"))
    mode = (_env("TOOLRENT_SMTP_TLS", "starttls") or "").lower()
    msg = EmailMessage()
    msg["From"] = _env("TOOLRENT_SMTP_FROM") or _env("TOOLRENT_SMTP_USER") or "toolrent@localhost"
    msg["To"] = job["to"]
    msg["Subject"] = job["subject"]
    msg.set_content(job["body"])
    smtp = smtplib.SMTP_SSL(host, port, timeout=20) if mode == "ssl" else smtplib.SMTP(host, port, timeout=20)
    try:
        if mode == "starttls":
            smtp.starttls()
        if _env("TOOLRENT_SMTP_USER"):
            smtp.login(_env("TOOLRENT_SMTP_USER"), _env("TOOLRENT_SMTP_PASSWORD", ""))
        smtp.send_message(msg)
    finally:
        try:
            smtp.quit()
        except Exception:
            pass


def _mail_worker():
    while True:
        job = _mail_q.get()
        try:
            _smtp_send(job)
        except Exception as e:
            job["tries"] += 1
            if job["tries"] <= len(MAIL_RETRY_SECONDS):
                delay = MAIL_RETRY_SECONDS[job["tries"] - 1]
                print(f"[ToolRent] письмо для {job['to']} не отправлено ({e}); повтор через {delay} с")
                threading.Timer(delay, _mail_q.put, [job]).start()
            else:
                print(f"[ToolRent] письмо для {job['to']} не отправлено, попытки исчерпаны: {e}")


def send_email(to, subject, body):
    """Ставит письмо в очередь и сразу возвращает управление: страница не ждёт SMTP."""
    global _mail_thread
    if not _env("TOOLRENT_SMTP_HOST"):
        print(f"[ToolRent] SMTP не настроен, письмо не отправлено.\n  Кому: {to}\n  Тема: {subject}\n  {body}\n")
        return
    if _mail_thread is None or not _mail_thread.is_alive():
        _mail_thread = threading.Thread(target=_mail_worker, daemon=True)
        _mail_thread.start()
    _mail_q.put({"to": to, "subject": subject, "body": body, "tries": 0})


def notify(uid, title, body, kind="Система", rent_id=None, mail=True):
    """Внутреннее уведомление + письмо (только на подтверждённую почту)."""
    with Tx() as t:
        t.insert("Уведомления", {"id_user": uid, "Заголовок": title[:255], "Текст": body, "Тип": kind,
                                 "ДатаСоздания": now(), "Прочитано": False, "id_rent": rent_id})
        u = t.one("SELECT * FROM [Пользователь] WHERE [id_user] = ?", [uid])
    if mail and u and u["email"] and u["email_ok"]:
        send_email(u["email"], f"ToolRent: {title}", f"{body}\n\nЛичный кабинет: {base_url()}/profile/messages")


def notify_all(notes):
    for n in notes:
        notify(*n)


def admin_ids(t=None):
    rows = (t.q if t else q)("SELECT [id_user] FROM [Пользователь] WHERE [Роль] = ?", [ROLE_ADMIN])
    return [r["id_user"] for r in rows]


def admin_log(t, admin_id, action, target_user=None, comment=""):
    t.insert("ИсторияДействийАдминистратора", {"id_user_admin": admin_id, "id_user_target": target_user,
                                               "ТипДействия": action[:255], "Дата": now(), "Комментарий": comment})


# ---------------------------------------------------------------------------
# Каталог
# ---------------------------------------------------------------------------
CATEGORY_IMG_BY_ID = {}      # при желании: {id_category: "файл.svg"}
CATEGORY_IMG_WORDS = [        # иначе — по началу слов в названии категории
    (("перфор",), "perforator.svg"), (("дрел", "шуруп"), "drill.svg"),
    (("шлиф", "резк", "болгар", "пил"), "grinder.svg"), (("генер", "электростан"), "generator.svg"),
    (("нивел", "лазер", "измер"), "laser.svg"), (("леса", "подмост"), "scaffold.svg"),
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


def get_categories():
    if "cats" not in g:
        g.cats = [{"id": r["id_category"], "name": clean_text(r["name"]),
                   "img": category_img(r["id_category"], r["name"])}
                  for r in q("SELECT * FROM [Категория] ORDER BY [Название]")]
    return g.cats


def get_category(cat_id):
    return next((c for c in get_categories() if c["id"] == cat_id), None)


def get_brands():
    if "brands" not in g:
        g.brands = {r["id_brand"]: clean_text(r["name"]) for r in q("SELECT * FROM [Бренд] ORDER BY [Название]")}
    return g.brands


def tool_img(photo, cat_id):
    """Фото из img_tool по ИмяФото; нет файла — иллюстрация категории."""
    name = os.path.basename(str(photo or "").strip())
    if name and os.path.isfile(os.path.join(IMG_DIR, name)):
        return name
    cat = get_category(cat_id)
    return cat["img"] if cat else DEFAULT_IMG


def reserved_map(start, end, t=None, exclude_rent=None):
    """{id_tool: занято штук} для одобренных и выданных позиций, пересекающихся с периодом."""
    rows = (t.q if t else q)(
        "SELECT [id_tool], [id_rent], [Количество], [ДатаНачала], [ДатаОкончания] FROM [ПозицияБронирования] "
        "WHERE [Статус] IN (?, ?) AND [ДатаНачала] <= ? AND [ДатаОкончания] >= ?", [P_OK, P_OUT, end, start])
    out = {}
    for r in rows:
        if exclude_rent and r["id_rent"] == exclude_rent:
            continue
        out[r["id_tool"]] = out.get(r["id_tool"], 0) + int(r["qty"] or 0)
    return out


def blocked_tools(start, end, t=None):
    rows = (t.q if t else q)("SELECT [id_tool], [ДатаНачала], [ДатаОкончания] FROM [БлокировкаИнструмента]")
    out = set()
    for r in rows:
        s, e = as_dt(r["start"]), as_dt(r["end"])
        if (s is None or s <= end) and (e is None or e >= start):
            out.add(r["id_tool"])
    return out


def tool_free(t, tool_id, stock, start, end, exclude_rent=None):
    if tool_id in blocked_tools(start, end, t):
        return 0
    return max(int(stock) - reserved_map(start, end, t, exclude_rent).get(tool_id, 0), 0)


def build_tool(r, reserved=None, blocked=None):
    stock = int(r["stock"] or 0)
    free = stock if reserved is None else max(stock - reserved.get(r["id_tool"], 0), 0)
    if blocked and r["id_tool"] in blocked:
        free = 0
    return {"id": r["id_tool"], "brand": get_brands().get(r["id_brand"], ""), "brand_id": r["id_brand"],
            "name": clean_text(r["name"]), "cat": r["id_category"], "price": money(r["price"]),
            "deposit": money(r["deposit"]), "stock": stock, "qty": free,
            "img": tool_img(r["photo"], r["id_category"]), "photo": clean_text(r["photo"]),
            "desc": clean_text(r["descr"]), "specs": clean_text(r["specs"]), "active": bool(r["active"])}


TOOL_SORT = {"": lambda t: t["id"], "id_desc": lambda t: -t["id"], "price_asc": lambda t: (t["price"], t["id"]),
             "price_desc": lambda t: (-t["price"], t["id"]), "name": lambda t: (t["name"].casefold(), t["id"])}


def list_tools(cat=0, brand=0, avail=False, sort="", top=0, include_inactive=False):
    where, params = ([] if include_inactive else ["[Активен] = ?"]), ([] if include_inactive else [True])
    if cat:
        where.append("[id_category] = ?")
        params.append(cat)
    if brand:
        where.append("[id_brand] = ?")
        params.append(brand)
    rows = q("SELECT * FROM [Инструмент]" + (" WHERE " + " AND ".join(where) if where else ""), params)
    day = today()
    reserved, blocked = reserved_map(day, day), blocked_tools(day, day)
    tools = [build_tool(r, reserved, blocked) for r in rows]
    if avail:
        tools = [t for t in tools if t["qty"] > 0]
    tools.sort(key=TOOL_SORT.get(sort, TOOL_SORT[""]))
    return tools[:top] if top else tools


def tools_by_ids(ids, include_inactive=False):
    ids = list(ids)
    if not ids:
        return {}
    rows = q(f"SELECT * FROM [Инструмент] WHERE [id_tool] IN ({', '.join(['?'] * len(ids))})", ids)
    return {r["id_tool"]: build_tool(r) for r in rows if include_inactive or r["active"]}


def get_tool(tool_id, include_inactive=False):
    r = q1("SELECT * FROM [Инструмент] WHERE [id_tool] = ?", [tool_id])
    if not r or not (include_inactive or r["active"]):
        return None
    day = today()
    return build_tool(r, reserved_map(day, day), blocked_tools(day, day))


# ---------------------------------------------------------------------------
# Заявка (корзина): позиции с собственным количеством и сроками, хранится в сессии
# ---------------------------------------------------------------------------
def cart_lines():
    lines = session.get("cart", [])
    tools = tools_by_ids({l["t"] for l in lines})
    out = []
    for l in lines:
        tool = tools.get(l["t"])
        s, e = parse_date(l["s"]), parse_date(l["e"])
        if not tool or not s or not e:
            continue
        days = rent_days(s, e)
        out.append({"k": l["k"], "tool": tool, "qty": l["q"], "start": s, "end": e, "days": days,
                    "rent": tool["price"] * l["q"] * days, "deposit": tool["deposit"] * l["q"]})
    return out


def validate_period(start, end):
    if not start or not end:
        raise ValueError("Укажите даты начала и окончания аренды.")
    if start < today():
        raise ValueError("Дата начала не может быть в прошлом.")
    if end < start:
        raise ValueError("Дата окончания раньше даты начала.")
    if rent_days(start, end) > MAX_RENT_DAYS:
        raise ValueError(f"Максимальный срок аренды — {MAX_RENT_DAYS} дней.")


def cart_add_line(tool_id, qty, start, end):
    validate_period(start, end)
    if not 1 <= qty <= 9999:
        raise ValueError("Количество должно быть от 1 до 9999.")
    tool = get_tool(tool_id)
    if not tool:
        raise ValueError("Инструмент не найден или снят с аренды.")
    cart = session.get("cart", [])
    s, e = start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")
    for l in cart:
        if l["t"] == tool_id and l["s"] == s and l["e"] == e:
            l["q"] = min(l["q"] + qty, 9999)
            break
    else:
        if len(cart) >= MAX_CART_LINES:
            raise ValueError(f"В заявке не может быть больше {MAX_CART_LINES} позиций.")
        cart.append({"k": secrets.token_hex(4), "t": tool_id, "q": qty, "s": s, "e": e})
    session["cart"] = cart


# ---------------------------------------------------------------------------
# Купоны, баллы, оценочная выгода
# ---------------------------------------------------------------------------
def coupon_discount(c, subtotal):
    v = D(subtotal) * D(c["amount"]) / 100 if c["disc_type"] == "Процент" else D(c["amount"])
    return money(min(v, D(subtotal)))


def usable_coupons(t, uid, subtotal=None):
    """Купоны пользователя с пометкой, можно ли применить их сейчас."""
    out = []
    for uc in t.q("SELECT * FROM [КупонПользователя] WHERE [id_user] = ? AND [Использован] = ?", [uid, False]):
        c = t.one("SELECT * FROM [Купон] WHERE [id_coupon] = ?", [uc["id_coupon"]])
        if not c:
            continue
        why = ""
        n = now()
        end = as_dt(uc["end"]) or as_dt(c["end"])
        if not c["active"]:
            why = "купон отключён"
        elif as_dt(c["start"]) and as_dt(c["start"]) > n:
            why = f"действует с {fmt_d(c['start'])}"
        elif end and end + timedelta(days=1) <= n:
            why = "срок действия истёк"
        elif c["limit"] and t.one("SELECT COUNT(*) AS n FROM [ИспользованиеКупона] WHERE [id_coupon] = ?",
                                  [c["id_coupon"]])["n"] >= c["limit"]:
            why = "лимит использований исчерпан"
        elif subtotal is not None and D(subtotal) < D(c["min_sum"]):
            why = f"минимальная сумма заказа {rub(c['min_sum'])}"
        out.append({"uc": uc["id_user_coupon"], "coupon": c, "end": end, "ok": not why, "why": why,
                    "discount": coupon_discount(c, subtotal) if subtotal is not None and not why else None})
    return out


def estimated_benefit(rent_total):
    """«Оценочная выгода от аренды»: условная величина, не подтверждённая рыночная экономия.
    Формула и коэффициент SAVINGS_COEF меняются только здесь."""
    return money(D(rent_total) * (SAVINGS_COEF - 1))


def user_stats(uid):
    orders = q("SELECT * FROM [Бронирование] WHERE [id_user] = ?", [uid])
    done = [o for o in orders if o["status"] == O_DONE]
    rent_total = sum((D(o["rent_sum"]) for o in done), Decimal(0))
    return {"orders": len(orders), "done": len(done), "rent_total": money(rent_total),
            "benefit": estimated_benefit(rent_total)}


# ---------------------------------------------------------------------------
# Заявки: создание, статусы, пересчёт, баллы
# ---------------------------------------------------------------------------
def derive_status(statuses):
    s = set(statuses)
    if not s:
        return O_NEW
    if s <= {P_CANCEL}:
        return O_CANCEL
    if s <= set(DEAD):
        return O_NO
    if s <= {P_BACK, P_NO, P_CANCEL}:
        return O_DONE
    if P_OUT in s:
        return O_RENT
    if s <= {P_NEW}:
        return O_NEW
    return O_WORK


def create_order(user, lines, uc_id, points, delivery):
    if not lines:
        raise ValueError("Заявка пуста.")
    if delivery not in ("Самовывоз", "Доставка"):
        raise ValueError("Выберите способ получения.")
    notes = []
    with Tx() as t:
        taken = []        # позиции этой же заявки, уже занявшие инструмент
        for l in lines:
            tool = t.one("SELECT * FROM [Инструмент] WHERE [id_tool] = ? AND [Активен] = ?", [l["tool"]["id"], True])
            if not tool:
                raise ValueError(f"«{l['tool']['name']}» больше недоступен для аренды.")
            validate_period(l["start"], l["end"])
            free = tool_free(t, tool["id_tool"], tool["stock"], l["start"], l["end"])
            free -= sum(q_ for (tid, s, e, q_) in taken if tid == tool["id_tool"] and s <= l["end"] and e >= l["start"])
            if l["qty"] > free:
                raise ValueError(f"«{l['tool']['name']}»: на {fmt_d(l['start'])}–{fmt_d(l['end'])} "
                                 f"свободно только {max(free, 0)} шт.")
            taken.append((tool["id_tool"], l["start"], l["end"], l["qty"]))
        subtotal = sum((l["rent"] for l in lines), Decimal(0))
        deposit = sum((l["deposit"] for l in lines), Decimal(0))
        disc_c, coupon = Decimal(0), None
        if uc_id:
            cand = [c for c in usable_coupons(t, user["id_user"], subtotal) if c["uc"] == uc_id]
            if not cand or not cand[0]["ok"]:
                raise ValueError("Выбранный купон нельзя применить: " + (cand[0]["why"] if cand else "он не найден") + ".")
            coupon = cand[0]
            disc_c = coupon["discount"]
        pts_used, disc_p = 0, Decimal(0)
        if points and POINT_RATE:
            disc_p = min(D(points) * POINT_RATE, subtotal - disc_c)
            pts_used = int((disc_p / POINT_RATE).to_integral_value(ROUND_DOWN))
            disc_p = money(D(pts_used) * POINT_RATE)
        discount = money(disc_c + disc_p)
        rid = t.insert("Бронирование", {
            "id_user": user["id_user"], "ДатаСоздания": now(), "Статус": O_NEW,
            "СуммаДоСкидки": money(subtotal), "СуммаСкидки": discount, "СуммаАренды": money(subtotal - discount),
            "СуммаЗалога": money(deposit), "СпособДоставки": delivery})
        for l in lines:
            item = t.insert("ПозицияБронирования", {
                "id_rent": rid, "id_tool": l["tool"]["id"], "Количество": l["qty"], "ЦенаНаМоментБрони": l["tool"]["price"],
                "ДатаНачала": l["start"], "ДатаОкончания": l["end"], "Статус": P_NEW})
            t.insert("ИсторияСтатусовПозиций", {"id_item": item, "СтарыйСтатус": "—", "НовыйСтатус": P_NEW,
                                               "КтоИзменил_id_user": user["id_user"], "Когда": now(),
                                               "Комментарий": "Позиция создана"})
        t.insert("ИсторияСтатусов", {"id_rent": rid, "СтарыйСтатус": "—", "НовыйСтатус": O_NEW,
                                     "КтоИзменил_id_user": user["id_user"], "Когда": now(),
                                     "Комментарий": "Заявка создана пользователем"})
        if coupon:
            t.x("UPDATE [КупонПользователя] SET [Использован] = ? WHERE [id_user_coupon] = ?", [True, uc_id])
            t.insert("ИспользованиеКупона", {"id_coupon": coupon["coupon"]["id_coupon"], "id_user": user["id_user"],
                                             "id_rent": rid, "РазмерСкидки": disc_c, "ДатаИспользования": now()})
        if pts_used:
            if t.x("UPDATE [Пользователь] SET [Баллы] = [Баллы] - ? WHERE [id_user] = ? AND [Баллы] >= ?",
                   [pts_used, user["id_user"], pts_used]) != 1:
                raise ValueError("Недостаточно баллов.")
            t.insert("ОперацииБаллов", {"id_user": user["id_user"], "id_rent": rid, "Количество": -pts_used,
                                        "ТипОперации": "Списание", "Дата": now(),
                                        "Комментарий": f"Оплата части заявки №{rid} баллами"})
        notes.append((user["id_user"], f"Заявка №{rid} отправлена",
                      f"Мы получили вашу заявку №{rid} ({len(lines)} поз.). Администратор рассмотрит её, "
                      f"статус можно смотреть в личном кабинете.", "Заявка", rid))
        for aid in admin_ids(t):
            notes.append((aid, f"Новая заявка №{rid}", f"Пользователь {user['login']} отправил заявку №{rid}.",
                          "Заявка", rid, False))
    notify_all(notes)
    return rid


def refresh_order(t, rid, who, comment=""):
    """Пересчитывает суммы и общий статус заявки; при завершении начисляет баллы, при полном отказе возвращает скидки."""
    notes = []
    order = t.one("SELECT * FROM [Бронирование] WHERE [id_rent] = ?", [rid])
    pos = t.q("SELECT * FROM [ПозицияБронирования] WHERE [id_rent] = ?", [rid])
    live = [p for p in pos if p["status"] not in DEAD]
    sub = sum((D(p["price_at"]) * p["qty"] * rent_days(p["start"], p["end"]) for p in live), Decimal(0))
    dep = Decimal(0)
    for p in live:
        tool = t.one("SELECT * FROM [Инструмент] WHERE [id_tool] = ?", [p["id_tool"]])
        dep += D(tool["deposit"]) * p["qty"] if tool else 0
    new = derive_status([p["status"] for p in pos])
    disc_c = disc_p = Decimal(0)
    usage = t.one("SELECT * FROM [ИспользованиеКупона] WHERE [id_rent] = ?", [rid])
    if live and usage:
        c = t.one("SELECT * FROM [Купон] WHERE [id_coupon] = ?", [usage["id_coupon"]])
        if c:
            disc_c = coupon_discount(c, sub)
            if D(usage["amount"]) != disc_c:
                t.update("ИспользованиеКупона", "id_usage", usage["id_usage"], {"РазмерСкидки": disc_c})
    if live and POINT_RATE:
        spent = -sum(int(o["delta"]) for o in t.q(
            "SELECT [Количество] AS delta FROM [ОперацииБаллов] WHERE [id_rent] = ? AND [ТипОперации] = ?",
            [rid, "Списание"]))
        disc_p = min(D(spent) * POINT_RATE, sub - disc_c)
    discount = money(min(disc_c + disc_p, sub))
    t.update("Бронирование", "id_rent", rid, {"СуммаДоСкидки": money(sub), "СуммаСкидки": discount,
                                              "СуммаАренды": money(sub - discount), "СуммаЗалога": money(dep),
                                              "Статус": new})
    old = order["status"]
    if new != old:
        t.insert("ИсторияСтатусов", {"id_rent": rid, "СтарыйСтатус": old, "НовыйСтатус": new,
                                     "КтоИзменил_id_user": who, "Когда": now(), "Комментарий": comment or None})
    uid = order["id_user"]
    if new in (O_NO, O_CANCEL):
        notes += release_discounts(t, rid, uid)
    if new == O_DONE:
        order = t.one("SELECT * FROM [Бронирование] WHERE [id_rent] = ?", [rid])
        notes += accrue_points(t, order)
        notes.append((uid, f"Заявка №{rid} завершена", f"Все позиции заявки №{rid} возвращены. Спасибо, что выбрали ToolRent!",
                      "Заявка", rid))
    return old, new, notes


def release_discounts(t, rid, uid):
    """Полный отказ или отмена: купон снова доступен, списанные баллы возвращаются."""
    notes = []
    usage = t.one("SELECT * FROM [ИспользованиеКупона] WHERE [id_rent] = ?", [rid])
    if usage:
        uc = t.one("SELECT * FROM [КупонПользователя] WHERE [id_user] = ? AND [id_coupon] = ? AND [Использован] = ?",
                   [uid, usage["id_coupon"], True])
        if uc:
            t.x("UPDATE [КупонПользователя] SET [Использован] = ? WHERE [id_user_coupon] = ?", [False, uc["id_user_coupon"]])
        t.x("DELETE FROM [ИспользованиеКупона] WHERE [id_usage] = ?", [usage["id_usage"]])
        notes.append((uid, "Купон возвращён", f"Заявка №{rid} не состоялась, купон снова доступен.", "Купон", rid))
    spent = -sum(int(o["delta"]) for o in t.q(
        "SELECT [Количество] AS delta FROM [ОперацииБаллов] WHERE [id_rent] = ? AND [ТипОперации] = ?", [rid, "Списание"]))
    back = t.one("SELECT [id_operation] FROM [ОперацииБаллов] WHERE [id_rent] = ? AND [ТипОперации] = ?",
                 [rid, "Возврат списанных баллов"])
    if spent > 0 and not back:
        t.x("UPDATE [Пользователь] SET [Баллы] = [Баллы] + ? WHERE [id_user] = ?", [spent, uid])
        t.insert("ОперацииБаллов", {"id_user": uid, "id_rent": rid, "Количество": spent,
                                    "ТипОперации": "Возврат списанных баллов", "Дата": now(),
                                    "Комментарий": f"Заявка №{rid} не состоялась"})
        notes.append((uid, "Баллы возвращены", f"Заявка №{rid} не состоялась, возвращено баллов: {spent}.", "Баллы", rid))
    return notes


def accrue_points(t, order):
    rid, uid = order["id_rent"], order["id_user"]
    if t.one("SELECT [id_operation] FROM [ОперацииБаллов] WHERE [id_rent] = ? AND [ТипОперации] = ?",
             [rid, "Начисление за аренду"]):
        return []                                   # повторно за одну заявку не начисляем
    pts = int((D(order["rent_sum"]) * POINTS_PERCENT / 100).to_integral_value(ROUND_DOWN))
    if pts <= 0:
        return []
    t.x("UPDATE [Пользователь] SET [Баллы] = [Баллы] + ? WHERE [id_user] = ?", [pts, uid])
    t.insert("ОперацииБаллов", {"id_user": uid, "id_rent": rid, "Количество": pts, "ТипОперации": "Начисление за аренду",
                                "Дата": now(), "Комментарий": f"{POINTS_PERCENT}% от стоимости аренды по заявке №{rid}"})
    return [(uid, "Начислены баллы", f"За заявку №{rid} начислено баллов: {pts}.", "Баллы", rid)]


ITEM_ACTIONS = {"approve": (P_NEW, P_OK, "одобрена"), "reject": (P_NEW, P_NO, "отклонена"),
                "issue": (P_OK, P_OUT, "выдана"), "return": (P_OUT, P_BACK, "возвращена")}


def admin_item_action(admin, item_id, action, reason=""):
    if action not in ITEM_ACTIONS:
        raise ValueError("Неизвестное действие.")
    need, new, word = ITEM_ACTIONS[action]
    if action == "reject" and not reason.strip():
        raise ValueError("При отказе нужно указать причину.")
    notes = []
    with Tx() as t:
        it = t.one("SELECT * FROM [ПозицияБронирования] WHERE [id_item] = ?", [item_id])
        if not it:
            raise ValueError("Позиция не найдена.")
        if it["status"] != need:
            raise ValueError(f"Действие невозможно: позиция сейчас «{it['status']}».")
        order = t.one("SELECT * FROM [Бронирование] WHERE [id_rent] = ?", [it["id_rent"]])
        tool = t.one("SELECT * FROM [Инструмент] WHERE [id_tool] = ?", [it["id_tool"]])
        if action == "approve" and tool_free(t, it["id_tool"], tool["stock"], as_dt(it["start"]), as_dt(it["end"])) < it["qty"]:
            raise ValueError("Недостаточно свободного инструмента на эти даты.")
        t.update("ПозицияБронирования", "id_item", item_id, {"Статус": new})
        t.insert("ИсторияСтатусовПозиций", {"id_item": item_id, "СтарыйСтатус": it["status"], "НовыйСтатус": new,
                                           "КтоИзменил_id_user": admin["id_user"], "Когда": now(),
                                           "Комментарий": reason.strip() or None})
        old, st, more = refresh_order(t, order["id_rent"], admin["id_user"], f"После изменения позиции №{item_id}")
        what = f"«{clean_text(tool['name'])}» × {it['qty']}, {fmt_d(it['start'])}–{fmt_d(it['end'])}"
        titles = {"approve": "Позиция одобрена", "reject": "Позиция отклонена", "issue": "Оборудование выдано",
                  "return": "Возврат подтверждён"}
        body = f"Заявка №{order['id_rent']}: {what} — {word}."
        if action == "reject":
            body += f"\nПричина: {reason.strip()}"
        if st != old:
            body += f"\nСтатус заявки: {st}."
        notes.append((order["id_user"], titles[action], body, "Заявка", order["id_rent"]))
        notes += more
        admin_log(t, admin["id_user"], f"Позиция {word}", order["id_user"],
                  f"Заявка №{order['id_rent']}, позиция №{item_id}: {what}" + (f". Причина: {reason.strip()}" if reason.strip() else ""))
    notify_all(notes)


def user_cancel_order(user, rid):
    notes = []
    with Tx() as t:
        order = t.one("SELECT * FROM [Бронирование] WHERE [id_rent] = ? AND [id_user] = ?", [rid, user["id_user"]])
        if not order:
            raise ValueError("Заявка не найдена.")
        pos = t.q("SELECT * FROM [ПозицияБронирования] WHERE [id_rent] = ?", [rid])
        if any(p["status"] in (P_OUT, P_BACK) for p in pos):
            raise ValueError("Заявку с выданным оборудованием отменить нельзя: обратитесь к администратору.")
        for p in pos:
            if p["status"] in (P_NEW, P_OK):
                t.update("ПозицияБронирования", "id_item", p["id_item"], {"Статус": P_CANCEL})
                t.insert("ИсторияСтатусовПозиций", {"id_item": p["id_item"], "СтарыйСтатус": p["status"],
                                                   "НовыйСтатус": P_CANCEL, "КтоИзменил_id_user": user["id_user"],
                                                   "Когда": now(), "Комментарий": "Отменено пользователем"})
        old, new, more = refresh_order(t, rid, user["id_user"], "Отменено пользователем")
        notes.append((user["id_user"], f"Заявка №{rid} отменена", "Вы отменили заявку.", "Заявка", rid))
        notes += more
    notify_all(notes)


# ---------------------------------------------------------------------------
# Напоминания о возврате и просрочки (фоновая проверка раз в 30 минут)
# ---------------------------------------------------------------------------
def due_check():
    day = today()
    n = 0
    for p in q("SELECT * FROM [ПозицияБронирования] WHERE [Статус] = ?", [P_OUT]):
        end = as_dt(p["end"])
        if not end:
            continue
        if end < day:
            title, kind = f"Просрочка возврата: позиция №{p['id_item']}", "Просрочка"
        elif end <= day + timedelta(days=1):
            title, kind = f"Скоро возврат: позиция №{p['id_item']}", "Напоминание"
        else:
            continue
        order = q1("SELECT * FROM [Бронирование] WHERE [id_rent] = ?", [p["id_rent"]])
        if not order or q1("SELECT [id_notification] FROM [Уведомления] WHERE [id_user] = ? AND [Заголовок] = ?",
                           [order["id_user"], title]):
            continue
        tool = q1("SELECT * FROM [Инструмент] WHERE [id_tool] = ?", [p["id_tool"]])
        text = (f"Заявка №{p['id_rent']}: «{clean_text(tool['name']) if tool else 'инструмент'}» × {p['qty']}. "
                f"Срок возврата — {fmt_d(end)}.")
        notify(order["id_user"], title, text + (" Срок аренды истёк, пожалуйста, верните оборудование." if kind == "Просрочка" else ""),
               kind, p["id_rent"])
        for aid in admin_ids():
            notify(aid, title, text, kind, p["id_rent"], mail=False)
        n += 1
    return n


def _due_loop():
    while True:
        try:
            due_check()
        except Exception as e:
            print("[ToolRent] проверка сроков возврата не удалась:", e)
        time.sleep(1800)


# ---------------------------------------------------------------------------
# Карта активности (по таблицам заявок, позиций и историй статусов)
# ---------------------------------------------------------------------------
ACT_KINDS = {"created": "Заявка создана", "approved": "Одобрено", "rejected": "Отказ", "issued": "Выдано",
             "pending": "Ждёт решения", "returned": "Возвращено", "completed": "Аренда завершена"}
_ITEM_KIND = {P_OK: "approved", P_NO: "rejected", P_OUT: "issued", P_BACK: "returned"}


def _chunks(seq, n=100):
    seq = list(seq)
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


def activity_data(uid):
    orders = q("SELECT * FROM [Бронирование] WHERE [id_user] = ?", [uid])
    ev = []                                           # (datetime, kind, text)
    rids = [o["id_rent"] for o in orders]
    by_rid = {o["id_rent"]: o for o in orders}
    for o in orders:
        ev.append((as_dt(o["created"]), "created", f"Заявка №{o['id_rent']} создана"))
    items = []
    for part in _chunks(rids):
        items += q(f"SELECT * FROM [ПозицияБронирования] WHERE [id_rent] IN ({', '.join(['?'] * len(part))})", part)
    for p in items:
        if p["status"] == P_NEW:
            ev.append((as_dt(by_rid[p["id_rent"]]["created"]), "pending", f"Заявка №{p['id_rent']}: позиция №{p['id_item']} ждёт решения"))
    item_ids = [p["id_item"] for p in items]
    owner = {p["id_item"]: p["id_rent"] for p in items}
    for part in _chunks(item_ids):
        for h in q(f"SELECT * FROM [ИсторияСтатусовПозиций] WHERE [id_item] IN ({', '.join(['?'] * len(part))})", part):
            k = _ITEM_KIND.get(h["new_status"])
            if k:
                ev.append((as_dt(h["at"]), k, f"Заявка №{owner[h['id_item']]}: позиция №{h['id_item']} — {ACT_KINDS[k].lower()}"))
    for part in _chunks(rids):
        for h in q(f"SELECT * FROM [ИсторияСтатусов] WHERE [id_rent] IN ({', '.join(['?'] * len(part))})", part):
            if h["new_status"] == O_DONE:
                ev.append((as_dt(h["at"]), "completed", f"Заявка №{h['id_rent']} завершена"))
    ev = [e for e in ev if e[0]]
    ev.sort(key=lambda e: e[0], reverse=True)
    per_day = {}
    for dtv, k, _ in ev:
        per_day.setdefault(dtv.date(), {}).setdefault(k, 0)
        per_day[dtv.date()][k] += 1
    end = date.today()
    start = end - timedelta(days=364)
    start -= timedelta(days=start.weekday())          # неделя начинается с понедельника
    weeks, d = [], start
    while d <= end:
        col = []
        for _ in range(7):
            kinds = per_day.get(d, {}) if d <= end else {}
            total = sum(kinds.values())
            col.append({"date": d, "future": d > end, "total": total, "kinds": kinds,
                        "level": 0 if not total else 1 if total == 1 else 2 if total == 2 else 3 if total <= 4 else 4,
                        "rej": "rejected" in kinds, "pend": "pending" in kinds,
                        "title": f"{d.strftime('%d.%m.%Y')}: " + (", ".join(f"{ACT_KINDS[k].lower()} — {v}" for k, v in kinds.items()) or "нет событий")})
            d += timedelta(days=1)
        weeks.append(col)
    return {"weeks": weeks, "events": ev[:15], "total": len(ev), "active_days": len(per_day)}


# ---------------------------------------------------------------------------
# Запрос: пользователь, CSRF, общий контекст
# ---------------------------------------------------------------------------
@app.before_request
def before():
    g.user = None
    if request.endpoint in (None, "css", "img"):
        return
    if request.method == "POST":
        if not hmac.compare_digest(request.form.get("csrf", ""), session.get("_csrf", "")):
            abort(400)
    uid = session.get("uid")
    if uid:
        u = q1("SELECT * FROM [Пользователь] WHERE [id_user] = ?", [uid])
        if u and u["active"]:
            g.user = u
        else:
            session.pop("uid", None)
            if u:
                flash("Аккаунт заблокирован" + (f": {u['block_reason']}" if u["block_reason"] else "") + ".")


@app.context_processor
def inject_globals():
    ctx = dict(
        me=g.user, is_admin=bool(g.user and g.user["role"] == ROLE_ADMIN),
        is_b2b=bool(g.user and g.user["role"] == ROLE_B2B), unread=0, mode=session.get("mode", "b2c"),
        categories=[], cart_lines=[], cart_count=0, cart_rent=Decimal(0), cart_deposit=Decimal(0),
        current=request.full_path.rstrip("?"), point_rate=POINT_RATE, points_percent=int(POINTS_PERCENT),
        mail_enabled=bool(_env("TOOLRENT_SMTP_HOST")), today_iso=date.today().isoformat())
    try:      # страницы ошибок должны открываться, даже если база недоступна
        lines = cart_lines()
        ctx.update(categories=get_categories(), cart_lines=lines, cart_count=len(lines),
                   cart_rent=sum((l["rent"] for l in lines), Decimal(0)),
                   cart_deposit=sum((l["deposit"] for l in lines), Decimal(0)))
        if g.user:
            ctx["unread"] = q1("SELECT COUNT(*) AS n FROM [Уведомления] WHERE [id_user] = ? AND [Прочитано] = ?",
                               [g.user["id_user"], False])["n"]
    except DBUnavailable:
        pass
    return ctx


def login_required(f):
    @wraps(f)
    def w(*a, **k):
        if not g.user:
            flash("Войдите в аккаунт, чтобы продолжить.")
            return redirect(url_for("login", next=request.full_path.rstrip("?")))
        return f(*a, **k)
    return w


def admin_required(f):
    @wraps(f)
    @login_required
    def w(*a, **k):
        if g.user["role"] != ROLE_ADMIN:
            abort(403)
        return f(*a, **k)
    return w


def message_page(title, text, status=200, link=None):
    return render_template("message.html", title=title, text=text, link=link), status


@app.errorhandler(DBUnavailable)
def db_down(e):
    html = ("<!doctype html><meta charset=utf-8><title>ToolRent</title><body style='font:16px system-ui;background:#0a0d12;"
            "color:#f3f4f6;display:grid;place-items:center;min-height:100vh;margin:0'><div style='max-width:560px;padding:24px'>"
            "<h1>База данных недоступна</h1><p>ToolRent не смог открыть <b>database.accdb</b>.</p>"
            f"<p style='color:#98a0ad'>Причина: {escape(str(e))}</p>"
            "<p>Проверьте драйвер Access и откройте <a style='color:#ffb020' href='/db-check'>/db-check</a>.</p></div>")
    return Response(html, 503)


@app.errorhandler(400)
def bad_request(_):
    return message_page("Сессия устарела", "Форма была отправлена с устаревшей страницы. Вернитесь назад, обновите страницу и повторите.", 400)


@app.errorhandler(403)
def forbidden(_):
    return message_page("Нет доступа", "У вас нет прав для просмотра этой страницы.", 403)


@app.errorhandler(404)
def not_found(_):
    return message_page("Страница не найдена", "Такой страницы нет. Попробуйте найти нужный инструмент в каталоге.", 404,
                        link=(url_for("catalog"), "Перейти в каталог"))


@app.errorhandler(413)
def too_big(_):
    return message_page("Файл слишком большой", "Максимальный размер загрузки — 8 МБ.", 413)


# ---------------------------------------------------------------------------
# Файлы и публичные страницы
# ---------------------------------------------------------------------------
@app.route("/style.css")
def css():
    return send_from_directory(BASE_DIR, "style.css", mimetype="text/css")


@app.route("/img/<path:name>")
def img(name):
    return send_from_directory(IMG_DIR, name)


def catalog_stats():
    tools = list_tools()
    return {"positions": len(tools), "units": sum(t["stock"] for t in tools), "free": sum(t["qty"] for t in tools),
            "categories": len({t["cat"] for t in tools})}


@app.route("/")
def index():
    tools = list_tools()
    return render_template("index.html", popular=sorted(tools, key=lambda t: -t["id"])[:4],
                           min_price=min((t["price"] for t in tools), default=None), stats=catalog_stats())


@app.route("/catalog")
def catalog():
    q_ = request.args.get("q", "").strip()
    cat, brand = to_int(request.args.get("cat")), to_int(request.args.get("brand"))
    avail, sort = request.args.get("avail") == "1", request.args.get("sort", "")
    result = list_tools(cat=cat, brand=brand, avail=avail, sort=sort)
    if q_:       # поиск в Python: регистр кириллицы считается одинаково на любом драйвере
        ql = q_.casefold()
        result = [t for t in result if ql in t["name"].casefold() or ql in t["brand"].casefold()]
    brands = sorted({(t["brand_id"], t["brand"]) for t in list_tools() if t["brand"]}, key=lambda b: b[1].casefold())
    return render_template("catalog.html", tools=result, brands=[{"id": i, "name": n} for i, n in brands],
                           q=q_, cat=cat, brand=brand, avail=avail, sort=sort, active_cat=get_category(cat))


@app.route("/tool/<int:tool_id>")
def tool(tool_id):
    t = get_tool(tool_id)
    if not t:
        abort(404)
    similar = [x for x in list_tools(cat=t["cat"]) if x["id"] != t["id"]][:3]
    return render_template("tool.html", t=t, durations=DURATIONS, category=get_category(t["cat"]), similar=similar)


@app.route("/mode/<m>")
def set_mode(m):
    if m in ("b2c", "b2b"):
        session["mode"] = m
    return redirect(safe_next())


@app.route("/b2b")
def b2b():
    return render_template("b2b.html", stats=catalog_stats(), org=q1("SELECT * FROM [B2B_Реквизиты] WHERE [id_user] = ?", [g.user["id_user"]]) if g.user else None)


# ---------------------------------------------------------------------------
# Корзина и оформление заявки
# ---------------------------------------------------------------------------
@app.route("/cart/add", methods=["POST"])
def cart_add():
    try:
        start = parse_date(request.form.get("start")) or today()
        days = to_int(request.form.get("days"), 1, 1, MAX_RENT_DAYS)
        end = parse_date(request.form.get("end")) or start + timedelta(days=days - 1)
        cart_add_line(to_int(request.form.get("tool_id")), to_int(request.form.get("qty"), 1, 1, 9999), start, end)
        return redirect(safe_next() + "#cart")
    except ValueError as e:
        flash(str(e))
        return redirect(safe_next())


@app.route("/cart/remove/<k>", methods=["POST"])
def cart_remove(k):
    session["cart"] = [l for l in session.get("cart", []) if l["k"] != k]
    return redirect(safe_next("booking") + ("#cart" if request.values.get("next", "") != "/booking" else ""))


@app.route("/cart/clear", methods=["POST"])
def cart_clear():
    session["cart"] = []
    return redirect(safe_next() + ("#cart" if request.values.get("next", "") != "/booking" else ""))


def apply_line_edits():
    """Сохраняет изменения количества и дат из формы оформления."""
    cart, errors = session.get("cart", []), []
    for l in cart:
        k = l["k"]
        if f"qty-{k}" not in request.form:
            continue
        qty = to_int(request.form.get(f"qty-{k}"), l["q"], 1, 9999)
        s, e = parse_date(request.form.get(f"start-{k}")), parse_date(request.form.get(f"end-{k}"))
        try:
            validate_period(s, e)
            l["q"], l["s"], l["e"] = qty, s.strftime("%Y-%m-%d"), e.strftime("%Y-%m-%d")
        except ValueError as ex:
            errors.append(str(ex))
    session["cart"] = cart
    return errors


def booking_view(selected_uc=0, points=0, delivery="Самовывоз"):
    lines = cart_lines()
    subtotal = sum((l["rent"] for l in lines), Decimal(0))
    deposit = sum((l["deposit"] for l in lines), Decimal(0))
    coupons, disc_c, disc_p = [], Decimal(0), Decimal(0)
    # свободное количество с учётом других позиций этой же заявки
    taken = []
    with Tx() as t:
        for l in lines:
            free = tool_free(t, l["tool"]["id"], l["tool"]["stock"], l["start"], l["end"])
            free -= sum(q_ for (tid, s, e, q_) in taken if tid == l["tool"]["id"] and s <= l["end"] and e >= l["start"])
            l["free"] = max(free, 0)
            taken.append((l["tool"]["id"], l["start"], l["end"], l["qty"]))
        if g.user:
            coupons = usable_coupons(t, g.user["id_user"], subtotal)
    chosen = next((c for c in coupons if c["uc"] == selected_uc and c["ok"]), None)
    if chosen:
        disc_c = chosen["discount"]
    if points and POINT_RATE and g.user:
        points = min(points, g.user["points"])
        disc_p = money(min(D(points) * POINT_RATE, subtotal - disc_c))
        points = int((disc_p / POINT_RATE).to_integral_value(ROUND_DOWN))
    else:
        points = 0
    total = money(subtotal - disc_c - disc_p)
    return dict(lines=lines, subtotal=money(subtotal), deposit=money(deposit), coupons=coupons, uc=selected_uc if chosen else 0,
                disc_c=disc_c, disc_p=disc_p, points=points, total=total, delivery=delivery,
                has_problem=any(l["qty"] > l["free"] for l in lines))


@app.route("/booking", methods=["GET", "POST"])
def booking():
    uc, points, delivery, errors = 0, 0, "Самовывоз", []
    if request.method == "POST":
        errors = apply_line_edits()
        uc, points = to_int(request.form.get("coupon")), to_int(request.form.get("points"))
        delivery = request.form.get("delivery", "Самовывоз")
        action = request.form.get("action", "update")
        if action == "submit" and not errors:
            if not g.user:
                flash("Войдите в аккаунт, чтобы отправить заявку.")
                return redirect(url_for("login", next=url_for("booking")))
            if not g.user["email_ok"]:
                errors.append("Подтвердите электронную почту, чтобы отправить заявку (письмо можно отправить повторно в личном кабинете).")
            else:
                try:
                    view = booking_view(uc, points, delivery)
                    rid = create_order(g.user, view["lines"], view["uc"], view["points"], delivery)
                    session["cart"] = []
                    flash(f"Заявка №{rid} отправлена. Статус можно отслеживать в личном кабинете.")
                    return redirect(url_for("profile_order", rid=rid))
                except ValueError as e:
                    errors.append(str(e))
    for e in errors:
        flash(e)
    return render_template("booking.html", **booking_view(uc, points, delivery))


# ---------------------------------------------------------------------------
# Регистрация, вход, подтверждение почты, восстановление пароля
# ---------------------------------------------------------------------------
LOGIN_RE = re.compile(r"^[A-Za-z0-9_.-]{3,32}$")
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]{2,}$")


def check_new_password(pwd, pwd2, login=""):
    if len(pwd) < 8:
        raise ValueError("Пароль должен быть не короче 8 символов.")
    if len(pwd) > 128:
        raise ValueError("Пароль слишком длинный (максимум 128 символов).")
    if not (re.search(r"[A-Za-zА-Яа-я]", pwd) and re.search(r"\d", pwd)):
        raise ValueError("В пароле должны быть и буквы, и цифры.")
    if login and pwd.lower() == login.lower():
        raise ValueError("Пароль не должен совпадать с логином.")
    if pwd != pwd2:
        raise ValueError("Пароли не совпадают.")


def check_phone(v):
    if v and not 10 <= len(re.sub(r"\D", "", v)) <= 15:
        raise ValueError("Телефон указан неверно.")
    return v


def check_email(v):
    if not EMAIL_RE.match(v) or len(v) > 255:
        raise ValueError("Email указан неверно.")
    return v.lower()


def send_verification(uid, email):
    with Tx() as t:
        raw = make_token(t, uid, "email")
    send_email(email, "Подтвердите почту на ToolRent",
               f"Здравствуйте!\n\nЧтобы подтвердить адрес электронной почты, перейдите по ссылке (действует {TOKEN_HOURS['email']} ч.):\n"
               f"{base_url()}/verify/{raw}\n\nЕсли вы не регистрировались на ToolRent, просто проигнорируйте это письмо.")


def login_user(uid, remember=True):
    session.clear()
    session["uid"] = uid
    session.permanent = remember
    csrf_token()


@app.route("/register", methods=["GET", "POST"])
def register():
    if g.user:
        return redirect(url_for("profile"))
    if request.method == "POST":
        try:
            fio = field("fio", 255, True, "ФИО")
            login = field("login", 32, True, "Логин")
            email = check_email(field("email", 255, True, "Email"))
            phone = check_phone(field("phone", 255, label="Телефон"))
            address = field("address", 255, label="Адрес")
            if not LOGIN_RE.match(login):
                raise ValueError("Логин: 3–32 символа, латинские буквы, цифры, точка, дефис и подчёркивание.")
            pwd = request.form.get("password", "")
            check_new_password(pwd, request.form.get("password2", ""), login)
            with Tx() as t:
                if t.one("SELECT [id_user] FROM [Пользователь] WHERE [Логин] = ? OR [Email] = ?", [login, email]):
                    raise ValueError("Пользователь с таким логином или email уже зарегистрирован.")
                uid = t.insert("Пользователь", {
                    "Логин": login, "Пароль": hash_password(pwd), "ФИО": fio, "Телефон": phone or None, "Email": email,
                    "Адрес": address or None, "Роль": ROLE_USER, "Баллы": 0, "EmailПодтвержден": False, "Активен": True,
                    "ДатаРегистрации": now()})
                if t.one("SELECT COUNT(*) AS n FROM [Пользователь] WHERE [Логин] = ? OR [Email] = ?", [login, email])["n"] > 1:
                    raise ValueError("Пользователь с таким логином или email уже зарегистрирован.")
        except ValueError as e:
            flash(str(e))
            return render_template("register.html", form=request.form), 400
        send_verification(uid, email)
        login_user(uid)
        flash("Аккаунт создан. " + ("Мы отправили письмо со ссылкой для подтверждения почты." if _env("TOOLRENT_SMTP_HOST")
                                    else "Отправка почты на сервере не настроена: подтвердить адрес может администратор."))
        return redirect(url_for("profile"))
    return render_template("register.html", form={})


@app.route("/login", methods=["GET", "POST"])
def login():
    if g.user:
        return redirect(safe_next("profile"))
    if request.method == "POST":
        ident = (request.form.get("login") or "").strip()
        key = (request.remote_addr, ident.lower())
        if too_many_failures(key):
            flash("Слишком много неудачных попыток. Попробуйте через несколько минут.")
            return render_template("login.html"), 429
        u = q1("SELECT * FROM [Пользователь] WHERE [Логин] = ? OR [Email] = ?", [ident, ident.lower()]) if ident else None
        if not u or not check_password_hash(u["pwd"] or "", request.form.get("password", "")):
            note_failure(key)
            flash("Неверный логин или пароль.")
            return render_template("login.html"), 401
        if not u["active"]:
            flash("Аккаунт заблокирован" + (f": {u['block_reason']}" if u["block_reason"] else "") + ". Обратитесь в поддержку.")
            return render_template("login.html"), 403
        nxt = safe_next("profile")
        login_user(u["id_user"], remember=request.form.get("remember") == "1")
        return redirect(nxt)
    return render_template("login.html")


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    flash("Вы вышли из аккаунта.")
    return redirect(url_for("index"))


@app.route("/verify/<token>")
def verify_email(token):
    with Tx() as t:
        row = find_token(t, token, "email")
        if not row:
            return message_page("Ссылка недействительна", "Ссылка подтверждения устарела или уже использована. "
                                "Запросите новое письмо в личном кабинете.", 400, link=(url_for("profile"), "Личный кабинет"))
        t.x("UPDATE [Пользователь] SET [EmailПодтвержден] = ? WHERE [id_user] = ?", [True, row["id_user"]])
        t.x("UPDATE [ТокеныАккаунта] SET [Использован] = ? WHERE [id_token] = ?", [True, row["id_token"]])
    return message_page("Почта подтверждена", "Спасибо! Теперь можно отправлять заявки на аренду.",
                        link=(url_for("catalog"), "Перейти в каталог"))


@app.route("/verify/resend", methods=["POST"])
@login_required
def verify_resend():
    u = g.user
    if u["email_ok"]:
        flash("Почта уже подтверждена.")
    else:
        last = q1("SELECT [ДатаСоздания] FROM [ТокеныАккаунта] WHERE [id_user] = ? AND [Тип] = ? ORDER BY [ДатаСоздания] DESC",
                  [u["id_user"], "email"])
        if last and (now() - as_dt(last["created"])).total_seconds() < 60:
            flash("Письмо уже отправлено. Повторить можно через минуту.")
        else:
            send_verification(u["id_user"], u["email"])
            flash("Письмо отправлено повторно." if _env("TOOLRENT_SMTP_HOST") else
                  "Отправка почты на сервере не настроена: подтвердить адрес может администратор.")
    return redirect(safe_next("profile"))


@app.route("/forgot", methods=["GET", "POST"])
def forgot():
    if request.method == "POST":
        email = (request.form.get("email") or "").strip().lower()
        u = q1("SELECT * FROM [Пользователь] WHERE [Email] = ?", [email]) if email else None
        if u and u["active"] and not too_many_failures(("reset", email)):
            note_failure(("reset", email))
            with Tx() as t:
                raw = make_token(t, u["id_user"], "reset")
            send_email(u["email"], "Восстановление пароля ToolRent",
                       f"Чтобы задать новый пароль, перейдите по ссылке (действует {TOKEN_HOURS['reset']} ч.):\n"
                       f"{base_url()}/reset/{raw}\n\nЕсли вы не запрашивали восстановление, проигнорируйте письмо.")
        flash("Если такой адрес зарегистрирован, мы отправили на него письмо со ссылкой для восстановления.")
        return redirect(url_for("forgot"))
    return render_template("forgot.html", token=None)


@app.route("/reset/<token>", methods=["GET", "POST"])
def reset_password(token):
    with Tx() as t:
        row = find_token(t, token, "reset")
        if not row:
            return message_page("Ссылка недействительна", "Ссылка для восстановления устарела или уже использована.", 400,
                                link=(url_for("forgot"), "Запросить новую"))
        if request.method == "POST":
            try:
                u = t.one("SELECT * FROM [Пользователь] WHERE [id_user] = ?", [row["id_user"]])
                check_new_password(request.form.get("password", ""), request.form.get("password2", ""), u["login"])
            except ValueError as e:
                flash(str(e))
                return render_template("forgot.html", token=token), 400
            t.x("UPDATE [Пользователь] SET [Пароль] = ? WHERE [id_user] = ?", [hash_password(request.form["password"]), row["id_user"]])
            t.x("UPDATE [ТокеныАккаунта] SET [Использован] = ? WHERE [id_token] = ?", [True, row["id_token"]])
            flash("Пароль изменён. Теперь можно войти.")
            return redirect(url_for("login"))
    return render_template("forgot.html", token=token)


# ---------------------------------------------------------------------------
# Личный кабинет
# ---------------------------------------------------------------------------
ORG_TYPES = ["ИП", "ООО", "АО", "ПАО", "Самозанятый", "Другое"]
KPP_REQUIRED = ("ООО", "АО", "ПАО")
SECTIONS = [("overview", "Личные данные"), ("orders", "История аренд"), ("points", "Баллы"), ("coupons", "Купоны"),
            ("messages", "Сообщения"), ("activity", "Карта активности"), ("password", "Пароль"), ("business", "Бизнес-профиль")]


def orders_with_counts(uid):
    orders = sorted(q("SELECT * FROM [Бронирование] WHERE [id_user] = ?", [uid]), key=lambda o: -o["id_rent"])
    counts = {}
    for part in _chunks([o["id_rent"] for o in orders]):
        for p in q(f"SELECT * FROM [ПозицияБронирования] WHERE [id_rent] IN ({', '.join(['?'] * len(part))})", part):
            counts[p["id_rent"]] = counts.get(p["id_rent"], 0) + 1
    for o in orders:
        o["n_items"] = counts.get(o["id_rent"], 0)
    return orders


@app.route("/profile")
@app.route("/profile/<section>")
@login_required
def profile(section="overview"):
    if section not in dict(SECTIONS):
        abort(404)
    u = g.user
    data = {}
    if section == "overview":
        data = dict(stats=user_stats(u["id_user"]), recent=orders_with_counts(u["id_user"])[:3],
                    org=q1("SELECT * FROM [B2B_Реквизиты] WHERE [id_user] = ?", [u["id_user"]]),
                    n_coupons=len([c for c in q_coupons(u["id_user"]) if c["state"] == "Доступен"]))
    elif section == "orders":
        data = dict(orders=orders_with_counts(u["id_user"]))
    elif section == "points":
        ops = q("SELECT [id_operation], [id_rent], [Количество] AS delta, [ТипОперации], [Дата], [Комментарий] "
                "FROM [ОперацииБаллов] WHERE [id_user] = ?", [u["id_user"]])
        data = dict(ops=sorted(ops, key=lambda o: o["id_operation"], reverse=True))
    elif section == "coupons":
        data = dict(coupons=q_coupons(u["id_user"]))
    elif section == "messages":
        data = dict(msgs=sorted(q("SELECT * FROM [Уведомления] WHERE [id_user] = ?", [u["id_user"]]),
                                key=lambda m: m["id_notification"], reverse=True)[:100])
    elif section == "activity":
        data = dict(act=activity_data(u["id_user"]), kinds=ACT_KINDS)
    elif section == "business":
        data = dict(org=q1("SELECT * FROM [B2B_Реквизиты] WHERE [id_user] = ?", [u["id_user"]]), org_types=ORG_TYPES)
    return render_template("profile.html", section=section, sections=SECTIONS, **data)


def q_coupons(uid):
    out = []
    with Tx() as t:
        for uc in t.q("SELECT * FROM [КупонПользователя] WHERE [id_user] = ?", [uid]):
            c = t.one("SELECT * FROM [Купон] WHERE [id_coupon] = ?", [uc["id_coupon"]])
            if not c:
                continue
            end = as_dt(uc["end"]) or as_dt(c["end"])
            state = ("Использован" if uc["used"] else "Не активен" if not c["active"] else
                     "Срок истёк" if end and end + timedelta(days=1) <= now() else
                     "Ещё не действует" if as_dt(c["start"]) and as_dt(c["start"]) > now() else "Доступен")
            out.append({"c": c, "end": end, "state": state, "issued": uc["issued"]})
    return sorted(out, key=lambda x: (x["state"] != "Доступен", x["end"] or datetime.max))


@app.route("/profile/orders/<int:rid>")
@login_required
def profile_order(rid):
    order = q1("SELECT * FROM [Бронирование] WHERE [id_rent] = ? AND [id_user] = ?", [rid, g.user["id_user"]])
    if not order:
        abort(404)
    pos = sorted(q("SELECT * FROM [ПозицияБронирования] WHERE [id_rent] = ?", [rid]), key=lambda p: p["id_item"])
    tools = tools_by_ids({p["id_tool"] for p in pos}, include_inactive=True)
    hist = {p["id_item"]: [] for p in pos}
    for part in _chunks(hist):
        for h in q(f"SELECT * FROM [ИсторияСтатусовПозиций] WHERE [id_item] IN ({', '.join(['?'] * len(part))})", part):
            hist[h["id_item"]].append(h)
    for p in pos:
        p["tool"] = tools.get(p["id_tool"])
        p["days"] = rent_days(p["start"], p["end"])
        p["sum"] = D(p["price_at"]) * p["qty"] * p["days"]
        p["history"] = sorted(hist[p["id_item"]], key=lambda h: h["id_log_item"] if "id_log_item" in h else 0)
        rej = [h for h in p["history"] if h["new_status"] == P_NO and h["comment"]]
        p["reason"] = clean_text(rej[-1]["comment"]) if rej else ""
    history = sorted(q("SELECT * FROM [ИсторияСтатусов] WHERE [id_rent] = ?", [rid]), key=lambda h: h["id_log"])
    can_cancel = order["status"] in (O_NEW, O_WORK) and not any(p["status"] in (P_OUT, P_BACK) for p in pos)
    return render_template("profile.html", section="order", sections=SECTIONS, order=order, positions=pos,
                           history=history, can_cancel=can_cancel)


@app.route("/profile/orders/<int:rid>/cancel", methods=["POST"])
@login_required
def profile_order_cancel(rid):
    try:
        user_cancel_order(g.user, rid)
        flash(f"Заявка №{rid} отменена.")
    except ValueError as e:
        flash(str(e))
    return redirect(url_for("profile_order", rid=rid))


@app.route("/profile/data", methods=["POST"])
@login_required
def profile_data():
    u = g.user
    try:
        fio = field("fio", 255, True, "ФИО")
        phone = check_phone(field("phone", 255, label="Телефон"))
        address = field("address", 255, label="Адрес")
        email = check_email(field("email", 255, True, "Email"))
        changed = email != (u["email"] or "").lower()
        with Tx() as t:
            if changed and t.one("SELECT [id_user] FROM [Пользователь] WHERE [Email] = ? AND [id_user] <> ?", [email, u["id_user"]]):
                raise ValueError("Этот email уже используется другим аккаунтом.")
            t.update("Пользователь", "id_user", u["id_user"], {"ФИО": fio, "Телефон": phone or None, "Адрес": address or None,
                                                                 **({"Email": email, "EmailПодтвержден": False} if changed else {})})
        if changed:
            send_verification(u["id_user"], email)
            flash("Данные сохранены. На новый адрес отправлено письмо для подтверждения." if _env("TOOLRENT_SMTP_HOST")
                  else "Данные сохранены. Новый адрес нужно подтвердить: отправка почты не настроена, это может сделать администратор.")
        else:
            flash("Данные сохранены.")
    except ValueError as e:
        flash(str(e))
    return redirect(url_for("profile"))


@app.route("/profile/password", methods=["POST"])
@login_required
def profile_password():
    u = g.user
    try:
        if not check_password_hash(u["pwd"] or "", request.form.get("current", "")):
            raise ValueError("Текущий пароль указан неверно.")
        check_new_password(request.form.get("password", ""), request.form.get("password2", ""), u["login"])
        q_ = Tx()
        with q_ as t:
            t.x("UPDATE [Пользователь] SET [Пароль] = ? WHERE [id_user] = ?", [hash_password(request.form["password"]), u["id_user"]])
        flash("Пароль изменён.")
    except ValueError as e:
        flash(str(e))
    return redirect(url_for("profile", section="password"))


@app.route("/profile/messages/read", methods=["POST"])
@login_required
def messages_read():
    with Tx() as t:
        t.x("UPDATE [Уведомления] SET [Прочитано] = ? WHERE [id_user] = ? AND [Прочитано] = ?", [True, g.user["id_user"], False])
    return redirect(url_for("profile", section="messages"))


@app.route("/profile/business", methods=["POST"])
@login_required
def business_save():
    u = g.user
    if u["role"] == ROLE_ADMIN:
        flash("Администратору бизнес-профиль не требуется.")
        return redirect(url_for("profile"))
    try:
        org = field("org_type", 50, True, "Тип организации")
        if org not in ORG_TYPES:
            raise ValueError("Выберите тип организации из списка.")
        company = field("company", 255, True, "Название компании")
        inn = re.sub(r"\s", "", field("inn", 20, True, "ИНН"))
        kpp = re.sub(r"\s", "", field("kpp", 20, label="КПП"))
        addr = field("legal_address", 255, True, "Юридический адрес")
        if not inn.isdigit() or len(inn) not in ((12,) if org in ("ИП", "Самозанятый") else (10,) if org in KPP_REQUIRED else (10, 12)):
            raise ValueError("ИНН указан неверно: у " + ("ИП и самозанятых 12 цифр." if org in ("ИП", "Самозанятый") else
                             "организаций 10 цифр." if org in KPP_REQUIRED else "организации 10 цифр, у предпринимателя 12."))
        if org in KPP_REQUIRED and not kpp:
            raise ValueError("Для выбранной формы организации КПП обязателен.")
        if kpp and not (kpp.isdigit() and len(kpp) == 9):
            raise ValueError("КПП должен состоять из 9 цифр.")
        with Tx() as t:
            row = t.one("SELECT * FROM [B2B_Реквизиты] WHERE [id_user] = ?", [u["id_user"]])
            vals = {"ТипОрганизации": org, "НазваниеКомпании": company, "ИНН": inn, "КПП": kpp or None, "ЮрАдрес": addr}
            if row:
                t.update("B2B_Реквизиты", "id_b2b", row["id_b2b"], vals)
            else:
                t.insert("B2B_Реквизиты", {"id_user": u["id_user"], **vals})
            if u["role"] == ROLE_USER:
                t.update("Пользователь", "id_user", u["id_user"], {"Роль": ROLE_B2B})
        if u["role"] == ROLE_USER:
            notify(u["id_user"], "Бизнес-профиль подключён", f"Реквизиты «{company}» сохранены. Заказы, баллы и купоны остались в вашем аккаунте.", "Система")
            flash("Бизнес-профиль подключён.")
        else:
            flash("Реквизиты сохранены.")
    except ValueError as e:
        flash(str(e))
    return redirect(url_for("profile", section="business"))


# ---------------------------------------------------------------------------
# Администрирование (все маршруты защищены декоратором admin_required)
# ---------------------------------------------------------------------------
def adm(section, **data):
    return render_template("admin.html", section=section, **data)


def user_map(ids):
    ids = [i for i in set(ids) if i]
    out = {}
    for part in _chunks(ids):
        for u in q(f"SELECT * FROM [Пользователь] WHERE [id_user] IN ({', '.join(['?'] * len(part))})", part):
            out[u["id_user"]] = u
    return out


@app.route("/admin")
@admin_required
def admin_home():
    pos = q("SELECT * FROM [ПозицияБронирования]")
    count = lambda s: len([p for p in pos if p["status"] == s])
    overdue = [p for p in pos if p["status"] == P_OUT and as_dt(p["end"]) and as_dt(p["end"]) < today()]
    users = q("SELECT * FROM [Пользователь]")
    recent = sorted(q("SELECT * FROM [Бронирование]"), key=lambda o: -o["id_rent"])[:8]
    um = user_map(o["id_user"] for o in recent)
    return adm("home", n_new=count(P_NEW), n_ok=count(P_OK), n_out=count(P_OUT), n_over=len(overdue),
               n_unconfirmed=len([u for u in users if not u["email_ok"]]), n_blocked=len([u for u in users if not u["active"]]),
               recent=recent, um=um)


# --- инструменты ---------------------------------------------------------
def sniff_image(data):
    if data[:3] == b"\xff\xd8\xff":
        return "jpg"
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "webp"
    return None


@app.route("/admin/tools")
@admin_required
def admin_tools():
    q_, state = request.args.get("q", "").strip().casefold(), request.args.get("state", "")
    tools = list_tools(include_inactive=True)
    if q_:
        tools = [t for t in tools if q_ in t["name"].casefold() or q_ in t["brand"].casefold()]
    if state in ("1", "0"):
        tools = [t for t in tools if t["active"] == (state == "1")]
    return adm("tools", tools=tools, q=request.args.get("q", ""), state=state)


@app.route("/admin/tools/new")
@app.route("/admin/tools/<int:tool_id>")
@admin_required
def admin_tool_form(tool_id=0):
    t = get_tool(tool_id, include_inactive=True) if tool_id else None
    if tool_id and not t:
        abort(404)
    return adm("tool_form", t=t, brands=get_brands(), cats=get_categories())


@app.route("/admin/tools/save", methods=["POST"])
@admin_required
def admin_tool_save():
    tool_id = to_int(request.form.get("id"))
    try:
        vals = {"Название": field("name", 255, True, "Название"), "Описание": field("descr", 20000) or None,
                "Характеристики": field("specs", 20000) or None,
                "ЦенаЗаСутки": to_money(request.form.get("price", ""), "Цена за сутки"),
                "Залог": to_money(request.form.get("deposit", "0") or "0", "Залог"),
                "КоличествоНаСкладе": to_int(request.form.get("stock"), 0, 0, 100000),
                "Активен": request.form.get("active") == "1"}
        if vals["ЦенаЗаСутки"] <= 0:
            raise ValueError("Цена за сутки должна быть больше нуля.")
        brand, cat = to_int(request.form.get("brand")), to_int(request.form.get("cat"))
        if brand not in get_brands() or cat not in {c["id"] for c in get_categories()}:
            raise ValueError("Выберите бренд и категорию.")
        vals["id_brand"], vals["id_category"] = brand, cat
        fs = request.files.get("photo")
        data = fs.read() if fs and fs.filename else b""
        ext = sniff_image(data) if data else None
        if data and (not ext or len(data) > 5 * 1024 * 1024):
            raise ValueError("Фото: допустимы JPG, PNG или WEBP размером до 5 МБ.")
        with Tx() as t:
            old_photo = None
            if tool_id:
                old = t.one("SELECT * FROM [Инструмент] WHERE [id_tool] = ?", [tool_id])
                if not old:
                    raise ValueError("Инструмент не найден.")
                old_photo = old["photo"]
                t.update("Инструмент", "id_tool", tool_id, vals)
            else:
                tool_id = t.insert("Инструмент", vals)
            if data:
                name = f"tool_{tool_id}_{secrets.token_hex(4)}.{ext}"
                with open(os.path.join(IMG_DIR, name), "wb") as f:
                    f.write(data)
                t.update("Инструмент", "id_tool", tool_id, {"ИмяФото": name})
                if old_photo and str(old_photo).startswith("tool_") and os.path.basename(old_photo) == old_photo:
                    try:           # удаляем только файлы, загруженные через админку
                        os.remove(os.path.join(IMG_DIR, old_photo))
                    except OSError:
                        pass
            admin_log(t, g.user["id_user"], "Инструмент сохранён", None,
                      f"Инструмент №{tool_id} «{vals['Название']}»" + (", фото заменено" if data else "") +
                      ("" if vals["Активен"] else ", снят с аренды"))
        flash("Инструмент сохранён.")
        return redirect(url_for("admin_tool_form", tool_id=tool_id))
    except ValueError as e:
        flash(str(e))
        return redirect(url_for("admin_tool_form", tool_id=tool_id) if tool_id else url_for("admin_tool_form"))


@app.route("/admin/tools/<int:tool_id>/toggle", methods=["POST"])
@admin_required
def admin_tool_toggle(tool_id):
    with Tx() as t:
        row = t.one("SELECT * FROM [Инструмент] WHERE [id_tool] = ?", [tool_id])
        if not row:
            abort(404)
        t.update("Инструмент", "id_tool", tool_id, {"Активен": not row["active"]})
        admin_log(t, g.user["id_user"], "Инструмент " + ("снят с аренды" if row["active"] else "активирован"), None,
                  f"Инструмент №{tool_id} «{clean_text(row['name'])}»")
    return redirect(request.referrer or url_for("admin_tools"))


# --- категории и бренды --------------------------------------------------
DICTS = {"brand": ("Бренд", "id_brand", "id_brand"), "category": ("Категория", "id_category", "id_category")}


@app.route("/admin/dicts", methods=["GET"])
@admin_required
def admin_dicts():
    tools = q("SELECT [id_brand], [id_category] FROM [Инструмент]")
    nb, nc = {}, {}
    for t_ in tools:
        nb[t_["id_brand"]] = nb.get(t_["id_brand"], 0) + 1
        nc[t_["id_category"]] = nc.get(t_["id_category"], 0) + 1
    return adm("dicts", brands=get_brands(), cat_rows=[(c["id"], c["name"]) for c in get_categories()], nb=nb, nc=nc)


@app.route("/admin/dicts/<kind>/<action>", methods=["POST"])
@admin_required
def admin_dict_action(kind, action):
    if kind not in DICTS or action not in ("add", "rename", "delete"):
        abort(404)
    table, key, col = DICTS[kind]
    label = "Бренд" if kind == "brand" else "Категория"
    try:
        with Tx() as t:
            if action in ("add", "rename"):
                name = field("name", 255, True, "Название")
                if t.one(f"SELECT [{key}] FROM [{table}] WHERE [Название] = ?" + (f" AND [{key}] <> ?" if action == "rename" else ""),
                         [name] + ([to_int(request.form.get("id"))] if action == "rename" else [])):
                    raise ValueError("Такое название уже есть.")
            if action == "add":
                new_id = t.insert(table, {"Название": name})
                admin_log(t, g.user["id_user"], f"{label}: создан", None, f"№{new_id} «{name}»")
            else:
                rid = to_int(request.form.get("id"))
                row = t.one(f"SELECT * FROM [{table}] WHERE [{key}] = ?", [rid])
                if not row:
                    raise ValueError("Запись не найдена.")
                if action == "rename":
                    t.update(table, key, rid, {"Название": name})
                    admin_log(t, g.user["id_user"], f"{label}: переименован", None, f"№{rid}: «{clean_text(row['name'])}» → «{name}»")
                else:
                    if t.one("SELECT COUNT(*) AS n FROM [Инструмент] WHERE [" + col + "] = ?", [rid])["n"]:
                        raise ValueError("Удалить нельзя: к записи привязаны инструменты.")
                    t.x(f"DELETE FROM [{table}] WHERE [{key}] = ?", [rid])
                    admin_log(t, g.user["id_user"], f"{label}: удалён", None, f"№{rid} «{clean_text(row['name'])}»")
        flash("Готово.")
    except ValueError as e:
        flash(str(e))
    return redirect(url_for("admin_dicts"))


# --- заявки --------------------------------------------------------------
@app.route("/admin/orders")
@admin_required
def admin_orders():
    status = request.args.get("status", "")
    orders = sorted(q("SELECT * FROM [Бронирование]"), key=lambda o: -o["id_rent"])
    if status:
        orders = [o for o in orders if o["status"] == status]
    return adm("orders", orders=orders[:300], um=user_map(o["id_user"] for o in orders[:300]), status=status,
               statuses=[O_NEW, O_WORK, O_RENT, O_DONE, O_NO, O_CANCEL])


@app.route("/admin/orders/<int:rid>")
@admin_required
def admin_order(rid):
    order = q1("SELECT * FROM [Бронирование] WHERE [id_rent] = ?", [rid])
    if not order:
        abort(404)
    pos = sorted(q("SELECT * FROM [ПозицияБронирования] WHERE [id_rent] = ?", [rid]), key=lambda p: p["id_item"])
    tools = tools_by_ids({p["id_tool"] for p in pos}, include_inactive=True)
    um = user_map(h["by_user"] for h in q("SELECT * FROM [ИсторияСтатусов] WHERE [id_rent] = ?", [rid]))
    um.update(user_map([order["id_user"]]))
    hist = {p["id_item"]: [] for p in pos}
    for part in _chunks(hist):
        for h in q(f"SELECT * FROM [ИсторияСтатусовПозиций] WHERE [id_item] IN ({', '.join(['?'] * len(part))})", part):
            hist[h["id_item"]].append(h)
    with Tx() as t:
        for p in pos:
            p["tool"] = tools.get(p["id_tool"])
            p["days"] = rent_days(p["start"], p["end"])
            p["sum"] = D(p["price_at"]) * p["qty"] * p["days"]
            p["free"] = tool_free(t, p["id_tool"], p["tool"]["stock"], as_dt(p["start"]), as_dt(p["end"])) if p["tool"] else 0
            p["history"] = sorted(hist[p["id_item"]], key=lambda h: h["id_item_log"])
    history = sorted(q("SELECT * FROM [ИсторияСтатусов] WHERE [id_rent] = ?", [rid]), key=lambda h: h["id_log"])
    um.update(user_map(h["by_user"] for p in pos for h in p["history"]))
    return adm("order", order=order, positions=pos, history=history, um=um, customer=um.get(order["id_user"]))


@app.route("/admin/orders/<int:rid>/item", methods=["POST"])
@admin_required
def admin_order_item(rid):
    item = to_int(request.form.get("item_id"))
    try:
        it = q1("SELECT * FROM [ПозицияБронирования] WHERE [id_item] = ? AND [id_rent] = ?", [item, rid])
        if not it:
            raise ValueError("Позиция не найдена.")
        admin_item_action(g.user, item, request.form.get("action", ""), field("reason", 255))
        flash("Готово.")
    except ValueError as e:
        flash(str(e))
    return redirect(url_for("admin_order", rid=rid))


# --- пользователи --------------------------------------------------------
@app.route("/admin/users")
@admin_required
def admin_users():
    q_ = request.args.get("q", "").strip().casefold()
    users = q("SELECT * FROM [Пользователь]")
    if q_:
        users = [u for u in users if any(q_ in (u[k] or "").casefold() for k in ("login", "fio", "email", "phone"))]
    return adm("users", users=sorted(users, key=lambda u: -u["id_user"])[:300], q=request.args.get("q", ""))


@app.route("/admin/users/<int:uid>")
@admin_required
def admin_user(uid):
    u = q1("SELECT * FROM [Пользователь] WHERE [id_user] = ?", [uid])
    if not u:
        abort(404)
    ops = sorted(q("SELECT [id_operation], [id_rent], [Количество] AS delta, [ТипОперации], [Дата], [Комментарий] "
                   "FROM [ОперацииБаллов] WHERE [id_user] = ?", [uid]), key=lambda o: -o["id_operation"])[:15]
    journal = sorted(q("SELECT * FROM [ИсторияДействийАдминистратора] WHERE [id_user_target] = ?", [uid]),
                     key=lambda j: -j["id_action"])[:15]
    all_coupons = q("SELECT * FROM [Купон] WHERE [Активен] = ?", [True])
    return adm("user", u=u, org=q1("SELECT * FROM [B2B_Реквизиты] WHERE [id_user] = ?", [uid]), stats=user_stats(uid),
               orders=orders_with_counts(uid)[:10], coupons=q_coupons(uid), ops=ops, journal=journal,
               coupon_list=all_coupons, adm_names=user_map(j["id_user_admin"] for j in journal))


@app.route("/admin/users/<int:uid>/<action>", methods=["POST"])
@admin_required
def admin_user_action(uid, action):
    me = g.user
    notes = []
    try:
        with Tx() as t:
            u = t.one("SELECT * FROM [Пользователь] WHERE [id_user] = ?", [uid])
            if not u:
                abort(404)
            if action == "block":
                if u["role"] == ROLE_ADMIN:
                    raise ValueError("Администратора заблокировать нельзя.")
                reason = field("reason", 255, True, "Причина блокировки")
                t.update("Пользователь", "id_user", uid, {"Активен": False, "ПричинаБлокировки": reason, "ДатаБлокировки": now()})
                admin_log(t, me["id_user"], "Пользователь заблокирован", uid, reason)
            elif action == "unblock":
                t.update("Пользователь", "id_user", uid, {"Активен": True, "ПричинаБлокировки": None, "ДатаБлокировки": None})
                admin_log(t, me["id_user"], "Пользователь разблокирован", uid, "")
            elif action == "confirm-email":
                if u["email_ok"]:
                    raise ValueError("Почта уже подтверждена.")
                t.update("Пользователь", "id_user", uid, {"EmailПодтвержден": True})
                admin_log(t, me["id_user"], "Почта подтверждена вручную", uid, f"Email: {u['email']}")
                notes.append((uid, "Почта подтверждена", "Администратор подтвердил ваш адрес электронной почты.", "Система", None, False))
            elif action == "points":
                try:
                    delta = int(request.form.get("delta", "0").replace(" ", ""))
                except ValueError:
                    raise ValueError("Укажите целое число баллов (отрицательное — списание).")
                comment = field("comment", 255, True, "Комментарий")
                if delta == 0 or abs(delta) > 1000000:
                    raise ValueError("Количество баллов должно быть ненулевым и не больше 1 000 000.")
                if delta < 0 and t.x("UPDATE [Пользователь] SET [Баллы] = [Баллы] + ? WHERE [id_user] = ? AND [Баллы] >= ?",
                                     [delta, uid, -delta]) != 1:
                    raise ValueError("Нельзя списать больше баллов, чем на балансе.")
                if delta > 0:
                    t.x("UPDATE [Пользователь] SET [Баллы] = [Баллы] + ? WHERE [id_user] = ?", [delta, uid])
                t.insert("ОперацииБаллов", {"id_user": uid, "id_rent": None, "Количество": delta,
                                            "ТипОперации": "Начисление администратором" if delta > 0 else "Списание администратором",
                                            "Дата": now(), "Комментарий": comment})
                admin_log(t, me["id_user"], "Баллы " + ("начислены" if delta > 0 else "списаны"), uid, f"{delta:+d}: {comment}")
                notes.append((uid, "Изменение баллов", f"{'Начислено' if delta > 0 else 'Списано'} баллов: {abs(delta)}. {comment}", "Баллы", None))
            elif action == "coupon":
                cid = to_int(request.form.get("coupon_id"))
                c = t.one("SELECT * FROM [Купон] WHERE [id_coupon] = ? AND [Активен] = ?", [cid, True])
                if not c:
                    raise ValueError("Выберите действующий купон.")
                end = parse_date(request.form.get("end")) if request.form.get("end") else None
                t.insert("КупонПользователя", {"id_user": uid, "id_coupon": cid, "ДатаВыдачи": now(), "ДатаОкончания": end, "Использован": False})
                admin_log(t, me["id_user"], "Купон выдан", uid, f"Купон №{cid} «{clean_text(c['name'])}»")
                notes.append((uid, "Вам выдан купон", f"Купон «{clean_text(c['name'])}» доступен в личном кабинете (раздел «Купоны»).", "Купон", None))
            elif action == "message":
                title, text = field("title", 255, True, "Тема"), field("text", 5000, True, "Сообщение")
                admin_log(t, me["id_user"], "Сообщение отправлено", uid, title)
                notes.append((uid, title, text, "Сообщение", None))
            else:
                abort(404)
        notify_all(notes)
        flash("Готово.")
    except ValueError as e:
        flash(str(e))
    return redirect(url_for("admin_user", uid=uid))


# --- купоны --------------------------------------------------------------
@app.route("/admin/coupons")
@admin_required
def admin_coupons():
    coupons = sorted(q("SELECT * FROM [Купон]"), key=lambda c: -c["id_coupon"])
    used, given = {}, {}
    for r in q("SELECT [id_coupon] FROM [ИспользованиеКупона]"):
        used[r["id_coupon"]] = used.get(r["id_coupon"], 0) + 1
    for r in q("SELECT [id_coupon] FROM [КупонПользователя]"):
        given[r["id_coupon"]] = given.get(r["id_coupon"], 0) + 1
    return adm("coupons", coupons=coupons, used=used, given=given)


@app.route("/admin/coupons/new")
@app.route("/admin/coupons/<int:cid>")
@admin_required
def admin_coupon_form(cid=0):
    c = q1("SELECT * FROM [Купон] WHERE [id_coupon] = ?", [cid]) if cid else None
    if cid and not c:
        abort(404)
    return adm("coupon_form", c=c)


@app.route("/admin/coupons/save", methods=["POST"])
@admin_required
def admin_coupon_save():
    cid = to_int(request.form.get("id"))
    try:
        kind = request.form.get("disc_type")
        if kind not in ("Процент", "Сумма"):
            raise ValueError("Выберите тип скидки.")
        amount = to_money(request.form.get("amount", ""), "Размер скидки")
        if amount <= 0 or (kind == "Процент" and amount > 100):
            raise ValueError("Размер скидки: процент от 1 до 100 или сумма больше нуля.")
        start, end = parse_date(request.form.get("start")), parse_date(request.form.get("end"))
        if start and end and end < start:
            raise ValueError("Дата окончания раньше даты начала.")
        vals = {"Название": field("name", 255, True, "Название"), "Описание": field("descr", 2000) or None, "ТипСкидки": kind,
                "РазмерСкидки": amount, "ДатаНачала": start, "ДатаОкончания": end,
                "МинимальнаяСумма": to_money(request.form.get("min_sum", "0") or "0", "Минимальная сумма"),
                "ЛимитИспользований": to_int(request.form.get("limit"), 0, 0, 1000000),
                "Активен": request.form.get("active") == "1"}
        with Tx() as t:
            if cid:
                t.update("Купон", "id_coupon", cid, vals)
            else:
                while True:       # поле «Код» в базе числовое: выдаём уникальный шестизначный номер
                    code = secrets.randbelow(900000) + 100000
                    if not t.one("SELECT [id_coupon] FROM [Купон] WHERE [Код] = ?", [code]):
                        break
                cid = t.insert("Купон", {**vals, "Код": code})
            admin_log(t, g.user["id_user"], "Купон сохранён", None, f"Купон №{cid} «{vals['Название']}»")
        flash("Купон сохранён.")
        return redirect(url_for("admin_coupons"))
    except ValueError as e:
        flash(str(e))
        return redirect(url_for("admin_coupon_form", cid=cid) if cid else url_for("admin_coupon_form"))


@app.route("/admin/coupons/<int:cid>/toggle", methods=["POST"])
@admin_required
def admin_coupon_toggle(cid):
    with Tx() as t:
        c = t.one("SELECT * FROM [Купон] WHERE [id_coupon] = ?", [cid])
        if not c:
            abort(404)
        t.update("Купон", "id_coupon", cid, {"Активен": not c["active"]})
        admin_log(t, g.user["id_user"], "Купон " + ("деактивирован" if c["active"] else "активирован"), None, f"Купон №{cid} «{clean_text(c['name'])}»")
    return redirect(url_for("admin_coupons"))


@app.route("/admin/coupons/<int:cid>/assign", methods=["POST"])
@admin_required
def admin_coupon_assign(cid):
    notes = []
    try:
        login_ = field("login", 32, True, "Логин")
        with Tx() as t:
            u = t.one("SELECT * FROM [Пользователь] WHERE [Логин] = ?", [login_])
            c = t.one("SELECT * FROM [Купон] WHERE [id_coupon] = ? AND [Активен] = ?", [cid, True])
            if not u or not c:
                raise ValueError("Пользователь или действующий купон не найден.")
            end = parse_date(request.form.get("end")) if request.form.get("end") else None
            t.insert("КупонПользователя", {"id_user": u["id_user"], "id_coupon": cid, "ДатаВыдачи": now(), "ДатаОкончания": end, "Использован": False})
            admin_log(t, g.user["id_user"], "Купон выдан", u["id_user"], f"Купон №{cid} «{clean_text(c['name'])}»")
            notes.append((u["id_user"], "Вам выдан купон", f"Купон «{clean_text(c['name'])}» доступен в личном кабинете (раздел «Купоны»).", "Купон", None))
        notify_all(notes)
        flash("Купон выдан.")
    except ValueError as e:
        flash(str(e))
    return redirect(url_for("admin_coupons"))


# --- журнал --------------------------------------------------------------
@app.route("/admin/journal")
@admin_required
def admin_journal():
    rows = sorted(q("SELECT * FROM [ИсторияДействийАдминистратора]"), key=lambda j: -j["id_action"])[:300]
    return adm("journal", rows=rows, um=user_map([r["id_user_admin"] for r in rows] + [r["id_user_target"] for r in rows]))


# ---------------------------------------------------------------------------
# Диагностика подключения (только с локального компьютера, содержимое таблиц не показывает)
# ---------------------------------------------------------------------------
DB_EXPECTED = {
    "Пользователь": ["id_user", "Логин", "Пароль", "ФИО", "Телефон", "Email", "Адрес", "Роль", "Баллы", "EmailПодтвержден",
                     "Активен", "ПричинаБлокировки", "ДатаБлокировки", "ДатаРегистрации"],
    "B2B_Реквизиты": ["id_b2b", "id_user", "НазваниеКомпании", "ИНН", "КПП", "ЮрАдрес", "ТипОрганизации"],
    "Инструмент": ["id_tool", "id_brand", "Название", "Описание", "Характеристики", "ЦенаЗаСутки", "Залог",
                   "КоличествоНаСкладе", "ИмяФото", "Активен", "id_category"],
    "Бренд": ["id_brand", "Название"], "Категория": ["id_category", "Название"],
    "Бронирование": ["id_rent", "id_user", "ДатаСоздания", "Статус", "СуммаАренды", "СуммаЗалога", "СпособДоставки",
                     "СуммаДоСкидки", "СуммаСкидки"],
    "ПозицияБронирования": ["id_item", "id_rent", "id_tool", "Количество", "ЦенаНаМоментБрони", "ДатаНачала", "ДатаОкончания", "Статус"],
    "БлокировкаИнструмента": ["id_block", "id_tool", "ДатаНачала", "ДатаОкончания", "Причина"],
    "ИсторияСтатусов": ["id_log", "id_rent", "СтарыйСтатус", "НовыйСтатус", "КтоИзменил_id_user", "Когда", "Комментарий"],
    "ИсторияСтатусовПозиций": ["id_item_log", "id_item", "СтарыйСтатус", "НовыйСтатус", "КтоИзменил_id_user", "Когда", "Комментарий"],
    "ОперацииБаллов": ["id_operation", "id_user", "id_rent", "Количество", "ТипОперации", "Дата", "Комментарий"],
    "Купон": ["id_coupon", "Код", "Название", "Описание", "ТипСкидки", "РазмерСкидки", "ДатаНачала", "ДатаОкончания",
              "МинимальнаяСумма", "ЛимитИспользований", "Активен"],
    "КупонПользователя": ["id_user_coupon", "id_user", "id_coupon", "ДатаВыдачи", "ДатаОкончания", "Использован"],
    "ИспользованиеКупона": ["id_usage", "id_coupon", "id_user", "id_rent", "РазмерСкидки", "ДатаИспользования"],
    "Уведомления": ["id_notification", "id_user", "Заголовок", "Текст", "Тип", "ДатаСоздания", "Прочитано", "id_rent"],
    "ТокеныАккаунта": ["id_token", "id_user", "ТокенХеш", "Тип", "ДатаСоздания", "СрокДействия", "Использован"],
    "ИсторияДействийАдминистратора": ["id_action", "id_user_admin", "id_user_target", "ТипДействия", "Дата", "Комментарий"],
}


@app.route("/db-check")
def db_check():
    if request.remote_addr not in ("127.0.0.1", "::1"):
        abort(404)
    L = ["ToolRent: диагностика Access (содержимое таблиц не показывается)",
         f"Python {sys.version.split()[0]}, {struct.calcsize('P') * 8}-bit, {platform.platform()}", f"Файл базы: {DB_PATH}"]
    if pyodbc is None:
        L.append("pyodbc: НЕ УСТАНОВЛЕН → pip install pyodbc")
        return Response("\n".join(L), mimetype="text/plain; charset=utf-8")
    drivers = pyodbc.drivers()
    L.append("Драйверы Access: " + (", ".join(d for d in drivers if "Access" in d) or "НЕ НАЙДЕНЫ"))
    try:
        conn = pyodbc.connect(CONN_STR)
    except Exception as e:
        L.append(f"\nПОДКЛЮЧЕНИЕ НЕ УДАЛОСЬ: {e}")
        return Response("\n".join(L), mimetype="text/plain; charset=utf-8")
    L.append("Подключение: OK\n")
    cur, bad = conn.cursor(), 0
    for table, expected in DB_EXPECTED.items():
        try:
            n = cur.execute(f"SELECT COUNT(*) FROM [{table}]").fetchone()[0]
            cols = {r.column_name: r.type_name for r in cur.columns(table=table)}
            miss, extra = [c for c in expected if c not in cols], [c for c in cols if c not in expected]
            bad += bool(miss)
            L.append(f"[{table}] строк: {n}" + (f"  !!! НЕТ колонок: {', '.join(miss)}" if miss else "  колонки совпадают")
                     + (f"  (лишние: {', '.join(extra)})" if extra else ""))
        except Exception as e:
            bad += 1
            L.append(f"[{table}] ОШИБКА: {e}")
    adm_n = cur.execute("SELECT COUNT(*) FROM [Пользователь] WHERE [Роль] = 'admin'").fetchone()[0]
    L.append(f"\nАдминистраторов: {adm_n} (должен быть ровно 1; назначить: python server.py make-admin ЛОГИН)")
    L.append("ИТОГ: структура соответствует ожидаемой." if not bad else f"ИТОГ: проблем в структуре: {bad}")
    conn.close()
    return Response("\n".join(L), mimetype="text/plain; charset=utf-8")


# ---------------------------------------------------------------------------
# Командная строка и запуск
# ---------------------------------------------------------------------------
def make_admin(login_):
    with Tx() as t:
        existing = t.q("SELECT * FROM [Пользователь] WHERE [Роль] = ?", [ROLE_ADMIN])
        u = t.one("SELECT * FROM [Пользователь] WHERE [Логин] = ?", [login_])
        if not u:
            return f"Пользователь «{login_}» не найден: сначала зарегистрируйте его на сайте."
        if existing and existing[0]["id_user"] != u["id_user"]:
            return (f"Администратор уже назначен: «{existing[0]['login']}». В системе может быть только один. "
                    "Чтобы передать права, измените роль вручную в таблице «Пользователь».")
        t.update("Пользователь", "id_user", u["id_user"], {"Роль": ROLE_ADMIN, "EmailПодтвержден": True, "Активен": True})
    return f"Готово: «{login_}» — администратор."


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "make-admin":
        print(make_admin(sys.argv[2]))
        sys.exit(0)
    if os.environ.get("WERKZEUG_RUN_MAIN") == "true" or not app.debug:
        threading.Thread(target=_due_loop, daemon=True).start()
    app.run(debug=True)
