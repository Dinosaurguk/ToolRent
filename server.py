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
import json
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

from flask import (Flask, Response, abort, flash, g, has_request_context, jsonify, redirect, render_template,
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
POINTS_MAX_PERCENT = Decimal("30")       # баллами можно оплатить не более 30% аренды после скидок (залог не оплачивается)
try:                                     # 1 балл = 1 рубль; для другого курса задайте TOOLRENT_POINT_RUB_RATE
    POINT_RUB_RATE = Decimal(_env("TOOLRENT_POINT_RUB_RATE", "1"))
except InvalidOperation:
    POINT_RUB_RATE = Decimal("1")
if POINT_RUB_RATE <= 0:
    POINT_RUB_RATE = Decimal("1")

# --- Правила скидок: единое место настройки, браузер получает эти же значения из Flask ---
QUANTITY_DISCOUNT_MIN_UNITS = 15         # скидка за количество: суммарно физических единиц во всей заявке
QUANTITY_DISCOUNT_PERCENT = Decimal("3") # ...на всю стоимость аренды (залог не участвует)
DURATION_TIERS = []                      # скидка за срок, ступени (от_суток, процент) для каждой позиции отдельно.
#   Проценты не согласованы, поэтому ступеней пока нет. Пример: [(7, Decimal("5")), (14, Decimal("8")), (30, Decimal("12"))]
CORPORATE_DISCOUNT_PERCENT = Decimal("3")  # скидка для бизнеса: на всю аренду (залог не скидывается), только для подтверждённых бизнес-клиентов в режиме «Бизнесу»
BIZ_MARK = "бизнес-скидка"                 # пометка в истории заявки: скидка была применена при оформлении
COUPON_PERCENT, COUPON_FIXED = "Процент", "Фиксированная"   # типы купонов в базе

DURATIONS = [1, 3, 7, 14, 30]
MAX_RENT_DAYS = 365
MAX_CART_LINES = 30
TOKEN_HOURS = {"email": 24, "reset": 1}
MAIL_RETRY_SECONDS = (60, 300, 1800)     # повторные попытки отправки письма
ORDER_LOCK = threading.RLock()           # одновременные заявки на один инструмент обрабатываются по очереди

ROLE_USER, ROLE_B2B, ROLE_ADMIN = "user", "b2b", "admin"

# Статусы позиций заявки и заявки целиком (хранятся в базе как текст)
P_NEW, P_OK, P_NO, P_OUT, P_BACK, P_CANCEL = (
    "На рассмотрении", "Одобрена", "Отклонена", "Выдана", "Возвращена", "Отменена")
O_NEW, O_WORK, O_RENT, O_DONE, O_NO, O_CANCEL = (
    "На рассмотрении", "В обработке", "В аренде", "Завершена", "Отклонена", "Отменена")
# Физические экземпляры: состояние экземпляра и статус его назначения на позицию заявки
STATE_OK, STATE_CHECK, STATE_REPAIR, STATE_OFF = "Доступен", "На проверке", "В ремонте", "Списан"
INSTANCE_STATES = (STATE_OK, STATE_CHECK, STATE_REPAIR, STATE_OFF)
A_ASSIGNED, A_OUT, A_BACK, A_CANCEL = "Назначен", "Выдан", "Возвращён", "Отменён"
ARTICLE_RE = re.compile(r"^([A-Z]{3})-(\d{6})$")
RESERVING = (P_NEW, P_OK, P_OUT)         # статусы, при которых инструмент занят (заявка, отправленная на рассмотрение, уже резервирует)
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
    "Артикул": "article", "Состояние": "state", "ИнвентарныйКод": "code_inv", "ДатаДобавления": "added", "Примечание": "note",
    "КодВыдачи": "issue_code", "ОценкаКачества": "grade", "ДатаВозврата": "returned_at", "КтоОсмотрел_id_user": "inspector",
    "Событие": "event", "СостояниеДо": "state_before", "СостояниеПосле": "state_after", "ИнвентарныйКодНаМомент": "code_at",
    "ДатаСобытия": "event_at",
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
    s = f"{int(v):,}".replace(",", "\u00a0") if v == v.to_integral() else f"{v:,.2f}".replace(",", "\u00a0").replace(".", ",")
    return s + "\u00a0₽"


def parse_specs(text):
    """Характеристики из Access → список (название, значение). Понимает «Название: значение» (по строке),
    «Название;значение», а также чередование строк «название / значение» без двоеточий.
    Строка без названия возвращается как ("", текст) и выводится на всю ширину."""
    raw = clean_text(text)
    if not raw.strip():
        return []
    if "\n" not in raw and raw.count(";") >= 1 and ":" in raw:
        raw = raw.replace(";", "\n")
    lines = [re.sub(r"^[•*]\s*|^[-–—]\s+", "", l.strip()) for l in raw.split("\n")]
    lines = [l for l in lines if l]
    rows = []
    if any(":" in l for l in lines):
        for l in lines:
            key, sep, val = l.partition(":")
            key, val = key.strip(), val.strip()
            if sep and key and len(key) <= 70:
                rows.append([key, val])
            elif rows and rows[-1][1] == "" and not sep:      # «Название:» и значение на следующей строке
                rows[-1][1] = l
            else:
                rows.append(["", l])
    elif any("\t" in l for l in lines):
        for l in lines:
            key, _, val = l.partition("\t")
            rows.append([key.strip(), val.strip()] if val.strip() else ["", l])
    elif len(lines) >= 2 and len(lines) % 2 == 0:             # название и значение на соседних строках
        rows = [[lines[i], lines[i + 1]] for i in range(0, len(lines), 2)]
    else:
        rows = [["", l] for l in lines]
    return [(k, v or "—") if k else ("", v) for k, v in rows]


def join_specs(keys, vals):
    """Строки формы редактора → текст для Access: «Название: значение», по строке на характеристику."""
    out = []
    for k, v in zip(keys, vals):
        k = re.sub(r"[:\r\n]+", " ", k or "").strip()
        v = re.sub(r"\s*[\r\n]+\s*", " ", v or "").strip()
        if k and v:
            out.append(f"{k}: {v}")
        elif v:
            out.append(v)
    return "\n".join(out)


def current_mode():
    """Режим просмотра: частный или бизнес. У подтверждённого бизнес-клиента по умолчанию бизнес."""
    m = session.get("mode")
    if m in ("b2c", "b2b"):
        return m
    return "b2b" if g.get("user") and g.user["role"] == ROLE_B2B else "b2c"


def is_biz_real():
    """Скидка для бизнеса действует только у подтверждённого бизнес-клиента в режиме «Бизнесу»."""
    return bool(g.get("user") and g.user["role"] == ROLE_B2B and current_mode() == "b2b")


def corp_bp_now():
    return _bp(CORPORATE_DISCOUNT_PERCENT) if has_request_context() and is_biz_real() else 0


@app.template_filter("price_html")
def price_html(value):
    """В режиме «Бизнесу» старая цена зачёркнута, рядом цена со скидкой и метка скидки."""
    if not has_request_context() or current_mode() != "b2b":
        return rub(value)
    k, bp = kop(value), _bp(CORPORATE_DISCOUNT_PERCENT)
    new = from_k(k - _off(k, bp))
    return Markup(f'<span class="price-old">{escape(rub(value))}</span><span class="price-new">{escape(rub(new))}</span>'
                  f'<span class="disc-chip">−{escape(fmt_pct(bp))}</span>')


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


app.jinja_env.globals["csrf_token"] = csrf_token
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
    (("маляр",), "paint.svg"), (("штукатур",), "plaster.svg"),
    (("газов", "свароч"), "welding.svg"), (("специал",), "special.svg"),
    (("электроинстр",), "powertool.svg"),
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


def reserved_map(start, end, t=None, exclude_rent=None, exclude_item=None):
    """{id_tool: занято штук}: позиции «На рассмотрении», «Одобрена», «Выдана», пересекающиеся с периодом.
    Каждая позиция считается один раз при любом из этих статусов, поэтому одобрение не резервирует повторно."""
    rows = (t.q if t else q)(
        "SELECT [id_item], [id_tool], [id_rent], [Количество], [ДатаНачала], [ДатаОкончания] FROM [ПозицияБронирования] "
        "WHERE [Статус] IN (?, ?, ?) AND [ДатаНачала] <= ? AND [ДатаОкончания] >= ?", [*RESERVING, end, start])
    out = {}
    for r in rows:
        if (exclude_rent and r["id_rent"] == exclude_rent) or (exclude_item and r["id_item"] == exclude_item):
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


def tool_free(t, tool_id, stock, start, end, exclude_rent=None, exclude_item=None):
    if tool_id in blocked_tools(start, end, t):
        return 0
    return max(int(stock) - reserved_map(start, end, t, exclude_rent, exclude_item).get(tool_id, 0), 0)


def capacity_map(t=None):
    """{id_tool: (доступных к аренде, всего не списанных)} по таблице ФизическийЭкземпляр."""
    out = {}
    for r in (t.q if t else q)("SELECT [id_tool], [Состояние] FROM [ФизическийЭкземпляр]"):
        a, o = out.get(r["id_tool"], (0, 0))
        if r["state"] != STATE_OFF:
            o += 1
        if r["state"] == STATE_OK:
            a += 1
        out[r["id_tool"]] = (a, o)
    return out


def _caps():
    if has_request_context():
        if "caps" not in g:
            g.caps = capacity_map()
        return g.caps
    return capacity_map()


def tool_capacity(t, row):
    """Сколько экземпляров модели можно сдавать (на проверке, в ремонте и списанные не входят).
    Если у модели ещё нет экземпляров, используется КоличествоНаСкладе."""
    a = capacity_map(t).get(row["id_tool"])
    return int(row["stock"] or 0) if a is None else a[0]


def build_tool(r, reserved=None, blocked=None):
    cap = _caps().get(r["id_tool"])
    stock = int(r["stock"] or 0) if cap is None else cap[0]
    free = stock if reserved is None else max(stock - reserved.get(r["id_tool"], 0), 0)
    if blocked and r["id_tool"] in blocked:
        free = 0
    return {"id": r["id_tool"], "brand": get_brands().get(r["id_brand"], ""), "brand_id": r["id_brand"],
            "name": clean_text(r["name"]), "cat": r["id_category"], "price": money(r["price"]),
            "deposit": money(r["deposit"]), "stock": stock, "owned": stock if cap is None else cap[1], "qty": free,
            "article": clean_text(r["article"]),
            "img": tool_img(r["photo"], r["id_category"]), "photo": clean_text(r["photo"]),
            "desc": clean_text(r["descr"]), "specs": clean_text(r["specs"]), "specs_rows": parse_specs(r["specs"]),
            "active": bool(r["active"])}


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
# Расчёт стоимости аренды. Все суммы считаются в целых копейках, округление «от половины вверх».
# Тот же алгоритм реализован в script.js: браузер показывает результат мгновенно, а окончательные
# суммы при добавлении в корзину и при оформлении заявки всегда считает Flask.
# ---------------------------------------------------------------------------
def kop(x):
    return int((D(x) * 100).to_integral_value(ROUND_HALF_UP))


def from_k(k):
    return Decimal(int(k)) / 100


def _bp(percent):
    """Процент → сотые доли процента (3 % → 300)."""
    return int((D(percent) * 100).to_integral_value(ROUND_HALF_UP))


def _off(amount_k, bp):
    """Размер скидки bp (сотых процента) от суммы в копейках."""
    return (amount_k * bp + 5000) // 10000


def duration_bp(days):
    bp = 0
    for d, p in sorted(DURATION_TIERS):
        if days >= d:
            bp = _bp(p)
    return bp


def rent_config(preview=False):
    """Правила расчёта для браузера (единый источник — настройки в начале server.py)."""
    return {"qty_min_units": QUANTITY_DISCOUNT_MIN_UNITS, "qty_bp": _bp(QUANTITY_DISCOUNT_PERCENT),
            "tiers": [[d, _bp(p)] for d, p in sorted(DURATION_TIERS)], "corp_bp": _bp(CORPORATE_DISCOUNT_PERCENT) if preview else corp_bp_now(),
            "points_max_bp": _bp(POINTS_MAX_PERCENT), "rate_k": kop(POINT_RUB_RATE), "max_days": MAX_RENT_DAYS,
            "accrual_percent": int(POINTS_PERCENT)}


def coupon_kind(c):
    return "percent" if c["disc_type"] == COUPON_PERCENT else "fixed"      # прежнее значение «Сумма» читается как фиксированная


def coupon_spec(c):
    return {"type": coupon_kind(c), "bp": _bp(c["amount"]), "amount_k": kop(c["amount"]), "min_k": kop(c["min_sum"])}


def calc_quote(lines, coupon=None, points=0, balance=None, enforce_max=True, corp_bp=None):
    """lines: [{price_k, deposit_k, qty, days}] → скидки и итоги в копейках.
    Порядок: срок (по позиции) → количество (на всю заявку) → корпоративная → купон → баллы. Залог не скидывается."""
    units = sum(l["qty"] for l in lines)
    qbp = _bp(QUANTITY_DISCOUNT_PERCENT) if units >= QUANTITY_DISCOUNT_MIN_UNITS else 0
    cbp = corp_bp_now() if corp_bp is None else corp_bp
    out, tot = [], dict(base=0, dur=0, qty=0, corp=0, rent=0, deposit=0)
    for l in lines:
        base = l["price_k"] * l["qty"] * l["days"]
        dbp = duration_bp(l["days"])
        dur = _off(base, dbp)
        a = base - dur
        qd = _off(a, qbp)
        b = a - qd
        cp = _off(b, cbp)
        rent = b - cp
        dep = l["deposit_k"] * l["qty"]
        out.append({"base": base, "dur_bp": dbp, "dur": dur, "qty_disc": qd, "corp": cp, "rent": rent, "deposit": dep})
        for k, v in (("base", base), ("dur", dur), ("qty", qd), ("corp", cp), ("rent", rent), ("deposit", dep)):
            tot[k] += v
    auto_rent = tot["rent"]
    coupon_k, coupon_ok, coupon_why = 0, True, ""
    if coupon:
        if auto_rent < coupon.get("min_k", 0):
            coupon_ok, coupon_why = False, "сумма аренды меньше минимальной для купона"
        elif coupon["type"] == "percent":
            coupon_k = _off(auto_rent, coupon["bp"])
        else:
            coupon_k = min(coupon["amount_k"], auto_rent)
    after_coupon = auto_rent - coupon_k
    rate_k = kop(POINT_RUB_RATE)
    max_points = (after_coupon * _bp(POINTS_MAX_PERCENT) // 10000) // rate_k       # не более 30 % аренды после скидок
    if enforce_max:
        cap = min(max_points, balance) if balance is not None else max_points
    else:
        cap = min(after_coupon // rate_k, balance) if balance is not None else after_coupon // rate_k
    used = max(0, min(int(points or 0), cap))
    points_k = used * rate_k
    rent_total = after_coupon - points_k
    return {"lines": out, "units": units, "qty_bp": qbp, "corp_bp": cbp, "base": tot["base"], "dur": tot["dur"],
            "qty": tot["qty"], "corp": tot["corp"], "auto_rent": auto_rent, "coupon": coupon_k, "coupon_ok": coupon_ok,
            "coupon_why": coupon_why, "max_points": cap, "points": used, "points_k": points_k, "rent": rent_total,
            "deposit": tot["deposit"], "grand": rent_total + tot["deposit"], "savings": tot["base"] - rent_total}


@app.template_filter("rubk")
def rub_k(k):
    return rub(from_k(k or 0))


@app.template_filter("coupon_value")
def coupon_value(c):
    """«10%» или «500 ₽» для купона любого происхождения (Процент / Фиксированная / прежняя «Сумма»)."""
    return fmt_pct(_bp(c["amount"])) if c["disc_type"] == COUPON_PERCENT else rub(c["amount"])


@app.template_filter("pct")
def fmt_pct(bp):
    v = Decimal(int(bp or 0)) / 100
    txt = f"{v:f}"
    if "." in txt:
        txt = txt.rstrip("0").rstrip(".")
    return txt.replace(".", ",") + "%"


# ---------------------------------------------------------------------------
# Доступность на период (учитывает заявки на рассмотрении, одобренные и выданные позиции, блокировки)
# ---------------------------------------------------------------------------
def _tool_rows(ids, t=None):
    ids = list(ids)
    if not ids:
        return {}
    rows = (t.q if t else q)(f"SELECT * FROM [Инструмент] WHERE [id_tool] IN ({', '.join(['?'] * len(ids))})", ids)
    return {r["id_tool"]: r for r in rows}


def lines_availability(items, t=None, rows=None):
    """items: [{tool_id, qty, start, end}] → для каждой позиции {free, ok, message}.
    Из свободного остатка вычитаются другие позиции этой же заявки на пересекающиеся даты."""
    rows = rows or _tool_rows({i["tool_id"] for i in items}, t)
    res = []
    for idx, it in enumerate(items):
        r = rows.get(it["tool_id"])
        if not r or not r["active"]:
            res.append({"free": 0, "stock": 0, "ok": False, "message": "Инструмент недоступен для аренды."})
            continue
        cap = tool_capacity(t, r)
        db_free = tool_free(t, it["tool_id"], cap, it["start"], it["end"])
        others = sum(o["qty"] for j, o in enumerate(items) if j != idx and o["tool_id"] == it["tool_id"]
                     and o["start"] <= it["end"] and o["end"] >= it["start"])
        free = max(db_free - others, 0)
        ok = it["qty"] <= free
        if ok:
            msg = ""
        elif free == 0:
            msg = "На выбранный период свободных экземпляров нет."
        else:
            msg = f"Доступно только {free} шт. на выбранный период" + (" с учётом других позиций заявки." if others else ".")
        res.append({"free": free, "stock": cap, "ok": ok, "message": msg})
    return res


# ---------------------------------------------------------------------------
# Физические экземпляры: серии артикулов, инвентарные коды, назначение, выдача, приёмка, списание.
# Резервирование остаётся по количеству (позиции «На рассмотрении», «Одобрена», «Выдана»); доступный склад модели —
# число экземпляров в состоянии «Доступен». Количество при выдаче и возврате не списывается второй раз:
# КоличествоНаСкладе уменьшается только при списании экземпляра.
# ---------------------------------------------------------------------------
def next_series(t):
    """Следующая серия артикулов: AAA → AAB → … → AAZ → ABA …"""
    used = []
    for r in t.q("SELECT [Артикул] FROM [Инструмент]"):
        m = ARTICLE_RE.match((r["article"] or "").strip().upper())
        if m:
            used.append(m.group(1))
    if not used:
        return "AAA"
    chars = list(max(used))
    for i in (2, 1, 0):
        if chars[i] != "Z":
            chars[i] = chr(ord(chars[i]) + 1)
            return "".join(chars)
        chars[i] = "A"
    raise ValueError("Серии артикулов закончились.")


def tool_series(t, tool):
    """Серия модели; пустой артикул заполняется автоматически, посторонний формат не трогаем (возвращает None)."""
    art = (tool["article"] or "").strip().upper()
    if not art:
        art = f"{next_series(t)}-000000"
        t.update("Инструмент", "id_tool", tool["id_tool"], {"Артикул": art})
    m = ARTICLE_RE.match(art)
    return m.group(1) if m else None


def instance_log(t, inst_id, event, before, after, assign_id=None, code=None, grade=None, comment=None, who=None):
    t.insert("ЖурналЭкземпляров", {
        "id_instance": inst_id, "id_instance_rent": assign_id, "ИнвентарныйКодНаМомент": code, "Событие": event,
        "СостояниеДо": before, "СостояниеПосле": after, "ОценкаКачества": grade,
        "Комментарий": (comment or "")[:255] or None, "ДатаСобытия": now(), "КтоИзменил_id_user": who})


def new_instance(t, tool_id, series, who=None):
    """Новый экземпляр с наименьшим свободным кодом серии (освобождённые коды списанных используются повторно)."""
    used = {i["code_inv"] for i in t.q("SELECT [ИнвентарныйКод] FROM [ФизическийЭкземпляр] WHERE [id_tool] = ?", [tool_id]) if i["code_inv"]}
    n = 1
    while f"{series}-{n:06d}" in used:
        n += 1
    code = f"{series}-{n:06d}"
    iid = t.insert("ФизическийЭкземпляр", {"id_tool": tool_id, "ИнвентарныйКод": code, "Состояние": STATE_OK,
                                          "ДатаДобавления": now(), "Примечание": None})
    instance_log(t, iid, "Добавлен", None, STATE_OK, code=code, who=who)
    return iid


def sync_alive_stock(t, tool_id):
    n = sum(1 for i in t.q("SELECT [Состояние] FROM [ФизическийЭкземпляр] WHERE [id_tool] = ?", [tool_id]) if i["state"] != STATE_OFF)
    t.update("Инструмент", "id_tool", tool_id, {"КоличествоНаСкладе": n})


def sync_tool(t, tool_id, who=None):
    """Назначает серию, создаёт недостающие экземпляры по КоличествоНаСкладе (повторный запуск дубликатов не создаёт)."""
    tool = t.one("SELECT * FROM [Инструмент] WHERE [id_tool] = ?", [tool_id])
    if not tool:
        return 0
    series = tool_series(t, tool)
    if not series:
        print(f"[ToolRent] Артикул «{tool['article']}» модели №{tool_id} не в формате AAA-000000: экземпляры не созданы.")
        return 0
    alive = [i for i in t.q("SELECT * FROM [ФизическийЭкземпляр] WHERE [id_tool] = ?", [tool_id]) if i["state"] != STATE_OFF]
    want, created = int(tool["stock"] or 0), 0
    for _ in range(max(want - len(alive), 0)):
        new_instance(t, tool_id, series, who)
        created += 1
    if want < len(alive):                           # уменьшить количество можно только списанием экземпляров
        t.update("Инструмент", "id_tool", tool_id, {"КоличествоНаСкладе": len(alive)})
    return created


def startup_sync():
    try:
        with ORDER_LOCK, Tx() as t:
            n = sum(sync_tool(t, r["id_tool"]) for r in t.q("SELECT [id_tool] FROM [Инструмент]"))
        if n:
            print(f"[ToolRent] Созданы физические экземпляры: {n}.")
    except Exception as e:
        print("[ToolRent] Синхронизация физических экземпляров не выполнена:", e)


def has_instances(t, tool_id):
    return bool(t.one("SELECT [id_instance] FROM [ФизическийЭкземпляр] WHERE [id_tool] = ?", [tool_id]))


def pick_instances(t, tool_id, start, end, qty, exclude_item=None, exclude_ids=()):
    """Свободные на период экземпляры в состоянии «Доступен». Экземпляр не назначается на две пересекающиеся аренды;
    выданный и не возвращённый считается занятым (в том числе после срока). Берутся те, что лучше «стыкуются» по датам."""
    insts = [i for i in t.q("SELECT * FROM [ФизическийЭкземпляр] WHERE [id_tool] = ? AND [Состояние] = ?", [tool_id, STATE_OK])
             if i["id_instance"] not in exclude_ids]
    ids = {i["id_instance"] for i in insts}
    items = {p["id_item"]: p for p in t.q("SELECT * FROM [ПозицияБронирования] WHERE [id_tool] = ? AND [Статус] IN (?, ?)", [tool_id, P_OK, P_OUT])}
    busy, day = {}, today()
    for a in t.q("SELECT * FROM [ЭкземплярыПозиции] WHERE [Статус] IN (?, ?)", [A_ASSIGNED, A_OUT]):
        p = items.get(a["id_item"])
        if a["id_instance"] not in ids or not p or p["id_item"] == exclude_item:
            continue
        s_, e_ = as_dt(p["start"]), as_dt(p["end"])
        busy.setdefault(a["id_instance"], []).append((s_, max(e_, day) if a["status"] == A_OUT else e_))
    free = []
    for i in insts:
        iv = busy.get(i["id_instance"], [])
        if any(s_ <= end and e_ >= start for s_, e_ in iv):
            continue
        before = [e_ for s_, e_ in iv if e_ < start]
        free.append(((start - max(before)).days if before else 10 ** 6, i["id_instance"], i))
    free.sort(key=lambda x: (x[0], x[1]))
    return [x[2] for x in free[:qty]]


def assign_instances(t, item, who):
    """Связывает позицию с конкретными экземплярами (ЭкземплярыПозиции.id_instance); недостающие добирает."""
    if not has_instances(t, item["id_tool"]):
        return
    have = [a for a in t.q("SELECT * FROM [ЭкземплярыПозиции] WHERE [id_item] = ?", [item["id_item"]]) if a["status"] in (A_ASSIGNED, A_OUT)]
    need = int(item["qty"]) - len(have)
    if need <= 0:
        return
    chosen = pick_instances(t, item["id_tool"], as_dt(item["start"]), as_dt(item["end"]), need, exclude_item=item["id_item"])
    if len(chosen) < need:
        raise ValueError(f"Не хватает физических экземпляров на эти даты: нужно {need}, свободно {len(chosen)} "
                         "(часть может быть на проверке, в ремонте или ещё не возвращена).")
    for inst in chosen:
        aid = t.insert("ЭкземплярыПозиции", {"id_item": item["id_item"], "id_instance": inst["id_instance"], "Статус": A_ASSIGNED})
        instance_log(t, inst["id_instance"], "Назначен на заявку", inst["state"], inst["state"], aid, inst["code_inv"], who=who)


def swap_assignment(t, a, item, bad_inst, who, reason):
    """Заменяет экземпляр в назначении другим свободным на те же даты."""
    repl = pick_instances(t, item["id_tool"], as_dt(item["start"]), as_dt(item["end"]), 1,
                          exclude_item=item["id_item"], exclude_ids={bad_inst["id_instance"]})
    if not repl:
        raise ValueError(f"Экземпляр {bad_inst['code_inv'] or '№' + str(bad_inst['id_instance'])} назначен на заявку №{item['id_rent']} "
                         f"({fmt_d(item['start'])}–{fmt_d(item['end'])}), заменить его нечем.")
    t.update("ЭкземплярыПозиции", "id_instance_rent", a["id_instance_rent"], {"id_instance": repl[0]["id_instance"]})
    instance_log(t, bad_inst["id_instance"], "Назначение снято", bad_inst["state"], bad_inst["state"], a["id_instance_rent"],
                 bad_inst["code_inv"], comment=reason, who=who)
    instance_log(t, repl[0]["id_instance"], "Назначен на заявку", STATE_OK, STATE_OK, a["id_instance_rent"], repl[0]["code_inv"],
                 comment="Замена", who=who)


def issue_instances(t, item, who):
    """Выдача: фиксирует код экземпляра (КодВыдачи) и дату; непригодные экземпляры заменяет."""
    if not has_instances(t, item["id_tool"]):
        return
    assign_instances(t, item, who)
    out_now = {a["id_instance"] for a in t.q("SELECT * FROM [ЭкземплярыПозиции] WHERE [Статус] = ?", [A_OUT])}
    for a in [x for x in t.q("SELECT * FROM [ЭкземплярыПозиции] WHERE [id_item] = ?", [item["id_item"]]) if x["status"] == A_ASSIGNED]:
        inst = t.one("SELECT * FROM [ФизическийЭкземпляр] WHERE [id_instance] = ?", [a["id_instance"]])
        if inst["state"] != STATE_OK or inst["id_instance"] in out_now:
            swap_assignment(t, a, item, inst, who, "Экземпляр недоступен к выдаче")
            a = t.one("SELECT * FROM [ЭкземплярыПозиции] WHERE [id_instance_rent] = ?", [a["id_instance_rent"]])
            inst = t.one("SELECT * FROM [ФизическийЭкземпляр] WHERE [id_instance] = ?", [a["id_instance"]])
        t.update("ЭкземплярыПозиции", "id_instance_rent", a["id_instance_rent"],
                 {"Статус": A_OUT, "КодВыдачи": inst["code_inv"], "ДатаВыдачи": now()})
        instance_log(t, inst["id_instance"], "Выдан", inst["state"], inst["state"], a["id_instance_rent"], inst["code_inv"], who=who)


def release_future_assignments(t, inst, who):
    """Перед уходом экземпляра с проверки/в ремонт/в списание подбирает замену для будущих назначений."""
    for a in t.q("SELECT * FROM [ЭкземплярыПозиции] WHERE [id_instance] = ? AND [Статус] = ?", [inst["id_instance"], A_ASSIGNED]):
        item = t.one("SELECT * FROM [ПозицияБронирования] WHERE [id_item] = ?", [a["id_item"]])
        if item:
            swap_assignment(t, a, item, inst, who, "Экземпляр выведен из доступных")


def apply_state(t, inst, new, who, comment=None, assign_id=None, grade=None, event=None):
    old = inst["state"]
    if old == STATE_OFF:
        raise ValueError("Списанный экземпляр изменить нельзя.")
    if new != STATE_OK:
        release_future_assignments(t, inst, who)
    code, vals = inst["code_inv"], {"Состояние": new}
    if new == STATE_OFF:        # код освобождается; прежний сохраняется в журнале, назначениях и примечании
        add = f"Списан {fmt_d(now())}; прежний инвентарный код {code}." + (f" {comment}" if comment else "")
        vals.update({"ИнвентарныйКод": None, "Примечание": ((inst["note"] + "\n") if inst["note"] else "") + add})
    t.update("ФизическийЭкземпляр", "id_instance", inst["id_instance"], vals)
    names = {STATE_OK: "Принят: доступен", STATE_CHECK: "Отправлен на проверку", STATE_REPAIR: "Отправлен в ремонт", STATE_OFF: "Списан"}
    instance_log(t, inst["id_instance"], event or names[new], old, new, assign_id, code, grade, comment, who)
    if new == STATE_OFF:
        sync_alive_stock(t, inst["id_tool"])


def return_instances(t, item, who, form, accept_all):
    """Возврат: осмотр каждого экземпляра (оценка 0–100, комментарий, решение) или одна кнопка «принять все»."""
    asg = [a for a in t.q("SELECT * FROM [ЭкземплярыПозиции] WHERE [id_item] = ?", [item["id_item"]]) if a["status"] == A_OUT]
    plan = []
    for a in asg:                                  # сначала проверка всех данных, потом запись
        k = a["id_instance_rent"]
        if accept_all:
            plan.append((a, 100, None, STATE_OK))
            continue
        try:
            grade = int(str(form.get(f"grade-{k}", "")).strip())
        except ValueError:
            grade = -1
        decision, comment = form.get(f"state-{k}"), (form.get(f"comment-{k}") or "").strip()[:255] or None
        if not 0 <= grade <= 100:
            raise ValueError("Оценка качества — число от 0 до 100.")
        if decision not in INSTANCE_STATES:
            raise ValueError("Выберите дальнейшее состояние каждого экземпляра.")
        if decision in (STATE_REPAIR, STATE_OFF) and not comment:
            raise ValueError("Для ремонта и списания укажите комментарий.")
        plan.append((a, grade, comment, decision))
    for a, grade, comment, decision in plan:
        k = a["id_instance_rent"]
        inst = t.one("SELECT * FROM [ФизическийЭкземпляр] WHERE [id_instance] = ?", [a["id_instance"]])
        t.update("ЭкземплярыПозиции", "id_instance_rent", k, {"Статус": A_BACK, "ДатаВозврата": now(), "ОценкаКачества": grade,
                                                              "Комментарий": comment, "КтоОсмотрел_id_user": who})
        apply_state(t, inst, decision, who, comment, k, grade, "Возврат и приёмка")


def cancel_assignments(t, item_id, who):
    for a in t.q("SELECT * FROM [ЭкземплярыПозиции] WHERE [id_item] = ?", [item_id]):
        if a["status"] == A_ASSIGNED:
            inst = t.one("SELECT * FROM [ФизическийЭкземпляр] WHERE [id_instance] = ?", [a["id_instance"]])
            t.update("ЭкземплярыПозиции", "id_instance_rent", a["id_instance_rent"], {"Статус": A_CANCEL})
            instance_log(t, inst["id_instance"], "Назначение снято", inst["state"], inst["state"], a["id_instance_rent"],
                         inst["code_inv"], comment="Заявка отменена", who=who)


def change_instance_state(admin, iid, new, comment):
    """Ручная смена состояния: проверка, ремонт, приёмка, списание. Возвращает id модели."""
    if new not in INSTANCE_STATES:
        raise ValueError("Выберите состояние.")
    with ORDER_LOCK, Tx() as t:
        inst = t.one("SELECT * FROM [ФизическийЭкземпляр] WHERE [id_instance] = ?", [iid])
        if not inst:
            raise ValueError("Экземпляр не найден.")
        if inst["state"] == STATE_OFF:
            raise ValueError("Списанный экземпляр изменить нельзя.")
        if any(a["status"] == A_OUT for a in t.q("SELECT [Статус] FROM [ЭкземплярыПозиции] WHERE [id_instance] = ?", [iid])):
            raise ValueError("Экземпляр сейчас выдан: сначала оформите возврат и приёмку по заявке.")
        if new == inst["state"]:
            raise ValueError("Экземпляр уже в этом состоянии.")
        if new in (STATE_REPAIR, STATE_OFF) and not comment:
            raise ValueError("Для ремонта и списания укажите комментарий.")
        code = inst["code_inv"]
        repaired = inst["state"] == STATE_REPAIR and new == STATE_CHECK      # ремонт завершён, нужна приёмка
        event = ("Ремонт завершён: ожидает подтверждения приёмки" if repaired else
                 "Приёмка подтверждена" if inst["state"] == STATE_CHECK and new == STATE_OK else None)
        apply_state(t, inst, new, admin["id_user"], comment or None, event=event)
        admin_log(t, admin["id_user"], "Экземпляр: " + new.lower(), None,
                  f"Модель №{inst['id_tool']}, экземпляр №{iid} ({code}): {inst['state']} → {new}" + (f". {comment}" if comment else ""))
        notes = []
        if repaired:                       # уведомление администраторам: инструмент починен, подтвердите приёмку
            tool = t.one("SELECT * FROM [Инструмент] WHERE [id_tool] = ?", [inst["id_tool"]])
            for aid in admin_ids(t):
                notes.append((aid, "Инструмент починен: подтвердите приёмку",
                              f"«{clean_text(tool['name'])}», экземпляр {code}: ремонт завершён" + (f" ({comment})" if comment else "") +
                              ". Откройте Админ-панель → Инструменты → Экземпляры и подтвердите приёмку, выбрав состояние «Доступен».",
                              "Система", None))
        result = (inst["id_tool"], repaired)
    notify_all(notes)
    return result


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
                    "price_k": kop(tool["price"]), "deposit_k": kop(tool["deposit"])})
    return out


def calc_input(lines):
    return [{"price_k": l["price_k"], "deposit_k": l["deposit_k"], "qty": l["qty"], "days": l["days"]} for l in lines]


def validate_period(start, end):
    if not start or not end:
        raise ValueError("Укажите даты начала и окончания аренды.")
    if start < today():
        raise ValueError("Дата начала не может быть в прошлом.")
    if end < start:
        raise ValueError("Дата окончания раньше даты начала.")
    if rent_days(start, end) > MAX_RENT_DAYS:
        raise ValueError(f"Максимальный срок аренды — {MAX_RENT_DAYS} дней.")


def prospective_items(tool_id, qty, start, end):
    """Содержимое корзины после добавления позиции (одинаковые даты объединяются) и индекс новой позиции."""
    cart = session.get("cart", [])
    s, e = start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")
    items = [{"tool_id": l["t"], "qty": l["q"], "start": parse_date(l["s"]), "end": parse_date(l["e"])} for l in cart]
    same = next((i for i, l in enumerate(cart) if l["t"] == tool_id and l["s"] == s and l["e"] == e), None)
    if same is not None:
        items[same]["qty"] = min(items[same]["qty"] + qty, 9999)
        return items, same
    items.append({"tool_id": tool_id, "qty": qty, "start": start, "end": end})
    return items, len(items) - 1


def cart_add_line(tool_id, qty, start, end):
    """Добавляет позицию в корзину после проверки дат и свободного количества (с учётом уже выбранных позиций)."""
    validate_period(start, end)
    if not 1 <= qty <= 9999:
        raise ValueError("Количество должно быть от 1 до 9999.")
    tool = get_tool(tool_id)
    if not tool:
        raise ValueError("Инструмент не найден или снят с аренды.")
    cart = session.get("cart", [])
    items, idx = prospective_items(tool_id, qty, start, end)
    if idx == len(cart) and len(cart) >= MAX_CART_LINES:
        raise ValueError(f"В заявке не может быть больше {MAX_CART_LINES} позиций.")
    av = lines_availability(items)[idx]
    if not av["ok"]:
        raise ValueError(f"Нельзя добавить «{tool['name']}»: {av['message']}")
    if idx < len(cart):
        cart[idx]["q"] = items[idx]["qty"]
    else:
        cart.append({"k": secrets.token_hex(4), "t": tool_id, "q": qty, "s": start.strftime("%Y-%m-%d"), "e": end.strftime("%Y-%m-%d")})
    session["cart"] = cart


# ---------------------------------------------------------------------------
# Купоны, баллы, оценочная выгода
# ---------------------------------------------------------------------------



def usable_coupons(t, uid, auto_rent_k=None):
    """Купоны пользователя: можно ли применить сейчас, и параметры для расчёта (spec)."""
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
        elif auto_rent_k is not None and auto_rent_k < kop(c["min_sum"]):
            why = f"минимальная сумма аренды {rub(c['min_sum'])}"
        out.append({"uc": uc["id_user_coupon"], "coupon": c, "end": end, "ok": not why, "why": why, "spec": coupon_spec(c),
                    "kind": coupon_kind(c)})
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


def create_order(user, lines, uc_id, points, delivery, logistics=None):
    """Создаёт заявку. Цены, доступность, скидки и итоги заново считаются здесь из данных Access;
    суммы из браузера не принимаются."""
    if not lines:
        raise ValueError("Заявка пуста.")
    if delivery not in ("Самовывоз", "Доставка"):
        raise ValueError("Выберите способ получения.")
    notes = []
    with ORDER_LOCK:                       # две одновременные заявки на один инструмент не пройдут мимо проверки остатка
        with Tx() as t:
            rows = _tool_rows({l["tool"]["id"] for l in lines}, t)
            items = []
            for l in lines:
                r = rows.get(l["tool"]["id"])
                if not r or not r["active"]:
                    raise ValueError(f"«{l['tool']['name']}» больше недоступен для аренды.")
                validate_period(l["start"], l["end"])
                items.append({"tool_id": r["id_tool"], "qty": l["qty"], "start": l["start"], "end": l["end"]})
            for l, a in zip(lines, lines_availability(items, t, rows)):
                if not a["ok"]:
                    raise ValueError(f"«{l['tool']['name']}», {fmt_d(l['start'])}–{fmt_d(l['end'])}: {a['message']}")
            cl = [{"price_k": kop(rows[l["tool"]["id"]]["price"]), "deposit_k": kop(rows[l["tool"]["id"]]["deposit"]),
                   "qty": l["qty"], "days": rent_days(l["start"], l["end"])} for l in lines]
            spec = cand = None
            if uc_id:
                cand = next((c for c in usable_coupons(t, user["id_user"], calc_quote(cl)["auto_rent"]) if c["uc"] == uc_id), None)
                if not cand or not cand["ok"]:
                    raise ValueError("Выбранный купон нельзя применить: " + (cand["why"] if cand else "он не найден") + ".")
                spec = cand["spec"]
            balance = t.one("SELECT [Баллы] FROM [Пользователь] WHERE [id_user] = ?", [user["id_user"]])["points"]
            qt = calc_quote(cl, spec, points, int(balance or 0))
            rid = t.insert("Бронирование", {
               "id_user": user["id_user"], "ДатаСоздания": now(), "Статус": O_NEW,
                "СуммаДоСкидки": from_k(qt["base"]), "СуммаСкидки": from_k(qt["base"] - qt["rent"]),
                "СуммаАренды": from_k(qt["rent"]), "СуммаЗалога": from_k(qt["deposit"]), "СпособДоставки": delivery})
            for l, ln in zip(lines, qt["lines"]):
                item = t.insert("ПозицияБронирования", {
                    "id_rent": rid, "id_tool": l["tool"]["id"], "Количество": l["qty"],
                    "ЦенаНаМоментБрони": money(rows[l["tool"]["id"]]["price"]), "ДатаНачала": l["start"],
                    "ДатаОкончания": l["end"], "Статус": P_NEW})
                t.insert("ИсторияСтатусовПозиций", {"id_item": item, "СтарыйСтатус": "—", "НовыйСтатус": P_NEW,
                                                   "КтоИзменил_id_user": user["id_user"], "Когда": now(),
                                                   "Комментарий": "Позиция создана"})
            t.insert("ИсторияСтатусов", {"id_rent": rid, "СтарыйСтатус": "—", "НовыйСтатус": O_NEW,
                                         "КтоИзменил_id_user": user["id_user"], "Когда": now(),
                                         "Комментарий": "Заявка создана пользователем" + (f" · {BIZ_MARK} {fmt_pct(qt['corp_bp'])}" if qt["corp_bp"] else "")})
            if cand and qt["coupon_ok"]:
                if t.x("UPDATE [КупонПользователя] SET [Использован] = ? WHERE [id_user_coupon] = ? AND [Использован] = ?",
                       [True, uc_id, False]) != 1:
                    raise ValueError("Этот купон уже использован.")
                t.insert("ИспользованиеКупона", {"id_coupon": cand["coupon"]["id_coupon"], "id_user": user["id_user"],
                                                 "id_rent": rid, "РазмерСкидки": from_k(qt["coupon"]),
                                                 "ДатаИспользования": now()})
            if qt["points"]:
                if t.x("UPDATE [Пользователь] SET [Баллы] = [Баллы] - ? WHERE [id_user] = ? AND [Баллы] >= ?",
                       [qt["points"], user["id_user"], qt["points"]]) != 1:
                    raise ValueError("Недостаточно баллов.")
                t.insert("ОперацииБаллов", {"id_user": user["id_user"], "id_rent": rid, "Количество": -qt["points"],
                                            "ТипОперации": "Списание", "Дата": now(),
                                            "Комментарий": f"Оплата части заявки №{rid} баллами"})
            notes.append((user["id_user"], f"Заявка №{rid} отправлена",
                          f"Мы получили вашу заявку №{rid} ({len(lines)} поз.). Администратор рассмотрит её, "
                          f"статус можно смотреть в личном кабинете.", "Заявка", rid))
            for aid in admin_ids(t):
                notes.append((aid, f"Новая заявка №{rid}", f"Пользователь {user['login']} отправил заявку №{rid}.",
                              "Заявка", rid, False))
    try:
        save_order_meta(rid, logistics)
    except OSError as e:
        print("[ToolRent] Не удалось записать order_meta.json:", e)
    notify_all(notes)
    return rid


def refresh_order(t, rid, who, comment=""):
    """Пересчитывает суммы и общий статус заявки; при завершении начисляет баллы, при полном отказе возвращает скидки."""
    notes = []
    order = t.one("SELECT * FROM [Бронирование] WHERE [id_rent] = ?", [rid])
    pos = t.q("SELECT * FROM [ПозицияБронирования] WHERE [id_rent] = ?", [rid])
    live = [p for p in pos if p["status"] not in DEAD]
    tools = _tool_rows({p["id_tool"] for p in live}, t)
    cl = [{"price_k": kop(p["price_at"]), "deposit_k": kop(tools[p["id_tool"]]["deposit"]) if p["id_tool"] in tools else 0,
           "qty": p["qty"], "days": rent_days(p["start"], p["end"])} for p in live]
    new = derive_status([p["status"] for p in pos])
    usage = t.one("SELECT * FROM [ИспользованиеКупона] WHERE [id_rent] = ?", [rid])
    spec = None
    if live and usage:
        c = t.one("SELECT * FROM [Купон] WHERE [id_coupon] = ?", [usage["id_coupon"]])
        if c:
            spec = coupon_spec(c)
            spec["min_k"] = 0                      # условия купона проверялись при оформлении
    spent = -sum(int(o["delta"]) for o in t.q(
        "SELECT [Количество] AS delta FROM [ОперацииБаллов] WHERE [id_rent] = ? AND [ТипОперации] = ?", [rid, "Списание"]))
    had_corp = any(BIZ_MARK in (h["comment"] or "") for h in t.q("SELECT * FROM [ИсторияСтатусов] WHERE [id_rent] = ?", [rid]))
    qt = calc_quote(cl, spec, spent if live else 0, None, enforce_max=False,
                    corp_bp=_bp(CORPORATE_DISCOUNT_PERCENT) if had_corp else 0)
    if usage and spec and D(usage["amount"]) != from_k(qt["coupon"]):
        t.update("ИспользованиеКупона", "id_usage", usage["id_usage"], {"РазмерСкидки": from_k(qt["coupon"])})
    t.update("Бронирование", "id_rent", rid, {"СуммаДоСкидки": from_k(qt["base"]), "СуммаСкидки": from_k(qt["base"] - qt["rent"]),
                                              "СуммаАренды": from_k(qt["rent"]), "СуммаЗалога": from_k(qt["deposit"]),
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
                "issue": (P_OK, P_OUT, "выдана"), "return": (P_OUT, P_BACK, "возвращена")}   # «принять все» — вариант return


def admin_item_action(admin, item_id, action, reason="", form=None):
    accept_all = action == "return_all"        # одна кнопка: все экземпляры позиции прошли проверку
    if accept_all:
        action = "return"
    if action not in ITEM_ACTIONS:
        raise ValueError("Неизвестное действие.")
    need, new, word = ITEM_ACTIONS[action]
    if action == "reject" and not reason.strip():
        raise ValueError("При отказе нужно указать причину.")
    notes = []
    with ORDER_LOCK, Tx() as t:
        it = t.one("SELECT * FROM [ПозицияБронирования] WHERE [id_item] = ?", [item_id])
        if not it:
            raise ValueError("Позиция не найдена.")
        if it["status"] != need:
            raise ValueError(f"Действие невозможно: позиция сейчас «{it['status']}».")
        order = t.one("SELECT * FROM [Бронирование] WHERE [id_rent] = ?", [it["id_rent"]])
        tool = t.one("SELECT * FROM [Инструмент] WHERE [id_tool] = ?", [it["id_tool"]])
        if action == "approve" and tool_free(t, it["id_tool"], tool_capacity(t, tool), as_dt(it["start"]), as_dt(it["end"]),
                                                exclude_item=item_id) < it["qty"]:
            raise ValueError("Недостаточно свободного инструмента на эти даты.")
        if action == "approve":
            assign_instances(t, it, admin["id_user"])
        elif action == "issue":
            issue_instances(t, it, admin["id_user"])
        elif action == "return":
            return_instances(t, it, admin["id_user"], form or {}, accept_all)
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
                if p["status"] == P_OK:
                    cancel_assignments(t, p["id_item"], user["id_user"])
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
            if h["new_status"] == O_DONE and not (h["comment"] or "").startswith(PAID_MARK):
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
        sent, expected = request.form.get("csrf") or request.headers.get("X-CSRF-Token", ""), session.get("_csrf", "")
        if not sent or not expected or not hmac.compare_digest(sent, expected):     # пустой токен не принимается
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


def rules_info():
    return dict(qty_units=QUANTITY_DISCOUNT_MIN_UNITS, qty_percent=fmt_pct(_bp(QUANTITY_DISCOUNT_PERCENT)),
                has_tiers=bool(DURATION_TIERS), tiers=[(d, fmt_pct(_bp(p))) for d, p in sorted(DURATION_TIERS)],
                points_percent=int(POINTS_PERCENT), points_max=fmt_pct(_bp(POINTS_MAX_PERCENT)))


def asset_version():
    """Метка версии style.css и script.js, чтобы браузер не держал старые файлы."""
    times = [os.path.getmtime(os.path.join(BASE_DIR, f)) for f in ("style.css", "script.js") if os.path.exists(os.path.join(BASE_DIR, f))]
    return int(max(times)) if times else 0


@app.context_processor
def inject_globals():
    ctx = dict(
        me=g.user, is_admin=bool(g.user and g.user["role"] == ROLE_ADMIN),
        is_b2b=bool(g.user and g.user["role"] == ROLE_B2B), unread=0, mode=current_mode(), b2b_pending=False,
        biz_view=current_mode() == "b2b", biz_real=is_biz_real(), biz_pct=fmt_pct(_bp(CORPORATE_DISCOUNT_PERCENT)),
        categories=[], cart_lines=[], cart_count=0,
        current=request.full_path.rstrip("?"), points_percent=int(POINTS_PERCENT), rent_cfg=json.dumps(rent_config(preview=request.endpoint == "tool" and current_mode() == "b2b")), rules=rules_info(),
        asset_v=asset_version(),
        mail_enabled=bool(_env("TOOLRENT_SMTP_HOST")), today_iso=date.today().isoformat())
    try:      # страницы ошибок должны открываться, даже если база недоступна
        lines = cart_lines()
        ctx.update(categories=get_categories(), cart_lines=lines, cart_count=len(lines))
        if g.user and g.user["role"] == ROLE_USER:
            ctx["b2b_pending"] = bool(q1("SELECT [id_b2b] FROM [B2B_Реквизиты] WHERE [id_user] = ?", [g.user["id_user"]]))
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


@app.route("/script.js")
def js():
    return send_from_directory(BASE_DIR, "script.js", mimetype="application/javascript")


@app.route("/img/<path:name>")
def img(name):
    return send_from_directory(IMG_DIR, name)

DOCS_DIR = os.path.join(BASE_DIR, "docs_files")
CLIENT_DOCS_DIR = os.path.join(BASE_DIR, "client_b2b_docs")
CLIENT_DOCS_PATH = os.path.join(BASE_DIR, "client_b2b_docs.json")
DOC_KINDS = {"card": "Карточка организации", "poa": "Доверенность на получение оборудования"}


def user_b2b_docs(uid):
    return json_load(CLIENT_DOCS_PATH, {}).get(str(uid)) or {}


def b2b_docs_approved(uid):
    """Оба документа загружены и имеют статус 'approved' (Подтверждён)."""
    d = user_b2b_docs(uid)
    return (d.get("card") or {}).get("status") == "approved" and (d.get("poa") or {}).get("status") == "approved"


@app.route("/docs/<path:name>")
def download_template_doc(name):
    return send_from_directory(DOCS_DIR, name, as_attachment=True)


@app.route("/b2b-doc/<int:uid>/<kind>")
@login_required
def download_client_doc(uid, kind):
    if kind not in DOC_KINDS:
        abort(404)
    if g.user["id_user"] != uid and g.user["role"] != ROLE_ADMIN:
        abort(403)
    info = user_b2b_docs(uid).get(kind)
    if not info:
        abort(404)
    return send_from_directory(CLIENT_DOCS_DIR, info["file"], as_attachment=True,
                               download_name=info.get("orig") or info["file"])


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
    cart_units = sum(l["q"] for l in session.get("cart", []))
    return render_template("tool.html", t=t, durations=DURATIONS, category=get_category(t["cat"]), similar=similar,
                           cart_units=cart_units)


@app.route("/mode/<m>")
def set_mode(m):
    if m in ("b2c", "b2b"):
        session["mode"] = m
        if g.user and g.user["role"] == ROLE_B2B and m == "b2c":
            flash(f"Режим для частных лиц: скидка для бизнеса ({fmt_pct(_bp(CORPORATE_DISCOUNT_PERCENT))}) не применяется. "
                  "Вернитесь в режим «Бизнесу», чтобы её получить.")
        elif m == "b2b" and not (g.user and g.user["role"] in (ROLE_B2B, ROLE_ADMIN)):
            flash("Режим просмотра для бизнеса. Скидка начнёт действовать после подтверждения бизнес-статуса администратором.")
    return redirect(safe_next())


@app.route("/b2b")
def b2b():
    return render_template("b2b.html", stats=catalog_stats(), org=q1("SELECT * FROM [B2B_Реквизиты] WHERE [id_user] = ?", [g.user["id_user"]]) if g.user else None)


# ---------------------------------------------------------------------------
# Корзина и оформление заявки
# ---------------------------------------------------------------------------
@app.route("/cart/add", methods=["POST"])
def cart_add():
    nxt = safe_next("catalog")
    try:
        start = parse_date(request.form.get("start")) or today()
        days = to_int(request.form.get("days"), 1, 1, MAX_RENT_DAYS)
        end = parse_date(request.form.get("end")) or start + timedelta(days=days - 1)
        cart_add_line(to_int(request.form.get("tool_id")), to_int(request.form.get("qty"), 1, 1, 9999), start, end)
        flash("Добавлено в корзину.")
        return redirect(url_for("cart") if request.form.get("go") == "cart" else nxt)
    except ValueError as e:
        flash(str(e))
        return redirect(nxt)


@app.route("/cart/remove/<k>", methods=["POST"])
def cart_remove(k):
    session["cart"] = [l for l in session.get("cart", []) if l["k"] != k]
    return redirect(url_for("cart"))


@app.route("/cart/clear", methods=["POST"])
def cart_clear():
    session["cart"] = []
    session.pop("cart_opts", None)
    return redirect(url_for("cart"))


DELIVERY = ("Самовывоз", "Доставка")


def cart_opts():
    o = session.get("cart_opts") or {}
    return {"coupon": to_int(o.get("coupon")), "points": to_int(o.get("points")),
            "delivery": o.get("delivery") if o.get("delivery") in DELIVERY else "Самовывоз"}


def save_cart_opts(coupon=None, points=None, delivery=None):
    o = cart_opts()
    if coupon is not None:
        o["coupon"] = to_int(coupon)
    if points is not None:
        o["points"] = to_int(points)
    if delivery in DELIVERY:
        o["delivery"] = delivery
    session["cart_opts"] = o


def cart_save_lines(raw):
    """Сохраняет изменения количества и дат позиций корзины. Инструмент позиции менять нельзя; возвращает {ключ: ошибка}."""
    cart, errors = session.get("cart", []), {}
    by_k = {l["k"]: l for l in cart}
    for r in raw:
        l = by_k.get(r.get("k"))
        if not l:
            continue
        try:
            qty = to_int(r.get("qty"), 0, 0, 9999)
            if not 1 <= qty <= 9999:
                raise ValueError("Количество должно быть от 1 до 9999.")
            start, end = parse_date(r.get("start")), parse_date(r.get("end"))
            validate_period(start, end)
            l["q"], l["s"], l["e"] = qty, start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")
        except ValueError as ex:
            errors[l["k"]] = str(ex)
    session["cart"] = cart
    return errors


def cart_context():
    """Всё, что нужно корзине и оформлению: позиции, доступность, купоны, итоги (считает Flask)."""
    opts = cart_opts()
    lines = cart_lines()
    for l, a in zip(lines, lines_availability([{"tool_id": l["tool"]["id"], "qty": l["qty"], "start": l["start"], "end": l["end"]}
                                               for l in lines])):
        l["avail"] = a
    ci = calc_input(lines)
    coupons, balance = [], 0
    if g.user:
        balance = int(g.user["points"] or 0)
        with Tx() as t:
            coupons = usable_coupons(t, g.user["id_user"], calc_quote(ci)["auto_rent"])
    chosen = next((c for c in coupons if c["uc"] == opts["coupon"] and c["ok"]), None)
    quote = calc_quote(ci, chosen["spec"] if chosen else None, opts["points"], balance)
    for l, ql in zip(lines, quote["lines"]):
        l["q"] = ql
    return dict(lines=lines, quote=quote, coupons=coupons, uc=chosen["uc"] if chosen else 0, balance=balance,
                points=quote["points"], delivery=opts["delivery"], has_problem=any(not l["avail"]["ok"] for l in lines))


def form_lines():
    out = []
    for l in session.get("cart", []):
        k = l["k"]
        if f"qty-{k}" not in request.form:
            continue
        start = parse_date(request.form.get(f"start-{k}"))
        end = parse_date(request.form.get(f"end-{k}")) or (
            start + timedelta(days=to_int(request.form.get(f"days-{k}"), 1, 1, MAX_RENT_DAYS) - 1) if start else None)
        out.append({"k": k, "qty": request.form.get(f"qty-{k}"), "start": request.form.get(f"start-{k}"),
                    "end": end.strftime("%Y-%m-%d") if end else ""})
    return out


@app.route("/cart", methods=["GET", "POST"])
def cart():
    if request.method == "POST":
        errors = cart_save_lines(form_lines())
        save_cart_opts(request.form.get("coupon", 0), request.form.get("points", 0), request.form.get("delivery"))
        for msg in errors.values():
            flash(msg)
        if request.form.get("action") == "checkout" and not errors:
            ctx = cart_context()
            if not ctx["lines"]:
                flash("Корзина пуста.")
            elif ctx["has_problem"]:
                flash("Проверьте доступность: для некоторых позиций выбрано больше, чем свободно на эти даты.")
            elif not g.user:
                flash("Чтобы оформить заявку, войдите в аккаунт или зарегистрируйтесь. Корзина сохранена.")
                return redirect(url_for("login", next=url_for("booking")))
            else:
                return redirect(url_for("booking"))
        return redirect(url_for("cart"))
    return render_template("cart.html", **cart_context())


@app.route("/api/cart/quote", methods=["POST"])
def api_cart_quote():
    """Живой пересчёт корзины: браузер присылает только количество, даты, купон и баллы; цены берутся из Access."""
    d = request.get_json(silent=True) or {}
    errors = cart_save_lines(d.get("lines") or [])
    save_cart_opts(d.get("coupon"), d.get("points"))
    ctx = cart_context()
    q_ = ctx["quote"]
    return jsonify(
        lines=[{"k": l["k"], "ok": l["avail"]["ok"], "free": l["avail"]["free"], "message": l["avail"]["message"],
                "error": errors.get(l["k"], ""), "base": l["q"]["base"], "dur": l["q"]["dur"], "dur_bp": l["q"]["dur_bp"],
                "qty_disc": l["q"]["qty_disc"], "corp": l["q"]["corp"], "rent": l["q"]["rent"], "deposit": l["q"]["deposit"]}
               for l in ctx["lines"]],
        quote={k: q_[k] for k in ("units", "qty_bp", "corp_bp", "base", "dur", "qty", "corp", "auto_rent", "coupon", "coupon_ok",
                                  "coupon_why", "max_points", "points", "points_k", "rent", "deposit", "grand", "savings")},
        coupons=[{"uc": c["uc"], "ok": c["ok"], "why": c["why"]} for c in ctx["coupons"]], coupon=ctx["uc"],
        balance=ctx["balance"], has_problem=ctx["has_problem"], count=len(ctx["lines"]))


@app.route("/api/availability", methods=["POST"])
def api_availability():
    """Свободное количество на период для страницы инструмента (с учётом позиций, уже лежащих в корзине)."""
    d = request.get_json(silent=True) or {}
    try:
        tool_id, qty = to_int(d.get("tool_id")), to_int(d.get("qty"), 0, 0, 9999)
        start, end = parse_date(d.get("start")), parse_date(d.get("end"))
        validate_period(start, end)
        if qty < 1:
            raise ValueError("Количество должно быть не меньше 1.")
        tool_row = _tool_rows([tool_id]).get(tool_id)
        if not tool_row or not tool_row["active"]:
            raise ValueError("Инструмент недоступен для аренды.")
        items, idx = prospective_items(tool_id, qty, start, end)
        a = lines_availability(items)[idx]
        return jsonify(ok=a["ok"], free=a["free"], stock=a["stock"], message=a["message"], error="")
    except ValueError as e:
        return jsonify(ok=False, free=0, stock=0, message="", error=str(e))

PICKUPS_PATH = os.path.join(BASE_DIR, "pickup_points.json")
ORDER_META_PATH = os.path.join(BASE_DIR, "order_meta.json")
JSON_LOCK = threading.RLock()


def json_load(path, default):
    with JSON_LOCK:
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except FileNotFoundError:
            return default
        except ValueError:                 # повреждённый файл не затираем, а откладываем в сторону
            os.replace(path, f"{path}.broken-{int(time.time())}")
            print(f"[ToolRent] {os.path.basename(path)} повреждён, сохранён как .broken-*, начат пустой.")
            return default


def json_save(path, data):
    with JSON_LOCK:
        tmp = path + ".tmp"                # сначала во временный файл, потом атомарная замена
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)


def pickups_load():
    d = json_load(PICKUPS_PATH, {"next_id": 1, "points": []})
    d.setdefault("points", [])
    d.setdefault("next_id", 1 + max([p["id_pickup"] for p in d["points"]] or [0]))
    return d


def all_pickups():
    return sorted(pickups_load()["points"], key=lambda p: p["id_pickup"])


def active_pickups():
    return sorted((p for p in pickups_load()["points"] if p.get("active")),
                  key=lambda p: (p["name"].lower(), p["id_pickup"]))


def save_order_meta(rid, lg):
    if not lg:
        return
    with JSON_LOCK:
        meta = json_load(ORDER_META_PATH, {})
        meta[str(rid)] = {"object": lg.get("object"), "pickup": lg.get("pickup"),
                          "delivery_address": lg.get("address")}
        json_save(ORDER_META_PATH, meta)


def attach_meta(orders):
    """Добавляет к строкам заявок object / pickup / delivery_address из order_meta.json (у старых заявок их нет → None)."""
    meta = json_load(ORDER_META_PATH, {})
    for o in orders:
        m = meta.get(str(o["id_rent"])) or {}
        o["object"], o["pickup"], o["delivery_address"] = m.get("object"), m.get("pickup"), m.get("delivery_address")
    return orders

def parse_logistics(delivery, pickups):
    """Проверяет объект, точку самовывоза / адрес доставки. Всё берётся из формы, но точка — только из активных."""
    my_addr = (g.user["address"] or "").strip()
    if request.form.get("use_my_address") == "1":
        if not my_addr:
            raise ValueError("В личном кабинете не указан адрес. Укажите его или впишите объект вручную.")
        obj = my_addr
    else:
        obj = field("object", 255, True, "Объект")
    pickup = address = None
    if delivery == "Самовывоз":
        p = next((x for x in pickups if x["id_pickup"] == to_int(request.form.get("pickup_id"))), None)
        if not p:
            raise ValueError("Выберите точку самовывоза.")
        pickup = f"{p['name']} — {p['address']} ({p['schedule']})"[:255]
    else:
        address = field("delivery_address", 255, True, "Точный адрес доставки")
    return {"object": obj[:255], "pickup": pickup, "address": address}


@app.route("/booking", methods=["GET", "POST"])
def booking():
    if not g.user:
        flash("Чтобы оформить заявку, войдите в аккаунт или зарегистрируйтесь. Корзина сохранена.")
        return redirect(url_for("login", next=url_for("booking")))
    ctx = cart_context()
    if not ctx["lines"]:
        flash("Корзина пуста.")
        return redirect(url_for("cart"))
    pickups = active_pickups()
    b2b_blocked = ""
    if current_mode() == "b2b":
        if g.user["role"] != ROLE_B2B:
            b2b_blocked = "В режиме «Бизнесу» оформление доступно только после подтверждения бизнес-статуса администратором."
        elif not b2b_docs_approved(g.user["id_user"]):
            b2b_blocked = "Оба документа (Карточка организации и Доверенность) должны быть загружены и подтверждены администратором."
    ctx.update(pickups=pickups, form={}, my_address=(g.user["address"] or "").strip(), b2b_blocked=b2b_blocked)
    if request.method == "POST":
        if b2b_blocked:
            flash(b2b_blocked)
            return redirect(url_for("booking"))

        save_cart_opts(delivery=request.form.get("delivery"))
        ctx["delivery"] = cart_opts()["delivery"]
        if ctx["has_problem"]:
            flash("Часть позиций недоступна на выбранные даты. Измените их в корзине.")
            return redirect(url_for("cart"))
        try:
            lg = parse_logistics(ctx["delivery"], pickups)
        except ValueError as e:                      # остаёмся на странице, введённое не теряется
            flash(str(e))
            ctx["form"] = request.form
            return render_template("booking.html", **ctx)
        try:
            rid = create_order(g.user, ctx["lines"], ctx["uc"], ctx["points"], ctx["delivery"], lg)
        except ValueError as e:
            flash(str(e))
            return redirect(url_for("cart"))
        session["cart"] = []
        session.pop("cart_opts", None)
        flash(f"Заявка №{rid} отправлена. Статус можно отслеживать в личном кабинете.")
        return redirect(url_for("profile_order", rid=rid))
    return render_template("booking.html", **ctx)

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
    """Письмо подтверждения. Подтверждение почты необязательно: любая ошибка здесь не должна ломать регистрацию."""
    try:
        with Tx() as t:
            raw = make_token(t, uid, "email")
        send_email(email, "Подтвердите почту на ToolRent",
                   f"Здравствуйте!\n\nЧтобы подтвердить адрес электронной почты, перейдите по ссылке (действует {TOKEN_HOURS['email']} ч.):\n"
                   f"{base_url()}/verify/{raw}\n\nЕсли вы не регистрировались на ToolRent, просто проигнорируйте это письмо.")
    except Exception as e:
        print("[ToolRent] письмо подтверждения не отправлено:", e)


def login_user(uid, remember=True):
    keep = {k: session[k] for k in ("cart", "cart_opts", "mode") if k in session}    # корзина гостя не теряется при входе
    session.clear()
    session.update(keep)
    session["uid"] = uid
    session.permanent = remember
    row = q1("SELECT [Роль] AS role FROM [Пользователь] WHERE [id_user] = ?", [uid])
    if row and row["role"] == ROLE_B2B:
        session["mode"] = "b2b"            # подтверждённый бизнес-клиент по умолчанию работает в бизнес-режиме
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
        nxt = safe_next("profile")
        login_user(uid)
        flash("Аккаунт создан. " + ("Мы отправили письмо со ссылкой для подтверждения почты, но она необязательна: "
                                    "сайтом можно пользоваться сразу." if _env("TOOLRENT_SMTP_HOST")
                                    else "Подтверждение почты необязательно: сайтом можно пользоваться сразу."))
        return redirect(nxt)
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
PAID_MARK = "Демо-оплата"      # запись в ИсторияСтатусов с этим началом комментария = демонстрационная оплата выполнена


def paid_set(rids):
    out = set()
    for part in _chunks(rids):
        for h in q(f"SELECT [id_rent], [Комментарий] FROM [ИсторияСтатусов] WHERE [id_rent] IN ({', '.join(['?'] * len(part))})", part):
            if (h["comment"] or "").startswith(PAID_MARK):
                out.add(h["id_rent"])
    return out


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
    paid = paid_set([o["id_rent"] for o in orders])
    for o in orders:
        o["n_items"] = counts.get(o["id_rent"], 0)
        o["paid"] = o["id_rent"] in paid
    attach_meta(orders)
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
        data = dict(org=q1("SELECT * FROM [B2B_Реквизиты] WHERE [id_user] = ?", [u["id_user"]]),
                    org_types=ORG_TYPES, b2b_docs=user_b2b_docs(u["id_user"]))
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
    attach_meta([order])
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
    paid = order["id_rent"] in paid_set([rid])
    pay_k = kop(order["rent_sum"]) + kop(order["deposit_sum"])
    return render_template("profile.html", section="order", sections=SECTIONS, order=order, positions=pos,
                           history=history, can_cancel=can_cancel, paid=paid, pay_k=pay_k,
                           can_pay=(not paid) and order["status"] not in (O_NO, O_CANCEL))


@app.route("/profile/orders/<int:rid>/pay", methods=["POST"])
@login_required
def profile_order_pay(rid):
    """Демонстрационная оплата: реальных денежных операций нет, платёжных данных не принимается и не хранится."""
    try:
        with Tx() as t:
            order = t.one("SELECT * FROM [Бронирование] WHERE [id_rent] = ? AND [id_user] = ?", [rid, g.user["id_user"]])
            if not order:
                abort(404)
            if order["status"] in (O_NO, O_CANCEL):
                raise ValueError("Эта заявка не подлежит оплате.")
            if any((h["comment"] or "").startswith(PAID_MARK) for h in t.q(
                    "SELECT [Комментарий] FROM [ИсторияСтатусов] WHERE [id_rent] = ?", [rid])):
                raise ValueError("Заявка уже оплачена.")
            amount = from_k(kop(order["rent_sum"]) + kop(order["deposit_sum"]))
            t.insert("ИсторияСтатусов", {"id_rent": rid, "СтарыйСтатус": order["status"], "НовыйСтатус": order["status"],
                                         "КтоИзменил_id_user": g.user["id_user"], "Когда": now(),
                                         "Комментарий": f"{PAID_MARK}: выполнена, {rub(amount)} (аренда и залог). Реальных денежных операций нет."})
        notify(g.user["id_user"], "Оплата успешно выполнена", f"Заявка №{rid}: демонстрационная оплата на {rub(amount)} выполнена.", "Заявка", rid, mail=False)
        flash("Оплата успешно выполнена")
    except ValueError as e:
        flash(str(e))
    return redirect(url_for("profile_order", rid=rid))


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
        cur_docs = user_b2b_docs(u["id_user"])
        saved_files = {}
        os.makedirs(CLIENT_DOCS_DIR, exist_ok=True)
        for kind, label in DOC_KINDS.items():
            fs = request.files.get(f"doc_{kind}")
            if fs and fs.filename:
                orig = os.path.basename(fs.filename)
                if not orig.lower().endswith(".pdf"):
                    raise ValueError(f"Файл «{label}» должен быть в формате PDF.")
                raw = fs.read()
                if not raw.startswith(b"%PDF-"):
                    raise ValueError(f"Файл «{label}» повреждён или не является PDF.")
                fname = f"u{u['id_user']}_{kind}_{secrets.token_hex(4)}.pdf"
                with open(os.path.join(CLIENT_DOCS_DIR, fname), "wb") as f:
                    f.write(raw)
                saved_files[kind] = {"file": fname, "orig": orig, "status": "pending",
                                     "reason": "", "at": now().strftime("%d.%m.%Y %H:%M")}
            elif not cur_docs.get(kind):
                raise ValueError(f"Прикрепите PDF-файл: {label}.")
        if saved_files:
            with JSON_LOCK:
                all_d = json_load(CLIENT_DOCS_PATH, {})
                u_d = all_d.get(str(u["id_user"]), {})
                for k, v in saved_files.items():
                    old = u_d.get(k, {}).get("file")
                    if old and old != v["file"]:
                        try:
                            os.remove(os.path.join(CLIENT_DOCS_DIR, old))
                        except OSError:
                            pass
                    u_d[k] = v
                all_d[str(u["id_user"])] = u_d
                json_save(CLIENT_DOCS_PATH, all_d)

        with Tx() as t:
            row = t.one("SELECT * FROM [B2B_Реквизиты] WHERE [id_user] = ?", [u["id_user"]])
            vals = {"ТипОрганизации": org, "НазваниеКомпании": company, "ИНН": inn, "КПП": kpp or None, "ЮрАдрес": addr}
            if row:
                t.update("B2B_Реквизиты", "id_b2b", row["id_b2b"], vals)
            else:
                t.insert("B2B_Реквизиты", {"id_user": u["id_user"], **vals})
            admins = admin_ids(t)
        if u["role"] == ROLE_USER:
            again = bool(row)
            notify(u["id_user"], "Заявка на бизнес-статус отправлена",
                   f"Реквизиты «{company}» (ИНН {inn}) отправлены администратору. Пока заявка не подтверждена, скидка для бизнеса "
                   "не действует. Мы сообщим о решении.", "Система")
            for aid in admins:
                notify(aid, "Заявка на бизнес-статус" + (" (данные обновлены)" if again else ""),
                       f"Пользователь {u['login']} ({u['fio']}) просит бизнес-статус: {org} «{company}», ИНН {inn}. "
                       "Откройте его карточку в админ-панели («Пользователи») и подтвердите или отклоните заявку.", "Система")
            flash("Заявка отправлена администратору. Бизнес-статус и скидка появятся после подтверждения.")
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
    orgs = {r["id_user"] for r in q("SELECT [id_user] FROM [B2B_Реквизиты]")}
    n_b2b = len([u for u in users if u["role"] == ROLE_USER and u["id_user"] in orgs])
    return adm("home", n_b2b=n_b2b, n_new=count(P_NEW), n_ok=count(P_OK), n_out=count(P_OUT), n_over=len(overdue),
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
                "Характеристики": (join_specs(request.form.getlist("spec_k"), request.form.getlist("spec_v"))
                                   + "\n" + "\n".join(f"{k}: {v}" if k else v for k, v in parse_specs(request.form.get("specs_bulk", "")))).strip()[:20000] or None,
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
        with ORDER_LOCK, Tx() as t:
            old_photo = None
            if tool_id:
                old = t.one("SELECT * FROM [Инструмент] WHERE [id_tool] = ?", [tool_id])
                if not old:
                    raise ValueError("Инструмент не найден.")
                old_photo = old["photo"]
                owned = capacity_map(t).get(tool_id, (0, 0))[1]
                if has_instances(t, tool_id) and vals["КоличествоНаСкладе"] < owned:
                    raise ValueError(f"Количество экземпляров ({owned}) нельзя уменьшить вручную: спишите лишние экземпляры "
                                     "на странице «Экземпляры».")
                t.update("Инструмент", "id_tool", tool_id, vals)
            else:
                art = (request.form.get("article") or "").strip().upper()      # необязательно: иначе серия назначается сама
                if art:
                    m = re.match(r"^([A-Z]{3})(-000000)?$", art)
                    if not m:
                        raise ValueError("Артикул: три латинские буквы серии, например ABC (или ABC-000000).")
                    taken = {(ARTICLE_RE.match((r["article"] or "").upper()) or [None, None])[1] for r in t.q("SELECT [Артикул] FROM [Инструмент]")}
                    if m.group(1) in taken:
                        raise ValueError("Такая серия артикулов уже используется другой моделью.")
                    vals["Артикул"] = f"{m.group(1)}-000000"
                tool_id = t.insert("Инструмент", vals)
            sync_tool(t, tool_id, g.user["id_user"])      # серия и экземпляры по количеству
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


@app.route("/admin/tools/<int:tool_id>/instances")
@admin_required
def admin_instances(tool_id):
    tool = get_tool(tool_id, include_inactive=True)
    if not tool:
        abort(404)
    insts = sorted(q("SELECT * FROM [ФизическийЭкземпляр] WHERE [id_tool] = ?", [tool_id]),
                   key=lambda i: (i["state"] == STATE_OFF, i["id_instance"]))
    ids = [i["id_instance"] for i in insts]
    asg, jr = {i: [] for i in ids}, {i: [] for i in ids}
    for part in _chunks(ids):
        ph = ", ".join(["?"] * len(part))
        for a in q(f"SELECT * FROM [ЭкземплярыПозиции] WHERE [id_instance] IN ({ph})", part):
            asg[a["id_instance"]].append(a)
        for j in q(f"SELECT * FROM [ЖурналЭкземпляров] WHERE [id_instance] IN ({ph})", part):
            jr[j["id_instance"]].append(j)
    items = {}
    for part in _chunks({a["id_item"] for v in asg.values() for a in v}):
        for p in q(f"SELECT * FROM [ПозицияБронирования] WHERE [id_item] IN ({', '.join(['?'] * len(part))})", part):
            items[p["id_item"]] = p
    for i in insts:
        i["asg"] = sorted(asg[i["id_instance"]], key=lambda a: -a["id_instance_rent"])
        for a in i["asg"]:
            a["item"] = items.get(a["id_item"])
        i["current"] = next((a for a in i["asg"] if a["status"] in (A_ASSIGNED, A_OUT)), None)
        i["log"] = sorted(jr[i["id_instance"]], key=lambda j: -j["id_log"])[:40]
        grades = [a["grade"] for a in i["asg"] if a["grade"] is not None]
        i["avg_grade"] = round(sum(grades) / len(grades)) if grades else None
    um = user_map(j["by_user"] for i in insts for j in i["log"])
    return adm("instances", tool=tool, insts=insts, states=INSTANCE_STATES, um=um)


@app.route("/admin/instances/<int:iid>/state", methods=["POST"])
@admin_required
def admin_instance_state(iid):
    tool_id = 0
    try:
        tool_id, repaired = change_instance_state(g.user, iid, request.form.get("state", ""), field("comment", 255))
        flash("Ремонт завершён, администраторам отправлено уведомление. Подтвердите приёмку, выбрав состояние «Доступен»."
              if repaired else "Состояние экземпляра изменено.")
    except ValueError as e:
        flash(str(e))
        row = q1("SELECT [id_tool] FROM [ФизическийЭкземпляр] WHERE [id_instance] = ?", [iid])
        tool_id = row["id_tool"] if row else 0
    return redirect(url_for("admin_instances", tool_id=tool_id) if tool_id else url_for("admin_tools"))


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

@app.route("/admin/pickups")
@admin_required
def admin_pickups():
    return adm("pickups", points=all_pickups())


@app.route("/admin/pickups/save", methods=["POST"])
@admin_required
def admin_pickup_save():
    pid = to_int(request.form.get("id"))
    try:
        vals = {"name": field("name", 255, True, "Название"), "address": field("address", 255, True, "Адрес"),
                "schedule": field("schedule", 255, True, "График работы")}
        with JSON_LOCK:
            d = pickups_load()
            if pid:
                p = next((x for x in d["points"] if x["id_pickup"] == pid), None)
                if not p:
                    raise ValueError("Точка не найдена.")
                p.update(vals)
                what = "изменена"
            else:
                pid = d["next_id"]
                d["next_id"] = pid + 1
                d["points"].append({"id_pickup": pid, **vals, "active": True})
                what = "создана"
            json_save(PICKUPS_PATH, d)
        with Tx() as t:
            admin_log(t, g.user["id_user"], f"Точка самовывоза: {what}", None, f"№{pid} «{vals['name']}»")
        flash("Готово.")
    except (ValueError, OSError) as e:
        flash(str(e))
    return redirect(url_for("admin_pickups"))


@app.route("/admin/pickups/<int:pid>/toggle", methods=["POST"])
@admin_required
def admin_pickup_toggle(pid):
    with JSON_LOCK:
        d = pickups_load()
        p = next((x for x in d["points"] if x["id_pickup"] == pid), None)
        if not p:
            abort(404)
        p["active"] = not p.get("active")
        json_save(PICKUPS_PATH, d)
    with Tx() as t:
        admin_log(t, g.user["id_user"], "Точка самовывоза: " + ("включена" if p["active"] else "отключена"),
                  None, f"№{pid} «{p['name']}»")
    return redirect(url_for("admin_pickups"))


@app.route("/admin/pickups/<int:pid>/delete", methods=["POST"])
@admin_required
def admin_pickup_delete(pid):
    with JSON_LOCK:
        d = pickups_load()
        p = next((x for x in d["points"] if x["id_pickup"] == pid), None)
        if not p:
            abort(404)
        d["points"] = [x for x in d["points"] if x["id_pickup"] != pid]
        json_save(PICKUPS_PATH, d)
    with Tx() as t:
        admin_log(t, g.user["id_user"], "Точка самовывоза: удалена", None, f"№{pid} «{p['name']}»")
    flash("Точка удалена. В уже созданных заявках она сохранится.")
    return redirect(url_for("admin_pickups"))

# --- заявки --------------------------------------------------------------
@app.route("/admin/orders")
@admin_required
def admin_orders():
    status = request.args.get("status", "")
    orders = sorted(q("SELECT * FROM [Бронирование]"), key=lambda o: -o["id_rent"])
    if status:
        orders = [o for o in orders if o["status"] == status]
    attach_meta(orders[:300])
    return adm("orders", orders=orders[:300], um=user_map(o["id_user"] for o in orders[:300]), status=status,
               paid=paid_set([o["id_rent"] for o in orders[:300]]),
               statuses=[O_NEW, O_WORK, O_RENT, O_DONE, O_NO, O_CANCEL])


@app.route("/admin/orders/<int:rid>")
@admin_required
def admin_order(rid):
    order = q1("SELECT * FROM [Бронирование] WHERE [id_rent] = ?", [rid])
    if not order:
        abort(404)
    attach_meta([order])
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
            p["free"] = tool_free(t, p["id_tool"], p["tool"]["stock"], as_dt(p["start"]), as_dt(p["end"]), exclude_item=p["id_item"]) if p["tool"] else 0
            p["asg"] = []
            for a in t.q("SELECT * FROM [ЭкземплярыПозиции] WHERE [id_item] = ?", [p["id_item"]]):
                if a["status"] != A_CANCEL:
                    a["inst"] = t.one("SELECT * FROM [ФизическийЭкземпляр] WHERE [id_instance] = ?", [a["id_instance"]])
                    p["asg"].append(a)
            p["history"] = sorted(hist[p["id_item"]], key=lambda h: h["id_item_log"])
    history = sorted(q("SELECT * FROM [ИсторияСтатусов] WHERE [id_rent] = ?", [rid]), key=lambda h: h["id_log"])
    um.update(user_map(h["by_user"] for p in pos for h in p["history"]))
    return adm("order", order=order, positions=pos, history=history, um=um, customer=um.get(order["id_user"]),
               paid=rid in paid_set([rid]))


@app.route("/admin/orders/<int:rid>/item", methods=["POST"])
@admin_required
def admin_order_item(rid):
    item = to_int(request.form.get("item_id"))
    try:
        it = q1("SELECT * FROM [ПозицияБронирования] WHERE [id_item] = ? AND [id_rent] = ?", [item, rid])
        if not it:
            raise ValueError("Позиция не найдена.")
        admin_item_action(g.user, item, request.form.get("action", ""), field("reason", 255), request.form)
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
    orgs = {r["id_user"] for r in q("SELECT [id_user] FROM [B2B_Реквизиты]")}
    wait = {u["id_user"] for u in users if u["role"] == ROLE_USER and u["id_user"] in orgs}
    flt = request.args.get("f", "")
    if flt == "b2b_wait":
        users = [u for u in users if u["id_user"] in wait]
    return adm("users", users=sorted(users, key=lambda u: -u["id_user"])[:300], q=request.args.get("q", ""), wait=wait, flt=flt)


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
    org = q1("SELECT * FROM [B2B_Реквизиты] WHERE [id_user] = ?", [uid])
    return adm("user", u=u, org=org, b2b_docs=user_b2b_docs(uid), pending=bool(org and u["role"] == ROLE_USER), stats=user_stats(uid),
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
            elif action == "b2b-approve":
                org_row = t.one("SELECT * FROM [B2B_Реквизиты] WHERE [id_user] = ?", [uid])
                if u["role"] != ROLE_USER or not org_row:
                    raise ValueError("Нет заявки на бизнес-статус.")
                if not b2b_docs_approved(uid):
                    raise ValueError("Сначала подтвердите оба документа клиента (Карточку организации и Доверенность).")
                t.update("Пользователь", "id_user", uid, {"Роль": ROLE_B2B})
                admin_log(t, me["id_user"], "Бизнес-статус подтверждён", uid, f"{org_row['company']}, ИНН {org_row['inn']}")
                notes.append((uid, "Бизнес-статус подтверждён",
                              f"Администратор подтвердил бизнес-профиль «{org_row['company']}». Включите режим «Бизнесу» в шапке сайта: "
                              f"скидка для бизнеса {fmt_pct(_bp(CORPORATE_DISCOUNT_PERCENT))} действует на всю аренду (залог без скидки).", "Система"))
            elif action == "doc-review":
                kind = request.form.get("kind", "")
                decision = request.form.get("decision", "")
                if kind not in DOC_KINDS or decision not in ("approved", "rejected"):
                    raise ValueError("Неверные параметры проверки документа.")
                reason = field("reason", 255, decision == "rejected", "Причина отклонения")
                with JSON_LOCK:
                    all_d = json_load(CLIENT_DOCS_PATH, {})
                    u_d = all_d.get(str(uid)) or {}
                    if kind not in u_d:
                        raise ValueError("Документ ещё не загружен клиентом.")
                    u_d[kind]["status"] = decision
                    u_d[kind]["reason"] = reason if decision == "rejected" else ""
                    all_d[str(uid)] = u_d
                    json_save(CLIENT_DOCS_PATH, all_d)
                word = "подтверждён" if decision == "approved" else "отклонён"
                admin_log(t, me["id_user"], f"Документ {word}: {DOC_KINDS[kind]}", uid, reason)
                notes.append((uid, f"Документ {word}",
                              f"Документ «{DOC_KINDS[kind]}» {word} администратором."
                              + (f" Причина: {reason}. Загрузите исправленный файл в бизнес-профиле." if decision == "rejected" else ""),
                              "Система"))

            elif action == "b2b-reject":
                org_row = t.one("SELECT * FROM [B2B_Реквизиты] WHERE [id_user] = ?", [uid])
                if u["role"] != ROLE_USER or not org_row:
                    raise ValueError("Нет заявки на бизнес-статус.")
                reason = field("reason", 255, True, "Причина отказа")
                t.x("DELETE FROM [B2B_Реквизиты] WHERE [id_user] = ?", [uid])
                admin_log(t, me["id_user"], "Заявка на бизнес-статус отклонена", uid, reason)
                notes.append((uid, "Заявка на бизнес-статус отклонена",
                              f"Причина: {reason}. Вы можете исправить данные и отправить заявку снова в личном кабинете.", "Система"))
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
        if kind not in (COUPON_PERCENT, COUPON_FIXED):
            raise ValueError("Выберите тип скидки.")
        amount = to_money(request.form.get("amount", ""), "Размер скидки")
        if amount <= 0 or (kind == COUPON_PERCENT and amount > 100):
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
                   "КоличествоНаСкладе", "ИмяФото", "Активен", "id_category", "Артикул"],
    "ФизическийЭкземпляр": ["id_instance", "id_tool", "ИнвентарныйКод", "Состояние", "ДатаДобавления", "Примечание"],
    "ЭкземплярыПозиции": ["id_instance_rent", "id_item", "КодВыдачи", "Статус", "ОценкаКачества", "Комментарий", "ДатаВыдачи",
                          "ДатаВозврата", "КтоОсмотрел_id_user", "id_instance"],
    "ЖурналЭкземпляров": ["id_log", "id_instance", "id_instance_rent", "ИнвентарныйКодНаМомент", "Событие", "СостояниеДо",
                          "СостояниеПосле", "ОценкаКачества", "Комментарий", "ДатаСобытия", "КтоИзменил_id_user"],
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


# ---------------------------------------------------------------------------
# Информационные страницы ToolRent
# ---------------------------------------------------------------------------

INFO_PAGES = {
    "about": {
        "title": "О нас",
        "lead": (
            "ToolRent — сервис для удобного подбора и аренды "
            "строительного и профессионального инструмента."
        ),
        "blocks": [
            (
                "Инструмент для ваших задач",
                "В каталоге ToolRent можно подобрать оборудование "
                "для строительных, монтажных, ремонтных и других работ. "
                "В карточке инструмента указаны основные характеристики, "
                "стоимость аренды и размер залога."
            ),
            (
                "Для частных и корпоративных клиентов",
                "Сервис рассчитан как на частных пользователей, так и на "
                "организации. Для компаний предусмотрена подача заявки "
                "на подтверждение бизнес-статуса администратором."
            ),
            (
                "Контроль состояния оборудования",
                "Каждый физический экземпляр инструмента учитывается отдельно. "
                "После возврата оборудование проходит осмотр, по результатам "
                "которого его можно вернуть в доступный фонд, отправить "
                "на проверку или ремонт либо списать."
            ),
        ],
    },

    "geography": {
        "title": "Где мы работаем",
        "lead": (
            "Доступность самовывоза и доставки зависит от действующих "
            "точек получения и условий, указанных при оформлении заказа."
        ),
        "blocks": [
            (
                "Точки самовывоза",
                "Если для выбранного заказа доступны активные точки самовывоза, "
                "их можно выбрать при оформлении. Адрес и условия получения "
                "уточняются в информации о конкретной точке."
            ),
            (
                "Доставка",
                "Если доставка доступна, её можно выбрать при оформлении заказа. "
                "Возможность доставки и её условия зависят от настроек сервиса "
                "и выбранного способа получения."
            ),
            (
                "Проверяйте условия заказа",
                "Перед отправкой заявки проверьте выбранный способ получения, "
                "даты аренды и доступность инструмента. Если нужный вариант "
                "получения не отображается, обратитесь к администратору."
            ),
        ],
    },

    "rules": {
        "title": "Правила аренды",
        "lead": (
            "Основные этапы аренды инструмента в ToolRent: "
            "выбор, бронирование, получение и возврат."
        ),
        "blocks": [
            (
                "1. Выбор и оформление",
                "Выберите инструмент, укажите количество и даты аренды, "
                "добавьте позиции в корзину и проверьте итоговый расчёт. "
                "Стоимость аренды и залог отображаются отдельно."
            ),
            (
                "2. Отправка заявки и резерв",
                "После оформления заявка поступает на рассмотрение администратора. "
                "Позиции заказа резервируются с учётом выбранных дат и доступности. "
                "Отправка заявки сама по себе не означает, что инструмент уже выдан."
            ),
            (
                "3. Подтверждение и получение",
                "Администратор проверяет заявку и её отдельные позиции. "
                "Получение оборудования производится согласованным способом "
                "после подтверждения и выполнения необходимых условий."
            ),
            (
                "4. Залог и стоимость аренды",
                "Стоимость аренды и обеспечительный залог рассчитываются отдельно. "
                "Скидки, купоны и бонусные баллы применяются к аренде согласно "
                "условиям сервиса; залог не уменьшается за счёт скидок и баллов."
            ),
            (
                "5. Возврат и осмотр",
                "Верните инструмент в согласованный срок и способом, "
                "указанным для заказа. После возврата администратор проверяет "
                "состояние каждого экземпляра. Решение по залогу принимается "
                "с учётом состояния оборудования и условий аренды."
            ),
        ],
    },

    "offer": {
        "title": "Публичная оферта",
        "lead": (
            "Основные принципы оформления аренды через ToolRent. "
            "Перед использованием сервиса ознакомьтесь с условиями заказа."
        ),
        "blocks": [
            (
                "1. Предмет аренды",
                "Сервис предназначен для оформления заявок на временное "
                "пользование строительным и профессиональным инструментом. "
                "Состав заказа, количество оборудования и период аренды "
                "указываются пользователем при оформлении."
            ),
            (
                "2. Стоимость и расчёты",
                "Стоимость аренды рассчитывается по тарифам, отображаемым "
                "в карточках инструмента и при оформлении заказа. Итоговая "
                "сумма может учитывать применимые скидки, купоны и бонусные "
                "баллы. Обеспечительный залог рассчитывается отдельно."
            ),
            (
                "3. Заявка и подтверждение",
                "После отправки заявка поступает администратору на рассмотрение. "
                "Состав заказа и доступность отдельных позиций могут быть "
                "проверены до подтверждения. Не следует считать заявку "
                "подтверждённой до изменения её статуса в сервисе."
            ),
            (
                "4. Получение и возврат",
                "Пользователь обязан получить и вернуть оборудование "
                "согласованным способом и в согласованные сроки. "
                "Факт выдачи и возврата отдельных экземпляров фиксируется "
                "в системе учёта."
            ),
            (
                "5. Сохранность оборудования",
                "Пользователь обязан бережно обращаться с полученным "
                "оборудованием и вернуть его в согласованном состоянии. "
                "Состояние инструмента проверяется при возврате."
            ),
            (
                "6. Важное уточнение",
                "Этот раздел содержит краткое описание принципов работы сервиса "
                "и не заменяет полный договор аренды. До публикации в качестве "
                "юридически значимой публичной оферты необходимо дополнить "
                "его сведениями об арендодателе, порядке оплаты, отмены, "
                "возврата залога, ответственности сторон и иных обязательных "
                "условиях."
            ),
        ],
    },

    "privacy": {
        "title": "Политика конфиденциальности",
        "lead": (
            "На этой странице описаны основные категории данных, "
            "которые могут использоваться для работы с аккаунтом и заказами."
        ),
        "blocks": [
            (
                "Какие данные используются",
                "При работе с сервисом могут использоваться данные профиля: "
                "ФИО, логин, адрес электронной почты, номер телефона и адрес. "
                "Для корпоративного профиля дополнительно используются "
                "сведения об организации и загруженные документы."
            ),
            (
                "Для чего нужны данные",
                "Данные используются для регистрации и обслуживания аккаунта, "
                "обработки заявок, связи с пользователем, подтверждения "
                "бизнес-статуса и ведения истории заказов."
            ),
            (
                "Доступ к информации",
                "Доступ к данным профиля и корпоративным документам должен "
                "предоставляться только в объёме, необходимом для работы сервиса. "
                "Пользователь может обратиться к администратору по вопросам "
                "своих данных."
            ),
            (
                "Защита информации",
                "Не передавайте другим лицам пароль от аккаунта. "
                "Корпоративные документы и персональные сведения не должны "
                "размещаться в общедоступных каталогах сайта."
            ),
            (
                "Важное уточнение",
                "Перед использованием этой страницы в качестве полноценной "
                "политики обработки персональных данных необходимо указать "
                "оператора данных, правовые основания обработки, сроки хранения, "
                "порядок реализации прав пользователей и сведения "
                "о привлечённых обработчиках данных."
            ),
        ],
    },
}


@app.route("/info/<slug>")
def info_page(slug):
    page = INFO_PAGES.get(slug)
    if page is None:
        abort(404)

    return render_template(
        "info.html",
        page=page,
        slug=slug,
    )


@app.route("/feedback", methods=["POST"])
def send_feedback():
    nxt = safe_next("index")

    try:
        kind = request.form.get("kind", "").strip()

        if kind not in ("Ошибка", "Предложение"):
            raise ValueError("Выберите тип обращения.")

        contact = field(
            "contact",
            255,
            not bool(g.user),
            "Контакт для связи"
        )

        message = field(
            "text",
            3000,
            True,
            "Текст сообщения"
        )

        if not message or not message.strip():
            raise ValueError("Введите текст сообщения.")

        if g.user:
            sender = (
                f"Пользователь: {g.user.get('login', 'Не указан')}\n"
                f"ФИО: {g.user.get('fio', 'Не указано')}\n"
                f"Email: {g.user.get('email', 'Не указан')}"
            )
            if contact:
                sender += f"\nДополнительный контакт: {contact}"
        else:
            sender = f"Гость сайта\nКонтакт: {contact}"

        admins = admin_ids()

        if not admins:
            raise ValueError("Администратор не найден.")

        # В проекте один администратор — отправляем ему одно уведомление.
        admin_id = admins[0]

        title = f"Обратная связь: {kind}"
        body = (
            f"{sender}\n"
            f"Страница: {nxt}\n\n"
            f"Сообщение:\n{message.strip()}"
        )

        notify(
            admin_id,
            title,
            body,
            "Обратная связь",
            mail=False
        )

        flash("Спасибо! Ваше сообщение отправлено администратору.")

    except ValueError as e:
        flash(str(e))

    except Exception:
        app.logger.exception("Ошибка отправки обратной связи")
        flash("Не удалось отправить сообщение. Попробуйте позже.")

    return redirect(nxt)


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "make-admin":
        print(make_admin(sys.argv[2]))
        sys.exit(0)
    startup_sync()                      # серии артикулов и экземпляры по КоличествоНаСкладе (без дубликатов)
    if os.environ.get("WERKZEUG_RUN_MAIN") == "true" or not app.debug:
        threading.Thread(target=_due_loop, daemon=True).start()
    app.run(host='0.0.0.0', debug=True)