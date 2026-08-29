#!/usr/bin/env python3
"""Backend дашборда 'Колбасное производство Азизии' (проект №3).

ХРАНЕНИЕ: Google Таблица (Excel на Google Диске).
  Листы: Заказы / Магазины / Поставщики / Сырьё / Оборудование
  Все данные дашборда живут в таблице. Новые заявки из дашборда
  сохраняются туда же. state.json НЕ используется (БД = таблица).

Flask:
  GET  /            -> дашборд
  GET  /api/state   -> состояние (из таблицы)
  POST /api/orders  -> новая заявка (пишется в лист Заказы)
  POST /api/orders/<id>/stage -> смена этапа
  POST /api/orders/<id>/note  -> заметка
  POST /api/raw/<name>/stock  -> обновить остаток сырья
"""
import os, json, uuid, datetime, time
from flask import Flask, request, jsonify, send_from_directory
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
import gspread

BASE = os.path.dirname(os.path.abspath(__file__))
TOKEN = os.path.expanduser("~/.hermes/google_token.json")
SCOPES = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "19TkAQDxrfaF8auyZsIiJFmSI-wxS9fRT8TPX1_r2XBE"

STAGES = ["заявка", "принято", "производство", "готово", "отгрузка", "закрыто"]
PRODUCTS = ["Варёная в/с (халяль)", "Копчёная (халяль)", "Сырокопчёная (халяль)", "Сосиски (халяль)", "Сардельки (халяль)",
            "Пельмени (халяль)", "Паштет (халяль)", "Колбаса детская (халяль)"]

app = Flask(__name__, static_folder=None)

# --- кэш чтения из Google Sheets (чтобы не упереться в квоту 429) ---
_CACHE = {"data": None, "ts": 0, "TTL": 30}  # сек


def creds():
    c = Credentials.from_authorized_user_file(TOKEN, SCOPES)
    if c.expired and c.refresh_token:
        c.refresh(Request())
    return c


_gc = None


def gc():
    global _gc
    if _gc is None:
        _gc = gspread.authorize(creds())
    return _gc


def sh():
    return gc().open_by_key(SPREADSHEET_ID)


def rows(ws_name):
    w = sh().worksheet(ws_name)
    data = w.get_all_values()
    if not data:
        return []
    header = data[0]
    return [dict(zip(header, r)) for r in data[1:] if any(x.strip() for x in r)]


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def next_id():
    existing = [o["id"] for o in rows("Заказы") if o.get("id")]
    n = len(existing) + 1
    return "ord-%04d" % n


# -------- STATE из таблицы (с кэшем) --------
def build_state():
    orders = {o["id"]: o for o in rows("Заказы")}
    mags = rows("Магазины")
    sups = rows("Поставщики")
    raw = [{"name": r["название"], "stock": int(r["остаток_кг"] or 0), "min": int(r["мин_кг"] or 0)}
           for r in rows("Сырьё")]
    equip = [{"name": e["название"], "status": e["статус"], "health": int(e["здоровье_%"] or 0)}
             for e in rows("Оборудование")]
    return {"orders": orders, "stages": STAGES, "products": PRODUCTS,
            "raw": raw, "equip": equip,
            "shops": mags, "suppliers": sups, "sheet_id": SPREADSHEET_ID}


def get_state():
    t = time.time()
    # свежий кэш — отдаём сразу
    if _CACHE["data"] is not None and t - _CACHE["ts"] < _CACHE["TTL"]:
        return _CACHE["data"]
    # нужно обновить — пытаемся прочитать таблицу, но ловим ошибки (429 и т.п.)
    try:
        data = build_state()
        _CACHE["data"] = data
        _CACHE["ts"] = t
        return data
    except Exception as e:
        # не уперлись в квоту/сеть: отдаём последний успешный кэш (даже просроченный)
        if _CACHE["data"] is not None:
            return _CACHE["data"]
        # совсем нет данных — вернём минимальный каркас, чтобы дашборд не упал
        return {"orders": {}, "stages": STAGES, "products": PRODUCTS, "raw": [], "equip": [],
                "shops": [], "suppliers": [], "sheet_id": SPREADSHEET_ID, "_stale": True}


def invalidate():
    # не стираем данные, только помечаем кэш просроченным -> при след. чтении обновим,
    # но старый data остаётся как fallback при ошибке 429
    _CACHE["ts"] = 0


@app.route("/")
def index():
    return send_from_directory(os.path.join(BASE, "static"), "index.html")


@app.route("/api/state")
def state():
    return jsonify(get_state())


@app.route("/api/orders", methods=["POST"])
def new_order():
    b = request.get_json(force=True, silent=True) or {}
    oid = next_id()
    cust = b.get("customer", "—")
    # найдём id магазина по названию (если есть)
    mag = next((m for m in rows("Магазины") if m["название"] == cust), None)
    cid = mag["id"] if mag else ""
    rec = [oid, "заявка", b.get("product", PRODUCTS[0]), int(b.get("weight", 50)),
           cid, cust, b.get("source", "заявка"), b.get("priority", "P2"),
           b.get("manager", "Акмаль"), b.get("note", ""), now()]
    sh().worksheet("Заказы").append_row(rec)
    invalidate()
    return jsonify(dict(zip(
        ["id", "этап", "продукт", "вес_кг", "клиент_id", "клиент", "источник", "приоритет", "менеджер", "заметка", "создан"], rec))), 201


def find_row(ws_name, col, val):
    w = sh().worksheet(ws_name)
    data = w.get_all_values()
    header = data[0]
    for i, r in enumerate(data[1:], start=2):
        if r and r[header.index(col)] == val:
            return w, i, header
    return w, None, header


@app.route("/api/orders/<oid>/stage", methods=["POST"])
def set_stage(oid):
    b = request.get_json(force=True, silent=True) or {}
    w, i, header = find_row("Заказы", "id", oid)
    if not i:
        return jsonify({"error": "no order"}), 404
    ns = b.get("stage")
    if ns in STAGES:
        w.update_cell(i, header.index("этап") + 1, ns)
    invalidate()
    o = dict(zip(header, w.row_values(i)))
    return jsonify(o)


@app.route("/api/orders/<oid>/note", methods=["POST"])
def set_note(oid):
    b = request.get_json(force=True, silent=True) or {}
    w, i, header = find_row("Заказы", "id", oid)
    if not i:
        return jsonify({"error": "no order"}), 404
    w.update_cell(i, header.index("заметка") + 1, b.get("note", ""))
    invalidate()
    return jsonify(dict(zip(header, w.row_values(i))))


@app.route("/api/raw/<name>/stock", methods=["POST"])
def set_raw(name):
    b = request.get_json(force=True, silent=True) or {}
    w, i, header = find_row("Сырьё", "название", name)
    if not i:
        return jsonify({"error": "no raw"}), 404
    w.update_cell(i, header.index("остаток_кг") + 1, int(b.get("stock", 0)))
    invalidate()
    return jsonify(dict(zip(header, w.row_values(i))))


if __name__ == "__main__":
    print("🌭 Колбасное производство Азизии (Google Sheets) на http://localhost:5002")
    port = int(os.environ.get("PORT", 5002))
    app.run(host="0.0.0.0", port=port, debug=False)
