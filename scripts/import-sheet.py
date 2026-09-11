#!/usr/bin/env python3
"""Импорт дневника питания из Google-таблицы в data/days/.

Вход — JSON-результат Google Drive MCP (download_file_content с экспортом в xlsx),
который инструмент сохраняет на диск. Скрипт сам достаёт base64, разбирает книгу
и печатает короткий отчёт: что изменилось. Ничего объёмного наружу не отдаёт —
чтобы вызывающему агенту не приходилось читать таблицу глазами.

Зависимостей нет, только стандартная библиотека.

    python3 scripts/import-sheet.py <путь-к-результату-download_file_content>
"""
import base64, hashlib, json, os, re, sys, zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DAYS = os.path.join(ROOT, "data", "days")
FOODS = os.path.join(ROOT, "data", "foods.json")
STATE = os.path.join(ROOT, "data", "sheet-state.json")

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
MEALS = ("Завтрак", "Обед", "Ужин", "Перекусы")
MONTHS = {m: i + 1 for i, m in enumerate(
    ["январь", "февраль", "март", "апрель", "май", "июнь",
     "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"])}
DAY_RE = re.compile(r"^\s*(\d{1,2})\s+(" + "|".join(MONTHS) + r")\s+(\d{4})", re.I)

TRANSLIT = {"а":"a","б":"b","в":"v","г":"g","д":"d","е":"e","ё":"e","ж":"zh","з":"z","и":"i",
            "й":"y","к":"k","л":"l","м":"m","н":"n","о":"o","п":"p","р":"r","с":"s","т":"t",
            "у":"u","ф":"f","х":"h","ц":"c","ч":"ch","ш":"sh","щ":"sch","ъ":"","ы":"y","ь":"",
            "э":"e","ю":"yu","я":"ya"}


def slug(name):
    out = "".join(TRANSLIT.get(ch, ch) for ch in name.lower())
    out = re.sub(r"[^a-z0-9]+", "_", out).strip("_")
    return out or "food"


def num(v):
    if v is None:
        return None
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return None


def open_book(path):
    """Достаёт xlsx из JSON-результата MCP, из голого base64 или из самого .xlsx."""
    with open(path, "rb") as fh:
        head = fh.read(1)
    if head == b"PK":
        return zipfile.ZipFile(path)
    raw = open(path, encoding="utf-8", errors="replace").read().strip()
    if raw.startswith("{"):
        raw = json.loads(raw)["content"]
    return zipfile.ZipFile(__import__("io").BytesIO(base64.b64decode(raw)))


def sheet_rows(z, shared, path):
    for row in ET.fromstring(z.read(path)).iter("{%s}row" % NS["m"]):
        cells = {}
        for cell in row.findall("m:c", NS):
            v = cell.find("m:v", NS)
            if v is None or v.text is None:
                continue
            val = shared[int(v.text)] if cell.get("t") == "s" else v.text
            cells[re.match(r"[A-Z]+", cell.get("r")).group(0)] = val
        if cells:
            yield cells


def load_sheets(z):
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        shared = ["".join(t.text or "" for t in si.iter("{%s}t" % NS["m"]))
                  for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS)]
    rels = {r.get("Id"): r.get("Target").lstrip("/")
            for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    out = {}
    for sh in ET.fromstring(z.read("xl/workbook.xml")).find("m:sheets", NS):
        tgt = rels[sh.get("{%s}id" % NS["r"])]
        out[sh.get("name")] = tgt if tgt.startswith("xl/") else "xl/" + tgt
    return shared, out


def parse_days(z, shared, sheets):
    days = {}
    for name, path in sheets.items():
        if not re.match(r"^(" + "|".join(MONTHS) + r")\s+\d{4}$", name.strip(), re.I):
            continue
        date = None
        for c in sheet_rows(z, shared, path):
            m = DAY_RE.match(str(c.get("A", "")))
            if m:
                date = "%s-%02d-%02d" % (m.group(3), MONTHS[m.group(2).lower()], int(m.group(1)))
                continue
            if date is None or c.get("A") not in MEALS:
                continue
            product = (c.get("B") or "").strip()
            if not product:
                continue
            grams, k = num(c.get("C")), num(c.get("D"))
            if not grams or k is None:
                continue
            item = {"name": product, "g": round(grams, 1), "k": round(k, 1),
                    "p": round(num(c.get("E")) or 0, 1), "f": round(num(c.get("F")) or 0, 1),
                    "c": round(num(c.get("G")) or 0, 1)}
            day = days.setdefault(date, {})
            day.setdefault(c["A"], []).append(item)
    return days


def parse_foods(z, shared, sheets):
    path = next((p for n, p in sheets.items() if n.strip().lower() == "база продуктов"), None)
    if not path:
        return []
    out = []
    for c in sheet_rows(z, shared, path):
        name = (c.get("A") or "").strip()
        k = num(c.get("B"))
        if not name or k is None or name.lower() == "продукт":
            continue
        out.append({"id": slug(name), "name": name, "cat": "Из таблицы",
                    "per100": {"k": round(k, 1), "p": round(num(c.get("C")) or 0, 1),
                               "f": round(num(c.get("D")) or 0, 1), "c": round(num(c.get("E")) or 0, 1)},
                    "piece": None, "src": "sheet"})
    return out


def write_days(days):
    os.makedirs(DAYS, exist_ok=True)
    state = json.load(open(STATE, encoding="utf-8")) if os.path.exists(STATE) else {}
    old = state.get("days", {})
    new, added, changed = {}, [], []
    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    for date, meals in sorted(days.items()):
        doc = {"date": date, "source": "google-sheets", "importedAt": stamp,
               "meals": [{"name": m, "items": meals[m]} for m in MEALS if meals.get(m)]}
        body = json.dumps({"date": date, "meals": doc["meals"]}, ensure_ascii=False, sort_keys=True)
        digest = hashlib.sha1(body.encode("utf-8")).hexdigest()[:12]
        new[date] = digest
        path = os.path.join(DAYS, date + ".json")
        if old.get(date) == digest and os.path.exists(path):
            continue
        (added if date not in old else changed).append(date)
        json.dump(doc, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    removed = []
    for date in old:
        if date not in new:
            path = os.path.join(DAYS, date + ".json")
            if os.path.exists(path):
                os.remove(path)
            removed.append(date)

    json.dump({"days": new, "syncedAt": stamp}, open(STATE, "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    return added, changed, removed


def merge_foods(sheet_foods):
    doc = json.load(open(FOODS, encoding="utf-8"))
    by_name = {f["name"].strip().lower(): f for f in doc["foods"]}
    added = 0
    for f in sheet_foods:
        prev = by_name.get(f["name"].strip().lower())
        if prev:
            prev["per100"] = f["per100"]      # значения из таблицы главнее
            prev["src"] = "sheet"
        else:
            by_name[f["name"].strip().lower()] = f
            added += 1
    doc["foods"] = sorted(by_name.values(), key=lambda x: (x.get("src") != "sheet", x["name"]))
    doc["updated"] = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    json.dump(doc, open(FOODS, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return added, len(sheet_foods)


def main():
    if len(sys.argv) < 2:
        sys.exit("Нужен путь к результату download_file_content (xlsx-экспорт таблицы)")
    z = open_book(sys.argv[1])
    shared, sheets = load_sheets(z)
    days = parse_days(z, shared, sheets)
    added, changed, removed = write_days(days)
    new_foods, total_foods = merge_foods(parse_foods(z, shared, sheets))

    print("дней с записями в таблице: %d%s" % (
        len(days), (" (%s … %s)" % (min(days), max(days))) if days else ""))
    print("новых: %s" % (", ".join(added) or "нет"))
    print("изменилось: %s" % (", ".join(changed) or "нет"))
    print("удалено: %s" % (", ".join(removed) or "нет"))
    print("база продуктов из таблицы: %d строк, новых в репозитории: %d" % (total_foods, new_foods))
    print("ИТОГ: %s" % ("есть изменения" if (added or changed or removed) else "без изменений"))


if __name__ == "__main__":
    main()
