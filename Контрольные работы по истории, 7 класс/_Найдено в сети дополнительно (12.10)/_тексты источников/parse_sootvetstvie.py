# -*- coding: utf-8 -*-
"""Разбор заданий «на соответствие» из листа «Работы» в структуру:
левая часть (А/Б/В/Г или 1/2/3), правая часть (1/2/3 или а/б/в), правильные пары.
Результат — JSON + печать для проверки глазами."""
import re, json, sys
from openpyxl import load_workbook

DIR = "/home/user/School/Контрольные тестовые работы/"
wb = load_workbook(DIR + "Контрольные работы 7 класс.xlsx")
ws = wb["Работы"]
ROWS = list(ws.iter_rows(min_row=2, values_only=True))

LAT2CYR = {"A": "А", "B": "Б", "C": "В", "D": "Г", "E": "Д"}
CYR = "АБВГДЕ"
MARKER = re.compile(r'(?:(?<![А-Яа-яЁёA-Za-z0-9])((?:[АБВГДЕабвгде]|[A-E])\s*[\)\.]\s|\d{1,2}\s*[\)\.]\s))')

def norm_label(m):
    m = m.strip()
    ch = m[0]
    if ch in LAT2CYR: ch = LAT2CYR[ch]
    return ch

def tokenize(text):
    """возвращает список (метка, текст) по маркерам"""
    parts = []
    ms = list(MARKER.finditer(text))
    for i, m in enumerate(ms):
        end = ms[i+1].start() if i+1 < len(ms) else len(text)
        label = norm_label(m.group(1))
        body = text[m.end():end].strip().strip("|").strip()
        body = re.sub(r'\s*\|\s*$', '', body).strip()
        parts.append((label, body))
    head = text[:ms[0].start()].strip() if ms else text
    return head, parts

def is_letter(l): return l in CYR or l in "абвгде"
def is_digit(l): return l.isdigit()

def parse_task(question, answer):
    head, parts = tokenize(question)
    if not parts: return None
    first = parts[0][0]
    left_is_letter = is_letter(first)
    left, right = [], []
    for label, body in parts:
        side = "L" if (is_letter(label) == left_is_letter) else "R"
        (left if side == "L" else right).append((label, body))
    return head, left, right

def parse_answer(answer, left, right):
    """нормализуем ответ в пары (метка левая -> метка правая)"""
    a = (answer or "").strip()
    pairs = {}
    # формат «А2 Б4 В1 Г3» / «А2, Б4»
    found = re.findall(r'([АБВГДE]|\d)\s*[—\-–]\s*([АБВГДабвгде]|\d)', a)
    if found and len(found) >= 2:
        for l, r in found: pairs[l] = r
        return pairs, "dash"
    # формат «А2 Б4 В1 Г3» без разделителей
    found = re.findall(r'([АБВГД])\s*(\d)', a)
    if len(found) >= 2:
        for l, r in found: pairs[l] = r
        return pairs, "letterdigit"
    # формат «4312» — строка цифр по числу левых элементов
    if re.fullmatch(r'\d{2,6}', a):
        lefts = [l for l, _ in left]
        if len(a) == len(lefts):
            for l, r in zip(lefts, a): pairs[l] = r
            return pairs, "digits"
    # формат «1АВ 2БГ» — цифра + набор букв
    found = re.findall(r'(\d)\s*([АБВГабвг]{1,4})', a)
    if found:
        for l, r in found: pairs[l] = r
        return pairs, "digitletters"
    # формат «1 — в; 2 — д» ловится выше; «1в 2д»
    found = re.findall(r'(\d)\s*[\)\.]?\s*([абвгде])', a)
    if found:
        for l, r in found: pairs[l] = r
        return pairs, "digitlettersmall"
    return {}, "unknown"

out = []
for r in ROWS:
    q = str(r[5] or "")
    if not (all(o is None for o in r[6:12]) and re.search(r'соответстви|Соотнесите', q)):
        continue
    p = parse_task(q, r[12])
    if not p:
        print("!! НЕ РАЗОБРАНО:", r[0]); continue
    head, left, right = p
    pairs, how = parse_answer(str(r[12] or ""), left, right)
    out.append(dict(num=r[0], razdel=r[1], work=r[2], var=r[3], task=r[4], head=head,
                    left=left, right=right, answer=r[12], ball=r[13], note=r[17],
                    pairs=pairs, how=how))

print("РАЗОБРАНО:", len(out))
for t in out:
    print("="*100)
    print("№%s | %s | вар.%s | зад.%s | формат ответа: %s" % (t["num"], str(t["work"])[:40], t["var"], t["task"], t["how"]))
    print("ВОПРОС:", t["head"][:120])
    print("ЛЕВО (%d):" % len(t["left"]), " || ".join("%s) %s" % (l, b[:55]) for l, b in t["left"]))
    print("ПРАВО (%d):" % len(t["right"]), " || ".join("%s) %s" % (l, b[:55]) for l, b in t["right"]))
    print("ОТВЕТ В ИСТОЧНИКЕ:", t["answer"], "| ПАРЫ:", t["pairs"])

json.dump(out, open(DIR + "_Найдено в сети дополнительно (12.10)/_тексты источников/sootvetstviya.json", "w"),
          ensure_ascii=False, indent=1, default=str)
print("\nJSON сохранён.")
