# -*- coding: utf-8 -*-
"""Полная пересборка «Контрольные работы 7 класс.xlsx»:
лист «Работы» (385 заданий, без заданий на соответствие) + лист «Соответствия» (37 заданий).
Ссылки и авторы проставляются заново по названию работы.
"""
import json, urllib.parse
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

DIR = "/home/user/School/Контрольные тестовые работы/"
PATH = DIR + "Контрольные работы 7 класс.xlsx"
SRC = "/tmp/backup4.xlsx"          # состояние ДО удаления строк: лист «Работы» = 422 строки
T = json.load(open(DIR + "_Найдено в сети дополнительно (12.10)/_тексты источников/sootvetstviya.json"))

S_BASE = "https://school105.ru/storage_drive/svedeniya-ob-oo/osnovnye-svedeniya/obrazovanie/om/5-9/История/"
def s105(f): return S_BASE + f
A_G = "Галсанов Ринчин-Доржи Пурбоевич (nsportal.ru)"
A_N = "Никитина Людмила Александровна (nsportal.ru)"
A_0 = "автор не указан (school105.ru)"
L_G = "https://nsportal.ru/vuz/istoricheskie-nauki/library/2026/04/09/testy-po-istorii-rossii-7-klass"
L_N = "https://nsportal.ru/shkola/istoriya/library/2019/10/09/test-po-istorii-rossii"

MAP = [
 ("Итоговый тест «История России", A_G, L_G),
 ("Тест «Россия в XVI в.»", A_G, L_G),
 ("Тест «Смута в России»", A_G, L_G),
 ("Тест «Россия при первых Романовых»", A_G, L_G),
 ("Тест «Россия в XVII веке»", A_0, s105("Контрольная работа по истории России 7 класс.pdf")),
 ("Итоговая контрольная работа по Новой истории", A_0, s105("Контрольная работа по истории 7 класс новая история.pdf")),
 ("Тест «Россия на рубеже XVI–XVII веков»", A_N, L_N),
 ("Контрольная работа по истории России XVI–XVII вв.", A_0, s105("ОМ история 7 класс.pdf")),
 ("Итоговая работа по всеобщей истории", A_0, s105("ОМ история 7 класс.pdf")),
 ("Входная контрольная работа, 7 класс", A_0, s105("Входная контрольная работа по истории 7 класс.pdf")),
 ("Входная контрольная работа, 8 класс", A_0, s105("Входная контрольная работа по истории 8 класс.pdf")),
]
def meta(work):
    w = str(work)
    for k, a, l in MAP:
        if k in w: return a, l
    raise SystemExit("нет метаданных для: " + w)

# ---- 1. читаем исходные 422 строки листа «Работы» (из бэкапа) ----
from openpyxl import load_workbook
wb0 = load_workbook(SRC)
ws0 = wb0["Работы"]
rows = [[c.value for c in r[:18]] for r in ws0.iter_rows(min_row=2)]
# выбрасываем строки соответствий (по JSON) и пустые
nums = {int(t["num"]) for t in T}
works, skipped = [], 0
for i, r in enumerate(rows, start=1):
    if i in nums or r[2] is None:
        skipped += 1; continue
    works.append(r)
print("перенесено в «Работы»:", len(works), "| исключено:", skipped)

def clean_body(s):
    s = (s or "").strip().strip("|").strip().rstrip(";").strip()
    return s

def answer_text(t):
    parts = []
    for l, _ in t["left"]:
        v = t["pairs"].get(l)
        if v is None: continue
        if len(v) > 1 and all(ch in [x for x, _ in t["right"]] for ch in v):
            v = ", ".join(v)
        parts.append("%s → %s" % (l, v))
    return "; ".join(parts)

# ---- 2. собираем новую книгу ----
hfont = Font(name="Calibri", size=11, bold=True)
hfill = PatternFill("solid", fgColor="DDEBF7")
dfont = Font(name="Calibri", size=11)
dalign = Alignment(wrap_text=True, vertical="top")
lfont = Font(name="Calibri", size=11, color="0563C1", underline="single")

wb = Workbook(); wb.remove(wb.active)
W, M = "Работы", "Соответствия"

# ----- лист «Работы» -----
ws = wb.create_sheet(W)
H = ["№","Раздел","Работа","Вариант","№ задания","Вопрос","Вариант 1","Вариант 2","Вариант 3","Вариант 4",
     "Вариант 5","Вариант 6","Правильный ответ","Баллы","Время на ответ, мин","Автор","Ссылка на источник","Примечание"]
WID = [5,14,36,8,9,58,22,22,22,22,22,22,30,7,9,30,62,32]
for j,h in enumerate(H,1):
    c=ws.cell(row=1,column=j,value=h); c.font=hfont; c.fill=hfill
    c.alignment=Alignment(vertical="top",wrap_text=True)
    ws.column_dimensions[get_column_letter(j)].width=WID[j-1]
for i,r in enumerate(works,start=2):
    author,link = meta(r[2])
    out = [i-1]+r[1:6]+list(r[6:10])+[None,None]+r[12:15]+[author,link,r[17]]
    # r: [№,Раздел,Работа,Вариант,№з,Q,V1..V4,Ответ,Баллы,Время,Автор,Ссылка,Примечание]
    out = [i-1, r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11],
           r[12], r[13], r[14], author, link, r[17]]
    for j,v in enumerate(out,1):
        c=ws.cell(row=i,column=j,value=v); c.font=dfont; c.alignment=dalign
    c=ws.cell(row=i,column=17); c.hyperlink=urllib.parse.quote(str(link),safe=":/?&=%"); c.font=lfont
ws.freeze_panes="G2"; ws.auto_filter.ref="A1:R%d"%ws.max_row

# ----- лист «Соответствия» -----
sm = wb.create_sheet(M)
HM = ["№","Раздел","Работа","Вариант","№ задания","Вопрос",
      "Левая часть 1","Левая часть 2","Левая часть 3","Левая часть 4","Левая часть 5",
      "Правая часть 1","Правая часть 2","Правая часть 3","Правая часть 4","Правая часть 5",
      "Правильный ответ","Ответ как в источнике","Баллы","Автор","Ссылка на источник","Примечание"]
WIDM = [5,14,34,8,9,46,28,28,28,28,28,28,28,28,28,28,30,22,7,30,62,32]
for j,h in enumerate(HM,1):
    c=sm.cell(row=1,column=j,value=h); c.font=hfont; c.fill=hfill
    c.alignment=Alignment(vertical="top",wrap_text=True)
    sm.column_dimensions[get_column_letter(j)].width=WIDM[j-1]
for i,t in enumerate(sorted(T,key=lambda x:x["num"]),start=2):
    author,link = meta(t["work"])
    out=[i-1,t["razdel"],t["work"],t["var"],t["task"],t["head"]]
    for n in range(5):
        out.append(("%s) %s"%(t["left"][n][0],clean_body(t["left"][n][1]))) if n<len(t["left"]) else None)
    for n in range(5):
        out.append(("%s) %s"%(t["right"][n][0],clean_body(t["right"][n][1]))) if n<len(t["right"]) else None)
    out += [answer_text(t), t["answer"], t["ball"], author, link, t["note"]]
    for j,v in enumerate(out,1):
        c=sm.cell(row=i,column=j,value=v); c.font=dfont; c.alignment=dalign
    c=sm.cell(row=i,column=21); c.hyperlink=urllib.parse.quote(str(link),safe=":/?&=%"); c.font=lfont
sm.freeze_panes="G2"; sm.auto_filter.ref="A1:V%d"%sm.max_row

wb.save(PATH)
print("готово:", wb.sheetnames, "| «Работы»:", ws.max_row-1, "| «Соответствия»:", sm.max_row-1)
