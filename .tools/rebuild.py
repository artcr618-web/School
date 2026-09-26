#!/usr/bin/env python3
"""Пересобирает 'Каталог документов.md' из CSV + карты разобранного (done.json).
Карта хранится рядом со скриптом, поэтому переживает перезапуск песочницы."""
import csv, json, os, urllib.parse

BASE = "/home/user/Документы школы 56"
HERE = os.path.dirname(os.path.abspath(__file__))
DONE = os.path.join(HERE, "done.json")

MARK = {"полностью": "✅", "разобрано": "✅", "частично": "⚠️", "реквизиты": "⚠️",
        "заглушка": "❗", "битая": "❌"}
LABEL = {"полностью": "полностью", "разобрано": "разобрано", "частично": "частично",
         "реквизиты": "реквизиты", "заглушка": "ЗАГЛУШКА", "битая": "ссылка не работает"}

def link(path):
    return urllib.parse.quote(path)

def main():
    done = json.load(open(DONE, encoding="utf-8"))
    rows = list(csv.DictReader(open(os.path.join(BASE, "Каталог документов.csv"),
                                    encoding="utf-8-sig")))
    cats, order = {}, []
    for r in rows:
        c = r["Категория"]
        if c not in cats:
            cats[c] = []; order.append(c)
        cats[c].append(r)

    out = ["# Каталог документов МБОУ «СШ № 56» (г. Иваново)", "",
           "Всего файлов: **%d**. Источник — [раздел «Документы» на сайте школы]"
           "(https://school56.gosuslugi.ru/svedeniya-ob-obrazovatelnoy-organizatsii/dokumenty/), "
           "срез на 25.09.2026." % len(rows), "",
           "**Как пользоваться таблицей.** В столбце «Файл на сайте» — прямая ссылка: по клику "
           "документ открывается или скачивается с сайта школы. В столбце «Текст у нас» — ссылка "
           "на текстовую версию в этой папке, если документ уже разобран. Прочерк означает, что "
           "до документа ещё не дошли.", "",
           "Обозначения: ✅ текст извлечён · ⚠️ извлечена часть · ❗ на сайте вместо документа "
           "заглушка · ❌ ссылка не работает · — не разбирался", "",
           "Разобрано на сегодня: **%d из %d**." % (len(done), len(rows)), "", "---", ""]

    for c in order:
        out.append("\n## %s (%d)\n" % (c, len(cats[c])))
        out.append("| Документ | Файл на сайте | Текст у нас |")
        out.append("|---|---|---|")
        for r in sorted(cats[c], key=lambda x: x["Название"].lower()):
            fn = r["Имя файла"]
            if fn in done:
                st, tgt = done[fn]
                cell = "%s [%s](%s)" % (MARK[st], LABEL[st], link(tgt))
            else:
                cell = "—"
            out.append("| %s | [открыть](%s) | %s |" % (r["Название"], r["Ссылка"], cell))

    open(os.path.join(BASE, "Каталог документов.md"), "w", encoding="utf-8").write("\n".join(out) + "\n")
    print("готово, записей: %d | разобрано: %d" % (len(rows), len(done)))

if __name__ == "__main__":
    main()
