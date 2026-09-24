# -*- coding: utf-8 -*-
"""Тесты режимов «Один файл поставщика» и «Несколько файлов поставщика».
Запуск:  python tests/test_multi.py  (из корня проекта)"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from openpyxl import Workbook, load_workbook  # noqa: E402
import app  # noqa: E402

TMP = tempfile.mkdtemp(prefix="kompleks_test_")


def make_sup(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(["KOD", "GLAVA1", "SUM_N1", "MZADOLG1", "GLAVA2", "SUM_N2",
               "MZADOLG2"])
    for r in rows:
        ws.append(r)
    wb.save(path)


def read_out(p):
    wb = load_workbook(p)
    ws = wb.active
    hdr = [str(c.value) for c in ws[1]]
    rows = {str(r[1]): dict(zip(hdr, r))
            for r in ws.iter_rows(min_row=2, values_only=True)}
    return hdr, rows


def g(row, c):
    v = row.get(c)
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v)


def main():
    S1 = os.path.join(TMP, "sup1.xlsx")
    S2 = os.path.join(TMP, "sup2.xlsx")
    T = os.path.join(TMP, "tgt.xlsx")
    make_sup(S1, [["100", "Договор А", 150.5, "", "", 0, ""],
                  ["200", "", 300, "7", "Ложь", "", ""]])
    make_sup(S2, [["200", "Дог-Б", -50, "", "GL2", 400, ""],
                  ["300", "", 999, "", "", "", ""]])
    wb = Workbook()
    ws = wb.active
    ws.append(["RAION", "KOD", "PERIOD", "GLAVA1", "GLAVA2", "SUM_N1",
               "ZADOLG1", "MZADOLG1", "SUM_N2", "ZADOLG2", "MZADOLG2"])
    ws.append(["Р1", "100", "2026-08", "ГлавА", "", "", "", "", "", "", ""])
    ws.append(["Р1", "200", "2026-08", "ГлавВ", "", "", "", "", "", "", ""])
    ws.append(["Р1", "999", "2026-08", "", "", "5", "", "", "", "", ""])
    wb.save(T)

    out1 = os.path.join(TMP, "out1.xlsx")
    out2 = os.path.join(TMP, "out2.xlsx")

    # --- одиночный режим (регрессия прежних правил) --------------------------
    res = app.process(S1, T, out1, multi_mode=False)
    hdr, rows = read_out(out1)
    assert g(rows["100"], "DOGOVOR1") == "1" and g(rows["100"], "SUM_N1") == "150.5"
    assert g(rows["100"], "DOGOVOR2") == "0"
    assert g(rows["200"], "DOGOVOR1") == "1"      # GLAVA1 цели + сумма
    assert g(rows["200"], "DOGOVOR2") == "0"      # GLAVA2 из sup не используется
    assert g(rows["200"], "ZADOLG1") == "1" and g(rows["200"], "MZADOLG1") == "7"
    assert all(g(rows["999"], f) == "0"
               for f in ("SUM_N1", "ZADOLG1", "MZADOLG1", "SUM_N2"))
    assert g(rows["100"], "GLAVA1") == "ГлавА"    # GLAVA сохранена
    print("SINGLE MODE OK")

    # --- мультирежим: два файла поставщика -----------------------------------
    res2 = app.process([S1, S2], T, out2, multi_mode=True)
    hdr2, rows2 = read_out(out2)
    r100, r200, r999 = rows2["100"], rows2["200"], rows2["999"]
    assert g(r100, "DOGOVOR1") == "1" and g(r100, "SUM_N1") == "150.5"
    assert g(r100, "DOGOVOR2") == "0"             # SUM_N2 = 0
    assert g(r100, "GLAVA1") == "" and g(r100, "GLAVA2") == ""   # очищена
    # дубль KOD 200 между файлами: sup2 имеет max сумму 400 > 300 —
    # переносится ЦЕЛИКОМ запись sup2 (SUM_N1 её пуст -> 0)
    assert g(r200, "SUM_N2") == "400" and g(r200, "SUM_N1") == "0", r200
    assert g(r200, "DOGOVOR2") == "1"             # есть сумма -> 1 без GLAVA
    assert g(r200, "DOGOVOR1") == "0"
    assert g(r200, "GLAVA1") == "" and g(r200, "GLAVA2") == ""
    assert g(r999, "DOGOVOR1") == "0" and g(r999, "SUM_N1") == "0"
    assert res2["dup_kods"] == ["200"] and res2["multi_mode"] is True
    txt = app.build_report_text(res2)
    assert "несколько файлов данных от поставщика" in txt and "sup1.xlsx" in txt and "sup2.xlsx" in txt
    assert "GLAVA<N> очищено" in txt
    assert "Определять тариф по площади" in hdr2
    print("MULTI MODE OK")
    print("ALL TESTS PASSED")


if __name__ == "__main__":
    main()
