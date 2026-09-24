# -*- coding: utf-8 -*-
"""
Обработчик данных поставщика для загрузки в Комплекс.

Логика (по требованиям):
1. Загружаются два файла:
   - файл данных от поставщика (.dbf или Excel .xls/.xlsx);
   - файл для загрузки в комплекс (Excel .xls/.xlsx).
2. Из файла поставщика в файл загрузки переносятся значения столбцов
   SUM_N1..SUM_N16, ZADOLG1..ZADOLG16, MZADOLG1..MZADOLG16
   (соединение строк по столбцу KOD).
3. DOGOVOR1..DOGOVOR16: если в файле поставщика для этого номера блока
   заполнены и GLAVA<N>, и SUM_N<N> -> 1, иначе 0.
4. В файл загрузки добавляется столбец "Определять тариф по площади" = 1.
5. Если KOD отсутствует в файле поставщика -> в SUM_N/ZADOLG/MZADOLG
   ставятся 0. Если столбца KOD нет в файле для загрузки -> сообщение
   "Код семьи не найден (нет столбца KOD в файле для загрузки)".
6. Если KOD повторяется в файле поставщика несколько раз -> берётся
   НАИБОЛЬШЕЕ значение SUM_N по каждому блоку (для ZADOLG/MZADOLG тоже
   берётся максимум).
7. GLAVA1..GLAVA16 имеют текстовый формат (значения из файла загрузки
   сохраняются как есть).
8. Если в MZADOLG1..MZADOLG16 есть какая-либо информация (не пусто и не 0),
   то ZADOLG<N> = 1, иначе 0.
9. Поля RAION, KOD, ID_FIAS, ID_KLADR, PUNKT, STREET, HOUSE, KORP, FLAT,
   KOM, PERIOD имеют текстовый формат.
10. После обработки предлагается скачать (сохранить) результат; имя файла
    соответствует имени файла для загрузки.

Запуск: python app.py  (графический интерфейс, Windows)
"""

import os
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import dbfread
import xlrd
from openpyxl import Workbook, load_workbook

MAX_BLOCKS = 16          # блоки 1..16
TEXT_FIELDS = {          # поля, которые всегда сохраняются в текстовом формате
    "RAION", "KOD", "ID_FIAS", "ID_KLADR", "PUNKT", "STREET",
    "HOUSE", "KORP", "FLAT", "KOM", "PERIOD",
} | {"GLAVA%d" % i for i in range(1, MAX_BLOCKS + 1)}


# ---------------------------------------------------------------------------
# Утилиты нормализации
# ---------------------------------------------------------------------------

def norm_name(name):
    """Нормализация имени столбца: верхний регистр, без пробелов."""
    return str(name).strip().upper() if name is not None else ""


def cell_to_str(v):
    """Любое значение ячейки -> строка без «хвостов» .0 у целых чисел."""
    if v is None:
        return ""
    if isinstance(v, float):
        if v == int(v):
            return str(int(v))
        return repr(v)
    return str(v).strip()


def to_number(v):
    """Попытка привести значение к float; None если не число."""
    s = cell_to_str(v).replace(",", ".")
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def fmt_num(n):
    """Число -> строка без лишнего '.0' (для записи в Excel)."""
    if n is None:
        return ""
    if isinstance(n, float) and n == int(n):
        return str(int(n))
    return str(n)


def has_info(v):
    """Есть ли в ячейке 'какая-либо информация' (не пусто и не ноль)."""
    s = cell_to_str(v)
    if s == "":
        return False
    n = to_number(s)
    if n is not None:
        return n != 0
    return True


def num_or_zero(v):
    """Значение для переноса в SUM_N/ZADOLG/MZADOLG: число или 0."""
    n = to_number(v)
    return 0 if n is None else n


# ---------------------------------------------------------------------------
# Чтение файлов поставщика
# ---------------------------------------------------------------------------

def read_dbf(path):
    """Чтение DBF -> список словарей {СТОЛБЕЦ: str}, порядок сохраняется."""
    last_err = None
    for enc in ("cp866", "cp1251", "utf-8", "latin1"):
        try:
            table = dbfread.DBF(path, encoding=enc)
            rows = []
            cols = [norm_name(f.name) for f in table.fields]
            for rec in table:
                rows.append({norm_name(k): cell_to_str(v) for k, v in rec.items()})
            return rows, cols
        except Exception as e:  # пробуем следующую кодировку
            last_err = e
    raise RuntimeError("Не удалось прочитать DBF-файл: %s" % last_err)


def _rows_from_xlrd(path):
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    if sh.nrows == 0:
        return [], []
    header = [norm_name(sh.cell_value(0, c)) for c in range(sh.ncols)]
    rows = []
    for r in range(1, sh.nrows):
        row = {}
        for c in range(sh.ncols):
            val = sh.cell_value(r, c)
            if sh.cell_type(r, c) == xlrd.XL_CELL_NUMBER:
                val = cell_to_str(val)
            row[header[c]] = cell_to_str(val)
        rows.append(row)
    return rows, header


def _rows_from_openpyxl(path):
    wb = load_workbook(path, data_only=True, read_only=True)
    sh = wb.worksheets[0]
    it = sh.iter_rows(values_only=True)
    try:
        header_row = next(it)
    except StopIteration:
        return [], []
    header = [norm_name(h) for h in header_row]
    rows = []
    for values in it:
        if values is None or all(v is None for v in values):
            continue
        row = {}
        for c, h in enumerate(header):
            v = values[c] if c < len(values) else None
            row[h] = cell_to_str(v)
        rows.append(row)
    wb.close()
    return rows, header


def read_excel_any(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        return _rows_from_xlrd(path)
    return _rows_from_openpyxl(path)


def read_tabular(path):
    """Чтение таблицы (DBF/Excel) -> (rows:[dict], header:[str])."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".dbf":
        return read_dbf(path)
    if ext in (".xls", ".xlsx", ".xlsm"):
        return read_excel_any(path)
    raise RuntimeError("Неподдерживаемый формат файла: %s\n"
                       "Можно загрузить .dbf, .xls или .xlsx" % ext)


# ---------------------------------------------------------------------------
# Основная обработка
# ---------------------------------------------------------------------------

def process(supplier_path, target_path, out_path):
    """Выполняет обработку, пишет результат в out_path. Возвращает сводку."""
    sup_rows, sup_header = read_tabular(supplier_path)
    tgt_rows, tgt_header = read_tabular(target_path)

    warnings = []

    # --- проверка наличия KOD ------------------------------------------------
    if "KOD" not in tgt_header:
        raise RuntimeError(
            "Код семьи не найден: в файле для загрузки отсутствует "
            "столбец KOD.")
    if "KOD" not in sup_header:
        raise RuntimeError(
            "Код семьи не найден: в файле поставщика отсутствует столбец KOD.")

    # --- столбцы файла поставщика, которых нет в файле загрузки --------------
    sup_data_cols = {c for c in sup_header
                     if c not in ("RAION", "KOD", "ID_FIAS", "ID_KLADR",
                                  "PUNKT", "STREET", "HOUSE", "KORP",
                                  "FLAT", "KOM", "PERIOD")}
    missing_cols = [c for c in sup_data_cols if c not in tgt_header]
    columns = list(tgt_header) + missing_cols
    if missing_cols:
        warnings.append("В файл загрузки добавлены столбцы из файла "
                        "поставщика: %s" % ", ".join(missing_cols))

    # --- индекс поставщика по KOD (максимумы при дублях) ---------------------
    sup_index = {}
    dup_kods = set()
    for row in sup_rows:
        kod = cell_to_str(row.get("KOD"))
        if kod == "":
            continue
        if kod in sup_index:
            dup_kods.add(kod)
        cur = sup_index.get(kod)
        merged = dict(cur) if cur else {}
        for key, val in row.items():
            if key.startswith(("SUM_N", "ZADOLG", "MZADOLG")):
                a, b = to_number(merged.get(key)), to_number(val)
                if a is None:
                    merged[key] = val
                elif b is not None and b > a:
                    merged[key] = val
            elif key.startswith("GLAVA"):
                if not has_info(merged.get(key)) and has_info(val):
                    merged[key] = val
            else:
                merged.setdefault(key, val)
        sup_index[kod] = merged
    if dup_kods:
        warnings.append("В файле поставщика KOD повторяется: %s — "
                        "взяты наибольшие значения."
                        % ", ".join(sorted(dup_kods)))

    # --- итоговый набор столбцов --------------------------------------------
    new_col = "Определять тариф по площади"
    if new_col.upper() not in [c.upper() for c in columns]:
        columns.append(new_col)

    # --- формирование результата ---------------------------------------------
    result = []
    matched = 0
    unmatched = []
    for row in tgt_rows:
        if all(cell_to_str(v) == "" for v in row.values()):
            continue
        kod = cell_to_str(row.get("KOD"))
        sup = sup_index.get(kod)
        out = dict(row)
        if sup is None:
            unmatched.append(kod)
            # данных поставщика нет -> во все блоки ставим 0
            for i in range(1, MAX_BLOCKS + 1):
                out["SUM_N%d" % i] = "0"
                out["ZADOLG%d" % i] = "0"
                out["MZADOLG%d" % i] = "0"
                out["DOGOVOR%d" % i] = "0"
        else:
            matched += 1
            for i in range(1, MAX_BLOCKS + 1):
                s_sum = sup.get("SUM_N%d" % i)
                s_mz = sup.get("MZADOLG%d" % i)
                s_glava = sup.get("GLAVA%d" % i)

                # перенос числовых данных из файла поставщика (пустое -> 0)
                out["SUM_N%d" % i] = fmt_num(num_or_zero(s_sum))
                out["MZADOLG%d" % i] = fmt_num(num_or_zero(s_mz))

                # DOGOVOR: GLAVA и SUM_N заполнены -> 1, иначе 0
                ok = has_info(s_glava) and has_info(s_sum)
                out["DOGOVOR%d" % i] = "1" if ok else "0"

                # ZADOLG: есть информация в MZADOLG -> 1, иначе 0
                out["ZADOLG%d" % i] = "1" if has_info(s_mz) else "0"

            # прочие столбцы поставщика, отсутствовавшие в файле загрузки
            for c in missing_cols:
                if c.startswith(("SUM_N", "ZADOLG", "MZADOLG")):
                    out[c] = fmt_num(num_or_zero(sup.get(c)))
                elif c.startswith("DOGOVOR"):
                    n = c[len("DOGOVOR"):]
                    ok = (has_info(sup.get("GLAVA%s" % n))
                          and has_info(sup.get("SUM_N%s" % n)))
                    out[c] = "1" if ok else "0"
                elif c.startswith("GLAVA"):
                    out[c] = cell_to_str(sup.get(c))
                else:
                    out[c] = cell_to_str(sup.get(c))

        out[new_col] = "1"
        result.append(out)

    if unmatched:
        warnings.append("Код семьи не найден в файле поставщика (проставлены 0): "
                        "%s" % ", ".join(unmatched))

    # --- запись результата ----------------------------------------------------
    out_path = write_output(out_path, columns, result, warnings)

    summary = {
        "out_path": out_path,
        "rows": len(result),
        "matched": matched,
        "warnings": warnings,
    }
    return summary


def write_output(out_path, columns, result, warnings):
    """Запись результата. .xls пишется через Excel COM (Windows),
    при недоступности — сохраняется .xlsx. Возвращает фактический путь."""
    ext = os.path.splitext(out_path)[1].lower()
    if ext == ".xls":
        tmp = os.path.join(tempfile.gettempdir(),
                           "conv_%d.xlsx" % os.getpid())
        _write_xlsx(tmp, columns, result)
        if _convert_xlsx_to_xls(tmp, out_path):
            try:
                os.remove(tmp)
            except OSError:
                pass
            return out_path
        alt = os.path.splitext(out_path)[0] + ".xlsx"
        shutil.move(tmp, alt)
        warnings.append("Формат .xls недоступен (не найден MS Excel), "
                        "результат сохранён как: %s" % alt)
        return alt
    if ext != ".xlsx":
        out_path += ".xlsx"
    _write_xlsx(out_path, columns, result)
    return out_path


def _write_xlsx(path, columns, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Данные"
    text_idx = {i for i, c in enumerate(columns)
                if c.upper() in TEXT_FIELDS or c == "Определять тариф по площади"}
    ws.append(list(columns))
    for row in rows:
        values = []
        for i, c in enumerate(columns):
            v = cell_to_str(row.get(c, ""))
            if i in text_idx:
                values.append(str(v))
            else:
                n = to_number(v)
                values.append(n if n is not None else str(v))
        ws.append(values)
    wb.save(path)


def _convert_xlsx_to_xls(xlsx_path, xls_path):
    """Конвертация xlsx -> xls через PowerShell + Excel COM (Windows)."""
    if os.name != "nt":
        return False
    ps = (
        "$e = New-Object -ComObject Excel.Application;"
        "$e.Visible = $false; $e.DisplayAlerts = $false;"
        "$wb = $e.Workbooks.Open('%s');"
        "$wb.SaveAs('%s', 56);"
        "$wb.Close($false); $e.Quit();"
    ) % (xlsx_path.replace("'", "''"), xls_path.replace("'", "''"))
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       check=True, capture_output=True, timeout=120)
        return os.path.exists(xls_path)
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Графический интерфейс
# ---------------------------------------------------------------------------

class App(tk.Tk):
    SUPPLIER_TYPES = [
        ("Файлы данных поставщика", "*.dbf *.DBF *.xls *.xlsx"),
        ("DBF файлы", "*.dbf *.DBF"),
        ("Файлы Excel", "*.xls *.xlsx"),
        ("Все файлы", "*.*"),
    ]
    TARGET_TYPES = [
        ("Файлы Excel", "*.xls *.xlsx"),
        ("Все файлы", "*.*"),
    ]

    def __init__(self):
        super().__init__()
        self.title("Обработчик данных поставщика для загрузки в Комплекс")
        self.geometry("720x430")
        self.resizable(False, False)

        self.supplier_var = tk.StringVar()
        self.target_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Шаг 1. Выберите файл данных от поставщика (.dbf или Excel).")

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="1. Данные от поставщика (DBF / Excel):").grid(
            row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.supplier_var, width=62).grid(
            row=1, column=0, sticky="we", pady=(2, 8))
        ttk.Button(frm, text="Обзор…", command=self.pick_supplier).grid(
            row=1, column=1, padx=6)

        ttk.Label(frm, text="2. Данные для загрузки в Комплекс (Excel):").grid(
            row=2, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.target_var, width=62).grid(
            row=3, column=0, sticky="we", pady=(2, 12))
        ttk.Button(frm, text="Обзор…", command=self.pick_target).grid(
            row=3, column=1, padx=6)

        self.run_btn = ttk.Button(frm, text="Обработать данные",
                                  command=self.start_process)
        self.run_btn.grid(row=4, column=0, sticky="w", pady=(4, 10))

        ttk.Label(frm, textvariable=self.status_var, foreground="#004085",
                  wraplength=640, justify="left").grid(
            row=5, column=0, columnspan=2, sticky="w")

        self.progress = ttk.Progressbar(frm, mode="indeterminate", length=640)
        self.progress.grid(row=6, column=0, columnspan=2, sticky="we", pady=(10, 0))

        frm.columnconfigure(0, weight=1)

        self._result = None
        self._busy = False

    # ---- выбор файлов -------------------------------------------------------
    def pick_supplier(self):
        path = filedialog.askopenfilename(title="Выберите файл данных поставщика",
                                          filetypes=self.SUPPLIER_TYPES)
        if path:
            self.supplier_var.set(path)
            self.status_var.set("Файл поставщика выбран. Шаг 2 — выберите файл для загрузки.")

    def pick_target(self):
        path = filedialog.askopenfilename(title="Выберите файл для загрузки в Комплекс",
                                          filetypes=self.TARGET_TYPES)
        if path:
            self.target_var.set(path)
            self.status_var.set("Оба файла выбраны. Нажмите «Обработать данные».")

    # ---- запуск обработки ---------------------------------------------------
    def start_process(self):
        if self._busy:
            return
        sup = self.supplier_var.get().strip()
        tgt = self.target_var.get().strip()
        if not sup or not os.path.isfile(sup):
            messagebox.showerror("Ошибка", "Не выбран файл данных от поставщика.")
            return
        if not tgt or not os.path.isfile(tgt):
            messagebox.showerror("Ошибка", "Не выбран файл для загрузки в Комплекс.")
            return

        self._busy = True
        self.run_btn.config(state="disabled")
        self.progress.pack() if False else self.progress.start(12)
        self.status_var.set("Идёт обработка…")
        threading.Thread(target=self._work, args=(sup, tgt), daemon=True).start()

    def _work(self, sup, tgt):
        try:
            out_dir = os.path.dirname(tgt) or os.getcwd()
            base = os.path.basename(tgt)
            root_, ext_ = os.path.splitext(base)
            out_path = os.path.join(out_dir, root_ + "_обработанный" + ext_)
            res = process(sup, tgt, out_path)
            self.after(0, self._done, res)
        except Exception as e:
            err = "".join(traceback.format_exception_only(type(e), e)).strip()
            self.after(0, self._fail, err)

    def _done(self, res):
        self._busy = False
        self.progress.stop()
        self.run_btn.config(state="normal")
        self._result = res
        msg = ("Обработка завершена. Строк: %d, совпадений по KOD: %d."
               % (res["rows"], res["matched"]))
        for w in res["warnings"]:
            msg += "\n⚠ " + w
        self.status_var.set(msg)
        ask = ("Данные обработаны.\n\n%s\n\n"
               "Скачать (сохранить) файл результатов?" % msg)
        if messagebox.askyesno("Готово", ask):
            self._save_and_offer(res["out_path"])

    def _fail(self, err):
        self._busy = False
        self.progress.stop()
        self.run_btn.config(state="normal")
        self.status_var.set("Ошибка обработки.")
        messagebox.showerror("Ошибка", err)

    # ---- скачивание результата ----------------------------------------------
    def _save_and_offer(self, auto_path):
        try:
            with open(auto_path, "rb") as f:
                data = f.read()
        except OSError as e:
            messagebox.showerror("Ошибка", "Не удалось прочитать результат:\n%s" % e)
            return
        default_name = os.path.basename(auto_path)
        path = filedialog.asksaveasfilename(
            title="Сохранить (скачать) файл для загрузки",
            initialfile=default_name,
            defaultextension=os.path.splitext(default_name)[1],
            filetypes=[("Файлы Excel", "*.xls *.xlsx"), ("Все файлы", "*.*")])
        if path:
            try:
                with open(path, "wb") as f:
                    f.write(data)
                self.status_var.set("Файл сохранён: %s" % path)
                if os.name == "nt":
                    try:
                        os.startfile(os.path.dirname(path))  # открыть папку
                    except Exception:
                        pass
                messagebox.showinfo("Сохранено", "Файл сохранён:\n%s" % path)
            except OSError as e:
                messagebox.showerror("Ошибка", "Не удалось сохранить файл:\n%s" % e)


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
