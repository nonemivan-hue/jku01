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
3. DOGOVOR1..DOGOVOR16: если в файле для загрузки GLAVA<N> (текстовое)
   имеет значение и SUM_N<N> после переноса > 0 (не пусто и не 0;
   отрицательная сумма переносится как 0) -> 1, иначе 0.
4. В файл загрузки добавляется столбец "Определять тариф по площади" = 1.
5. Если KOD отсутствует в файле поставщика -> в SUM_N/ZADOLG/MZADOLG
   ставятся 0. Если столбца KOD нет в файле для загрузки -> сообщение
   "Код семьи не найден (нет столбца KOD в файле для загрузки)".
6. Если KOD повторяется в файле поставщика несколько раз -> переносится
   ЦЕЛИКОМ запись с НАИБОЛЬШИМ значением SUM_N (сравнение максимальных
   сумм по всем блокам 1..16; при равенстве — первая из записей).
7. GLAVA1..GLAVA16 имеют текстовый формат (значения из файла загрузки
   сохраняются как есть).
8. Если в MZADOLG1..MZADOLG16 есть какая-либо информация (не пусто и не 0),
   то ZADOLG<N> = 1, иначе 0.
9. Поля RAION, KOD, ID_FIAS, ID_KLADR, PUNKT, STREET, HOUSE, KORP, FLAT,
   KOM, PERIOD имеют текстовый формат.
10. Если SUM_N<N> в файле поставщика отрицательное -> переносится 0.
11. После обработки предлагается скачать (сохранить) результат; имя файла
    соответствует имени файла для загрузки.
12. Вкладка «Несколько файлов поставщика»: загружается несколько файлов
    данных от поставщика (можно сразу несколько) и один файл для загрузки
    в Комплекс. Логика та же, кроме: DOGOVOR1..16 = 1 только при наличии
    данных в SUM_N1..16 (без проверки GLAVA), а поле GLAVA<N> очищается.

Запуск: python app.py  (графический интерфейс, Windows)
"""

import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

def _ensure_deps():
    """Автоматически устанавливает отсутствующие зависимости (xlrd,
    openpyxl) через pip. Вызывается до импорта этих модулей. Работает и при
    запуске из .exe, собранного PyInstaller. dbfread не требуется — для DBF
    есть встроенный читатель."""
    required = {"xlrd": "xlrd", "openpyxl": "openpyxl"}
    missing = []
    for module, pip_name in required.items():
        try:
            __import__(module)
        except ImportError:
            missing.append(pip_name)
    if not missing:
        return
    # Не пытаемся ставить пакеты внутри GUI-экзешника без Python — подскажем.
    if getattr(sys, "frozen", False):
        _fatal_missing_modules(missing)
    print("Установка недостающих зависимостей: %s ..." % ", ".join(missing))
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--yes"] + missing)
    except Exception:
        # не удалось с --yes (старые версии pip) — пробуем без флага
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing)
        except Exception as e:
            _fatal_missing_modules(missing, str(e))
    # Обновляем sys.path на случай установки в каталог пользователя (--user)
    import site
    try:
        user_site = site.getusersitepackages()
        if user_site and user_site not in sys.path:
            sys.path.append(user_site)
    except Exception:
        pass
    still = [m for m in missing if not _can_import(m)]
    if still:
        _fatal_missing_modules(still)


def _can_import(module):
    try:
        __import__(module)
        return True
    except ImportError:
        return False


def _fatal_missing_modules(modules, error=None):
    msg = ("Не найдены модули: %s.\n"
           "Установите их командой:\n"
           "    python -m pip install %s\n"
           "(если несколько версий Python — проверьте, каким python.exe "
           "вы запускаете программу)" % (", ".join(modules), " ".join(modules)))
    if error:
        msg += "\n\nТекст ошибки: %s" % error
    try:
        import tkinter as _tk
        from tkinter import messagebox as _mb
        root = _tk.Tk()
        root.withdraw()
        _mb.showerror("Отсутствуют зависимости", msg)
        root.destroy()
    except Exception:
        print(msg)
    raise SystemExit(msg)


_ensure_deps()

import xlrd
from openpyxl import Workbook, load_workbook

try:
    import dbfread
except ImportError:          # необязательно: есть встроенный читатель DBF
    dbfread = None

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


def _block_index(col):
    """'SUM_N3'/'MZADOLG12'/'ZADOLG4' -> 3/12/4, иначе None."""
    m = re.match(r"^(?:SUM_N|MZADOLG|ZADOLG)(\d+)$", col)
    return int(m.group(1)) if m else None


def pick_max_block(rows):
    """Из списка строк поставщика выбирает «наибольшую» по SUM_N.

    Сравниваются суммы во всех блоках SUM_N1..SUM_N16: берётся строка,
    у которой максимальная из сумм наибольшая (при равенстве — первая).
    Если ни в одной строке нет заполненных SUM_N — возвращается первая.
    """
    best, best_key = None, None
    for r in rows:
        nums = [to_number(v) for c, v in r.items()
                if c.startswith("SUM_N") and _block_index(c)]
        nums = [n for n in nums if n is not None]
        key = max(nums) if nums else None
        if best is None:
            best, best_key = r, key
        elif key is not None and (best_key is None or key > best_key):
            best, best_key = r, key
    return best


# ---------------------------------------------------------------------------
# Чтение файлов поставщика
# ---------------------------------------------------------------------------

def _read_dbf_builtin(path):
    """Встроенный читатель dBASE III/IV без внешних зависимостей.
    Возвращает (rows:[dict], header:[str]) — значения в виде строк."""
    with open(path, "rb") as fh:
        data = fh.read()
    if len(data) < 32 or data[0] not in (0x03, 0x30, 0x31, 0xF5, 0xFB, 0x83):
        raise RuntimeError("Файл не является DBF (dBASE).")
    num_records = int.from_bytes(data[4:8], "little")
    header_len = int.from_bytes(data[8:10], "little")
    record_len = int.from_bytes(data[10:12], "little")
    blk_size = int.from_bytes(data[28:30], "little") or 512
    fields = []
    off = 32
    while off + 32 <= header_len and data[off] != 0x0D:
        ftype = chr(data[off + 11])
        flen = data[off + 16]
        name = data[off:off + 11].split(b"\x00")[0]\
            .decode("ascii", "ignore").strip()
        fields.append((name, ftype, flen))
        off += 32
    has_memo = any(ft in ("M", "P") for _n, ft, _l in fields)
    memo_path = None
    mdata = b""
    if has_memo:
        memo_ext = ".dbt" if data[0] in (0x03, 0x83, 0xF5) else ".fpt"
        memo_path = os.path.join(
            os.path.dirname(os.path.abspath(path)),
            os.path.splitext(os.path.basename(path))[0] + memo_ext)
        if os.path.exists(memo_path):
            with open(memo_path, "rb") as mf:
                mdata = mf.read()

    def read_memo(raw):
        try:
            blknum = int(raw.decode("ascii", "ignore").strip() or 0)
        except ValueError:
            return ""
        if blknum <= 0 or not mdata or blknum * blk_size >= len(mdata):
            return ""
        chunk = mdata[blknum * blk_size:]
        end = chunk.find(b"\x1a\x1a")
        if end == -1:
            end = chunk.find(b"\x00")
        return chunk[:end if end != -1 else len(chunk)]

    def decode_text(b):
        for enc in ("cp866", "cp1251", "utf-8", "latin1"):
            try:
                return b.decode(enc)
            except UnicodeDecodeError:
                continue
        return b.decode("latin1")

    rows = []
    for i in range(num_records):
        rec_start = header_len + i * record_len
        if rec_start + record_len > len(data):
            break
        if data[rec_start:rec_start + 1] == b"*":   # помечена на удаление
            continue
        pos = rec_start + 1
        row = {}
        for name, ftype, flen in fields:
            raw = data[pos:pos + flen]
            pos += flen
            if ftype in ("C", "V"):
                val = decode_text(raw).strip()
            elif ftype in ("N", "F"):
                s = decode_text(raw).strip().replace(",", ".")
                if s == "":
                    val = ""
                else:
                    try:
                        num = float(s)
                        val = str(int(num)) if num == int(num) else repr(num)
                    except ValueError:
                        val = s
            elif ftype == "D":
                s = decode_text(raw).strip()
                val = "%s.%s.%s" % (s[4:6], s[2:4], s[0:2]) if len(s) == 8 else s
            elif ftype == "L":
                c = chr(raw[0]).upper() if raw else "?"
                val = "1" if c in "TY" else ("0" if c in "FN" else "")
            elif ftype in ("M", "P"):
                val = decode_text(read_memo(raw)).strip()
            else:
                val = decode_text(raw).strip()
            row[norm_name(name)] = val
        rows.append(row)
    header = [norm_name(n) for n, _t, _l in fields]
    return rows, header


def read_dbf(path):
    """Чтение DBF -> список словарей {СТОЛБЕЦ: str}, порядок сохраняется.
    Используется библиотека dbfread (если установлена), иначе — встроенный
    читатель, не требующий зависимостей."""
    if dbfread is not None:
        try:
            table = dbfread.DBF(path, encoding="cp866")
            cols = [norm_name(f.name) for f in table.fields]
            rows = [{norm_name(k): cell_to_str(v) for k, v in rec.items()}
                    for rec in table]
            return rows, cols
        except UnicodeDecodeError:
            pass                       # попробуем другие кодировки ниже
        except Exception as e:
            raise RuntimeError("Не удалось прочитать DBF-файл: %s" % e)
        last_err = None
        for enc in ("cp1251", "utf-8", "latin1"):
            try:
                table = dbfread.DBF(path, encoding=enc)
                cols = [norm_name(f.name) for f in table.fields]
                rows = [{norm_name(k): cell_to_str(v) for k, v in rec.items()}
                        for rec in table]
                return rows, cols
            except Exception as e:
                last_err = e
        try:
            return _read_dbf_builtin(path)
        except Exception:
            raise RuntimeError("Не удалось прочитать DBF-файл: %s" % last_err)
    return _read_dbf_builtin(path)


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

def process(supplier_paths, target_path, out_path, multi_mode=False):
    """Выполняет обработку, пишет результат в out_path. Возвращает сводку.

    supplier_paths — путь к файлу поставщика или список путей (несколько
    файлов). multi_mode=True (вкладка «Несколько файлов поставщика»):
    DOGOVOR<N>=1 только при наличии данных в SUM_N<N> (без проверки GLAVA),
    поле GLAVA<N> в результате очищается.
    """
    if isinstance(supplier_paths, str):
        supplier_paths = [supplier_paths]
    supplier_paths = [p for p in supplier_paths if p]
    if not supplier_paths:
        raise RuntimeError("Не выбран файл данных от поставщика.")

    tgt_rows, tgt_header = read_tabular(target_path)

    warnings = []

    # --- проверка наличия KOD ------------------------------------------------
    if "KOD" not in tgt_header:
        raise RuntimeError(
            "Код семьи не найден: в файле для загрузки отсутствует "
            "столбец KOD.")

    ID_COLS = {"RAION", "KOD", "ID_FIAS", "ID_KLADR", "PUNKT", "STREET",
               "HOUSE", "KORP", "FLAT", "KOM", "PERIOD"}

    # --- чтение всех файлов поставщика ---------------------------------------
    # Каждый файл имеет СВОЙ набор переносимых столбцов: данные переносятся
    # только из тех столбцов, которые реально есть в данном файле (иначе
    # пустые значения из одного файла затёрли бы данные из другого).
    all_sup_rows = []
    missing_cols = []          # новые столбцы для файла загрузки (по порядку)
    for sp in supplier_paths:
        rows, header = read_tabular(sp)
        if "KOD" not in header:
            raise RuntimeError(
                "Код семьи не найден: в файле поставщика «%s» отсутствует "
                "столбец KOD." % os.path.basename(sp))
        src = os.path.basename(sp)
        new_for_this = [c for c in header
                        if c not in ID_COLS and c not in tgt_header
                        and c not in missing_cols]
        missing_cols.extend(new_for_this)
        for r in rows:
            r["__src__"] = src
            r["__cols__"] = {c for c in header if c not in ID_COLS}
        all_sup_rows.extend(rows)

    columns = list(tgt_header) + missing_cols

    # гарантируем наличие столбцов DOGOVOR1..16 (могли быть в файле загрузки,
    # но не войти в расчётный список; новые — добавляем сразу за блоком SUM_N)
    def _ensure_after(col, anchor):
        if col in columns:
            return
        if anchor in columns:
            columns.insert(columns.index(anchor) + 1, col)
        else:
            columns.append(col)
    for i in range(1, MAX_BLOCKS + 1):
        _ensure_after("DOGOVOR%d" % i, "SUM_N%d" % i)

    if missing_cols:
        warnings.append("В файл загрузки добавлены столбцы из файла "
                        "поставщика: %s" % ", ".join(missing_cols))

    # --- индекс поставщика по KOD (при дублях — запись с наибольшей SUM_N) ---
    sup_groups = {}
    for row in all_sup_rows:
        kod = cell_to_str(row.get("KOD"))
        if kod == "":
            continue
        sup_groups.setdefault(kod, []).append(row)
    dup_kods = {k for k, v in sup_groups.items() if len(v) > 1}
    sup_index = {k: (v[0] if len(v) == 1 else pick_max_block(v))
                 for k, v in sup_groups.items()}
    if dup_kods:
        src_note = ("среди всех файлов поставщика" if len(supplier_paths) > 1
                    else "в файле поставщика")
        warnings.append("KOD повторяется %s: %s — "
                        "взята запись с наибольшей суммой (SUM_N)."
                        % (src_note, ", ".join(sorted(dup_kods))))

    # --- итоговый набор столбцов --------------------------------------------
    new_col = "Определять тариф по площади"
    if new_col.upper() not in [c.upper() for c in columns]:
        columns.append(new_col)

    # --- формирование результата ---------------------------------------------
    result = []
    matched = 0
    unmatched = []
    report_rows = []  # построчная информация для отчёта о выполнении
    for row in tgt_rows:
        if all(cell_to_str(v) == "" for v in row.values()):
            continue
        kod = cell_to_str(row.get("KOD"))
        sup = sup_index.get(kod)
        out = dict(row)
        neg_fixed = 0
        if sup is None:
            unmatched.append(kod)
            # данных поставщика нет -> во все блоки ставим 0
            for i in range(1, MAX_BLOCKS + 1):
                out["SUM_N%d" % i] = "0"
                out["ZADOLG%d" % i] = "0"
                out["MZADOLG%d" % i] = "0"
        else:
            matched += 1
            sup_cols = sup.get("__cols__") or set()

            # перенос блоков SUM_N/ZADOLG/MZADOLG — только если блок есть
            # в данном файле поставщика (иначе оставляем то, что уже есть
            # в строке файла загрузки; для пустого блока DOGOVOR станет 0)
            for i in range(1, MAX_BLOCKS + 1):
                blk = ("SUM_N%d" % i, "MZADOLG%d" % i, "ZADOLG%d" % i)
                if not any(b in sup_cols for b in blk):
                    continue
                s_sum = sup.get("SUM_N%d" % i)
                s_mz = sup.get("MZADOLG%d" % i)

                # перенос числовых данных из файла поставщика (пустое -> 0)
                # отрицательная сумма переносится как 0
                sv = num_or_zero(s_sum)
                if isinstance(sv, (int, float)) and sv < 0:
                    sv = 0
                    neg_fixed += 1
                out["SUM_N%d" % i] = fmt_num(sv)
                out["MZADOLG%d" % i] = fmt_num(num_or_zero(s_mz))

                # ZADOLG: есть информация в MZADOLG -> 1, иначе 0
                out["ZADOLG%d" % i] = "1" if has_info(s_mz) else "0"

            # прочие столбцы поставщика, отсутствовавшие в файле загрузки
            for c in missing_cols:
                if c not in sup_cols:
                    continue
                if c.startswith("SUM_N"):
                    mv = num_or_zero(sup.get(c))
                    if isinstance(mv, (int, float)) and mv < 0:
                        mv = 0
                    out[c] = fmt_num(mv)
                elif c.startswith(("ZADOLG", "MZADOLG")):
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

        # DOGOVOR1..16.
        # Обычный режим: если в GLAVA<N> (файл для загрузки, текстовое) ЕСТЬ
        # значение и в SUM_N<N> (после переноса) есть сумма -> 1, иначе 0.
        # Режим «Несколько файлов поставщика»: DOGOVOR<N>=1 только при
        # наличии данных в SUM_N<N> (GLAVA не проверяется), поле GLAVA<N>
        # очищается. Считается ПОСЛЕ переноса данных.
        for i in range(1, MAX_BLOCKS + 1):
            if multi_mode:
                ok = has_info(out.get("SUM_N%d" % i))
                out["GLAVA%d" % i] = ""
            else:
                ok = (has_info(out.get("GLAVA%d" % i))
                      and has_info(out.get("SUM_N%d" % i)))
            out["DOGOVOR%d" % i] = "1" if ok else "0"

        out[new_col] = "1"
        result.append(out)

        # строка отчёта о выполнении по данной семье
        filled_blocks = sum(1 for i in range(1, MAX_BLOCKS + 1)
                            if has_info(out.get("SUM_N%d" % i)))
        dogovor_cnt = sum(1 for i in range(1, MAX_BLOCKS + 1)
                          if out.get("DOGOVOR%d" % i) == "1")
        zadolg_cnt = sum(1 for i in range(1, MAX_BLOCKS + 1)
                         if out.get("ZADOLG%d" % i) == "1")
        report_rows.append({
            "kod": kod,
            "status": ("совпал с файлом поставщика" if sup is not None
                       else "код семьи не найден в файле поставщика — "
                            "проставлены 0"),
            "blocks": filled_blocks,
            "dogovor": dogovor_cnt,
            "zadolg": zadolg_cnt,
            "neg": neg_fixed,
        })

    if unmatched:
        warnings.append("Код семьи не найден в файле поставщика (проставлены 0): "
                        "%s" % ", ".join(unmatched))

    # --- запись результата ----------------------------------------------------
    out_path = write_output(out_path, columns, result, warnings)

    summary = {
        "out_path": out_path,
        "rows": len(result),
        "matched": matched,
        "unmatched": unmatched,
        "warnings": warnings,
        "supplier_paths": [os.path.abspath(p) for p in supplier_paths],
        "target_path": target_path,
        "report_rows": report_rows,
        "dup_kods": sorted(dup_kods),
        "missing_cols": missing_cols,
        "new_col": new_col,
        "multi_mode": multi_mode,
    }
    return summary


def build_report_text(summary):
    """Формирует текст отчёта о выполнении (для скачивания в .txt)."""
    lines = []
    add = lines.append
    add("ОТЧЁТ О ВЫПОЛНЕНИИ ОБРАБОТКИ ДАННЫХ")
    if summary.get("multi_mode"):
        add("Режим: несколько файлов данных от поставщика "
            "(DOGOVOR<N>=1 при наличии SUM_N<N>, GLAVA<N> очищается)")
    add("=" * 60)
    add("Дата/время формирования: %s"
        % datetime.now().strftime("%d.%m.%Y %H:%M:%S"))
    add("")
    sup_paths = summary.get("supplier_paths") or [summary.get("supplier_path")]
    if len(sup_paths) > 1:
        add("Файлы данных от поставщика (%d):" % len(sup_paths))
        for i, p in enumerate(sup_paths, 1):
            add("    %d. %s" % (i, p))
    else:
        add("Файл данных от поставщика:   %s" % sup_paths[0])
    add("Файл для загрузки в Комплекс: %s" % summary["target_path"])
    add("Файл результата:              %s" % summary["out_path"])
    add("")
    add("-" * 60)
    add("ИТОГИ")
    add("-" * 60)
    add("Всего строк обработано:            %d" % summary["rows"])
    add("Совпадений по KOD с поставщиком:   %d" % summary["matched"])
    add("KOD не найдено у поставщика:       %d%s"
        % (len(summary["unmatched"]),
           " (%s)" % ", ".join(summary["unmatched"]) if summary["unmatched"] else ""))
    add("Дублирующихся KOD у поставщика:    %d%s"
        % (len(summary["dup_kods"]),
           " (%s) — взята запись с наибольшей суммой" % ", ".join(summary["dup_kods"])
           if summary["dup_kods"] else ""))
    add("Столбец «%s»: добавлен, значение 1 во всех строках."
        % summary["new_col"])
    if summary["missing_cols"]:
        add("Добавлены столбцы из файла поставщика: %s"
            % ", ".join(summary["missing_cols"]))
    add("")
    add("-" * 60)
    add("ПОСТРОЧНО (по KOD семьи)")
    add("-" * 60)
    for i, r in enumerate(summary["report_rows"], 1):
        add("%d. KOD %s — %s; блоков с суммой: %d из 16; "
            "DOGOVOR=1: %d; ZADOLG=1: %d; отрицательных SUM_N заменено на 0: %d"
            % (i, r["kod"], r["status"], r["blocks"], r["dogovor"],
               r["zadolg"], r["neg"]))
    add("")
    add("-" * 60)
    add("ПРЕДУПРЕЖДЕНИЯ")
    add("-" * 60)
    if summary["warnings"]:
        for w in summary["warnings"]:
            add("⚠ " + w)
    else:
        add("Предупреждений нет.")
    add("")
    add("Правила обработки:")
    add("— SUM_N/ZADOLG/MZADOLG 1..16 переносятся из файла поставщика "
        "по KOD; при отсутствии KOD — 0;")
    add("— при дублях KOD переносится целиком запись с наибольшим SUM_N;")
    add("— отрицательная сумма SUM_N переносится как 0;")
    if summary.get("multi_mode"):
        add("— DOGOVOR<N>=1, если SUM_N<N> содержит сумму (без проверки "
            "GLAVA<N>), иначе 0; поле GLAVA<N> очищено;")
    else:
        add("— DOGOVOR<N>=1, если GLAVA<N> заполнена и SUM_N<N> содержит "
            "сумму, иначе 0;")
    add("— есть данные в MZADOLG<N> -> ZADOLG<N>=1, иначе 0;")
    add("— столбец «%s» = 1." % summary["new_col"])
    return "\r\n".join(lines) + "\r\n"


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
        self.geometry("760x520")
        self.resizable(False, False)

        self.supplier_var = tk.StringVar()      # путь (одна вкладка)
        self.target_var = tk.StringVar()        # путь (вкладка 1)
        self.multi_suppliers = []               # список путей (вкладка 2)
        self.multi_target_var = tk.StringVar()  # путь (вкладка 2)
        self.status_var = tk.StringVar(
            value="Шаг 1. Выберите файл данных от поставщика (.dbf или Excel).")
        self.multi_status_var = tk.StringVar(
            value="Можно выбрать сразу несколько файлов поставщика.")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        # ---------------- Вкладка 1: один файл поставщика -------------------
        frm = ttk.Frame(nb, padding=12)
        nb.add(frm, text="  Один файл поставщика  ")

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
                                  command=lambda: self.start_process(False))
        self.run_btn.grid(row=4, column=0, sticky="w", pady=(4, 10))

        self.report_btn = ttk.Button(frm, text="Скачать отчёт о выполнении (.txt)",
                                     command=self.save_report, state="disabled")
        self.report_btn.grid(row=4, column=1, sticky="we", pady=(4, 10))

        ttk.Label(frm, textvariable=self.status_var, foreground="#004085",
                  wraplength=640, justify="left").grid(
            row=5, column=0, columnspan=2, sticky="w")

        self.progress = ttk.Progressbar(frm, mode="indeterminate", length=640)
        self.progress.grid(row=6, column=0, columnspan=2, sticky="we", pady=(10, 0))

        frm.columnconfigure(0, weight=1)

        # ------- Вкладка 2: несколько файлов поставщика ---------------------
        frm2 = ttk.Frame(nb, padding=12)
        nb.add(frm2, text="  Несколько файлов поставщика  ")

        ttk.Label(frm2, text="1. Данные от поставщика (можно выбрать "
                             "сразу несколько файлов, DBF / Excel):").grid(
            row=0, column=0, columnspan=2, sticky="w")

        lbfrm = ttk.Frame(frm2)
        lbfrm.grid(row=1, column=0, sticky="we", pady=(2, 4))
        self.multi_list = tk.Listbox(lbfrm, width=64, height=7,
                                     exportselection=False)
        sb = ttk.Scrollbar(lbfrm, command=self.multi_list.yview)
        self.multi_list.config(yscrollcommand=sb.set)
        self.multi_list.pack(side="left", fill="both")
        sb.pack(side="right", fill="y")

        btns = ttk.Frame(frm2)
        btns.grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Button(btns, text="Добавить файлы…",
                   command=self.add_multi_suppliers).pack(side="left", padx=(0, 6))
        ttk.Button(btns, text="Удалить выделенные",
                   command=self.remove_multi_suppliers).pack(side="left", padx=6)
        ttk.Button(btns, text="Очистить список",
                   command=self.clear_multi_suppliers).pack(side="left", padx=6)

        ttk.Label(frm2, text="2. Данные для загрузки в Комплекс (один Excel-файл):").grid(
            row=3, column=0, sticky="w")
        ttk.Entry(frm2, textvariable=self.multi_target_var, width=62).grid(
            row=4, column=0, sticky="we", pady=(2, 12))
        ttk.Button(frm2, text="Обзор…",
                   command=self.pick_multi_target).grid(row=4, column=1, padx=6)

        self.multi_run_btn = ttk.Button(
            frm2, text="Обработать данные",
            command=lambda: self.start_process(True))
        self.multi_run_btn.grid(row=5, column=0, sticky="w", pady=(4, 10))

        self.multi_report_btn = ttk.Button(
            frm2, text="Скачать отчёт о выполнении (.txt)",
            command=self.save_report, state="disabled")
        self.multi_report_btn.grid(row=5, column=1, sticky="we", pady=(4, 10))

        ttk.Label(frm2, textvariable=self.multi_status_var,
                  foreground="#004085", wraplength=660,
                  justify="left").grid(
            row=6, column=0, columnspan=2, sticky="w")

        self.multi_progress = ttk.Progressbar(frm2, mode="indeterminate",
                                              length=640)
        self.multi_progress.grid(row=7, column=0, columnspan=2, sticky="we",
                                 pady=(10, 0))

        frm2.columnconfigure(0, weight=1)

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

    # ---- вкладка «Несколько файлов поставщика» ------------------------------
    def add_multi_suppliers(self):
        paths = filedialog.askopenfilenames(
            title="Выберите файлы данных поставщика (можно несколько)",
            filetypes=self.SUPPLIER_TYPES)
        added = 0
        for p in paths:
            if p and p not in self.multi_suppliers:
                self.multi_suppliers.append(p)
                self.multi_list.insert("end", p)
                added += 1
        self.multi_status_var.set(
            ("Добавлено файлов: %d. " % added if added else "") +
            "Всего в списке: %d." % len(self.multi_suppliers))

    def remove_multi_suppliers(self):
        sel = list(self.multi_list.curselection())
        for idx in reversed(sel):
            self.multi_list.delete(idx)
            del self.multi_suppliers[idx]
        self.multi_status_var.set("Всего в списке: %d."
                                  % len(self.multi_suppliers))

    def clear_multi_suppliers(self):
        self.multi_list.delete(0, "end")
        self.multi_suppliers = []
        self.multi_status_var.set("Список очищен.")

    def pick_multi_target(self):
        path = filedialog.askopenfilename(
            title="Выберите файл для загрузки в Комплекс",
            filetypes=self.TARGET_TYPES)
        if path:
            self.multi_target_var.set(path)
            self.multi_status_var.set("Файл для загрузки выбран. "
                                      "Нажмите «Обработать данные».")

    # ---- запуск обработки ---------------------------------------------------
    def start_process(self, multi_mode):
        if self._busy:
            return
        if multi_mode:
            sups = list(self.multi_suppliers)
            tgt = self.multi_target_var.get().strip()
            status = self.multi_status_var
            run_btn, progress = self.multi_run_btn, self.multi_progress
            if not sups:
                messagebox.showerror(
                    "Ошибка",
                    "Не добавлено ни одного файла данных от поставщика.")
                return
            missing = [p for p in sups if not os.path.isfile(p)]
            if missing:
                messagebox.showerror(
                    "Ошибка",
                    "Файл поставщика не найден:\n%s" % "\n".join(missing))
                return
        else:
            s = self.supplier_var.get().strip()
            sups = [s] if s else []
            tgt = self.target_var.get().strip()
            status = self.status_var
            run_btn, progress = self.run_btn, self.progress
            if not sups or not os.path.isfile(s):
                messagebox.showerror("Ошибка", "Не выбран файл данных от поставщика.")
                return
        if not tgt or not os.path.isfile(tgt):
            messagebox.showerror("Ошибка", "Не выбран файл для загрузки в Комплекс.")
            return

        self._busy = True
        self._multi_mode = multi_mode
        self._active_btn = run_btn
        self._active_progress = progress
        run_btn.config(state="disabled")
        progress.start(12)
        status.set("Идёт обработка…")
        threading.Thread(target=self._work, args=(sups, tgt, multi_mode),
                         daemon=True).start()

    def _work(self, sups, tgt, multi_mode):
        try:
            out_dir = os.path.dirname(tgt) or os.getcwd()
            base = os.path.basename(tgt)
            root_, ext_ = os.path.splitext(base)
            suffix = "_обработанный_неск" if multi_mode else "_обработанный"
            out_path = os.path.join(out_dir, root_ + suffix + ext_)
            res = process(sups, tgt, out_path, multi_mode=multi_mode)
            self.after(0, self._done, res)
        except Exception as e:
            err = "".join(traceback.format_exception_only(type(e), e)).strip()
            self.after(0, self._fail, err)

    def _done(self, res):
        self._busy = False
        self._active_progress.stop()
        self._active_btn.config(state="normal")
        multi = bool(res.get("multi_mode"))
        self._result = res
        self.report_btn.config(state="normal")
        self.multi_report_btn.config(state="normal")
        # отчёт о выполнении сохраняется рядом с файлом результата
        report_path = os.path.splitext(res["out_path"])[0] + "_отчёт.txt"
        try:
            with open(report_path, "w", encoding="utf-8-sig", newline="") as f:
                f.write(build_report_text(res))
            self._report_path = report_path
        except OSError:
            self._report_path = None
        msg = ("Обработка завершена. Строк: %d, совпадений по KOD: %d."
               % (res["rows"], res["matched"]))
        if multi:
            msg += " Режим «Несколько файлов поставщика»: DOGOVOR по SUM_N, GLAVA очищена."
        for w in res["warnings"]:
            msg += "\n⚠ " + w
        status = self.multi_status_var if multi else self.status_var
        status.set(msg)
        ask = ("Данные обработаны.\n\n%s\n\n"
               "Скачать (сохранить) файл результатов?" % msg)
        if messagebox.askyesno("Готово", ask):
            self._save_and_offer(res["out_path"])
        if self._report_path and messagebox.askyesno(
                "Отчёт о выполнении",
                "Скачать отчёт о выполнении в формате .txt?"):
            self.save_report()

    def _fail(self, err):
        self._busy = False
        getattr(self, "_active_progress", self.progress).stop()
        getattr(self, "_active_btn", self.run_btn).config(state="normal")
        self.status_var.set("Ошибка обработки.")
        messagebox.showerror("Ошибка", err)

    # ---- скачивание отчёта о выполнении -------------------------------------
    def save_report(self):
        res = getattr(self, "_result", None)
        if not res:
            messagebox.showinfo("Отчёт",
                                "Сначала выполните обработку данных.")
            return
        text = build_report_text(res)
        default_name = os.path.splitext(os.path.basename(res["out_path"]))[0] \
            + "_отчёт.txt"
        path = filedialog.asksaveasfilename(
            title="Скачать отчёт о выполнении (.txt)",
            initialfile=default_name,
            defaultextension=".txt",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                f.write(text)
            self.status_var.set("Отчёт сохранён: %s" % path)
            if os.name == "nt":
                try:
                    os.startfile(os.path.dirname(path))
                except Exception:
                    pass
            messagebox.showinfo("Сохранено", "Отчёт сохранён:\n%s" % path)
        except OSError as e:
            messagebox.showerror("Ошибка",
                                 "Не удалось сохранить отчёт:\n%s" % e)

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
