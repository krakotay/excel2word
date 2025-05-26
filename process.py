import tomllib
from pathlib import Path
import polars as pl
from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt
from docx.oxml.ns import qn
from pydantic import BaseModel
from time import time
import polars.selectors as sc
from humanize import intcomma

OUTPUT_DIR = Path("output")


class OfficeConfig(BaseModel):
    filename: Path
    sheet_names: list[str] | None = None

class Processor:
    word: OfficeConfig
    excel: OfficeConfig
    def __init__(self) -> None:
        with open("config.toml", "rb") as f:
            cfg = tomllib.load(f)
        self.word = OfficeConfig.model_validate(cfg["word"])
        self.excel = OfficeConfig.model_validate(cfg["excel"])
        pass
    def make_word(self, word_filename, excel_filename) -> str:
        OUTPUT_DIR.mkdir(exist_ok=True)
        timestamp = str(int(time()))
        tmp_dir = OUTPUT_DIR / timestamp
        tmp_dir.mkdir()

        # 1. конфиг
        print('filenames = ', word_filename, excel_filename)
        
        self.word.filename = Path(str(word_filename))
        self.excel.filename = Path(str(excel_filename))
        df = pl.read_excel(self.excel.filename, sheet_name=self.excel.sheet_names[0])
        df = df.with_columns(sc.by_dtype(pl.Float64).round(2))

        # 3. Открываем Word
        doc = Document(self.word.filename)
        tmp = tmp_dir / f"temp_{self.word.filename.name}"

        # 4. Ищем L6 внутри ячеек всех таблиц
        doc = insert_l6_table(doc, df)
        if not doc:
            raise RuntimeError("Маркер L6 не найден ни в одной ячейке таблиц.")
        doc.save(tmp)
        
        df = pl.read_excel(self.excel.filename, sheet_name=self.excel.sheet_names[1])
        doc = insert_k_table(doc, df)
        if not doc:
            raise RuntimeError("Маркер L6 не найден ни в одной ячейке таблиц.")
        doc.save(tmp)

        df = pl.read_excel(self.excel.filename, sheet_name=self.excel.sheet_names[2])
        doc = insert_d_table(doc, df)
        if not doc:
            raise RuntimeError("Маркер Общая сумма СПОД не найден ни в одной ячейке таблиц.")
        # 5. Сохраняем документ
        doc.save(tmp)
        tmp = tmp.relative_to('.')
        print('tmp = ',tmp, str(tmp))
        return str(tmp)


def insert_l6_table(doc: DocumentObject, df: pl.DataFrame):
    all_cells = (
        cell
        for tbl in doc.tables
        for row in tbl.rows
        for cell in row.cells
    )

    # находим первую ячейку с маркером L6
    try:
        cell = next(cell for cell in all_cells if "L6" in cell.text)
    except StopIteration:
        return None
    clear_cell_shading(cell)
    if "L6" in cell.text:
        # очистить маркер
        cell.text = ""
        # вставить вложенную таблицу
        nested = cell.add_table(rows=1, cols=len(df.columns))
        nested.style = "Table Grid"
        nested.alignment = WD_TABLE_ALIGNMENT.LEFT
        # nested.autofit = False

        # шапка

        hdr = nested.rows[0].cells
        for i, col in enumerate(df.columns):
            hdr[i].text = str(col).strip()
        # данные
        for data_row in df.iter_rows():
            new_cells = nested.add_row().cells
            for i, val in enumerate(data_row):
                new_cells[i].text = intcomma(str(val).strip()).replace(",", " ").replace(".", ",") if i >= 2 else str(val).strip()
        for hdr_cell in nested.rows[0].cells:
            for p in hdr_cell.paragraphs:
                pf = p.paragraph_format
                pf.first_line_indent = Pt(0)
                pf.left_indent = Pt(0)

        # и то же для всех строк с данными
        for data_row in nested.rows[1:]:
            for cell in data_row.cells:
                for p in cell.paragraphs:
                    pf = p.paragraph_format
                    pf.first_line_indent = Pt(0)
                    pf.left_indent = Pt(0)

        return doc

def insert_k_table(doc: DocumentObject, df: pl.DataFrame):
    df = df.with_columns(sc.by_dtype(pl.Float64).round(4))
    # 1) Находим таблицу и номер заголовочной строки
    target_table = None
    header_row_idx = None

    for tbl in doc.tables:
        for idx, row in enumerate(tbl.rows):
            # проверяем, есть ли в этой строке нужная ячейка
            if any("Общая сумма СПОД" in cell.text for cell in row.cells):
                target_table = tbl
                header_row_idx = idx
                break
        if target_table:
            break

    if not target_table:
        # таблица не найдена — выходим
        return

    # 2) Берём значения из df
    # Предположим, что у вас в df есть колонка "value", и в ней ровно столько строк,
    # сколько строк таблицы (начиная с header_row_idx), куда нужно записать.
    values = df.row(0)

    # 3) Проверяем, хватает ли строк в таблице.
    # Нужны строки от header_row_idx до header_row_idx + len(values) - 1
    needed = header_row_idx + len(values) - len(target_table.rows)
    for _ in range(needed):
        target_table.add_row()

    # 4) Записываем по одной строке в ячейку (row_idx, col_idx=1)
    # print(target_table.rows)
    # print(len(target_table.rows))

    row_indices = []
    for row_idx in range(header_row_idx, header_row_idx + len(target_table.rows) - 1):
        # проверяем первый столбец на непустоту
        # print('row_idx = ', row_idx)
        if not target_table.cell(row_idx, 0).text.strip() or not target_table.cell(row_idx, 1).text.strip():
            continue
        # проверяем второй столбец на "Д"
        if target_table.cell(row_idx, 1).text.startswith("Д"):
            continue
        row_indices.append(row_idx)

    # 2) Теперь просто zip по values и row_indices
    for val, row_idx in zip(values, row_indices):
        cell = target_table.cell(row_idx, 1)
        cell.text = intcomma(str(val)).replace(",", " ").replace(".", ",")
    return doc

def insert_d_table(doc: DocumentObject, df: pl.DataFrame):
    df = df.with_columns(sc.by_dtype(pl.Float64).round(2))
    # 1) Находим таблицу и номер заголовочной строки
    target_table = None
    header_row_idx = None

    for tbl in doc.tables:
        for idx, row in enumerate(tbl.rows):
            # проверяем, есть ли в этой строке нужная ячейка
            if any("Общая сумма СПОД" in cell.text for cell in row.cells):
                target_table = tbl
                header_row_idx = idx
                break
        if target_table:
            break

    if not target_table:
        # таблица не найдена — выходим
        return

    # 2) Берём значения из df
    values = df.to_series(1).to_list()

    # 3) Проверяем, хватает ли строк в таблице.
    # Нужны строки от header_row_idx до header_row_idx + len(values) - 1
    needed = header_row_idx + len(values) - len(target_table.rows)
    for _ in range(needed):
        target_table.add_row()

    # 4) Записываем по одной строке в ячейку (row_idx, col_idx=1)
    # print(target_table.rows)
    # print(len(target_table.rows))

    row_indices = []
    for row_idx in range(header_row_idx, header_row_idx + len(target_table.rows) - 1):
        # проверяем первый столбец на непустоту
        # print('row_idx = ', row_idx)
        if not target_table.cell(row_idx, 0).text.strip() or not target_table.cell(row_idx, 1).text.strip():
            continue
        # проверяем второй столбец на "Д"
        if not target_table.cell(row_idx, 1).text.startswith("Д"):
            continue
        row_indices.append(row_idx)

    # 2) Теперь просто zip по values и row_indices
    for val, row_idx in zip(values, row_indices):
        cell = target_table.cell(row_idx, 1)
        cell.text = intcomma(str(val)).replace(",", " ").replace(".", ",")
    return doc



def clear_cell_shading(cell):
    """
    Убирает любую заливку из ячейки cell:
    удаляет все <w:shd> в свойствах ячейки.
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for shd in tcPr.findall(qn("w:shd")):
        tcPr.remove(shd)


