from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import json
import os


def set_cell_size(cell, width_cm, height_cm):
    """Set cell width and row height in cm (XML-based, precise)"""
    cell.width = Cm(width_cm)
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # Width (EMU)
    tcW = tcPr.get_or_add_tcW()
    tcW.w = int(width_cm * 567)  # 1 cm = 567 twips/dxa (1/20 pt)
    tcW.type = "dxa"

    # Height via row
    tr = tc.getparent()
    trPr = tr.get_or_add_trPr()
    # Fallback for older python-docx
    trHeight_el = trPr.find(qn('w:trHeight'))
    if trHeight_el is None:
        trHeight_el = OxmlElement('w:trHeight')
        trPr.append(trHeight_el)
    trHeight_el.set(qn('w:val'), str(int(height_cm * 567)))
    trHeight_el.set(qn('w:hRule'), 'exact')


def set_cell_text(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=12):
    """Set cell text properties"""
    # Ensure at least one paragraph
    if not cell.paragraphs:
        cell.add_paragraph()
    p = cell.paragraphs[0]
    p.alignment = align
    p.clear()
    run = p.add_run(str(text) if text is not None else '')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(font_size)
    p_format = p.paragraph_format
    p_format.line_spacing = 1.0
    p_format.space_before = Pt(0)
    p_format.space_after = Pt(0)
    cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER


def set_table_width(table, width_cm):
    """Set table width and disable autofit"""
    table.width = Cm(width_cm)
    table.autofit = False
    # Optional: also set column widths explicitly in XML
    for row in table.rows:
        for cell in row.cells:
            # already done in set_cell_size
            pass


def wrap_text_into_paragraphs(paragraph, text, max_width_cm, font_size=12):
    """Wrap long text into multiple lines (by word wrapping), return list of lines"""
    from textwrap import wrap
    # Approximate conversion: 1 cm ≈ 28 pt at 12pt font (TNR), but better to use actual width
    # Instead, we'll split manually by words to fit width
    # For simplicity, use a rough rule: 21.7 cm ≈ 300 pt, and 12pt * 0.6 avg char width * 5 chars/cm ≈ 30 chars/cm -> 21.7*30 ≈ 650 chars per line
    # But better: use python-docx width measurement — not reliable without measuring.
    # We'll do a simple heuristic: assume ~60 chars per 27 cm line (conservative)
    chars_per_line = int(max_width_cm * 25)  # ~25 chars per cm (depends on font)
    return wrap(str(text) if text is not None else '', chars_per_line)


def split_array_elements(arr):
    """If arr is list, return list of elements. Else wrap in list."""
    if isinstance(arr, list):
        return arr
    return [arr] if arr is not None else []


def insert_row_A(doc, item):
    # Required keys
    graph = item.get('graph')
    цех = item.get('Цех')
    операция = item.get('Номер, наименование операции')
    обозначения = split_array_elements(item.get('Обозначение документа'))
    проф = split_array_elements(item.get('Проф.'))
    р = split_array_elements(item.get('Р'))
    кр = split_array_elements(item.get('КР'))
    ут = split_array_elements(item.get('УТ'))
    ен = split_array_elements(item.get('ЕН'))
    тпз = split_array_elements(item.get('Тпз'))
    топ = split_array_elements(item.get('Топ'))

    # Ensure all arrays same length — pad with None if needed
    max_len = max(len(обозначения), len(проф), len(р), len(кр), len(ут), len(ен), len(тпз), len(топ))
    def pad_list(lst, n):
        return lst + [None] * (n - len(lst))
    обозначения = pad_list(обозначения, max_len)
    проф = pad_list(проф, max_len)
    р = pad_list(р, max_len)
    кр = pad_list(кр, max_len)
    ут = pad_list(ут, max_len)
    ен = pad_list(ен, max_len)
    тпз = pad_list(тпз, max_len)
    топ = pad_list(топ, max_len)

    # ✅ First row: full data (graph, Цех, операция + first doc)
    row0 = doc.add_table(rows=1, cols=11)
    row0.style = 'Table Grid'
    cells = [row0.cell(0, j) for j in range(11)]
    widths = [1.6, 1.6, 8, 6.9, 1.8, 1.1, 1.1, 1.1, 1.9, 1.8, 1.8]
    for j, cell in enumerate(cells):
        set_cell_size(cell, widths[j], 0.7)

    set_cell_text(cells[0], graph, WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[1], цех, WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[2], операция, WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(cells[3], обозначения[0], WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(cells[4], проф[0], WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[5], р[0], WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[6], кр[0], WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[7], ут[0], WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[8], ен[0], WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[9], тпз[0], WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[10], топ[0], WD_ALIGN_PARAGRAPH.CENTER)

    set_table_width(row0, 28.7)

    # ✅ Next rows: only doc + prof + R/KR/UT/EN/TPZ/TOP (without graph, Цех, операция)
    for i in range(1, max_len):
        row_n = doc.add_table(rows=1, cols=11)
        row_n.style = 'Table Grid'
        cells_n = [row_n.cell(0, j) for j in range(11)]
        for j, cell in enumerate(cells_n):
            set_cell_size(cell, widths[j], 0.7)

        set_cell_text(cells_n[0], '', WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(cells_n[1], '', WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(cells_n[2], '', WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_text(cells_n[3], обозначения[i], WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_text(cells_n[4], проф[i], WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(cells_n[5], р[i], WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(cells_n[6], кр[i], WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(cells_n[7], ут[i], WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(cells_n[8], ен[i], WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(cells_n[9], тпз[i], WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(cells_n[10], топ[i], WD_ALIGN_PARAGRAPH.CENTER)

        set_table_width(row_n, 28.7)

    # Insert blank row before — as required
    blank_row = doc.add_table(rows=1, cols=1)
    blank_row.style = 'Table Grid'
    cell = blank_row.cell(0, 0)
    set_cell_size(cell, 28.7, 0.7)
    blank_row.cell(0, 0).paragraphs[0].add_run()
    set_table_width(blank_row, 28.7)


def insert_row_O(doc, item):
    """Insert О-row with long text wrapping across lines"""
    graph = item.get('graph')
    содержание = item.get('Содержание операции')
    if not содержание:
        return

    # Wrap content into lines
    lines = wrap_text_into_paragraphs(None, содержание, max_width_cm=27.1)

    # One row per line, each row with full 27.1 cm cell
    for line in lines:
        row = doc.add_table(rows=1, cols=2)
        row.style = 'Table Grid'
        cells = [row.cell(0, 0), row.cell(0, 1)]

        set_table_width(row, 28.7)  # ✅ explicitly set total width

        # Set sizes: [1.6, 27.1]
        set_cell_size(cells[0], 1.6, 0.7)
        set_cell_size(cells[1], 27.1, 0.7)

        set_cell_text(cells[0], graph, WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(cells[1], line, WD_ALIGN_PARAGRAPH.LEFT)


def insert_row_KM(doc, item):
    """Insert К/М-row"""
    graph = item.get('graph')
    наименование = item.get('Наименование детали или сборочной единицы/наименование, марка материала')
    ен = item.get('ЕН')
    ев = item.get('ЕВ')
    нр = item.get('НР')

    row = doc.add_table(rows=1, cols=5)
    row.style = 'Table Grid'
    cells = [row.cell(0, j) for j in range(5)]
    set_table_width(row, 28.7)  # ✅ explicitly set total width

    widths = [1.6, 21.6, 1.9, 1.8, 1.8]
    for i, cell in enumerate(cells):
        set_cell_size(cell, widths[i], 0.7)

    set_cell_text(cells[0], graph, WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[1], наименование, WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(cells[2], ен, WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[3], ев, WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[4], нр, WD_ALIGN_PARAGRAPH.CENTER)


def insert_row_T(doc, item):
    """Insert Т-row"""
    graph = item.get('graph')
    наименование = item.get('Наименование детали или сборочной единицы/наименование, марка материала')
    ев = item.get('ЕВ')
    колво = item.get('Кол-во / HP')

    row = doc.add_table(rows=1, cols=3)
    row.style = 'Table Grid'
    cells = [row.cell(0, 0), row.cell(0, 1), row.cell(0, 2)]
    set_table_width(row, 28.7)  # ✅ explicitly set total width

    widths = [1.6, 23.5, 1.8]
    for i, cell in enumerate(cells):
        set_cell_size(cell, widths[i], 0.7)

    set_cell_text(cells[0], graph, WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[1], наименование, WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(cells[2], ев, WD_ALIGN_PARAGRAPH.CENTER)
    # Количество (ЕВ? Or "Кол-во / HP"? Let's assume 3rd is "ЕВ", 4th is "Кол-во / HP")
    # But spec says: [graph, наименование, ЕВ, Кол-во / HP] — but only 3 columns!
    # Reread spec: row "Т" has only 3 cells: [graph, наименование, ЕВ], then "Кол-во / HP" — likely typo, should be 4 cells.

    # ✅ Correction: The spec lists 3 items but then says "Кол-во / HP" as 5th cell. This seems like error.
    # Most likely: 1) graph, 2) название, 3) ЕВ, 4) Кол-во / HP.
    # Let's fix table to 4 columns (per spec line: "3) Ячейка содержит значение ключа "ЕВ"... 5) Ячейка содержит... Кол-во / HР")
    # => So the row has 4 columns, not 3. Let's fix widths accordingly:
    # [1.6, 23.5, 1.8, 1.8]
    # => Let's recreate with 4 columns

def insert_row_T_fixed(doc, item):
    graph = item.get('graph')
    наименование = item.get('Наименование детали или сборочной единицы/наименование, марка материала')
    ев = item.get('ЕВ')
    колво = item.get('Кол-во / HP')

    row = doc.add_table(rows=1, cols=4)
    row.style = 'Table Grid'
    cells = [row.cell(0, j) for j in range(4)]
    set_table_width(row, 28.7)  # ✅ explicitly set total width

    widths = [1.6, 23.5, 1.8, 1.8]
    for i, cell in enumerate(cells):
        set_cell_size(cell, widths[i], 0.7)

    set_cell_text(cells[0], graph, WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[1], наименование, WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(cells[2], ев, WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(cells[3], колво, WD_ALIGN_PARAGRAPH.CENTER)


def process_json_to_doc(json_path, form_type="КТП"):
    """
    Converts JSON array (from file at `json_path`) to Word document.
    form_type must be "КТП" — only supported for now.
    Returns path to generated .docx file.
    """
    # Load JSON
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # Validate array
    if not isinstance(data, list):
        raise ValueError("JSON must contain an array of dictionaries")

    # Create document
    doc = Document()

    # --- Page setup: A4 landscape ---
    section = doc.sections[0]
    # Set A4 landscape manually (width > height)
    section.page_width = Cm(29.7)  # 297 mm = 29.7 cm
    section.page_height = Cm(21.0)  # 210 mm = 21.0 cm
    # Page margins 0.5 cm (already set below)
    section.top_margin = Cm(0.5)
    section.bottom_margin = Cm(0.5)
    section.left_margin = Cm(0.5)
    section.right_margin = Cm(0.5)

    # Add title
    p = doc.add_paragraph(f"Таблица формата {form_type}")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    run.bold = True

    # Process each item
    for item in data:
        # Detect row types: item should have keys like "graph", etc.
        graph = item.get('graph')
        if not graph:
            continue  # Skip invalid entries

        # Infer type by presence of keys:
        if 'Номер, наименование операции' in item:
            insert_row_A(doc, item)
        elif 'Содержание операции' in item:
            insert_row_O(doc, item)
        elif 'Наименование детали или сборочной единицы/наименование, марка материала' in item:
            # Distinguish К/М vs Т by other keys
            if 'ЕН' in item and 'ЕВ' in item and 'НР' in item and 'Кол-во / HP' not in item:
                insert_row_KM(doc, item)
            elif 'Кол-во / HP' in item:
                insert_row_T_fixed(doc, item)
            else:
                # Default fallback — try to guess
                insert_row_KM(doc, item)
        # else: skip or log unknown

    # Save
    output_path = os.path.splitext(json_path)[0] + "_output.docx"
    doc.save(output_path)
    return output_path


# Example usage (uncomment to run):
if __name__ == "__main__":
    # Replace with your actual file path
    input_file = "../../projects/ASOS Desalination Plant/temp/ktp_tube_exchanger.json"
    output_file = process_json_to_doc(input_file, "КТП")
    print(f"✅ Generated: {output_file}")