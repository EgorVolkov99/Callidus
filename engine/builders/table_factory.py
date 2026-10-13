import logging
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

logger = logging.getLogger("callidus")


class TableFactory:
    """
    Фабрика для безопасной генерации таблиц Word.
    Гарантирует соблюдение спецификации OOXML.
    """

    def __init__(self, tpl):
        self.tpl = tpl

    def _create_stub(self):
        return self.tpl.new_subdoc()

    def _extract_headers_and_rows(self, data: dict):
        """Умный извлекатель: понимает и сырой JSON, и данные из адаптера."""
        if not data:
            return [], []

        # Если адаптер уже превратил словарь в формат {"headers": [...], "rows": [...]}
        if "headers" in data and "rows" in data and isinstance(data["headers"], list):
            return data["headers"], data["rows"]

        # Если пришел оригинальный JSON (словарь колонок)
        keys = [k for k, v in data.items() if isinstance(v, list)]
        if not keys:
            return [], []

        num_rows = len(data[keys[0]])
        rows_data = []
        for r in range(num_rows):
            row = [str(data[k][r]) if r < len(data[k]) else "" for k in keys]
            rows_data.append(row)

        return keys, rows_data

    # ==========================================
    # БЕЗОПАСНЫЕ ФОРМАТТЕРЫ ЯЧЕЕК (OOXML)
    # ==========================================
    def _set_table_full_width(self, table):
        tblPr = table._tbl.tblPr
        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:w'), '5000')
        tblW.set(qn('w:type'), 'pct')

        existing_tblW = tblPr.find(qn('w:tblW'))
        if existing_tblW is not None:
            tblPr.remove(existing_tblW)

        tblPr.append(tblW)

    def _insert_in_tcPr(self, tcPr, element):
        tags = [
            'w:cnfStyle', 'w:tcW', 'w:gridSpan', 'w:hMerge', 'w:vMerge',
            'w:tcBorders', 'w:shd', 'w:noWrap', 'w:tcMar', 'w:textDirection',
            'w:tcFitText', 'w:vAlign', 'w:hideMark', 'w:headers', 'w:cellIns',
            'w:cellDel', 'w:cellMerge', 'w:tcPrChange'
        ]
        tag_order = {qn(tag): i for i, tag in enumerate(tags)}
        element_tag = element.tag

        existing = tcPr.find(element_tag)
        if existing is not None:
            tcPr.remove(existing)

        if element_tag not in tag_order:
            tcPr.append(element)
            return

        for i, child in enumerate(tcPr):
            if child.tag in tag_order and tag_order[child.tag] > tag_order[element_tag]:
                tcPr.insert(i, element)
                return
        tcPr.append(element)

    def _format_cell(self, cell, text: str, is_header=False, align="center", bg_color=None, direction=None,
                     bottom_border=None):
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = getattr(WD_ALIGN_PARAGRAPH, align.upper(), WD_ALIGN_PARAGRAPH.CENTER)
        fmt = p.paragraph_format
        fmt.first_line_indent = Cm(0)
        fmt.left_indent = Cm(0)
        fmt.right_indent = Cm(0)
        fmt.space_before = Pt(3)
        fmt.space_after = Pt(3)
        fmt.line_spacing = 1

        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(11)
        if is_header:
            r.bold = True

        tcPr = cell._tc.get_or_add_tcPr()

        if bg_color:
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), bg_color)
            self._insert_in_tcPr(tcPr, shd)

        if direction:
            textDirection = OxmlElement('w:textDirection')
            textDirection.set(qn('w:val'), direction)
            self._insert_in_tcPr(tcPr, textDirection)

        if bottom_border:
            tcBorders = tcPr.find(qn('w:tcBorders'))
            if tcBorders is None:
                tcBorders = OxmlElement('w:tcBorders')
                self._insert_in_tcPr(tcPr, tcBorders)
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'double')
            bottom.set(qn('w:sz'), str(bottom_border))
            bottom.set(qn('w:space'), '0')
            bottom.set(qn('w:color'), 'auto')
            tcBorders.append(bottom)

    def _set_repeat_header(self, row):
        trPr = row._tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        tblHeader.set(qn('w:val'), "true")
        trPr.append(tblHeader)

    # ==========================================
    # 1. БАЗОВЫЕ И ГРУППИРОВАННЫЕ ТАБЛИЦЫ
    # ==========================================

    def build_generic_table(self, headers: list, rows_data: list, col_widths_cm: list | None = None,
                            table_name: str = "Generic"):
        logger.info(f"[TableFactory] Сборка: {table_name}")
        subdoc = self.tpl.new_subdoc()
        if not headers and not rows_data:
            return subdoc

        num_cols = len(headers) if headers else (len(rows_data[0]) if rows_data else 0)
        if num_cols == 0:
            return subdoc

        table = subdoc.add_table(rows=0, cols=num_cols)
        table.style = 'Table Grid'
        table.autofit = False
        self._set_table_full_width(table)

        if headers:
            hdr_row = table.add_row()
            self._set_repeat_header(hdr_row)
            for i, text in enumerate(headers):
                cell = hdr_row.cells[i]
                if col_widths_cm and i < len(col_widths_cm):
                    cell.width = Cm(col_widths_cm[i])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                self._format_cell(cell, str(text), is_header=True, bg_color="F2F2F2", bottom_border="4")

        for item_row in rows_data:
            row = table.add_row()
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            row.height = Cm(1.25)
            for i in range(num_cols):
                cell = row.cells[i]
                if col_widths_cm and i < len(col_widths_cm):
                    cell.width = Cm(col_widths_cm[i])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

                val = str(item_row[i]) if i < len(item_row) and item_row[i] is not None else ""
                if headers and "Серийный номер" in headers and i == 2 and not val:
                    val = "_______________"
                self._format_cell(cell, val)

        subdoc.add_paragraph("")
        return subdoc

    def build_grouped_table(self, data_dict: dict, headers: list, col_widths_cm: list | None = None,
                            table_name: str = "Grouped"):
        logger.info(f"[TableFactory] Сборка: {table_name}")
        subdoc = self.tpl.new_subdoc()
        if not data_dict:
            return subdoc

        num_cols = len(headers) if headers else 2
        table = subdoc.add_table(rows=0, cols=num_cols)
        table.style = 'Table Grid'
        table.autofit = False
        self._set_table_full_width(table)

        if not col_widths_cm:
            col_widths_cm = [10.4, 7.5]

        if headers:
            hdr_row = table.add_row()
            hdr_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            hdr_row.height = Cm(1)
            self._set_repeat_header(hdr_row)
            for i, header_text in enumerate(headers):
                cell = hdr_row.cells[i]
                if i < len(col_widths_cm): cell.width = Cm(col_widths_cm[i])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                self._format_cell(cell, str(header_text), is_header=True, bg_color="F2F2F2", bottom_border="4")

        for group_name, params in data_dict.items():
            grp_row = table.add_row()
            grp_row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            grp_row.height = Cm(0.75)

            grp_row.cells[0].merge(grp_row.cells[-1])
            grp_cell = grp_row.cells[0]
            grp_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            self._format_cell(grp_cell, str(group_name), is_header=True, bg_color="E0E0E0")

            if isinstance(params, dict):
                for key, val in params.items():
                    param_row = table.add_row()
                    param_row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                    param_row.height = Cm(0.75)

                    cell_left, cell_right = param_row.cells[0], param_row.cells[1]
                    if len(col_widths_cm) >= 2:
                        cell_left.width, cell_right.width = Cm(col_widths_cm[0]), Cm(col_widths_cm[1])
                    cell_left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    cell_right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

                    self._format_cell(cell_left, str(key), align="left")
                    self._format_cell(cell_right, str(val), align="center")

        subdoc.add_paragraph("")
        return subdoc

    def build_basic_info_table(self, items: list):
        subdoc = self.tpl.new_subdoc()
        if not items:
            return subdoc

        table = subdoc.add_table(rows=0, cols=2)
        table.style = 'Table Grid'
        table.autofit = False
        self._set_table_full_width(table)

        for item in items:
            row = table.add_row()
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            row.height = Cm(1)

            cell_left, cell_right = row.cells[0], row.cells[1]
            cell_left.width = Cm(7.5)
            cell_right.width = Cm(10.4)

            cell_left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell_right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

            self._format_cell(cell_left, str(item.get('name', '')), is_header=True, align="left", bg_color="F2F2F2")
            self._format_cell(cell_right, str(item.get('value', '')), align="left")

        subdoc.add_paragraph("")
        return subdoc

    def build_approval_tests_table(self, test_data: list, unit_info: dict, has_extra: bool, is_eng=False):
        subdoc = self.tpl.new_subdoc()
        if not test_data: return subdoc

        cols = [
            {"hdr": "Name and designation" if is_eng else "Наименование и обозначение", "key": "stamp", "width": Cm(3)}]
        if has_extra:
            cols.extend([
                {"hdr": "Type of test" if is_eng else "Вид испытаний", "key": "test_type", "width": Cm(5.2)},
                {"hdr": "Test medium" if is_eng else "Среда испытательная", "key": "medium", "width": Cm(3)},
                {"hdr": "Test pressure, MPa" if is_eng else "Давление, МПа", "key": "pressure", "width": Cm(3)},
                {"hdr": "Test temperature, °C" if is_eng else "Температура, °C", "key": "temperature", "width": Cm(3)},
                {
                    "hdr": "Test results according to documentation actual:" if is_eng else "Результаты испытаний по документации фактические:",
                    "key": "result", "width": Cm(8.5)}
            ])
        else:
            cols.extend([
                {"hdr": "Type of test" if is_eng else "Вид испытаний", "key": "test_type", "width": Cm(11.8)},
                {
                    "hdr": "Test results:\naccording to documentation | actual:" if is_eng else "Результаты испытаний:\nпо документации | фактические:",
                    "key": "result", "width": Cm(11.7)}
            ])

        table = subdoc.add_table(rows=0, cols=len(cols))
        table.style = 'Table Grid'
        table.autofit = False
        self._set_table_full_width(table)

        hdr_row = table.add_row()
        self._set_repeat_header(hdr_row)
        for i, col in enumerate(cols):
            cell = hdr_row.cells[i]
            cell.width = col["width"]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            self._format_cell(cell, col["hdr"], is_header=True, bg_color="F2F2F2", bottom_border="4")

        for data_item in test_data:
            row = table.add_row()
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            row.height = Cm(3)
            for i, col in enumerate(cols):
                cell = row.cells[i]
                cell.width = col["width"]
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                key = col["key"]
                if key == "stamp":
                    pass
                elif key == "result":
                    self._format_cell(cell,
                                      str(data_item.get(key) or '') + "\n_________________________________________\n\n",
                                      align="left")
                else:
                    val = str(data_item.get(key) or '-')
                    if not val or val == 'None': val = '-'
                    self._format_cell(cell, val)

        if len(test_data) > 0:
            start_cell = table.cell(1, 0)
            end_cell = table.cell(len(test_data), 0)
            if start_cell != end_cell:
                start_cell.merge(end_cell)

            start_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            prod_name = unit_info.get("product_name_eng" if is_eng else "product_name_rus", "")
            brand = unit_info.get("product_brand", "")
            model = unit_info.get("product_model", "")
            protocol = "Approval test protocol\nNo. _________ dated _________" if is_eng else "Протокол приёмо-сдаточных испытаний\n№ _________ от _________"

            text = f"{prod_name} \n{brand} {model}\n{protocol}"
            self._format_cell(start_cell, text, direction='btLr')

        subdoc.add_paragraph("")
        return subdoc

    # ==========================================
    # 2. УНИКАЛЬНЫЕ ТАБЛИЦЫ РАЗДЕЛА MAIN PARTS
    # ==========================================

    def build_materials_table(self, data: dict, is_eng: bool = False):
        """Таблица ТБ1: Сведения о материалах."""
        logger.info(f"[TableFactory] Сборка: materials (is_eng={is_eng})")
        subdoc = self.tpl.new_subdoc()

        headers, rows_data = self._extract_headers_and_rows(data)
        if not headers:
            return subdoc

        num_cols = len(headers)
        col_widths = [3.0, 3.0, 1.2, 2.5, 1.8, 1.8]
        if num_cols > 6:
            col_widths.extend([0.8] * (num_cols - 6))

        try:
            table = subdoc.add_table(rows=0, cols=num_cols)
            table.style = 'Table Grid'
            table.autofit = False
            self._set_table_full_width(table)

            # Шапка
            hdr_row = table.add_row()
            self._set_repeat_header(hdr_row)
            for i, text in enumerate(headers):
                cell = hdr_row.cells[i]
                if i < len(col_widths): cell.width = Cm(col_widths[i])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                direction = "btLr" if i >= 6 else None
                self._format_cell(cell, text, is_header=True, bg_color="F2F2F2", bottom_border="4", direction=direction)

            # Нумерация колонок
            num_row = table.add_row()
            self._set_repeat_header(num_row)
            for i in range(num_cols):
                cell = num_row.cells[i]
                if i < len(col_widths): cell.width = Cm(col_widths[i])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                self._format_cell(cell, str(i + 1), is_header=True, bg_color="E0E0E0", bottom_border="4")

            # Данные
            for item_row in rows_data:
                row = table.add_row()
                row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                row.height = Cm(0.75)
                for i in range(num_cols):
                    cell = row.cells[i]
                    if i < len(col_widths): cell.width = Cm(col_widths[i])
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    val = str(item_row[i]) if i < len(item_row) else ""
                    self._format_cell(cell, val)

            subdoc.add_paragraph("")
        except Exception as e:
            logger.error(f"[TableFactory] Ошибка в materials: {e}")

        return subdoc

    def build_welds_table(self, data: dict, is_eng: bool = False):
        """Таблица ТБ2: Сварные швы."""
        logger.info(f"[TableFactory] Сборка: welds (is_eng={is_eng})")
        subdoc = self.tpl.new_subdoc()

        headers, rows_data = self._extract_headers_and_rows(data)
        if not headers:
            return subdoc

        try:
            # --- ЧАСТЬ 1 ---
            headers1 = headers[:11]
            num_cols1 = len(headers1)
            col_widths1 = [1.2, 3.0, 1.5, 1.2, 2.5, 2.5, 2.0, 2.0, 2.0, 2.0, 2.0][:num_cols1]

            table1 = subdoc.add_table(rows=0, cols=num_cols1)
            table1.style = 'Table Grid'
            table1.autofit = False
            self._set_table_full_width(table1)

            hdr_row1 = table1.add_row()
            self._set_repeat_header(hdr_row1)
            for i, text in enumerate(headers1):
                cell = hdr_row1.cells[i]
                if i < len(col_widths1): cell.width = Cm(col_widths1[i])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                self._format_cell(cell, text, is_header=True, bg_color="F2F2F2", bottom_border="4")

            num_row1 = table1.add_row()
            self._set_repeat_header(num_row1)
            for i in range(num_cols1):
                cell = num_row1.cells[i]
                if i < len(col_widths1): cell.width = Cm(col_widths1[i])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                self._format_cell(cell, str(i + 1), is_header=True, bg_color="E0E0E0", bottom_border="4")

            for item_row in rows_data:
                row = table1.add_row()
                row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                row.height = Cm(0.75)
                for i in range(num_cols1):
                    cell = row.cells[i]
                    if i < len(col_widths1): cell.width = Cm(col_widths1[i])
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    val = str(item_row[i]) if i < len(item_row) else ""
                    self._format_cell(cell, val)

            # --- ЧАСТЬ 2 ---
            if len(headers) > 11:
                text_continue = "Продолжение таблицы / Continuation of the table" if is_eng else "Продолжение таблицы"
                p = subdoc.add_paragraph(text_continue)
                p.paragraph_format.space_before = Pt(12)
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.keep_with_next = True

                headers2 = [headers[0]] + headers[11:]
                num_cols2 = len(headers2)
                col_widths2 = [1.2] + [0.8] * (num_cols2 - 1)

                table2 = subdoc.add_table(rows=0, cols=num_cols2)
                table2.style = 'Table Grid'
                table2.autofit = False
                self._set_table_full_width(table2)

                hdr_row2 = table2.add_row()
                self._set_repeat_header(hdr_row2)
                for i, text in enumerate(headers2):
                    cell = hdr_row2.cells[i]
                    if i < len(col_widths2): cell.width = Cm(col_widths2[i])
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    direction = "btLr" if i > 0 else None
                    self._format_cell(cell, text, is_header=True, bg_color="F2F2F2", bottom_border="4",
                                      direction=direction)

                num_row2 = table2.add_row()
                self._set_repeat_header(num_row2)

                cell_0 = num_row2.cells[0]
                if 0 < len(col_widths2): cell_0.width = Cm(col_widths2[0])
                cell_0.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                self._format_cell(cell_0, "1", is_header=True, bg_color="E0E0E0", bottom_border="4")

                for i in range(1, num_cols2):
                    cell = num_row2.cells[i]
                    if i < len(col_widths2): cell.width = Cm(col_widths2[i])
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    self._format_cell(cell, str(11 + i), is_header=True, bg_color="E0E0E0", bottom_border="4")

                for item_row in rows_data:
                    row = table2.add_row()
                    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                    row.height = Cm(0.75)

                    cell_0 = row.cells[0]
                    if 0 < len(col_widths2): cell_0.width = Cm(col_widths2[0])
                    cell_0.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    val_0 = str(item_row[0]) if 0 < len(item_row) else ""
                    self._format_cell(cell_0, val_0)

                    for c in range(1, num_cols2):
                        cell = row.cells[c]
                        if c < len(col_widths2): cell.width = Cm(col_widths2[c])
                        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                        val = str(item_row[11 + c - 1]) if (11 + c - 1) < len(item_row) else ""
                        self._format_cell(cell, val)

            subdoc.add_paragraph("")
        except Exception as e:
            logger.error(f"[TableFactory] Ошибка в welds: {e}")

        return subdoc

    def build_valves_table(self, data: dict, is_eng: bool = False):
        """Таблица ТБ3: Арматура и КИП."""
        logger.info(f"[TableFactory] Сборка: valves (is_eng={is_eng})")
        subdoc = self.tpl.new_subdoc()

        headers, rows_data = self._extract_headers_and_rows(data)
        if not headers:
            return subdoc

        num_cols = len(headers)
        col_widths = [3.0, 3.0, 1.5, 2.0, 1.5, 2.0, 2.0, 2.0, 2.0, 2.0][:num_cols]

        try:
            table = subdoc.add_table(rows=0, cols=num_cols)
            table.style = 'Table Grid'
            table.autofit = False
            self._set_table_full_width(table)

            hdr_row = table.add_row()
            self._set_repeat_header(hdr_row)
            for i, text in enumerate(headers):
                cell = hdr_row.cells[i]
                if i < len(col_widths): cell.width = Cm(col_widths[i])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                self._format_cell(cell, text, is_header=True, bg_color="F2F2F2", bottom_border="4")

            num_row = table.add_row()
            self._set_repeat_header(num_row)
            for i in range(num_cols):
                cell = num_row.cells[i]
                if i < len(col_widths): cell.width = Cm(col_widths[i])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                self._format_cell(cell, str(i + 1), is_header=True, bg_color="E0E0E0", bottom_border="4")

            for item_row in rows_data:
                row = table.add_row()
                row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                row.height = Cm(0.75)
                for i in range(num_cols):
                    cell = row.cells[i]
                    if i < len(col_widths): cell.width = Cm(col_widths[i])
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    val = str(item_row[i]) if i < len(item_row) else ""
                    self._format_cell(cell, val)

            subdoc.add_paragraph("")
        except Exception as e:
            logger.error(f"[TableFactory] Ошибка в valves: {e}")

        return subdoc