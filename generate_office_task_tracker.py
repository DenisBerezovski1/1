#!/usr/bin/env python3
"""Generate an Excel workbook for tracking office tasks."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

OUTPUT_PATH = Path(__file__).resolve().parent / "office_task_tracker.xlsx"

MAX_TASK_ROWS = 200
MAX_CONTACT_ROWS = 120

TASKS_SHEET_NAME = "Задачи"
SUMMARY_SHEET_NAME = "Сводка"
CONTACTS_SHEET_NAME = "Контакты"
SETTINGS_SHEET_NAME = "Настройки"

CATEGORIES = [
    "Документы",
    "Закупки",
    "Встречи",
    "IT",
    "Администрирование",
    "Прочее",
]
PRIORITIES = ["Низкий", "Средний", "Высокий"]
STATUSES = ["Новая", "В работе", "На паузе", "Выполнено"]
EMPLOYEES = [f"Сотрудник {index}" for index in range(1, 6)]

TASK_HEADERS = [
    "ID",
    "Дата постановки",
    "Задача",
    "Категория",
    "Ответственный",
    "Приоритет",
    "Статус",
    "Срок",
    "Осталось дней",
    "Просрочено",
    "Комментарий",
]
TASK_COLUMN_WIDTHS = [8, 16, 40, 18, 20, 12, 14, 14, 14, 12, 32]
CONTACT_HEADERS = ["ФИО", "Номер телефона", "Почтовый адрес"]
CONTACT_COLUMN_WIDTHS = [28, 20, 42]

TASK_TITLE_ROW = 1
TASK_HEADER_ROW = 2
TASK_FIRST_DATA_ROW = 3
TASK_LAST_DATA_ROW = TASK_FIRST_DATA_ROW + MAX_TASK_ROWS - 1

CONTACTS_TITLE_ROW = 1
CONTACTS_HEADER_ROW = 2
CONTACTS_FIRST_DATA_ROW = 3
CONTACTS_LAST_DATA_ROW = CONTACTS_FIRST_DATA_ROW + MAX_CONTACT_ROWS - 1

SETTINGS_TITLE_ROW = 1
SETTINGS_HEADER_ROW = 2
SETTINGS_FIRST_ITEM_ROW = 3


def quoted_sheet_name(sheet_name: str) -> str:
    return f"'{sheet_name}'"


def inline_string_cell(ref: str, value: str, style_id: int = 0) -> str:
    return (
        f'<c r="{ref}" s="{style_id}" t="inlineStr">'
        f"<is><t>{escape(value)}</t></is>"
        f"</c>"
    )


def blank_cell(ref: str, style_id: int = 0) -> str:
    return f'<c r="{ref}" s="{style_id}"/>'


def formula_cell(
    ref: str,
    formula: str,
    style_id: int = 0,
    cell_type: str | None = None,
) -> str:
    type_attr = f' t="{cell_type}"' if cell_type else ""
    return f'<c r="{ref}" s="{style_id}"{type_attr}><f>{escape(formula)}</f></c>'


def row_xml(
    row_number: int,
    cells: list[str],
    spans: str,
    *,
    height: int | None = None,
) -> str:
    height_attr = f' ht="{height}" customHeight="1"' if height is not None else ""
    return (
        f'<row r="{row_number}" spans="{spans}"{height_attr}>'
        f"{''.join(cells)}"
        f"</row>"
    )


def column_widths_xml(
    widths: list[float],
    *,
    hidden_columns: set[int] | None = None,
) -> str:
    hidden_columns = hidden_columns or set()
    columns = []
    for index, width in enumerate(widths, start=1):
        hidden_attr = ' hidden="1"' if index in hidden_columns else ""
        columns.append(
            f'<col min="{index}" max="{index}" width="{width}" customWidth="1"{hidden_attr}/>'
        )
    return "".join(columns)


def merge_cells_xml(ranges: list[str]) -> str:
    if not ranges:
        return ""
    return (
        f'<mergeCells count="{len(ranges)}">'
        + "".join(f'<mergeCell ref="{cell_range}"/>' for cell_range in ranges)
        + "</mergeCells>"
    )


def sheet_pr_xml(tab_color: str | None = None) -> str:
    if not tab_color:
        return ""
    return f'<sheetPr><tabColor rgb="{tab_color}"/></sheetPr>'


def build_status_conditional_formatting(start_row: int, end_row: int) -> str:
    return f"""
  <conditionalFormatting sqref="A{start_row}:K{end_row}">
    <cfRule type="expression" dxfId="0" priority="1" stopIfTrue="1">
      <formula>AND($J{start_row}="Да",$C{start_row}&lt;&gt;"")</formula>
    </cfRule>
    <cfRule type="expression" dxfId="1" priority="2">
      <formula>$G{start_row}="Новая"</formula>
    </cfRule>
    <cfRule type="expression" dxfId="2" priority="3">
      <formula>$G{start_row}="В работе"</formula>
    </cfRule>
    <cfRule type="expression" dxfId="3" priority="4">
      <formula>$G{start_row}="На паузе"</formula>
    </cfRule>
    <cfRule type="expression" dxfId="4" priority="5">
      <formula>$G{start_row}="Выполнено"</formula>
    </cfRule>
  </conditionalFormatting>
""".strip()


def build_tasks_sheet() -> str:
    rows = [
        row_xml(
            TASK_TITLE_ROW,
            [
                inline_string_cell("A1", "Трекер офисных задач", 5),
                *[blank_cell(f"{column}1", 5) for column in "BCDEFGHIJK"],
            ],
            "1:11",
            height=28,
        ),
        row_xml(
            TASK_HEADER_ROW,
            [
                inline_string_cell(f"{column}{TASK_HEADER_ROW}", header, 1)
                for column, header in zip("ABCDEFGHIJK", TASK_HEADERS)
            ],
            "1:11",
            height=22,
        ),
    ]

    for row_number in range(TASK_FIRST_DATA_ROW, TASK_LAST_DATA_ROW + 1):
        rows.append(
            row_xml(
                row_number,
                [
                    formula_cell(
                        f"A{row_number}",
                        f'IF(C{row_number}="","",ROW()-{TASK_HEADER_ROW})',
                        4,
                    ),
                    blank_cell(f"B{row_number}", 3),
                    blank_cell(f"C{row_number}", 2),
                    blank_cell(f"D{row_number}", 2),
                    blank_cell(f"E{row_number}", 2),
                    blank_cell(f"F{row_number}", 4),
                    blank_cell(f"G{row_number}", 4),
                    blank_cell(f"H{row_number}", 3),
                    formula_cell(
                        f"I{row_number}",
                        f'IF(H{row_number}="","",H{row_number}-TODAY())',
                        4,
                    ),
                    formula_cell(
                        f"J{row_number}",
                        (
                            f'IF(OR(H{row_number}="",G{row_number}="Выполнено"),"",'
                            f'IF(H{row_number}<TODAY(),"Да","Нет"))'
                        ),
                        4,
                        cell_type="str",
                    ),
                    blank_cell(f"K{row_number}", 2),
                ],
                "1:11",
            )
        )

    data_validations = f"""
  <dataValidations count="4">
    <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="D{TASK_FIRST_DATA_ROW}:D{TASK_LAST_DATA_ROW}">
      <formula1>task_categories</formula1>
    </dataValidation>
    <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="E{TASK_FIRST_DATA_ROW}:E{TASK_LAST_DATA_ROW}">
      <formula1>task_employees</formula1>
    </dataValidation>
    <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="F{TASK_FIRST_DATA_ROW}:F{TASK_LAST_DATA_ROW}">
      <formula1>task_priorities</formula1>
    </dataValidation>
    <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="G{TASK_FIRST_DATA_ROW}:G{TASK_LAST_DATA_ROW}">
      <formula1>task_statuses</formula1>
    </dataValidation>
  </dataValidations>
""".strip()

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  {sheet_pr_xml("FF2F75B5")}
  <dimension ref="A1:K{TASK_LAST_DATA_ROW}"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="2" topLeftCell="A{TASK_FIRST_DATA_ROW}" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A{TASK_FIRST_DATA_ROW}" sqref="A{TASK_FIRST_DATA_ROW}"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="20"/>
  <cols>{column_widths_xml(TASK_COLUMN_WIDTHS)}</cols>
  <sheetData>{''.join(rows)}</sheetData>
  <autoFilter ref="A{TASK_HEADER_ROW}:K{TASK_LAST_DATA_ROW}"/>
  {merge_cells_xml(["A1:K1"])}
  {build_status_conditional_formatting(TASK_FIRST_DATA_ROW, TASK_LAST_DATA_ROW)}
  {data_validations}
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"""


def build_summary_sheet() -> str:
    tasks_sheet = quoted_sheet_name(TASKS_SHEET_NAME)
    metrics = [
        (
            "Всего задач",
            f"COUNTA({tasks_sheet}!$C${TASK_FIRST_DATA_ROW}:$C${TASK_LAST_DATA_ROW})",
        ),
        (
            "Новая",
            f'COUNTIF({tasks_sheet}!$G${TASK_FIRST_DATA_ROW}:$G${TASK_LAST_DATA_ROW},"Новая")',
        ),
        (
            "В работе",
            f'COUNTIF({tasks_sheet}!$G${TASK_FIRST_DATA_ROW}:$G${TASK_LAST_DATA_ROW},"В работе")',
        ),
        (
            "На паузе",
            f'COUNTIF({tasks_sheet}!$G${TASK_FIRST_DATA_ROW}:$G${TASK_LAST_DATA_ROW},"На паузе")',
        ),
        (
            "Выполнено",
            f'COUNTIF({tasks_sheet}!$G${TASK_FIRST_DATA_ROW}:$G${TASK_LAST_DATA_ROW},"Выполнено")',
        ),
        (
            "Просрочено",
            f'COUNTIF({tasks_sheet}!$J${TASK_FIRST_DATA_ROW}:$J${TASK_LAST_DATA_ROW},"Да")',
        ),
        (
            "Высокий приоритет",
            f'COUNTIF({tasks_sheet}!$F${TASK_FIRST_DATA_ROW}:$F${TASK_LAST_DATA_ROW},"Высокий")',
        ),
        (
            "Срок сегодня",
            (
                f'COUNTIFS({tasks_sheet}!$H${TASK_FIRST_DATA_ROW}:$H${TASK_LAST_DATA_ROW},TODAY(),'
                f'{tasks_sheet}!$G${TASK_FIRST_DATA_ROW}:$G${TASK_LAST_DATA_ROW},"<>Выполнено")'
            ),
        ),
        (
            "Назначено другим",
            (
                f'COUNTIFS({tasks_sheet}!$E${TASK_FIRST_DATA_ROW}:$E${TASK_LAST_DATA_ROW},"<>",'
                f'{tasks_sheet}!$C${TASK_FIRST_DATA_ROW}:$C${TASK_LAST_DATA_ROW},"<>")'
            ),
        ),
        (
            "Без ответственного",
            (
                f'COUNTIFS({tasks_sheet}!$E${TASK_FIRST_DATA_ROW}:$E${TASK_LAST_DATA_ROW},"",'
                f'{tasks_sheet}!$C${TASK_FIRST_DATA_ROW}:$C${TASK_LAST_DATA_ROW},"<>")'
            ),
        ),
    ]

    rows = [
        row_xml(
            1,
            [
                inline_string_cell("A1", "Сводка по задачам", 5),
                blank_cell("B1", 5),
            ],
            "1:2",
            height=28,
        ),
        row_xml(
            2,
            [
                inline_string_cell(
                    "A2",
                    (
                        "Лист помогает контролировать личные задачи: ответственных можно указывать "
                        "в основной таблице, но работа ведется из одной вкладки."
                    ),
                    6,
                ),
                blank_cell("B2", 6),
            ],
            "1:2",
            height=24,
        ),
        row_xml(
            3,
            [
                inline_string_cell("A3", "Показатель", 1),
                inline_string_cell("B3", "Значение", 1),
            ],
            "1:2",
            height=22,
        ),
    ]

    table_height = len(metrics)
    for offset in range(table_height):
        row_number = 4 + offset
        label, formula = metrics[offset]
        rows.append(
            row_xml(
                row_number,
                [
                    inline_string_cell(f"A{row_number}", label, 6),
                    formula_cell(f"B{row_number}", formula, 7),
                ],
                "1:2",
            )
        )

    last_row = 3 + table_height
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  {sheet_pr_xml("FF8064A2")}
  <dimension ref="A1:B{last_row}"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A4" sqref="A4"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="20"/>
  <cols>{column_widths_xml([28, 16])}</cols>
  <sheetData>{''.join(rows)}</sheetData>
  {merge_cells_xml(["A1:B1", "A2:B2"])}
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"""


def build_contacts_sheet() -> str:
    rows = [
        row_xml(
            CONTACTS_TITLE_ROW,
            [
                inline_string_cell("A1", "Контакты для рабочих взаимодействий", 5),
                blank_cell("B1", 5),
                blank_cell("C1", 5),
            ],
            "1:3",
            height=28,
        ),
        row_xml(
            CONTACTS_HEADER_ROW,
            [
                inline_string_cell(f"{column}{CONTACTS_HEADER_ROW}", header, 1)
                for column, header in zip("ABC", CONTACT_HEADERS)
            ],
            "1:3",
            height=22,
        ),
    ]

    for row_number in range(CONTACTS_FIRST_DATA_ROW, CONTACTS_LAST_DATA_ROW + 1):
        rows.append(
            row_xml(
                row_number,
                [
                    blank_cell(f"A{row_number}", 2),
                    blank_cell(f"B{row_number}", 4),
                    blank_cell(f"C{row_number}", 2),
                ],
                "1:3",
            )
        )

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  {sheet_pr_xml("FF70AD47")}
  <dimension ref="A1:C{CONTACTS_LAST_DATA_ROW}"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="2" topLeftCell="A{CONTACTS_FIRST_DATA_ROW}" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A{CONTACTS_FIRST_DATA_ROW}" sqref="A{CONTACTS_FIRST_DATA_ROW}"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="20"/>
  <cols>{column_widths_xml(CONTACT_COLUMN_WIDTHS)}</cols>
  <sheetData>{''.join(rows)}</sheetData>
  <autoFilter ref="A{CONTACTS_HEADER_ROW}:C{CONTACTS_LAST_DATA_ROW}"/>
  {merge_cells_xml(["A1:C1"])}
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"""


def build_settings_sheet() -> str:
    rows = [
        row_xml(
            SETTINGS_TITLE_ROW,
            [
                inline_string_cell("A1", "Настройки и справочники", 5),
                blank_cell("B1", 5),
                blank_cell("C1", 5),
                blank_cell("D1", 5),
            ],
            "1:4",
            height=28,
        ),
        row_xml(
            SETTINGS_HEADER_ROW,
            [
                inline_string_cell("A2", "Категории", 1),
                inline_string_cell("B2", "Приоритеты", 1),
                inline_string_cell("C2", "Статусы", 1),
                inline_string_cell("D2", "Сотрудники", 1),
            ],
            "1:4",
            height=22,
        ),
    ]

    max_items = max(len(CATEGORIES), len(PRIORITIES), len(STATUSES), len(EMPLOYEES))
    for offset in range(max_items):
        row_number = SETTINGS_FIRST_ITEM_ROW + offset
        cells = [
            (
                inline_string_cell(f"A{row_number}", CATEGORIES[offset], 2)
                if offset < len(CATEGORIES)
                else blank_cell(f"A{row_number}", 2)
            ),
            (
                inline_string_cell(f"B{row_number}", PRIORITIES[offset], 4)
                if offset < len(PRIORITIES)
                else blank_cell(f"B{row_number}", 4)
            ),
            (
                inline_string_cell(f"C{row_number}", STATUSES[offset], 4)
                if offset < len(STATUSES)
                else blank_cell(f"C{row_number}", 4)
            ),
            (
                inline_string_cell(f"D{row_number}", EMPLOYEES[offset], 2)
                if offset < len(EMPLOYEES)
                else blank_cell(f"D{row_number}", 2)
            ),
        ]
        rows.append(row_xml(row_number, cells, "1:4"))

    last_row = SETTINGS_FIRST_ITEM_ROW + max_items - 1
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  {sheet_pr_xml("FF7F7F7F")}
  <dimension ref="A1:D{last_row}"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A3" sqref="A3"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="20"/>
  <cols>{column_widths_xml([24, 14, 16, 20])}</cols>
  <sheetData>{''.join(rows)}</sheetData>
  <autoFilter ref="A2:D{last_row}"/>
  {merge_cells_xml(["A1:D1"])}
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"""


def build_styles() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="dd.mm.yyyy"/>
  </numFmts>
  <fonts count="4">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color rgb="FFFFFFFF"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color rgb="FF1F1F1F"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="14"/>
      <color rgb="FFFFFFFF"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF1F4E78"/>
        <bgColor rgb="FF1F4E78"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF4F81BD"/>
        <bgColor rgb="FF4F81BD"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFDEEAF6"/>
        <bgColor rgb="FFDEEAF6"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFFDF2CC"/>
        <bgColor rgb="FFFDF2CC"/>
      </patternFill>
    </fill>
  </fills>
  <borders count="2">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FFD9D9D9"/></top>
      <bottom style="thin"><color rgb="FFD9D9D9"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="8">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment vertical="top" wrapText="1"/>
    </xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="2" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="5">
    <dxf>
      <font><color rgb="FF9C0006"/></font>
      <fill>
        <patternFill patternType="solid">
          <fgColor rgb="FFFFC7CE"/>
          <bgColor rgb="FFFFC7CE"/>
        </patternFill>
      </fill>
    </dxf>
    <dxf>
      <font><color rgb="FF1F4E78"/></font>
      <fill>
        <patternFill patternType="solid">
          <fgColor rgb="FFEAF3FF"/>
          <bgColor rgb="FFEAF3FF"/>
        </patternFill>
      </fill>
    </dxf>
    <dxf>
      <font><color rgb="FF7F6000"/></font>
      <fill>
        <patternFill patternType="solid">
          <fgColor rgb="FFFFF2CC"/>
          <bgColor rgb="FFFFF2CC"/>
        </patternFill>
      </fill>
    </dxf>
    <dxf>
      <font><color rgb="FFC65911"/></font>
      <fill>
        <patternFill patternType="solid">
          <fgColor rgb="FFFCE4D6"/>
          <bgColor rgb="FFFCE4D6"/>
        </patternFill>
      </fill>
    </dxf>
    <dxf>
      <font><color rgb="FF385723"/></font>
      <fill>
        <patternFill patternType="solid">
          <fgColor rgb="FFE2F0D9"/>
          <bgColor rgb="FFE2F0D9"/>
        </patternFill>
      </fill>
    </dxf>
  </dxfs>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"""


def build_content_types(sheet_count: int) -> str:
    sheet_overrides = "".join(
        (
            '  <Override PartName="/xl/worksheets/sheet'
            f'{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
        )
        for index in range(1, sheet_count + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{sheet_overrides}  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def build_root_relationships() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def build_workbook(sheet_names: list[str]) -> str:
    settings_sheet = quoted_sheet_name(SETTINGS_SHEET_NAME)
    sheets_xml = "\n".join(
        f'    <sheet name="{escape(sheet_name)}" sheetId="{index}" r:id="rId{index}"/>'
        for index, sheet_name in enumerate(sheet_names, start=1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews>
    <workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="14805"/>
  </bookViews>
  <sheets>
{sheets_xml}
  </sheets>
  <definedNames>
    <definedName name="task_categories">{settings_sheet}!$A${SETTINGS_FIRST_ITEM_ROW}:$A${SETTINGS_FIRST_ITEM_ROW + len(CATEGORIES) - 1}</definedName>
    <definedName name="task_priorities">{settings_sheet}!$B${SETTINGS_FIRST_ITEM_ROW}:$B${SETTINGS_FIRST_ITEM_ROW + len(PRIORITIES) - 1}</definedName>
    <definedName name="task_statuses">{settings_sheet}!$C${SETTINGS_FIRST_ITEM_ROW}:$C${SETTINGS_FIRST_ITEM_ROW + len(STATUSES) - 1}</definedName>
    <definedName name="task_employees">{settings_sheet}!$D${SETTINGS_FIRST_ITEM_ROW}:$D${SETTINGS_FIRST_ITEM_ROW + len(EMPLOYEES) - 1}</definedName>
  </definedNames>
  <calcPr calcId="171027" fullCalcOnLoad="1"/>
</workbook>
"""


def build_workbook_relationships(sheet_count: int) -> str:
    sheet_relationships = "\n".join(
        (
            f'  <Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
        for index in range(1, sheet_count + 1)
    )
    styles_relationship_id = sheet_count + 1
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{sheet_relationships}
  <Relationship Id="rId{styles_relationship_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""


def build_core_properties() -> str:
    created_at = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Office Task Tracker</dc:title>
  <dc:creator>Cursor Cloud Agent</dc:creator>
  <cp:lastModifiedBy>Cursor Cloud Agent</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{created_at}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{created_at}</dcterms:modified>
</cp:coreProperties>
"""


def build_app_properties(sheet_names: list[str]) -> str:
    titles = "\n".join(
        f"      <vt:lpstr>{escape(sheet_name)}</vt:lpstr>" for sheet_name in sheet_names
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties
  xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>{len(sheet_names)}</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="{len(sheet_names)}" baseType="lpstr">
{titles}
    </vt:vector>
  </TitlesOfParts>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0300</AppVersion>
</Properties>
"""


def write_workbook(output_path: Path) -> None:
    sheet_names = [
        TASKS_SHEET_NAME,
        SUMMARY_SHEET_NAME,
        CONTACTS_SHEET_NAME,
        SETTINGS_SHEET_NAME,
    ]
    worksheets = [
        build_tasks_sheet(),
        build_summary_sheet(),
        build_contacts_sheet(),
        build_settings_sheet(),
    ]

    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", build_content_types(len(worksheets)))
        workbook.writestr("_rels/.rels", build_root_relationships())
        workbook.writestr("docProps/core.xml", build_core_properties())
        workbook.writestr("docProps/app.xml", build_app_properties(sheet_names))
        workbook.writestr("xl/workbook.xml", build_workbook(sheet_names))
        workbook.writestr(
            "xl/_rels/workbook.xml.rels",
            build_workbook_relationships(len(worksheets)),
        )
        workbook.writestr("xl/styles.xml", build_styles())

        for sheet_index, worksheet_xml in enumerate(worksheets, start=1):
            workbook.writestr(
                f"xl/worksheets/sheet{sheet_index}.xml",
                worksheet_xml,
            )


def main() -> None:
    write_workbook(OUTPUT_PATH)
    print(f"Created {OUTPUT_PATH.name}")


if __name__ == "__main__":
    main()
