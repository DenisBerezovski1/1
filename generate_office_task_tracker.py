#!/usr/bin/env python3
"""Generate an Excel workbook for tracking office tasks."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

OUTPUT_PATH = Path(__file__).resolve().parent / "office_task_tracker.xlsx"

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


def column_widths_xml(widths: list[float]) -> str:
    columns = []
    for index, width in enumerate(widths, start=1):
        columns.append(
            f'<col min="{index}" max="{index}" width="{width}" customWidth="1"/>'
        )
    return "".join(columns)


def build_tasks_sheet() -> str:
    headers = [
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

    rows = [
        row_xml(
            1,
            [inline_string_cell(f"{column}1", header, 1) for column, header in zip("ABCDEFGHIJK", headers)],
            "1:11",
            height=24,
        )
    ]

    for row_number in range(2, 202):
        rows.append(
            row_xml(
                row_number,
                [
                    formula_cell(
                        f"A{row_number}",
                        f'IF(C{row_number}="","",ROW()-1)',
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

    data_validations = """
    <dataValidations count="3">
      <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="D2:D201">
        <formula1>task_categories</formula1>
      </dataValidation>
      <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="F2:F201">
        <formula1>task_priorities</formula1>
      </dataValidation>
      <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="G2:G201">
        <formula1>task_statuses</formula1>
      </dataValidation>
    </dataValidations>
    """.strip()

    columns = column_widths_xml([8, 16, 38, 18, 20, 12, 14, 14, 14, 12, 30])

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:K201"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>{columns}</cols>
  <sheetData>{''.join(rows)}</sheetData>
  <autoFilter ref="A1:K201"/>
  {data_validations}
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"""


def build_summary_sheet() -> str:
    metrics = [
        ("Всего задач", "COUNTA('Задачи'!$C$2:$C$201)"),
        ("Новая", 'COUNTIF(\'Задачи\'!$G$2:$G$201,"Новая")'),
        ("В работе", 'COUNTIF(\'Задачи\'!$G$2:$G$201,"В работе")'),
        ("На паузе", 'COUNTIF(\'Задачи\'!$G$2:$G$201,"На паузе")'),
        ("Выполнено", 'COUNTIF(\'Задачи\'!$G$2:$G$201,"Выполнено")'),
        ("Просрочено", 'COUNTIF(\'Задачи\'!$J$2:$J$201,"Да")'),
        ("Высокий приоритет", 'COUNTIF(\'Задачи\'!$F$2:$F$201,"Высокий")'),
        (
            "Срок сегодня",
            'COUNTIFS(\'Задачи\'!$H$2:$H$201,TODAY(),\'Задачи\'!$G$2:$G$201,"<>Выполнено")',
        ),
    ]

    rows = [
        row_xml(
            1,
            [
                inline_string_cell("A1", "Показатель", 1),
                inline_string_cell("B1", "Значение", 1),
            ],
            "1:2",
            height=24,
        )
    ]

    for row_number, (label, formula) in enumerate(metrics, start=2):
        rows.append(
            row_xml(
                row_number,
                [
                    inline_string_cell(f"A{row_number}", label, 2),
                    formula_cell(f"B{row_number}", formula, 4),
                ],
                "1:2",
            )
        )

    columns = column_widths_xml([24, 16])

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B9"/>
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>{columns}</cols>
  <sheetData>{''.join(rows)}</sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"""


def build_lists_sheet() -> str:
    rows = [
        row_xml(
            1,
            [
                inline_string_cell("A1", "Категории", 1),
                inline_string_cell("B1", "Приоритеты", 1),
                inline_string_cell("C1", "Статусы", 1),
            ],
            "1:3",
            height=24,
        )
    ]

    max_items = max(len(CATEGORIES), len(PRIORITIES), len(STATUSES))
    for offset in range(max_items):
        row_number = offset + 2
        cells = []
        cells.append(
            inline_string_cell(f"A{row_number}", CATEGORIES[offset], 2)
            if offset < len(CATEGORIES)
            else blank_cell(f"A{row_number}", 2)
        )
        cells.append(
            inline_string_cell(f"B{row_number}", PRIORITIES[offset], 2)
            if offset < len(PRIORITIES)
            else blank_cell(f"B{row_number}", 2)
        )
        cells.append(
            inline_string_cell(f"C{row_number}", STATUSES[offset], 2)
            if offset < len(STATUSES)
            else blank_cell(f"C{row_number}", 2)
        )
        rows.append(row_xml(row_number, cells, "1:3"))

    columns = column_widths_xml([22, 14, 16])

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C7"/>
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>{columns}</cols>
  <sheetData>{''.join(rows)}</sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"""


def build_styles() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="dd.mm.yyyy"/>
  </numFmts>
  <fonts count="2">
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
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF1F4E78"/>
        <bgColor rgb="FF1F4E78"/>
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
  <cellXfs count="5">
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
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"""


def build_content_types() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
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


def build_workbook() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews>
    <workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="14805"/>
  </bookViews>
  <sheets>
    <sheet name="Задачи" sheetId="1" r:id="rId1"/>
    <sheet name="Сводка" sheetId="2" r:id="rId2"/>
    <sheet name="Списки" sheetId="3" state="hidden" r:id="rId3"/>
  </sheets>
  <definedNames>
    <definedName name="task_categories">'Списки'!$A$2:$A$7</definedName>
    <definedName name="task_priorities">'Списки'!$B$2:$B$4</definedName>
    <definedName name="task_statuses">'Списки'!$C$2:$C$5</definedName>
  </definedNames>
  <calcPr calcId="171027" fullCalcOnLoad="1"/>
</workbook>
"""


def build_workbook_relationships() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
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


def build_app_properties() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties
  xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>3</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="3" baseType="lpstr">
      <vt:lpstr>Задачи</vt:lpstr>
      <vt:lpstr>Сводка</vt:lpstr>
      <vt:lpstr>Списки</vt:lpstr>
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
    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", build_content_types())
        workbook.writestr("_rels/.rels", build_root_relationships())
        workbook.writestr("docProps/core.xml", build_core_properties())
        workbook.writestr("docProps/app.xml", build_app_properties())
        workbook.writestr("xl/workbook.xml", build_workbook())
        workbook.writestr("xl/_rels/workbook.xml.rels", build_workbook_relationships())
        workbook.writestr("xl/styles.xml", build_styles())
        workbook.writestr("xl/worksheets/sheet1.xml", build_tasks_sheet())
        workbook.writestr("xl/worksheets/sheet2.xml", build_summary_sheet())
        workbook.writestr("xl/worksheets/sheet3.xml", build_lists_sheet())


def main() -> None:
    write_workbook(OUTPUT_PATH)
    print(f"Created {OUTPUT_PATH.name}")


if __name__ == "__main__":
    main()
