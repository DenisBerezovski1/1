from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape
import zipfile


OUTPUT_PATH = Path(__file__).resolve().with_name("vedomost_scheta_ucheta_1c.xlsx")
HEADER = [
    "Наименование",
    "Счет учета",
    "Счет списания",
    "Расшифровка назначения",
]

ROWS = [
    (
        "Датчик температуры ET-K PT1000",
        "Датчик системы автоматики; используется в текущих монтажных/производственных работах, списывается в себестоимость",
        "Датчик системы автоматики; используется при строительстве, монтаже или дооборудовании объекта ОС, включается в капитальные вложения",
    ),
    (
        "Датчик температуры комнатный SHUFT RTF1-PT1000",
        "Датчик системы автоматики и вентиляции; используется в текущих монтажных/производственных работах, списывается в себестоимость",
        "Датчик системы автоматики и вентиляции; используется при строительстве, монтаже или дооборудовании объекта ОС, включается в капитальные вложения",
    ),
    (
        "Датчик перепада давления DVL-200",
        "Датчик контроля перепада давления; используется в текущих монтажных/производственных работах, списывается в себестоимость",
        "Датчик контроля перепада давления; используется при строительстве, монтаже или дооборудовании объекта ОС, включается в капитальные вложения",
    ),
    (
        "Датчик потока SHUFT SI FE",
        "Датчик контроля потока; используется в текущих монтажных/производственных работах, списывается в себестоимость",
        "Датчик контроля потока; используется при строительстве, монтаже или дооборудовании объекта ОС, включается в капитальные вложения",
    ),
    (
        "Реле перепада давления NKA217804955-CN6",
        "Реле контроля перепада давления; используется в текущих монтажных/производственных работах, списывается в себестоимость",
        "Реле контроля перепада давления; используется при строительстве, монтаже или дооборудовании объекта ОС, включается в капитальные вложения",
    ),
    (
        "Реле перепада давления (контроль запыленности фильтра) NKA217804956-CN6 / NKA217804951-CN6",
        "Реле контроля загрязненности фильтра; используется в текущих монтажных/производственных работах, списывается в себестоимость",
        "Реле контроля загрязненности фильтра; используется при строительстве, монтаже или дооборудовании объекта ОС, включается в капитальные вложения",
    ),
    (
        "Реле потока (L22...)",
        "Реле контроля потока; используется в текущих монтажных/производственных работах, списывается в себестоимость",
        "Реле контроля потока; используется при строительстве, монтаже или дооборудовании объекта ОС, включается в капитальные вложения",
    ),
    (
        "Кабель монтажный МКШнг(А)-LS 2x0,75",
        "Монтажный кабель для подключения оборудования и цепей управления; списывается в себестоимость",
        "Монтажный кабель для подключения оборудования и цепей управления; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Кабель монтажный МКШнг(А)-LS 1x0,75",
        "Монтажный кабель для внутренних соединений и слаботочных цепей; списывается в себестоимость",
        "Монтажный кабель для внутренних соединений и слаботочных цепей; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Кабель интерфейсный RS-485 КППЭнг(А)-LS 2x2x0,78 (2x2x0,75)",
        "Интерфейсный кабель для линий связи автоматики; списывается в себестоимость",
        "Интерфейсный кабель для линий связи автоматики; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Кабель контрольный КВВГнг(А)-LS 7x0,75",
        "Контрольный кабель для цепей управления и сигнализации; списывается в себестоимость",
        "Контрольный кабель для цепей управления и сигнализации; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Кабель контрольный КВВГнг(А)-LS 10x0,75",
        "Контрольный кабель для цепей управления и сигнализации; списывается в себестоимость",
        "Контрольный кабель для цепей управления и сигнализации; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Кабель контрольный КВВГнг(А)-LS 14x0,75",
        "Контрольный кабель для многожильных цепей управления и сигнализации; списывается в себестоимость",
        "Контрольный кабель для многожильных цепей управления и сигнализации; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Провод установочный ПуГВнг(А)-LS 1x0,75 (белый/синий)",
        "Установочный провод для внутренних соединений в шкафах и щитах; списывается в себестоимость",
        "Установочный провод для внутренних соединений в шкафах и щитах; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Гофротруба армированная d20 с протяжкой",
        "Материал для защиты и прокладки кабельных линий; списывается в себестоимость",
        "Материал для защиты и прокладки кабельных линий; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Держатель с защелкой d20/d25",
        "Крепежный элемент для монтажа гофры и кабельных трасс; списывается в себестоимость",
        "Крепежный элемент для монтажа гофры и кабельных трасс; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Лента перфорированная оцинкованная",
        "Монтажный и крепежный материал для фиксации трасс и оборудования; списывается в себестоимость",
        "Монтажный и крепежный материал для фиксации трасс и оборудования; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Короб перфорированный RL 60x60",
        "Короб для укладки и организации проводов в шкафу; списывается в себестоимость",
        "Короб для укладки и организации проводов в шкафу; включается в капитальные вложения по объекту ОС",
    ),
    (
        "DIN-рейка (AK2..., OMEGA 35)",
        "Монтажная рейка для установки модульного оборудования; списывается в себестоимость",
        "Монтажная рейка для установки модульного оборудования; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Клемма TUR-4 (серая)",
        "Клемма для подключения и коммутации проводников; списывается в себестоимость",
        "Клемма для подключения и коммутации проводников; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Клемма TUR-2.5 (серая/синяя)",
        "Клемма для подключения и разводки электрических цепей; списывается в себестоимость",
        "Клемма для подключения и разводки электрических цепей; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Изолятор торцевой для клемм (D-TUR-2.5-10)",
        "Дополнительный элемент к клеммным сборкам; списывается в себестоимость",
        "Дополнительный элемент к клеммным сборкам; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Перемычка вставная SBF-10-6",
        "Комплектующее для объединения клемм в электрических цепях; списывается в себестоимость",
        "Комплектующее для объединения клемм в электрических цепях; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Наконечник кабельный ТМЛ(DIN) 16-10",
        "Расходный материал для оконцевания кабельных жил; списывается в себестоимость",
        "Расходный материал для оконцевания кабельных жил; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Наконечник штыревой втулочный НШВИ 0,75-8",
        "Расходный материал для подготовки и подключения проводников; списывается в себестоимость",
        "Расходный материал для подготовки и подключения проводников; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Клеммник с держателем предохранителя",
        "Защитное электромонтажное комплектующее для подключения цепей; списывается в себестоимость",
        "Защитное электромонтажное комплектующее для подключения цепей; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Предохранитель стеклянный 5x20",
        "Защитный расходный элемент электрических цепей; списывается в себестоимость",
        "Защитный расходный элемент электрических цепей; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Концевой элемент на клеммы",
        "Дополнительный элемент для сборки клеммных групп; списывается в себестоимость",
        "Дополнительный элемент для сборки клеммных групп; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Интерфейсное реле (в т.ч. PLC-RSC-24DC/21 и аналоги)",
        "Реле для коммутации и согласования сигналов автоматики; списывается в себестоимость",
        "Реле для коммутации и согласования сигналов автоматики; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Модуль дискретного ввода (K2.DI...)",
        "Модуль приема дискретных сигналов системы автоматики; списывается в себестоимость",
        "Модуль приема дискретных сигналов системы автоматики; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Модуль дискретного вывода (K2.DO...)",
        "Модуль управления дискретными выходами системы автоматики; списывается в себестоимость",
        "Модуль управления дискретными выходами системы автоматики; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Терминально-соединительный модуль (K2.TM...)",
        "Модуль подключения и коммутации системы автоматики; списывается в себестоимость",
        "Модуль подключения и коммутации системы автоматики; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Контроллер программный (pCO mini DIN и аналоги)",
        "Контроллер системы автоматики; используется в текущих монтажных/производственных работах, списывается в себестоимость",
        "Контроллер системы автоматики; используется при строительстве, монтаже или дооборудовании объекта ОС, включается в капитальные вложения",
    ),
    (
        "Разъем для контроллера",
        "Комплектующее для подключения контроллера; списывается в себестоимость",
        "Комплектующее для подключения контроллера; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Пульт управления (P-14U)",
        "Пульт оператора и управления оборудованием; списывается в себестоимость",
        "Пульт оператора и управления оборудованием; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Блок контроллера-пусковой ШКП-3",
        "Шкафное комплектующее для автоматизации и пуска оборудования; списывается в себестоимость",
        "Шкафное комплектующее для автоматизации и пуска оборудования; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Блок управления задвижкой ШУЗ-3С",
        "Блок управления исполнительным механизмом инженерной системы; списывается в себестоимость",
        "Блок управления исполнительным механизмом инженерной системы; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Автоматический выключатель модульный",
        "Аппарат защиты электрических цепей; списывается в себестоимость",
        "Аппарат защиты электрических цепей; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Независимый расцепитель OAT/DIN BMS-6H/...",
        "Дополнительный аппарат дистанционного отключения автомата; списывается в себестоимость",
        "Дополнительный аппарат дистанционного отключения автомата; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Модуль свободных сигнальных контактов OAT/DIN BMS-MCCK2",
        "Дополнительный модуль сигнализации состояния автомата; списывается в себестоимость",
        "Дополнительный модуль сигнализации состояния автомата; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Светильник СД 36Вт IP65",
        "Светотехническое изделие для освещения шкафа, помещения или техзоны; списывается в себестоимость",
        "Светотехническое изделие для освещения шкафа, помещения или техзоны; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Лампа AD22DS LED d22 (24В)",
        "Сигнальная лампа и элемент индикации состояния оборудования; списывается в себестоимость",
        "Сигнальная лампа и элемент индикации состояния оборудования; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Корпус лотка с крышкой (IEK)",
        "Элемент кабельной трассы для прокладки и защиты кабеля; списывается в себестоимость",
        "Элемент кабельной трассы для прокладки и защиты кабеля; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Кабельный ввод в шкаф",
        "Комплектующее для ввода и фиксации кабеля в шкафу; списывается в себестоимость",
        "Комплектующее для ввода и фиксации кабеля в шкафу; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Заклепка вытяжная М8",
        "Крепежный расходный материал для сборки и монтажа; списывается в себестоимость",
        "Крепежный расходный материал для сборки и монтажа; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Болт М8",
        "Крепежный материал для сборки шкафов, конструкций и оборудования; списывается в себестоимость",
        "Крепежный материал для сборки шкафов, конструкций и оборудования; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Шайба М8",
        "Крепежный расходный материал для монтажных соединений; списывается в себестоимость",
        "Крепежный расходный материал для монтажных соединений; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Шайба пружинная М8",
        "Крепежный расходный материал для фиксации соединений; списывается в себестоимость",
        "Крепежный расходный материал для фиксации соединений; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Саморезы (Hard-Fix и аналоги)",
        "Крепежный расходный материал для монтажных работ; списывается в себестоимость",
        "Крепежный расходный материал для монтажных работ; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Маркировка для клемм/групп зажимов",
        "Маркировочный материал для идентификации клемм и цепей; списывается в себестоимость",
        "Маркировочный материал для идентификации клемм и цепей; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Маркеры кабельные (0…9, +, -, и т.д.)",
        "Маркировочный материал для кабелей и проводников; списывается в себестоимость",
        "Маркировочный материал для кабелей и проводников; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Табличка «Молния-24»",
        "Информационная табличка для обозначения оборудования или линии; списывается в себестоимость",
        "Информационная табличка для обозначения оборудования или линии; включается в капитальные вложения по объекту ОС",
    ),
    (
        "Шнур-адаптер CN_WIRE",
        "Соединительное комплектующее и сервисный адаптер; списывается в себестоимость",
        "Соединительное комплектующее и сервисный адаптер; включается в капитальные вложения по объекту ОС",
    ),
]


def excel_column(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def cell_xml(reference: str, value: str, style_id: int) -> str:
    attributes = ""
    if value.startswith(" ") or value.endswith(" ") or "\n" in value:
        attributes = ' xml:space="preserve"'
    return (
        f'<c r="{reference}" t="inlineStr" s="{style_id}">'
        f"<is><t{attributes}>{escape(value)}</t></is>"
        "</c>"
    )


def worksheet_xml(rows: list[list[str]]) -> str:
    total_rows = len(rows) + 1
    xml_rows = []

    header_cells = "".join(
        cell_xml(f"{excel_column(column)}1", value, 1)
        for column, value in enumerate(HEADER, start=1)
    )
    xml_rows.append(f'<row r="1" ht="28" customHeight="1">{header_cells}</row>')

    for row_number, row in enumerate(rows, start=2):
        row_cells = []
        for column, value in enumerate(row, start=1):
            style_id = 3 if column in (2, 3) else 2
            row_cells.append(cell_xml(f"{excel_column(column)}{row_number}", value, style_id))
        xml_rows.append(
            f'<row r="{row_number}" ht="36" customHeight="1">{"".join(row_cells)}</row>'
        )

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="55" customWidth="1"/>
    <col min="2" max="2" width="12" customWidth="1"/>
    <col min="3" max="3" width="14" customWidth="1"/>
    <col min="4" max="4" width="85" customWidth="1"/>
  </cols>
  <sheetData>
    {"".join(xml_rows)}
  </sheetData>
  <autoFilter ref="A1:D{total_rows}"/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"""


def workbook_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="00000"/>
  <workbookPr defaultThemeVersion="124226"/>
  <bookViews>
    <workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>
  </bookViews>
  <sheets>
    <sheet name="20.01" sheetId="1" r:id="rId1"/>
    <sheet name="08.03" sheetId="2" r:id="rId2"/>
  </sheets>
  <calcPr calcId="191029"/>
</workbook>
"""


def workbook_relationships_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""


def root_relationships_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def content_types_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFD9EAF7"/>
        <bgColor indexed="64"/>
      </patternFill>
    </fill>
  </fills>
  <borders count="2">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
    <border>
      <left style="thin"><color auto="1"/></left>
      <right style="thin"><color auto="1"/></right>
      <top style="thin"><color auto="1"/></top>
      <bottom style="thin"><color auto="1"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment vertical="top" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="top"/>
    </xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"""


def app_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>2</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="2" baseType="lpstr">
      <vt:lpstr>20.01</vt:lpstr>
      <vt:lpstr>08.03</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0300</AppVersion>
</Properties>
"""


def core_xml() -> str:
    timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Ведомость счетов учета для 1С</dc:title>
  <dc:creator>Cursor</dc:creator>
  <cp:lastModifiedBy>Cursor</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:modified>
</cp:coreProperties>
"""


def build_workbook() -> None:
    rows_20 = [[name, "10.08", "20.01", description_20] for name, description_20, _ in ROWS]
    rows_08 = [[name, "10.08", "08.03", description_08] for name, _, description_08 in ROWS]

    with zipfile.ZipFile(OUTPUT_PATH, "w", compression=zipfile.ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types_xml())
        workbook.writestr("_rels/.rels", root_relationships_xml())
        workbook.writestr("docProps/app.xml", app_xml())
        workbook.writestr("docProps/core.xml", core_xml())
        workbook.writestr("xl/workbook.xml", workbook_xml())
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_relationships_xml())
        workbook.writestr("xl/styles.xml", styles_xml())
        workbook.writestr("xl/worksheets/sheet1.xml", worksheet_xml(rows_20))
        workbook.writestr("xl/worksheets/sheet2.xml", worksheet_xml(rows_08))


if __name__ == "__main__":
    build_workbook()
    print(f"Created {OUTPUT_PATH.name} with {len(ROWS)} rows on each sheet.")
