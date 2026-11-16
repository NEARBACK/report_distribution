"""Build a daily distribution report from district Excel workbooks."""
import argparse
import calendar
from datetime import date
import math
from pathlib import Path
import re
import sys
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


DISHES = [
 ('Плов с мантами', 'плов|плов ман'),
 ('Плов курица', 'плов курица|плов с кур'),
 ('Плов мясо', 'плов мясо|плов с мясом|плов эт|эт плов'),
 ('Куурдак', 'курдак|куурдак'),
 ('Ман дом', 'мант дом|мант по дом|манты дом|манты по дом|по домаш'),
 ('Себет манты', 'манты себет|себет м|себет манты|себетманты'),
 ('Курица кисло сладкий', 'кур кисло|курица|курица в кс'),
 ('Бифштекс', 'бифштекс|биштекс'),
 ('Бризоль', 'бризоль'),
 ('Ган фан', 'ган фан|ганфан'),
 ('Шницель', 'шницель|шнизель'),
 ('Кебаб', 'кебаб|кебап'),
 ('Цезарь', 'цезарь'),
 ('Нежный салат', 'нежный'),
 ('Оромо', 'оромо'),
 ('Эт нан', 'эт нан'),
 ('Хошан', 'хошан'),
 ('Самсы', 'самсы кур|сл кур|сл курица|см кур'),
 ('Самсы', 'самсы сыр|сам сыр|сл сыр|под сыром'),
 ('Лагман', 'лагман|босо лагман'),
 ('Кфс', 'кфс|крылышки'),
]
DISH_INDEX = {label: i for i, (_, labels) in enumerate(DISHES) for label in labels.split('|')}
ZONES = ['Кызыл-аскер','Пригород','Пишпек','Джал','Горький','Кок-жар'] + [f'{i} зона' for i in range(7,27)]
GENERIC_SAMSA = {"самсы", "самса"}
SAMSA_FIRST = 18
SAMSA_SECOND = 19


def collect_files(inputs):
    files = set()
    for path in inputs:
        if path.is_dir():
            reports = [p for p in path.iterdir() if p.is_file() and p.suffix.lower() == ".xlsx"
                       and re.match(r"^\d+\s*[-–—]", p.name)]
            if not reports:
                raise ValueError(f"В папке {path} нет районных файлов .xlsx")
            files.update(p.resolve() for p in reports)
        elif path.is_file() and path.suffix.lower() == ".xlsx":
            files.add(path.resolve())
        else:
            raise ValueError(f"Укажите папку или районный файл .xlsx: {path}")

    districts = {}
    for path in sorted(files):
        match = re.match(r"^(\d+)\s*[-–—]", path.name)
        if match is None or not 1 <= int(match[1]) <= len(ZONES):
            raise ValueError(f"Имя файла должно начинаться с номера зоны 1–26 и дефиса: {path.name}")
        district = int(match[1])
        if district in districts:
            raise ValueError(f"Два файла для зоны {district}: {districts[district].name}, {path.name}")
        districts[district] = path
    return districts


def normalize(value):
    return " ".join(str(value or "").casefold().split())


def read_quantities(received, products, location):
    values = [None] * len(DISHES)
    generic_samsa = []
    for column, (quantity, product) in enumerate(zip(received[3:], products[3:]), start=4):
        label = normalize(product)
        index = DISH_INDEX.get(label)
        if quantity is None or (index is None and label not in GENERIC_SAMSA):
            continue
        if type(quantity) not in (int, float) or not math.isfinite(quantity) or quantity < 0:
            cell = get_column_letter(column) + "3"
            raise ValueError(f"{location}, {cell}: ожидалось количество, найдено {quantity!r}")
        if label in GENERIC_SAMSA:
            generic_samsa.append((column, quantity))
        else:
            values[index] = (values[index] or 0) + quantity

    # Preserve the two slots from the reference when the filling is unspecified.
    first, second = SAMSA_FIRST - 1, SAMSA_SECOND - 1
    for column, quantity in generic_samsa:
        index = first if values[first] is None and (values[second] is not None or column <= 38) else second
        values[index] = (values[index] or 0) + quantity
    return values


def read_reports(districts):
    periods = set()
    daily_rows = {}
    for district, path in districts.items():
        workbook = load_workbook(path, read_only=True, data_only=True)
        seen_days = set()
        try:
            for sheet in workbook:
                title = sheet.title.strip()
                if not re.fullmatch(r"[0-9]{1,2}", title):
                    continue
                day = int(title)
                location = f"{path.name}, вкладка {sheet.title}"
                metadata, _, received, products = sheet.iter_rows(max_row=4, values_only=True)
                label = normalize(received[0])
                if label != "принял" and not (district == 13 and not label and "чина" in normalize(received[2])):
                    raise ValueError(f"{location}: в строке 3 не найдена отметка «Принял»")
                dates = [value for value in metadata if isinstance(value, date)]
                if len(dates) != 1:
                    raise ValueError(f"{location}: в строке 1 должна быть одна дата Excel")
                source_date = dates[0]
                year, month = source_date.year, source_date.month
                periods.add((year, month))
                if not 1 <= day <= calendar.monthrange(year, month)[1] or day in seen_days:
                    raise ValueError(f"{location}: повторяющийся или недопустимый день")
                seen_days.add(day)
                if day != source_date.day:
                    print(f"{location}: в шапке {source_date:%d.%m}, использую день {day:02} из вкладки.",
                          file=sys.stderr)
                daily_rows.setdefault(day, {})[district] = read_quantities(received, products, location)
            if not seen_days:
                raise ValueError(f"{path.name}: нет вкладок с номерами дней")
        finally:
            workbook.close()
    if len(periods) != 1:
        raise ValueError("Все входные файлы должны относиться к одному месяцу и году")
    (year, month), = periods
    return year, month, daily_rows


def make_workbook(month, daily_rows):
    workbook = Workbook()
    workbook.remove(workbook.active)
    thin = Side(style="thin", color="DCE5EC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for day, by_district in sorted(daily_rows.items()):
        sheet = workbook.create_sheet(f"{day:02}.{month:02}")
        sheet.append(["Зоны", *(label for label, _ in DISHES)])
        end = 2 + len(ZONES)
        columns = [get_column_letter(i) for i in range(2, 2 + len(DISHES))]
        sheet.append(["ОБЩИЙ", *(f"=SUM({col}3:{col}{end})" for col in columns)])
        for district, zone in enumerate(ZONES, 1):
            sheet.append([zone, *by_district.get(district, [None] * len(DISHES))])
        sheet.freeze_panes = "A3"
        sheet.sheet_view.showGridLines = False
        sheet.column_dimensions["A"].width = 17
        for col in columns:
            sheet.column_dimensions[col].width = 12
        sheet.column_dimensions["H"].width = 17
        for row in sheet.iter_rows():
            row_number = row[0].row
            sheet.row_dimensions[row_number].height = 34 if row_number == 1 else 17
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="center", wrap_text=row_number == 1,
                                           horizontal="left" if cell.column == 1 else "center")
                cell.font = Font(name="Arial", size=10, bold=row_number <= 2,
                                 color="FFFFFF" if row_number == 1 else "172B3D")
                if row_number <= 2:
                    cell.fill = PatternFill("solid", fgColor="183D58" if row_number == 1 else "D9EAF4")
    return workbook


def save_xlsx(workbook, output):
    output.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix=".report-", dir=output.parent) as directory:
        report = Path(directory) / "report.xlsx"
        workbook.save(report)
        # Publish the finished workbook without overwriting an existing file.
        output.hardlink_to(report)


def main():
    parser = argparse.ArgumentParser(description="Собрать дневные таблицы «Принял» по 26 зонам в один .xlsx.")
    parser.add_argument("inputs", nargs="+", type=Path, metavar="ВХОД",
                        help="Папка с отчётами или один либо несколько районных файлов .xlsx")
    parser.add_argument("-o", "--output", type=Path, help="Путь к итоговому .xlsx")
    args = parser.parse_args()
    try:
        districts = collect_files(args.inputs)
        year, month, daily_rows = read_reports(districts)
        output = (args.output or Path(f"Раздача {year}-{month:02}.xlsx")).absolute()
        if output.suffix.lower() != ".xlsx":
            raise ValueError("У итогового файла должно быть расширение .xlsx")
        if output.exists():
            raise ValueError(f"Файл уже существует: {output}. Укажите новый путь через -o.")
        save_xlsx(make_workbook(month, daily_rows), output)
    except (OSError, ValueError) as exc:
        parser.exit(1, f"Ошибка: {exc}\n")
    print(f"{output}\nЗон с исходными файлами: {len(districts)}. Дневных вкладок: {len(daily_rows)}.")


if __name__ == "__main__":
    main()
