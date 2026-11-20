"""Structural baseline and synthetic capacity migration."""
import json
import re
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import get_column_letter, range_boundaries, rows_from_range
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableColumn

BASELINE = Path("local") / "workbook-baseline.json"
FIRST_ROW = 5
OLD_END = 104
NEW_END = 1004
MARKER = "SYNTHETIC CAPACITY TEST ONLY"
RANGE_END = re.compile(r"(\$?[A-Z]{1,3}\$?5:\$?[A-Z]{1,3}\$?)104\b")


def schema(book):
    result = {}
    for sheet in book:
        tables = {}
        for table in sheet.tables.values():
            tables[table.name] = {
                "ref": table.ref,
                "filter": table.autoFilter.ref if table.autoFilter else None,
                "headers": [column.name for column in table.tableColumns],
            }
        result[sheet.title] = {
            "state": sheet.sheet_state,
            "headers": [cell.value for cell in sheet[4] if cell.value is not None],
            "protected": sheet.protection.sheet,
            "tables": tables,
            "formulas": {
                cell.coordinate: cell.value
                for row in sheet.iter_rows()
                for cell in row
                if cell.data_type == "f"
            },
            "validation": [dict(rule) | {"formula1": rule.formula1, "formula2": rule.formula2}
                           for rule in sheet.data_validations.dataValidation],
        }
    return result


def capture_baseline(source):
    """Read the reference workbook and save only its structure locally."""
    book = load_workbook(source, keep_links=False)
    try:
        structure = schema(book)
        counts = tuple(sum(len(entry[key]) for entry in structure.values())
                       for key in ("tables", "formulas", "validation"))
        if len(structure) != 21 or counts != (19, 1116, 21):
            raise ValueError("Reference does not match the reviewed V2 structure counts")
        BASELINE.parent.mkdir(exist_ok=True)
        with BASELINE.open("x", encoding="utf-8") as output:
            json.dump(structure, output, ensure_ascii=False, indent=2)
    finally:
        book.close()


def baseline():
    return json.loads(BASELINE.read_text(encoding="utf-8"))


def extend_range(reference):
    return re.sub(r"104$", str(NEW_END), reference)


def expanded_formulas(formulas):
    result = {}
    for address, formula in formulas.items():
        # Only the template's fixed source ranges are widened.
        result[address] = RANGE_END.sub(r"\g<1>1004", formula)
    for address, formula in list(result.items()):
        if address.endswith(str(OLD_END)):
            column = address[:-len(str(OLD_END))]
            translator = Translator(formula, origin=address)
            for row in range(OLD_END + 1, NEW_END + 1):
                target = f"{column}{row}"
                result[target] = translator.translate_formula(target)
    return result


def expected_schema(expanded=False):
    expected = baseline()
    if not expanded:
        return expected
    intake_sheet = list(expected)[5]
    for name, entry in expected.items():
        entry["formulas"] = expanded_formulas(entry["formulas"])
        for rule in entry["validation"]:
            rule["sqref"] = extend_range(rule["sqref"])
        for table in entry["tables"].values():
            table["ref"] = extend_range(table["ref"])
            if table["filter"]:
                table["filter"] = extend_range(table["filter"])
            if name == intake_sheet:
                table["ref"] = "A4:R1004"
                if table["filter"]:
                    table["filter"] = "A4:R1004"
                table["headers"].append("Workfront ID")
                entry["headers"].append("Workfront ID")
    return expected


def validate(book, expanded=False):
    expected = expected_schema(expanded)
    observed = schema(book)
    if list(observed) != list(expected):
        raise ValueError("Worksheet names or order differ from baseline")
    for name, entry in expected.items():
        sheet = book[name]
        for field, value in entry.items():
            if observed[name][field] != value:
                raise ValueError(f"{name}: {field} differs from baseline")
        for table in entry["tables"].values():
            first, top, last, _ = range_boundaries(table["ref"])
            if [sheet.cell(top, column).value for column in range(first, last + 1)] != table["headers"]:
                raise ValueError(f"{name}: worksheet headers differ from table metadata")
        for rule in entry["validation"]:
            allowed = rule["formula1"][1:-1].split(",")
            for address, in rows_from_range(rule["sqref"]):
                value = sheet[address].value
                if value is not None and value not in allowed:
                    raise ValueError(f"{name}!{address}: invalid list value")
    if book.worksheets[1]["A1"].value != MARKER:
        raise ValueError("Only generated synthetic workbooks are accepted")


def make_fixture(path):
    book = Workbook()
    book.remove(book.active)
    for name, entry in baseline().items():
        sheet = book.create_sheet(name)
        for column, header in enumerate(entry["headers"], 1):
            sheet.cell(4, column, header)
        for table_name, metadata in entry["tables"].items():
            table = Table(displayName=table_name, ref=metadata["ref"])
            table._initialise_columns()
            if not metadata["filter"]:
                table.autoFilter = None
            for column, header in zip(table.tableColumns, metadata["headers"]):
                column.name = header
            sheet.add_table(table)
        for address, formula in entry["formulas"].items():
            sheet[address] = formula
        for rule in entry["validation"]:
            sheet.add_data_validation(DataValidation(**rule))
    book.worksheets[1]["A1"] = MARKER
    validate(book)
    with Path(path).open("xb") as output:
        book.save(output)


def values(book):
    return {sheet.title: {cell.coordinate: (cell.data_type, cell.value)
                         for row in sheet.iter_rows() for cell in row
                         if cell.value is not None and cell.data_type != "f"}
            for sheet in book}


def stage_migration(source, destination):
    book = load_workbook(source)
    try:
        validate(book)
        original_values = values(book)
        expected = expected_schema(expanded=True)
        for sheet in book:
            entry = expected[sheet.title]
            for address, formula in entry["formulas"].items():
                if sheet[address].value != formula:
                    sheet[address] = formula
            for table in sheet.tables.values():
                metadata = entry["tables"][table.name]
                first, _, last, _ = range_boundaries(metadata["ref"])
                for column in range(first, last + 1):
                    template = sheet.cell(OLD_END, column)
                    for row in range(OLD_END + 1, NEW_END + 1):
                        cell = sheet.cell(row, column)
                        if template.has_style:
                            cell._style = copy(template._style)
                table.ref = metadata["ref"]
                if table.autoFilter:
                    table.autoFilter.ref = metadata["filter"]
                if sheet == book.worksheets[5]:
                    sheet.cell(4, last, "Workfront ID")
                    table.tableColumns.append(TableColumn(id=last, name="Workfront ID"))
                    original_values[sheet.title][f"{get_column_letter(last)}4"] = ("s", "Workfront ID")
            for rule in sheet.data_validations.dataValidation:
                rule.sqref = extend_range(str(rule.sqref))
        book.save(destination)
    finally:
        book.close()
    check = load_workbook(destination)
    try:
        validate(check, expanded=True)
        if values(check) != original_values:
            raise ValueError("Migration changed existing non-formula values")
    finally:
        check.close()
