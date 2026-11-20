"""Build disposable local demo data for the five synthetic reports."""
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook

from .journal import Journal
from .workbook import make_fixture, stage_migration
from .workfront import create


REPORTS = Path(__file__).resolve().parent.parent / "examples" / "five-synthetic-reports.txt"


def seed(directory):
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    workbook = directory / "synthetic-five.xlsx"
    database = directory / "audit.sqlite3"
    if workbook.exists() or database.exists():
        raise ValueError("Choose an empty demo directory")
    journal = Journal(database)
    workfronts = {}
    for name, activity in (("North", "SYN-CUT"), ("East", "SYN-DRAIN"), ("West", "SYN-TOOL"),
                           ("South", "SYN-CROSS"), ("Central", "SYN-ACCESS")):
        workfronts[name] = create(journal, "Synthetic Setup", "Mock Section", f"Mock {name} Zone",
                                  [activity], [f"Mock {name}"])
    with TemporaryDirectory(dir=directory) as temp:
        base = Path(temp) / "fixture.xlsx"
        expanded = Path(temp) / "expanded.xlsx"
        make_fixture(base)
        stage_migration(base, expanded)
        book = load_workbook(expanded)
        try:
            task = book.worksheets[20]
            task["A5"] = "SYN-TASK-1"
            task["C5"] = workfronts["East"]
            task["I5"] = "Open"
            with workbook.open("xb") as output:
                book.save(output)
        finally:
            book.close()
    return workbook, journal


def main():
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory")
    args = parser.parse_args()
    workbook, journal = seed(args.directory)
    print(f"python -m kvd.desktop --workbook \"{workbook}\" --journal \"{journal.path}\"")
    print(f"Paste the contents of {REPORTS.name} from the examples folder.")


if __name__ == "__main__":
    main()
