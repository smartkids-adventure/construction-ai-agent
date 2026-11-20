"""Windows-compatible command-line interface for P0 and P1."""
import argparse
import json

from openpyxl import load_workbook

from .journal import Journal, fingerprint, migrate
from .workbook import capture_baseline, make_fixture, validate
from .workfront import create, resolve, retire


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--journal", default="audit.sqlite3")
    commands = parser.add_subparsers(dest="command", required=True)
    for name in ("baseline", "fixture", "validate"):
        command = commands.add_parser(name)
        command.add_argument("workbook")
        if name == "validate":
            command.add_argument("--expanded", action="store_true")
    command = commands.add_parser("migrate")
    command.add_argument("source")
    command.add_argument("output")
    command.add_argument("--actor", required=True)
    command.add_argument("--approved-sha", required=True)
    commands.add_parser("recover").add_argument("--actor", required=True)
    command = commands.add_parser("workfront-add")
    for name in ("actor", "section", "location"):
        command.add_argument(f"--{name}", required=True)
    command.add_argument("--activity", action="append", default=[])
    command.add_argument("--alias", action="append", required=True)
    commands.add_parser("workfront-find").add_argument("lookup")
    command = commands.add_parser("workfront-retire")
    command.add_argument("identifier")
    command.add_argument("--actor", required=True)
    args = parser.parse_args()
    try:
        if args.command == "baseline":
            capture_baseline(args.workbook)
        elif args.command == "fixture":
            make_fixture(args.workbook)
            print(fingerprint(args.workbook))
        elif args.command == "validate":
            book = load_workbook(args.workbook)
            try:
                validate(book, expanded=args.expanded)
                print(fingerprint(args.workbook))
            finally:
                book.close()
        else:
            journal = Journal(args.journal)
            if args.command == "migrate":
                print(migrate(args.source, args.output, journal, args.actor, args.approved_sha))
            elif args.command == "recover":
                print(json.dumps(journal.recover(args.actor)))
            elif args.command == "workfront-add":
                print(create(journal, args.actor, args.section, args.location,
                             args.activity, args.alias))
            elif args.command == "workfront-find":
                print(json.dumps(resolve(journal, args.lookup), ensure_ascii=False))
            else:
                retire(journal, args.actor, args.identifier)
    except (ValueError, OSError) as error:
        parser.exit(1, f"{error}\n")


if __name__ == "__main__":
    main()
