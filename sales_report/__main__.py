"""Command-line entry point:  python -m sales_report --input data/raw --output output/report.xlsx"""
import argparse
import logging
import sys
from datetime import date
from pathlib import Path

from .cleaner import load_all
from .report import build_report


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Clean sales exports and build an Excel report.")
    parser.add_argument("--input", "-i", type=Path, default=Path("data/raw"),
                        help="folder containing CSV/Excel exports (default: data/raw)")
    parser.add_argument("--output", "-o", type=Path,
                        default=Path("output") / f"sales_report_{date.today():%Y%m%d}.xlsx",
                        help="where to save the report")
    parser.add_argument("--quiet", "-q", action="store_true", help="only print errors")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.WARNING if args.quiet else logging.INFO, format="%(message)s")
    try:
        df, stats = load_all(args.input)
        path = build_report(df, stats, args.output)
    except (FileNotFoundError, ValueError) as err:
        logging.error("Error: %s", err)
        return 1
    logging.info("Report saved to %s (%d clean rows, $%s revenue)", path, len(df), f"{df['revenue'].sum():,.2f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
