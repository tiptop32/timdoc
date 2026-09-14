from __future__ import annotations

import argparse

from timdoc_app.desktop import run_desktop


def main() -> None:
    parser = argparse.ArgumentParser(description="Timdoc: заполнение документов по акту")
    parser.add_argument("--self-test", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.self_test:
        from timdoc_app.facade import create_default_facade

        create_default_facade().bootstrap()
        print("Timdoc self-test: OK")
        return
    run_desktop()


if __name__ == "__main__":
    main()
