#!/usr/bin/env python3
"""Entrada principal: interface PySide6 por padrão e inventário via CLI."""

from __future__ import annotations

import sys


def main() -> None:
    args = sys.argv[1:]

    if args and args[0] in {"gui", "--gui"}:
        del sys.argv[1]
        from drive_inventory_gui import main as gui_main

        gui_main()
        return

    if args and args[0] == "--smoke-test-gui":
        del sys.argv[1]
        from drive_inventory_gui import main as gui_main

        gui_main(smoke_test=True)
        return

    if args and args[0] in {"cli", "--cli"}:
        del sys.argv[1]
        from drive_inventory import main as cli_main

        cli_main()
        return

    if args:
        from drive_inventory import main as cli_main

        cli_main()
        return

    from drive_inventory_gui import main as gui_main

    gui_main()


if __name__ == "__main__":
    main()
