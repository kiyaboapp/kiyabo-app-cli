# run_cli.py  (place next to cli/, python32/, python64/)
import os
import sys
import pathlib

# ------------------------------------------------------------
# 1. Make the project root importable
ROOT = pathlib.Path(__file__).parent.resolve()
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# ------------------------------------------------------------
# 2. Import and run the Typer app
try:
    from cli.main import app
    app()
except Exception as e:
    import traceback, rich
    console = rich.get_console()
    console.print("[bold red]FATAL ERROR[/]")
    console.print(traceback.format_exc())
    sys.exit(1)