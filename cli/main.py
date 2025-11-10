#!/usr/bin/env python
import typer
from typing import Literal
from .banner import print_banner
from .colors import console

# IMPORT ONLY WHAT EXISTS IN YOUR ORIGINAL FILES
from .alevel.ranking import process_exam as alevel_process_exam
from .olevel.processDS import OlevelProcessor
from .olevel.resultImport import OlevelResultImporter
from .alevel.importer import ExamDataImporter
from .alevel.exporter import StudentExamExporter
from . import testing

app = typer.Typer(
    name="kiyabo",
    help="KIYABO APP – Tanzania School Information System",
    add_completion=False,
)

Level = Literal["primary", "olevel", "alevel"]

@app.command()
def test(level: Level = typer.Argument(..., help="School level")):
    testing.main()

@app.command()
def upload(
    level: Level = typer.Argument(..., help="School level"),
    exam_id: str = typer.Option(..., "--exam-id"),
    excel_path: str = typer.Option(..., "--excel"),
    db_path: str = typer.Option(r"C:\Kiyabo App\backend\Kiyabo App Backend v4.0.0.accdb", "--db"),
    process_after: bool = typer.Option(False, "--process", help="Process after completion")

):
    if level == "alevel":
        importer = ExamDataImporter()
        success = importer.import_exam_data(exam_id, excel_path, db_path)
        console.print("[green]UPLOAD SUCCESS[/]" if success else "[red]UPLOAD FAILED[/]")
        if success and process_after:
            process(level=level, exam_id=exam_id, db_path=db_path)
    
    elif level == "olevel":
        importer = OlevelResultImporter(exam_id=exam_id,excel_file=excel_path,db_path=db_path)
        importer.run()
        console.print("[green]UPLOAD COMPLETE[/]")
    else:
        console.print(f"[yellow]upload not implemented for {level}[/]")

@app.command()
def export(
    level: Level = typer.Argument(..., help="School level"),
    exam_id: str = typer.Option(..., "--exam-id"),
    db_path: str = typer.Option(r"C:\Kiyabo App\backend\Kiyabo App Backend v4.0.0.accdb", "--db"),
    include_comb: bool = typer.Option(False, "--comb"),
    top_n: int = typer.Option(10, "--top"),
    bottom_n: int = typer.Option(10, "--bottom"),
):
    if level == "alevel":
        StudentExamExporter(
            exam_id=exam_id,
            db_path=db_path,
            include_comb_sheets=include_comb,
            top_n=top_n,
            bottom_n=bottom_n,
        )
        console.print("[green]EXPORT STARTED → C:\\Kiyabo App\\Results[/]")
    else:
        console.print(f"[yellow]export not implemented for {level}[/]")

@app.command()
def process(
    level: Level = typer.Argument(..., help="School level"),
    exam_id: str = typer.Option(..., "--exam-id"),
    db_path: str = typer.Option(r"C:\Users\droge\OneDrive\Documents\Kiyabo App Backend v4.0.0.accdb", "--db"),
    include_inc: bool = typer.Option(True, "--inc/--no-inc"),
):
    level = level.lower()
    if level == "alevel":
        alevel_process_exam(exam_id, db_path, include_INC=include_inc)
        console.print("[green]PROCESSING COMPLETE[/]")
    
    elif level == "olevel":
        processor = OlevelProcessor(exam_id=exam_id, db_path=db_path)
        processor.run()
        console.print("[green]PROCESSING COMPLETE[/]")
    else:
        console.print(f"[yellow]process not implemented for {level}[/]")

@app.callback(invoke_without_command=True)
def main(ctx: typer.Context):
    if ctx.invoked_subcommand is None:
        print_banner()
        console.print("\n[bold magenta]KIYABO APP[/] — type `kiyabo --help` for commands.\n")

if __name__ == "__main__":
    app()