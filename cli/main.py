#!/usr/bin/env python
import typer
import sys
import subprocess
from pathlib import Path
from typing import List, Literal
from .banner import print_banner
from .colors import console

# IMPORT ONLY WHAT EXISTS IN YOUR ORIGINAL FILES
from .alevel.ranking import process_exam as alevel_process_exam
from .olevel.processDS import OlevelProcessor
from .olevel.resultImport import OlevelResultImporter
from .olevel.insert import OlevelStudentImporter
from .alevel.insert import AlevelStudentImporter
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
        importer = OlevelResultImporter(exam_id=exam_id,excel_file=excel_path,db_path=db_path,force_import=True,process_after=process_after)
        importer.run()
        console.print("[green]UPLOAD COMPLETE[/]")
    else:
        console.print(f"[yellow]upload not implemented for {level}[/]")



@app.command(name="import")
def import_(
    level: Level = typer.Argument(..., help="School level (primary, olevel, alevel)"),
    class_id: str = typer.Argument(..., help="Class ID (I, II, III, IV, V, VI, PC)"),
    excel_path: str = typer.Option(..., "--excel", "-e", help="Path to Excel file"),
    db_path: str = typer.Option(..., "--db", "-d", help="Path to database file"),
    save_folder: str = typer.Option(r"C:\Kiyabo App\admission", "--save-folder", "-s", help="Folder to save processed files"),
):
    level=level.lower()
    class_id=class_id.upper()

    if level=="olevel":
        importer=OlevelStudentImporter(excel_path=excel_path,db_path=db_path,save_folder=save_folder,class_id=class_id)
        importer.run()
    elif level=="alevel":
        importer=AlevelStudentImporter(excel_path=excel_path,db_path=db_path,save_folder=save_folder,class_id=class_id)
        importer.run()
    else:
        console.print(f"[yellow]import not implemented for {level}[/]")



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
    sort_cols: str=typer.Option(None, "--sort-cols", help="Columns to sort by (comma-separated)"),
):
    print_banner()
    level = level.lower()
    if level == "alevel":
        alevel_process_exam(exam_id, db_path, include_INC=include_inc)
        console.print("[green]PROCESSING COMPLETE[/]")
    
    elif level == "olevel":
        if sort_cols:
            sort_cols_list = [col.strip() for col in sort_cols.split(",")]
        else:
            sort_cols_list = None

        processor = OlevelProcessor(exam_id=exam_id, db_path=db_path,sort_columns=sort_cols_list,include_inc=include_inc)
        processor.run()
        console.print("[green]PROCESSING COMPLETE[/]")
    else:
        console.print(f"[yellow]process not implemented for {level}[/]")

@app.callback(invoke_without_command=True)
def main(ctx: typer.Context):
    if ctx.invoked_subcommand is None:
        print_banner()
        console.print("\n[bold magenta]KIYABO APP[/] — type `kiyabo --help` for commands.\n")


@app.command(context_settings={"allow_extra_args": True, "ignore_unknown_options": True})
def run(file_path: str, args: List[str] = typer.Argument(None)):
    """Run a Python script with arguments"""
    script_path = Path(file_path)
    if not script_path.exists():
        typer.echo(f"Error: File {file_path} not found")
        raise typer.Exit(1)
    
    # Build command: python script.py [args...]
    command = [sys.executable, str(script_path)]
    if args:
        command.extend(args)
    
    try:
        subprocess.run(command, check=True)
    except subprocess.CalledProcessError as e:
        typer.echo(f"Script execution failed with exit code {e.returncode}")
        raise typer.Exit(e.returncode)


@app.command(context_settings={"allow_extra_args": True, "ignore_unknown_options": True})
def python(args: List[str] = typer.Argument(None)):
    """Run Python commands (pip, python -m, etc.)"""
    if not args:
        # Start interactive Python if no args
        subprocess.run([sys.executable])
    else:
        command = [sys.executable] + args
        try:
            subprocess.run(command, check=True)
        except subprocess.CalledProcessError as e:
            typer.echo(f"Command failed with exit code {e.returncode}")
            raise typer.Exit(e.returncode)


@app.command(context_settings={"allow_extra_args": True, "ignore_unknown_options": True})
def pip(args: List[str] = typer.Argument(None)):
    """Run pip commands"""
    print_banner()
    if not args:
        typer.echo("Please specify pip command (install, list, etc.)")
        return
    
    command = [sys.executable, "-m", "pip"] + args
    try:
        subprocess.run(command, check=True)
    except subprocess.CalledProcessError as e:
        typer.echo(f"pip command failed with exit code {e.returncode}")
        raise typer.Exit(e.returncode)
    

if __name__ == "__main__":
    app()