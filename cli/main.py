#!/usr/bin/env python
import typer
import sys
import subprocess
from pathlib import Path
from typing import List, Literal
from .banner import print_banner
from .colors import console

# IMPORT ONLY WHAT EXISTS IN YOUR ORIGINAL FILES

from .olevel.processDS import OlevelProcessor
from .olevel.resultImport import OlevelResultImporter
from .olevel.insert import OlevelStudentImporter
from .olevel.combiner import DualExamProcessor as DualOlevelProcessor
from .alevel.combiner import DualAlevelProcessor as DualAlevelProcessor

from .alevel.ranking import AlevelProcessor
from .alevel.insert import AlevelStudentImporter
from .alevel.importer import ExamDataImporter
from .alevel.exporter import StudentExamExporter

from .primary.insert import PrimaryPupilImporter
from .primary.resultImport import PrimaryResultImporter
from .primary.process import ExamProcessor as PrimaryProcessor


from . import testing
from .time_travel import TimeTravelProcessor

from .bulk import BulkSMSSender

app = typer.Typer(
    name="kiyabo",
    help="KIYABO APP – Tanzania School Information System",
    add_completion=False,
)

Level = Literal["primary", "olevel", "alevel"]

@app.command()
def test(level: Level = typer.Argument(..., help="School level")):
    print_banner()
    testing.main()

@app.command()
def upload(
    level: Level = typer.Argument(..., help="School level"),
    exam_id: str = typer.Option(..., "--exam-id"),
    excel_path: str = typer.Option(..., "--excel"),
    db_path: str = typer.Option(r"C:\Kiyabo App\backend\Kiyabo App Backend v4.0.0.accdb", "--db"),
    process_after: bool = typer.Option(False, "--process", is_flag=True, help="Process after completion"),
    sort_columns: str = typer.Option(None, "--sort-cols", help="Columns to sort by (comma-separated)")
):
    print_banner()
    print(f"Uploading {level} data for exam ID {exam_id} from {excel_path} to {db_path} process_after={process_after}")
    if level == "alevel":
        importer = ExamDataImporter()
        success = importer.import_exam_data(exam_id, excel_path, db_path)
        console.print("[green]UPLOAD SUCCESS[/]" if success else "[red]UPLOAD FAILED[/]")
        if success and process_after:
            processor = AlevelProcessor(exam_id=exam_id, 
                        db_path=db_path, sort_columns=[col.strip() for col in sort_columns.split(",")] if sort_columns else None,
                        )
            processor.run()
            console.print("[green]PROCESSING COMPLETE[/]")
    
    elif level == "olevel":
        importer = OlevelResultImporter(exam_id=exam_id,excel_file=excel_path,db_path=db_path,force_import=True,process_after=process_after)
        importer.run()
        console.print("[green]UPLOAD COMPLETE[/]")
    
    elif level == "primary":
        importer=PrimaryResultImporter(db_path=db_path,excel_path=excel_path,exam_id=exam_id)
        success=importer.run()
        if success and process_after:
            processor=PrimaryProcessor(db_path=db_path,exam_id=exam_id)
            processor.complete_exam()
    else:
        console.print(f"[yellow]upload not implemented for {level}[/]")



@app.command(name="import")
def import_(
    level: Level = typer.Argument(..., help="School level (primary, olevel, alevel)"),
    class_id: str = typer.Argument(..., help="Class ID (I, II, III, IV, V, VI, PC)"),
    excel_path: str = typer.Option(..., "--excel", "-e", help="Path to Excel file"),
    db_path: str = typer.Option(..., "--db", "-d", help="Path to database file"),
    save_folder: str = typer.Option(r"C:\Kiyabo App\admission", "--save-folder", "-s", help="Folder to save processed files"),
    academic_year: int = typer.Option(None, "--academic-year", "-y", help="Academic year (defaults to current year if not specified)"),
):
    level=level.lower()
    class_id=class_id.upper()
    print_banner()

    if level=="olevel":
        importer=OlevelStudentImporter(excel_path=excel_path,db_path=db_path,save_folder=save_folder,class_id=class_id)
        importer.run()

    elif level=="alevel":
        importer=AlevelStudentImporter(excel_path=excel_path,db_path=db_path,save_folder=save_folder,class_id=class_id)
        importer.run()

    elif level=="primary":
        importer=PrimaryPupilImporter(excel_path=excel_path,db_path=db_path,save_folder=save_folder,class_id=class_id,academic_year=academic_year)
        importer.run()
        
    else:
        console.print(f"[yellow]import not implemented for {level}[/]")




@app.command()
def export(
    level: Level = typer.Argument(..., help="School level"),
    exam_id: str = typer.Option(..., "--exam-id"),
    db_path: str = typer.Option(r"C:\Kiyabo App\backend\Kiyabo App Backend v4.0.0.accdb", "--db"),
    include_comb: bool = typer.Option(True, "--comb/--no-comb"),
    top_n: int = typer.Option(10, "--top"),
    bottom_n: int = typer.Option(10, "--bottom"),
    order_by: str = typer.Option("position", "--order-by"),
    paper_size: str = typer.Option("A4", "--paper-size"),
    orientation: str = typer.Option(None, "--orientation"),
):
    print_banner()
    level = level.lower()

    if level == "alevel":
        exporter=StudentExamExporter(
            exam_id=exam_id,
            db_path=db_path,
            include_comb_sheets=include_comb,
            order_by=order_by,
            top_n=top_n,
            bottom_n=bottom_n,
            paper_size=paper_size,
            orientation=orientation,
        )
        exporter.run()
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
    flat_rate:bool=typer.Option(False,"--flat/--no-flat",help="Will yiu take Top Bests Subjects or Average All?")
):
    print_banner()
    try:
        level = level.lower()
        if level == "alevel":
            if sort_cols:
                sort_cols_list = [col.strip() for col in sort_cols.split(",")]
            else:
                sort_cols_list = None

            processor = AlevelProcessor(
                exam_id=exam_id,
                db_path=db_path,
                sort_columns=sort_cols_list,
                include_inc=include_inc
                )
            processor.run()
        
        elif level == "olevel":
            if sort_cols:
                sort_cols_list = [col.strip() for col in sort_cols.split(",")]
            else:
                sort_cols_list = None

            processor = OlevelProcessor(exam_id=exam_id, db_path=db_path,sort_columns=sort_cols_list,include_inc=include_inc,flat_rate=flat_rate)
            processor.run()
            console.print("[green]PROCESSING COMPLETE[/]")
        
        elif level == "primary":
            processor = PrimaryProcessor(exam_id=exam_id, db_path=db_path)
            processor.complete_exam()
            console.print("[green]PROCESSING COMPLETE[/]")
        else:
            console.print(f"[yellow]process not implemented for {level}[/]")
        

    except Exception as e:
        console.print(f"[red]ERROR:[/] {e}")

@app.callback(invoke_without_command=True)
def main(ctx: typer.Context):
    if ctx.invoked_subcommand is None:
        print_banner()
        console.print("\n[bold magenta]KIYABO APP[/] — type `kiyabo --help` for commands.\n")


@app.command(context_settings={"allow_extra_args": True, "ignore_unknown_options": True})
def run(file_path: str, args: List[str] = typer.Argument(None)):
    """Run a Python script with arguments"""
    print_banner()
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
    print_banner()
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


@app.command(name='future')
def future_from_antique(
    level: str = typer.Argument(..., help="School level"),
    db_path: str = typer.Option(None, "--db"),
):
    """Run future_from_antique to update old Kiyabo App databases"""
    print_banner()
    if db_path is None:
        raise typer.BadParameter("Database path is required", param_hint="--db")
    
    try:
        space_ship = TimeTravelProcessor(level=level, db_path=db_path)
        space_ship.into_the_future()
    except Exception as e:
        console.print(f"[red]ERROR:[/] {e}")


@app.command("sendsms")
def bulk_sms(
    level: str = typer.Argument(..., help="School level"),
    excel_path: str = typer.Option(..., "--excel", "-e", help="Path to Excel file with SMS data"),
    mpe_exe_path: str = typer.Option(r"C:\Program Files (x86)\MyPhoneExplorer\MyPhoneExplorer.exe", "--mpe", "-m", help="Path to MyPhoneExplorer.exe"),
    mode: str = typer.Option("direct", "--mode", "-M", help="Sending mode: 'direct' for one-by-one, 'batch' for XML batches"),
    batch_size: int = typer.Option(1, "--batch-size", "-b", help="Batch size for 'batch' mode"),
    min_sleep: int = typer.Option(10, "--min-sleep","-min", help="Minimum sleep time between messages (seconds)"),
    max_sleep: int = typer.Option(20, "--max-sleep", "-max", help="Maximum sleep time between messages (seconds)"),
):
    """Send bulk SMS using MyPhoneExplorer"""
    print_banner()
    sender = BulkSMSSender(
        excel_path=excel_path,
        mpe_exe_path=mpe_exe_path,
        mode=mode,
        batch_size=batch_size,
        min_sleep=min_sleep,
        max_sleep=max_sleep
    )
    sender.run()



@app.command("combine")
def combine(
    level: str = typer.Argument(..., help="School level (e.g. Form4, O-Level, A-Level)"),

    e1: str = typer.Option(..., "--exam1", "-e1", help="First exam ID"),
    e2: str = typer.Option(..., "--exam2", "-e2", help="Second exam ID"),
    db: str = typer.Option(..., "--db", "-d", help="Path to Access database (.accdb)"),

    exam_name_1: str = typer.Option(None, "--exam1-name", help="Name of the first exam"),
    exam_name_2: str = typer.Option(None, "--exam2-name", help="Name of the second exam"),

    query_name: str = typer.Option("qry_CombinedExamResults", "--query-name", "-q", help="Name of the query to create"),
    base_subjects: int = typer.Option(7, "--base-subjects", "-b", help="Number of base/best subjects"),
    flat_rate: bool = typer.Option(True, "--flat-rate/--no-flat-rate", help="Use flat division rate instead of weighted"),
    include_inc: bool = typer.Option(True, "--include-inc/--no-include-inc", help="Include INC subjects in average calculation"),
    ranking_method: str = typer.Option("min", "--ranking", "-r", help="Ranking method: 'min', 'max' or 'total'"),
    necta_decimal_places: int = typer.Option(1, "--necta-dp", help="Decimal places for NECTA-style division (0 or 1)"),
    sort_columns: str = typer.Option(None, "--sort-cols", "-s", help="Columns to sort by (comma-separated)"),
):
    """
    Combine two exams and save results as Access query.
    """
    print_banner()
    level=level.lower()

    if sort_columns:
        sort_columns_list = [col.strip() for col in sort_columns.split(",")]
    else:
        sort_columns_list = None

    if level=="olevel":
        processor = DualOlevelProcessor(
            exam_id_1=e1,
            exam_id_2=e2,
            exam_name_1=exam_name_1,
            exam_name_2=exam_name_2,
            db_path=db,
            query_name=query_name,
            base_subjects=base_subjects,
            flat_rate=flat_rate,
            include_inc=include_inc,
            ranking_method=ranking_method,
            necta_decimal_places=necta_decimal_places,
            sort_columns=sort_columns_list
        )
        processor.run() 
    
    elif level=="alevel":
        processor = DualAlevelProcessor(
            exam_id_1=e1,
            exam_id_2=e2,
            exam_name_1=exam_name_1,
            exam_name_2=exam_name_2,
            db_path=db,
            sort_columns=sort_columns_list,
            include_inc=include_inc,
            rank_method=ranking_method,
        )
        processor.run()
    else:
        console.print(f"[yellow]combine not implemented for {level}[/]")
    

if __name__ == "__main__":
    app()