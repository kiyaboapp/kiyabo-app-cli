# build.py
import subprocess
from pathlib import Path
import argparse
import shutil

def clean_previous_builds(onefile=False):
    """Remove only necessary previous build files"""
    # Always clean build directory (safe to remove)
    if Path("build").exists():
        print("Cleaning build directory...")
        shutil.rmtree("build")
    
    # Only clean dist if we're changing modes to avoid conflicts
    if onefile:
        # Remove previous onefile builds to avoid confusion
        if Path("dist/Kiyabo32.exe").exists():
            Path("dist/Kiyabo32.exe").unlink(missing_ok=True)
        if Path("dist/Kiyabo64.exe").exists():
            Path("dist/Kiyabo64.exe").unlink(missing_ok=True)
    else:
        # Remove previous onedir builds to avoid confusion  
        if Path("dist/Kiyabo32").exists():
            shutil.rmtree("dist/Kiyabo32", ignore_errors=True)
        if Path("dist/Kiyabo64").exists():
            shutil.rmtree("dist/Kiyabo64", ignore_errors=True)
    
    # Clean spec files (safe to regenerate)
    spec_files = ["Kiyabo32.spec", "Kiyabo64.spec"]
    for spec_file in spec_files:
        if Path(spec_file).exists():
            print(f"Removing {spec_file}...")
            Path(spec_file).unlink(missing_ok=True)

def main():
    parser = argparse.ArgumentParser(description='Build Kiyabo executables')
    parser.add_argument('--onefile', action='store_true', default=False,
                       help='Build as single executable file (default: False)')
    parser.add_argument('--all', action='store_true', default=False,
                       help='Build all 32/64 onefile and onedir versions')
    
    args = parser.parse_args()
    
    if args.all:
        # Build all 4 combinations: 32/64 onefile/onedir
        print("Building ALL versions (32/64 onefile/onedir)...")
        
        # Clean everything for all build
        if Path("build").exists():
            shutil.rmtree("build")
        if Path("dist").exists():
            shutil.rmtree("dist")
        
        # Build 32-bit onefile
        print("Building Kiyabo32 (onefile)...")
        build_cmd = ["python32/python.exe", "-m", "PyInstaller", "--clean", "--noconfirm", "--noupx", "--onefile", "--icon=icon.ico", "--name=Kiyabo32", "--hidden-import=win32timezone", "run_cli.py"]
        # subprocess.run(build_cmd)
        
        # Build 32-bit onedir
        print("Building Kiyabo32 (onedir)...")
        build_cmd = ["python32/python.exe", "-m", "PyInstaller", "--clean", "--noconfirm", "--noupx", "--onedir", "--icon=icon.ico", "--name=Kiyabo32", "--hidden-import=win32timezone", "run_cli.py"]
        subprocess.run(build_cmd)
        
        # Build 64-bit onefile
        print("Building Kiyabo64 (onefile)...")
        build_cmd = ["python64/python.exe", "-m", "PyInstaller", "--clean", "--noconfirm", "--noupx", "--onefile", "--icon=icon.ico", "--name=Kiyabo64", "--hidden-import=win32timezone", "run_cli.py"]
        # subprocess.run(build_cmd)
        
        # Build 64-bit onedir
        print("Building Kiyabo64 (onedir)...")
        build_cmd = ["python64/python.exe", "-m", "PyInstaller", "--clean", "--noconfirm", "--noupx", "--onedir", "--icon=icon.ico", "--name=Kiyabo64", "--hidden-import=win32timezone", "run_cli.py"]
        subprocess.run(build_cmd)
    else:
        # Original behavior - keep everything exactly the same
        # Clean previous builds (smarter cleaning)
        clean_previous_builds(args.onefile)
        
        # Common PyInstaller arguments to reduce AV detection
        common_args = [
            "--clean",
            "--noconfirm",
            "--noupx",           # Disable UPX (commonly flagged)
            "--hidden-import=win32timezone"
        ]
        
        # Build 32-bit
        print("Building Kiyabo32...")
        build_cmd = ["python32/python.exe", "-m", "PyInstaller"] + common_args
        
        if args.onefile:
            build_cmd.append("--onefile")
            print("  Mode: onefile")
        else:
            build_cmd.append("--onedir")
            print("  Mode: onedir")
        
        build_cmd.extend(["--icon=icon.ico", "--name=Kiyabo32", "run_cli.py"])
        subprocess.run(build_cmd)

        # Build 64-bit  
        print("Building Kiyabo64...")
        build_cmd = ["python64/python.exe", "-m", "PyInstaller"] + common_args
        
        if args.onefile:
            build_cmd.append("--onefile")
            print("  Mode: onefile")
        else:
            build_cmd.append("--onedir")
            print("  Mode: onedir")
        
        build_cmd.extend(["--icon=icon.ico", "--name=Kiyabo64", "run_cli.py"])
        subprocess.run(build_cmd)

    print("DONE")
    print(f"\nBuilt executables are in: {Path('dist').absolute()}")
    print("AV Reduction Techniques Applied:")
    print("✓ --noupx  (UPX disabled - commonly flagged by AV)")

if __name__ == "__main__":
    main()