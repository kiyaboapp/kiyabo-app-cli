#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Comprehensive Python Environment Testing Script
Tests all aspects of Python integration with VBA/MS Access
"""

import sys
import os
import platform
import subprocess
import json
from pathlib import Path

# ANSI color codes for better readability
class Colors:
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    MAGENTA = '\033[95m'
    RESET = '\033[0m'
    BOLD = '\033[1m'

def print_header(text):
    """Print a formatted header"""
    print(f"\n{Colors.CYAN}{Colors.BOLD}{'='*70}{Colors.RESET}")
    print(f"{Colors.CYAN}{Colors.BOLD}{text.center(70)}{Colors.RESET}")
    print(f"{Colors.CYAN}{Colors.BOLD}{'='*70}{Colors.RESET}\n")

def print_test(test_name, passed, message=""):
    """Print test result"""
    status = f"{Colors.GREEN}✓ PASSED{Colors.RESET}" if passed else f"{Colors.RED}✗ FAILED{Colors.RESET}"
    print(f"{Colors.BOLD}Test {test_name}:{Colors.RESET} {status}")
    if message:
        print(f"  {message}")
    print()

def test_python_version():
    """Test 1: Python version check"""
    try:
        version = sys.version_info
        version_str = f"{version.major}.{version.minor}.{version.micro}"
        
        # Check if version is acceptable (3.7+)
        if version.major >= 3 and version.minor >= 7:
            print_test("1: Python Version", True, 
                      f"Python {version_str} (Minimum 3.7 required)")
            return True
        else:
            print_test("1: Python Version", False, 
                      f"Python {version_str} is too old. Upgrade to 3.7+")
            return False
    except Exception as e:
        print_test("1: Python Version", False, f"Error: {e}")
        return False

def test_system_info():
    """Test 2: System information"""
    try:
        info = {
            "OS": platform.system(),
            "OS Version": platform.version(),
            "Machine": platform.machine(),
            "Processor": platform.processor(),
            "Python Implementation": platform.python_implementation()
        }
        
        message = "\n  ".join([f"{k}: {v}" for k, v in info.items()])
        print_test("2: System Information", True, message)
        return True
    except Exception as e:
        print_test("2: System Information", False, f"Error: {e}")
        return False

def test_environment_variables():
    """Test 3: Environment variables check"""
    try:
        python_path = sys.executable
        path_env = os.environ.get('PATH', '')
        pythonpath = os.environ.get('PYTHONPATH', 'Not set')
        
        # Check if Python is in PATH
        in_path = any(python_path.lower().startswith(p.lower()) 
                     for p in path_env.split(os.pathsep) if p)
        
        message = f"Python executable: {python_path}\n"
        message += f"  In PATH: {Colors.GREEN}Yes{Colors.RESET}" if in_path else f"  In PATH: {Colors.RED}No{Colors.RESET}"
        message += f"\n  PYTHONPATH: {pythonpath}"
        
        print_test("3: Environment Variables", in_path, message)
        return in_path
    except Exception as e:
        print_test("3: Environment Variables", False, f"Error: {e}")
        return False

def test_site_packages():
    """Test 4: Site packages paths"""
    try:
        import site
        
        user_site = site.getusersitepackages()
        site_packages = site.getsitepackages()
        
        message = f"User site-packages: {user_site}\n"
        message += f"  Exists: {os.path.exists(user_site)}\n"
        message += "  System site-packages:\n"
        for sp in site_packages:
            message += f"    - {sp} (exists: {os.path.exists(sp)})\n"
        
        print_test("4: Site Packages Paths", True, message)
        return True
    except Exception as e:
        print_test("4: Site Packages Paths", False, f"Error: {e}")
        return False

def test_package_import(package_name):
    """Helper: Test if a package can be imported"""
    try:
        module = __import__(package_name)
        version = getattr(module, '__version__', 'Unknown')
        location = getattr(module, '__file__', 'Unknown')
        return True, version, location
    except ImportError as e:
        return False, None, str(e)

def test_required_packages():
    """Test 5: Required packages (numpy, pandas, tqdm)"""
    packages = ['numpy', 'pandas', 'tqdm']
    results = {}
    all_passed = True
    
    message = ""
    for pkg in packages:
        success, version, location = test_package_import(pkg)
        results[pkg] = success
        
        if success:
            message += f"  {Colors.GREEN}✓{Colors.RESET} {pkg} {version}\n"
            message += f"    Location: {location}\n"
        else:
            message += f"  {Colors.RED}✗{Colors.RESET} {pkg} NOT FOUND\n"
            message += f"    Error: {location}\n"
            all_passed = False
    
    print_test("5: Required Packages", all_passed, message)
    return all_passed

def test_pip_functionality():
    """Test 6: pip functionality"""
    try:
        result = subprocess.run(
            [sys.executable, '-m', 'pip', '--version'],
            capture_output=True,
            text=True,
            timeout=10
        )
        
        if result.returncode == 0:
            message = result.stdout.strip()
            print_test("6: pip Functionality", True, message)
            return True
        else:
            print_test("6: pip Functionality", False, 
                      f"pip error: {result.stderr}")
            return False
    except subprocess.TimeoutExpired:
        print_test("6: pip Functionality", False, "pip command timed out")
        return False
    except Exception as e:
        print_test("6: pip Functionality", False, f"Error: {e}")
        return False

def test_file_operations():
    """Test 7: File read/write operations"""
    try:
        test_file = Path("test_file_operations.tmp")
        test_data = "Testing file operations from Python"
        
        # Write test
        test_file.write_text(test_data)
        
        # Read test
        read_data = test_file.read_text()
        
        # Cleanup
        test_file.unlink()
        
        passed = (test_data == read_data)
        print_test("7: File Operations", passed, 
                  "Can read and write files successfully" if passed 
                  else "File read/write mismatch")
        return passed
    except Exception as e:
        print_test("7: File Operations", False, f"Error: {e}")
        return False

def test_command_line_args():
    """Test 8: Command line arguments"""
    try:
        args = sys.argv[1:]  # Exclude script name
        
        if len(args) > 0:
            message = f"Received {len(args)} argument(s):\n"
            for i, arg in enumerate(args, 1):
                message += f"  Arg {i}: '{arg}'\n"
            print_test("8: Command Line Arguments", True, message)
            return True
        else:
            print_test("8: Command Line Arguments", True, 
                      "No arguments passed (this is OK for basic test)")
            return True
    except Exception as e:
        print_test("8: Command Line Arguments", False, f"Error: {e}")
        return False

def test_encoding():
    """Test 9: Unicode/encoding support"""
    try:
        # Test various encodings
        test_strings = {
            "ASCII": "Hello World",
            "Unicode": "Hello 世界 🌍",
            "Special": "Café résumé naïve"
        }
        
        all_passed = True
        message = ""
        
        for encoding_type, test_str in test_strings.items():
            try:
                # Try to encode and decode
                encoded = test_str.encode('utf-8')
                decoded = encoded.decode('utf-8')
                
                if decoded == test_str:
                    message += f"  {Colors.GREEN}✓{Colors.RESET} {encoding_type}: '{test_str}'\n"
                else:
                    message += f"  {Colors.RED}✗{Colors.RESET} {encoding_type}: Mismatch\n"
                    all_passed = False
            except Exception as e:
                message += f"  {Colors.RED}✗{Colors.RESET} {encoding_type}: {e}\n"
                all_passed = False
        
        print_test("9: Encoding/Unicode", all_passed, message)
        return all_passed
    except Exception as e:
        print_test("9: Encoding/Unicode", False, f"Error: {e}")
        return False

def test_console_output():
    """Test 10: Console output with colors"""
    try:
        print_test("10: Console Output", True, 
                  f"{Colors.GREEN}Colors{Colors.RESET} {Colors.YELLOW}are{Colors.RESET} {Colors.BLUE}working{Colors.RESET}!")
        return True
    except Exception as e:
        print_test("10: Console Output", False, f"Error: {e}")
        return False

def test_comprehensive_imports():
    """Test 11: Comprehensive package functionality"""
    try:
        import numpy as np
        import pandas as pd
        from tqdm import tqdm
        import time
        
        message = ""
        all_passed = True
        
        # Test numpy
        try:
            arr = np.array([1, 2, 3, 4, 5])
            mean_val = np.mean(arr)
            message += f"  {Colors.GREEN}✓{Colors.RESET} NumPy: Array mean = {mean_val}\n"
        except Exception as e:
            message += f"  {Colors.RED}✗{Colors.RESET} NumPy failed: {e}\n"
            all_passed = False
        
        # Test pandas
        try:
            df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
            sum_val = df['A'].sum()
            message += f"  {Colors.GREEN}✓{Colors.RESET} Pandas: DataFrame sum = {sum_val}\n"
        except Exception as e:
            message += f"  {Colors.RED}✗{Colors.RESET} Pandas failed: {e}\n"
            all_passed = False
        
        # Test tqdm
        try:
            items = list(tqdm(range(5), desc="Testing tqdm", ncols=50, leave=False))
            message += f"  {Colors.GREEN}✓{Colors.RESET} tqdm: Progress bar working\n"
        except Exception as e:
            message += f"  {Colors.RED}✗{Colors.RESET} tqdm failed: {e}\n"
            all_passed = False
        
        print_test("11: Package Functionality", all_passed, message)
        return all_passed
    except ImportError as e:
        print_test("11: Package Functionality", False, 
                  f"Cannot import required packages: {e}")
        return False
    except Exception as e:
        print_test("11: Package Functionality", False, f"Error: {e}")
        return False

def test_memory_and_performance():
    """Test 12: Memory usage and basic performance"""
    try:
        import psutil
        process = psutil.Process()
        memory_mb = process.memory_info().rss / 1024 / 1024
        
        message = f"Memory usage: {memory_mb:.2f} MB\n"
        
        # Simple performance test
        import time
        start = time.time()
        result = sum(range(1000000))
        elapsed = time.time() - start
        
        message += f"  Performance test: {elapsed*1000:.2f}ms for 1M iterations"
        
        print_test("12: Memory & Performance", True, message)
        return True
    except ImportError:
        print_test("12: Memory & Performance", True, 
                  "psutil not installed (optional test)")
        return True
    except Exception as e:
        print_test("12: Memory & Performance", False, f"Error: {e}")
        return False

def test_json_operations():
    """Test 13: JSON serialization (important for data exchange)"""
    try:
        test_data = {
            "name": "Test User",
            "values": [1, 2, 3, 4, 5],
            "nested": {
                "key1": "value1",
                "key2": 123
            }
        }
        
        # Serialize
        json_str = json.dumps(test_data)
        
        # Deserialize
        loaded_data = json.loads(json_str)
        
        passed = (test_data == loaded_data)
        print_test("13: JSON Operations", passed, 
                  "JSON serialization/deserialization working" if passed 
                  else "JSON data mismatch")
        return passed
    except Exception as e:
        print_test("13: JSON Operations", False, f"Error: {e}")
        return False

def test_subprocess_execution():
    """Test 14: Subprocess execution (for running other commands)"""
    try:
        result = subprocess.run(
            [sys.executable, '-c', 'print("subprocess_test_ok")'],
            capture_output=True,
            text=True,
            timeout=5
        )
        
        passed = (result.returncode == 0 and 
                 "subprocess_test_ok" in result.stdout)
        
        print_test("14: Subprocess Execution", passed, 
                  "Can execute Python subprocesses" if passed 
                  else "Subprocess execution failed")
        return passed
    except Exception as e:
        print_test("14: Subprocess Execution", False, f"Error: {e}")
        return False

def test_exception_handling():
    """Test 15: Exception handling"""
    try:
        # Intentional error
        try:
            result = 1 / 0
        except ZeroDivisionError as e:
            message = f"Caught expected exception: {type(e).__name__}"
            print_test("15: Exception Handling", True, message)
            return True
        
        # Should not reach here
        print_test("15: Exception Handling", False, 
                  "Exception was not caught")
        return False
    except Exception as e:
        print_test("15: Exception Handling", False, f"Unexpected error: {e}")
        return False

def generate_diagnostic_report(results):
    """Generate a comprehensive diagnostic report"""
    print_header("DIAGNOSTIC REPORT")
    
    total_tests = len(results)
    passed_tests = sum(results.values())
    failed_tests = total_tests - passed_tests
    
    print(f"{Colors.BOLD}Total Tests:{Colors.RESET} {total_tests}")
    print(f"{Colors.GREEN}{Colors.BOLD}Passed:{Colors.RESET} {passed_tests}")
    print(f"{Colors.RED}{Colors.BOLD}Failed:{Colors.RESET} {failed_tests}")
    print(f"{Colors.BLUE}{Colors.BOLD}Success Rate:{Colors.RESET} {(passed_tests/total_tests)*100:.1f}%\n")
    
    if failed_tests == 0:
        print(f"{Colors.GREEN}{Colors.BOLD}✓ ALL TESTS PASSED!{Colors.RESET}")
        print(f"{Colors.GREEN}Your Python environment is fully functional.{Colors.RESET}\n")
    else:
        print(f"{Colors.YELLOW}{Colors.BOLD}⚠ SOME TESTS FAILED{Colors.RESET}")
        print(f"{Colors.YELLOW}Review the failed tests above for details.{Colors.RESET}\n")

def print_troubleshooting_guide():
    """Print common troubleshooting steps"""
    print_header("COMMON ISSUES & SOLUTIONS")
    
    issues = {
        "Python not in PATH": [
            "Reinstall Python with 'Add Python to PATH' checked",
            "Manually add Python to PATH in System Environment Variables",
            "Restart your computer after installation"
        ],
        "Packages not found": [
            "Run: pip install --user numpy pandas tqdm",
            "Check if pip is working: python -m pip --version",
            "Try: python -m pip install --upgrade pip"
        ],
        "Permission errors": [
            "Use --user flag when installing: pip install --user package_name",
            "Run Command Prompt as Administrator",
            "Check antivirus isn't blocking Python"
        ],
        "Import errors after install": [
            "Close and reopen MS Access completely",
            "Restart your computer",
            "Check Python version matches (32-bit vs 64-bit)"
        ],
        "VBA execution fails": [
            "Verify Python path in VBA matches actual location",
            "Check working directory has write permissions",
            "Disable antivirus temporarily to test",
            "Check Windows Firewall isn't blocking"
        ]
    }
    
    for issue, solutions in issues.items():
        print(f"{Colors.CYAN}{Colors.BOLD}{issue}:{Colors.RESET}")
        for i, solution in enumerate(solutions, 1):
            print(f"  {i}. {solution}")
        print()

def print_system_diagnostic():
    """Print detailed system diagnostic information"""
    print_header("SYSTEM DIAGNOSTIC INFORMATION")
    
    print(f"{Colors.BOLD}Python Executable:{Colors.RESET}")
    print(f"  {sys.executable}\n")
    
    print(f"{Colors.BOLD}Python Path:{Colors.RESET}")
    for path in sys.path:
        print(f"  {path}")
    print()
    
    print(f"{Colors.BOLD}Environment Variables:{Colors.RESET}")
    important_vars = ['PATH', 'PYTHONPATH', 'PYTHONHOME', 'USERPROFILE', 'TEMP']
    for var in important_vars:
        value = os.environ.get(var, 'Not set')
        if var == 'PATH':
            print(f"  {var}:")
            for p in value.split(os.pathsep):
                if p:
                    print(f"    {p}")
        else:
            print(f"  {var}: {value}")
    print()

def main():
    """Main test execution"""
    print_header("COMPREHENSIVE PYTHON ENVIRONMENT TEST")
    print(f"{Colors.MAGENTA}Testing Python integration for VBA/MS Access{Colors.RESET}")
    print(f"{Colors.MAGENTA}Script started at: {__file__}{Colors.RESET}\n")
    
    # Track all test results
    results = {}
    
    # Run all tests
    print_header("RUNNING TESTS")
    
    results['python_version'] = test_python_version()
    results['system_info'] = test_system_info()
    results['environment'] = test_environment_variables()
    results['site_packages'] = test_site_packages()
    results['required_packages'] = test_required_packages()
    results['pip'] = test_pip_functionality()
    results['file_operations'] = test_file_operations()
    results['command_args'] = test_command_line_args()
    results['encoding'] = test_encoding()
    results['console_output'] = test_console_output()
    results['package_functionality'] = test_comprehensive_imports()
    results['memory_performance'] = test_memory_and_performance()
    results['json_operations'] = test_json_operations()
    results['subprocess'] = test_subprocess_execution()
    results['exception_handling'] = test_exception_handling()
    
    # Generate report
    generate_diagnostic_report(results)
    
    # Print system diagnostic
    print_system_diagnostic()
    
    # Print troubleshooting guide if any tests failed
    if not all(results.values()):
        print_troubleshooting_guide()
    
    # Final message
    print_header("TEST COMPLETED")
    
    if all(results.values()):
        print(f"{Colors.GREEN}{Colors.BOLD}SUCCESS!{Colors.RESET}")
        print(f"{Colors.GREEN}All tests passed. Python is ready for VBA integration.{Colors.RESET}\n")
        sys.exit(0)
    else:
        print(f"{Colors.YELLOW}{Colors.BOLD}WARNING!{Colors.RESET}")
        print(f"{Colors.YELLOW}Some tests failed. Review the report above.{Colors.RESET}\n")
        sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n{Colors.YELLOW}Test interrupted by user.{Colors.RESET}")
        sys.exit(1)
    except Exception as e:
        print(f"\n{Colors.RED}{Colors.BOLD}CRITICAL ERROR:{Colors.RESET}")
        print(f"{Colors.RED}{e}{Colors.RESET}")
        import traceback
        traceback.print_exc()
        sys.exit(1)