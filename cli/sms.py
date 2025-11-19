import os
import time
import random
import subprocess
import threading
import sys
import argparse
from typing import Optional, Tuple, List, Dict
import openpyxl
from openpyxl.styles import PatternFill
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeRemainingColumn
from rich.prompt import Prompt, Confirm, IntPrompt
from rich.table import Table
from rich.text import Text
import tkinter as tk
from tkinter import filedialog, messagebox
from tabulate import tabulate
from datetime import datetime

class SMSSender:
    def __init__(self, excel_path: str = "", delay_min: int = 5, delay_max: int = 10, 
                 show_window: bool = True, test_mode: bool = False):
        self.console = Console()
        self.should_cancel = False
        self.mpe_path = r"C:\Program Files (x86)\MyPhoneExplorer\MyPhoneExplorer.exe"
        self.excel_path = excel_path
        self.delay_min = delay_min
        self.delay_max = delay_max
        self.show_window = show_window
        self.test_mode = test_mode
        
        # Timing statistics
        self.timing_stats = {
            'min_time': float('inf'),
            'max_time': 0,
            'total_time': 0,
            'count': 0,
            'start_time': None,
            'end_time': None
        }
        
        self.stats = {
            'total': 0,
            'success': 0,
            'failed': 0,
            'skipped': 0,
            'cancelled': 0,
            'pending': 0,
            'already_sent': 0,
            'details': []
        }
        
        # Initialize Tkinter root (hidden)
        self.root = tk.Tk()
        self.root.withdraw()  # Hide the main window
        
    def clear_screen(self):
        """Clear the console screen"""
        os.system('cls' if os.name == 'nt' else 'clear')
        
    def show_header(self):
        """Display elegant header"""
        self.clear_screen()
        self.console.print("\n")
        self.console.print("🎯 SMS SENDING SYSTEM", style="bold cyan")
        self.console.print("MyPhoneExplorer Automation", style="dim")
        
        # Show current configuration
        config_table = [
            ["Excel File", self.excel_path or "Not loaded"],
            ["Delay Range", f"{self.delay_min}-{self.delay_max} seconds"],
            ["Show Window", "Yes" if self.show_window else "No"],
            ["Test Mode", "Yes" if self.test_mode else "No"]
        ]
        self.console.print("\n⚙️  Current Configuration:")
        config_display = tabulate(config_table, tablefmt="simple")
        self.console.print(config_display)
        self.console.print("\n")
        
    def show_menu(self):
        """Display main menu"""
        self.console.print("1. ", style="bold cyan", end="")
        self.console.print("Load Excel File & Preview", style="white")
        
        self.console.print("2. ", style="bold cyan", end="")
        self.console.print("Check MyPhoneExplorer Status", style="white")
        
        self.console.print("3. ", style="bold cyan", end="")
        self.console.print("Connect MyPhoneExplorer", style="white")
        
        self.console.print("4. ", style="bold cyan", end="")
        self.console.print("Send SMS Messages", style="white")
        
        self.console.print("5. ", style="bold cyan", end="")
        self.console.print("Test SMS Sending", style="white")
        
        self.console.print("6. ", style="bold cyan", end="")
        self.console.print("Adjust Settings", style="white")
        
        self.console.print("7. ", style="bold cyan", end="")
        self.console.print("View Statistics", style="white")
        
        self.console.print("0. ", style="bold cyan", end="")
        self.console.print("Exit", style="white")
        
        self.console.print("\nSelect an option: ", style="bold yellow", end="")
    
    def adjust_settings(self):
        """Allow user to adjust settings"""
        self.console.print("\n⚙️  Adjust Settings", style="bold")
        self.console.print("=" * 50, style="dim")
        
        # Adjust delay settings
        self.console.print(f"\nCurrent delay: {self.delay_min}-{self.delay_max} seconds")
        new_min = IntPrompt.ask("Enter minimum delay in seconds", default=self.delay_min)
        new_max = IntPrompt.ask("Enter maximum delay in seconds", default=self.delay_max)
        
        if new_min > new_max:
            self.console.print("❌ Minimum cannot be greater than maximum", style="red")
        else:
            self.delay_min = new_min
            self.delay_max = new_max
            self.console.print(f"✅ Delay set to {self.delay_min}-{self.delay_max} seconds", style="green")
        
        # Adjust window visibility
        current_window_setting = "shown" if self.show_window else "hidden"
        self.console.print(f"\nCurrent window setting: {current_window_setting}")
        self.show_window = Confirm.ask("Show MyPhoneExplorer window during sending?")
        
        window_status = "shown" if self.show_window else "hidden"
        self.console.print(f"✅ Window will be {window_status}", style="green")
        
        # Show updated configuration
        self.console.print("\n📋 Updated Configuration:", style="bold")
        config_table = [
            ["Delay Range", f"{self.delay_min}-{self.delay_max} seconds"],
            ["Show Window", "Yes" if self.show_window else "No"]
        ]
        config_display = tabulate(config_table, tablefmt="grid")
        self.console.print(config_display)
    
    def test_sms_sending(self):
        """Test SMS sending with user-provided number and message"""
        self.console.print("\n🧪 TEST SMS SENDING", style="bold")
        self.console.print("=" * 50, style="dim")
        
        if not self.is_sms_direct_installed():
            self.console.print("❌ MyPhoneExplorer is not installed", style="red")
            return
        
        # Get phone number from user
        phone_number = Prompt.ask("📞 Enter phone number to test")
        normalized_phone = self.normalize_phone(phone_number)
        self.console.print(f"📞 Normalized phone: {normalized_phone}", style="blue")
        
        # Get message from user
        message = Prompt.ask("💬 Enter test message")
        
        # Confirm before sending
        self.console.print(f"\n📋 Test SMS Details:", style="bold")
        self.console.print(f"   To: {normalized_phone}")
        self.console.print(f"   Message: {message}")
        
        if not Confirm.ask("\nSend this test SMS?"):
            self.console.print("❌ Test cancelled", style="yellow")
            return
        
        # Send test SMS
        self.console.print("\n🚀 Sending test SMS...", style="yellow")
        
        success, result_message = self.try_send_text_sms(normalized_phone, message)
        
        if success:
            self.console.print(f"✅ Test SMS sent successfully!", style="green")
            self.console.print(f"📝 Result: {result_message}", style="dim")
        else:
            self.console.print(f"❌ Test SMS failed!", style="red")
            self.console.print(f"📝 Error: {result_message}", style="dim")
    
    def browse_excel_file(self) -> str:
        """Open file dialog to browse for Excel file"""
        try:
            file_path = filedialog.askopenfilename(
                title="Select Excel File",
                filetypes=[
                    ("Excel files", "*.xlsx *.xls"),
                    ("All files", "*.*")
                ]
            )
            return file_path
        except Exception as e:
            self.console.print(f"❌ File dialog error: {str(e)}", style="red")
            return ""
    
    def is_sms_direct_installed(self) -> bool:
        """Check if MyPhoneExplorer is installed"""
        return os.path.exists(self.mpe_path)
    
    def is_app_running(self, app_name: str = "MyPhoneExplorer.exe") -> bool:
        """Check if MyPhoneExplorer is running"""
        try:
            result = subprocess.run(
                ['tasklist', '/FI', f'IMAGENAME eq {app_name}'],
                capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW
            )
            return app_name.lower() in result.stdout.lower()
        except:
            return False
    
    def connect_myphone_explorer(self) -> bool:
        """Initialize connection to MyPhoneExplorer"""
        if not self.is_sms_direct_installed():
            self.console.print("\n❌ MyPhoneExplorer is not installed.", style="red")
            self.console.print("Please install MyPhoneExplorer first.", style="dim")
            return False
        
        try:
            self.console.print("\n🔌 Initializing connection to MyPhoneExplorer...", style="yellow")
            
            # Use window setting
            creation_flags = 0 if self.show_window else subprocess.CREATE_NO_WINDOW
            
            args = "action=connect"
            command = f'"{self.mpe_path}" {args}'
            
            process = subprocess.run(
                command,
                shell=True,
                capture_output=True,
                text=True,
                timeout=30,
                creationflags=creation_flags
            )
            
            if process.returncode == 0:
                self.console.print("✅ MyPhoneExplorer connected successfully", style="green")
                return True
            else:
                self.console.print("⚠️ MyPhoneExplorer opened but connection may need manual setup", style="yellow")
                return True
                
        except subprocess.TimeoutExpired:
            self.console.print("✅ MyPhoneExplorer started (timeout reached)", style="green")
            return True
        except Exception as e:
            self.console.print(f"❌ Failed to connect: {str(e)}", style="red")
            return False
    
    def sanitize_sms(self, message: str) -> str:
        """Sanitize SMS message - replace newlines with %n"""
        if not message:
            return ""
        return message.replace('\n', '%n').replace('\r', '%n')
    
    def normalize_phone(self, phone: str) -> str:
        """Normalize phone number to Tanzanian format"""
        if not phone:
            return ""
            
        phone = str(phone).strip().replace(" ", "")
        
        # Remove any non-digit characters except +
        cleaned_phone = ''.join(c for c in phone if c.isdigit() or c == '+')
        
        if cleaned_phone.startswith('255') and len(cleaned_phone) == 12:
            return "+" + cleaned_phone
        elif len(cleaned_phone) == 9:
            return "+255" + cleaned_phone
        elif len(cleaned_phone) == 10 and cleaned_phone.startswith('0'):
            return "+255" + cleaned_phone[1:]
        elif cleaned_phone.startswith('+'):
            return cleaned_phone
        else:
            return "+255" + cleaned_phone[-9:]  # Take last 9 digits
    
    def load_excel_preview(self) -> bool:
        """Load Excel file and show preview with FULL messages"""
        if not self.excel_path:
            self.console.print("\n📁 Select Excel file...", style="bold")
            self.console.print("1. ", style="cyan", end="")
            self.console.print("Browse with file dialog", style="white")
            self.console.print("2. ", style="cyan", end="")
            self.console.print("Enter path manually", style="white")
            self.console.print("Choose option: ", style="yellow", end="")
            
            choice = input().strip()
            
            if choice == '1':
                # Use Tkinter file dialog
                self.console.print("🗔 Opening file dialog...", style="dim")
                file_path = self.browse_excel_file()
                if file_path:
                    self.excel_path = file_path
                    self.console.print(f"✅ Selected: {self.excel_path}", style="green")
                else:
                    self.console.print("❌ No file selected", style="red")
                    return False
            else:
                # Manual path entry
                path_input = Prompt.ask("📁 Enter path to Excel file")
                # Remove surrounding quotes if user included them
                self.excel_path = path_input.strip('"\'')
        
        # Normalize path
        self.excel_path = os.path.abspath(self.excel_path)
        
        if not os.path.exists(self.excel_path):
            self.console.print(f"❌ File not found: {self.excel_path}", style="red")
            self.excel_path = ""  # Reset path
            return False
        
        try:
            # Save a backup immediately when opening
            backup_path = self.excel_path.replace('.xlsx', f'_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
            import shutil
            shutil.copy2(self.excel_path, backup_path)
            self.console.print(f"💾 Backup created: {os.path.basename(backup_path)}", style="dim")
            
            workbook = openpyxl.load_workbook(self.excel_path)
            sheet = workbook.active
            
            # Find data range
            last_row = 1
            for row in range(2, sheet.max_row + 1):
                if sheet.cell(row=row, column=4).value:
                    last_row = row
                else:
                    break
            
            if last_row <= 1:
                self.console.print("❌ No data found in Excel file", style="red")
                workbook.close()
                return False
            
            total_messages = last_row - 1
            
            # Show preview
            self.console.print(f"\n📊 Excel File Loaded: {os.path.basename(self.excel_path)}", style="bold green")
            self.console.print(f"📋 Total messages to send: {total_messages}", style="cyan")
            
            # Preview first 3 messages with FULL messages
            self.console.print("\n👥 Preview of first 3 messages:", style="bold")
            self.console.print("=" * 80, style="dim")
            
            for i in range(2, min(5, last_row + 1)):
                student = str(sheet.cell(row=i, column=2).value or "").strip()
                phone = str(sheet.cell(row=i, column=4).value or "")
                message = str(sheet.cell(row=i, column=3).value or "")
                status = str(sheet.cell(row=i, column=5).value or "").strip()
                
                normalized_phone = self.normalize_phone(phone)
                
                # Show FULL message without any truncation
                self.console.print(f"\n#{i-1}", style="bold yellow")
                self.console.print(f"🎓 Student: {student}", style="white")
                self.console.print(f"📞 Phone: {phone} → {normalized_phone}", style="blue")
                self.console.print(f"📝 Status: {status if status else 'PENDING'}", style="cyan")
                self.console.print(f"💬 Message:", style="bold white")
                self.console.print(f"{message}", style="dim")
                
                if i < min(5, last_row + 1) - 1:
                    self.console.print("-" * 80, style="dim")
            
            # Show status summary in table format
            status_counts = {}
            for i in range(2, last_row + 1):
                status = str(sheet.cell(row=i, column=5).value or "").strip().upper()
                if status:
                    status_counts[status] = status_counts.get(status, 0) + 1
                else:
                    status_counts['PENDING'] = status_counts.get('PENDING', 0) + 1
            
            if status_counts:
                status_table = []
                for status, count in status_counts.items():
                    status_table.append([status, count])
                
                self.console.print("\n📈 Current Status Summary:", style="bold")
                headers = ["Status", "Count"]
                table = tabulate(status_table, headers=headers, tablefmt="grid")
                self.console.print(table)
            
            workbook.close()
            return True
            
        except Exception as e:
            self.console.print(f"❌ Error loading Excel file: {str(e)}", style="red")
            self.excel_path = ""  # Reset path on error
            return False
    
    def try_send_text_sms(self, phone_number: str, text_sms: str) -> Tuple[bool, str]:
        """Send SMS using MyPhoneExplorer and wait for completion"""
        if not self.is_sms_direct_installed():
            return False, "MyPhoneExplorer not installed"
        
        try:
            sanitized_message = self.sanitize_sms(text_sms)
            
            # Build command with proper formatting
            args = f'action=sendmessage savetosent=1 number={phone_number} Text="{sanitized_message}"'
            command = f'"{self.mpe_path}" {args}'
            
            # Use window setting
            creation_flags = 0 if self.show_window else subprocess.CREATE_NO_WINDOW
            
            # Time the execution
            start_time = time.time()
            
            process = subprocess.run(
                command,
                shell=True,
                capture_output=True,
                text=True,
                timeout=120,
                creationflags=creation_flags
            )
            
            execution_time = time.time() - start_time
            
            # Update timing statistics
            self.timing_stats['min_time'] = min(self.timing_stats['min_time'], execution_time)
            self.timing_stats['max_time'] = max(self.timing_stats['max_time'], execution_time)
            self.timing_stats['total_time'] += execution_time
            self.timing_stats['count'] += 1
            
            # Wait random delay after sending
            delay = random.randint(self.delay_min, self.delay_max)
            time.sleep(delay)
            
            if process.returncode == 0:
                return True, f"Sent in {execution_time:.2f}s (waited {delay}s)"
            else:
                error_msg = process.stderr.strip() if process.stderr else "Unknown error"
                return False, f"Failed in {execution_time:.2f}s: {error_msg} (waited {delay}s)"
                
        except subprocess.TimeoutExpired:
            return False, "Timeout - SMS sending took too long"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def send_sms_messages(self):
        """Main function to send SMS messages"""
        if not self.excel_path:
            self.console.print("❌ Please load an Excel file first", style="red")
            return
        
        if not self.is_sms_direct_installed():
            self.console.print("❌ MyPhoneExplorer is not installed", style="red")
            return
        
        # Show current configuration
        self.console.print("\n⚙️  Sending Configuration:", style="bold")
        config_table = [
            ["Delay Range", f"{self.delay_min}-{self.delay_max} seconds"],
            ["Show Window", "Yes" if self.show_window else "No"],
            ["Test Mode", "Yes" if self.test_mode else "No"]
        ]
        config_display = tabulate(config_table, tablefmt="grid")
        self.console.print(config_display)
        
        # Check if MyPhoneExplorer is running
        if not self.is_app_running():
            self.console.print("⚠️ MyPhoneExplorer is not running", style="yellow")
            if Confirm.ask("Would you like to start and connect it now?"):
                if not self.connect_myphone_explorer():
                    self.console.print("❌ Cannot proceed without MyPhoneExplorer", style="red")
                    return
            else:
                self.console.print("❌ Cannot send SMS without MyPhoneExplorer", style="red")
                return
        
        try:
            workbook = openpyxl.load_workbook(self.excel_path)
            sheet = workbook.active
            
            # Find last row
            last_row = 1
            for row in range(2, sheet.max_row + 1):
                if sheet.cell(row=row, column=4).value:
                    last_row = row
                else:
                    break
            
            if last_row <= 1:
                self.console.print("❌ No data found in Excel file", style="red")
                workbook.close()
                return
            
            total_rows = last_row - 1
            
            # Check for already sent messages
            already_sent_count = 0
            for i in range(2, last_row + 1):
                status = str(sheet.cell(row=i, column=5).value or "").strip().upper()
                if status == "OK" or status == "SENT":
                    already_sent_count += 1
            
            resend_all = False
            if already_sent_count > 0:
                self.console.print(f"\n📊 Found {already_sent_count} messages already marked as SENT/OK", style="yellow")
                resend_all = Confirm.ask("Do you want to resend ALL messages (including already sent)?")
                if not resend_all:
                    self.console.print("✅ Will skip already sent messages", style="green")
            
            # Reset timing stats for this session
            self.timing_stats = {
                'min_time': float('inf'),
                'max_time': 0,
                'total_time': 0,
                'count': 0,
                'start_time': datetime.now(),
                'end_time': None
            }
            
            self.stats = {
                'total': total_rows,
                'success': 0,
                'failed': 0,
                'skipped': 0,
                'cancelled': 0,
                'pending': 0,
                'already_sent': already_sent_count,
                'details': []
            }
            
            self.console.print(f"\n🚀 Starting to send {total_rows} SMS messages...", style="bold green")
            self.console.print("⏸️  Press 'q' + Enter to cancel at any time", style="yellow")
            
            # Start cancellation listener
            self.should_cancel = False
            cancel_thread = threading.Thread(target=self._listen_for_cancel)
            cancel_thread.daemon = True
            cancel_thread.start()
            
            # Process with progress bar
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
                TimeRemainingColumn(),
            ) as progress:
                
                task = progress.add_task("[cyan]Sending...", total=total_rows)
                processed_count = 0
                
                for i in range(2, last_row + 1):
                    if self.should_cancel:
                        remaining = total_rows - processed_count
                        self.stats['cancelled'] = remaining
                        break
                    
                    student_name = str(sheet.cell(row=i, column=2).value or "").strip()
                    phone_raw = str(sheet.cell(row=i, column=4).value or "")
                    message_raw = str(sheet.cell(row=i, column=3).value or "")
                    current_status = str(sheet.cell(row=i, column=5).value or "").strip().upper()
                    
                    clean_phone = self.normalize_phone(phone_raw)
                    text_sms = self.sanitize_sms(message_raw).strip()
                    
                    status_cell = sheet.cell(row=i, column=5)
                    detail_cell = sheet.cell(row=i, column=6)
                    
                    # Skip already sent messages unless resend_all is True
                    if not resend_all and current_status in ["OK", "SENT"]:
                        self.stats['skipped'] += 1
                        progress.console.print(f"⏭️ {student_name}: Already sent - skipping", style="dim")
                        processed_count += 1
                        progress.update(task, advance=1)
                        continue
                    
                    if clean_phone and text_sms:
                        progress.console.print(f"📤 Sending to {student_name}...", style="yellow")
                        
                        send_result, status_message = self.try_send_text_sms(clean_phone, text_sms)
                        
                        if send_result:
                            status_cell.value = "OK"
                            status_cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                            self.stats['success'] += 1
                            progress.console.print(f"✅ {student_name}: {status_message}", style="green")
                        else:
                            status_cell.value = "PENDING"
                            status_cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                            self.stats['failed'] += 1
                            progress.console.print(f"❌ {student_name}: {status_message}", style="red")
                        
                        # Store details for statistics
                        self.stats['details'].append({
                            'student': student_name,
                            'phone': clean_phone,
                            'status': 'OK' if send_result else 'PENDING',
                            'message': status_message
                        })
                    else:
                        status_cell.value = "INVALID"
                        status_cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                        detail_cell.value = "Invalid phone or message"
                        self.stats['skipped'] += 1
                        progress.console.print(f"⚠️ {student_name}: Skipped - invalid data", style="yellow")
                    
                    processed_count += 1
                    progress.update(task, advance=1)
                
                # Update end time
                self.timing_stats['end_time'] = datetime.now()
                
                # Save workbook
                if not self.should_cancel:
                    try:
                        workbook.save(self.excel_path)
                        self.console.print("💾 Excel file saved with status updates", style="green")
                    except Exception as e:
                        self.console.print(f"❌ Error saving Excel file: {str(e)}", style="red")
                        # Try to save with a different name
                        try:
                            new_path = self.excel_path.replace('.xlsx', f'_updated_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
                            workbook.save(new_path)
                            self.console.print(f"💾 Saved as backup: {os.path.basename(new_path)}", style="yellow")
                        except:
                            self.console.print("❌ Could not save changes", style="red")
                
            workbook.close()
            
            # Show final statistics
            self.show_statistics()
            
        except Exception as e:
            self.console.print(f"❌ Error: {str(e)}", style="red")
    
    def show_statistics(self):
        """Display detailed statistics in tabular format"""
        stats = self.stats
        timing = self.timing_stats
        
        self.console.print("\n" + "="*60, style="bold cyan")
        self.console.print("📊 SMS SENDING STATISTICS", style="bold cyan")
        self.console.print("="*60, style="bold cyan")
        
        # Main statistics table
        main_stats = [
            ["Total Messages", stats['total']],
            ["✅ Successfully Sent", stats['success']],
            ["❌ Failed", stats['failed']],
            ["⚠️ Skipped (Invalid)", stats['skipped']],
            ["⏭️ Already Sent", stats['already_sent']],
        ]
        
        if stats['cancelled'] > 0:
            main_stats.append(["🚫 Cancelled", stats['cancelled']])
        
        # Add percentages
        percentage_stats = []
        for row in main_stats:
            if row[0] != "Total Messages" and not row[0].startswith("🚫"):
                percentage = (row[1] / stats['total']) * 100
                percentage_stats.append([row[0], row[1], f"{percentage:.1f}%"])
            else:
                percentage_stats.append([row[0], row[1], "-"])
        
        headers = ["Category", "Count", "Percentage"]
        main_table = tabulate(percentage_stats, headers=headers, tablefmt="grid")
        self.console.print(main_table)
        
        # Timing statistics
        if timing['count'] > 0:
            avg_time = timing['total_time'] / timing['count']
            timing_stats = [
                ["Total Execution Time", f"{timing['total_time']:.2f}s"],
                ["Number of SMS Sent", timing['count']],
                ["Fastest SMS", f"{timing['min_time']:.2f}s"],
                ["Slowest SMS", f"{timing['max_time']:.2f}s"],
                ["Average Time per SMS", f"{avg_time:.2f}s"]
            ]
            
            if timing['start_time'] and timing['end_time']:
                duration = timing['end_time'] - timing['start_time']
                timing_stats.append(["Total Session Duration", str(duration).split('.')[0]])
            
            self.console.print("\n⏱️  Timing Statistics:", style="bold")
            timing_table = tabulate(timing_stats, headers=["Metric", "Value"], tablefmt="grid")
            self.console.print(timing_table)
        
        # Recent activity table
        if stats['details']:
            recent_data = []
            for detail in stats['details'][:8]:
                status_icon = "✅" if detail['status'] == 'OK' else "❌"
                recent_data.append([
                    status_icon,
                    detail['student'],
                    detail['phone'],
                    detail['message']
                ])
            
            self.console.print(f"\n📋 Recent Activity ({len(stats['details'])} messages processed):", style="bold")
            recent_headers = ["Status", "Student", "Phone", "Result"]
            recent_table = tabulate(recent_data, headers=recent_headers, tablefmt="grid")
            self.console.print(recent_table)
    
    def show_app_status(self):
        """Show MyPhoneExplorer installation and connection status in tabular format"""
        self.console.print("\n🔍 MyPhoneExplorer Status Check", style="bold")
        self.console.print("─" * 40, style="dim")
        
        status_data = []
        
        # Installation status
        installed = self.is_sms_direct_installed()
        status_data.append(["MyPhoneExplorer Installed", "✅ Yes" if installed else "❌ No"])
        
        if installed:
            status_data.append(["Installation Path", self.mpe_path])
        
        # Running status
        running = self.is_app_running()
        status_data.append(["MyPhoneExplorer Running", "✅ Yes" if running else "❌ No"])
        
        # Excel status
        excel_loaded = self.excel_path and os.path.exists(self.excel_path)
        status_data.append(["Excel File Loaded", "✅ Yes" if excel_loaded else "❌ No"])
        
        if excel_loaded:
            try:
                workbook = openpyxl.load_workbook(self.excel_path)
                sheet = workbook.active
                last_row = 1
                for row in range(2, sheet.max_row + 1):
                    if sheet.cell(row=row, column=4).value:
                        last_row = row
                total_messages = last_row - 1
                workbook.close()
                status_data.append(["Messages Ready", total_messages])
            except:
                status_data.append(["Excel Status", "⚠️ File has issues"])
        
        status_table = tabulate(status_data, headers=["Component", "Status"], tablefmt="grid")
        self.console.print(status_table)
    
    def _listen_for_cancel(self):
        """Listen for cancellation input"""
        while not self.should_cancel:
            try:
                user_input = input().strip().lower()
                if user_input == 'q':
                    self.should_cancel = True
                    self.console.print("\n🛑 Cancellation requested...", style="yellow")
                    break
            except:
                break
    
    def run(self):
        """Main application loop"""
        while True:
            self.show_header()
            self.show_menu()
            
            try:
                choice = input().strip()
                
                if choice == '1':
                    self.load_excel_preview()
                elif choice == '2':
                    self.show_app_status()
                elif choice == '3':
                    self.connect_myphone_explorer()
                elif choice == '4':
                    self.send_sms_messages()
                elif choice == '5':
                    self.test_sms_sending()
                elif choice == '6':
                    self.adjust_settings()
                elif choice == '7':
                    if self.stats['total'] > 0:
                        self.show_statistics()
                    else:
                        self.console.print("📊 No statistics available yet", style="yellow")
                elif choice == '0':
                    self.console.print("\n👋 Thank you for using SMS Sending System!", style="green")
                    sys.exit(0)
                else:
                    self.console.print("❌ Invalid choice. Please try again.", style="red")
                    continue
                
                # Wait for user to continue
                if choice != '0':
                    self.console.print("\n↵ Press Enter to continue...", style="dim", end="")
                    input()
                    
            except KeyboardInterrupt:
                self.console.print("\n\n👋 Application interrupted. Goodbye!", style="yellow")
                sys.exit(0)
            except Exception as e:
                self.console.print(f"\n❌ Error: {str(e)}", style="red")
                self.console.print("↵ Press Enter to continue...", style="dim", end="")
                input()

def parse_arguments():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(description='SMS Sending System using MyPhoneExplorer')
    parser.add_argument('--excel', '-e', type=str, help='Path to Excel file')
    parser.add_argument('--delay-min', '-min', type=int, default=5, help='Minimum delay between SMS (seconds)')
    parser.add_argument('--delay-max', '-max', type=int, default=10, help='Maximum delay between SMS (seconds)')
    parser.add_argument('--show-window', '-w', action='store_true', help='Show MyPhoneExplorer window')
    parser.add_argument('--test', '-t', action='store_true', help='Enable test mode')
    
    return parser.parse_args()

def main():
    """Main function with command line argument support"""
    args = parse_arguments()
    
    app = SMSSender(
        excel_path=args.excel,
        delay_min=args.delay_min,
        delay_max=args.delay_max,
        show_window=args.show_window,
        test_mode=args.test
    )
    
    app.run()

if __name__ == "__main__":
    main()