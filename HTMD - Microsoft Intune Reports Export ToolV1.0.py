#!/usr/bin/env python3
"""
Microsoft Intune Reports Export Tool - Dynamic Column Selection v1.0
Exports all available data first, then allows users to select/deselect columns post-export
"""

# Auto-install required packages
import sys
import subprocess
import importlib
import os

def check_and_install_packages():
    """Check and install required packages automatically"""
    print("🔍 Checking required packages...")
    
    # Define required packages (pyautogui is optional - will be installed when needed for PowerBI)
    required_packages = {
        'requests': 'requests',
        'pandas': 'pandas'
    }
    
    # Optional packages that enhance functionality
    optional_packages = {
        'pyautogui': 'pyautogui'  # Used for PowerBI automation
    }
    
    # Built-in modules that don't need installation
    builtin_modules = [
        'tkinter', 'json', 'os', 'csv', 'datetime', 'threading', 
        'webbrowser', 'urllib', 'http', 'socket', 'traceback', 
        'time', 'zipfile', 'subprocess'
    ]
    
    missing_packages = []
    
    # Check each required package
    for module_name, package_name in required_packages.items():
        try:
            importlib.import_module(module_name)
            print(f"✅ {module_name} - already installed")
        except ImportError:
            print(f"❌ {module_name} - not found")
            missing_packages.append(package_name)
    
    # Install missing packages
    if missing_packages:
        print(f"\n📦 Installing missing packages: {', '.join(missing_packages)}")
        print("⏳ This may take a moment...")
        
        for package in missing_packages:
            try:
                print(f"Installing {package}...")
                result = subprocess.run(
                    [sys.executable, '-m', 'pip', 'install', package], 
                    capture_output=True, 
                    text=True,
                    check=True
                )
                print(f"✅ {package} installed successfully")
            except subprocess.CalledProcessError as e:
                print(f"❌ Failed to install {package}: {e}")
                print(f"Please manually install: pip install {package}")
                return False
            except Exception as e:
                print(f"❌ Error installing {package}: {e}")
                return False
    
    print("✅ All required packages are ready!")
    
    # Check optional packages (don't block startup if missing)
    print("\n🔍 Checking optional packages...")
    for module_name, package_name in optional_packages.items():
        try:
            importlib.import_module(module_name)
            print(f"✅ {module_name} - available (enhanced PowerBI automation)")
        except ImportError:
            print(f"ℹ️ {module_name} - not installed (PowerBI automation will offer to install it)")
    
    return True

# Run package check before importing other modules
if __name__ == "__main__":
    if not check_and_install_packages():
        print("\n❌ Some packages could not be installed automatically.")
        print("Please install them manually using:")
        print("pip install requests pandas")
        input("Press Enter to continue anyway or Ctrl+C to exit...")

# Now import all required modules
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import requests
import json
import os
import csv
import pandas as pd
from datetime import datetime
import threading
import webbrowser
import urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
import socket
from urllib.parse import parse_qs, urlparse
import traceback
import time
import zipfile
import sys
import random
from collections import deque

class RateLimiter:
    """Rate limiter for Microsoft Graph API calls"""
    
    def __init__(self, requests_per_minute=600, requests_per_second=10):
        """
        Setup rate limiting - default 600/min, 10/sec for Graph API
        """
        self.requests_per_minute = requests_per_minute
        self.requests_per_second = requests_per_second
        self.minute_requests = deque()
        self.second_requests = deque()
        self.last_429_time = None
        self.throttle_until = None
    
    def wait_if_needed(self):
        """Wait before request if we're hitting limits"""
        import time
        from datetime import datetime, timedelta
        
        now = datetime.now()
        
        # Still throttled from previous 429 response
        if self.throttle_until and now < self.throttle_until:
            wait_time = (self.throttle_until - now).total_seconds()
            if wait_time > 0:
                time.sleep(wait_time)
        
        # Clean old requests (older than 1 minute)
        cutoff_minute = now - timedelta(minutes=1)
        while self.minute_requests and self.minute_requests[0] < cutoff_minute:
            self.minute_requests.popleft()
        
        # Clean old requests (older than 1 second)
        cutoff_second = now - timedelta(seconds=1)
        while self.second_requests and self.second_requests[0] < cutoff_second:
            self.second_requests.popleft()
        
        # Check per-minute limit
        if len(self.minute_requests) >= self.requests_per_minute:
            sleep_time = 60 - (now - self.minute_requests[0]).total_seconds()
            if sleep_time > 0:
                time.sleep(sleep_time)
        
        # Check per-second limit
        if len(self.second_requests) >= self.requests_per_second:
            sleep_time = 1 - (now - self.second_requests[0]).total_seconds()
            if sleep_time > 0:
                time.sleep(sleep_time)
        
        # Record this request
        current_time = datetime.now()
        self.minute_requests.append(current_time)
        self.second_requests.append(current_time)
    
    def handle_429_response(self, response):
        """Handle 429 rate limit response"""
        import time
        from datetime import datetime, timedelta
        
        # Get retry-after header (in seconds)
        retry_after = response.headers.get('Retry-After', '60')
        try:
            retry_seconds = int(retry_after)
        except ValueError:
            retry_seconds = 60
        
        # Add random delay to avoid simultaneous retries
        jitter = random.uniform(0.1, 0.3) * retry_seconds
        total_wait = retry_seconds + jitter
        
        # Set throttle state
        self.throttle_until = datetime.now() + timedelta(seconds=total_wait)
        self.last_429_time = datetime.now()
        
        return total_wait

class TimeoutManager:
    """Manages timeouts for different API operations"""
    
    @staticmethod
    def get_timeout_for_operation(operation_type, estimated_records=None):
        """Get timeout value based on operation type and data size"""
        
        base_timeouts = {
            'authentication': 60,
            'token_refresh': 30,
            'api_call': 120,
            'export_job_creation': 180,
            'export_job_status': 60,
            'file_download': 300,
            'large_export': 600
        }
        
        base_timeout = base_timeouts.get(operation_type, 120)
        
        # Adjust for estimated data size
        if estimated_records and operation_type in ['export_job_creation', 'large_export']:
            # Add time based on estimated records (1 second per 1000 records)
            additional_time = (estimated_records / 1000) * 1
            base_timeout = int(base_timeout + additional_time)
            
            # Cap at reasonable maximums
            if operation_type == 'export_job_creation':
                base_timeout = min(base_timeout, 1800)  # Max 30 minutes
            elif operation_type == 'large_export':
                base_timeout = min(base_timeout, 3600)  # Max 1 hour
        
        return base_timeout
    
    @staticmethod
    def get_exponential_backoff_delay(attempt, base_delay=1, max_delay=60):
        """Calculate retry delay with random component"""
        delay = min(base_delay * (2 ** attempt), max_delay)
        jitter = random.uniform(0.1, 0.3) * delay
        return delay + jitter

class AuthCallbackHandler(BaseHTTPRequestHandler):
    """Handle OAuth callback"""
    def do_GET(self):
        parsed_url = urlparse(self.path)
        query_params = parse_qs(parsed_url.query)
        
        if 'code' in query_params:
            self.server.auth_code = query_params['code'][0]
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(b"""
            <html>
                <body style='font-family: Arial; text-align: center; margin-top: 100px;'>
                    <h2 style='color: green;'>Authentication Successful!</h2>
                    <p>You can close this window and return to the application.</p>
                </body>
            </html>
            """)
        else:
            self.server.auth_code = None
            self.send_response(400)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            error_desc = query_params.get('error_description', ['Unknown error'])[0]
            self.wfile.write(f"""
            <html>
                <body style='font-family: Arial; text-align: center; margin-top: 100px;'>
                    <h2 style='color: red;'>Authentication Failed!</h2>
                    <p>Error: {error_desc}</p>
                </body>
            </html>
            """.encode())
    
    def log_message(self, format, *args):
        pass

class ReportViewer:
    """A dedicated window for viewing report data in a table format"""
    
    def __init__(self, parent, report_name, data, columns):
        self.parent = parent
        self.report_name = report_name
        self.data = data
        self.columns = columns
        self.viewer_window = None
        
        self.create_viewer_window()
    
    def create_viewer_window(self):
        """Create the report viewer window"""
        self.viewer_window = tk.Toplevel(self.parent.root)
        self.viewer_window.title("")  # Empty title for custom title bar
        self.viewer_window.geometry("1200x700")
        self.viewer_window.minsize(800, 500)
        
        # Configure window icon and properties
        self.viewer_window.configure(bg='#f5f5f5')
        self.viewer_window.transient(self.parent.root)
        self.viewer_window.grab_set()
        
        # Enable maximize/minimize buttons
        self.viewer_window.resizable(True, True)
        
        # Add window state tracking
        self.is_maximized = False
        
        # Bind window state events
        self.viewer_window.bind('<F11>', self.toggle_fullscreen)
        self.viewer_window.bind('<Double-Button-1>', self.on_title_double_click)
        
        # Custom title bar frame (white)
        title_bar_frame = tk.Frame(self.viewer_window, bg='white', height=40, relief='solid', bd=1)
        title_bar_frame.pack(fill='x')
        title_bar_frame.pack_propagate(False)
        
        # Title text on left
        title_text = tk.Label(title_bar_frame, text=f"Report Viewer - {self.report_name}", 
                             font=('Segoe UI', 11, 'bold'), 
                             bg='white', fg='#323130')
        title_text.pack(side='left', padx=15, pady=10)
        
        # Window controls frame (right side of white title bar)
        controls_frame = tk.Frame(title_bar_frame, bg='white')
        controls_frame.pack(side='right', padx=10, pady=5)
        
        # Minimize button
        minimize_btn = tk.Button(controls_frame, text="−", 
                               command=self.minimize_window,
                               font=('Segoe UI', 14, 'bold'),
                               bg='#e1e1e1', fg='#323130', width=3, height=1,
                               relief='flat', cursor='hand2',
                               activebackground='#d1d1d1', bd=0)
        minimize_btn.pack(side='left', padx=2)
        
        # Maximize/Restore button
        self.maximize_btn = tk.Button(controls_frame, text="□", 
                                    command=self.toggle_maximize,
                                    font=('Segoe UI', 12, 'bold'),
                                    bg='#e1e1e1', fg='#323130', width=3, height=1,
                                    relief='flat', cursor='hand2',
                                    activebackground='#d1d1d1', bd=0)
        self.maximize_btn.pack(side='left', padx=2)
        
        # Close button
        close_btn = tk.Button(controls_frame, text="×", 
                             command=self.close_viewer,
                             font=('Segoe UI', 14, 'bold'),
                             bg='#e81123', fg='white', width=3, height=1,
                             relief='flat', cursor='hand2',
                             activebackground='#c50e1f', bd=0)
        close_btn.pack(side='left', padx=2)
        
        # Header frame (blue header with report info)
        header_frame = tk.Frame(self.viewer_window, bg='#0078d4', height=60)
        header_frame.pack(fill='x', pady=(0, 10))
        header_frame.pack_propagate(False)
        
        # Report title
        title_label = tk.Label(header_frame, text=f"📊 {self.report_name}", 
                              font=('Segoe UI', 14, 'bold'), 
                              bg='#0078d4', fg='white')
        title_label.pack(side='left', padx=20, pady=15)
        
        # Record count label
        record_count = 0
        if self.data is not None:
            if hasattr(self.data, 'empty'):  # DataFrame
                record_count = len(self.data) if not self.data.empty else 0
            else:  # List or other iterable
                record_count = len(self.data) if self.data else 0
        
        count_label = tk.Label(header_frame, text=f"📈 {record_count:,} records", 
                              font=('Segoe UI', 10), 
                              bg='#0078d4', fg='white')
        count_label.pack(side='right', padx=20, pady=15)
        
        # Action buttons frame - pack BEFORE content frame to ensure visibility
        buttons_frame = tk.Frame(self.viewer_window, bg='#f5f5f5', height=60)
        buttons_frame.pack(fill='x', side='bottom', padx=10, pady=10)
        buttons_frame.pack_propagate(False)  # Prevent shrinking
        
        # Export CSV button
        export_btn = tk.Button(buttons_frame, text="📥 Export CSV", 
                              command=self.export_csv,
                              font=('Segoe UI', 10, 'bold'),
                              bg='#107c10', fg='white', padx=20, pady=8,
                              relief='flat', cursor='hand2',
                              activebackground='#0d5c0d')
        export_btn.pack(side='left', padx=(10, 10), pady=10)
        
        # Refresh button
        refresh_btn = tk.Button(buttons_frame, text="🔄 Refresh", 
                               command=self.refresh_data,
                               font=('Segoe UI', 10, 'bold'),
                               bg='#0078d4', fg='white', padx=20, pady=8,
                               relief='flat', cursor='hand2',
                               activebackground='#106ebe')
        refresh_btn.pack(side='left', padx=(0, 10), pady=10)
        
        # Status label
        self.status_label = tk.Label(buttons_frame, text="Ready", 
                                    font=('Segoe UI', 9), 
                                    bg='#f5f5f5', fg='#605e5c')
        self.status_label.pack(side='left', padx=20, pady=10)
        
        # Main content frame - pack AFTER buttons frame
        content_frame = tk.Frame(self.viewer_window, bg='#f5f5f5')
        content_frame.pack(fill='both', expand=True, padx=10, pady=(0, 0))
        
        # Create treeview for data display
        self.create_data_table(content_frame)
    
    def create_data_table(self, parent):
        """Create the data table with scrollbars"""
        # Table frame
        table_frame = tk.Frame(parent, bg='#f5f5f5')
        table_frame.pack(fill='both', expand=True)
        
        # Create treeview
        self.tree = ttk.Treeview(table_frame, show='headings', height=20)
        
        # Configure columns
        display_columns = self.columns[:50] if len(self.columns) > 50 else self.columns  # Limit columns for performance
        self.tree['columns'] = display_columns
        
        # Configure column headings and widths
        for col in display_columns:
            self.tree.heading(col, text=col, anchor='w')
            # Set dynamic column width based on content length
            max_width = max(len(col) * 8, 100)  # Minimum 100px
            self.tree.column(col, width=max_width, minwidth=80, anchor='w')
        
        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack scrollbars and treeview
        v_scrollbar.pack(side='right', fill='y')
        h_scrollbar.pack(side='bottom', fill='x')
        self.tree.pack(side='left', fill='both', expand=True)
        
        # Populate data
        self.populate_data(display_columns)
        
        # Configure treeview styling
        style = ttk.Style()
        style.configure("Treeview", rowheight=25, font=('Segoe UI', 9))
        style.configure("Treeview.Heading", font=('Segoe UI', 9, 'bold'))
    
    def populate_data(self, display_columns):
        """Populate the treeview with data"""
        # Check if data is empty - handle both DataFrame and list
        data_is_empty = False
        if self.data is None:
            data_is_empty = True
        elif hasattr(self.data, 'empty'):  # DataFrame
            data_is_empty = self.data.empty
        else:  # List or other iterable
            data_is_empty = len(self.data) == 0
            
        if data_is_empty:
            # No data available
            self.tree.insert('', 'end', values=['No data available'] + [''] * (len(display_columns) - 1))
            return
        
        # Add data rows (limit to first 1000 for performance)
        max_rows = min(len(self.data), 1000)
        
        # Different handling for DataFrame vs dict
        if hasattr(self.data, 'iterrows'):  # It's a pandas DataFrame
            for i, (index, row) in enumerate(self.data.head(max_rows).iterrows()):
                values = []
                for col in display_columns:
                    value = row.get(col, '') if col in row else ''
                    # Convert to string and limit length
                    str_value = str(value)[:100] if value is not None else ''
                    values.append(str_value)
                
                self.tree.insert('', 'end', values=values)
        else:  # It's list of dicts (traditional CSV data)
            for i, row in enumerate(self.data[:max_rows]):
                values = []
                for col in display_columns:
                    value = row.get(col, '') if isinstance(row, dict) else ''
                    # Convert to string and limit length
                    str_value = str(value)[:100] if value is not None else ''
                    values.append(str_value)
                
                self.tree.insert('', 'end', values=values)
        
        # Update status if data was truncated
        if len(self.data) > 1000:
            self.status_label.config(text=f"Showing first 1,000 of {len(self.data):,} records")
        elif len(display_columns) > 50:
            self.status_label.config(text=f"Showing first 50 of {len(self.columns)} columns")
    
    def export_csv(self):
        """Export the current data to CSV"""
        try:
            # Check if data is empty - handle both DataFrame and list
            data_is_empty = False
            if self.data is None:
                data_is_empty = True
            elif hasattr(self.data, 'empty'):  # DataFrame
                data_is_empty = self.data.empty
            else:  # List or other iterable
                data_is_empty = len(self.data) == 0
                
            if data_is_empty:
                messagebox.showwarning("No Data", "No data available to export.")
                return
            
            # Ask for file location
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Save Report As",
                initialfile=f"{self.report_name}_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            )
            
            if filename:
                # Export using pandas for better formatting
                import pandas as pd
                df = pd.DataFrame(self.data)
                df.to_csv(filename, index=False)
                
                file_size = os.path.getsize(filename) / (1024 * 1024)  # Size in MB
                self.status_label.config(text=f"✅ Exported to {filename} ({file_size:.2f} MB)")
                messagebox.showinfo("Export Complete", f"Report exported successfully!\n\nFile: {filename}")
                
        except Exception as e:
            error_msg = f"Export failed: {str(e)}"
            self.status_label.config(text=error_msg)
            messagebox.showerror("Export Error", error_msg)
    
    def close_viewer(self):
        """Close the viewer window"""
        self.viewer_window.grab_release()
        self.viewer_window.destroy()
    
    def minimize_window(self):
        """Minimize the viewer window"""
        self.viewer_window.iconify()
        self.status_label.config(text="Window minimized")
    
    def toggle_maximize(self):
        """Toggle between maximized and normal window state"""
        if self.is_maximized:
            # Restore to normal size
            self.viewer_window.state('normal')
            self.maximize_btn.config(text="□")
            self.is_maximized = False
            self.status_label.config(text="Window restored")
        else:
            # Maximize window
            self.viewer_window.state('zoomed')
            self.maximize_btn.config(text="❐")
            self.is_maximized = True
            self.status_label.config(text="Window maximized")
    
    def toggle_fullscreen(self, event=None):
        """Toggle fullscreen mode (F11)"""
        current_state = self.viewer_window.attributes('-fullscreen')
        self.viewer_window.attributes('-fullscreen', not current_state)
        if not current_state:
            self.status_label.config(text="Fullscreen mode (Press F11 to exit)")
        else:
            self.status_label.config(text="Exited fullscreen mode")
    
    def on_title_double_click(self, event):
        """Handle double-click on title bar to maximize/restore"""
        # Only trigger if double-click is on the title label area
        if event.widget == self.viewer_window or hasattr(event.widget, 'master'):
            self.toggle_maximize()
    
    def refresh_data(self):
        """Refresh the data display"""
        try:
            # Clear existing data
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Repopulate with current data
            display_columns = self.columns[:50] if len(self.columns) > 50 else self.columns
            self.populate_data(display_columns)
            
            self.status_label.config(text="✅ Data refreshed")
            
        except Exception as e:
            self.status_label.config(text=f"❌ Refresh failed: {str(e)}")

class ParameterDialog:
    """Dialog for collecting report parameters"""
    
    def __init__(self, parent, report_name, parameter_config):
        self.parent = parent
        self.report_name = report_name
        self.parameter_config = parameter_config
        self.result = None
        self.parameters = {}
        
        self.dialog = tk.Toplevel(parent.root)
        self.dialog.title(f"Configure Parameters - {report_name}")
        self.dialog.geometry("700x600")
        self.dialog.configure(bg='#f5f5f5')
        self.dialog.transient(parent.root)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)  # Allow resizing to see if content is hidden
        
        # Center the dialog
        self.dialog.geometry("+%d+%d" % (parent.root.winfo_rootx() + 50, parent.root.winfo_rooty() + 50))
        
        self.create_dialog_ui()
        
        # Wait for dialog to close
        self.dialog.wait_window()
    
    def has_date_parameters(self):
        """Check if this report has date parameters"""
        if 'parameters' not in self.parameter_config:
            return False
        
        for param_name, param_config in self.parameter_config['parameters'].items():
            if param_config.get('type') == 'date':
                return True
        
        return False
    
    def create_dialog_ui(self):
        """Create the parameter dialog UI"""
        
        # Header
        header_frame = tk.Frame(self.dialog, bg='#0078d4', height=80)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text=f"📊 Configure Parameters", 
                              font=('Segoe UI', 14, 'bold'), 
                              bg='#0078d4', fg='white')
        title_label.pack(side='left', padx=20, pady=20)
        
        report_label = tk.Label(header_frame, text=f"Report: {self.report_name}", 
                               font=('Segoe UI', 10), 
                               bg='#0078d4', fg='white')
        report_label.pack(side='right', padx=20, pady=20)
        
        # Main content with scrollbar if needed
        main_frame = tk.Frame(self.dialog, bg='#f5f5f5')
        main_frame.pack(fill='both', expand=True)
        
        canvas = tk.Canvas(main_frame, bg='#f5f5f5')
        scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        content_frame = tk.Frame(canvas, bg='#f5f5f5')
        
        content_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=content_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        scrollbar.pack(side="right", fill="y")
        
        # Description
        desc_text = self.parameter_config.get('description', 'Configure parameters for this report')
        desc_label = tk.Label(content_frame, text=desc_text, 
                             font=('Segoe UI', 10), 
                             bg='#f5f5f5', fg='#323130')
        desc_label.pack(anchor='w', pady=(0, 20))
        
        # Parameters section
        if 'parameters' in self.parameter_config:
            self.create_parameter_widgets(content_frame)
        
        # Quick templates section - only for reports with date parameters
        if self.has_date_parameters():
            self.create_templates_section(content_frame)
        
        # Buttons section (ensure it's always visible)
        button_frame = tk.Frame(self.dialog, bg='#f5f5f5', height=60)
        button_frame.pack(side='bottom', fill='x', padx=20, pady=20)
        button_frame.pack_propagate(False)  # Maintain fixed height
        
        # Cancel button
        cancel_btn = tk.Button(button_frame, text="Cancel", 
                              command=self.cancel_dialog,
                              font=('Segoe UI', 10),
                              bg='#e1e1e1', fg='#323130', padx=20, pady=8,
                              relief='flat', cursor='hand2')
        cancel_btn.pack(side='right', padx=(10, 0), pady=10)
        
        # OK button
        ok_btn = tk.Button(button_frame, text="Export with Parameters", 
                          command=self.ok_dialog,
                          font=('Segoe UI', 10, 'bold'),
                          bg='#0078d4', fg='white', padx=20, pady=8,
                          relief='flat', cursor='hand2')
        ok_btn.pack(side='right', pady=10)
        
        print(f"DEBUG: Created buttons for {self.report_name} dialog")  # Debug output
    
    def create_parameter_widgets(self, parent):
        """Create widgets for each parameter"""
        
        params_frame = tk.LabelFrame(parent, text="Required Parameters", 
                                    font=('Segoe UI', 10, 'bold'),
                                    bg='#f5f5f5', fg='#323130')
        params_frame.pack(fill='x', pady=(0, 20))
        
        self.param_widgets = {}
        
        for param_name, param_config in self.parameter_config['parameters'].items():
            param_frame = tk.Frame(params_frame, bg='#f5f5f5')
            param_frame.pack(fill='x', padx=10, pady=10)
            
            # Parameter label
            label_text = param_config.get('description', param_name)
            if param_config.get('required', False):
                label_text += " *"
            
            param_label = tk.Label(param_frame, text=label_text,
                                  font=('Segoe UI', 9),
                                  bg='#f5f5f5', fg='#323130')
            param_label.pack(anchor='w')
            
            # Add helpful hint for device selector
            if param_config.get('type') == 'device_selector':
                hint_label = tk.Label(param_frame, 
                                     text="💡 Tip: Type to search, paste device name from Intune portal, or select from dropdown",
                                     font=('Segoe UI', 8, 'italic'),
                                     bg='#f5f5f5', fg='#605e5c')
                hint_label.pack(anchor='w', pady=(2, 5))
            
            # Parameter widget based on type
            param_type = param_config.get('type', 'text')
            
            if param_type == 'device_selector':
                widget = self.create_device_selector(param_frame, param_name)
            elif param_type == 'policy_selector':
                widget = self.create_policy_selector(param_frame, param_name)
            elif param_type == 'date':
                widget = self.create_date_selector(param_frame, param_name)
            elif param_type == 'number':
                widget = self.create_number_input(param_frame, param_name)
            else:  # text
                widget = self.create_text_input(param_frame, param_name)
            
            self.param_widgets[param_name] = widget
    
    def create_device_selector(self, parent, param_name):
        """Create enhanced device selector with search and paste functionality"""
        
        # Container frame
        container = tk.Frame(parent, bg='#f5f5f5')
        container.pack(fill='x', pady=(5, 0))
        
        # Device variable and mapping
        device_var = tk.StringVar()
        self.device_id_mapping = {}  # Maps display name to device ID
        self.device_search_data = []  # All devices for search
        
        # Create searchable combobox (editable for search and paste)
        device_combo = ttk.Combobox(container, textvariable=device_var, 
                                   state='normal', width=50)
        device_combo.pack(side='left', padx=(0, 5))
        
        # Add placeholder text
        device_combo.insert(0, "Type device name or paste from Intune portal...")
        device_combo.bind('<FocusIn>', lambda e: self.on_device_focus(device_combo, device_var))
        device_combo.bind('<KeyRelease>', lambda e: self.on_device_search(device_combo, device_var, e))
        device_combo.bind('<<ComboboxSelected>>', lambda e: self.on_device_selected(device_combo, device_var))
        
        # Refresh button
        refresh_btn = tk.Button(container, text="🔄", 
                               command=lambda: self.load_devices(device_combo, device_var),
                               font=('Segoe UI', 8),
                               bg='#e1e1e1', fg='#323130',
                               width=3, relief='flat', cursor='hand2')
        refresh_btn.pack(side='left', padx=(2, 5))
        
        # Paste button
        paste_btn = tk.Button(container, text="📋", 
                             command=lambda: self.paste_device_name(device_combo, device_var),
                             font=('Segoe UI', 8),
                             bg='#0078d4', fg='white',
                             width=3, relief='flat', cursor='hand2')
        paste_btn.pack(side='left', padx=(0, 5))
        
        # Clear button
        clear_btn = tk.Button(container, text="✖", 
                             command=lambda: self.clear_device_selection(device_combo, device_var),
                             font=('Segoe UI', 8),
                             bg='#d13438', fg='white',
                             width=3, relief='flat', cursor='hand2')
        clear_btn.pack(side='left')
        
        # Store references for later use
        device_var.combo_widget = device_combo
        
        # Load devices initially
        self.parent.root.after(100, lambda: self.load_devices(device_combo, device_var))
        
        return device_var
    
    def create_policy_selector(self, parent, param_name):
        """Create enhanced policy selector dropdown with search and paste functionality"""
        
        container = tk.Frame(parent, bg='#f5f5f5')
        container.pack(fill='x', pady=(5, 0))
        
        # Top row: Search box and buttons
        search_frame = tk.Frame(container, bg='#f5f5f5')
        search_frame.pack(fill='x', pady=(0, 5))
        
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side='left', padx=(0, 5))
        
        # Add context-aware placeholder text
        policy_type = self.get_policy_type_name().lower()
        placeholder_text = f"Search {policy_type}..."
        search_entry.insert(0, placeholder_text)
        search_entry.config(fg='gray')
        
        def on_focus_in(event):
            if search_entry.get() == placeholder_text:
                search_entry.delete(0, tk.END)
                search_entry.config(fg='black')
        
        def on_focus_out(event):
            if not search_entry.get():
                search_entry.insert(0, placeholder_text)
                search_entry.config(fg='gray')
        
        search_entry.bind('<FocusIn>', on_focus_in)
        search_entry.bind('<FocusOut>', on_focus_out)
        
        paste_btn = tk.Button(search_frame, text="📋", 
                             command=lambda: self.paste_policy_name(search_var),
                             font=('Segoe UI', 8),
                             bg='#e1e1e1', fg='#323130',
                             width=3, relief='flat', cursor='hand2')
        paste_btn.pack(side='left', padx=(0, 5))
        
        clear_btn = tk.Button(search_frame, text="✕", 
                             command=lambda: self.clear_policy_search(search_var, policy_combo),
                             font=('Segoe UI', 8),
                             bg='#e1e1e1', fg='#323130',
                             width=3, relief='flat', cursor='hand2')
        clear_btn.pack(side='left', padx=(0, 5))
        
        refresh_btn = tk.Button(search_frame, text="🔄", 
                               command=lambda: self.load_policies(policy_combo, policy_var),
                               font=('Segoe UI', 8),
                               bg='#e1e1e1', fg='#323130',
                               width=3, relief='flat', cursor='hand2')
        refresh_btn.pack(side='left')
        
        # Bottom row: Policy dropdown
        policy_var = tk.StringVar()
        policy_combo = ttk.Combobox(container, textvariable=policy_var, 
                                   state='readonly', width=70)
        
        policy_combo['values'] = ['Loading policies...']
        policy_combo.pack(fill='x', pady=(0, 0))
        
        # Bind search functionality
        search_var.trace('w', lambda *args: self.filter_policies(search_var.get(), policy_combo, policy_var))
        
        # Store references for filtering
        policy_combo.all_policies = []
        policy_combo.search_var = search_var
        
        self.parent.root.after(100, lambda: self.load_policies(policy_combo, policy_var))
        
        return policy_var
    
    def create_date_selector(self, parent, param_name):
        """Create date selector"""
        
        container = tk.Frame(parent, bg='#f5f5f5')
        container.pack(fill='x', pady=(5, 0))
        
        date_var = tk.StringVar()
        date_entry = tk.Entry(container, textvariable=date_var, width=20)
        date_entry.pack(side='left', padx=(0, 10))
        
        # Set default date based on parameter name
        from datetime import datetime, timedelta
        if 'start' in param_name.lower():
            default_date = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
        else:
            default_date = datetime.now().strftime("%Y-%m-%d")
        
        date_var.set(default_date)
        
        # Quick buttons
        today_btn = tk.Button(container, text="Today", 
                             command=lambda: date_var.set(datetime.now().strftime("%Y-%m-%d")),
                             font=('Segoe UI', 8),
                             bg='#e1e1e1', fg='#323130', relief='flat', cursor='hand2')
        today_btn.pack(side='left', padx=(0, 5))
        
        week_btn = tk.Button(container, text="7 Days Ago", 
                            command=lambda: date_var.set((datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")),
                            font=('Segoe UI', 8),
                            bg='#e1e1e1', fg='#323130', relief='flat', cursor='hand2')
        week_btn.pack(side='left', padx=(0, 5))
        
        month_btn = tk.Button(container, text="30 Days Ago", 
                             command=lambda: date_var.set((datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")),
                             font=('Segoe UI', 8),
                             bg='#e1e1e1', fg='#323130', relief='flat', cursor='hand2')
        month_btn.pack(side='left')
        
        return date_var
    
    def create_number_input(self, parent, param_name):
        """Create number input"""
        
        number_var = tk.StringVar()
        number_entry = tk.Entry(parent, textvariable=number_var, width=20)
        number_entry.pack(anchor='w', pady=(5, 0))
        
        # Set default based on parameter
        if 'top' in param_name.lower():
            number_var.set("1000")
        
        return number_var
    
    def create_text_input(self, parent, param_name):
        """Create text input"""
        
        text_var = tk.StringVar()
        text_entry = tk.Entry(parent, textvariable=text_var, width=50)
        text_entry.pack(anchor='w', pady=(5, 0))
        
        return text_var
    
    def create_templates_section(self, parent):
        """Create quick template section"""
        
        templates_frame = tk.LabelFrame(parent, text="Quick Templates", 
                                       font=('Segoe UI', 10, 'bold'),
                                       bg='#f5f5f5', fg='#323130')
        templates_frame.pack(fill='x', pady=(0, 20))
        
        templates = {
            "Last 30 Days": {"startDate": "30_days_ago", "endDate": "today"},
            "Last 7 Days": {"startDate": "7_days_ago", "endDate": "today"},
            "This Month": {"startDate": "month_start", "endDate": "today"}
        }
        
        for template_name, template_params in templates.items():
            btn = tk.Button(templates_frame, text=template_name,
                           command=lambda t=template_params: self.apply_template(t),
                           font=('Segoe UI', 9),
                           bg='#0078d4', fg='white', padx=15, pady=5,
                           relief='flat', cursor='hand2')
            btn.pack(side='left', padx=5, pady=10)
    
    def apply_template(self, template_params):
        """Apply a parameter template"""
        from datetime import datetime, timedelta
        
        for param_name, param_value in template_params.items():
            if param_name in self.param_widgets:
                widget = self.param_widgets[param_name]
                
                if param_value == "today":
                    widget.set(datetime.now().strftime("%Y-%m-%d"))
                elif param_value == "7_days_ago":
                    widget.set((datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d"))
                elif param_value == "30_days_ago":
                    widget.set((datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"))
                elif param_value == "month_start":
                    widget.set(datetime.now().replace(day=1).strftime("%Y-%m-%d"))
    
    def load_devices(self, combo_widget, var_widget):
        """Load devices from Microsoft Graph with enhanced display"""
        try:
            if not self.parent.access_token:
                combo_widget['values'] = ['Please login first']
                return
            
            combo_widget['values'] = ['Loading devices...']
            var_widget.set('Loading devices...')
            
            # Make API call to get devices with more details
            url = f"{self.parent.graph_base_url}/deviceManagement/managedDevices"
            params = {
                '$select': 'id,deviceName,userPrincipalName,model,manufacturer,operatingSystem,lastSyncDateTime,complianceState',
                '$top': 1000  # Increased limit
            }
            
            devices = self.parent.fetch_all_graph_pages(
                url,
                params=params,
                operation_type='api_call'
            )
                
            # Store raw device data for search
            self.device_search_data = devices
            
            # Clear previous mappings
            self.device_id_mapping = {}
            
            device_display_list = []
            for device in devices:
                device_name = device.get('deviceName', 'Unknown')
                user_name = device.get('userPrincipalName', 'No User')
                device_id = device.get('id', '')
                
                # Create user-friendly display name (just device and user)
                display_name = f"{device_name} ({user_name})"
                
                # Map display name to device ID
                self.device_id_mapping[display_name] = device_id
                device_display_list.append(display_name)
            
            if device_display_list:
                # Sort alphabetically
                device_display_list.sort()
                combo_widget['values'] = device_display_list
                var_widget.set('')  # Clear loading text
                
                # Update placeholder
                if hasattr(combo_widget, 'set'):
                    placeholder = f"Search {len(device_display_list)} devices or paste device name..."
                    combo_widget.delete(0, 'end')
                    combo_widget.insert(0, placeholder)
            else:
                combo_widget['values'] = ['No devices found in tenant']
                var_widget.set('')
                
        except Exception as e:
            combo_widget['values'] = [f'Error: {str(e)}']
            var_widget.set('')
    
    def on_device_focus(self, combo_widget, var_widget):
        """Handle focus event on device selector"""
        current_text = var_widget.get()
        if current_text == "Type device name or paste from Intune portal...":
            var_widget.set("")
    
    def on_device_search(self, combo_widget, var_widget, event):
        """Handle real-time search in device selector"""
        search_text = var_widget.get().lower()
        
        if len(search_text) < 2:  # Only search after 2 characters
            return
        
        # Filter devices based on search text
        filtered_devices = []
        for device in self.device_search_data:
            device_name = device.get('deviceName', '').lower()
            user_name = device.get('userPrincipalName', '').lower()
            
            if (search_text in device_name or 
                search_text in user_name):
                display_name = f"{device.get('deviceName', 'Unknown')} ({device.get('userPrincipalName', 'No User')})"
                filtered_devices.append(display_name)
        
        # Update dropdown values
        if filtered_devices:
            combo_widget['values'] = filtered_devices[:20]  # Limit to 20 results
        else:
            combo_widget['values'] = ['No matching devices found']
    
    def on_device_selected(self, combo_widget, var_widget):
        """Handle device selection from dropdown"""
        selected = var_widget.get()
        if selected and selected in self.device_id_mapping:
            # Store the selected device info
            var_widget.selected_device_id = self.device_id_mapping[selected]
    
    def paste_device_name(self, combo_widget, var_widget):
        """Paste device name from clipboard and attempt to match"""
        try:
            # Get clipboard content
            clipboard_text = self.dialog.clipboard_get().strip()
            
            if not clipboard_text:
                return
            
            # Try to find matching device
            matched_device = None
            for device in self.device_search_data:
                device_name = device.get('deviceName', '')
                
                # Try exact match first
                if clipboard_text.lower() == device_name.lower():
                    matched_device = device
                    break
                # Try partial match
                elif clipboard_text.lower() in device_name.lower():
                    matched_device = device
                    break
            
            if matched_device:
                display_name = f"{matched_device.get('deviceName', 'Unknown')} ({matched_device.get('userPrincipalName', 'No User')})"
                var_widget.set(display_name)
                var_widget.selected_device_id = matched_device.get('id')
                
                # Update status
                self.show_paste_status("✅ Device found and selected")
            else:
                # Just set the pasted text and let user search
                var_widget.set(clipboard_text)
                self.show_paste_status("📝 Pasted text - use search to find device")
                
        except Exception as e:
            self.show_paste_status(f"❌ Paste failed: {str(e)}")
    
    def clear_device_selection(self, combo_widget, var_widget):
        """Clear device selection"""
        var_widget.set("")
        var_widget.selected_device_id = None
        combo_widget['values'] = [device for device in self.device_id_mapping.keys()]
    
    def show_paste_status(self, message):
        """Show temporary status message for paste operation"""
        # Find status label in dialog if it exists
        try:
            status_label = getattr(self, 'temp_status_label', None)
            if not status_label:
                # Create temporary status label
                status_frame = tk.Frame(self.dialog, bg='#f5f5f5')
                status_frame.pack(fill='x', padx=20)
                
                self.temp_status_label = tk.Label(status_frame, text="", 
                                                 font=('Segoe UI', 8), 
                                                 bg='#f5f5f5', fg='#605e5c')
                self.temp_status_label.pack(anchor='w', pady=2)
                status_label = self.temp_status_label
            
            status_label.config(text=message)
            
            # Clear message after 3 seconds
            self.dialog.after(3000, lambda: status_label.config(text=""))
        except:
            pass  # Fail silently if UI elements not available
    
    def load_policies(self, combo_widget, var_widget):
        """Load policies from Microsoft Graph based on report type"""
        try:
            if not self.parent.access_token:
                combo_widget['values'] = ['Please login first']
                return
            
            combo_widget['values'] = ['Loading...']
            var_widget.set('Loading...')
            
            # Check what policy type to load
            policy_endpoint = self.get_policy_endpoint_for_report()
            
            # Make API call to get policies
            url = f"{self.parent.graph_base_url}{policy_endpoint}"
            params = {'$select': 'id,displayName', '$top': 100}
            
            policies = self.parent.fetch_all_graph_pages(
                url,
                params=params,
                operation_type='api_call'
            )
                
            policy_list = []
            policy_mapping = {}  # Map policy names to IDs
            policy_type = self.get_policy_type_name()
            
            for policy in policies:
                display_name = policy.get('displayName', 'Unnamed Policy')
                policy_id = policy.get('id')
                policy_list.append(display_name)
                policy_mapping[display_name] = policy_id
            
            if policy_list:
                combo_widget.all_policies = policy_list
                combo_widget.policy_mapping = policy_mapping  # Store name-to-ID mapping
                combo_widget['values'] = policy_list
                # Store mapping at dialog level for easy access during parameter collection
                if not hasattr(self, 'policy_name_to_id_mapping'):
                    self.policy_name_to_id_mapping = {}
                self.policy_name_to_id_mapping.update(policy_mapping)
                var_widget.set('')
                self.log_policy_load_success(policy_type, len(policy_list))
            else:
                combo_widget.all_policies = []
                combo_widget.policy_mapping = {}
                combo_widget['values'] = [f'No {policy_type.lower()} found']
                
        except Exception as e:
            combo_widget['values'] = [f'Error loading policies: {str(e)}']
    
    def get_policy_endpoint_for_report(self):
        """Get the appropriate policy endpoint based on the current report"""
        report_name = getattr(self, 'report_name', '')
        
        # Map report types to their appropriate policy endpoints
        policy_mappings = {
            # Windows Update Policies
            'qualityupdate': '/deviceManagement/windowsQualityUpdateProfiles',
            'featureupdate': '/deviceManagement/windowsFeatureUpdateProfiles', 
            'driverupdate': '/deviceManagement/windowsDriverUpdateProfiles',
            'updatering': '/deviceManagement/windowsUpdateForBusinessConfigurations',
            
            # Compliance Policies  
            'compliance': '/deviceManagement/deviceCompliancePolicies',
            'compliant': '/deviceManagement/deviceCompliancePolicies',
            'noncompliant': '/deviceManagement/deviceCompliancePolicies',
            
            # Configuration Policies
            'configuration': '/deviceManagement/deviceConfigurations',
            'config': '/deviceManagement/deviceConfigurations',
            'setting': '/deviceManagement/deviceConfigurations',
            
            # Enrollment Policies
            'enrollment': '/deviceManagement/deviceEnrollmentConfigurations',
            
            # App Protection Policies
            'appprotection': '/deviceAppManagement/managedAppPolicies',
            'mam': '/deviceAppManagement/managedAppPolicies',
            
            # Conditional Access
            'conditionalaccess': '/identity/conditionalAccess/policies',
            
            # Endpoint Security
            'security': '/deviceManagement/intents',
            'antivirus': '/deviceManagement/intents',
            'firewall': '/deviceManagement/intents',
        }
        
        # Check report name against mappings
        report_lower = report_name.lower()
        for keyword, endpoint in policy_mappings.items():
            if keyword in report_lower:
                return endpoint
                
        # Default to compliance policies if no match found
        return '/deviceManagement/deviceCompliancePolicies'
    
    def get_policy_type_name(self):
        """Get user-friendly policy type name based on endpoint"""
        endpoint = self.get_policy_endpoint_for_report()
        
        type_names = {
            '/deviceManagement/windowsQualityUpdateProfiles': 'Quality Update Profiles',
            '/deviceManagement/windowsFeatureUpdateProfiles': 'Feature Update Profiles',
            '/deviceManagement/windowsDriverUpdateProfiles': 'Driver Update Profiles', 
            '/deviceManagement/windowsUpdateForBusinessConfigurations': 'Windows Update Rings',
            '/deviceManagement/deviceCompliancePolicies': 'Compliance Policies',
            '/deviceManagement/deviceConfigurations': 'Configuration Profiles',
            '/deviceManagement/deviceEnrollmentConfigurations': 'Enrollment Configurations',
            '/deviceAppManagement/managedAppPolicies': 'App Protection Policies',
            '/identity/conditionalAccess/policies': 'Conditional Access Policies',
            '/deviceManagement/intents': 'Endpoint Security Policies'
        }
        
        return type_names.get(endpoint, 'Policies')
    
    def log_policy_load_success(self, policy_type, count):
        """Log successful policy loading"""
        if hasattr(self.parent, 'log_message'):
            self.parent.log_message(f"✅ Loaded {count} {policy_type} for {self.report_name}", 'info')
    
    def paste_policy_name(self, search_var):
        """Paste policy name from clipboard"""
        try:
            clipboard_text = self.dialog.clipboard_get().strip()
            if clipboard_text:
                # Clean up the pasted text (remove extra spaces, newlines)
                cleaned_text = ' '.join(clipboard_text.split())
                search_var.set(cleaned_text)
                self.show_paste_status("📝 Pasted policy name for search")
        except Exception as e:
            self.show_paste_status(f"❌ Paste failed: {str(e)}")
    
    def clear_policy_search(self, search_var, combo_widget):
        """Clear policy search"""
        search_var.set("")
        if hasattr(combo_widget, 'all_policies'):
            combo_widget['values'] = combo_widget.all_policies
    
    def filter_policies(self, search_text, combo_widget, policy_var):
        """Filter policies based on search text"""
        if not hasattr(combo_widget, 'all_policies'):
            return
        
        # Ignore placeholder text (dynamic based on policy type)
        policy_type = self.get_policy_type_name().lower()
        placeholder_text = f"Search {policy_type}..."
        if not search_text or search_text == placeholder_text:
            combo_widget['values'] = combo_widget.all_policies
            return
        
        search_lower = search_text.lower()
        filtered_policies = []
        
        for policy in combo_widget.all_policies:
            if search_lower in policy.lower():
                filtered_policies.append(policy)
        
        combo_widget['values'] = filtered_policies if filtered_policies else ['No policies match search']
        
        # Auto-select if only one match
        if len(filtered_policies) == 1 and filtered_policies[0] != 'No policies match search':
            policy_var.set(filtered_policies[0])
    
    def validate_parameters(self):
        """Validate all parameters"""
        
        if 'parameters' not in self.parameter_config:
            return True
        
        for param_name, param_config in self.parameter_config['parameters'].items():
            if param_config.get('required', False):
                widget = self.param_widgets.get(param_name)
                if widget and not widget.get().strip():
                    messagebox.showerror("Validation Error", 
                                       f"Parameter '{param_config.get('description', param_name)}' is required")
                    return False
        
        return True
    
    def collect_parameters(self):
        """Collect all parameter values"""
        
        parameters = {}
        
        if 'parameters' in self.parameter_config:
            for param_name, param_config in self.parameter_config['parameters'].items():
                widget = self.param_widgets.get(param_name)
                if widget:
                    value = widget.get().strip()
                    
                    if value and value != "Type device name or paste from Intune portal..." and not value.startswith("Search "):
                        # Special date formats
                        if param_config.get('type') == 'device_selector':
                            # Check if we have a selected device ID stored
                            if hasattr(widget, 'selected_device_id') and widget.selected_device_id:
                                parameters[param_name] = widget.selected_device_id
                            elif value in self.device_id_mapping:
                                # Use mapping to get device ID
                                parameters[param_name] = self.device_id_mapping[value]
                            else:
                                # Try to find device by name search
                                matched_device = self.find_device_by_name(value)
                                if matched_device:
                                    parameters[param_name] = matched_device.get('id')
                                else:
                                    # Keep the user-entered value (might be manually entered ID)
                                    parameters[param_name] = value
                        elif param_config.get('type') == 'policy_selector':
                            # Get policy ID from policy name using the stored mapping
                            if hasattr(self, 'policy_name_to_id_mapping') and value in self.policy_name_to_id_mapping:
                                parameters[param_name] = self.policy_name_to_id_mapping[value]
                            else:
                                # Fallback: if it's still in old format, extract ID
                                if '|' in value:
                                    parameters[param_name] = value.split('|')[0]
                                else:
                                    parameters[param_name] = value
                        else:
                            parameters[param_name] = value
        
        return parameters
    
    def find_device_by_name(self, search_name):
        """Find device by name from loaded device data"""
        if not hasattr(self, 'device_search_data'):
            return None
            
        search_lower = search_name.lower()
        
        # Try exact match first
        for device in self.device_search_data:
            device_name = device.get('deviceName', '').lower()
            if search_lower == device_name:
                return device
        
        # Try partial match
        for device in self.device_search_data:
            device_name = device.get('deviceName', '').lower()
            if search_lower in device_name:
                return device
        
        return None
    
    def ok_dialog(self):
        """Handle OK button click"""
        
        if self.validate_parameters():
            self.parameters = self.collect_parameters()
            self.result = 'ok'
            self.dialog.destroy()
    
    def cancel_dialog(self):
        """Handle Cancel button click"""
        
        self.result = 'cancel'
        self.dialog.destroy()

class IntuneReportsGUI:
    def __init__(self):
        # Configuration
        self.client_id = "enter the client id here"
        self.client_secret = "enter the client secret here"
        self.tenant_id = "enter the tenant id here"
        self.redirect_uri = "http://localhost:8080/callback"
        
        # Graph API endpoints
        self.graph_base_url = "https://graph.microsoft.com/v1.0"
        self.beta_base_url = "https://graph.microsoft.com/beta"
        self.export_endpoint = f"{self.beta_base_url}/deviceManagement/reports/exportJobs"
        
        # Authentication state
        self.access_token = None
        self.refresh_token = None
        self.user_info = None
        self.token_expires_at = None
        self.token_issued_at = None
        
        # Enterprise features
        self.rate_limiter = RateLimiter()
        self.timeout_manager = TimeoutManager()
        
        # Export data
        self.current_export_data = None
        self.current_columns = []
        self.column_vars = {}
        self.learned_parameters = {}  # Cache for learned report parameters
        
        # Available reports (all 179 Intune reports from official Microsoft list)
        self.available_reports = {
            # Security & Antivirus Reports
            "ActiveMalware": "Active Malware Detections",
            "DefenderAgents": "Microsoft Defender Agent Status", 
            "Malware": "Detected Malware Reports",
            "UnhealthyDefenderAgents": "Unhealthy Defender Endpoints",
            "FirewallStatus": "MDM Firewall Status for Windows 10+",
            "FirewallUnhealthyStatus": "Firewall Unhealthy Status",
            
            # Apps & Application Management
            "AllAppsList": "All Apps List",
            "FilteredAppsList": "Filtered Apps List", 
            "AppInstallStatusAggregate": "App Install Status Aggregate",
            "AppInvAggregate": "Discovered Apps Aggregate",
            "AppInvByDevice": "Discovered Apps by Device",
            "AppInvRawData": "Discovered Apps Raw Data",
            "CatalogAppsUpdateList": "Enterprise App Catalog Updates",
            "DependentAppsInstallStatus": "Dependent Apps Install Status",
            "DeviceInstallStatusByApp": "Device Install Status by App",
            "DevicesByAppInv": "Devices by App Inventory",
            "OrgAppsInstallStatus": "Org Apps Install Status",
            "UserInstallStatusAggregateByApp": "User Install Status by App",
            
            # Mobile Application Management (MAM)
            "MAMAppConfigurationStatus": "MAM App Configuration Status",
            "MAMAppConfigurationStatusScopedV2": "MAM App Config Status Scoped V2",
            "MAMAppConfigurationStatusV2": "MAM App Configuration Status V2",
            "MAMAppProtectionStatus": "MAM App Protection Status (iOS/Android)",
            "MAMAppProtectionStatusScopedV2": "MAM App Protection Status Scoped V2",
            "MAMAppProtectionStatusV2": "MAM App Protection Status V2",
            
            # Device Management & Inventory
            "Devices": "All Managed Devices",
            "DevicesWithInventory": "Devices with Hardware Inventory",
            "AllDeviceCertificates": "All Device Certificates",
            "CertificatesByRAPolicy": "Certificates by Registration Authority Policy",
            "TpmAttestationStatus": "TPM Attestation Status",
            "MEMUpgradeReadinessOrgAsset": "MEM Upgrade Readiness Org Assets",
            
            # Device Compliance
            "DeviceCompliance": "Device Compliance Status",
            "DeviceComplianceTrend": "Device Compliance Trend",
            "DeviceNonCompliance": "Device Non-Compliance Report",
            "DevicesWithoutCompliancePolicy": "Devices Without Compliance Policy",
            "DevicesWithoutCompliancePolicyV3": "Devices Without Compliance Policy V3",
            "NonCompliantDevicesByCompliancePolicy": "Non-Compliant Devices by Policy",
            "NonCompliantDevicesByCompliancePolicyV3": "Non-Compliant Devices by Policy V3",
            "NoncompliantDevicesAndSettings": "Non-Compliant Devices and Settings",
            "NoncompliantDevicesAndSettingsV3": "Non-Compliant Devices and Settings V3",
            "NoncompliantDevicesToBeRetired": "Non-Compliant Devices to be Retired",
            "NonCompliantCompliancePoliciesAggregate": "Non-Compliant Policies Aggregate",
            "NonCompliantCompliancePoliciesAggregateV3": "Non-Compliant Policies Aggregate V3",
            "NonCompliantConfigurationPoliciesAggregateWithPF": "Non-Compliant Config Policies with PF",
            "NonCompliantConfigurationPoliciesAggregateWithPFV3": "Non-Compliant Config Policies with PF V3",
            
            # Configuration Policies
            "ConfigurationPolicyAggregate": "Configuration Policy Aggregate",
            "ConfigurationPolicyAggregateV3": "Configuration Policy Aggregate V3",
            "ConfigurationPolicyDeviceAggregates": "Configuration Policy Device Aggregates",
            "ConfigurationPolicyDeviceAggregatesV3": "Configuration Policy Device Aggregates V3",
            "ConfigurationPolicyDeviceAggregatesWithPF": "Configuration Policy Device Aggregates with PF",
            "ConfigurationPolicyDeviceAggregatesWithPFV3": "Configuration Policy Device Aggregates with PF V3",
            "DeviceConfigurationPolicyStatuses": "Device Configuration Policy Status",
            "DeviceConfigurationPolicyStatusesV3": "Device Configuration Policy Status V3",
            "DeviceConfigurationPolicyStatusesWithPF": "Device Configuration Policy Status with PF",
            "DeviceConfigurationPolicyStatusesWithPFV3": "Device Configuration Policy Status with PF V3",
            "Policies": "All Device Policies",
            
            # Device Assignment & Status Reports
            "DeviceAssignmentStatusByConfigurationPolicy": "Device Assignment Status by Config Policy",
            "DeviceAssignmentStatusByConfigurationPolicyForAC": "Device Assignment Status for App Control",
            "DeviceAssignmentStatusByConfigurationPolicyForASR": "Device Assignment Status for ASR",
            "DeviceAssignmentStatusByConfigurationPolicyForEDR": "Device Assignment Status for EDR",
            "DeviceAssignmentStatusByConfigurationPolicyV3": "Device Assignment Status by Config Policy V3",
            "DeviceStatusesByConfigurationProfile": "Device Status by Configuration Profile",
            "DeviceStatusesByConfigurationProfileForAppControl": "Device Status by Config Profile for App Control",
            "DeviceStatusesByConfigurationProfileForASR": "Device Status by Config Profile for ASR",
            "DeviceStatusesByConfigurationProfileForEDR": "Device Status by Config Profile for EDR",
            "DeviceStatusesByConfigurationProfileV3": "Device Status by Configuration Profile V3",
            "DeviceStatusesByConfigurationProfileWithPF": "Device Status by Config Profile with PF",
            "DeviceStatusesByConfigurationProfileWithPFV3": "Device Status by Config Profile with PF V3",
            "DeviceStatusesByInventoryPolicyWithPF": "Device Status by Inventory Policy with PF",
            "DeviceStatusesByInventoryPolicyWithPFV3": "Device Status by Inventory Policy with PF V3",
            
            # Compliance Policy Reports
            "DevicePoliciesComplianceReport": "Device Policies Compliance Report",
            "DevicePoliciesComplianceReportV3": "Device Policies Compliance Report V3",
            "DevicePolicySettingsComplianceReport": "Device Policy Settings Compliance Report",
            "DevicePolicySettingsComplianceReportV3": "Device Policy Settings Compliance Report V3",
            "DeviceStatusByCompliacePolicyReport": "Device Status by Compliance Policy Report",
            "DeviceStatusByCompliacePolicyReportV3": "Device Status by Compliance Policy Report V3",
            "DeviceStatusByCompliancePolicySettingReport": "Device Status by Compliance Policy Setting",
            "DeviceStatusByCompliancePolicySettingReportV3": "Device Status by Compliance Policy Setting V3",
            "DeviceStatusSummaryByCompliacePolicyReport": "Device Status Summary by Compliance Policy",
            "DeviceStatusSummaryByCompliacePolicyReportV3": "Device Status Summary by Compliance Policy V3",
            "DeviceStatusSummaryByCompliancePolicySettingsReport": "Device Status Summary by Compliance Policy Settings",
            "DeviceStatusSummaryByCompliancePolicySettingsReportV3": "Device Status Summary by Compliance Policy Settings V3",
            "DevicesStatusByPolicyPlatformComplianceReport": "Devices Status by Policy Platform Compliance",
            "DevicesStatusByPolicyPlatformComplianceReportV3": "Devices Status by Policy Platform Compliance V3",
            "DevicesStatusBySettingReport": "Devices Status by Setting Report",
            "DevicesStatusBySettingReportV3": "Devices Status by Setting Report V3",
            "PolicyComplianceAggReport": "Policy Compliance Aggregate Report",
            "PolicyComplianceAggReportV3": "Policy Compliance Aggregate Report V3",
            "PolicyNonComplianceAgg": "Policy Non-Compliance Aggregate",
            "PolicyNonComplianceAggVer3": "Policy Non-Compliance Aggregate V3",
            "PolicyNonComplianceNew": "Policy Non-Compliance New",
            "PolicyNonComplianceNewV3": "Policy Non-Compliance New V3",
            "SettingComplianceAggReport": "Setting Compliance Aggregate Report",
            "SettingComplianceAggReportV3": "Setting Compliance Aggregate Report V3",
            
            # Windows Updates
            "FeatureUpdateDeviceState": "Feature Update Device State",
            "FeatureUpdatePolicyFailuresAggregate": "Feature Update Policy Failures Aggregate",
            "FeatureUpdatePolicyStatusSummary": "Feature Update Policy Status Summary",
            "QualityUpdateDeviceErrorsByPolicy": "Quality Update Device Errors by Policy",
            "QualityUpdateDeviceStatusByPolicy": "Quality Update Device Status by Policy",
            "QualityUpdatePolicyStatusSummary": "Quality Update Policy Status Summary",
            "WindowsUpdatePerPolicyPerDeviceStatus": "Windows Update per Policy per Device Status",
            "DriverUpdatePolicyStatusSummary": "Driver Update Policy Status Summary",
            "DeviceFailuresByFeatureUpdatePolicy": "Device Failures by Feature Update Policy",
            "WindowsDeviceHealthAttestationReport": "Windows Device Health Attestation Report",
            
            # Enrollment & Autopilot
            "EnrollmentActivity": "Device Enrollment Activity",
            "DeviceEnrollmentFailures": "Device Enrollment Failures",
            "EnrollmentConfigurationPoliciesByDevice": "Enrollment Configuration Policies by Device",
            "AutopilotV1DeploymentStatus": "Autopilot V1 Deployment Status",
            "AutopilotV2DeploymentStatus": "Autopilot V2 Deployment Status",
            "AutopilotV2DeploymentStatusDetailedAppInfo": "Autopilot V2 Deployment Status - App Info",
            "AutopilotV2DeploymentStatusDetailedScriptInfo": "Autopilot V2 Deployment Status - Script Info",
            
            # Endpoint Analytics - Device Performance
            "EADevicePerformance": "Endpoint Analytics Device Performance",
            "EADevicePerformanceV2": "Endpoint Analytics Device Performance V2",
            "EADeviceModelPerformance": "Endpoint Analytics Device Model Performance",
            "EADeviceModelPerformanceV2": "Endpoint Analytics Device Model Performance V2",
            "EADeviceScoresV2": "Endpoint Analytics Device Scores V2",
            "EAModelScoresV2": "Endpoint Analytics Model Scores V2",
            
            # Endpoint Analytics - Application Performance
            "EAAppPerformance": "Endpoint Analytics App Performance",
            "EAOSVersionsPerformance": "Endpoint Analytics OS Versions Performance",
            
            # Endpoint Analytics - Startup Performance
            "EAStartupPerfDevicePerformance": "Endpoint Analytics Startup Performance - Device",
            "EAStartupPerfDevicePerformanceV2": "Endpoint Analytics Startup Performance - Device V2",
            "EAStartupPerfDeviceProcesses": "Endpoint Analytics Startup Performance - Device Processes",
            "EAStartupPerfModelPerformance": "Endpoint Analytics Startup Performance - Model",
            "EAStartupPerfModelPerformanceV2": "Endpoint Analytics Startup Performance - Model V2",
            
            # Endpoint Analytics - Resource Performance
            "EAResourcePerfAggByDevice": "Endpoint Analytics Resource Performance by Device",
            "EAResourcePerfAggByModel": "Endpoint Analytics Resource Performance by Model",
            "EAResourcePerfCpuSpikeProcess": "Endpoint Analytics Resource Performance - CPU Spikes",
            "EAResourcePerfRamSpikeProcess": "Endpoint Analytics Resource Performance - RAM Spikes",
            "ResourcePerformanceAggregateByDevice": "Resource Performance Aggregate by Device",
            "ResourcePerformanceAggregateByModel": "Resource Performance Aggregate by Model",
            
            # Endpoint Analytics - Battery Health
            "BRBatteryByModel": "Battery Report by Model Performance",
            "BRBatteryByOs": "Battery Report by OS Performance",
            "BRDeviceBatteryAgg": "Battery Report Device Performance Aggregate",
            "BREnergyUsage": "Battery Report Energy Usage",
            
            # Endpoint Analytics - Work from Anywhere
            "EAWFADeviceList": "Endpoint Analytics Work from Anywhere - Device List",
            "EAWFAModelPerformance": "Endpoint Analytics Work from Anywhere - Model Performance",
            "EAWFAPerDevicePerformance": "Endpoint Analytics Work from Anywhere - Per Device Performance",
            "WorkFromAnywhereDeviceList": "Work from Anywhere Device List",
            
            # Endpoint Analytics - Anomalies
            "EAAnomalyAsset": "Endpoint Analytics Anomaly Assets",
            "EAAnomalyAssetV2": "Endpoint Analytics Anomaly Assets V2",
            "EAAnomalyDeviceAsset": "Endpoint Analytics Anomaly Device Assets",
            "EAAnomalyDeviceAssetV2": "Endpoint Analytics Anomaly Device Assets V2",
            
            # Scripts & Proactive Remediations
            "DeviceRunStatesByScript": "Device Run States by Script",
            "DeviceRunStatesByProactiveRemediation": "Device Run States by Proactive Remediation",
            "PolicyRunStatesByProactiveRemediation": "Policy Run States by Proactive Remediation",
            
            # Endpoint Privilege Management (EPM)
            "EpmAggregationReportByApplication": "EPM Elevation Report by Application",
            "EpmAggregationReportByApplicationV2": "EPM Elevation Report by Application V2",
            "EpmAggregationReportByPublisher": "EPM Elevation Report by Publisher",
            "EpmAggregationReportByPublisherV2": "EPM Elevation Report by Publisher V2",
            "EpmAggregationReportByUser": "EPM Elevation Report by User",
            "EpmAggregationReportByUserV2": "EPM Elevation Report by User V2",
            "EpmAggregationReportByUserAppByMonth": "EPM Elevation Report by User App by Month",
            "EpmDeniedReport": "EPM Denied Elevations Report",
            "EpmElevationReportByUserAppByDayToReporting": "EPM Elevation Report by User App by Day",
            "EpmElevationReportElevationEvent": "EPM Elevation Report - Elevation Events",
            "EpmInsightsElevationTrend": "EPM Insights Elevation Trend",
            "EpmInsightsMostFrequentElevations": "EPM Insights Most Frequent Elevations",
            "EpmInsightsReport": "EPM Insights Report",
            
            # Device Settings & Configuration
            "DeviceIntentPerSettingStatus": "Device Intent per Setting Status",
            "DeviceInventoryPolicyStatusesV3": "Device Inventory Policy Status V3",
            "DeviceInventoryPolicyStatusesWithPF": "Device Inventory Policy Status with PF",
            "InventoryPolicyDeviceAggregatesV3": "Inventory Policy Device Aggregates V3",
            "InventoryPolicyDeviceAggregatesWithPF": "Inventory Policy Device Aggregates with PF",
            "PerSettingDeviceSummaryByConfigurationPolicy": "Per Setting Device Summary by Configuration Policy",
            "PerSettingDeviceSummaryByConfigurationPolicyForAppControl": "Per Setting Device Summary for App Control",
            "PerSettingDeviceSummaryByConfigurationPolicyForEDR": "Per Setting Device Summary for EDR",
            "PerSettingDeviceSummaryByInventoryPolicy": "Per Setting Device Summary by Inventory Policy",
            "PerSettingDeviceSummaryByInventoryPolicyV3": "Per Setting Device Summary by Inventory Policy V3",
            "PerSettingSummaryByDeviceConfigurationPolicy": "Per Setting Summary by Device Configuration Policy",
            "ADMXSettingsByDeviceByPolicy": "ADMX Settings by Device by Policy",
            
            # Co-management & Cloud Attached Devices
            "ComanagedDeviceWorkloads": "Co-managed Device Workloads",
            "ComanagementEligibilityTenantAttachedDevices": "Co-management Eligibility Tenant Attached Devices",
            
            # Group Policy Analytics
            "GPAnalyticsSettingMigrationReadiness": "Group Policy Analytics Migration Readiness",
            
            # Security Tasks & Monitoring
            "TicketingSecurityTaskAppsList": "Security Task Apps List",
            "OrgDeviceInstallStatus": "Org Device Install Status",
            
            # Users & Remote Assistance
            "Users": "All Users",
            "AllGroupsInMyOrg": "All Groups in My Organization",
            "RemoteAssistanceSessions": "Remote Assistance Sessions",
            "UserScaleTest": "User Scale Test Report"
        }
        
        # Parameter requirements for reports
        self.report_parameter_requirements = {
            # Reports requiring mandatory input
            "AppInvByDevice": {
                "requirement_level": "mandatory",
                "icon": "🔴",
                "parameters": {
                    "deviceId": {"type": "device_selector", "required": True, "description": "Select target device"}
                },
                "description": "Requires specific device selection"
            },
            "DeviceRunStatesByProfiles": {
                "requirement_level": "mandatory", 
                "icon": "🔴",
                "parameters": {
                    "deviceId": {"type": "device_selector", "required": True, "description": "Select target device"}
                },
                "description": "Requires device context"
            },
            "FeatureUpdatePolicyFailuresAggregate": {
                "requirement_level": "mandatory",
                "icon": "🔴", 
                "parameters": {
                    "startDate": {"type": "date", "required": True, "description": "Start date for report"},
                    "endDate": {"type": "date", "required": True, "description": "End date for report"}
                },
                "description": "Requires date range"
            },
            "QualityUpdatePolicyFailuresAggregate": {
                "requirement_level": "mandatory",
                "icon": "🔴",
                "parameters": {
                    "startDate": {"type": "date", "required": True, "description": "Start date for report"},
                    "endDate": {"type": "date", "required": True, "description": "End date for report"}
                },
                "description": "Requires date range"
            },
            
            # Reports with no parameters (changed from optional to none)
            "Devices": {"requirement_level": "none", "icon": "🟢", "description": "No input required"},
            "DevicesWithInventory": {"requirement_level": "none", "icon": "�", "description": "No input required"},
            "PolicyNonCompliance": {"requirement_level": "none", "icon": "�", "description": "No input required"},

            # Reports requiring no input (green)
            "AllAppsList": {"requirement_level": "none", "icon": "🟢", "description": "No input required"},
            "DefenderAgents": {"requirement_level": "none", "icon": "🟢", "description": "No input required"},
            "Users": {"requirement_level": "none", "icon": "🟢", "description": "No input required"},
            "AllGroupsInMyOrg": {"requirement_level": "none", "icon": "🟢", "description": "No input required"}
        }
        
        # Current parameters for active export
        self.current_export_parameters = {}
        
        # Add these predefined templates
        self.export_templates = {
            "Device Overview": ['deviceName', 'userPrincipalName', 'manufacturer', 'model', 'complianceState'],
            "Compliance Report": ['deviceName', 'complianceState', 'lastSyncDateTime', 'userPrincipalName'],
            "Hardware Inventory": ['deviceName', 'manufacturer', 'model', 'serialNumber', 'totalStorageSpace'],
            "User Report": ['userPrincipalName', 'displayName', 'deviceName', 'enrolledDateTime']
        }
        
        # Permission discovery and caching
        self.user_permissions_cache = None
        self.user_access_level = 'unknown'
        self.filtered_available_reports = None
        
        # Permission test endpoints mapping for all 177 reports
        self.permission_test_endpoints = {
            # Security & Antivirus Reports
            "ActiveMalware": "/deviceManagement/detectedApps",
            "DefenderAgents": "/deviceManagement/managedDevices", 
            "Malware": "/deviceManagement/detectedApps",
            "UnhealthyDefenderAgents": "/deviceManagement/managedDevices",
            "FirewallStatus": "/deviceManagement/managedDevices",
            "FirewallUnhealthyStatus": "/deviceManagement/managedDevices",
            
            # Apps & Application Management
            "AllAppsList": "/deviceAppManagement/mobileApps",
            "FilteredAppsList": "/deviceAppManagement/mobileApps", 
            "AppInstallStatusAggregate": "/deviceAppManagement/mobileApps",
            "AppInvAggregate": "/deviceAppManagement/mobileApps",
            "AppInvByDevice": "/deviceAppManagement/mobileApps",
            "AppInvRawData": "/deviceAppManagement/mobileApps",
            "CatalogAppsUpdateList": "/deviceAppManagement/mobileApps",
            "DependentAppsInstallStatus": "/deviceAppManagement/mobileApps",
            "DeviceInstallStatusByApp": "/deviceAppManagement/mobileApps",
            "DevicesByAppInv": "/deviceAppManagement/mobileApps",
            "OrgAppsInstallStatus": "/deviceAppManagement/mobileApps",
            "OrgDeviceInstallStatus": "/deviceAppManagement/mobileApps",
            "UserInstallStatusAggregateByApp": "/deviceAppManagement/mobileApps",
            
            # Mobile Application Management (MAM)
            "MAMAppConfigurationStatus": "/deviceAppManagement/managedAppStatuses",
            "MAMAppConfigurationStatusScopedV2": "/deviceAppManagement/managedAppStatuses",
            "MAMAppConfigurationStatusV2": "/deviceAppManagement/managedAppStatuses",
            "MAMAppProtectionStatus": "/deviceAppManagement/managedAppStatuses",
            "MAMAppProtectionStatusScopedV2": "/deviceAppManagement/managedAppStatuses",
            "MAMAppProtectionStatusV2": "/deviceAppManagement/managedAppStatuses",
            
            # Device Management & Inventory
            "Devices": "/deviceManagement/managedDevices",
            "DevicesWithInventory": "/deviceManagement/managedDevices",
            "AllDeviceCertificates": "/deviceManagement/managedDevices",
            "CertificatesByRAPolicy": "/deviceManagement/managedDevices",
            "TpmAttestationStatus": "/deviceManagement/managedDevices",
            "MEMUpgradeReadinessOrgAsset": "/deviceManagement/managedDevices",
            
            # Device Compliance
            "DeviceCompliance": "/deviceManagement/deviceCompliancePolicyDeviceStateSummary",
            "DeviceComplianceTrend": "/deviceManagement/deviceCompliancePolicyDeviceStateSummary",
            "DeviceNonCompliance": "/deviceManagement/deviceCompliancePolicyDeviceStateSummary",
            "DevicesWithoutCompliancePolicy": "/deviceManagement/deviceCompliancePolicyDeviceStateSummary",
            "DevicesWithoutCompliancePolicyV3": "/deviceManagement/deviceCompliancePolicyDeviceStateSummary",
            "NonCompliantDevicesByCompliancePolicy": "/deviceManagement/deviceCompliancePolicyDeviceStateSummary",
            "NonCompliantDevicesByCompliancePolicyV3": "/deviceManagement/deviceCompliancePolicyDeviceStateSummary",
            "NoncompliantDevicesAndSettings": "/deviceManagement/deviceCompliancePolicyDeviceStateSummary",
            "NoncompliantDevicesAndSettingsV3": "/deviceManagement/deviceCompliancePolicyDeviceStateSummary",
            "NoncompliantDevicesToBeRetired": "/deviceManagement/deviceCompliancePolicyDeviceStateSummary",
            "NonCompliantCompliancePoliciesAggregate": "/deviceManagement/deviceCompliancePolicies",
            "NonCompliantCompliancePoliciesAggregateV3": "/deviceManagement/deviceCompliancePolicies",
            "NonCompliantConfigurationPoliciesAggregateWithPF": "/deviceManagement/deviceConfigurations",
            "NonCompliantConfigurationPoliciesAggregateWithPFV3": "/deviceManagement/deviceConfigurations",
            
            # Configuration Policies
            "ConfigurationPolicyAggregate": "/deviceManagement/deviceConfigurations",
            "ConfigurationPolicyAggregateV3": "/deviceManagement/deviceConfigurations",
            "ConfigurationPolicyDeviceAggregates": "/deviceManagement/deviceConfigurations",
            "ConfigurationPolicyDeviceAggregatesV3": "/deviceManagement/deviceConfigurations",
            "ConfigurationPolicyDeviceAggregatesWithPF": "/deviceManagement/deviceConfigurations",
            "ConfigurationPolicyDeviceAggregatesWithPFV3": "/deviceManagement/deviceConfigurations",
            "DeviceConfigurationPolicyStatuses": "/deviceManagement/deviceConfigurations",
            "DeviceConfigurationPolicyStatusesV3": "/deviceManagement/deviceConfigurations",
            "DeviceConfigurationPolicyStatusesWithPF": "/deviceManagement/deviceConfigurations",
            "DeviceConfigurationPolicyStatusesWithPFV3": "/deviceManagement/deviceConfigurations",
            "Policies": "/deviceManagement/deviceCompliancePolicies",
            
            # Default mapping for remaining reports - use managedDevices as it's commonly accessible
        }
        
        # Add default endpoints for reports not explicitly mapped
        for report_name in self.available_reports.keys():
            if report_name not in self.permission_test_endpoints:
                self.permission_test_endpoints[report_name] = "/deviceManagement/managedDevices"
        
        # Smart Parameter System for Reports requiring additional filters
        self.report_parameters = self.initialize_report_parameters()
        self.learned_parameters = {}  # Cache for dynamically learned parameters
        
        # Direct API Call Configuration - Reports that use GET calls instead of export jobs
        self.direct_api_reports = self.initialize_direct_api_reports()
        
        # Initialize GUI
        self.create_gui()
        
    def create_gui(self):
        """Create the main GUI"""
        self.root = tk.Tk()
        self.root.title("HTMD - Microsoft Intune Reports Export Tool")
        self.root.geometry("1400x900")
        self.root.minsize(1200, 700)
        
        # Create container for pages
        self.container = tk.Frame(self.root)
        self.container.pack(fill='both', expand=True)
        
        # Show login page immediately
        self.show_login_page()
        
        # Bind keyboard shortcuts
        self.root.bind('<Control-e>', lambda e: self.export_report() if hasattr(self, 'export_btn') and self.export_btn['state'] == 'normal' else None)
        self.root.bind('<Control-s>', lambda e: self.generate_filtered_csv() if hasattr(self, 'generate_btn') else None)
        self.root.bind('<Control-l>', lambda e: self.clear_console() if hasattr(self, 'console_text') else None)
        self.root.bind('<F5>', lambda e: self.export_report() if hasattr(self, 'export_btn') and self.export_btn['state'] == 'normal' else None)
        
    def initialize_report_parameters(self):
        """Initialize known parameter requirements for specific reports"""
        from datetime import datetime, timedelta
        
        return {
            # Windows Update Reports
            "FeatureUpdateDeviceState": {
                "required": {
                    "filter": "PolicyId ne null",
                    "startDate": "auto_30_days_ago",
                    "endDate": "auto_today"
                },
                "optional": {"top": 1000}
            },
            
            "QualityUpdateDeviceStatusByPolicy": {
                "required": {
                    "filter": "PolicyId ne null",
                    "startDate": "auto_7_days_ago", 
                    "endDate": "auto_today"
                }
            },
            
            "WindowsUpdatePerPolicyPerDeviceStatus": {
                "required": {
                    "filter": "PolicyId ne null",
                    "startDate": "auto_30_days_ago",
                    "endDate": "auto_today"
                }
            },
            
            # Enrollment & Autopilot Reports
            "DeviceEnrollmentFailures": {
                "required": {
                    "startDate": "auto_30_days_ago",
                    "endDate": "auto_today"
                },
                "optional": {"filter": "FailureCategory ne null"}
            },
            
            "EnrollmentActivity": {
                "required": {
                    "startDate": "auto_30_days_ago",
                    "endDate": "auto_today"
                }
            },
            
            "AutopilotV1DeploymentStatus": {
                "required": {
                    "startDate": "auto_30_days_ago",
                    "endDate": "auto_today"
                }
            },
            
            "AutopilotV2DeploymentStatus": {
                "required": {
                    "startDate": "auto_30_days_ago",
                    "endDate": "auto_today"
                }
            },
            
            # Compliance & Policy Reports
            "DeviceStatusByCompliacePolicyReport": {
                "required": {
                    "filter": "PolicyId ne null"
                },
                "optional": {
                    "startDate": "auto_30_days_ago",
                    "endDate": "auto_today"
                }
            },
            
            "DeviceStatusByCompliancePolicySettingReport": {
                "required": {
                    "filter": "PolicyId ne null AND SettingName ne null"
                }
            },
            
            # Endpoint Analytics Reports
            "EADevicePerformance": {
                "required": {
                    "startDate": "auto_30_days_ago",
                    "endDate": "auto_today"
                }
            },
            
            "EAStartupPerfDevicePerformance": {
                "required": {
                    "startDate": "auto_30_days_ago",
                    "endDate": "auto_today"
                }
            }
        }
    
    def initialize_direct_api_reports(self):
        """Initialize reports that use direct GET API calls instead of export jobs"""
        return {
            "Users": {
                "endpoint": "/users",
                "base_url": "https://graph.microsoft.com/v1.0",
                "required_permission": "User.ReadBasic.All",
                "description": "Retrieves all users in the organization",
                "parameters": {
                    "$top": 999
                }
            },
            
            "AllGroupsInMyOrg": {
                "endpoint": "/groups",
                "base_url": "https://graph.microsoft.com/v1.0",
                "required_permission": "Group.Read.All",
                "description": "Retrieves all groups in the organization",
                "parameters": {
                    "$top": 999
                }
            },
            
            "OrgAppsInstallStatus": {
                "endpoint": "/deviceAppManagement/mobileApps",
                "base_url": "https://graph.microsoft.com/beta",
                "required_permission": "DeviceManagementApps.Read.All",
                "description": "Retrieves organization apps install status",
                "parameters": {
                    "$top": 999,
                    "$filter": "isAssigned eq true"
                }
            },
            
            "OrgDeviceInstallStatus": {
                "endpoint": "/deviceAppManagement/mobileApps",
                "base_url": "https://graph.microsoft.com/beta",
                "required_permission": "DeviceManagementApps.Read.All",
                "description": "Retrieves organization device install status for apps",
                "parameters": {
                    "$top": 999,
                    "$expand": "installSummary,deviceStatuses,userStatuses"
                }
            },
            
            "Devices": {
                "endpoint": "/deviceManagement/managedDevices", 
                "base_url": "https://graph.microsoft.com/beta",
                "required_permission": "DeviceManagementManagedDevices.Read.All",
                "description": "Retrieves all managed devices",
                "parameters": {
                    "$top": 999
                }
            },
            
            "AllAppsList": {
                "endpoint": "/deviceAppManagement/mobileApps",
                "base_url": "https://graph.microsoft.com/beta", 
                "required_permission": "DeviceManagementApps.Read.All",
                "description": "Retrieves all mobile applications",
                "parameters": {
                    "$top": 999
                }
            },
            
            "Policies": {
                "endpoint": "/deviceManagement/deviceCompliancePolicies",
                "base_url": "https://graph.microsoft.com/beta",
                "required_permission": "DeviceManagementConfiguration.Read.All",
                "description": "Retrieves all device compliance policies",
                "parameters": {
                    "$top": 999,
                    "$expand": "deviceStatusOverview,userStatusOverview"
                }
            },
            
            "DevicesByAppInv": {
                "endpoint": "/deviceAppManagement/mobileApps",
                "base_url": "https://graph.microsoft.com/beta",
                "required_permission": "DeviceManagementApps.Read.All",
                "description": "Retrieves devices by app inventory",
                "parameters": {
                    "$top": 999
                }
            },
            
            "AppInvByDevice": {
                "endpoint": "/deviceAppManagement/mobileApps",
                "base_url": "https://graph.microsoft.com/beta",
                "required_permission": "DeviceManagementApps.Read.All",
                "description": "Retrieves app inventory by device",
                "parameters": {
                    "$top": 999
                }
            }
        }
    
    def get_report_parameters(self, report_name):
        """Get complete parameters for a report using smart defaults and collected parameters"""
        try:
            # Start with base parameters
            params = {
                "reportName": report_name,
                "format": "csv",
                "localizationType": "LocalizedValuesAsAdditionalColumn"
            }
            
            # First, apply any collected parameters from the parameter dialog
            if hasattr(self, 'current_export_parameters') and self.current_export_parameters:
                # Convert UI parameters to API parameters
                api_params = self.convert_ui_params_to_api(self.current_export_parameters)
                params.update(api_params)
                self.log_message(f"Applied collected parameters for {report_name}: {list(api_params.keys())}", 'info')
            
            # Check if we have specific configuration for this report
            elif report_name in self.report_parameters:
                config = self.report_parameters[report_name]
                params.update(self.apply_parameter_config(config))
                self.log_message(f"Applied known parameters for {report_name}", 'info')
                return params
            
            # Check if we've learned parameters for this report before
            if report_name in self.learned_parameters:
                learned_params = self.learned_parameters[report_name].copy()
                
                # Check for special instructions
                if "_remove_filter" in learned_params:
                    # Remove filter from base params and from learned params
                    if "filter" in params:
                        del params["filter"]
                    del learned_params["_remove_filter"]
                    self.log_message(f"Removed problematic filter for {report_name}", 'info')
                
                if "_minimal_params" in learned_params:
                    # Use only the absolute minimum parameters
                    params = {
                        "reportName": report_name,
                        "format": "csv",
                        "localizationType": "LocalizedValuesAsAdditionalColumn"
                    }
                    del learned_params["_minimal_params"]
                    self.log_message(f"Using minimal parameters for {report_name}", 'info')
                
                params.update(learned_params)
                self.log_message(f"Applied learned parameters for {report_name}", 'info')
                return params
            
            # Apply smart pattern-based defaults
            smart_params = self.apply_smart_defaults(report_name)
            if smart_params:
                params.update(smart_params)
                self.log_message(f"Applied smart defaults for {report_name}", 'info')
            
            return params
            
        except Exception as e:
            self.log_message(f"Error building parameters for {report_name}: {str(e)}", 'warning')
            # Return basic parameters as fallback
            return {
                "reportName": report_name,
                "format": "csv",
                "localizationType": "LocalizedValuesAsAdditionalColumn"
            }
    
    def apply_parameter_config(self, config):
        """Apply parameter configuration with auto-calculated dates"""
        from datetime import datetime, timedelta
        
        params = {}
        
        # Apply required parameters
        if "required" in config:
            for key, value in config["required"].items():
                if value == "auto_30_days_ago":
                    params[key] = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
                elif value == "auto_7_days_ago":
                    params[key] = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
                elif value == "auto_today":
                    params[key] = datetime.now().strftime("%Y-%m-%d")
                else:
                    params[key] = value
        
        # Apply optional parameters  
        if "optional" in config:
            for key, value in config["optional"].items():
                params[key] = value
                
        return params
    
    def apply_smart_defaults(self, report_name):
        """Set default values based on report type"""
        from datetime import datetime, timedelta
        
        params = {}
        report_lower = report_name.lower()
        
        # Date-based reports (updates, enrollment, etc.)
        if any(word in report_lower for word in ["update", "feature", "quality", "enrollment", "autopilot"]):
            params.update({
                "startDate": (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"),
                "endDate": datetime.now().strftime("%Y-%m-%d")
            })
            
            # Update reports often need policy filters (but not app inventory)
            if "update" in report_lower and "appinv" not in report_lower:
                params["filter"] = "PolicyId ne null"
        
        # Performance and analytics reports
        elif any(word in report_lower for word in ["performance", "analytics", "ea"]):
            params.update({
                "startDate": (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"),
                "endDate": datetime.now().strftime("%Y-%m-%d")
            })
        
        # Policy and compliance reports (but not app inventory)
        elif any(word in report_lower for word in ["policy", "compliance", "setting"]) and "appinv" not in report_lower:
            params["filter"] = "PolicyId ne null"
        
        # Device status reports
        elif "device" in report_lower and "status" in report_lower:
            params["top"] = 1000
            
        return params if params else None
    
    def learn_from_error(self, report_name, error_response):
        """Learn parameter requirements from API error responses"""
        try:
            error_message = error_response.get('error', {}).get('message', '').lower()
            
            # Common parameter patterns from error messages
            learned_params = {}
            
            # Check for property errors first (like PolicyId not found)
            if "could not find a property named" in error_message:
                self.log_message(f"Property error detected for {report_name}. This report doesn't support the attempted filter.", 'warning')
                
                # For reports like AppInvByDevice, we need minimal parameters
                if "appinv" in report_name.lower():
                    # App inventory reports typically just need basic parameters
                    learned_params = {
                        "_remove_filter": True,
                        "_minimal_params": True
                    }
                    self.log_message(f"Setting minimal parameters for app inventory report {report_name}", 'info')
                else:
                    # For other property errors, try with dates but no filters
                    from datetime import datetime, timedelta
                    learned_params = {
                        "startDate": (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"),
                        "endDate": datetime.now().strftime("%Y-%m-%d"),
                        "_remove_filter": True
                    }
                
            # Check for missing filters (if no property error)
            elif "filter" in error_message and "required" in error_message:
                # Don't add PolicyId filter for app inventory or install status reports
                if any(keyword in report_name.lower() for keyword in ["appinv", "installstatus", "orgdeviceinstallstatus", "deviceinstallstatus"]):
                    # Try with date filters for app and install status reports
                    from datetime import datetime, timedelta
                    learned_params = {
                        "startDate": (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"),
                        "endDate": datetime.now().strftime("%Y-%m-%d")
                    }
                else:
                    learned_params["filter"] = "PolicyId ne null"
                
            # Check for date/time errors
            elif "date" in error_message or "time" in error_message:
                from datetime import datetime, timedelta
                learned_params.update({
                    "startDate": (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"),
                    "endDate": datetime.now().strftime("%Y-%m-%d")
                })
            
            if learned_params:
                # Cache the learned parameters
                self.learned_parameters[report_name] = learned_params
                self.log_message(f"Learned new parameters for {report_name}: {list(learned_params.keys())}", 'info')
                return learned_params
                
        except Exception as e:
            self.log_message(f"Error learning from API response: {str(e)}", 'debug')
            
        return None

    def show_login_page(self):
        """Show login page"""
        # Clear container
        for widget in self.container.winfo_children():
            widget.destroy()
        
        # Login page frame
        login_frame = tk.Frame(self.container, bg='#f5f5f5')
        login_frame.pack(fill='both', expand=True)
        
        # Center content
        center_frame = tk.Frame(login_frame, bg='#f5f5f5')
        center_frame.place(relx=0.5, rely=0.5, anchor='center')
        
        # Header
        header_frame = tk.Frame(center_frame, bg='#0078d4', height=100, width=1200)
        header_frame.pack(pady=(0, 30))
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text="HTMD - Microsoft Intune Reports Export Tool", 
                              font=('Segoe UI', 22, 'bold'), 
                              bg='#0078d4', fg='white')
        title_label.pack(expand=True)
        
        # Login card
        card_frame = tk.Frame(center_frame, bg='white', relief='solid', bd=1)
        card_frame.pack(padx=40, pady=30, ipadx=30, ipady=30)
        
        welcome_label = tk.Label(card_frame, text="Corporate Login", 
                                font=('Segoe UI', 18, 'bold'), bg='white', fg='#323130')
        welcome_label.pack(pady=(0, 20))
        
        info_text = """Sign in with your corporate Microsoft account to access:

✅ Dynamic Report Export (All Available Data)
✅ Post-Export Column Selection
✅ CSV Column Filtering & Customization
✅ Power BI Integration
✅ OData Feed Access
✅ No Pre-defined Parameter Limitations
✅ Works with Any Tenant Configuration

Export first, choose columns later!"""
        
        info_label = tk.Label(card_frame, text=info_text, justify='left', 
                             font=('Segoe UI', 11), bg='white', fg='#605e5c')
        info_label.pack(pady=15)
        
        # README toggle switch
        readme_frame = tk.Frame(card_frame, bg='white')
        readme_frame.pack(pady=10)
        
        readme_label = tk.Label(readme_frame, text="README", 
                               font=('Segoe UI', 10, 'bold'), bg='white', fg='#323130')
        readme_label.pack(side='left', padx=(0, 10))
        
        # Custom toggle switch
        self.readme_var = tk.BooleanVar()
        self.create_toggle_switch(readme_frame)
        
        # Feedback button
        feedback_frame = tk.Frame(card_frame, bg='white')
        feedback_frame.pack(pady=(10, 0))
        
        feedback_btn = tk.Button(feedback_frame, text="📝 Provide Feedback", 
                                command=self.open_feedback_form,
                                font=('Segoe UI', 10, 'bold'),
                                bg='#107c10', fg='white',
                                padx=15, pady=8, cursor='hand2',
                                relief='flat', bd=0)
        feedback_btn.pack()
        
        # Login button
        self.login_btn = tk.Button(card_frame, text="Sign in with Corporate Account", 
                                  command=self.login,
                                  font=('Segoe UI', 12, 'bold'),
                                  bg='#0078d4', fg='white',
                                  padx=20, pady=12, cursor='hand2',
                                  relief='flat', bd=0)
        self.login_btn.pack(pady=20)
        
        # Status label
        self.login_status = tk.Label(card_frame, text="Ready to authenticate", 
                                    font=('Segoe UI', 10), bg='white', fg='#605e5c')
        self.login_status.pack()
        
    def show_reports_page(self):
        """Show reports page after login"""
        # Clear container
        for widget in self.container.winfo_children():
            widget.destroy()
        
        # Discover user permissions and filter available reports (like v1.3)
        self.refresh_available_reports()
        
        # Reports page frame
        reports_frame = tk.Frame(self.container)
        reports_frame.pack(fill='both', expand=True)
        
        # Header
        header_frame = tk.Frame(reports_frame, bg='#0078d4', height=80)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        header_content = tk.Frame(header_frame, bg='#0078d4')
        header_content.pack(fill='x', padx=20, pady=20)
        
        title_label = tk.Label(header_content, text="HTMD - Microsoft Intune Reports Export Tool", 
                              font=('Segoe UI', 16, 'bold'), 
                              bg='#0078d4', fg='white')
        title_label.pack(side='left')
        
        # User info and logout
        user_frame = tk.Frame(header_content, bg='#0078d4')
        user_frame.pack(side='right')
        
        user_name = self.user_info.get('displayName', 'User') if self.user_info else 'User'
        user_email = self.user_info.get('mail', '') if self.user_info else ''
        user_label = tk.Label(user_frame, text=f"User: {user_name} ({user_email})", 
                             font=('Segoe UI', 10), bg='#0078d4', fg='white')
        user_label.pack(side='left', padx=(0, 15))
        
        # Re-authenticate button for token refresh
        reauth_btn = tk.Button(user_frame, text="Refresh Token", command=self.manual_token_refresh,
                              font=('Segoe UI', 9), bg='#106ebe', fg='white',
                              padx=10, pady=4, relief='flat')
        reauth_btn.pack(side='right', padx=(0, 10))
        
        # Feedback button
        feedback_btn = tk.Button(user_frame, text="📝 Feedback", command=self.open_feedback_form,
                                font=('Segoe UI', 9), bg='#107c10', fg='white',
                                padx=10, pady=4, relief='flat')
        feedback_btn.pack(side='right', padx=(0, 10))
        
        logout_btn = tk.Button(user_frame, text="Logout", command=self.logout,
                              font=('Segoe UI', 10), bg='#d13438', fg='white',
                              padx=12, pady=6, relief='flat')
        logout_btn.pack(side='right', padx=(0, 10))
        
        # Create notebook for tabbed interface
        self.notebook = ttk.Notebook(reports_frame)
        self.notebook.pack(fill='both', expand=True, padx=15, pady=15)
        
        # Tab 1: Export
        self.export_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.export_tab, text="1. Export Report")
        
        # Tab 2: Column Selection (initially hidden)
        self.columns_tab = ttk.Frame(self.notebook)
        
        self.create_export_tab()
        
    def create_export_tab(self):
        """Create the export tab interface"""
        # Main paned window
        main_paned = ttk.PanedWindow(self.export_tab, orient='horizontal')
        main_paned.pack(fill='both', expand=True, padx=15, pady=15)
        
        # Left panel - Report Selection
        left_frame = ttk.LabelFrame(main_paned, text="Report Selection & Export", padding=15)
        main_paned.add(left_frame, weight=1)
        
        # Report selection with integrated search
        report_frame = tk.Frame(left_frame)
        report_frame.pack(fill='x', pady=(0, 20))
        
        tk.Label(report_frame, text="Select Report:", font=('Segoe UI', 12, 'bold')).pack(anchor='w')
        
        # Report dropdown with integrated search (editable combobox)
        self.selected_report = tk.StringVar()
        
        # Use clean report names without visual indicators
        available_reports_to_use = getattr(self, 'filtered_available_reports', self.available_reports)
        
        # Use clean report names directly
        report_values = sorted(list(available_reports_to_use.keys()))
        
        # No display mapping needed since we're using actual report names
        self.report_display_mapping = {name: name for name in report_values}
        
        self.sorted_reports = report_values
        self.filtered_reports = report_values.copy()  # Track filtered results
        self.last_search_text = ""  # Track last search to avoid unnecessary updates
        
        self.report_combo = ttk.Combobox(report_frame, textvariable=self.selected_report,
                                        values=report_values,
                                        state='normal', font=('Segoe UI', 10), width=50)
        self.report_combo.pack(fill='x', pady=5)
        
        # Bind events for search functionality
        self.report_combo.bind('<KeyRelease>', self.on_search_type)
        self.report_combo.bind('<<ComboboxSelected>>', self.on_report_selected)
        self.report_combo.bind('<FocusIn>', self.on_dropdown_focus)
        self.report_combo.bind('<Button-1>', self.on_dropdown_click)
        
        # Add specific key bindings for better search experience
        self.report_combo.bind('<Down>', self.on_arrow_down)
        self.report_combo.bind('<Up>', self.on_arrow_up)
        
        # Report description (like v1.3)
        total_reports = len(self.available_reports)
        accessible_count = len(report_values)
        if accessible_count == total_reports:
            desc_text = "Select a report to export all available data"
        else:
            desc_text = f"Select from {accessible_count} reports available to you"
            
        self.report_desc = tk.Label(report_frame, text=desc_text, 
                                   font=('Segoe UI', 10), fg='#605e5c', wraplength=400, justify='left')
        self.report_desc.pack(anchor='w', pady=(5, 15))
        
        # Info box
        info_frame = tk.Frame(left_frame, bg='#e7f3ff', relief='solid', bd=1)
        info_frame.pack(fill='x', pady=(0, 20), padx=10, ipady=15)
        
        info_text = """How it works:
        
1. Select a report from the dropdown
2. Click 'Export Full Report' to get ALL available data
3. After export, you'll see a 'Column Selection' tab
4. Choose which columns to keep in your final CSV
5. Generate customized CSV with only selected columns
6. Generate customized Power BI report with exported CSV as data source
7. Access OData feed for direct data integration with any OData compatible tool"""
        
        tk.Label(info_frame, text=info_text, justify='left', 
                font=('Segoe UI', 10), bg='#e7f3ff', fg='#0078d4').pack(padx=15)
        
        # Export button frame
        export_frame = tk.Frame(left_frame)
        export_frame.pack(fill='x')
        
        # Single export button
        self.export_btn = tk.Button(export_frame, text="📥 Export Report", 
                                   command=self.export_report,
                                   font=('Segoe UI', 11, 'bold'),
                                   bg='#0078d4', fg='white', padx=15, pady=8, relief='flat',
                                   state='disabled')
        self.export_btn.pack(pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(export_frame, mode='indeterminate')
        self.progress.pack(fill='x', pady=5)
        
        self.progress_label = tk.Label(export_frame, text="Ready", 
                                      font=('Segoe UI', 10), fg='#605e5c')
        self.progress_label.pack()
        
        # Right panel - Status Console  
        right_frame = ttk.LabelFrame(main_paned, text="Export Console & API Logs", padding=15)
        main_paned.add(right_frame, weight=2)
        
        # Console controls
        console_controls = tk.Frame(right_frame)
        console_controls.pack(fill='x', pady=(0, 10))
        
        ttk.Button(console_controls, text="Clear", command=self.clear_console).pack(side='left', padx=2)
        ttk.Button(console_controls, text="Save Log", command=self.save_log).pack(side='left', padx=2)
        
        self.auto_scroll = tk.BooleanVar(value=True)
        ttk.Checkbutton(console_controls, text="Auto-scroll", variable=self.auto_scroll).pack(side='right')
        
        # Console text
        console_frame = tk.Frame(right_frame)
        console_frame.pack(fill='both', expand=True)
        
        self.console_text = tk.Text(console_frame, wrap='word', font=('Consolas', 9),
                                   bg='#1e1e1e', fg='#ffffff', insertbackground='white')
        console_v_scroll = ttk.Scrollbar(console_frame, orient='vertical', command=self.console_text.yview)
        console_h_scroll = ttk.Scrollbar(console_frame, orient='horizontal', command=self.console_text.xview)
        self.console_text.configure(yscrollcommand=console_v_scroll.set, xscrollcommand=console_h_scroll.set)
        
        # Configure tags
        self.console_text.tag_configure('success', foreground='#4EC9B0')
        self.console_text.tag_configure('error', foreground='#F44747') 
        self.console_text.tag_configure('warning', foreground='#FFCC02')
        self.console_text.tag_configure('info', foreground='#9CDCFE')
        self.console_text.tag_configure('api', foreground='#C586C0')
        self.console_text.tag_configure('debug', foreground='#808080')
        
        self.console_text.grid(row=0, column=0, sticky='nsew')
        console_v_scroll.grid(row=0, column=1, sticky='ns')
        console_h_scroll.grid(row=1, column=0, sticky='ew')
        
        console_frame.grid_rowconfigure(0, weight=1)
        console_frame.grid_columnconfigure(0, weight=1)
        
        # Initialize
        self.log_message("=== Microsoft Intune Reports Export Tool v1.0 ===", 'info')
        self.log_message(f"Loaded {len(self.available_reports)} available reports", 'info')
        self.log_message("Dynamic approach: Export all data, then select columns", 'success')
        
    def create_columns_tab(self):
        """Create the column selection tab after export"""
        # Clear existing content
        for widget in self.columns_tab.winfo_children():
            widget.destroy()
            
        # Add tab if not already added
        try:
            self.notebook.add(self.columns_tab, text="2. Select Columns & Export")
        except:
            pass  # Tab already exists
        
        # Main frame with scrolling
        main_frame = tk.Frame(self.columns_tab)
        main_frame.pack(fill='both', expand=True, padx=15, pady=15)
        
        # Top section - Fixed (non-scrolling)
        top_frame = tk.Frame(main_frame)
        top_frame.pack(fill='x', pady=(0, 10))
        
        # Header with instructions - improved styling
        header_frame = tk.Frame(top_frame, bg='#f3f9ff', relief='solid', bd=1)
        header_frame.pack(fill='x', pady=(0, 15), ipady=10)
        
        instruction_text = f"✅ Report '{self.selected_report.get()}' exported successfully"
        stats_text = f"📊 {len(self.current_columns)} columns • {len(self.current_export_data)} rows"
        
        tk.Label(header_frame, text=instruction_text, 
                font=('Segoe UI', 13, 'bold'), bg='#f3f9ff', fg='#0078d4').pack(padx=20, pady=(5, 2))
        tk.Label(header_frame, text=stats_text, 
                font=('Segoe UI', 10), bg='#f3f9ff', fg='#605e5c').pack(padx=20, pady=(0, 5))
        
        # Export Actions Section - prominently placed
        export_section = tk.Frame(top_frame, bg='#fafafa', relief='solid', bd=1)
        export_section.pack(fill='x', pady=(0, 15), padx=5, ipady=10)
        
        # Export section header
        export_header = tk.Label(export_section, text="📤 Export Actions", 
                                font=('Segoe UI', 11, 'bold'), bg='#fafafa', fg='#323130')
        export_header.pack(pady=(5, 8))
        
        # Export buttons container
        export_frame = tk.Frame(export_section, bg='#fafafa')
        export_frame.pack(pady=(0, 5))
        
        # Buttons container - centered
        buttons_container = tk.Frame(export_frame, bg='#fafafa')
        buttons_container.pack()
        
        # CSV Export button - compact and professional
        self.generate_btn = tk.Button(buttons_container, text="📥 Export CSV", 
                                     command=self.generate_filtered_csv,
                                     font=('Segoe UI', 10, 'bold'),
                                     bg='#107c10', fg='white', padx=12, pady=5, 
                                     relief='flat', cursor='hand2', width=16,
                                     activebackground='#0d5c0d', activeforeground='white',
                                     highlightthickness=0, borderwidth=0)
        self.generate_btn.pack(side='left', padx=(0, 8))
        
        # View Report button - new feature for data preview
        self.view_btn = tk.Button(buttons_container, text="� View Report", 
                                 command=self.view_report_data,
                                 font=('Segoe UI', 10, 'bold'),
                                 bg='#8764b8', fg='white', padx=12, pady=5, 
                                 relief='flat', cursor='hand2', width=16,
                                 activebackground='#6b4d93', activeforeground='white',
                                 highlightthickness=0, borderwidth=0,
                                 state='disabled')
        self.view_btn.pack(side='left', padx=(0, 8))
        
        # PowerBI button - simplified approach
        self.powerbi_btn = tk.Button(buttons_container, text="📊 Open Power BI", 
                                    command=self.open_powerbi_simple,
                                    font=('Segoe UI', 10, 'bold'),
                                    bg='#FFCD00', fg='#323130', padx=12, pady=5, 
                                    relief='flat', cursor='hand2', width=16,
                                    activebackground='#FFB900', activeforeground='#323130',
                                    highlightthickness=0, borderwidth=0)
        self.powerbi_btn.pack(side='left', padx=(0, 8))
        
        # OData Feed button - for direct data feed access
        self.odata_btn = tk.Button(buttons_container, text="🔗 Get OData Feed", 
                                  command=self.get_odata_feed,
                                  font=('Segoe UI', 10, 'bold'),
                                  bg='#0078d4', fg='white', padx=12, pady=5, 
                                  relief='flat', cursor='hand2', width=16,
                                  activebackground='#106ebe', activeforeground='white',
                                  highlightthickness=0, borderwidth=0)
        self.odata_btn.pack(side='left')
        
        # Status label - below buttons
        self.export_status = tk.Label(export_section, text="Ready to export", 
                                     font=('Segoe UI', 9), fg='#605e5c', bg='#fafafa')
        self.export_status.pack(pady=(5, 0))
        
        # Column Controls Section
        controls_section = tk.Frame(top_frame, bg='#f9f9f9', relief='solid', bd=1)
        controls_section.pack(fill='x', pady=(0, 15), padx=5, ipady=8)
        
        # Controls header
        controls_header = tk.Label(controls_section, text="🔧 Column Controls", 
                                  font=('Segoe UI', 10, 'bold'), bg='#f9f9f9', fg='#323130')
        controls_header.pack(pady=(5, 8))
        
        # Control bar
        control_bar = tk.Frame(controls_section, bg='#f9f9f9')
        control_bar.pack(fill='x', padx=10)
        
        # Left side - Search and counter
        left_controls = tk.Frame(control_bar, bg='#f9f9f9')
        left_controls.pack(side='left', fill='x', expand=True)
        
        # Search
        search_frame = tk.Frame(left_controls, bg='#f9f9f9')
        search_frame.pack(side='left')
        
        tk.Label(search_frame, text="🔍 Search:", font=('Segoe UI', 9), bg='#f9f9f9').pack(side='left', padx=(0, 5))
        self.column_search = tk.StringVar()
        self.column_search.trace('w', self.filter_columns)
        search_entry = ttk.Entry(search_frame, textvariable=self.column_search, width=20)
        search_entry.pack(side='left')
        
        # Selection counter
        self.selection_counter = tk.Label(left_controls, text=f"Selected: {len(self.current_columns)} columns", 
                                         font=('Segoe UI', 9, 'bold'), fg='#107c10', bg='#f9f9f9')
        self.selection_counter.pack(side='left', padx=15)
        
        # Right side - Control buttons
        right_controls = tk.Frame(control_bar, bg='#f9f9f9')
        right_controls.pack(side='right')
        
        ttk.Button(right_controls, text="Select All", command=self.select_all_columns, width=10).pack(side='left', padx=1)
        ttk.Button(right_controls, text="Clear All", command=self.clear_all_columns, width=10).pack(side='left', padx=1)
        ttk.Button(right_controls, text="Common", command=self.select_common_columns, width=10).pack(side='left', padx=1)
        
        # Column selection area - enhanced styling with better visibility
        columns_container = ttk.LabelFrame(main_frame, text="📋 Select Columns for Export", padding=10)
        columns_container.pack(fill='both', expand=True, padx=5, pady=(10, 0))
        
        # Create container for canvas and scrollbar
        scroll_container = tk.Frame(columns_container, height=400)
        scroll_container.pack(fill='both', expand=True)
        scroll_container.pack_propagate(False)  # Maintain minimum height
        
        # Create scrollable frame with better styling - increased height
        canvas = tk.Canvas(scroll_container, bg='#fcfcfc', highlightthickness=0, height=400)
        scrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#fcfcfc')
        
        # Configure scrolling
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas_frame = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Bind canvas width to scrollable frame width
        def configure_canvas_width(event):
            canvas.itemconfig(canvas_frame, width=event.width)
        canvas.bind('<Configure>', configure_canvas_width)
        
        # Create checkboxes with better spacing and visibility
        self.column_vars = {}
        self.column_checkboxes = {}  # Store checkbox references for filtering
        columns_per_row = 2  # Use 2 columns for better readability
        
        # Debug info
        print(f"Creating {len(self.current_columns)} column checkboxes...")
        if len(self.current_columns) == 0:
            tk.Label(scrollable_frame, text="⚠️ No columns found in the exported data", 
                    font=('Segoe UI', 12), fg='red', bg='#fcfcfc').pack(pady=20)
        
        for i, column in enumerate(self.current_columns):
            var = tk.BooleanVar(value=True)
            self.column_vars[column] = var
            
            row = i // columns_per_row
            col = i % columns_per_row
            
            cb = tk.Checkbutton(scrollable_frame, text=column, variable=var,
                               font=('Segoe UI', 10), anchor='w', bg='#ffffff',
                               command=self.update_selection_count,
                               wraplength=300, justify='left',
                               activebackground='#e7f3ff', selectcolor='#ffffff',
                               relief='solid', bd=1, padx=8, pady=4,
                               highlightbackground='#cccccc')
            cb.grid(row=row, column=col, sticky='ew', padx=8, pady=3, ipadx=5, ipady=2)
            
            # Store reference for filtering
            self.column_checkboxes[column] = cb
            print(f"Created checkbox for column: {column}")  # Debug
        
        # Configure grid columns to expand evenly
        for col in range(columns_per_row):
            scrollable_frame.columnconfigure(col, weight=1, minsize=250)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Force canvas update
        self.root.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        
        # Auto-switch to columns tab
        self.notebook.select(self.columns_tab)
        
        # Update initial count
        self.update_selection_count()
    
    def update_selection_count(self):
        """Update the selection counter"""
        if hasattr(self, 'column_vars') and hasattr(self, 'selection_counter'):
            selected = sum(1 for var in self.column_vars.values() if var.get())
            total = len(self.column_vars)
            self.selection_counter.config(text=f"✅ Selected: {selected} of {total} columns")
            
            # Update export status and button states
            if hasattr(self, 'export_status'):
                if selected == 0:
                    self.export_status.config(text="⚠️ No columns selected - Please select at least one column", fg='#d13438')
                    self.generate_btn.config(state='disabled')
                    if hasattr(self, 'view_btn'):
                        self.view_btn.config(state='disabled')
                else:
                    self.export_status.config(text=f"📊 Ready to export {selected} columns × {len(self.current_export_data)} rows", fg='#107c10')
                    self.generate_btn.config(state='normal')
                    if hasattr(self, 'view_btn'):
                        self.view_btn.config(state='normal')
                    
                # OData button is always enabled regardless of column selection (it provides feed URL)
                if hasattr(self, 'odata_btn'):
                    self.odata_btn.config(state='normal')
    
    def select_all_columns(self):
        """Select all columns"""
        for var in self.column_vars.values():
            var.set(True)
        self.update_selection_count()
        self.log_message("All columns selected", 'info')
    
    def clear_all_columns(self):
        """Clear all column selections"""
        for var in self.column_vars.values():
            var.set(False)
        self.update_selection_count()
        self.log_message("All columns cleared", 'info')
    
    def select_common_columns(self):
        """Select commonly used columns"""
        # Clear all first
        self.clear_all_columns()
        
        # Common column patterns to look for
        common_patterns = [
            'name', 'id', 'user', 'device', 'status', 'state', 'date', 'time',
            'compliance', 'version', 'manufacturer', 'model', 'serial'
        ]
        
        selected_count = 0
        for column, var in self.column_vars.items():
            column_lower = column.lower()
            if any(pattern in column_lower for pattern in common_patterns):
                var.set(True)
                selected_count += 1
        
        self.update_selection_count()
        self.log_message(f"Selected {selected_count} common columns", 'info')
    
    def filter_columns(self, *args):
        """Filter columns based on search term"""
        if not hasattr(self, 'column_checkboxes') or not hasattr(self, 'column_search'):
            return
            
        search_term = self.column_search.get().lower()
        
        # Filter checkboxes based on search term
        visible_count = 0
        for column, checkbox in self.column_checkboxes.items():
            if search_term == "" or search_term in column.lower():
                checkbox.grid()  # Show checkbox
                visible_count += 1
            else:
                checkbox.grid_remove()  # Hide checkbox
        
        # Update the canvas scroll region after filtering
        if hasattr(self, 'root'):
            self.root.after_idle(lambda: self.update_canvas_scroll())
    
    def update_canvas_scroll(self):
        """Update canvas scroll region"""
        try:
            # Find the canvas (this is a bit hacky but works)
            for widget in self.columns_tab.winfo_children():
                if isinstance(widget, tk.Frame):
                    for subwidget in widget.winfo_children():
                        if isinstance(subwidget, ttk.LabelFrame):
                            for container in subwidget.winfo_children():
                                if isinstance(container, tk.Frame):
                                    for canvas_widget in container.winfo_children():
                                        if isinstance(canvas_widget, tk.Canvas):
                                            canvas_widget.configure(scrollregion=canvas_widget.bbox("all"))
                                            return
        except:
            pass
    
    def view_report_data(self):
        """Open a new window to view the report data in a table format"""
        try:
            # Check if we have exported data
            if not self.current_export_data:
                messagebox.showwarning("No Data", 
                                     "No report data available to view.\n\n" +
                                     "Please export a report first by clicking 'Export CSV'.")
                return
            
            # Get the selected report name
            selected_report = self.selected_report.get()
            if not selected_report:
                selected_report = "Report Data"
            
            # Open the report viewer
            viewer = ReportViewer(
                parent=self,
                report_name=selected_report,
                data=self.current_export_data,
                columns=self.current_columns
            )
            
            # Update status
            self.log_message(f"📊 Opened report viewer for {selected_report}", 'success')
            
        except Exception as e:
            error_msg = f"Failed to open report viewer: {str(e)}"
            messagebox.showerror("Viewer Error", error_msg)
            self.log_message(error_msg, 'error')
    
    def generate_filtered_csv(self):
        """Generate CSV with only selected columns and show immediate feedback"""
        # Get selected columns
        selected_columns = []
        for column, var in self.column_vars.items():
            if var.get():
                selected_columns.append(column)
        
        if not selected_columns:
            messagebox.showwarning("No Columns Selected", "Please select at least one column")
            return
        
        # Choose save location with better default name
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"{self.selected_report.get()}_export_{timestamp}.csv"
        
        filepath = filedialog.asksaveasfilename(
            title="Save Filtered CSV Export",
            initialfile=default_filename,
            defaultextension=".csv",
            filetypes=[
                ("CSV files", "*.csv"), 
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        
        if not filepath:
            return
        
        try:
            # Update UI to show progress
            self.export_status.config(text="📝 Creating CSV file...", fg='#0078d4')
            self.generate_btn.config(state='disabled', text="📝 Creating...")
            self.root.update()
            
            # Create filtered CSV
            if filepath.lower().endswith('.xlsx'):
                # Export as Excel
                import pandas as pd
                df_data = []
                for row in self.current_export_data:
                    filtered_row = {col: row.get(col, '') for col in selected_columns}
                    df_data.append(filtered_row)
                
                df = pd.DataFrame(df_data)
                df.to_excel(filepath, index=False)
                file_type = "Excel"
            else:
                # Export as CSV
                with open(filepath, 'w', newline='', encoding='utf-8-sig') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=selected_columns)
                    writer.writeheader()
                    
                    # Write filtered data
                    for row in self.current_export_data:
                        filtered_row = {col: row.get(col, '') for col in selected_columns}
                        writer.writerow(filtered_row)
                file_type = "CSV"
            
            # Create metadata
            metadata = {
                'original_report': self.selected_report.get(),
                'export_time': datetime.now().isoformat(),
                'user': self.user_info.get('displayName', 'Unknown') if self.user_info else 'Unknown',
                'total_columns_available': len(self.current_columns),
                'selected_columns_count': len(selected_columns),
                'selected_columns': selected_columns,
                'excluded_columns': [col for col in self.current_columns if col not in selected_columns],
                'total_rows': len(self.current_export_data),
                'file_format': file_type.lower(),
                'file_path': filepath
            }
            
            metadata_file = filepath.replace(filepath.split('.')[-1], 'json')
            with open(metadata_file, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, indent=2)
            
            # Update UI with success
            file_size = os.path.getsize(filepath)
            file_size_mb = file_size / (1024 * 1024)
            
            self.export_status.config(text=f"✅ {file_type} file created successfully!", fg='#107c10')
            # Update status with file details instead of separate file_info label
            detailed_status = f"✅ {file_type} created: {os.path.basename(filepath)} ({file_size_mb:.2f} MB)"
            self.export_status.config(text=detailed_status, fg='#107c10')
            
            self.log_message(f"Filtered {file_type} generated successfully!", 'success')
            self.log_message(f"File: {filepath}", 'success')
            self.log_message(f"Selected {len(selected_columns)} of {len(self.current_columns)} columns", 'info')
            self.log_message(f"File size: {file_size_mb:.2f} MB", 'info')
            
            # Show success dialog without automatic folder opening
            messagebox.showinfo(
                "Export Complete!", 
                f"✅ {file_type} file created successfully!\n\n"
                f"📁 File: {os.path.basename(filepath)}\n"
                f"📊 Data: {len(selected_columns)} columns × {len(self.current_export_data)} rows\n"
                f"💾 Size: {file_size_mb:.2f} MB\n\n"
                f"📂 Saved to: {os.path.dirname(filepath)}")
            
            # No automatic folder opening - let user access file manually
                    
        except Exception as e:
            self.export_status.config(text="❌ Export failed", fg='#d13438')
            self.log_message(f"Failed to generate CSV: {str(e)}", 'error')
            messagebox.showerror("Export Failed", f"Failed to generate filtered CSV:\n\n{str(e)}")
        finally:
            # Restore button
            self.generate_btn.config(state='normal', text="📥 Export CSV")
    
    def on_search_type(self, event):
        """Handle real-time search as user types in the dropdown"""
        # Skip certain keys that shouldn't trigger search
        if event.keysym in ['Down', 'Up', 'Return', 'Tab', 'Escape', 'Left', 'Right']:
            return
            
        search_text = self.selected_report.get().lower()
        
        # Only update if search text has actually changed
        if search_text == self.last_search_text:
            return
            
        self.last_search_text = search_text
        
        # Use clean report names for search
        available_reports = getattr(self, 'filtered_available_reports', None)
        if available_reports is None:
            available_reports = getattr(self, 'available_reports', {})
        if not available_reports:
            available_reports = {}
        
        if not search_text:
            # If search is empty, show all available reports (clean names)
            self.filtered_reports = sorted(list(available_reports.keys()))
            self.report_desc.config(text="Select a report to export all available data")
        else:
            # Filter reports that contain the search text
            self.filtered_reports = []
            for report_name in sorted(available_reports.keys()):
                report_description = available_reports[report_name].lower()
                report_key_lower = report_name.lower()
                
                if (search_text in report_key_lower or 
                    search_text in report_description):
                    self.filtered_reports.append(report_name)
            
            # Update status
            matching_count = len(self.filtered_reports)
            if matching_count > 0:
                self.report_desc.config(text=f"Found {matching_count} reports matching '{search_text}' - Click dropdown arrow to see results")
            else:
                self.report_desc.config(text=f"No reports found matching '{search_text}'")
        
        # Update the combobox values (but don't force it open)
        self.report_combo['values'] = self.filtered_reports
    
    def on_dropdown_click(self, event):
        """Handle dropdown arrow click - show current filtered results"""
        # When user clicks dropdown arrow, show current filtered results
        # This preserves the current filter if user was searching
        if not self.filtered_reports:
            self.filtered_reports = self.sorted_reports.copy()
        self.report_combo['values'] = self.filtered_reports
        
        if not self.selected_report.get():
            if len(self.filtered_reports) == len(self.sorted_reports):
                self.report_desc.config(text="Browse all 179 available reports")
            else:
                self.report_desc.config(text=f"Showing {len(self.filtered_reports)} filtered reports")
    
    def on_arrow_down(self, event):
        """Handle down arrow key"""
        # When user presses down arrow, open dropdown with current filtered results
        self.report_combo['values'] = self.filtered_reports
        return "break"  # Allow normal dropdown navigation
    
    def on_arrow_up(self, event):
        """Handle up arrow key"""
        # Similar to down arrow
        self.report_combo['values'] = self.filtered_reports
        return "break"  # Allow normal dropdown navigation
        
    def on_dropdown_focus(self, event):
        """Handle when dropdown gets focus"""
        # Show placeholder text or instructions
        current_text = self.selected_report.get()
        if not current_text:
            self.report_desc.config(text="Type to filter reports or click dropdown arrow to browse all 179 available reports")
    
    def on_report_selected(self, event=None):
        """Handle report selection with parameter indication in description only"""
        selected = self.selected_report.get()
        
        # Use filtered reports if available, otherwise fall back to all reports
        available_reports = getattr(self, 'filtered_available_reports', None)
        if available_reports is None:
            available_reports = getattr(self, 'available_reports', {})
        if not available_reports:
            available_reports = {}
        
        # Check if it's a valid report selection
        if selected and selected in available_reports:
            description = available_reports[selected]
            
            # Get parameter requirements for this report
            param_info = self.get_parameter_info(selected)
            
            # Update description with parameter information (keep visual indicators in description only)
            full_description = f"Selected: {description}"
            if param_info:
                full_description += f"\n{param_info['icon']} {param_info['description']}"
                
                if param_info['requirement_level'] == 'mandatory':
                    full_description += "\n⚠️ This report requires configuration before export"
                # Skip optional parameter message - no longer showing optional dialogs
            
            self.report_desc.config(text=full_description)
            
            # Reset search tracking
            self.last_search_text = selected.lower()
            # Reset to show all available reports for next search
            sorted_available = sorted(list(available_reports.keys()))
            self.filtered_reports = sorted_available.copy()
            self.report_combo['values'] = sorted_available
            
            # Enable export button when valid report is selected
            self.export_btn.config(state='normal')
            
            # Update export button text based on parameter requirements
            if param_info and param_info['requirement_level'] == 'mandatory':
                self.export_btn.config(text="🔧 Configure & Export*")
            else:
                self.export_btn.config(text="📥 Export Report")
            
            self.log_message(f"Selected report: {description}", 'info')
            
        elif not selected:
            self.report_desc.config(text="Select a report to export all available data")
            self.last_search_text = ""
            # Reset to show all available reports
            sorted_available = sorted(list(available_reports.keys())) if hasattr(self, 'filtered_available_reports') else self.sorted_reports
            self.filtered_reports = sorted_available.copy()
            self.report_combo['values'] = sorted_available
            # Disable both export buttons when no report is selected
            self.export_btn.config(state='disabled', text="📥 Export Report")
    
    def get_parameter_info(self, report_name):
        """Get parameter information for a report with auto-detection fallback"""
        
        # First check our predefined requirements
        if report_name in self.report_parameter_requirements:
            return self.report_parameter_requirements[report_name]
        
        # Fall back to auto-detection
        return self.auto_detect_parameter_requirements(report_name)
    
    def get_report_with_indicators(self):
        """Get clean report list without visual indicators"""
        clean_reports = {}
        
        for report_name, description in self.available_reports.items():
            # Return clean description without any icons
            clean_reports[report_name] = description
        
        return clean_reports
            
    def find_free_port(self):
        """Find free port for callback"""
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('', 0))
            return s.getsockname()[1]
    
    def login(self):
        """Start OAuth2 login"""
        self.login_status.config(text="Initializing authentication...", fg='#0078d4')
        self.login_btn.config(state='disabled', text="Authenticating...")
        self.root.update()
        
        thread = threading.Thread(target=self.authenticate)
        thread.daemon = True
        thread.start()
    
    def authenticate(self):
        """OAuth2 authentication flow"""
        try:
            port = self.find_free_port()
            callback_url = f"http://localhost:{port}/callback"
            
            server = HTTPServer(('localhost', port), AuthCallbackHandler)
            server.timeout = 300
            server.auth_code = None
            
            # Build auth URL
            params = {
                'client_id': self.client_id,
                'response_type': 'code', 
                'redirect_uri': callback_url,
                'scope': 'https://graph.microsoft.com/.default offline_access',
                'state': f'intune_reports_{int(time.time())}',
                'prompt': 'login',
                'domain_hint': 'organizations'
            }
            
            auth_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/authorize?" + urllib.parse.urlencode(params)
            
            webbrowser.open(auth_url)
            
            self.root.after(0, lambda: self.login_status.config(
                text="Waiting for authentication...", fg='#0078d4'))
            
            # Wait for callback
            start_time = time.time()
            while time.time() - start_time < 300:
                try:
                    server.handle_request()
                    break
                except socket.timeout:
                    continue
            
            if server.auth_code:
                # Exchange code for token
                self.root.after(0, lambda: self.login_status.config(
                    text="Processing authorization...", fg='#0078d4'))
                
                token_data = {
                    'grant_type': 'authorization_code',
                    'client_id': self.client_id,
                    'code': server.auth_code,
                    'redirect_uri': callback_url,
                    'scope': 'https://graph.microsoft.com/.default',
                    'client_secret': self.client_secret
                }
                
                token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
                timeout = self.timeout_manager.get_timeout_for_operation('authentication')
                response = requests.post(token_url, data=token_data, timeout=timeout)
                
                if response.status_code == 200:
                    token_info = response.json()
                    self.access_token = token_info['access_token']
                    self.refresh_token = token_info.get('refresh_token')
                    
                    # Track token expiry
                    from datetime import datetime, timedelta
                    expires_in = token_info.get('expires_in', 3600)  # Default 1 hour
                    self.token_issued_at = datetime.now()
                    self.token_expires_at = self.token_issued_at + timedelta(seconds=expires_in)
                    
                    # Get user info
                    headers = {'Authorization': f'Bearer {self.access_token}'}
                    user_response = requests.get(f"{self.graph_base_url}/me", headers=headers)
                    
                    if user_response.status_code == 200:
                        self.user_info = user_response.json()
                        user_name = self.user_info.get('displayName', 'User')
                        
                        self.root.after(0, lambda: self.login_status.config(
                            text=f"Authentication successful! Welcome {user_name}", fg='#107c10'))
                        self.root.after(2000, self.show_reports_page)
                    else:
                        self.root.after(0, lambda: self.login_status.config(
                            text="Failed to get user info", fg='#d13438'))
                        self.reset_login_button()
                else:
                    error_detail = self.parse_error_response(response)
                    self.root.after(0, lambda: self.login_status.config(
                        text=f"Token exchange failed: {error_detail}", fg='#d13438'))
                    self.reset_login_button()
            else:
                self.root.after(0, lambda: self.login_status.config(
                    text="Authentication timeout or cancelled", fg='#d13438'))
                self.reset_login_button()
                
        except Exception as e:
            self.root.after(0, lambda: self.login_status.config(
                text=f"Authentication error: {str(e)}", fg='#d13438'))
            self.reset_login_button()
    
    def parse_error_response(self, response):
        """Parse error information from API response"""
        try:
            if response.headers.get('content-type', '').startswith('application/json'):
                error_data = response.json()
                if 'error' in error_data:
                    if isinstance(error_data['error'], dict):
                        return f"{error_data['error'].get('code', 'Unknown')}: {error_data['error'].get('message', 'No details')}"
                    else:
                        return f"{error_data['error']}: {error_data.get('error_description', 'No details')}"
            return f"HTTP {response.status_code}: {response.reason}"
        except:
            return f"HTTP {response.status_code}: Unable to parse error"
    
    def reset_login_button(self):
        """Reset login button state"""
        self.root.after(0, lambda: self.login_btn.config(
            state='normal', text="Sign in with Corporate Account"))
    
    def refresh_access_token(self, max_retries=3):
        """Enhanced token refresh with retry logic and expiry tracking"""
        try:
            if not self.refresh_token:
                self.log_message("No refresh token available, need to re-authenticate", 'warning')
                return False
            
            self.log_message("Attempting to refresh access token...", 'info')
            
            token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
            
            token_data = {
                'client_id': self.client_id,
                'client_secret': self.client_secret,
                'scope': 'https://graph.microsoft.com/.default',
                'refresh_token': self.refresh_token,
                'grant_type': 'refresh_token'
            }
            
            # Retry logic for token refresh
            for attempt in range(max_retries):
                try:
                    timeout = self.timeout_manager.get_timeout_for_operation('token_refresh')
                    token_response = requests.post(token_url, data=token_data, timeout=timeout)
                    
                    if token_response.status_code == 200:
                        token_info = token_response.json()
                        
                        # Store token with expiry tracking
                        self.access_token = token_info['access_token']
                        if 'refresh_token' in token_info:
                            self.refresh_token = token_info['refresh_token']
                        
                        # Calculate token expiry (default 1 hour if not specified)
                        from datetime import datetime, timedelta
                        expires_in = token_info.get('expires_in', 3600)  # Default 1 hour
                        self.token_issued_at = datetime.now()
                        self.token_expires_at = self.token_issued_at + timedelta(seconds=expires_in)
                        
                        self.log_message(f"✅ Access token refreshed successfully (expires at {self.token_expires_at.strftime('%H:%M:%S')})", 'success')
                        return True
                    
                    elif token_response.status_code == 400:
                        # Bad request - refresh token might be expired
                        error_data = token_response.json() if token_response.text else {}
                        error_code = error_data.get('error', 'unknown')
                        
                        if error_code in ['invalid_grant', 'expired_token']:
                            self.log_message("❌ Refresh token expired, need to re-authenticate", 'error')
                            self.access_token = None
                            self.refresh_token = None
                            self.token_expires_at = None
                            return False
                        
                        self.log_message(f"❌ Token refresh failed (400): {token_response.text}", 'error')
                        return False
                    
                    elif token_response.status_code == 429:
                        # Rate limited
                        wait_time = self.rate_limiter.handle_429_response(token_response)
                        self.log_message(f"Token refresh rate limited, waiting {wait_time:.1f} seconds", 'warning')
                        time.sleep(wait_time)
                        continue
                    
                    else:
                        self.log_message(f"❌ Token refresh failed (HTTP {token_response.status_code}): {token_response.text}", 'error')
                        
                        if attempt < max_retries - 1:
                            wait_time = self.timeout_manager.get_exponential_backoff_delay(attempt)
                            self.log_message(f"Retrying token refresh in {wait_time:.1f} seconds (attempt {attempt + 2}/{max_retries})", 'info')
                            time.sleep(wait_time)
                            continue
                        
                        return False
                
                except requests.exceptions.Timeout:
                    if attempt < max_retries - 1:
                        wait_time = self.timeout_manager.get_exponential_backoff_delay(attempt)
                        self.log_message(f"Token refresh timeout, retrying in {wait_time:.1f} seconds", 'warning')
                        time.sleep(wait_time)
                        continue
                    else:
                        self.log_message("❌ Token refresh failed: timeout", 'error')
                        return False
                
                except requests.exceptions.RequestException as e:
                    if attempt < max_retries - 1:
                        wait_time = self.timeout_manager.get_exponential_backoff_delay(attempt)
                        self.log_message(f"Token refresh network error, retrying in {wait_time:.1f} seconds: {str(e)}", 'warning')
                        time.sleep(wait_time)
                        continue
                    else:
                        self.log_message(f"❌ Token refresh failed: {str(e)}", 'error')
                        return False
            
            return False
                
        except Exception as e:
            self.log_message(f"❌ Error refreshing token: {str(e)}", 'error')
            return False
    
    def token_expires_soon(self, buffer_minutes=5):
        """Check if token expires within buffer_minutes"""
        if not self.token_expires_at:
            return False
        
        from datetime import datetime, timedelta
        buffer_time = datetime.now() + timedelta(minutes=buffer_minutes)
        return buffer_time >= self.token_expires_at
    
    def is_token_valid(self):
        """Check if current token is valid and not expired"""
        if not self.access_token:
            return False
        
        if not self.token_expires_at:
            return True  # Assume valid if no expiry info
        
        from datetime import datetime
        return datetime.now() < self.token_expires_at

    def make_authenticated_request(self, method, url, operation_type='api_call', max_retries=3, **kwargs):
        """Make API calls with error handling and retries"""
        
        # Refresh token if it expires soon
        if self.token_expires_soon():
            self.log_message("Token expires soon, refreshing now...", 'info')
            self.refresh_access_token()
        
        # Check if token is still valid
        if not self.is_token_valid():
            self.log_message("❌ Invalid or expired token", 'error')
            self.root.after(0, lambda: messagebox.showerror("Authentication Expired", 
                "Your session has expired. Please logout and login again."))
            return None
        
        # Set adaptive timeout if not provided
        if 'timeout' not in kwargs:
            kwargs['timeout'] = self.timeout_manager.get_timeout_for_operation(operation_type)
        
        # Prepare headers
        headers = kwargs.get('headers', {})
        headers['Authorization'] = f'Bearer {self.access_token}'
        kwargs['headers'] = headers
        
        last_response = None
        
        for attempt in range(max_retries):
            try:
                # Apply rate limiting
                self.rate_limiter.wait_if_needed()
                
                # Make the request
                self.log_message(f"Making {method.upper()} request to {url} (attempt {attempt + 1}/{max_retries})", 'debug')
                
                response = requests.request(method, url, **kwargs)
                last_response = response
                
                self.log_message(f"Response: HTTP {response.status_code}", 'debug')
                
                # Check response codes
                if response.status_code in [200, 201, 204]:
                    # Success
                    return response
                
                elif response.status_code == 401:
                    # Unauthorized - token expired or invalid
                    self.log_message("⚠️ Authentication failed (401), attempting token refresh...", 'warning')
                    
                    if self.refresh_access_token():
                        # Update headers with new token
                        headers['Authorization'] = f'Bearer {self.access_token}'
                        kwargs['headers'] = headers
                        self.log_message("🔄 Token refreshed, retrying request...", 'info')
                        continue
                    else:
                        # Refresh failed
                        self.log_message("❌ Token refresh failed, authentication required", 'error')
                        self.root.after(0, lambda: messagebox.showerror("Authentication Expired", 
                            "Your session has expired. Please logout and login again."))
                        return response
                
                elif response.status_code == 429:
                    # Rate limited
                    wait_time = self.rate_limiter.handle_429_response(response)
                    self.log_message(f"⚠️ Rate limited (429), waiting {wait_time:.1f} seconds...", 'warning')
                    time.sleep(wait_time)
                    continue
                
                elif response.status_code == 403:
                    # Forbidden - insufficient permissions
                    self.log_message(f"❌ Insufficient permissions (403): {response.text}", 'error')
                    return response
                
                elif response.status_code == 404:
                    # Not found
                    self.log_message(f"❌ Resource not found (404): {url}", 'error')
                    return response
                
                elif response.status_code >= 500:
                    # Server error - retry with backoff
                    if attempt < max_retries - 1:
                        wait_time = self.timeout_manager.get_exponential_backoff_delay(attempt)
                        self.log_message(f"⚠️ Server error ({response.status_code}), retrying in {wait_time:.1f} seconds...", 'warning')
                        time.sleep(wait_time)
                        continue
                    else:
                        self.log_message(f"❌ Server error ({response.status_code}) after {max_retries} attempts", 'error')
                        return response
                
                elif 400 <= response.status_code < 500:
                    # Client error - don't retry
                    self.log_message(f"❌ Client error ({response.status_code}): {response.text}", 'error')
                    return response
                
                else:
                    # Unexpected status code
                    if attempt < max_retries - 1:
                        wait_time = self.timeout_manager.get_exponential_backoff_delay(attempt)
                        self.log_message(f"⚠️ Unexpected response ({response.status_code}), retrying in {wait_time:.1f} seconds...", 'warning')
                        time.sleep(wait_time)
                        continue
                    else:
                        return response
            
            except requests.exceptions.Timeout:
                if attempt < max_retries - 1:
                    wait_time = self.timeout_manager.get_exponential_backoff_delay(attempt)
                    self.log_message(f"⚠️ Request timeout, retrying in {wait_time:.1f} seconds (attempt {attempt + 2}/{max_retries})", 'warning')
                    time.sleep(wait_time)
                    continue
                else:
                    self.log_message(f"❌ Request timeout after {max_retries} attempts", 'error')
                    raise
            
            except requests.exceptions.ConnectionError as e:
                if attempt < max_retries - 1:
                    wait_time = self.timeout_manager.get_exponential_backoff_delay(attempt)
                    self.log_message(f"⚠️ Connection error, retrying in {wait_time:.1f} seconds: {str(e)}", 'warning')
                    time.sleep(wait_time)
                    continue
                else:
                    self.log_message(f"❌ Connection error after {max_retries} attempts: {str(e)}", 'error')
                    raise
            
            except requests.exceptions.RequestException as e:
                if attempt < max_retries - 1:
                    wait_time = self.timeout_manager.get_exponential_backoff_delay(attempt)
                    self.log_message(f"⚠️ Request error, retrying in {wait_time:.1f} seconds: {str(e)}", 'warning')
                    time.sleep(wait_time)
                    continue
                else:
                    self.log_message(f"❌ Request error after {max_retries} attempts: {str(e)}", 'error')
                    raise
        
        # If we get here, all retries failed
        return last_response

    def fetch_all_graph_pages(self, url, params=None, operation_type='api_call', max_pages=2000, progress_callback=None):
        """Fetch all paged Microsoft Graph results using @odata.nextLink."""
        all_items = []
        current_url = url
        current_params = dict(params) if params else {}
        page_number = 0

        while current_url:
            page_number += 1
            if page_number > max_pages:
                raise Exception(f"Paging limit reached ({max_pages} pages). Aborting to prevent infinite loop.")

            response = self.make_authenticated_request(
                'GET',
                current_url,
                operation_type=operation_type,
                params=current_params
            )

            if not response:
                raise Exception("No response received from Microsoft Graph.")

            if response.status_code != 200:
                error_detail = self.parse_error_response(response)
                raise Exception(f"Graph paging request failed (HTTP {response.status_code}): {error_detail}")

            try:
                data = response.json()
            except json.JSONDecodeError:
                raise Exception("Invalid JSON response returned by Microsoft Graph.")

            page_items = data.get('value', [])
            if not isinstance(page_items, list):
                raise Exception("Unexpected Microsoft Graph response format: 'value' is not a list.")

            all_items.extend(page_items)

            if progress_callback:
                progress_callback(page_number, len(page_items), len(all_items))

            current_url = data.get('@odata.nextLink')
            # nextLink already includes query string parameters.
            current_params = None

        return all_items

    def manual_token_refresh(self):
        """Manually refresh the token when user clicks the button"""
        try:
            if self.refresh_access_token():
                messagebox.showinfo("Token Refreshed", "✅ Authentication token refreshed successfully!")
                self.log_message("✅ Manual token refresh successful", 'success')
            else:
                messagebox.showerror("Refresh Failed", "❌ Token refresh failed. Please logout and login again.")
                self.log_message("❌ Manual token refresh failed", 'error')
        except Exception as e:
            messagebox.showerror("Refresh Error", f"❌ Error refreshing token: {str(e)}")
            self.log_message(f"❌ Manual token refresh error: {str(e)}", 'error')
    
    def estimate_export_size(self, report_name):
        """Estimate the number of records in an export for adaptive timeout calculation"""
        
        # Define estimated record counts based on report type
        size_estimates = {
            # Device reports - can be large in enterprise
            'Devices': 100000,
            'DevicesWithInventory': 100000,
            'DeviceCompliance': 50000,
            'DeviceNonCompliance': 25000,
            'DevicesWithoutInventory': 10000,
            
            # User reports
            'UserInstallStateSummary': 50000,
            'UserDeviceAssociations': 25000,
            
            # App reports
            'AppInstallStatusAggregate': 75000,
            'AllAppsList': 5000,
            
            # Policy reports
            'PolicyNonCompliance': 15000,
            'SettingsNonCompliance': 20000,
            
            # Security reports
            'ActiveMalware': 1000,
            'Malware': 5000,
            'DefenderAgents': 100000,
            
            # Certificate reports
            'CertificateReport': 10000,
            
            # Default for unknown reports
            'default': 10000
        }
        
        estimated_size = size_estimates.get(report_name, size_estimates['default'])
        
        # Add some variance based on typical enterprise sizes
        # Small: < 1K devices, Medium: 1K-10K, Large: 10K-100K, Enterprise: 100K+
        if report_name in ['Devices', 'DevicesWithInventory', 'DefenderAgents']:
            # These could be very large in enterprise environments
            estimated_size = max(estimated_size, 50000)
        
        return estimated_size
    
    def convert_ui_params_to_api(self, ui_parameters):
        """Convert UI parameter names to API parameter names"""
        
        api_params = {}
        
        for ui_param, value in ui_parameters.items():
            if ui_param == 'deviceId':
                # For device-specific reports, often need a filter
                api_params['filter'] = f"DeviceId eq '{value}'"
            elif ui_param == 'policyId':
                # For policy-specific reports
                api_params['filter'] = f"PolicyId eq '{value}'"
            elif ui_param in ['startDate', 'endDate']:
                # Date parameters map directly
                api_params[ui_param] = value
            elif ui_param == 'filter':
                # Filter parameters map directly
                api_params['filter'] = value
            elif ui_param == 'top':
                # Top parameter for limiting results
                try:
                    api_params['top'] = int(value)
                except ValueError:
                    api_params['top'] = 1000  # Default fallback
        
        return api_params
    
    def merge_api_parameters(self, default_params, user_params, report_name):
        """Merge user parameters with default parameters for direct API calls"""
        
        merged_params = default_params.copy()
        
        if not user_params:
            return merged_params
        
        # Handle different parameter types based on report type
        filters = []
        
        for param_name, param_value in user_params.items():
            if param_name == 'deviceId' and param_value:
                # For app reports, we need to modify approach since /mobileApps doesn't filter by device
                if report_name in ['DevicesByAppInv', 'AppInvByDevice']:
                    # Store device ID for post-processing filtering (Graph API limitation)
                    merged_params['_post_filter_deviceId'] = param_value
                    self.log_message(f"Will filter results by deviceId: {param_value} after API call", 'info')
                else:
                    filters.append(f"managedDeviceId eq '{param_value}'")
                    
            elif param_name == 'policyId' and param_value:
                # Add policy filter based on report type
                if report_name == 'Policies':
                    filters.append(f"id eq '{param_value}'")
                else:
                    filters.append(f"policyId eq '{param_value}'")
                
            elif param_name == 'userId' and param_value:
                # Add user filter
                filters.append(f"userId eq '{param_value}'")
                
            elif param_name == 'applicationId' and param_value:
                # Add application filter
                filters.append(f"id eq '{param_value}'")
                
            elif param_name in ['startDate', 'endDate'] and param_value:
                # Handle date parameters
                if param_name == 'startDate':
                    filters.append(f"createdDateTime ge {param_value}T00:00:00Z")
                elif param_name == 'endDate':
                    filters.append(f"createdDateTime le {param_value}T23:59:59Z")
                    
            elif param_name == 'top' and param_value:
                # Override default $top
                try:
                    merged_params['$top'] = int(param_value)
                except ValueError:
                    pass  # Keep default
        
        # Combine filters with AND logic
        if filters:
            existing_filter = merged_params.get('$filter', '')
            if existing_filter:
                combined_filter = f"({existing_filter}) and ({' and '.join(filters)})"
            else:
                combined_filter = ' and '.join(filters)
            merged_params['$filter'] = combined_filter
            
        self.log_message(f"Merged parameters for {report_name}: {merged_params}", 'debug')
        return merged_params
    
    def apply_post_processing_filters(self, df, report_name, user_params):
        """Apply client-side filtering for cases where API doesn't support certain filters"""
        
        original_count = len(df)
        original_df = df.copy()  # Keep a copy for safety fallback
        
        self.log_message(f"=== FILTERING DEBUG START ===", 'warning')
        self.log_message(f"Original DataFrame: {original_count} records", 'warning')
        self.log_message(f"Report name: {report_name}", 'warning')
        self.log_message(f"User parameters: {user_params}", 'warning')
        
        # SAFETY CHECK: If we started with no data, don't filter
        if original_count == 0:
            self.log_message("Original data is already empty - skipping filters", 'warning')
            return df
        
        # If no meaningful parameters provided, return original data
        has_filters = any(param_value and str(param_value).strip() and str(param_value).strip() != '' 
                         for param_name, param_value in user_params.items() 
                         if param_name in ['deviceId', 'policyId', 'userId', 'applicationId'] and param_value)
        
        self.log_message(f"Has meaningful filters: {has_filters}", 'warning')
        
        if not has_filters:
            self.log_message("No applicable filters found - returning ALL original data", 'warning')
            self.log_message(f"=== FILTERING DEBUG END: RETURNING {len(df)} RECORDS ===", 'warning')
            return df
        
        # TEMPORARILY DISABLE DEVICE FILTERING - IT'S CAUSING ISSUES
        if 'deviceId' in user_params and user_params['deviceId']:
            device_id = user_params['deviceId']
            self.log_message(f"TEMPORARILY SKIPPING device filtering for device: {device_id}", 'warning')
            self.log_message(f"Device filtering is disabled until the issue is resolved", 'warning')
            
            # Original complex device filtering logic is commented out
            # if report_name in ['DevicesByAppInv', 'AppInvByDevice']:
            #     ... complex filtering logic that was causing issues ...
        
        # Handle other filters based on column names
        for param_name, param_value in user_params.items():
            if not param_value:
                continue
                
            if param_name == 'policyId' and 'id' in df.columns:
                df = df[df['id'] == param_value]
                self.log_message(f"Filtered by policyId: {len(df)} records remaining", 'info')
                
            elif param_name == 'userId' and 'userId' in df.columns:
                df = df[df['userId'] == param_value]
                self.log_message(f"Filtered by userId: {len(df)} records remaining", 'info')
                
            elif param_name == 'applicationId' and 'id' in df.columns:
                df = df[df['id'] == param_value]
                self.log_message(f"Filtered by applicationId: {len(df)} records remaining", 'info')
        
        filtered_count = len(df)
        
        # SAFETY CHECK: If filtering removed ALL data, return original data instead
        if filtered_count == 0 and original_count > 0:
            self.log_message(f"WARNING: Filtering removed ALL data ({original_count} → 0). Returning original data to avoid empty result.", 'warning')
            self.log_message(f"=== FILTERING DEBUG END: SAFETY FALLBACK - RETURNING {original_count} RECORDS ===", 'warning')
            return original_df
        
        if filtered_count != original_count:
            self.log_message(f"Post-processing filters applied: {original_count} → {filtered_count} records", 'info')
        
        self.log_message(f"=== FILTERING DEBUG END: RETURNING {filtered_count} RECORDS ===", 'warning')
        return df
    
    def get_device_specific_apps(self, device_id):
        """Get apps installed on a specific device"""
        try:
            # Use the device management endpoint to get apps for specific device
            url = f"https://graph.microsoft.com/beta/deviceManagement/managedDevices/{device_id}/detectedApps"
            
            response = self.make_authenticated_request('GET', url, 
                                                     operation_type='api_call',
                                                     params={'$select': 'id,displayName,version,publisher,sizeInByte'})
            
            if response and response.status_code == 200:
                data = response.json()
                if 'value' in data and data['value']:
                    import pandas as pd
                    return pd.DataFrame(data['value'])
            
            return None
            
        except Exception as e:
            self.log_message(f"Error getting device-specific apps: {str(e)}", 'debug')
            return None
    
    def auto_detect_parameter_requirements(self, report_name):
        """Auto-detect parameter requirements for reports not in our database"""
        
        report_lower = report_name.lower()
        
        # Device-specific reports - make mandatory for "ByDevice" reports, none for others
        if 'bydevice' in report_lower or 'perdevice' in report_lower:
            return {
                "requirement_level": "mandatory",
                "icon": "🔴",
                "parameters": {
                    "deviceId": {"type": "device_selector", "required": True, "description": "Select target device"}
                },
                "description": "Requires specific device selection"
            }
        elif any(word in report_lower for word in ['device', 'inventory']):
            return {
                "requirement_level": "none",
                "icon": "�",
                "description": "No input required"
            }
        
        # Policy-specific reports - no longer showing optional parameters
        elif any(word in report_lower for word in ['policy', 'compliance', 'configuration', 'settings']):
            return {
                "requirement_level": "none",
                "icon": "�",
                "description": "No input required"
            }
        
        # Update reports - no longer showing optional parameters
        elif any(word in report_lower for word in ['update', 'feature', 'quality']):
            return {
                "requirement_level": "none",
                "icon": "�",
                "description": "No input required"
            }
        
        # Default: no requirements
        return {
            "requirement_level": "none",
            "icon": "�",
            "description": "No input required"
        }

    def refresh_available_reports(self):
        """Refresh the available reports list - show all immediately, refine in background"""
        try:
            # FAST APPROACH: Show all reports immediately for instant UI response
            self.log_message("Loading reports interface - permission checking in background", 'info')
            
            # Show all reports immediately so user can see the interface
            self.filtered_available_reports = self.available_reports.copy()
            
            # Start permission discovery in background (non-blocking)
            threading.Thread(target=self.discover_permissions_background, daemon=True).start()
            
            self.log_message(f"Showing all {len(self.filtered_available_reports)} reports - access validation in progress", 'success')
            
            return True
            
        except Exception as e:
            self.log_message(f"Error refreshing available reports: {str(e)}", 'error')
            # Fallback to showing all reports
            self.filtered_available_reports = self.available_reports.copy()
            return False

    def discover_permissions_background(self):
        """Discover user permissions in background and update UI when complete"""
        try:
            import threading
            import time
            
            # Add a small delay to let UI finish loading
            time.sleep(1)
            
            self.log_message("Starting background permission discovery", 'info')
            
            # Simple and fast permission check - just test one key endpoint
            if not self.access_token:
                return
                
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Content-Type': 'application/json'
            }
            
            # Test just one endpoint to determine if user has admin access
            try:
                # Test device management access
                test_url = f"{self.graph_base_url}/deviceManagement/managedDevices?$top=1"
                response = requests.get(test_url, headers=headers, timeout=10)
                
                if response.status_code in [200, 206]:
                    # User has device management access - likely admin
                    user_name = self.user_info.get('displayName', 'User') if self.user_info else 'User'
                    self.log_message(f"✅ {user_name} has administrative access to all {len(self.available_reports)} reports", 'success')
                    self.user_access_level = 'admin'
                    # Keep all reports available
                else:
                    # Limited access user
                    user_name = self.user_info.get('displayName', 'User') if self.user_info else 'User'
                    self.log_message(f"⚠️ {user_name} has limited access - some reports may be restricted", 'warning')
                    self.user_access_level = 'limited'
                    # You could add more detailed checking here if needed
                    
            except Exception as e:
                self.log_message(f"Background permission check completed with fallback", 'info')
                self.user_access_level = 'unknown'
                
        except Exception as e:
            self.log_message(f"Background permission discovery error: {str(e)}", 'debug')

    def logout(self):
        """Logout and return to login"""
        self.access_token = None
        self.refresh_token = None
        self.user_info = None
        self.current_export_data = None
        self.current_columns = []
        # Clear permission cache for next user
        self.user_permissions_cache = None
        self.filtered_available_reports = None
        self.user_access_level = 'unknown'
        self.show_login_page()
    
    def export_report(self):
        """Export the selected report with parameter collection if needed"""
        try:
            self.log_message("Export button clicked - starting validation", 'debug')
            
            if not self.access_token:
                self.log_message("No access token found", 'error')
                messagebox.showerror("Authentication Error", "Please login first")
                return
            
            selected_report = self.selected_report.get()
            self.log_message(f"Selected report: '{selected_report}'", 'debug')
            
            if not selected_report:
                self.log_message("No report selected", 'error')
                messagebox.showwarning("Selection Error", "Please select a report")
                return
            
            # Check if this report requires parameters
            param_config = self.get_parameter_info(selected_report)
            
            # Show parameter dialog only if mandatory parameters are required
            if param_config and param_config['requirement_level'] == 'mandatory':
                self.log_message(f"Report {selected_report} has mandatory parameters - showing dialog", 'info')
                
                # Show parameter dialog for mandatory parameters
                dialog = ParameterDialog(self, selected_report, param_config)
                
                if dialog.result != 'ok':
                    self.log_message("Parameter dialog cancelled by user", 'info')
                    return
                
                # Store the collected parameters
                self.current_export_parameters = dialog.parameters
                self.log_message(f"Collected parameters: {list(self.current_export_parameters.keys())}", 'info')
            else:
                # No parameter dialog needed
                self.current_export_parameters = {}
                self.log_message("No mandatory parameters required - proceeding with export", 'info')
            
            self.log_message("=== EXPORT STARTED ===", 'info')
            self.log_message(f"Report: {selected_report}", 'info')
            self.log_message("Strategy: Export all available columns", 'info')
            self.log_message(f"Access token available: {bool(self.access_token)}", 'debug')
            
            # Check if this is a direct API call report
            if selected_report in self.direct_api_reports:
                self.log_message(f"Using direct API call for {selected_report}", 'info')
                self.export_direct_api_report(selected_report)
            else:
                self.log_message(f"Using export job for {selected_report}", 'info')
                self.export_via_export_job(selected_report)
                
        except Exception as e:
            self.log_message(f"Error in export_report function: {str(e)}", 'error')
            self.log_message(f"Traceback: {traceback.format_exc()}", 'debug')
            messagebox.showerror("Export Error", f"Error starting export: {str(e)}")
    
    def export_direct_api_report(self, report_name):
        """Export reports using direct API GET calls"""
        # Disable export button and start progress
        self.export_btn.config(state='disabled')
        self.progress.start()
        self.progress_label.config(text="Starting direct API export...")
        
        # Start export in background thread
        self.log_message("Starting background thread for direct API export", 'debug')
        thread = threading.Thread(target=self.direct_api_thread, args=(report_name,))
        thread.daemon = True
        thread.start()
    
    def export_via_export_job(self, report_name):
        """Export reports using traditional export job method"""
        # Disable export button and start progress
        self.export_btn.config(state='disabled')
        self.progress.start()
        self.progress_label.config(text="Starting export job...")
        
        # Start export in background thread
        self.log_message("Starting background thread for export job", 'debug')
        thread = threading.Thread(target=self.export_thread, args=(report_name,))
        thread.daemon = True
        thread.start()
    
    def direct_api_thread(self, report_name):
        """Thread for direct API calls to Graph endpoints"""
        try:
            self.log_message(f"Direct API thread started for report: {report_name}", 'debug')
            
            # Get configuration for this report
            report_config = self.direct_api_reports[report_name]
            endpoint = report_config["endpoint"]
            base_url = report_config["base_url"]
            required_permission = report_config["required_permission"]
            default_parameters = report_config["parameters"].copy()
            
            # Collect user-provided parameters if parameter dialog was used
            user_parameters = {}
            if hasattr(self, 'current_export_parameters') and self.current_export_parameters:
                user_parameters = self.current_export_parameters.copy()
                self.log_message(f"User provided parameters: {user_parameters}", 'info')
                
                # Provide user feedback about filtering
                filter_info = []
                for param, value in user_parameters.items():
                    if value:
                        if param == 'deviceId':
                            filter_info.append(f"Device: {value}")
                        elif param == 'policyId':
                            filter_info.append(f"Policy: {value}")
                        elif param == 'userId':
                            filter_info.append(f"User: {value}")
                        elif param == 'applicationId':
                            filter_info.append(f"Application: {value}")
                        elif param in ['startDate', 'endDate']:
                            filter_info.append(f"{param}: {value}")
                
                if filter_info:
                    filter_text = f"Applying filters: {', '.join(filter_info)}"
                    self.log_message(filter_text, 'info')
                    self.root.after(0, lambda: self.progress_label.config(text=f"Filtering data: {', '.join(filter_info)}"))
            
            # Merge user parameters with default parameters
            final_parameters = self.merge_api_parameters(default_parameters, user_parameters, report_name)
            
            # Build full URL with parameters
            url = f"{base_url}{endpoint}"
            
            self.log_message(f"Direct API Endpoint: {url}", 'api')
            self.log_message(f"Required Permission: {required_permission}", 'info')
            self.log_message(f"Final API Parameters: {final_parameters}", 'debug')
            
            # Debug: Check token permissions
            self.debug_token_permissions()
            
            self.root.after(0, lambda: self.progress_label.config(text="Making API call..."))
            
            try:
                def on_page_fetched(page_number, page_count, total_count):
                    self.log_message(
                        f"Fetched page {page_number}: {page_count} items (total: {total_count})",
                        'debug'
                    )
                    self.root.after(0, lambda: self.progress_label.config(
                        text=f"Fetching page {page_number}... Total records: {total_count:,}"
                    ))

                items = self.fetch_all_graph_pages(
                    url,
                    params=final_parameters,
                    operation_type='api_call',
                    progress_callback=on_page_fetched
                )

            except requests.exceptions.Timeout:
                raise Exception("Request timeout - API took too long to respond")
            except requests.exceptions.ConnectionError as e:
                raise Exception(f"Connection error - Unable to reach Microsoft Graph API: {str(e)}")
            except requests.exceptions.RequestException as e:
                raise Exception(f"Request failed: {str(e)}")

            self.log_message(f"API returned {len(items)} total items", 'success')

            if items:
                # Log first item structure for debugging
                self.log_message(f"First item keys: {list(items[0].keys()) if items[0] else 'No keys'}", 'debug')

                # Convert to DataFrame
                import pandas as pd
                df = pd.DataFrame(items)
                self.log_message(f"Created DataFrame with shape: {df.shape}", 'debug')

                # Show the data
                self.root.after(0, lambda: self.progress_label.config(text="Processing data..."))
                self.process_direct_api_data(df, report_name)
            else:
                self.log_message("No data returned from API", 'warning')
                raise Exception("No data available for this report")
                
        except Exception as e:
            self.log_message(f"Direct API thread failed: {str(e)}", 'error')
            self.log_message(f"Error trace: {traceback.format_exc()}", 'debug')
            
            # Capture error message before lambda to avoid scoping issues
            error_msg = str(e)
            self.root.after(0, lambda: messagebox.showerror(
                "Export Failed", f"Failed to export {report_name}:\n\n{error_msg}"))
        
        finally:
            # Re-enable export button and stop progress
            self.log_message("Direct API thread finishing - restoring UI state", 'debug')
            self.root.after(0, lambda: self.export_btn.config(state='normal'))
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.progress_label.config(text="Ready"))
    
    def process_direct_api_data(self, df, report_name):
        """Process data from direct API calls and show export interface"""
        try:
            self.log_message(f"Processing {len(df)} rows of data for {report_name}", 'info')
            
            # Apply data transformations based on report type
            if report_name == "AllGroupsInMyOrg" and 'groupTypes' in df.columns:
                self.log_message("Transforming groupTypes to user-friendly format", 'debug')
                df = self.transform_group_types(df)
            
            # Convert DataFrame to list of dictionaries (same format as CSV data)
            data_as_dicts = df.to_dict('records')
            
            # Store the data in the same format as traditional reports
            self.current_export_data = data_as_dicts
            self.current_columns = list(df.columns)
            
            # Update status and show success message
            row_count = len(data_as_dicts)
            col_count = len(self.current_columns)
            
            self.root.after(0, lambda: self.progress_label.config(text="Export completed"))
            self.log_message(f"✅ {report_name} exported successfully!", 'success')
            
            # Create the column selection tab (same as traditional exports)
            self.root.after(0, self.create_columns_tab)
            
            # Show success message (same as traditional exports)
            self.root.after(0, lambda: messagebox.showinfo(
                "Export Successful", 
                f"Report exported successfully!\n\nColumns: {col_count}\nRows: {row_count}\n\nNow go to the 'Select Columns' tab to customize your output."))
            
        except Exception as e:
            self.log_message(f"Error processing direct API data: {str(e)}", 'error')
            raise

    def transform_group_types(self, df):
        """Transform groupTypes array to user-friendly strings"""
        try:
            def convert_group_type(group_types):
                """Convert groupTypes array to readable string"""
                if not group_types or group_types == [] or group_types is None:
                    return "Security"
                
                # Array as string format
                if isinstance(group_types, str):
                    try:
                        import ast
                        group_types = ast.literal_eval(group_types)
                    except:
                        # If it's already a simple string, check its content
                        if 'Unified' in group_types:
                            return "Microsoft 365"
                        elif 'DynamicMembership' in group_types:
                            return "Dynamic Membership"
                        else:
                            return "Security"
                
                # Array format
                if isinstance(group_types, list):
                    if not group_types:  # Empty array
                        return "Security"
                    elif 'Unified' in group_types and 'DynamicMembership' in group_types:
                        return "Microsoft 365 (Dynamic)"
                    elif 'Unified' in group_types:
                        return "Microsoft 365"
                    elif 'DynamicMembership' in group_types:
                        return "Dynamic Membership"
                    else:
                        return "Security"
                
                # Fallback
                return "Security"
            
            # Apply transformation to groupTypes column
            df['groupTypes'] = df['groupTypes'].apply(convert_group_type)
            self.log_message("groupTypes transformation completed", 'debug')
            return df
            
        except Exception as e:
            self.log_message(f"Error transforming groupTypes: {str(e)}", 'warning')
            # Return original DataFrame if transformation fails
            return df

    def debug_token_permissions(self):
        """Debug method to check what permissions the current token has"""
        try:
            import base64
            import json
            
            if not self.access_token:
                self.log_message("No access token available for permission debugging", 'warning')
                return
            
            # Decode the JWT token to see the scopes (just for debugging)
            # Note: This is just for debugging, don't use in production
            token_parts = self.access_token.split('.')
            if len(token_parts) >= 2:
                # Add padding if needed
                payload = token_parts[1]
                payload += '=' * (4 - len(payload) % 4)
                
                try:
                    decoded_payload = base64.b64decode(payload)
                    token_data = json.loads(decoded_payload)
                    
                    # Log relevant token information
                    if 'scp' in token_data:
                        self.log_message(f"Token delegated scopes (scp): {token_data['scp']}", 'debug')
                    if 'roles' in token_data:
                        self.log_message(f"Token application roles: {token_data['roles']}", 'debug')
                    if 'aud' in token_data:
                        self.log_message(f"Token audience: {token_data['aud']}", 'debug')
                    if 'appid' in token_data:
                        self.log_message(f"Token app ID: {token_data['appid']}", 'debug')
                        
                except Exception as decode_error:
                    self.log_message(f"Could not decode token for debugging: {str(decode_error)}", 'debug')
            
        except Exception as e:
            self.log_message(f"Error in token debugging: {str(e)}", 'debug')

    def export_thread(self, report_name):
        """Export thread using smart parameter system"""
        try:
            self.log_message(f"Export thread started for report: {report_name}", 'debug')
            
            # Use smart parameter system to build request body
            request_body = self.get_report_parameters(report_name)
            
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Content-Type': 'application/json'
            }
            
            self.log_message(f"API Endpoint: {self.export_endpoint}", 'api')
            self.log_message(f"Request Body: {json.dumps(request_body, indent=2)}", 'debug')
            self.log_message(f"Headers prepared (token length: {len(self.access_token) if self.access_token else 0})", 'debug')
            
            # Step 1: Create export job with POST
            self.root.after(0, lambda: self.progress_label.config(text="Creating export job..."))
            self.log_message("Step 1: Creating export job with POST request", 'info')
            
            try:
                self.log_message(f"Making POST request to: {self.export_endpoint}", 'debug')
                response = self.make_authenticated_request('POST', self.export_endpoint, 
                                                         operation_type='export_job_creation',
                                                         json=request_body)
                
                self.log_message(f"POST Request completed", 'debug')
                self.log_message(f"POST Response Status: {response.status_code}", 'api')
                self.log_message(f"POST Response Headers: {dict(response.headers)}", 'debug')
                self.log_message(f"POST Response Body: {response.text}", 'debug')
                
                # Handle 400 errors with smart retry
                if response.status_code == 400:
                    error_detail = self.parse_error_response(response)
                    if "ReportTypeNotFlighted" in str(error_detail):
                        raise Exception(f"Report '{report_name}' is not available in your tenant. Try a different report like 'DevicesWithInventory' or 'Devices'.")
                    elif ("required filters" in str(error_detail).lower() or 
                          "required parameter" in str(error_detail).lower() or
                          "could not find a property named" in str(error_detail).lower()):
                        # Try to learn from the error and retry once
                        self.log_message(f"API parameter error detected. Attempting smart retry...", 'warning')
                        
                        learned_params = self.learn_from_error(report_name, response.json() if response.text else {})
                        if learned_params:
                            # Retry with learned parameters (regardless of whether it's in report_parameters)
                            self.log_message(f"Retrying {report_name} with learned parameters: {list(learned_params.keys())}", 'info')
                            self.root.after(0, lambda: self.progress_label.config(text="Retrying with corrected parameters..."))
                            
                            # Build new request with learned parameters
                            retry_body = self.get_report_parameters(report_name)  # This will now include learned params
                            self.log_message(f"Retry Request Body: {json.dumps(retry_body, indent=2)}", 'debug')
                            
                            retry_response = self.make_authenticated_request('POST', self.export_endpoint, 
                                                                           operation_type='export_job_creation',
                                                                           json=retry_body)
                            
                            if retry_response.status_code in [200, 201]:
                                self.log_message(f"Smart retry successful!", 'success')
                                response = retry_response  # Use retry response for processing
                            else:
                                retry_error = self.parse_error_response(retry_response)
                                raise Exception(f"Smart retry failed: {retry_error}. Original error: {error_detail}")
                        else:
                            raise Exception(f"Could not determine correct parameters for {report_name}: {error_detail}")
                    else:
                        raise Exception(f"Bad Request: {error_detail}")
                
            except requests.exceptions.Timeout:
                raise Exception("Request timeout - API took too long to respond")
            except requests.exceptions.ConnectionError as e:
                raise Exception(f"Connection error - Unable to reach Microsoft Graph API: {str(e)}")
            except requests.exceptions.RequestException as e:
                raise Exception(f"Request failed: {str(e)}")

            # Process successful response (either original or retry)
            if response.status_code in [200, 201]:
                try:
                    export_job = response.json()
                    self.log_message(f"Export job response: {json.dumps(export_job, indent=2)}", 'debug')
                except json.JSONDecodeError:
                    raise Exception(f"Invalid JSON response from API: {response.text}")
                
                job_id = export_job.get('id')
                
                if not job_id:
                    raise Exception(f"No job ID returned from export request. Response: {export_job}")
                
                self.log_message(f"Export job created successfully with ID: {job_id}", 'success')
                
                # Step 2: Poll for job completion using GET with job_id in single quotes
                self.root.after(0, lambda: self.progress_label.config(text="Waiting for job completion..."))
                self.log_message("Step 2: Polling for job completion", 'info')
                
                job_status_url = f"{self.export_endpoint}('{job_id}')"
                self.log_message(f"Status check URL: {job_status_url}", 'api')
                
                # Adaptive timeout based on export job type and potential data size
                estimated_records = self.estimate_export_size(report_name)
                max_wait_time = self.timeout_manager.get_timeout_for_operation('large_export', estimated_records)
                poll_interval = min(10, max(5, max_wait_time // 60))  # Adaptive polling: 5-10 seconds
                elapsed_time = 0
                
                self.log_message(f"Export job timeout set to {max_wait_time} seconds (estimated {estimated_records} records)", 'info')
                
                while elapsed_time < max_wait_time:
                    self.log_message(f"Sleeping for {poll_interval} seconds before status check", 'debug')
                    time.sleep(poll_interval)
                    elapsed_time += poll_interval
                    
                    # Calculate percentage completion based on elapsed time
                    progress_percentage = min(int((elapsed_time / max_wait_time) * 100), 100)
                    self.root.after(0, lambda p=progress_percentage: self.progress_label.config(
                        text=f"Checking job status... ({p}% complete)"))
                    
                    # GET request to check status
                    try:
                        self.log_message(f"Making GET request to check status", 'debug')
                        status_response = self.make_authenticated_request('GET', job_status_url, 
                                                                          operation_type='export_job_status')
                        self.log_message(f"GET Status Response: {status_response.status_code}", 'api')
                        self.log_message(f"GET Status Body: {status_response.text}", 'debug')
                    except requests.exceptions.RequestException as e:
                        self.log_message(f"Status check request failed: {str(e)}", 'warning')
                        continue
                    
                    if status_response.status_code == 200:
                        try:
                            job_status = status_response.json()
                            status = job_status.get('status', 'unknown')
                            self.log_message(f"Job status: {status}", 'info')
                            self.log_message(f"Full status response: {json.dumps(job_status, indent=2)}", 'debug')
                        except json.JSONDecodeError:
                            self.log_message(f"Invalid JSON in status response: {status_response.text}", 'warning')
                            continue
                        
                        if status.lower() == 'completed':
                            download_url = job_status.get('url')
                            if download_url:
                                self.log_message("Export job completed successfully!", 'success')
                                self.log_message(f"Download URL: {download_url}", 'api')
                                
                                # Step 3: Download and process the CSV
                                self.root.after(0, lambda: self.progress_label.config(text="Downloading and processing CSV..."))
                                self.log_message("Step 3: Downloading and processing CSV data", 'info')
                                
                                try:
                                    # The download URL from Microsoft is pre-authenticated and doesn't need our Bearer token
                                    # Use direct requests call instead of our authenticated method
                                    timeout = self.timeout_manager.get_timeout_for_operation('file_download')
                                    self.log_message(f"Downloading from pre-authenticated URL (timeout: {timeout}s)", 'info')
                                    
                                    download_response = requests.get(download_url, timeout=timeout)
                                    
                                    self.log_message(f"Download response status: {download_response.status_code}", 'api')
                                    
                                    if download_response.status_code == 200:
                                        self.log_message(f"Downloaded content size: {len(download_response.content)} bytes", 'info')
                                    else:
                                        raise Exception(f"Download failed: HTTP {download_response.status_code}")
                                        
                                except requests.exceptions.RequestException as e:
                                    raise Exception(f"Failed to download export file: {str(e)}")
                                except Exception as e:
                                    raise Exception(f"Failed to download export file: {str(e)}")
                                
                                if download_response.status_code == 200:
                                    # Process the downloaded content
                                    self.process_downloaded_content(download_response.content, report_name)
                                    break
                                else:
                                    raise Exception(f"Failed to download file: HTTP {download_response.status_code}")
                            else:
                                raise Exception("No download URL provided in completed job")
                                
                        elif status.lower() in ['failed', 'cancelled', 'error']:
                            error_msg = job_status.get('errorMessage', job_status.get('message', 'Unknown error'))
                            raise Exception(f"Export job {status}: {error_msg}")
                        
                        elif status.lower() in ['running', 'queued', 'inprogress']:
                            self.log_message(f"Job still {status}, continuing to wait...", 'info')
                            continue
                            
                    else:
                        self.log_message(f"Failed to check job status: HTTP {status_response.status_code}", 'warning')
                        self.log_message(f"Status response body: {status_response.text}", 'debug')
                        
                else:
                    raise Exception(f"Export job timeout after {max_wait_time} seconds")
            elif response.status_code == 401:
                raise Exception("Authentication failed - Token may be expired. Please logout and login again.")
            elif response.status_code == 403:
                error_detail = self.parse_error_response(response)
                raise Exception(f"Insufficient permissions: {error_detail}")
            else:
                error_detail = self.parse_error_response(response)
                raise Exception(f"API Error (HTTP {response.status_code}): {error_detail}")
                
        except Exception as e:
            self.log_message(f"Export thread failed: {str(e)}", 'error')
            self.log_message(f"Error trace: {traceback.format_exc()}", 'debug')
            
            # Capture error message before lambda to avoid scoping issues
            error_msg = str(e)
            self.root.after(0, lambda: messagebox.showerror(
                "Export Failed", f"Failed to export report:\n\n{error_msg}"))
        
        finally:
            # Re-enable export button and stop progress
            self.log_message("Export thread finishing - restoring UI state", 'debug')
            self.root.after(0, lambda: self.export_btn.config(state='normal'))
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.progress_label.config(text="Ready"))
    
    def process_downloaded_content(self, content, report_name):
        """Process downloaded content and prepare for column selection"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Check if it's a ZIP file
            if content.startswith(b'PK') or b'PK\x03\x04' in content[:10]:
                self.log_message("Processing ZIP file", 'info')
                
                # Save ZIP temporarily
                temp_zip = f"temp_{report_name}_{timestamp}.zip"
                with open(temp_zip, 'wb') as f:
                    f.write(content)
                
                # Extract CSV from ZIP
                csv_content = None
                with zipfile.ZipFile(temp_zip, 'r') as zip_ref:
                    file_list = zip_ref.namelist()
                    self.log_message(f"Files in ZIP: {file_list}", 'debug')
                    
                    # Find CSV file
                    csv_files = [f for f in file_list if f.lower().endswith('.csv')]
                    if csv_files:
                        csv_filename = csv_files[0]
                        csv_content = zip_ref.read(csv_filename).decode('utf-8-sig')
                        self.log_message(f"Extracted CSV: {csv_filename}", 'success')
                    else:
                        raise Exception("No CSV file found in ZIP")
                
                # Clean up temp file
                os.remove(temp_zip)
                
            else:
                # Direct CSV content
                self.log_message("Processing direct CSV content", 'info')
                csv_content = content.decode('utf-8-sig')
            
            if csv_content:
                # Parse CSV to get columns and data
                import io
                csv_reader = csv.DictReader(io.StringIO(csv_content))
                
                # Get column names
                self.current_columns = list(csv_reader.fieldnames)
                
                # Read all data
                self.current_export_data = list(csv_reader)
                original_data_count = len(self.current_export_data)
                
                # Apply filtering if this is a filtered export (not full export)
                if hasattr(self, 'current_export_is_filtered') and self.current_export_is_filtered:
                    self.log_message(f"Applying post-processing filters to export job data...", 'info')
                    
                    # Convert to DataFrame for filtering
                    import pandas as pd
                    df = pd.DataFrame(self.current_export_data)
                    
                    # Apply the same filtering logic as direct API reports
                    filtered_df = self.apply_post_processing_filters(df, report_name)
                    
                    # Convert back to list of dictionaries
                    if filtered_df is not None and not filtered_df.empty:
                        self.current_export_data = filtered_df.to_dict('records')
                        filtered_count = len(self.current_export_data)
                        self.log_message(f"Filtering applied: {original_data_count} -> {filtered_count} rows", 'info')
                    else:
                        self.log_message(f"Filtering resulted in empty dataset, keeping original {original_data_count} rows", 'warning')
                        # Keep original data as fallback
                
                rows_count = len(self.current_export_data)
                cols_count = len(self.current_columns)
                
                self.log_message(f"CSV processed successfully!", 'success')
                self.log_message(f"Columns: {cols_count}, Rows: {rows_count}", 'success')
                self.log_message(f"Available columns: {', '.join(self.current_columns[:10])}{'...' if cols_count > 10 else ''}", 'info')
                
                # Store the raw CSV for later use
                self.raw_csv_content = csv_content
                
                # Validate data
                if not self.current_columns:
                    raise Exception("No columns found in CSV data")
                    
                if not self.current_export_data:
                    raise Exception("No data rows found in CSV")
                    
                # Check for duplicate columns
                duplicate_cols = [col for col in self.current_columns if self.current_columns.count(col) > 1]
                if duplicate_cols:
                    self.log_message(f"Warning: Duplicate columns found: {set(duplicate_cols)}", 'warning')
                
                # Create column selection tab
                self.root.after(0, self.create_columns_tab)
                
                # Show success message
                self.root.after(0, lambda: messagebox.showinfo(
                    "Export Successful", 
                    f"Report exported successfully!\n\nColumns: {cols_count}\nRows: {rows_count}\n\nNow go to the 'Select Columns' tab to customize your output."))
            else:
                raise Exception("No CSV content found")
                
        except Exception as e:
            raise Exception(f"Failed to process downloaded content: {str(e)}")
    
    def log_message(self, message, tag='info'):
        """Log message to console with timestamp and color"""
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        
        prefixes = {
            'success': '[SUCCESS]',
            'error': '[ERROR]', 
            'warning': '[WARNING]',
            'info': '[INFO]',
            'api': '[API]',
            'debug': '[DEBUG]'
        }
        
        prefix = prefixes.get(tag, '[INFO]')
        log_message = f"[{timestamp}] {prefix} {message}\n"
        
        def update():
            # Check if console exists and is still valid before trying to log
            try:
                if (hasattr(self, 'console_text') and self.console_text and 
                    self.console_text.winfo_exists()):
                    self.console_text.insert(tk.END, log_message, tag)
                    if hasattr(self, 'auto_scroll') and self.auto_scroll.get():
                        self.console_text.see(tk.END)
                else:
                    # Fallback to print if console not available
                    print(log_message.strip())
            except (tk.TclError, AttributeError):
                # Widget was destroyed or invalid, fallback to print
                print(log_message.strip())
        
        if hasattr(self, 'root'):
            self.root.after(0, update)
        else:
            update()
    
    def clear_console(self):
        """Clear console log"""
        self.console_text.delete(1.0, tk.END)
        self.log_message("Console cleared", 'info')
    
    def save_log(self):
        """Save console log to file"""
        content = self.console_text.get(1.0, tk.END)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"intune_export_log_{timestamp}.log"
        
        filename = filedialog.asksaveasfilename(
            initialfile=default_filename,
            defaultextension=".log",
            filetypes=[("Log files", "*.log"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(f"Microsoft Intune Reports Export Tool v1.0 - Log File\n")
                    f.write(f"Generated: {datetime.now().isoformat()}\n")
                    f.write(f"User: {self.user_info.get('displayName', 'Unknown') if self.user_info else 'Unknown'}\n")
                    f.write("=" * 80 + "\n\n")
                    f.write(content)
                
                self.log_message(f"Log saved: {filename}", 'success')
                messagebox.showinfo("Log Saved", f"Log saved successfully to:\n{filename}")
            except Exception as e:
                self.log_message(f"Failed to save log: {str(e)}", 'error')
                messagebox.showerror("Save Error", f"Failed to save log file:\n{str(e)}")
    
    def open_feedback_form(self):
        """Open the Microsoft Form for tool feedback"""
        feedback_url = "https://htmd.in/toolfeedback"
        try:
            webbrowser.open(feedback_url)
            self.log_message("Feedback form opened in browser", 'info')
        except Exception as e:
            self.log_message(f"Failed to open feedback form: {str(e)}", 'error')
            messagebox.showerror("Error", f"Failed to open feedback form:\n{str(e)}")
    
    def export_to_powerbi_old_complex_version(self):
        """Export selected columns to PowerBI Desktop and open directly"""
        # Get selected columns
        selected_columns = []
        for column, var in self.column_vars.items():
            if var.get():
                selected_columns.append(column)
        
        if not selected_columns:
            messagebox.showwarning("No Columns Selected", "Please select at least one column")
            return
        
        try:
            # Update UI
            self.export_status.config(text="Opening Power BI...", fg='#F2C811')
            self.powerbi_btn.config(state='disabled', text="⚡ Opening...")
            self.root.update()
            
            import subprocess
            import platform
            
            # Create CSV in user's Documents folder for easy access
            documents_path = os.path.expanduser("~/Documents")
            powerbi_folder = os.path.join(documents_path, "PowerBI_Imports")
            
            # Create folder if it doesn't exist
            os.makedirs(powerbi_folder, exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_filename = f"Intune_{self.selected_report.get()}_{timestamp}.csv"
            csv_path = os.path.join(powerbi_folder, csv_filename)
            
            # Write CSV file with error checking
            with open(csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=selected_columns)
                writer.writeheader()
                
                for row in self.current_export_data:
                    filtered_row = {col: row.get(col, '') for col in selected_columns}
                    writer.writerow(filtered_row)
            
            # Verify file was created successfully
            if not os.path.exists(csv_path):
                raise Exception(f"Failed to create CSV file at {csv_path}")
            
            file_size = os.path.getsize(csv_path)
            if file_size == 0:
                raise Exception("CSV file was created but is empty")
                
            self.log_message(f"CSV created successfully: {csv_path} ({file_size} bytes)", 'success')
            
            # Create PowerBI template for automated data loading
            pbit_path = self.create_powerbi_template(csv_path, selected_columns)
            
            # Open PowerBI with automated data import
            if platform.system() == "Windows":
                # Try to find PowerBI Desktop and launch with template
                powerbi_paths = [
                    r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe",
                    r"C:\Program Files (x86)\Microsoft Power BI Desktop\bin\PBIDesktop.exe",
                    os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WindowsApps\PBIDesktop.exe")
                ]
                
                powerbi_found = False
                for pbi_path in powerbi_paths:
                    if os.path.exists(pbi_path):
                        try:
                            if pbit_path and os.path.exists(pbit_path):
                                # Launch PowerBI with template file for automatic data loading
                                subprocess.Popen([pbi_path, pbit_path], shell=False)
                                powerbi_found = True
                                self.log_message(f"PowerBI Desktop launched with template: {pbit_path}", 'success')
                            else:
                                # Fallback: Launch PowerBI and try automated CSV import
                                subprocess.Popen([pbi_path], shell=False)
                                powerbi_found = True
                                self.log_message(f"PowerBI Desktop launched: {pbi_path}", 'success')
                                
                                # Try to automate CSV import using Windows automation
                                self.root.after(4000, lambda: self.automate_csv_import(csv_path))
                            break
                            
                        except Exception as e:
                            self.log_message(f"Failed to launch PowerBI: {str(e)}", 'warning')
                            continue
                
                if not powerbi_found:
                    # PowerBI not found, try to open CSV with default handler
                    self.log_message("PowerBI Desktop not found, opening CSV with default program", 'warning')
                    os.startfile(csv_path)
                    self.show_powerbi_not_found_message(csv_path, selected_columns)
                else:
                    # Success - show folder location
                    self.export_status.config(text="✅ Opened in Power BI!", fg='#107c10')
                    
                    # Show success message with instructions
                    messagebox.showinfo(
                        "Power BI Export Success", 
                        f"✅ Data exported successfully!\n\n"
                        f"📁 File: {csv_filename}\n"
                        f"📊 Columns: {len(selected_columns)}\n"
                        f"📊 Rows: {len(self.current_export_data)}\n\n"
                        f"🚀 Power BI Desktop is starting...\n"
                        f"🤖 Attempting automated data import...\n"
                        f"� CSV path copied to clipboard\n\n"
                        f"If automation doesn't work:\n"
                        f"1. In Power BI: Home → Get Data → Text/CSV\n"
                        f"2. Press Ctrl+V to paste file path\n"
                        f"3. Click 'Load' or 'Transform Data'"
                    )
            else:
                # Non-Windows systems
                os.system(f'open "{csv_path}"')  # macOS
                messagebox.showinfo("CSV Export", f"CSV file created:\n{csv_path}\n\nOpen this file in Power BI Desktop")
            
        except Exception as e:
            self.log_message(f"PowerBI export failed: {str(e)}", 'error')
            messagebox.showerror("Export Failed", f"Failed to export to Power BI:\n\n{str(e)}")
            self.export_status.config(text="❌ Export failed", fg='#d13438')
        finally:
            self.powerbi_btn.config(state='normal', text="⚡ Open in Power BI")
    
    def create_powerbi_template(self, csv_path, selected_columns):
        """Create a PowerBI template file for automated CSV import"""
        try:
            import json
            import zipfile
            import tempfile
            
            # PowerBI template folder
            template_folder = os.path.dirname(csv_path)
            template_name = f"Intune_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pbit"
            template_path = os.path.join(template_folder, template_name)
            
            # Create a basic PowerBI template JSON structure
            # This is a simplified version - PowerBI templates are complex
            template_data = {
                "version": "1.0",
                "dataModel": {
                    "tables": [
                        {
                            "name": "IntuneData",
                            "columns": [{"name": col, "dataType": "string"} for col in selected_columns]
                        }
                    ]
                },
                "mashup": f"""
let
    Source = Csv.Document(File.Contents("{csv_path.replace(chr(92), '/')}"), [Delimiter=",", Columns={len(selected_columns)}, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    PromotedHeaders
"""
            }
            
            # Note: This is a simplified approach. Real PBIT files are more complex.
            # For now, we'll use the CSV automation approach instead.
            return None
            
        except Exception as e:
            self.log_message(f"Failed to create PowerBI template: {str(e)}", 'warning')
            return None
    
    def automate_csv_import(self, csv_path):
        """Attempt to automate CSV import in PowerBI using various methods"""
        try:
            import time
            
            # Method 1: Copy CSV path to clipboard for easy access
            import subprocess
            subprocess.run(['powershell', '-command', f'Set-Clipboard -Value "{csv_path}"'], check=True, capture_output=True)
            self.log_message(f"CSV path copied to clipboard: {csv_path}", 'success')
            
            # Method 2: Create a PowerBI script file for automation
            script_content = f'''
# PowerBI M Query Script for automatic CSV import
let
    Source = Csv.Document(File.Contents("{csv_path.replace(chr(92), '/')}"), 
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    PromotedHeaders
'''
            
            script_path = os.path.join(os.path.dirname(csv_path), "PowerBI_Import_Script.txt")
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(script_content)
            
            self.log_message(f"PowerBI M script created: {script_path}", 'success')
            
            # Method 3: Try Windows automation if available
            try:
                import pyautogui
                
                # Small delay to ensure PowerBI is ready
                time.sleep(2)
                
                # Send keyboard shortcuts for Get Data
                pyautogui.hotkey('ctrl', 'shift', 'g')  # Get Data shortcut in PowerBI
                time.sleep(1)
                
                # Type to search for CSV option
                pyautogui.typewrite('csv')
                time.sleep(0.5)
                pyautogui.press('enter')
                
                # Wait for file dialog
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'v')  # Paste file path
                time.sleep(0.5)
                pyautogui.press('enter')
                
                # Click Load button (assuming it's the default)
                time.sleep(1)
                pyautogui.press('enter')
                
                self.log_message("PowerBI automation attempted successfully", 'success')
                
            except ImportError:
                self.log_message("PyAutoGUI not available - clipboard method used instead", 'info')
                # Offer to install automation library
                if messagebox.askyesno("Enhanced Automation", 
                    "Want full PowerBI automation?\n\n"
                    "Install pyautogui for automatic GUI control?\n"
                    "(This will enable complete hands-free PowerBI data import)"):
                    self.install_automation_library()
            except Exception as automation_error:
                self.log_message(f"PowerBI automation failed: {str(automation_error)}", 'warning')
                
        except Exception as e:
            self.log_message(f"CSV import automation failed: {str(e)}", 'warning')
    
    def install_automation_library(self):
        """Install pyautogui for PowerBI automation"""
        try:
            import subprocess
            import sys
            
            self.log_message("Installing PyAutoGUI for enhanced automation...", 'info')
            
            # Run pip install in a separate thread to avoid blocking UI
            def install_thread():
                try:
                    # Use a more compatible approach for Python 3.13
                    import os
                    result = os.system(f'"{sys.executable}" -m pip install pyautogui')
                    if result == 0:
                        self.log_message("PyAutoGUI installed successfully!", 'success')
                        messagebox.showinfo("Installation Complete", 
                            "PyAutoGUI installed successfully!\n\n"
                            "Full PowerBI automation will be available on next export.")
                    else:
                        self.log_message(f"Installation failed with exit code: {result}", 'error')
                        messagebox.showerror("Installation Failed", 
                            f"Failed to install PyAutoGUI.\nTry installing manually: pip install pyautogui")
                except Exception as e:
                    self.log_message(f"Installation error: {str(e)}", 'error')
                    messagebox.showerror("Installation Error", 
                        f"Error installing PyAutoGUI:\n{str(e)}")
            
            import threading
            thread = threading.Thread(target=install_thread)
            thread.daemon = True
            thread.start()
            
        except Exception as e:
            self.log_message(f"Failed to start installation: {str(e)}", 'error')
    
    def open_powerbi_simple(self):
        """Simple PowerBI launcher - just opens PowerBI with instructions"""
        try:
            # Update UI
            self.export_status.config(text="Opening Power BI...", fg='#F2C811')
            self.powerbi_btn.config(state='disabled', text="📊 Opening...")
            self.root.update_idletasks()  # Use update_idletasks instead of update to prevent zoom issues
            
            import subprocess
            import platform
            
            if platform.system() == "Windows":
                # Try to find PowerBI Desktop
                powerbi_paths = [
                    r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe",
                    r"C:\Program Files (x86)\Microsoft Power BI Desktop\bin\PBIDesktop.exe",
                    os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WindowsApps\PBIDesktop.exe")
                ]
                
                powerbi_found = False
                for pbi_path in powerbi_paths:
                    if os.path.exists(pbi_path):
                        try:
                            # Launch PowerBI Desktop
                            subprocess.Popen([pbi_path], shell=False)
                            powerbi_found = True
                            self.log_message(f"PowerBI Desktop launched: {pbi_path}", 'success')
                            break
                        except Exception as e:
                            self.log_message(f"Failed to launch PowerBI: {str(e)}", 'warning')
                            continue
                
                if not powerbi_found:
                    # PowerBI not found
                    self.log_message("PowerBI Desktop not found", 'warning')
                    self.export_status.config(text="❌ PowerBI not found", fg='#d13438')
                    messagebox.showerror("PowerBI Not Found", 
                        "PowerBI Desktop not found on your system.\n\n"
                        "Please install PowerBI Desktop from:\n"
                        "https://powerbi.microsoft.com/desktop/")
                    return
                
                # Show simple instructions
                self.export_status.config(text="✅ PowerBI opened!", fg='#107c10')
                
                messagebox.showinfo(
                    "PowerBI Opened", 
                    "📊 PowerBI Desktop is now opening!\n\n"
                    "📋 To import your data:\n\n"
                    "1️⃣ First: Click 'Export CSV' button to save your data\n"
                    "2️⃣ In PowerBI: Home → Get Data → Text/CSV\n"
                    "3️⃣ Browse to your CSV file and select it\n"
                    "4️⃣ Click 'Load' or 'Transform Data'\n\n"
                    "💡 Tip: CSV files are saved to your Downloads folder by default"
                )
            else:
                messagebox.showinfo("PowerBI", 
                    "Please install PowerBI Desktop and manually import your CSV file.")
                    
        except Exception as e:
            self.log_message(f"PowerBI launch failed: {str(e)}", 'error')
            messagebox.showerror("Launch Failed", f"Failed to launch PowerBI:\n\n{str(e)}")
            self.export_status.config(text="❌ Launch failed", fg='#d13438')
        finally:
            self.powerbi_btn.config(state='normal', text="📊 Open Power BI")
    
    def get_odata_feed(self):
        """Get OData feed URL - copies the standard Data Warehouse OData URL to clipboard"""
        try:
            # Check if a report is selected
            if not self.selected_report.get():
                messagebox.showwarning("No Report Selected", "Please select a report first")
                return
            
            # Update UI
            self.export_status.config(text="Copying OData feed URL...", fg='#0078d4')
            self.odata_btn.config(state='disabled', text="🔗 Copying...")
            self.root.update_idletasks()
            
            # Get the selected report
            report_name = self.selected_report.get()
            self.log_message(f"Getting OData feed URL for report: {report_name}", 'info')
            
            # Standard OData feed URL that works for all Intune Data Warehouse reports
            odata_url = "https://fef.msua08.manage.microsoft.com/ReportingService/DataWarehouseFEService?api-version=v1.0"
            
            # Copy to clipboard using PowerShell
            import subprocess
            subprocess.run([
                'powershell', '-Command', 
                f'Set-Clipboard -Value "{odata_url}"'
            ], check=True, capture_output=True)
            
            self.log_message(f"OData URL copied to clipboard: {odata_url}", 'info')
            
            # Show success message with instructions
            messagebox.showinfo("OData Feed URL Copied", 
                              f"✅ OData feed URL copied to clipboard!\n\n"
                              f"📋 URL: {odata_url}\n\n"
                              f"📊 This URL works for all Intune Data Warehouse reports including:\n"
                              f"• {report_name}\n"
                              f"• All other available reports\n\n"
                              f"🔗 You can now paste this URL into:\n"
                              f"• Power BI (Get Data > OData feed)\n"
                              f"• Excel (Data > From Web > OData)\n"
                              f"• Any OData-compatible tool\n\n"
                              f"ℹ️ You may need to authenticate with your Microsoft credentials when using the URL.")
            
            self.export_status.config(text="✅ OData URL copied to clipboard!", fg='#107c10')
            
        except subprocess.CalledProcessError as e:
            # If clipboard copy fails, still show the URL
            error_message = f"Could not copy to clipboard: {e}"
            self.log_message(f"Clipboard error: {error_message}", 'warning')
            
            messagebox.showinfo("OData Feed URL", 
                              f"📋 OData feed URL for all Intune Data Warehouse reports:\n\n"
                              f"{odata_url}\n\n"
                              f"⚠️ {error_message}\n\n"
                              f"Please copy this URL manually to use in Power BI, Excel, or other OData tools.")
            
            self.export_status.config(text="✅ OData URL provided (copy manually)", fg='#107c10')
            
        except Exception as e:
            error_message = f"Error getting OData feed: {str(e)}"
            self.log_message(f"OData feed error: {error_message}", 'error')
            
            messagebox.showerror("OData Feed Error", error_message)
            self.export_status.config(text="❌ OData feed failed", fg='#d13438')
            
        finally:
            self.odata_btn.config(state='normal', text="🔗 Get OData Feed")
    
    def parse_api_error(self, response):
        """Parse API error response and provide detailed guidance"""
        status_code = response.status_code
        error_message = "Unknown error occurred"
        guidance = ""
        
        try:
            # Try to parse JSON error response
            error_data = response.json()
            if 'error' in error_data:
                error_info = error_data['error']
                error_message = error_info.get('message', 'API request failed')
                error_code = error_info.get('code', 'UnknownError')
                
                # Provide specific guidance based on error type
                if status_code == 401:
                    guidance = """🔐 AUTHENTICATION ERROR:
Your access token has expired or is invalid.

📋 Required Actions:
1. Click 'Logout' in the application
2. Click 'Login' to re-authenticate
3. Try the OData request again

🛡️ If problem persists:
• Check with your IT administrator
• Verify your account has Intune access"""

                elif status_code == 403:
                    guidance = """🚫 PERMISSION ERROR:
Your account lacks required permissions for this operation.

📋 Required Microsoft Graph Permissions:
• DeviceManagementApps.Read.All
• DeviceManagementConfiguration.Read.All  
• DeviceManagementManagedDevices.Read.All
• DeviceManagementRBAC.Read.All

👤 Contact your IT Administrator to:
1. Grant the above permissions to your account
2. Or add you to appropriate Intune admin roles:
   • Intune Service Administrator
   • Global Reader
   • Reports Reader"""

                elif status_code == 404:
                    guidance = """❓ RESOURCE NOT FOUND:
The requested report or endpoint was not found.

📋 Possible Causes:
1. Report name might be incorrect
2. Endpoint URL might have changed
3. Report might not be available in your tenant

🔧 Troubleshooting:
• Try a different report from the list
• Check if the report exists in Intune portal
• Contact support if issue persists"""

                elif status_code == 429:
                    guidance = """⏰ RATE LIMIT EXCEEDED:
Too many requests sent to Microsoft Graph API.

📋 Required Actions:
1. Wait 1-2 minutes before trying again
2. Reduce frequency of API calls
3. Try again later during off-peak hours"""

                elif status_code >= 500:
                    guidance = """🔧 SERVER ERROR:
Microsoft Graph API is experiencing issues.

📋 Recommended Actions:
1. Wait 5-10 minutes and try again
2. Check Microsoft 365 Service Health
3. Try a different report
4. Contact Microsoft support if persistent"""

                else:
                    guidance = f"""❌ API ERROR (HTTP {status_code}):
An unexpected error occurred.

📋 Troubleshooting Steps:
1. Try logging out and logging back in
2. Try a different report
3. Check your internet connection
4. Contact your IT administrator

🔍 Error Code: {error_code}"""

            else:
                error_message = response.text[:300] if response.text else f"HTTP {status_code} Error"
                
        except (ValueError, KeyError):
            # Not valid JSON or missing expected fields
            error_message = response.text[:300] if response.text else f"HTTP {status_code} Error"
            
            if status_code == 401:
                guidance = "🔐 Authentication failed. Please logout and login again."
            elif status_code == 403:
                guidance = "🚫 Access denied. Contact your IT administrator for Intune permissions."
            else:
                guidance = f"❌ HTTP {status_code} error. Check your connection and try again."
        
        return {
            'message': error_message,
            'guidance': guidance,
            'status_code': status_code
        }
    
    def show_detailed_error_dialog(self, title, error_message):
        """Show detailed error dialog with guidance"""
        # Create a custom dialog with scrollable text
        error_window = tk.Toplevel(self.root)
        error_window.title(title)
        error_window.geometry("600x400")
        error_window.configure(bg='#f0f0f0')
        
        # Make it modal
        error_window.transient(self.root)
        error_window.grab_set()
        
        # Center the window
        error_window.update_idletasks()
        x = (error_window.winfo_screenwidth() // 2) - (600 // 2)
        y = (error_window.winfo_screenheight() // 2) - (400 // 2)
        error_window.geometry(f"600x400+{x}+{y}")
        
        # Header
        header_frame = tk.Frame(error_window, bg='#d13438', height=50)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="❌ " + title, 
                font=('Segoe UI', 14, 'bold'), 
                bg='#d13438', fg='white').pack(pady=15)
        
        # Scrollable text area
        text_frame = tk.Frame(error_window)
        text_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side='right', fill='y')
        
        text_widget = tk.Text(text_frame, wrap='word', yscrollcommand=scrollbar.set,
                             font=('Segoe UI', 10), bg='white', padx=15, pady=15)
        text_widget.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=text_widget.yview)
        
        # Insert error message
        text_widget.insert('1.0', error_message)
        text_widget.config(state='disabled')  # Make read-only
        
        # Close button
        button_frame = tk.Frame(error_window, bg='#f0f0f0')
        button_frame.pack(pady=10)
        
        close_btn = tk.Button(button_frame, text="Close", 
                             command=error_window.destroy,
                             font=('Segoe UI', 10, 'bold'),
                             bg='#0078d4', fg='white', 
                             padx=20, pady=8, relief='flat',
                             cursor='hand2')
        close_btn.pack()
        
        # Focus on close button
        close_btn.focus_set()
        error_window.bind('<Return>', lambda e: error_window.destroy())
        error_window.bind('<Escape>', lambda e: error_window.destroy())
    
    def show_odata_info(self, report_name, odata_url, download_url, job_id):
        """Show OData feed information to the user"""
        # Copy URLs to clipboard
        import subprocess
        clipboard_content = f"OData URL: {odata_url}\nDownload URL: {download_url}\nJob ID: {job_id}"
        
        try:
            subprocess.run(['powershell', '-command', f'Set-Clipboard -Value @"\\n{clipboard_content}\\n"@'], 
                         check=True, capture_output=True)
            clipboard_msg = "📋 URLs copied to clipboard!"
        except:
            clipboard_msg = "📋 Copy URLs manually from below"
        
        # Show detailed OData information
        messagebox.showinfo(
            "OData Feed Retrieved", 
            f"🔗 OData Feed for '{report_name}'\n\n"
            f"📊 Export Job ID:\n{job_id}\n\n"
            f"🌐 OData Query URL:\n{odata_url}\n\n"
            f"📥 Direct Download URL:\n{download_url}\n\n"
            f"{clipboard_msg}\n\n"
            f"💡 Use these URLs in:\n"
            f"• PowerBI: Get Data → OData Feed\n"
            f"• Excel: Data → Get Data → From Other Sources → OData Feed\n"
            f"• Power Automate: HTTP connector\n"
            f"• Custom applications: REST API calls"
        )
        
        self.log_message(f"OData feed URLs retrieved for {report_name}", 'success')
        self.log_message(f"Job ID: {job_id}", 'info')
        self.log_message(f"OData URL: {odata_url}", 'info')
    
    def show_graph_odata_info(self, report_name):
        """Show Microsoft Graph OData information as fallback"""
        # Construct Graph API OData URLs
        base_graph_url = "https://graph.microsoft.com/beta/deviceManagement/reports"
        odata_metadata_url = f"{base_graph_url}/$metadata"
        export_jobs_url = f"{base_graph_url}/exportJobs"
        
        # Copy to clipboard
        import subprocess
        clipboard_content = f"""Microsoft Graph OData Endpoints:
Metadata: {odata_metadata_url}
Export Jobs: {export_jobs_url}
Report: {report_name}"""
        
        try:
            subprocess.run(['powershell', '-command', f'Set-Clipboard -Value @"\\n{clipboard_content}\\n"@'], 
                         check=True, capture_output=True)
            clipboard_msg = "📋 URLs copied to clipboard!"
        except:
            clipboard_msg = "📋 Copy URLs manually from below"
        
        messagebox.showinfo(
            "Microsoft Graph OData Info", 
            f"🔗 Microsoft Graph OData for '{report_name}'\n\n"
            f"📊 OData Metadata URL:\n{odata_metadata_url}\n\n"
            f"🚀 Export Jobs Endpoint:\n{export_jobs_url}\n\n"
            f"🔑 Authentication: Bearer token required\n"
            f"📝 Report Name: {report_name}\n\n"
            f"{clipboard_msg}\n\n"
            f"💡 Usage Examples:\n"
            f"• PowerBI: Get Data → OData Feed → Use metadata URL\n"
            f"• Excel: Data → From Web → Use export jobs URL\n"
            f"• REST API: POST to export jobs with report name\n"
            f"• Documentation: docs.microsoft.com/graph/api/intune-reporting"
        )
        
        self.log_message(f"Graph OData information provided for {report_name}", 'success')
    
    def open_powerbi_folder(self, folder_path):
        """Open the PowerBI imports folder in Windows Explorer"""
        try:
            import subprocess
            # Open the folder in Windows Explorer
            subprocess.Popen(f'explorer "{folder_path}"', shell=True)
            self.log_message(f"Opened folder: {folder_path}", 'success')
        except Exception as e:
            self.log_message(f"Failed to open folder: {str(e)}", 'warning')

    def show_powerbi_not_found_message(self, csv_path, selected_columns):
        """Show message when PowerBI is not found"""
        message = f"""Power BI Desktop not found on your system.

Your data has been saved to:
{csv_path}

To use in Power BI:
1. Install Power BI Desktop from Microsoft Store
2. Open Power BI Desktop
3. Click 'Get Data' → 'Text/CSV'
4. Navigate to the file above
5. Click 'Load'

Data Summary:
📊 Columns: {len(selected_columns)}
📊 Rows: {len(self.current_export_data)}"""
        
        messagebox.showinfo("Power BI Desktop Not Found", message)
    
    def create_toggle_switch(self, parent):
        """Create a custom toggle switch"""
        # Toggle switch container
        toggle_frame = tk.Frame(parent, bg='white')
        toggle_frame.pack(side='left')
        
        # Toggle switch background
        self.toggle_bg = tk.Frame(toggle_frame, width=50, height=25, bg='#cccccc', 
                                 relief='solid', bd=1)
        self.toggle_bg.pack()
        self.toggle_bg.pack_propagate(False)
        
        # Toggle switch slider
        self.toggle_slider = tk.Frame(self.toggle_bg, width=23, height=23, bg='white',
                                     relief='solid', bd=1)
        self.toggle_slider.place(x=1, y=1)
        
        # Bind click events
        self.toggle_bg.bind("<Button-1>", self.on_toggle_click)
        self.toggle_slider.bind("<Button-1>", self.on_toggle_click)
        
        # Initial state
        self.update_toggle_appearance()
    
    def on_toggle_click(self, event=None):
        """Handle toggle switch click"""
        self.readme_var.set(not self.readme_var.get())
        self.update_toggle_appearance()
        self.toggle_readme()
    
    def update_toggle_appearance(self):
        """Update toggle switch appearance based on state"""
        if self.readme_var.get():
            # ON state - blue background, slider to right
            self.toggle_bg.config(bg='#0078d4')
            self.toggle_slider.place(x=25, y=1)
        else:
            # OFF state - gray background, slider to left
            self.toggle_bg.config(bg='#cccccc')
            self.toggle_slider.place(x=1, y=1)
    
    def toggle_readme(self):
        """Toggle README window on/off"""
        if self.readme_var.get():
            self.show_readme_window()
        else:
            self.close_readme_window()
    
    def show_readme_window(self):
        """Show README window"""
        if hasattr(self, 'readme_window') and self.readme_window.winfo_exists():
            self.readme_window.window.lift()
            return
            
        self.readme_window = ReadmeWindow(self.root, self)
    
    def close_readme_window(self):
        """Close README window"""
        if hasattr(self, 'readme_window') and self.readme_window.winfo_exists():
            self.readme_window.window.destroy()
        self.readme_var.set(False)
        if hasattr(self, 'update_toggle_appearance'):
            self.update_toggle_appearance()
    
    def run(self):
        """Start the application"""
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            pass
        except Exception as e:
            messagebox.showerror("Application Error", f"Application error: {str(e)}")

class ReadmeWindow:
    """README window with comprehensive documentation"""
    
    def __init__(self, parent, main_app):
        self.parent = parent
        self.main_app = main_app
        self.window = tk.Toplevel(parent)
        self.window.title("README - HTMD Intune Reports Tool")
        self.window.geometry("900x700")
        self.window.minsize(800, 600)
        
        # Configure window
        self.window.configure(bg='#f5f5f5')
        self.window.resizable(True, True)
        
        # Icon (if available)
        try:
            self.window.iconbitmap(parent.tk.call('wm', 'iconbitmap', parent))
        except:
            pass
        
        # Window close event
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)
        
        self.create_content()
        
        # Center window
        self.center_window()
    
    def center_window(self):
        """Center the window on screen"""
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def create_content(self):
        """Create README content"""
        
        # Header with developer info - increased height to accommodate content
        header_frame = tk.Frame(self.window, bg='#0078d4', height=120)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        header_content = tk.Frame(header_frame, bg='#0078d4')
        header_content.pack(expand=True, fill='both', padx=20, pady=15)
        
        # Title on its own line
        title_label = tk.Label(header_content, text="HTMD - Microsoft Intune Reports Export Tool", 
                              font=('Segoe UI', 16, 'bold'), 
                              bg='#0078d4', fg='white')
        title_label.pack(anchor='w', pady=(0, 8))
        
        # Developer info on separate line with proper spacing
        dev_frame = tk.Frame(header_content, bg='#0078d4')
        dev_frame.pack(anchor='w', fill='x')
        
        dev_label = tk.Label(dev_frame, text="Developer: HTMD Community", 
                            font=('Segoe UI', 11, 'bold'), 
                            bg='#0078d4', fg='white')
        dev_label.pack(side='left')
        
        contact_label = tk.Label(dev_frame, text="Contact: +91 8971222240", 
                                font=('Segoe UI', 11, 'bold'), 
                                bg='#0078d4', fg='white')
        contact_label.pack(side='left', padx=(30, 0))
        
        # Main content area with scrollbar
        main_frame = tk.Frame(self.window, bg='#f5f5f5')
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Create scrollable text widget
        text_frame = tk.Frame(main_frame)
        text_frame.pack(fill='both', expand=True)
        
        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side='right', fill='y')
        
        self.text_widget = tk.Text(text_frame, wrap='word', yscrollcommand=scrollbar.set,
                                  font=('Segoe UI', 11), bg='white', fg='#323130',
                                  padx=20, pady=20, relief='solid', bd=1)
        self.text_widget.pack(side='left', fill='both', expand=True)
        
        scrollbar.config(command=self.text_widget.yview)
        
        # Configure text tags for formatting BEFORE inserting content
        self.text_widget.tag_configure("title", font=('Segoe UI', 14, 'bold', 'underline'), foreground='#0078d4', spacing1=10, spacing3=5)
        self.text_widget.tag_configure("header", font=('Segoe UI', 12, 'bold'), foreground='#0078d4', spacing1=8, spacing3=3)
        self.text_widget.tag_configure("subheader", font=('Segoe UI', 11, 'bold'), foreground='#323130', spacing1=5, spacing3=2)
        self.text_widget.tag_configure("important", font=('Segoe UI', 10, 'bold'), foreground='#d13438')
        self.text_widget.tag_configure("success", font=('Segoe UI', 10, 'bold'), foreground='#107c10')
        self.text_widget.tag_configure("normal", font=('Segoe UI', 10), foreground='#323130')
        
        # Insert README content
        self.insert_formatted_content()
        self.text_widget.config(state='disabled')  # Make read-only
    
    def insert_formatted_content(self):
        """Insert formatted README content"""
        # Main title
        self.text_widget.insert(tk.END, "PREREQUISITES AND SETUP GUIDE\n", "title")
        self.text_widget.insert(tk.END, "\n")
        
        # System Requirements
        self.text_widget.insert(tk.END, "SYSTEM REQUIREMENTS\n", "header")
        self.text_widget.insert(tk.END, "• Python Version: 3.8 or higher (Recommended: 3.11+)\n", "normal")
        self.text_widget.insert(tk.END, "• Operating System: Windows 10/11 (Primary), Linux, macOS\n", "normal")
        self.text_widget.insert(tk.END, "• Memory: Minimum 4GB RAM, Recommended 8GB+\n", "normal")
        self.text_widget.insert(tk.END, "• Storage: 1GB free space for exports and dependencies\n\n", "normal")
        
        # Required Modules
        self.text_widget.insert(tk.END, "REQUIRED PYTHON MODULES\n", "header")
        self.text_widget.insert(tk.END, "• requests (HTTP requests to Microsoft Graph API)\n", "normal")
        self.text_widget.insert(tk.END, "• pandas (Data processing and CSV handling)\n", "normal")
        self.text_widget.insert(tk.END, "• tkinter (GUI framework - usually included with Python)\n\n", "normal")
        
        # Optional Modules
        self.text_widget.insert(tk.END, "OPTIONAL MODULES (Enhanced Features)\n", "header")
        self.text_widget.insert(tk.END, "• pyautogui (PowerBI automation)\n", "normal")
        self.text_widget.insert(tk.END, "• openpyxl (Excel export support)\n\n", "normal")
        
        # Installation
        self.text_widget.insert(tk.END, "INSTALLATION\n", "header")
        self.text_widget.insert(tk.END, "1. Install Python from python.org\n", "normal")
        self.text_widget.insert(tk.END, "2. Install required modules:\n", "normal")
        self.text_widget.insert(tk.END, "   pip install requests pandas pyautogui openpyxl\n\n", "normal")
        
        # Azure AD Permissions - Critical Section
        self.text_widget.insert(tk.END, "ENTRA ID APPLICATION PERMISSIONS\n", "header")
        self.text_widget.insert(tk.END, "CRITICAL: Use DELEGATED Permissions, NOT Application Permissions\n\n", "important")
        
        # Required Permissions
        self.text_widget.insert(tk.END, "REQUIRED DELEGATED PERMISSIONS:\n", "subheader")
        permissions = [
            "• DeviceManagementConfiguration.Read.All", 
            "• DeviceManagementManagedDevices.Read.All",
            "• DeviceManagementApps.Read.All",
            "• User.ReadBasic.All",
            "• Group.Read.All",
            "• Directory.Read.All"
        ]
        for perm in permissions:
            self.text_widget.insert(tk.END, f"{perm}\n", "success")
        self.text_widget.insert(tk.END, "\n")
        
        # Delegated vs Application
        self.text_widget.insert(tk.END, "DELEGATED vs APPLICATION PERMISSIONS - Why Delegated?\n", "header")
        
        self.text_widget.insert(tk.END, "DELEGATED PERMISSIONS:\n", "subheader")
        delegated_benefits = [
            "✅ Acts on behalf of the signed-in user",
            "✅ Inherits user's existing permissions", 
            "✅ Works with corporate authentication",
            "✅ Supports interactive login flow",
            "✅ Respects user's role-based access",
            "✅ More secure for end-user tools"
        ]
        for benefit in delegated_benefits:
            self.text_widget.insert(tk.END, f"{benefit}\n", "success")
        self.text_widget.insert(tk.END, "\n")
        
        self.text_widget.insert(tk.END, "APPLICATION PERMISSIONS:\n", "subheader")
        app_issues = [
            "❌ Acts as the application itself",
            "❌ Requires admin consent for entire tenant",
            "❌ No user context - runs with app identity", 
            "❌ Complex certificate-based authentication",
            "❌ Overly broad access for user tools",
            "❌ Not suitable for interactive tools"
        ]
        for issue in app_issues:
            self.text_widget.insert(tk.END, f"{issue}\n", "important")
        self.text_widget.insert(tk.END, "\n")
        
        # Why Delegated
        self.text_widget.insert(tk.END, "WHY THIS TOOL USES DELEGATED PERMISSIONS:\n", "subheader")
        why_text = """This tool is designed as an interactive end-user application that authenticates with your corporate account and exports data based on your existing Intune permissions. Using delegated permissions ensures:

1. You only see data you're already authorized to access
2. No need for tenant-wide admin consent
3. Corporate authentication support
4. Compliance with least-privilege security principles

"""
        self.text_widget.insert(tk.END, why_text, "normal")
        
        # Permission Assignment
        self.text_widget.insert(tk.END, "HOW TO ASSIGN PERMISSIONS IN ENTRA ID\n", "header")
        
        self.text_widget.insert(tk.END, "METHOD 1: Azure Portal (Recommended)\n", "subheader")
        steps = [
            "1. Go to portal.azure.com",
            "2. Navigate to Azure Active Directory",
            "3. Select \"App registrations\"",
            "4. Find your application or create new one",
            "5. Go to \"Authentication\" and add Redirect URI:",
            "   - Platform: Web",
            "   - Redirect URI: http://localhost:8080/callback",
            "6. Go to \"API permissions\"",
            "7. Click \"Add a permission\"",
            "8. Select \"Microsoft Graph\"",
            "9. Choose \"Delegated permissions\"",
            "10. Search and add each required permission:",
            "   - DeviceManagementConfiguration.Read.All",
            "   - DeviceManagementManagedDevices.Read.All", 
            "   - DeviceManagementApps.Read.All",
            "   - DeviceManagementServiceConfig.Read.All",
            "   - User.Read",
            "11. Click \"Add permissions\"",
            "12. Click \"Grant admin consent for [Organization]\"",
            "13. IMPORTANT: Admin consent is required for all permissions"
        ]
        for step in steps:
            self.text_widget.insert(tk.END, f"{step}\n", "normal")
        self.text_widget.insert(tk.END, "\n")
        
        # Additional sections
        self.text_widget.insert(tk.END, "SUPPORTED REPORTS\n", "header")
        reports = [
            "• 177+ Intune Reports (Dynamic discovery)",
            "• Device Management Reports",
            "• Application Reports", 
            "• Compliance Reports",
            "• Configuration Reports",
            "• User and Group Information"
        ]
        for report in reports:
            self.text_widget.insert(tk.END, f"{report}\n", "normal")
        self.text_widget.insert(tk.END, "\n")
        
        # Support
        self.text_widget.insert(tk.END, "SUPPORT AND UPDATES\n", "header")
        self.text_widget.insert(tk.END, "• Developer: HTMD Community\n", "success")
        self.text_widget.insert(tk.END, "• Contact: +91 8971222240\n", "success")
        self.text_widget.insert(tk.END, "• Version: 1.0 (November 2025)\n", "normal")
        self.text_widget.insert(tk.END, "\nFor additional support or feature requests, contact the HTMD Community.", "normal")
        """Get comprehensive README content"""
        return """
PREREQUISITES AND SETUP GUIDE

SYSTEM REQUIREMENTS
• Python Version: 3.8 or higher (Recommended: 3.11+)
• Operating System: Windows 10/11 (Primary), Linux, macOS
• Memory: Minimum 4GB RAM, Recommended 8GB+
• Storage: 1GB free space for exports and dependencies

REQUIRED PYTHON MODULES
• requests (HTTP requests to Microsoft Graph API)
• pandas (Data processing and CSV handling)
• tkinter (GUI framework - usually included with Python)

OPTIONAL MODULES (Enhanced Features)
• pyautogui (PowerBI automation)
• openpyxl (Excel export support)

INSTALLATION
1. Install Python from python.org
2. Install required modules:
   pip install requests pandas pyautogui openpyxl

ENTRA ID APPLICATION PERMISSIONS

CRITICAL: Use DELEGATED Permissions, NOT Application Permissions

REQUIRED DELEGATED PERMISSIONS:
• DeviceManagementConfiguration.Read.All
• DeviceManagementManagedDevices.Read.All
• DeviceManagementApps.Read.All
• User.ReadBasic.All
• Group.Read.All
• Directory.Read.All

DELEGATED vs APPLICATION PERMISSIONS - Why Delegated?

DELEGATED PERMISSIONS:
✅ Acts on behalf of the signed-in user
✅ Inherits user's existing permissions
✅ Works with corporate authentication
✅ Supports interactive login flow
✅ Respects user's role-based access
✅ More secure for end-user tools

APPLICATION PERMISSIONS:
❌ Acts as the application itself
❌ Requires admin consent for entire tenant
❌ No user context - runs with app identity
❌ Complex certificate-based authentication
❌ Overly broad access for user tools
❌ Not suitable for interactive tools

WHY THIS TOOL USES DELEGATED PERMISSIONS:
This tool is designed as an interactive end-user application that authenticates with your corporate account and exports data based on your existing Intune permissions. Using delegated permissions ensures:

1. You only see data you're already authorized to access
2. No need for tenant-wide admin consent
3. Corporate authentication support
4. Compliance with least-privilege security principles

HOW TO ASSIGN PERMISSIONS IN ENTRA ID

Azure Portal Method (Recommended)
1. Go to https://entra.microsoft.com/
2. Navigate to Entra ID 
3. Select "App registrations"
4. Find your application or create new one
5. Go to "API permissions"
6. Click "Add a permission"
7. Select "Microsoft Graph"
8. Choose "Delegated permissions"
9. Search and add each required permission:
   - DeviceManagementConfiguration.Read.All
   - DeviceManagementManagedDevices.Read.All
   - DeviceManagementApps.Read.All
   - User.ReadBasic.All
   - Group.Read.All
   - Directory.Read.All
10. Click "Add permissions"
11. IMPORTANT: Admin consent may be required for some permissions

METHOD 2: PowerShell (Advanced Users)
Use Microsoft Graph PowerShell module to assign permissions programmatically.

METHOD 3: Azure CLI
Use Azure CLI commands for automation scenarios.

AUTHENTICATION FLOW
1. Corporate Login: Click "Sign in with Corporate Account"
2. Browser Redirect: Redirects to Microsoft login page
3. Corporate Authentication: Sign in with your work account
4. Permission Consent: Accept permissions (first time only)
5. Token Exchange: Application receives access token
6. Graph API Access: Tool can now access Microsoft Graph APIs

SUPPORTED REPORTS
• 177+ Intune Reports (Dynamic discovery)
• Device Management Reports
• Application Reports
• Compliance Reports
• Configuration Reports
• User and Group Information
• Custom API endpoints

EXPORT FORMATS
• CSV (Comma-separated values)
• Excel (XLSX format)
• PowerBI Integration
• OData Feed URLs
• Direct API access

TROUBLESHOOTING

PERMISSION ERRORS:
• Ensure all delegated permissions are granted
• Check if admin consent is required
• Verify your user account has Intune access
• Contact your IT administrator for permission issues

AUTHENTICATION ERRORS:
• Clear browser cache and cookies
• Try incognito/private browsing mode
• Ensure corporate account is active
• Check network connectivity

EXPORT ERRORS:
• Verify report exists and has data
• Check file permissions in save location
• Ensure sufficient disk space
• Try smaller date ranges for large reports

PERFORMANCE TIPS:
• Use column filtering to reduce data size
• Export data in smaller time windows
• Close other applications during large exports
• Monitor system memory usage

SECURITY CONSIDERATIONS
• Tool uses OAuth2 with Microsoft identity platform
• No passwords stored locally
• Access tokens are temporary and auto-expire
• All communication uses HTTPS encryption
• Respects your existing Intune role permissions

SUPPORT AND UPDATES
• Developer: HTMD Community
• Contact: +91 8971222240
• Regular updates for new Intune features
• Community support and documentation

VERSION INFORMATION
• Current Version: 1.0
• Release Date: November 2025
• Compatibility: Microsoft Graph API v1.0 and beta
• Supported Intune: All current versions

For additional support or feature requests, contact the HTMD Community.
"""
    
    def on_close(self):
        """Handle window close event"""
        if self.main_app and hasattr(self.main_app, 'readme_var'):
            self.main_app.readme_var.set(False)
            if hasattr(self.main_app, 'update_toggle_appearance'):
                self.main_app.update_toggle_appearance()
        self.window.destroy()
    
    def winfo_exists(self):
        """Check if window exists"""
        try:
            return self.window.winfo_exists()
        except:
            return False

def main():
    """Main function - Direct GUI launch"""
    try:
        print("🚀 Starting Intune Reports GUI...")
        app = IntuneReportsGUI()
        print("✅ GUI initialized successfully")
        app.run()
    except Exception as e:
        try:
            print(f"❌ Error during startup: {str(e)}")
            import traceback
            traceback.print_exc()
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Startup Error", f"Failed to start application:\n\n{str(e)}")
        except:
            print(f"Critical Error: {str(e)}")

if __name__ == "__main__":
    main()