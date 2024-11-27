import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
import json
import webbrowser
from typing import Dict, List
import platform

class RMAConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("RMA to Image Converter")
        self.root.geometry("900x700")
        
        # Enhanced color scheme with hover states
        self.colors = {
            'bg': '#004d40',  # Dark emerald green
            'fg': '#ffffff',  # White
            'accent': '#ffeb3b',  # Yellow
            'table_header': '#00796b',  # Lighter emerald green
            'hover': '#00897b',  # Hover state color
            'error': '#ff5252',  # Error color
            'success': '#4caf50',  # Success color
            'input_bg': '#f5f5f5',  # Light background for input
            'border': '#00897b'  # Border color
        }
        
        # Configure root window
        self.root.configure(bg=self.colors['bg'])
        self.setup_styles()
        
        # Create main container
        self.create_main_container()
        
        # Initialize history storage
        self.history_file = 'rma_history.json'
        self.history = []
        self.load_history()

    def setup_styles(self):
        """Configure custom styles for widgets"""
        style = ttk.Style()
        
        # Configure frame style
        style.configure(
            'Custom.TFrame',
            background=self.colors['bg']
        )
        
        # Configure label style
        style.configure(
            'Custom.TLabel',
            background=self.colors['bg'],
            foreground=self.colors['fg'],
            font=('Arial', 10)
        )
        
        # Configure button style
        style.configure(
            'Custom.TButton',
            background=self.colors['accent'],
            foreground=self.colors['bg'],
            padding=10,
            font=('Arial', 10, 'bold')
        )
        
        # Configure entry style
        style.configure(
            'Custom.TEntry',
            fieldbackground=self.colors['input_bg'],
            padding=5
        )

    def create_main_container(self):
        """Create and setup the main container and all widgets"""
        main_frame = ttk.Frame(self.root, padding="15", style='Custom.TFrame')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Create header frame
        self.create_header_frame(main_frame)
        
        # Create main content area
        self.create_content_area(main_frame)
        
        # Create footer
        self.create_footer(main_frame)

    def create_header_frame(self, parent):
        """Create the header section with date and settings"""
        header_frame = ttk.Frame(parent, style='Custom.TFrame')
        header_frame.grid(row=0, column=0, columnspan=3, pady=(0, 10), sticky='ew')

        # Date selection
        date_label = ttk.Label(
            header_frame,
            text="Due Date:",
            style='Custom.TLabel',
            font=('Arial', 10, 'bold')
        )
        date_label.pack(side='left', padx=(0, 10))

        self.due_date = DateEntry(
            header_frame,
            width=12,
            background=self.colors['table_header'],
            foreground='white',
            borderwidth=2,
            date_pattern='yyyy-mm-dd'
        )
        self.due_date.pack(side='left')

        # Settings button
        settings_btn = ttk.Button(
            header_frame,
            text="⚙️ Settings",
            style='Custom.TButton',
            command=self.show_settings
        )
        settings_btn.pack(side='right', padx=5)

    def create_content_area(self, parent):
        """Create the main content area with text input and preview"""
        # Text input frame
        input_frame = ttk.Frame(parent, style='Custom.TFrame')
        input_frame.grid(row=1, column=0, columnspan=3, sticky='nsew')

        # Instructions
        instructions = ttk.Label(
            input_frame,
            text="Paste RMA text data below:",
            style='Custom.TLabel'
        )
        instructions.pack(fill='x', pady=(0, 5))

        # Text area
        self.text_area = tk.Text(
            input_frame,
            height=20,
            width=90,
            font=('Consolas', 10),
            wrap='word',
            bg=self.colors['input_bg'],
            fg='#333333'
        )
        self.text_area.pack(side='left', fill='both', expand=True)

        # Scrollbar
        scrollbar = ttk.Scrollbar(
            input_frame,
            orient='vertical',
            command=self.text_area.yview
        )
        scrollbar.pack(side='right', fill='y')
        self.text_area['yscrollcommand'] = scrollbar.set

        # Total amount display
        total_frame = ttk.Frame(parent, style='Custom.TFrame')
        total_frame.grid(row=2, column=0, columnspan=3, pady=10, sticky='e')

        self.total_label = ttk.Label(
            total_frame,
            text="Total Potential Chargeback: $0.00",
            style='Custom.TLabel',
            font=('Arial', 12, 'bold')
        )
        self.total_label.pack(side='right')

    def create_footer(self, parent):
        """Create the footer with action buttons and status"""
        button_frame = ttk.Frame(parent, style='Custom.TFrame')
        button_frame.grid(row=3, column=0, columnspan=3, pady=10)

        # Convert button
        convert_btn = ttk.Button(
            button_frame,
            text="🔄 Convert to Image",
            style='Custom.TButton',
            command=self.convert_to_image
        )
        convert_btn.pack(side='left', padx=5)

        # Clear button
        clear_btn = ttk.Button(
            button_frame,
            text="🗑️ Clear",
            style='Custom.TButton',
            command=self.clear_text
        )
        clear_btn.pack(side='left', padx=5)

        # History button
        history_btn = ttk.Button(
            button_frame,
            text="📋 History",
            style='Custom.TButton',
            command=self.show_history
        )
        history_btn.pack(side='left', padx=5)

        # Status label
        self.status_label = ttk.Label(
            parent,
            text="Ready",
            style='Custom.TLabel'
        )
        self.status_label.grid(row=4, column=0, columnspan=3)

    def show_settings(self):
        """Display settings dialog"""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Settings")
        settings_window.geometry("400x300")
        settings_window.configure(bg=self.colors['bg'])
        
        # Add settings content here
        ttk.Label(
            settings_window,
            text="Settings",
            style='Custom.TLabel',
            font=('Arial', 14, 'bold')
        ).pack(pady=10)

    def show_history(self):
        """Display conversion history"""
        history_window = tk.Toplevel(self.root)
        history_window.title("Conversion History")
        history_window.geometry("600x400")
        history_window.configure(bg=self.colors['bg'])

        # Create treeview for history
        columns = ('Date', 'Items', 'Total', 'File')
        tree = ttk.Treeview(history_window, columns=columns, show='headings')
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)

        tree.pack(fill='both', expand=True, padx=10, pady=10)

        # Populate with history data
        for entry in self.history:
            tree.insert('', 'end', values=(
                entry.get('date'),
                entry.get('items'),
                entry.get('total'),
                entry.get('file')
            ))

    def load_history(self):
        """Load conversion history from file"""
        try:
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r') as f:
                    self.history = json.load(f)
        except Exception as e:
            self.show_error(f"Error loading history: {str(e)}")

    def save_history(self):
        """Save conversion history to file"""
        try:
            with open(self.history_file, 'w') as f:
                json.dump(self.history, f)
        except Exception as e:
            self.show_error(f"Error saving history: {str(e)}")

    def show_error(self, message: str):
        """Display error message"""
        self.status_label.config(
            text=f"Error: {message}",
            foreground=self.colors['error']
        )
        messagebox.showerror("Error", message)

    def show_success(self, message: str):
        """Display success message"""
        self.status_label.config(
            text=message,
            foreground=self.colors['success']
        )

    def clear_text(self):
        """Clear the text area and reset status"""
        self.text_area.delete('1.0', tk.END)
        self.total_label.config(text="Total Potential Chargeback: $0.00")
        self.status_label.config(
            text="Ready",
            foreground=self.colors['fg']
        )

    def parse_rma_text(self, text):
        try:
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            
            header_index = None
            for i, line in enumerate(lines):
                if 'RMA' in line and 'Model/Serial' in line:
                    header_index = i
                    break

            if header_index is None:
                raise ValueError("Could not find header line in the text")

            data = []
            for line in lines[header_index + 1:]:
                parts = [part.strip() for part in line.split('\t')]
                if len(parts) == 1:
                    parts = [part.strip() for part in line.split('  ') if part.strip()]
                
                if len(parts) >= 4:
                    price = parts[3].replace('$', '').replace(',', '')
                    if '.' in price:
                        price = price.split('.')[0]
                    parts[3] = price
                    data.append(parts[:4])

            df = pd.DataFrame(data, columns=['RMA', 'Model/Serial', 'Part', 'Price'])
            return df

        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")
            return pd.DataFrame()

    def convert_to_image(self):
        try:
            text = self.text_area.get('1.0', tk.END)
            df = self.parse_rma_text(text)
            
            if df.empty:
                return

            df['Price'] = pd.to_numeric(df['Price'], errors='coerce')
            total = df['Price'].sum()

            fig = go.Figure(data=[go.Table(
                header=dict(
                    values=list(df.columns),
                    fill_color=self.colors['table_header'],
                    font=dict(color='white', size=12),
                    align='left'
                ),
                cells=dict(
                    values=[df[col] for col in df.columns],
                    fill_color='white',
                    align='left',
                    font=dict(size=11)
                )
            )])

            due_date = self.due_date.get_date().strftime("%Y-%m-%d")
            
            fig.update_layout(
                title=f'Pending RMA Part Removal (Due: {due_date})<br>Total Potential Chargeback: ${total:,.2f}',
                width=1200,
                height=800,
                margin=dict(l=20, r=20, t=60, b=20)
            )

            if not os.path.exists('exports'):
                os.makedirs('exports')

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"exports/rma_parts_list_{timestamp}.png"

            fig.write_image(filename)

            if platform.system() == 'Windows':
                os.startfile(os.path.abspath(filename))
            elif platform.system() == 'Darwin':  # macOS
                os.system(f'open "{filename}"')
            else:  # Linux
                webbrowser.open(filename)

            self.status_label.config(text=f"Image saved and opened: {filename}")
            self.total_label.config(text=f"Total Potential Chargeback: ${total:,.2f}")

        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")

def main():
    try:
        root = tk.Tk()
        app = RMAConverter(root)
        root.mainloop()
    except Exception as e:
        print(f"Error during initialization: {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    main()
