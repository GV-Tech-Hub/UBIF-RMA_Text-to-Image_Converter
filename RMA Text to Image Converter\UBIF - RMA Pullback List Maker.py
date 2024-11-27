import tkinter as tk
from tkinter import ttk
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
import os

class RMAConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("RMA to Image Converter")
        self.root.geometry("800x600")

        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Create and configure text area
        self.text_area = tk.Text(main_frame, height=20, width=80)
        self.text_area.grid(row=0, column=0, columnspan=2, pady=10)

        # Create scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient='vertical', command=self.text_area.yview)
        scrollbar.grid(row=0, column=2, sticky='ns')
        self.text_area['yscrollcommand'] = scrollbar.set

        # Create convert button
        convert_btn = ttk.Button(main_frame, text="Convert to Image", command=self.convert_to_image)
        convert_btn.grid(row=1, column=0, pady=10)

        # Create clear button
        clear_btn = ttk.Button(main_frame, text="Clear", command=self.clear_text)
        clear_btn.grid(row=1, column=1, pady=10)

        # Status label
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.grid(row=2, column=0, columnspan=2)

    def clear_text(self):
        self.text_area.delete('1.0', tk.END)
        self.status_label.config(text="")

    def parse_rma_text(self, text):
        # Split the text into lines and filter out empty lines
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        # Find the header line (usually contains 'RMA', 'Model/Serial', etc.)
        header_index = None
        for i, line in enumerate(lines):
            if 'RMA' in line and 'Model/Serial' in line:
                header_index = i
                break

        if header_index is None:
            raise ValueError("Could not find header line in the text")

        # Parse the data lines
        data = []
        for line in lines[header_index + 1:]:
            # Split on tabs or multiple spaces
            parts = [part.strip() for part in line.split('\t')]
            if len(parts) == 1:
                parts = [part.strip() for part in line.split('  ') if part.strip()]
            
            if len(parts) >= 4:  # Ensure we have at least RMA, Model/Serial, Part, and Price
                data.append(parts[:4])  # Take only the first 4 columns

        # Create DataFrame
        df = pd.DataFrame(data, columns=['RMA', 'Model/Serial', 'Part', 'Price'])
        return df

    def convert_to_image(self):
        try:
            # Get text from text area
            text = self.text_area.get('1.0', tk.END)
            
            # Parse the text into a DataFrame
            df = self.parse_rma_text(text)

            # Create table figure
            fig = go.Figure(data=[go.Table(
                header=dict(
                    values=list(df.columns),
                    fill_color='#0066cc',
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

            # Update layout
            fig.update_layout(
                title='Pending RMA Part Removal',
                width=1200,
                height=800,
                margin=dict(l=20, r=20, t=40, b=20)
            )

            # Create 'exports' directory if it doesn't exist
            if not os.path.exists('exports'):
                os.makedirs('exports')

            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"exports/rma_parts_list_{timestamp}.png"

            # Save as PNG
            fig.write_image(filename)

            # Open the image file using the default system viewer
            os.startfile(os.path.abspath(filename))

            self.status_label.config(text=f"Image saved and opened: {filename}")

        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")

def main():
    root = tk.Tk()
    app = RMAConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()