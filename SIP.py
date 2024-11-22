import pandas as pd
import re
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

class SIPLogAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("SIP Log Analyzer")
        self.root.geometry("400x200")
        self.create_widgets()
        
    def create_widgets(self):
        self.select_button = tk.Button(self.root, text="Select TXT Log File", command=self.select_file)
        self.select_button.pack(pady=20)
        
        self.analyze_button = tk.Button(self.root, text="Analyze Logs", command=self.analyze_logs)
        self.analyze_button.pack(pady=10)
        
        self.status_label = tk.Label(self.root, text="")
        self.status_label.pack(pady=20)
        
    def select_file(self):
        self.filename = filedialog.askopenfilename(
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        if self.filename:
            self.status_label.config(text="File selected: " + self.filename.split('/')[-1])
    
    def parse_log_line(self, line):
        # Regular expressions for your specific log format
        patterns = {
            'timestamp': r'^(\d{2}:\d{2}:\d{2}\.\d{3})',
            'ip': r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
            'sid': r'\[SID=([\w:]+)\]',
            'message_type': r'(OPTIONS|INVITE|ACK|BYE|CANCEL|REGISTER|NOTIFY|SIP/2\.0 \d{3})',
            'call_id': r'CALL-ID: ([^\n]+)',
            'from_uri': r'FROM: ([^\n]+)',
            'to_uri': r'TO: ([^\n]+)',
            'user_agent': r'USER-AGENT: ([^\n]+)'
        }
        
        entry = {}
        
        # Extract basic information
        for key, pattern in patterns.items():
            match = re.search(pattern, line)
            entry[key] = match.group(1) if match else ''
            
        # Determine message direction and type
        entry['direction'] = 'Incoming' if 'Incoming SIP Message' in line else 'Outgoing' if 'Outgoing SIP Message' in line else 'Internal'
        
        # Extract status code if it's a response
        status_match = re.search(r'SIP/2\.0 (\d{3})', line)
        entry['status_code'] = status_match.group(1) if status_match else ''
        
        return entry

    def analyze_logs(self):
        if not hasattr(self, 'filename'):
            messagebox.showerror("Error", "Please select a log file first")
            return
            
        try:
            # Read the log file
            with open(self.filename, 'r', encoding='utf-8') as file:
                log_content = file.read()
            
            # Split into message blocks (each complete SIP message)
            messages = []
            current_message = []
            
            for line in log_content.split('\n'):
                if re.match(r'^\d{2}:\d{2}:\d{2}\.\d{3}', line):  # New message starts
                    if current_message:
                        messages.append('\n'.join(current_message))
                        current_message = []
                current_message.append(line)
            
            if current_message:  # Add the last message
                messages.append('\n'.join(current_message))
            
            # Parse each message
            parsed_data = []
            for message in messages:
                if message.strip():
                    entry = self.parse_log_line(message)
                    parsed_data.append(entry)
            
            # Create DataFrame
            df = pd.DataFrame(parsed_data)
            
            # Create Excel writer object
            output_file = self.filename.rsplit('.', 1)[0] + '_analysis.xlsx'
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                # Raw Data sheet
                df.to_excel(writer, sheet_name='Raw Data', index=False)
                
                # Message Types Analysis
                message_types = pd.DataFrame({
                    'Message Type': df['message_type'].value_counts().index,
                    'Count': df['message_type'].value_counts().values
                })
                message_types.to_excel(writer, sheet_name='Message Types', index=False)
                
                # Statistics sheet
                stats = pd.DataFrame({
                    'Metric': [
                        'Total Messages',
                        'Incoming Messages',
                        'Outgoing Messages',
                        'Internal Messages',
                        'Unique Sessions',
                        'OPTIONS Messages',
                        'Success Responses (2xx)',
                        'Client Errors (4xx)',
                        'Server Errors (5xx)'
                    ],
                    'Count': [
                        len(df),
                        len(df[df['direction'] == 'Incoming']),
                        len(df[df['direction'] == 'Outgoing']),
                        len(df[df['direction'] == 'Internal']),
                        df['sid'].nunique(),
                        len(df[df['message_type'] == 'OPTIONS']),
                        len(df[df['status_code'].str.startswith('2', na=False)]),
                        len(df[df['status_code'].str.startswith('4', na=False)]),
                        len(df[df['status_code'].str.startswith('5', na=False)])
                    ]
                })
                stats.to_excel(writer, sheet_name='Statistics', index=False)
                
                # Session Analysis
                session_analysis = df.groupby('sid').agg({
                    'message_type': 'count',
                    'direction': lambda x: x.value_counts().to_dict(),
                    'status_code': lambda x: x[x != ''].value_counts().to_dict()
                }).reset_index()
                session_analysis.to_excel(writer, sheet_name='Session Analysis', index=False)
                
                # IP Analysis
                ip_analysis = df.groupby('ip').agg({
                    'message_type': 'count',
                    'direction': lambda x: x.value_counts().to_dict()
                }).reset_index()
                ip_analysis.to_excel(writer, sheet_name='IP Analysis', index=False)
                
                # Format worksheets
                workbook = writer.book
                for worksheet in writer.sheets.values():
                    worksheet.set_column('A:Z', 20)  # Set column width
                    
            messagebox.showinfo("Success", f"Analysis complete! File saved as {output_file}")
            self.status_label.config(text="Analysis completed!")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_label.config(text="Analysis failed!")

def main():
    root = tk.Tk()
    app = SIPLogAnalyzer(root)
    root.mainloop()

if __name__ == "__main__":
    main()
