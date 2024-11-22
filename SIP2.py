import pandas as pd
import re
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

class SIPLogAnalyzer:
    def __init__(self):
        self.log_entries = []
        
    def parse_log_file(self, filename):
        current_session = {}
        
        with open(filename, 'r') as file:
            for line in file:
                # Extract timestamp and Session ID
                timestamp_match = re.match(r'(\d{2}:\d{2}:\d{2}\.\d{3})', line)
                sid_match = re.search(r'\[SID=([\w:]+)\]', line)
                
                if timestamp_match:
                    timestamp = timestamp_match.group(1)
                    
                    # Extract IP addresses
                    ip_matches = re.findall(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', line)
                    remote_ip = ip_matches[1] if len(ip_matches) > 1 else ''
                    
                    # Extract IP Group
                    ip_group_match = re.search(r'IPG_(\w+)', line)
                    ip_group = ip_group_match.group(1) if ip_group_match else ''
                    
                    # Extract Direction
                    direction = ''
                    if 'Incoming SIP Message' in line:
                        direction = 'Incoming'
                    elif 'Outgoing SIP Message' in line:
                        direction = 'Outgoing'
                    
                    # Extract Caller and Callee from SIP messages
                    if 'FROM:' in line:
                        from_match = re.search(r'FROM: <(.+?)>', line)
                        if from_match:
                            current_session['caller'] = from_match.group(1)
                    
                    if 'TO:' in line:
                        to_match = re.search(r'TO: <(.+?)>', line)
                        if to_match:
                            current_session['callee'] = to_match.group(1)
                    
                    # Extract Termination Reason
                    term_reason_match = re.search(r'Call End Reason: (.+?)$', line)
                    if term_reason_match:
                        current_session['termination_reason'] = term_reason_match.group(1)
                    
                    # When we detect a call end
                    if 'Call End' in line or 'Released' in line:
                        self.log_entries.append({
                            'Call End Time': timestamp,
                            'Session ID': sid_match.group(1) if sid_match else '',
                            'IP Group': ip_group,
                            'Caller': current_session.get('caller', ''),
                            'Callee': current_session.get('callee', ''),
                            'Direction': direction,
                            'Remote IP': remote_ip,
                            'Termination Reason': current_session.get('termination_reason', 'Normal'),
                            'Duration': current_session.get('duration', '')
                        })
                        current_session = {}  # Reset for next call
                        
    def create_excel(self, output_filename):
        df = pd.DataFrame(self.log_entries)
        
        # Create Excel writer object
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            # Write the data
            df.to_excel(writer, sheet_name='Call Details', index=False)
            
            # Get the workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Call Details']
            
            # Add some formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D3D3D3',
                'border': 1
            })
            
            # Format the header
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 20)  # Set column width

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Create analyzer instance
    analyzer = SIPLogAnalyzer()
    
    # Ask for input file
    input_file = filedialog.askopenfilename(
        title="Select SIP Log File",
        filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
    )
    
    if input_file:
        try:
            # Process the file
            analyzer.parse_log_file(input_file)
            
            # Create output filename
            output_file = input_file.rsplit('.', 1)[0] + '_analysis.xlsx'
            
            # Generate Excel file
            analyzer.create_excel(output_file)
            
            messagebox.showinfo("Success", f"Analysis complete!\nFile saved as: {output_file}")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    root.destroy()

if __name__ == "__main__":
    main()
