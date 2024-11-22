import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime
import logging
from typing import Tuple, Dict, Optional

class SIPCodeAnalyzer:
    def __init__(self):
        self.setup_logging()
        self.load_explanations()

    def setup_logging(self) -> None:
        """Configure logging for the application."""
        log_dir = "logs"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        log_file = os.path.join(log_dir, f"sip_analyzer_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        logging.basicConfig(
            filename=log_file,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

    def load_explanations(self) -> None:
        """Load all SIP code explanations and related data."""
        self.sip_categories = {
            "1xx": "Provisional",
            "2xx": "Success",
            "3xx": "Redirection",
            "4xx": "Client Error",
            "5xx": "Server Error",
            "6xx": "Global Failure"
        }

        # Enhanced error categories with more specific information
        self.error_categories = {
            "Authentication": ["401", "407"],
            "Routing": ["404", "480", "484", "485", "502", "504"],
            "Capacity": ["486", "503", "513"],
            "Protocol": ["400", "405", "415", "420", "421", "505"],
            "Network": ["408", "483", "500", "502"],
            "Configuration": ["403", "488", "501", "606"]
        }

        # Load the same SIP explanations as before, but structure them better
        self.sip_explanations = {
            # Your existing sip_explanations dictionary here
        }

        # Enhanced Microsoft subcode explanations
        self.microsoft_subcode_explanations = {
            560486: {
                "description": "Network busy or user unavailable",
                "cause": "The destination user agent or network is temporarily unavailable",
                "resolution": "Retry the call after a brief delay. If persistent, check network conditions."
            },
            560487: {
                "description": "Call cancelled or network timeout",
                "cause": "The call was terminated due to timeout or user cancellation",
                "resolution": "Check network latency and connection stability."
            },
            560404: {
                "description": "User or number not found",
                "cause": "The dialed number is invalid or user does not exist",
                "resolution": "Verify the phone number and user existence in the system."
            },
            560480: {
                "description": "Temporary service interruption",
                "cause": "Service is temporarily unavailable",
                "resolution": "Wait for service restoration and retry. Check service status."
            },
            560503: {
                "description": "Service currently unavailable",
                "cause": "System overload or maintenance",
                "resolution": "Wait for service restoration. If persistent, contact support."
            }
        }

    def get_error_category(self, sip_code: int) -> str:
        """Determine the error category for a given SIP code."""
        sip_str = str(sip_code)
        for category, codes in self.error_categories.items():
            if sip_str in codes:
                return category
        return "Uncategorized"

    def explain_sip(self, final_sip_code: int, microsoft_subcode: int, sip_phrase: str) -> Tuple[str, str, Dict]:
        """
        Provide detailed explanation for SIP codes and Microsoft subcodes.
        Returns tuple of (simple explanation, detailed explanation, diagnostic info)
        """
        try:
            # Get category information
            category_prefix = str(final_sip_code)[0] + "xx"
            category = self.sip_categories.get(category_prefix, "Unknown Category")
            error_category = self.get_error_category(final_sip_code)

            # Get basic explanations
            simple_explanation = self.sip_explanations.get(final_sip_code, f"Unknown call issue: {final_sip_code}")
            
            # Get Microsoft subcode details
            subcode_info = self.microsoft_subcode_explanations.get(microsoft_subcode, {
                "description": f"Unrecognized network condition: {microsoft_subcode}",
                "cause": "Unknown",
                "resolution": "Contact support for more information"
            })

            # Create diagnostic information
            diagnostic_info = {
                "sip_code": final_sip_code,
                "category": category,
                "error_category": error_category,
                "subcode": microsoft_subcode,
                "subcode_description": subcode_info["description"],
                "cause": subcode_info["cause"],
                "resolution": subcode_info["resolution"],
                "original_phrase": sip_phrase
            }

            # Construct detailed explanation
            detailed_explanation = (
                f"Category: {category} ({error_category})\n"
                f"SIP Code {final_sip_code}: {simple_explanation}\n"
                f"Microsoft Subcode {microsoft_subcode}: {subcode_info['description']}\n"
                f"Cause: {subcode_info['cause']}\n"
                f"Resolution: {subcode_info['resolution']}\n"
                f"Original Phrase: {sip_phrase}"
            )

            return simple_explanation, detailed_explanation, diagnostic_info

        except Exception as e:
            logging.error(f"Error in explain_sip: {str(e)}")
            return "Error processing code", "Error in processing", {}

    def process_file(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        """
        Process the input Excel file and generate enhanced output with error analysis.
        Returns tuple of (success boolean, message string)
        """
        try:
            # Load and validate input file
            if not os.path.exists(input_path):
                return False, "Input file does not exist"

            logging.info(f"Processing file: {input_path}")
            data = pd.read_excel(input_path)

            # Validate required columns
            required_fields = [
                'UPN', 'Display Name', 'Final SIP code',
                'Final Microsoft subcode', 'Final SIP Phrase'
            ]
            missing_fields = [field for field in required_fields if field not in data.columns]
            if missing_fields:
                return False, f"Missing required fields: {', '.join(missing_fields)}"

            # Process the data
            results = []
            for _, row in data.iterrows():
                simple_exp, detailed_exp, diagnostic = self.explain_sip(
                    int(row['Final SIP code']),
                    int(row['Final Microsoft subcode']),
                    row['Final SIP Phrase']
                )
                results.append({
                    'Simple_Explanation': simple_exp,
                    'Detailed_Explanation': detailed_exp,
                    **diagnostic
                })

            # Add results to dataframe
            results_df = pd.DataFrame(results)
            enhanced_data = pd.concat([data, results_df], axis=1)

            # Add summary statistics
            error_summary = enhanced_data.groupby(['error_category', 'Final SIP code']).size().reset_index(name='count')
            error_summary = error_summary.sort_values('count', ascending=False)

            # Save results
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                enhanced_data.to_excel(writer, sheet_name='Detailed Analysis', index=False)
                error_summary.to_excel(writer, sheet_name='Error Summary', index=False)

            logging.info(f"Successfully processed file. Output saved to: {output_path}")
            return True, f"Successfully processed file. Output saved to: {output_path}"

        except Exception as e:
            error_msg = f"Error processing file: {str(e)}"
            logging.error(error_msg)
            return False, error_msg

class ImprovedGUI:
    def __init__(self):
        self.analyzer = SIPCodeAnalyzer()
        self.create_gui()

    def create_gui(self):
        self.root = tk.Tk()
        self.root.title("Enhanced Call Log Analyzer")
        self.root.geometry("800x600")
        
        # Create main frame with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="5")
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)

        # Input file selection
        self.input_path_var = tk.StringVar()
        ttk.Label(file_frame, text="Input Excel File:").grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.input_path_var, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_input_file).grid(row=0, column=2, padx=5)

        # Output file selection
        self.output_path_var = tk.StringVar()
        ttk.Label(file_frame, text="Output Excel File:").grid(row=1, column=0, padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_path_var, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_output_file).grid(row=1, column=2, padx=5)

        # Progress frame
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="5")
        progress_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)

        self.progress_var = tk.StringVar(value="Ready to process...")
        ttk.Label(progress_frame, textvariable=self.progress_var).grid(row=0, column=0, padx=5, pady=5)
        self.progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)

        # Process button
        ttk.Button(
            main_frame,
            text="Process Files",
            command=self.process_and_generate,
            style="Accent.TButton"
        ).grid(row=2, column=0, pady=10)

        # Status log
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding="5")
        log_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        self.log_text = tk.Text(log_frame, height=10, width=70)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text['yscrollcommand'] = scrollbar.set

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        file_frame.columnconfigure(1, weight=1)
        progress_frame.columnconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

    def browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if file_path:
            self.input_path_var.set(file_path)
            self.log_message(f"Selected input file: {file_path}")

    def browse_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="Select Output Excel File",
            defaultextension=".xlsx",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if file_path:
            self.output_path_var.set(file_path)
            self.log_message(f"Selected output file: {file_path}")

    def log_message(self, message: str):
        """Add message to log with timestamp."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)

    def process_and_generate(self):
        """Process the input file and generate output with error handling."""
        input_path = self.input_path_var.get()
        output_path = self.output_path_var.get()

        if not input_path or not output_path:
            messagebox.showerror("Error", "Please select both input and output file paths.")
            return

        self.progress_bar.start()
        self.progress_var.set("Processing...")
        self.log_message("Starting file processing...")

        try:
            success, message = self.analyzer.process_file(input_path, output_path)
            self.progress_bar.stop()
            
            if success:
                self.progress_var.set("Processing complete!")
                messagebox.showinfo("Success", message)
            else:
                self.progress_var.set("Processing failed!")
                messagebox.showerror("Error", message)
            
            self.log_message(message)

        except Exception as e:
            self.progress_bar.stop()
            self.progress_var.set("Error occurred!")
            error_message = f"An unexpected error occurred: {str(e)}"
            self.log_message(error_message)
            messagebox.showerror("Error", error_message)

def main():
    app = ImprovedGUI()
    app.root.mainloop()

if __name__ == "__main__":
    main()