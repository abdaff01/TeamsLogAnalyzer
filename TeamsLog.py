import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def explain_sip(final_sip_code, microsoft_subcode, sip_phrase):
    sip_explanations = {
        # 1xx Provisional Responses
        100: "Call is being processed. Please wait.",
        180: "Phone is ringing at the destination.",
        181: "Call is being redirected to another number.",
        182: "Call is in a queue and will be handled soon.",
        183: "Call setup is in progress.",

        # 2xx Successful Responses
        200: "Call connected successfully.",
        202: "Request accepted and will be processed.",

        # 3xx Redirection Responses
        300: "Multiple call destinations found.",
        301: "Call destination has permanently changed.",
        302: "Call destination is temporarily different.",
        305: "You must use a specific network route.",
        380: "Alternative communication method available.",

        # 4xx Request Failure Responses
        400: "Invalid call request. Check the number.",
        401: "Authentication required to complete the call.",
        403: "Call blocked or not allowed.",
        404: "Phone number or user not found.",
        405: "Calling method not permitted.",
        406: "Call cannot be completed due to incompatible settings.",
        407: "Proxy authentication needed.",
        408: "No response from the destination. Timeout occurred.",
        410: "Number is no longer in service.",
        413: "Call request too large to process.",
        414: "Phone number too complicated to dial.",
        415: "Unsupported communication method.",
        416: "Unrecognized phone number format.",
        420: "Unsupported communication feature.",
        421: "Missing required communication feature.",
        423: "Call setup time too short.",
        480: "Destination temporarily unavailable.",
        481: "Call cannot be found or tracked.",
        482: "Call routing has created a loop.",
        483: "Too many network hops to complete call.",
        484: "Incomplete phone number.",
        485: "Unclear which number to call.",
        486: "Destination is currently busy.",
        487: "Call was cancelled or stopped.",
        488: "Call cannot be accepted by recipient.",

        # 5xx Server Failure Responses
        500: "Network error. Unable to complete call.",
        501: "Call feature not supported.",
        502: "Network routing problem.",
        503: "Network overloaded or maintenance.",
        504: "Network route timeout.",
        505: "Unsupported communication protocol.",
        513: "Call request too large.",

        # 6xx Global Failure Responses
        600: "User is busy everywhere.",
        603: "Call explicitly rejected.",
        604: "User does not exist.",
        606: "Call settings prevent connection."
    }

    microsoft_subcode_explanations = {
        560486: "Network busy or user unavailable.",
        560487: "Call cancelled or network timeout.",
        560404: "User or number not found.",
        560480: "Temporary service interruption.",
        560503: "Service currently unavailable.",
        0: "No additional details available."
    }

    # Detailed technical explanation
    detailed_explanations = {
        # 1xx Provisional Responses
        100: "Trying - The request is being processed, but no definitive response is available yet.",
        180: "Ringing - The destination user agent is alerting the user.",
        181: "Call Is Being Forwarded - The call is being redirected to another destination.",
        182: "Queued - The request is queued and will be processed soon.",
        183: "Session Progress - Provides progress information about the call setup.",

        # 2xx Successful Responses
        200: "OK - The request has been successfully processed and accepted.",
        202: "Accepted - The request has been accepted for processing, but not completed yet.",

        # 3xx Redirection Responses
        300: "Multiple Choices - The requested address resolves to multiple destinations.",
        301: "Moved Permanently - The requested address is no longer valid and has a new permanent address.",
        302: "Moved Temporarily - The requested address is temporarily unavailable and has a new temporary address.",
        305: "Use Proxy - The client must use the specified proxy to reach the destination.",
        380: "Alternative Service - The request cannot be fulfilled, but an alternative service is available.",

        # 4xx Request Failure Responses
        400: "Bad Request - The request could not be understood due to malformed syntax.",
        401: "Unauthorized - The request requires user authentication.",
        403: "Forbidden - The server understood the request but refuses to authorize it.",
        404: "Not Found - The requested user could not be located on the server.",
        405: "Method Not Allowed - The specified method is not allowed for the requested address.",
        406: "Not Acceptable - The requested resource cannot generate content matching the client's Accept headers.",
        407: "Proxy Authentication Required - The client must first authenticate with the proxy.",
        408: "Request Timeout - No response was received from the destination in a timely manner.",
        410: "Gone - The requested resource is no longer available and will not be available again.",
        413: "Request Entity Too Large - The request payload exceeds server processing capabilities.",
        414: "Request-URI Too Long - The request URI exceeds the server's maximum processing length.",
        415: "Unsupported Media Type - The request includes a media type the server cannot process.",
        416: "Unsupported URI Scheme - The request contains a URI scheme the server does not support.",
        420: "Bad Extension - The server does not understand a specified SIP extension.",
        421: "Extension Required - The server requires a specific extension not present in the request.",
        423: "Interval Too Brief - The request's expiration interval is too short.",
        480: "Temporarily Unavailable - The destination cannot be reached but might be available later.",
        481: "Call/Transaction Does Not Exist - The call or transaction referenced does not exist.",
        482: "Loop Detected - The request indicates a loop in the routing path.",
        483: "Too Many Hops - Maximum number of routing hops has been exceeded.",
        484: "Address Incomplete - The request URI is incomplete.",
        485: "Ambiguous - The request URI is ambiguous and could not be resolved uniquely.",
        486: "Busy Here - The destination is currently busy and cannot accept the call.",
        487: "Request Terminated - The request was terminated by the user or network.",
        488: "Not Acceptable Here - The request cannot be accepted by the recipient.",

        # 5xx Server Failure Responses
        500: "Server Internal Error - An unexpected condition prevented request fulfillment.",
        501: "Not Implemented - The server does not support the functionality required.",
        502: "Bad Gateway - The server received an invalid response from another server.",
        503: "Service Unavailable - The server is temporarily overloaded or under maintenance.",
        504: "Server Time-out - No response received from an upstream server.",
        505: "Version Not Supported - The SIP version is not supported.",
        513: "Message Too Large - The message exceeds the server's processing capabilities.",

        # 6xx Global Failure Responses
        600: "Busy Everywhere - The requested user is busy across all possible locations.",
        603: "Decline - The user explicitly declines the call.",
        604: "Does Not Exist Anywhere - The user cannot be found at any location.",
        606: "Not Acceptable - The user's preferences do not allow the call to be completed."
    }

    # Primary SIP code explanation
    simple_explanation = sip_explanations.get(final_sip_code, f"Unknown call issue: {final_sip_code}")
    
    # Microsoft subcode explanation
    subcode_explanation = microsoft_subcode_explanations.get(microsoft_subcode, f"Unrecognized network condition: {microsoft_subcode}")
    
    # Detailed technical explanation
    detailed_explanation = detailed_explanations.get(final_sip_code, "Detailed technical explanation not available.")
    
    # Simple explanation with subcode
    full_simple_explanation = f"{simple_explanation} {subcode_explanation}"
    
    # Detailed explanation with original phrase
    full_detailed_explanation = f"SIP Code: {final_sip_code} | {detailed_explanation} | Subcode: {microsoft_subcode} | Original Phrase: {sip_phrase}"
    
    return full_simple_explanation, full_detailed_explanation

def process_file(input_path, output_path):
    # Load the Excel file
    data = pd.read_excel(input_path)

    # Select only the specified fields
    fields = [
        'UPN', 'Display Name', 'User country', 'External Country', 'Invite time',
        'Start time', 'Failure time', 'End time', 'Duration (seconds)', 'Success',
        'Caller Number', 'Callee Number', 'Call type', 'Call Direction',
        'Azure region for Media', 'Azure region for Signaling', 'Final SIP code',
        'Final Microsoft subcode', 'Final SIP Phrase', 'SBC FQDN', 'Media bypass'
    ]
    data = data[fields]

    # Add two new columns with explanations
    data[['Explanation', 'Detailed Explanation']] = data.apply(
        lambda row: pd.Series(explain_sip(row['Final SIP code'], row['Final Microsoft subcode'], row['Final SIP Phrase'])), axis=1
    )

    # Save the output to the specified path
    data.to_excel(output_path, index=False)
    return output_path

# GUI application
def create_gui():
    def browse_input_file():
        file_path = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if file_path:
            input_path_var.set(file_path)

    def browse_output_file():
        file_path = filedialog.asksaveasfilename(
            title="Select Output Excel File",
            defaultextension=".xlsx",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if file_path:
            output_path_var.set(file_path)

    def process_and_generate():
        input_path = input_path_var.get()
        output_path = output_path_var.get()

        if not input_path or not output_path:
            messagebox.showerror("Error", "Please select both input and output file paths.")
            return

        try:
            process_file(input_path, output_path)
            messagebox.showinfo("Success", f"File processed successfully!\nOutput saved at: {output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    # Create main window
    root = tk.Tk()
    root.title("Call Log Analyzer")

    # Input file selection
    tk.Label(root, text="Input Excel File:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
    input_path_var = tk.StringVar()
    tk.Entry(root, textvariable=input_path_var, width=50).grid(row=0, column=1, padx=10, pady=5)
    tk.Button(root, text="Browse", command=browse_input_file).grid(row=0, column=2, padx=10, pady=5)

    # Output file selection
    tk.Label(root, text="Output Excel File:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    output_path_var = tk.StringVar()
    tk.Entry(root, textvariable=output_path_var, width=50).grid(row=1, column=1, padx=10, pady=5)
    tk.Button(root, text="Browse", command=browse_output_file).grid(row=1, column=2, padx=10, pady=5)

    # Process button
    tk.Button(root, text="Generate", command=process_and_generate, bg="green", fg="white").grid(row=2, column=1, pady=20)

    # Start the GUI event loop
    root.mainloop()

# Run the GUI application
if __name__ == "__main__":
    create_gui()
