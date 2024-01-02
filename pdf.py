import csv
import PyPDF2
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from ttkthemes import ThemedTk

def extract_invoice_number(page_text):
    # Extract the invoice numbers
    match = re.search(r'INVOICE NO: (\d+)', page_text)
    if match:
        return match.group(1)
    else:
        return None

def extract_info_from_page(page_text):
    # Extract phone number, customer name, and message from the text
    phone_number_match = re.search(r'(?:(?:\+\d{1,2}\s?)|0)?(\d{10})', page_text)
    customer_name_match = re.search(r'Ship to: (.+?)(?=\n)', page_text)

    if phone_number_match:
        # Get the matched phone number
        raw_phone_number = phone_number_match.group(1)

        # Ensure the number starts with '3' and has a length of 10
        cleaned_phone_number = '3' + raw_phone_number[-9:]
    else:
        cleaned_phone_number = ''

    customer_name = customer_name_match.group(1).strip() if customer_name_match else 'name not found'
    message = "Hello, here is your order invoice"

    return cleaned_phone_number, customer_name, message


def split_pdf(input_pdf, output_folder):
    # Using PyPDF2 to read and write PDF files
    pdf_reader = PyPDF2.PdfReader(input_pdf)
    pdf_writer = None
    output_pdf = None

    # Open the CSV file for writing
    csv_file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])

    if not csv_file_path:
        # User canceled the file dialog
        return

    with open(csv_file_path, 'w', newline='') as csvfile:
        # Create a CSV writer
        csv_writer = csv.writer(csvfile)

        # Write the header row
        csv_writer.writerow(['Phone Number', 'Customer Name', 'Message', 'PDF Path'])

        # Loop through each page in the PDF
        for page_number in range(len(pdf_reader.pages)):
            # Extract text from the current page
            page_text = pdf_reader.pages[page_number].extract_text()
            # Extract the invoice number from the text
            invoice_number = extract_invoice_number(page_text)

            if invoice_number:
                # Extract information from the page
                phone_number, customer_name, message = extract_info_from_page(page_text)

                # If a new invoice is found, create a new PDF file
                if pdf_writer and output_pdf:
                    # Save the current PDF
                    with open(output_pdf, "wb") as output_file:
                        pdf_writer.write(output_file)

                    # Write the row to the CSV file
                    csv_writer.writerow([phone_number, customer_name, message, output_pdf])

                # Initialize a new PDF writer and set the output file path
                pdf_writer = PyPDF2.PdfWriter()
                output_pdf = os.path.join(output_folder, f"{invoice_number}.pdf")

                # Update the customer_name for the current row
                customer_name = customer_name.strip(" \n")

            # Add the current page to the current PDF
            pdf_writer.add_page(pdf_reader.pages[page_number])

        # Save the last PDF if any
        if pdf_writer and output_pdf:
            with open(output_pdf, "wb") as output_file:
                pdf_writer.write(output_file)

            # Write the last row to the CSV file
            csv_writer.writerow([phone_number, customer_name, message, output_pdf])

    # Show a success message
    messagebox.showinfo("Success", "PDFs distributed successfully in the desired folder and CSV file.") 


def browse_pdf(entry_pdf):
    # Opens a file dialog to browse and select a PDF file.
    # Updates the Entry widget with the selected file path.
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])

    # If a file is selected, update the entry widget with the file path
    if file_path:
        entry_pdf.delete(0, tk.END)
        entry_pdf.insert(0, file_path)

def browse_output_folder(entry_output_folder):
    # Opens a folder dialog to browse and select an output folder.
    # Updates the Entry widget with the selected folder path.

    # Open a folder dialog to choose an output folder
    folder_path = filedialog.askdirectory()

    # If a folder is selected, update the entry widget with the folder path
    if folder_path:
        entry_output_folder.delete(0, tk.END)
        entry_output_folder.insert(0, folder_path)

def browse_csv_file(entry_csv_file):
    # Opens a file dialog to browse and select a CSV file.
    # Updates the Entry widget with the selected file path.
    csv_file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])

    # If a file is selected, update the entry widget with the file path
    if csv_file_path:
        entry_csv_file.delete(0, tk.END)
        entry_csv_file.insert(0, csv_file_path)

def split_pdf_gui():
    global entry_pdf, entry_output_folder, entry_csv_file

    # Create a themed Tkinter window with Adapta theme
    window = ThemedTk(theme="equilux")

    # Set the window title
    window.title("PDF Splitter")

    # Set the window size
    window.geometry("500x400")

    # Calculate the position to center the window
    x = (window.winfo_screenwidth() - window.winfo_reqwidth()) / 2
    y = (window.winfo_screenheight() - window.winfo_reqheight()) / 2

    # Set the window position
    window.geometry("+%d+%d" % (x, y))

    # Create a label for the first step
    tk.Label(window, text="Step 1: Select PDF File", font=("Helvetica", 14), bg="#f0f0f0", fg="#2E2E2E").pack(pady=10)

    # Create an entry widget for the PDF file path
    entry_pdf = tk.Entry(window, width=50, font=("Helvetica", 12))
    entry_pdf.pack(pady=10)

    # Create a button to browse and select a PDF file
    tk.Button(window, text="Browse", command=lambda: browse_pdf(entry_pdf), font=("Helvetica", 12), relief=tk.GROOVE, padx=10, pady=5, borderwidth=5).pack()

    # Create a label for the second step
    tk.Label(window, text="Step 2: Select Output Folder", font=("Helvetica", 14), bg="#f0f0f0", fg="#2E2E2E").pack(pady=10)

    # Create an entry widget for the output folder path
    entry_output_folder = tk.Entry(window, width=50, font=("Helvetica", 12))
    entry_output_folder.pack(pady=10)

    # Create a button to browse and select an output folder
    tk.Button(window, text="Browse", command=lambda: browse_output_folder(entry_output_folder), font=("Helvetica", 12), relief=tk.GROOVE, padx=10, pady=5, borderwidth=5).pack()

    # Create a label for the third step
    tk.Label(window, text="Step 3: Select CSV File", font=("Helvetica", 14), bg="#f0f0f0", fg="#2E2E2E").pack(pady=10)

    # Create an entry widget for the CSV file path
    entry_csv_file = tk.Entry(window, width=50, font=("Helvetica", 12))
    entry_csv_file.pack(pady=10)

    # Create a button to browse and select a CSV file
    tk.Button(window, text="Browse", command=lambda: browse_csv_file(entry_csv_file), font=("Helvetica", 12), relief=tk.GROOVE, padx=10, pady=5, borderwidth=5).pack()

    # Create a button to initiate PDF splitting
    tk.Button(window, text="Distribute The Invoices", command=lambda: split_pdf(entry_pdf.get(), entry_output_folder.get()), font=("Helvetica", 14), bg="#2E2E2E", fg="#f0f0f0", relief=tk.GROOVE, padx=20, pady=10, borderwidth=5).pack(pady=20)

    # Start the Tkinter main loop
    window.mainloop()

# Check if the script is run as the main program
if __name__ == "__main__":
    
    # Call the main GUI function
    split_pdf_gui()