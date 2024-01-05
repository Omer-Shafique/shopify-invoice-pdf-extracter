import csv
import PyPDF2
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from ttkthemes import ThemedTk
import threading
import phonenumbers

processed_invoices = set()

def extract_invoice_number(page_text):
    match = re.search(r'INVOICE NO: (\d+)', page_text)
    return match.group(1) if match else None

def extract_info_from_page(page_text):
    # Extract all phone numbers
    raw_phone_numbers = re.findall(r'\b\d{9,12}\b', page_text)
    
    # Extract message
    message_match = re.search(r'This Is Your Invoice', page_text)
    message = "This Is Your Invoice" if message_match else None

    # Extract customer name from "BILL TO" and "SHIP TO" sections
    bill_to_match = re.search(r'BILL TO[^\n]*\n(.*?)(?=\n)', page_text, re.DOTALL)
    ship_to_match = re.search(r'SHIP TO[^\n]*\n(.*?)(?=\n)', page_text, re.DOTALL)

    customer_name = None

    if bill_to_match:
        customer_name = clean_customer_name(bill_to_match.group(1).strip())
    elif ship_to_match:
        customer_name = clean_customer_name(ship_to_match.group(1).strip())
    
    formatted_phone_numbers = []
    phone_numbers_included_in_csv = []

    for raw_phone_number in raw_phone_numbers:
        try:
            parsed_number = phonenumbers.parse(raw_phone_number, None)
            formatted_phone_numbers.append(phonenumbers.format_number(parsed_number, phonenumbers.PhoneNumberFormat.E164))
        except Exception as e:
            phone_numbers_included_in_csv.append(raw_phone_number)
            print(f"Including phone number: {raw_phone_number}. Exception: {e}")

    # return formatted_phone_numbers, customer_name, message, phone_numbers_included_in_csv
    return phone_numbers_included_in_csv, customer_name, message

def write_to_csv(data, filename='output.csv'):
    with open(filename, 'w', newline='') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerow(['Formatted Phone Numbers', 'Customer Name', 'Message', 'Ignored Phone Numbers'])
        csv_writer.writerow(data)

def format_phone_number(phone_number):
    # Add any additional formatting logic for phone numbers if needed
    return phone_number



def clean_customer_name(raw_name):
    words = raw_name.split()
    unique_words = []
    for word in words:
        cleaned_word = word.lower().title()  # Convert to title case
        if cleaned_word not in unique_words:
            unique_words.append(cleaned_word)
    cleaned_name = ' '.join(unique_words)
    return cleaned_name


def split_pdf(input_pdf, output_folder):
    global processed_invoices

    pdf_reader = PyPDF2.PdfReader(input_pdf)
    csv_file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])

    if not csv_file_path:
        return

    try:
        with open(csv_file_path, 'w', newline='', encoding='utf-8-sig', errors='replace') as csvfile:
            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(['Phone Number', 'Customer Name', 'Message', 'PDF Path'])

            pdf_writer = PyPDF2.PdfWriter()  # Initialize pdf_writer outside the loop
            output_pdf = None  
            phone_number, customer_name, message, output_pdf = "", "", "", ""  # Initialize variables with default values

            for page_number in range(len(pdf_reader.pages)):
                page_text = pdf_reader.pages[page_number].extract_text()
                invoice_number = extract_invoice_number(page_text)

                if invoice_number and invoice_number not in processed_invoices:
                    processed_invoices.add(invoice_number)

                    if output_pdf:
                        with open(output_pdf, "wb") as output_file:
                            pdf_writer.write(output_file)

                        csv_writer.writerow([phone_number, customer_name, message, output_pdf])

                    output_pdf = os.path.join(output_folder, f"{invoice_number}.pdf")
                    customer_name = None

                    # Extract phone number and customer name
                    phone_number, customer_name, message = extract_info_from_page(page_text)
                    customer_name = customer_name.strip(" \n")

                    pdf_writer = PyPDF2.PdfWriter()
                
                pdf_writer.add_page(pdf_reader.pages[page_number])

            # Handle the last invoice outside the loop
            if pdf_writer.pages:
                with open(output_pdf, "wb") as output_file:
                    pdf_writer.write(output_file)

                csv_writer.writerow([phone_number, customer_name, message, output_pdf])

        messagebox.showinfo("Success", "PDFs distributed successfully in the desired folder and CSV file.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


def browse_pdf(entry_pdf):
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        entry_pdf.delete(0, tk.END)
        entry_pdf.insert(0, file_path)

def browse_output_folder(entry_output_folder):
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_output_folder.delete(0, tk.END)
        entry_output_folder.insert(0, folder_path)

def split_pdf_gui():
    global entry_pdf, entry_output_folder

    window = ThemedTk(theme="equilux")
    window.title("PDF Splitter")
    window.geometry("500x500")
    x = (window.winfo_screenwidth() - window.winfo_reqwidth()) / 2
    y = (window.winfo_screenheight() - window.winfo_reqheight()) / 2
    window.geometry("+%d+%d" % (x, y))

    tk.Label(window, text="Select PDF File Of Shopify Invoices", font=("Helvetica", 14), bg="#f0f0f0", fg="#2E2E2E").pack(pady=10)

    entry_pdf = tk.Entry(window, width=50, font=("Helvetica", 12))
    entry_pdf.pack(pady=10)

    tk.Button(window, text="Browse", command=lambda: browse_pdf(entry_pdf), font=("Helvetica", 12), relief=tk.GROOVE, padx=10, pady=5, borderwidth=5).pack()

    tk.Label(window, text="Select Output Folder For Distributed PDFs \n(Empty Folder Recommended)", font=("Helvetica", 14), bg="#f0f0f0", fg="#2E2E2E").pack(pady=10)

    entry_output_folder = tk.Entry(window, width=50, font=("Helvetica", 12))
    entry_output_folder.pack(pady=10)

    tk.Button(window, text="Browse", command=lambda: browse_output_folder(entry_output_folder), font=("Helvetica", 12), relief=tk.GROOVE, padx=10, pady=5, borderwidth=5).pack()

    tk.Button(window, text="Create Excel Sheet", command=lambda: threading.Thread(target=lambda: split_pdf(entry_pdf.get(), entry_output_folder.get())).start(), font=("Helvetica", 14), bg="#2E2E2E", fg="#f0f0f0", relief=tk.GROOVE, padx=20, pady=10, borderwidth=5).pack(pady=20)

    window.mainloop()

if __name__ == "__main__":
    split_pdf_gui()