# Lathe Workshop Invoice System

A lightweight, user-friendly invoice management system designed to run on Raspberry Pi or Windows using Python and Excel. Perfect for small workshops like lathe machine operations to easily create, save, retrieve, and print invoices with digital signatures.

---

## Features

- **Automatic Sequential Bill Numbers**: The system auto-generates unique invoice IDs starting from INV001.
- **Vehicle Number Input**: Record vehicle details for each invoice.
- **Multiple Jobs and Amounts**: Add any number of jobs or repairs with respective costs.
- **Auto-calculation of Total Amount**: Total sum updates automatically as jobs are added.
- **Signature Capture**: Sign directly on the canvas area to authorize the invoice.
- **Save & Retrieve Invoices**: All data saved in a single Excel file (`invoices.xlsx`), with easy retrieval by Bill Number.
- **Printable Invoice Preview**: View a formatted invoice in a popup for easy printing.
- **Clean and Simple UI**: Intuitive interface built with Tkinter for quick data entry and management.

---

## Installation

1. Clone the repository or download the script.

2. Install Python 3 (if not installed).

3. Install dependencies using pip in Terminal:
     ```  bash
     pip install openpyxl
     ```


---

## Usage

Run the program in Terminal:   
  ```  bash
python invoice_system.py
  ```


- The app will generate unique bill numbers automatically.
- Enter vehicle number, add jobs and amounts one by one.
- Draw signature in the signature area.
- Save invoices to the Excel file.
- You can retrieve previous invoices using their bill number.
- Preview invoices before printing using the print preview button.
- Clear fields easily using the clear button.

---

## File Structure

- `invoice_system.py` — Main Python application.
- `invoices.xlsx` — Excel file storing all invoice data (generated automatically).
- `signature_INVxxx.ps` — Postscript file storing signature images per invoice.

---
