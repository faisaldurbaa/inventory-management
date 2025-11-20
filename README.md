# Inventory Management System (v3.7.0) Nov-2023

This project features a modern, GUI-based Inventory Management System developed for a Turkish school. It offers comprehensive stock control, transaction tracking, and reporting functionalities through an intuitive graphical interface.

## Features

*   **Graphical User Interface (GUI):** Built with `tkinter` and `CustomTkinter` for a user-friendly and visually appealing experience.
*   **Excel-based Data Storage:** All inventory data is securely stored and managed in a single Excel file (`database.xlsx`) across three distinct sheets:
    *   **İşlemler (Transactions):** Records all product movements (additions, sales, edits).
    *   **Stok (Stock):** Maintains the current stock levels for all products.
    *   **Geçmiş (History):** Logs all modifications and deletions of transactions.
*   **Comprehensive Product Operations:**
    *   **Ürün Girişi (Product Entry):** Add new products to the inventory or increase the quantity of existing items.
    *   **Ürün Çıkışı (Product Sale/Output):** Process product sales, with real-time checks for available stock to prevent overselling.
    *   **İşlemi Düzenle (Edit Operation):** Modify details of past transactions. This intelligently updates stock levels and records the modification in the transaction history.
    *   **İşlemi Sil (Delete Operation):** Remove a transaction, which automatically reverses its impact on the stock and logs the deletion for auditing purposes.
*   **Dynamic Data View:** View and navigate through "İşlemler", "Stok", and "Geçmiş" data directly within the application's interactive tables.
*   **Automated Reporting:** Generate detailed Excel reports containing all transaction, stock, and history data at any time. Reports are timestamped for easy organization.
*   **Smart Archiving:** The system automatically archives transaction data into a new report file and clears the "İşlemler" sheet if it exceeds 500 entries, ensuring optimal performance and data manageability while preserving historical records.
*   **Unique Identifiers:** Automatic generation of unique product codes (e.g., "AA01") and sequential operation numbers for efficient tracking.
*   **Robust Error Handling:** Provides clear pop-up error messages for invalid inputs, insufficient stock, or other operational issues.

## Technologies Used

*   Python
*   Tkinter (for GUI)
*   CustomTkinter (for modern GUI widgets)
*   Openpyxl (for Excel file handling)
*   Pillow (PIL - for image handling, specifically the logo)

## How to Run

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/faisaldurbaa/inventory-management.git
    cd inventory-management
    ```
2.  **Install dependencies:**
    ```bash
    pip install openpyxl customtkinter pillow
    ```
3.  **Run the application:**
    ```bash
    python app.py
    ```

## How to Create a Standalone Executable (.exe)

You can convert the `app.py` script into a single executable file for Windows using `PyInstaller`. This will allow the application to run without needing a Python environment installed on the user's machine, and without a console window opening.

1. **Install PyInstaller:**

    ```bash
    pip install pyinstaller
    ```

2. **Create the executable:**
    Navigate to the project root directory in your terminal and run the following command. The `--noconsole` flag ensures no terminal window appears, and `--onefile` packages everything into a single executable.

    ```bash
    pyinstaller --noconsole --onefile --add-data "VDFL_logo.png;." app.py
    ```

    This will generate the executable in the `dist` folder.
