Route Sheet Automation
Route Sheet Automation is a Python-based desktop application designed to simplify the process of creating, processing, and printing route sheets for scouting organizations. The application leverages OCR technology to extract data from receipts and automatically generates route sheets in Excel format. It also includes features to view generated files, print them, and streamline workflow.

Features
Receipt Processing: Extract data from image or PDF receipts using OCR (Optical Character Recognition).
Automated Route Sheet Creation: Generate customized route sheets in Excel format with pre-filled data.
Multiple File Handling: Process multiple receipts and generate corresponding route sheets in a single session.
User-Friendly Interface:
Select and process receipts easily.
View processed data in a clean, readable format.
Clickable links to open generated route sheets directly.
Printing Support: Print all generated route sheets with a single click, automatically formatted to fit on one page (A1:K44).
Customizable Layout: Green button for route sheet creation, intuitive file selection, and processing workflow.
Requirements
Python 3.8+
Required libraries (install via pip):
customtkinter
pillow
pytesseract
pdf2image
win32com
Microsoft Excel (required for printing functionality)
Installation
Clone the repository:

bash
Copy code
git clone https://github.com/vibrantantelope/routesheetassistant.git
cd routesheetassistant
Set up a virtual environment:

bash
Copy code
python -m venv venv
source venv/bin/activate   # For Linux/Mac
venv\Scripts\activate      # For Windows
Install dependencies:

bash
Copy code
pip install -r requirements.txt
Install Tesseract OCR:

Windows: Download and install
Linux/Mac: Use your package manager (e.g., sudo apt install tesseract-ocr)
Usage
Launch the application:

bash
Copy code
python main.py
Select receipt files:

Click Select Receipts to upload image or PDF files.
Supported formats: .png, .jpg, .jpeg, .tiff, .bmp, .pdf.
Process receipts:

Click Process Receipts to extract data and preview results.
Create route sheets:

Click Create Route Sheet(s) to generate Excel files based on processed data.
Open or print route sheets:

Click the Open link for a file to view it.
Click Print All Route Sheets to print the generated files.
File Structure
main.py: Entry point for the application.
gui.py: Contains the graphical user interface logic.
receipt_processing.py: Handles OCR processing and receipt data extraction.
route_sheet.py: Contains logic for generating and updating route sheets.
assets/: Contains templates and generated files.
data/: Stores debug logs and intermediate outputs.
Contributions
Contributions, issues, and feature requests are welcome! Feel free to fork the repository and submit a pull request.

License
This project is licensed under the MIT License. See the LICENSE file for more information.

Acknowledgements
Tesseract OCR for text extraction.
CustomTkinter for the modern GUI.
Microsoft Excel and win32com for printing support.
Screenshots
Main Interface

Processed Data and Links

Printable Route Sheets

With Route Sheet Automation, managing your scouting route sheets has never been easier. 🚀
