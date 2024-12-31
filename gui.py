import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import webbrowser
from receipt_processing import extract_receipt_text_to_json
from route_sheet import update_route_sheet_from_json
import win32com.client


class RouteSheetApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Route Sheet Assistant")
        self.root.geometry("900x700")

        # Set appearance
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.selected_files = []  # List of selected receipt files
        self.generated_files = []  # List of generated route sheet files

        self.create_widgets()

    def create_widgets(self):
        # Main container
        container = ctk.CTkFrame(self.root)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        # Header
        header = ctk.CTkLabel(
            container,
            text="Upload Image and PDF Receipts to Create Route Sheets",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        header.pack(pady=20)

        # Button frame for Select, Process, and Create
        button_frame = ctk.CTkFrame(container)
        button_frame.pack(fill="x", pady=10)

        # Select Receipts Button
        self.select_button = ctk.CTkButton(
            button_frame,
            text="Select Receipts",
            command=self.select_receipts,
            height=40
        )
        self.select_button.pack(side="left", padx=10)

        # Process Receipts Button
        self.process_button = ctk.CTkButton(
            button_frame,
            text="Process Receipts",
            command=self.process_receipts,
            height=40
        )
        self.process_button.pack(side="left", padx=10)

        # Create Route Sheets Button (Green)
        self.create_button = ctk.CTkButton(
            button_frame,
            text="Create Route Sheet(s)",
            command=self.process_receipts,
            height=40,
            fg_color="green",
            hover_color="dark green"
        )
        self.create_button.pack(side="left", padx=10)

        # Status indicator
        self.status = ctk.CTkLabel(
            container,
            text="No files selected",
            text_color="gray"
        )
        self.status.pack(pady=5)

        # Data display frame
        self.data_frame = ctk.CTkFrame(container)
        self.data_frame.pack(fill="both", expand=True, pady=20)

        data_label = ctk.CTkLabel(
            self.data_frame,
            text="Processed Receipt Data and Outputs",
            font=ctk.CTkFont(weight="bold")
        )
        data_label.pack(pady=10)

        self.data_display = ctk.CTkTextbox(
            self.data_frame,
            height=300,
            width=600
        )
        self.data_display.pack(pady=10, padx=20)

        # Generated links container
        self.links_frame = ctk.CTkFrame(container)
        self.links_frame.pack(fill="x", pady=10)

        # Print Button
        self.print_button = ctk.CTkButton(
            container,
            text="Print All Route Sheets",
            command=self.print_route_sheets,
            fg_color="blue",
            hover_color="dark blue",
            height=40
        )
        self.print_button.pack(pady=10)
        self.print_button.configure(state="disabled")  # Disabled by default

    def select_receipts(self):
        file_paths = filedialog.askopenfilenames(
            title="Select Receipt Documents",
            filetypes=[("Image and PDF Files", "*.png;*.jpg;*.jpeg;*.tiff;*.bmp;*.pdf")]
        )

        if not file_paths:
            return

        self.selected_files = file_paths
        self.status.configure(
            text=f"{len(file_paths)} file(s) selected",
            text_color="green"
        )
        self.process_button.configure(state="normal")  # Enable process button

    def process_receipts(self):
        self.generated_files = []  # Reset generated files list
        self.data_display.delete("1.0", "end")  # Clear data display

        for file_path in self.selected_files:
            try:
                receipt_data = extract_receipt_text_to_json(file_path)
                output_path = update_route_sheet_from_json(receipt_data)
                self.generated_files.append(output_path)

                # Display receipt data
                formatted_data = self.format_receipt_data(receipt_data)
                self.data_display.insert("end", formatted_data + "\n")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to process {os.path.basename(file_path)}: {str(e)}")

        # Add clickable links for generated files
        if self.generated_files:
            for file in self.generated_files:
                self.add_clickable_link(file)

            self.print_button.configure(state="normal")  # Enable print button

    def add_clickable_link(self, file_path):
        """
        Add a clickable link to the links frame for the generated route sheet.
        """
        short_file_name = os.path.basename(file_path)  # Abbreviate the file path
        link_button = ctk.CTkButton(
            self.links_frame,
            text=f"Open {short_file_name}",
            command=lambda: webbrowser.open(file_path),
            height=30
        )
        link_button.pack(pady=5, padx=10, side="top")

    def format_receipt_data(self, data):
        """
        Format receipt data dictionary for display in the GUI.
        """
        formatted_data = "\n".join([f"{key}: {value}" for key, value in data.items()])
        return formatted_data

    def print_route_sheets(self):
        """
        Print all generated route sheets, restricting the print range to A1:K44
        and ensuring it fits on one page.
        """
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        try:
            for file in self.generated_files:
                workbook = excel.Workbooks.Open(file)
                sheet = workbook.Sheets(1)

                # Set the print area and fit to page
                sheet.PageSetup.PrintArea = "A1:K44"
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = 1

                # Print the sheet
                workbook.PrintOut()
                workbook.Close(SaveChanges=False)

            messagebox.showinfo("Success", "All route sheets have been printed successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to print route sheets: {str(e)}")

        finally:
            excel.Quit()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = RouteSheetApp()
    app.run()
