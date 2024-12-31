import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from receipt_processing import extract_receipt_text_to_json
from route_sheet import update_route_sheet_from_json


class RouteSheetApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Route Sheet Assistant")
        self.root.geometry("1024x768")
        self.receipt_data_list = []
        self.file_paths = []

        # Set appearance
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.create_widgets()

    def create_widgets(self):
        # Main container
        container = ctk.CTkFrame(self.root)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        # Header
        header = ctk.CTkLabel(
            container,
            text="Create Route Sheets from Images or PDFs",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        header.pack(pady=10)

        # Top button section
        button_frame = ctk.CTkFrame(container)
        button_frame.pack(pady=10)

        self.file_button = ctk.CTkButton(
            button_frame,
            text="Select Receipts",
            command=self.select_receipts,
            height=40,
            width=200
        )
        self.file_button.pack(side="left", padx=10)

        self.process_button = ctk.CTkButton(
            button_frame,
            text="Process Receipts",
            command=self.process_receipts,
            height=40,
            width=200,
            state="disabled"
        )
        self.process_button.pack(side="left", padx=10)

        self.create_button = ctk.CTkButton(
            button_frame,
            text="Create Route Sheet(s)",
            command=self.create_route_sheets,
            height=40,
            width=200,
            fg_color="green",
            hover_color="dark green",
            state="disabled"
        )
        self.create_button.pack(side="left", padx=10)

        # Status indicator
        self.status = ctk.CTkLabel(
            container,
            text="No files selected",
            text_color="gray"
        )
        self.status.pack(pady=5)

        # Data display section
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
            height=400,
            width=800
        )
        self.data_display.pack(pady=10, padx=20)

    def select_receipts(self):
        file_paths = filedialog.askopenfilenames(
            title="Select Receipt Documents",
            filetypes=[("Image and PDF Files", "*.png;*.jpg;*.jpeg;*.tiff;*.bmp;*.pdf")]
        )

        if not file_paths:
            return

        self.file_paths = file_paths
        self.status.configure(
            text=f"{len(file_paths)} file(s) selected",
            text_color="green"
        )

        # Enable the "Process Receipts" button
        self.process_button.configure(state="normal")

    def process_receipts(self):
        self.receipt_data_list.clear()
        self.data_display.delete("1.0", "end")

        for file_path in self.file_paths:
            try:
                # Extract receipt data
                receipt_data = extract_receipt_text_to_json(file_path)
                self.receipt_data_list.append({"file_path": file_path, "data": receipt_data})

                # Format and display the receipt data
                formatted_data = self.format_receipt_data(receipt_data)
                self.data_display.insert(
                    "end",
                    f"File: {os.path.basename(file_path)}\n{formatted_data}\n{'-'*80}\n"
                )
            except Exception as e:
                self.data_display.insert(
                    "end",
                    f"Error processing {os.path.basename(file_path)}: {str(e)}\n{'-'*80}\n"
                )

        # Enable the "Create Route Sheet(s)" button if receipts were successfully processed
        if self.receipt_data_list:
            self.create_button.configure(state="normal")

    def create_route_sheets(self):
        for receipt in self.receipt_data_list:
            try:
                # Update route sheet and get output path
                output_path = update_route_sheet_from_json(receipt["data"])
                file_name = os.path.basename(receipt["file_path"])
                self.data_display.insert(
                    "end",
                    f"Route sheet created for {file_name}: {output_path}\n{'='*80}\n"
                )
            except Exception as e:
                self.data_display.insert(
                    "end",
                    f"Error creating route sheet for {receipt['file_path']}: {str(e)}\n"
                )

        messagebox.showinfo("Success", "All route sheets created successfully!")

    def format_receipt_data(self, data):
        """
        Format receipt data dictionary for display in the GUI.
        """
        formatted_data = "\n".join(
            [f"{key.replace('_', ' ').capitalize()}: {value}" for key, value in data.items()]
        )
        return formatted_data

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = RouteSheetApp()
    app.run()
