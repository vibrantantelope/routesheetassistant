import pytesseract
from pytesseract import Output
from PIL import Image, ImageEnhance, ImageFilter
from pdf2image import convert_from_path
import json
import os
from datetime import datetime, timedelta
import re
import logging

# Set up logging
logging.basicConfig(
    filename="data/extract_receipt_debug.log",
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# Ensure data folder exists
os.makedirs("data", exist_ok=True)

def preprocess_image(image):
    """
    Preprocess the image to improve OCR accuracy.
    """
    logging.info("Preprocessing image for OCR.")
    # Convert to grayscale
    image = image.convert("L")
    # Resize the image for better OCR
    image = image.resize((image.width * 2, image.height * 2), Image.Resampling.LANCZOS)  # Use LANCZOS instead of ANTIALIAS
    # Enhance contrast
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(3)
    # Apply sharpening filter
    image = image.filter(ImageFilter.SHARPEN)
    logging.info("Image preprocessing complete.")
    return image

def extract_receipt_text_to_json(receipt_path):
    """
    Extract text from receipt (PDF or image) using OCR and save it as a JSON file.
    """
    try:
        logging.info(f"Processing receipt: {receipt_path}")

        # Convert PDF to image if necessary
        if receipt_path.lower().endswith(".pdf"):
            logging.info("Detected PDF. Converting to image.")
            try:
                images = convert_from_path(receipt_path, dpi=400)
                image = images[0]  # Use the first page
            except Exception as e:
                logging.error(f"Error converting PDF to image: {e}")
                raise RuntimeError("Failed to process receipt: Unable to convert PDF to image.")
        else:
            # Load image directly
            logging.info("Loading image file.")
            image = Image.open(receipt_path)

        # Preprocess the image
        image = preprocess_image(image)

        # Perform OCR on the image
        text = pytesseract.image_to_string(image, lang='eng', config='--psm 4')

        # Debugging: Save raw OCR output for review
        with open("data/raw_ocr_output.txt", "w") as f:
            f.write(text)
        logging.debug("OCR output saved to raw_ocr_output.txt.")

        # Initialize data dictionary
        data = {
            "council_number": "456",  # Always the same
            "effective_date": datetime.now().strftime("%Y-%m-01"),  # First day of the current month
            "term": "12 months",  # Always 12 months
        }

        # Extract district, unit, program type, and price information
        district_map = {
            "Calumet": 1, "Prairie Dunes": 3, "Thunderbird": 4, "Checaugau": 5, 
            "Iron Horse": 6, "Tri-Star": 7, "Five Creeks": 9, "Tall Grass": 11, "Trailblazer": 12
        }

        price_fields = {
            "Youth Registration": "C9",
            "Youth SL Subscription": "C10",
            "Youth Transfer": "C11",
            "Adult Registration": "C12",
            "Multiple/Position Change": "C13",
            "Adult Transfer": "C14",
            "Adult SL Subscription": "C15",
            "Youth Exploring": "C16",
            "Adult Exploring": "C17",
            "Program Fee": "C18",
        }

        # Prices to map
        prices = {field: 0 for field in price_fields.keys()}

        # Process each line for data
        logging.info("Processing OCR lines for data extraction.")
        lines = text.splitlines()
        for line in lines:
            logging.debug(f"Processing line: {line}")
            line = line.strip()

            # Extract district
            for district, number in district_map.items():
                if district.lower() in line.lower():
                    data["district_name"] = district
                    data["district_number"] = number
                    logging.info(f"Matched district: {district} ({number})")

            # Extract local unit number
            if "Troop" in line or "Pack" in line or "Crew" in line or "Ship" in line or "Post" in line:
                match = re.search(r"(\d+)", line)
                if match:
                    data["local_unit_number"] = match.group(1)
                    logging.info(f"Matched local unit number: {data['local_unit_number']}")

            # Extract program type
            if "Scouts BSA" in line or "Troop" in line:
                data["program"] = "Scouts BSA"
            elif "Cub Scouts" in line or "Pack" in line:
                data["program"] = "Cub Scouts"
            elif "Venturing" in line or "Crew" in line:
                data["program"] = "Venturing"
            elif "Sea Scouts" in line or "Ship" in line:
                data["program"] = "Sea Scouts"
            elif "Exploring" in line or "Post" in line:
                data["program"] = "Exploring"

            # Extract price information
                       # Extract price information
            match = re.search(r"(\d+)\s+(Youth BL|Youth Renewal|Adult Renewal|Adult New|Youth Program Fee|Adult Program Fee)", line, re.IGNORECASE)
            if match:
                count = int(match.group(1))
                label = match.group(2).strip().lower()
                logging.info(f"Matched price line: {label} ({count})")

                if "youth bl" in label:
                    prices["Youth SL Subscription"] += count
                elif "youth renewal" in label or "youth new" in label:
                    prices["Youth Registration"] += count
                elif "adult renewal" in label or "adult new" in label:
                    prices["Adult Registration"] += count
                elif "program fee" in label:
                    prices["Program Fee"] += count


        # Add prices to data
        data["prices"] = prices

        # Infer expiration date
        # Infer expiration date as 12 months minus 1 day
        effective_date = datetime.strptime(data["effective_date"], "%Y-%m-%d")
        # Add 12 months minus 1 day
        expiration_date = effective_date + timedelta(days=365) - timedelta(days=1)
        data["expiration_date"] = expiration_date.strftime("%Y-%m-%d")


        # Save data to JSON
        json_path = "data/receipt_data.json"
        with open(json_path, "w") as json_file:
            json.dump(data, json_file, indent=4)
        logging.info("Data successfully extracted and saved to JSON.")

        return data

    except Exception as e:
        logging.error(f"Error processing receipt: {e}")
        raise RuntimeError(f"Failed to process receipt: {e}")
