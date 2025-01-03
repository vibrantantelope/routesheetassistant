import logging
from openpyxl import load_workbook


def update_route_sheet_from_json(data):
    """
    Update the route sheet based on extracted JSON data.
    """
    # Load the route sheet template
    template_path = r"Z:\johnv\routesheetassistant\assets\RouteSheetTemplateV2.xlsx"
    workbook = load_workbook(template_path)
    sheet = workbook.active

    # Format dates
    effective_date = data["effective_date"]
    expiration_date = data["expiration_date"]

    # Ensure dates are in MM/DD/YYYY format
    effective_date_formatted = "/".join(effective_date.split("-")[1:] + [effective_date.split("-")[0]])
    expiration_date_formatted = "/".join(expiration_date.split("-")[1:] + [expiration_date.split("-")[0]])

    # Map extracted fields to cells
    sheet["B4"] = data["program"]
    sheet["C4"] = data["council_number"]
    sheet["D4"] = data["district_number"]
    sheet["G4"] = data["local_unit_number"]
    sheet["H4"] = effective_date_formatted
    sheet["I4"] = data["term"]
    sheet["J4"] = expiration_date_formatted

    # Write the program type (Troop, Pack, etc.) into E4
    logging.info("Updating E4 with the unit type based on program.")
    program_to_unit_type = {
        "Scouts BSA": "Troop",
        "Cub Scouts": "Pack",
        "Venturing": "Crew",
        "Sea Scouts": "Ship",
        "Exploring": "Post",
        "District": "Non-Unit",
        "Council": "Non-Unit"
    }
    unit_type = program_to_unit_type.get(data["program"], "Unknown")
    sheet["E4"] = unit_type

    # Map prices to their respective cells
    logging.info("Mapping prices to the correct cells.")
    for field, cell in {
        "Unit Charter": "C8",
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
    }.items():
        # Only update the price cells explicitly defined
        if field in data["prices"]:
            sheet[cell] = data["prices"].get(field, 0)

    # Save the updated route sheet with a new name
    district_name = data.get("district_name", "Unknown").replace(" ", "_")
    local_unit_number = data.get("local_unit_number", "Unknown")
    current_date = effective_date_formatted.replace("/", "-")
    output_path = rf"Z:\johnv\routesheetassistant\assets\Route_Sheet_{district_name}_{local_unit_number}_{current_date}.xlsx"
    workbook.save(output_path)
    logging.info(f"Route sheet successfully updated and saved to {output_path}.")
    return output_path
