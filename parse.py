import pandas as pd
import json
import logging

def parse_excel_to_json(excel_file, output_file):
    """
    Parses the Excel workbook and generates a JSON file with the timetable data.
    """
    logging.info("Reading Excel file...")
    excel = pd.ExcelFile(excel_file)
    sheet_names = excel.sheet_names
    timetable_data = []

    for sheet_name in sheet_names:
        logging.info(f"Processing sheet: {sheet_name}")
        df = pd.read_excel(excel, sheet_name=sheet_name, header=None)
        course_data = parse_course_sheet(df)
        if course_data:  # Ensure course_data is not empty
            timetable_data.append(course_data)

    # Write to JSON file
    logging.info(f"Writing data to {output_file}...")
    with open(output_file, 'w') as json_file:
        json.dump(timetable_data, json_file, indent=4)

    logging.info("Parsing completed successfully.")

def parse_course_sheet(df):
    """
    Parses a single course sheet and extracts course information.
    """
    # Initialize dictionary to hold course data
    course_info = {}
    sections = {}

    # Fill merged cells downwards
    df = df.fillna(method="ffill")

    # Skip header rows
    start_row = find_data_start_row(df)
    if start_row is None:
        logging.error("Could not find the start of data in the sheet.")
        return None  # Return None if no data found

    # Extract course information from the merged cells
    # Update the column indices to correctly extract course_code and course_title
    course_info["course_code"] = df.iloc[start_row, 1]  # Assuming course_code is in column 1
    course_info["course_title"] = df.iloc[start_row, 2]  # Course title in column 2
    course_info["credits"] = {
        "lecture": df.iloc[start_row, 3],
        "practical": df.iloc[start_row, 4],
        "units": df.iloc[start_row, 5],
    }

    # Columns mapping based on your layout
    col_sec = 6
    col_instr = 7
    col_room = 8
    col_time = 9

    # Process the rest of the data to extract sections
    logging.info("Extracting sections...")
    idx = start_row
    while idx < len(df):
        row = df.iloc[idx]

        # Identify section rows
        section_id = row[col_sec]
        if pd.notnull(section_id):
            section_key = (section_id, get_section_type(section_id))
            if section_key not in sections:
                section = {}
                section["section_type"] = get_section_type(section_id)
                section["section_number"] = section_id
                section["instructors"] = []
                section["room"] = (
                    str(row[col_room]) if pd.notnull(row[col_room]) else ""
                )
                section["timing"] = []
                sections[section_key] = section
            else:
                section = sections[section_key]

            # Add instructor
            if pd.notnull(row[col_instr]):
                instructor = row[col_instr]
                if instructor not in section["instructors"]:
                    section["instructors"].append(instructor)

            # Add time slots
            time_slots = parse_time_slots(row[col_time])
            for time_slot in time_slots:
                if time_slot not in section["timing"]:
                    section["timing"].append(time_slot)

            # Check for additional instructors in subsequent rows
            next_idx = idx + 1
            while (
                next_idx < len(df) and
                pd.isnull(df.iloc[next_idx, col_sec]) and
                pd.isnull(df.iloc[next_idx, 0]) and
                pd.isnull(df.iloc[next_idx, col_sec])
            ):
                next_row = df.iloc[next_idx]
                # Add instructor
                if pd.notnull(next_row[col_instr]):
                    instructor = next_row[col_instr]
                    if instructor not in section["instructors"]:
                        section["instructors"].append(instructor)
                # Add time slots
                time_slots = parse_time_slots(next_row[col_time])
                for time_slot in time_slots:
                    if time_slot not in section["timing"]:
                        section["timing"].append(time_slot)
                next_idx += 1

            idx = next_idx
        else:
            idx += 1

    course_info["sections"] = list(sections.values())
    return course_info

def find_data_start_row(df):
    """
    Finds the index of the row where actual data starts.
    """
    for idx in range(len(df)):
        value = df.iloc[idx, 0]
        if pd.notnull(value):
            if is_course_code(value):
                return idx
    return None

def is_course_code(value):
    """
    Checks if the value is a valid course code.
    """
    try:
        int_value = int(value)
        return True
    except (ValueError, TypeError):
        return False

def get_section_type(section_id):
    """
    Determines the section type based on the section identifier.
    """
    if isinstance(section_id, str):
        if section_id.startswith('L'):
            return 'lecture'
        elif section_id.startswith('T'):
            return 'tutorial'
        elif section_id.startswith('P'):
            return 'practical'
    return 'Unknown'

def parse_time_slots(time_str):
    """
    Parses time slot strings and converts them to a list of day-slot combinations,
    including actual time ranges.
    """
    time_slots = []
    if pd.isnull(time_str):
        return time_slots

    # Normalize the time string by replacing multiple spaces with a single space
    time_str = ' '.join(time_str.strip().split())

    # Split into parts
    parts = time_str.split()

    valid_days = ['M', 'T', 'W', 'Th', 'F', 'S']
    time_mapping = {
        1: '8-9',
        2: '9-10',
        3: '10-11',
        4: '11-12',
        5: '12-1',
        6: '2-3',
        7: '3-4',
        8: '4-5',
        9: '5-6'
    }
    idx = 0
    while idx < len(parts):
        day_group = []
        slot_group = []
        # Collect days
        while idx < len(parts) and parts[idx] in valid_days:
            day_group.append(parts[idx])
            idx += 1
        # Collect slots
        while idx < len(parts) and parts[idx] not in valid_days:
            slot_part = ''.join(filter(str.isdigit, parts[idx]))
            if slot_part.isdigit():
                slot_group.append(int(slot_part))
            else:
                logging.warning(f"Unexpected slot value '{parts[idx]}' in time string '{time_str}'")
            idx += 1
        # Combine days and slots
        for day in day_group:
            time_slot = {
                'day': day,
                'slots': slot_group,
                'timings': [time_mapping.get(slot, 'Unknown') for slot in slot_group]
            }
            time_slots.append(time_slot)
    return time_slots

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(message)s')
    try:
        excel_file = 'Timetable Workbook - SUTT Task 1.xlsx'
        output_file = 'timetable_data.json'
        parse_excel_to_json(excel_file, output_file)
    except Exception as e:
        logging.error(f"An error occurred: {e}")