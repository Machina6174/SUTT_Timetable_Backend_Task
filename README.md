# Timetable Excel Parsing - Task 1

## Author

**Daksh Gargi**

## Description

This project involves parsing an Excel workbook containing the institute's timetable data and generating a JSON file based on the extracted information. Each sheet in the workbook represents data about a single course.

The script:

- Reads the Excel workbook using pandas.
- Extracts relevant data such as course codes, titles, credits, sections, instructors, rooms, days, slots, and timings.
- Handles complex time strings and maps slot numbers to actual time ranges.
- Generates a JSON file (`timetable_data.json`) containing the structured timetable data.

## Instructions

1. **Install Dependencies:**

    ```bash
    pip install pandas
    ```

2. **Prepare the Files:**

    - Place `Timetable Workbook - SUTT Task 1.xlsx` in the same directory as `parse.py`.

3. **Run the Script:**

    ```bash
    python parse.py
    ```

4. **Output:**

    - The JSON output will be generated as `timetable_data.json` in the same directory.

## Notes

- **Time Slot Parsing:**
  
  - The script handles complex time strings (e.g., `"T Th F  2"`) and correctly parses them into individual day and slot combinations.
  - Slot numbers are mapped to actual time ranges (e.g., slot 1 corresponds to `"8-9"`).

- **Data Integrity:**
  
  - Ensure the Excel file is properly formatted and placed correctly.
  - The script fills merged cells appropriately and combines multiple instructors and time slots into single entries for each section.

- **Logging:**
  
  - The script uses the `logging` module to provide informative messages during execution.

## Dependencies

- Python 3.x
- pandas library

## Contact

For any questions or issues, please contact **Daksh Gargi** at [f20240888@pilani.bits-pilani.ac.in].
