# Excel Automation with OpenPyXL

This Python program automates the process of reading, modifying, and visualizing data in an Excel workbook using the `openpyxl` library. The script specifically processes a worksheet named "Sheet1" within the given workbook, applies a transformation to numeric data in a specific column, and then generates a bar chart to visualize the results.

## Features

- **Reads Data**: Loads data from a specified Excel workbook.
- **Processes Data**: Applies a 10% discount to numeric values in the third column and writes the corrected values to the fourth column.
- **Handles Non-Numeric Data**: Identifies and reports non-numeric data in the third column.
- **Visualizes Data**: Generates a bar chart based on the corrected values and adds it to the worksheet.
- **Saves Workbook**: Saves the modified workbook under the same filename.

## Prerequisites

- Python 3.x
- `openpyxl` library

You can install the `openpyxl` library using pip:

```bash
pip install openpyxl
```

## Usage

1. Clone the repository to your local machine.
2. Navigate to the directory containing the script.
3. Run the script and provide the name of your Excel file when prompted.

```bash
python process_workbook.py
```

## Script Details

### Function: `process_workbook(filename)`

- **Parameters**:
  - `filename` (str): The name of the Excel file to process.
  
- **Workflow**:
  1. **Load Workbook**: Opens the Excel workbook specified by `filename`.
  2. **Select Worksheet**: Accesses the worksheet named "Sheet1".
  3. **Iterate Through Rows**: For each row, starting from the second row to the last row:
      - Reads the value in the third column.
      - If the value is numeric (either integer or float), it applies a 10% discount and writes the corrected value to the fourth column.
      - If the value is non-numeric, it prints a message indicating the non-numeric data.
  4. **Generate Bar Chart**: Creates a bar chart based on the corrected values in the fourth column and adds it to the worksheet.
  5. **Save Workbook**: Saves the changes to the workbook.

- **Sample Output**:
  - Messages indicating the original and corrected values for each row.
  - A message confirming the workbook has been saved successfully.

```python
import openpyxl as xl
from openpyxl.chart import BarChart, Reference

def process_workbook(filename):
    wb = xl.load_workbook(filename)
    sheet = wb['Sheet1']

    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row, 3)
        if isinstance(cell.value, (int, float)):
            corrected_price = cell.value * 0.9
            corrected_price_cell = sheet.cell(row, 4)
            corrected_price_cell.value = corrected_price
            print(f"Row {row}: {cell.value} -> {corrected_price}")
        else:
            print(f"Row {row}: Non-numeric data in cell {cell.coordinate}")
        
    values = Reference(sheet,
                       min_row = 2,
                       max_row = sheet.max_row,
                       min_col = 4,
                       max_col = 4)
    chart = BarChart()
    chart.add_data(values)
    sheet.add_chart(chart, 'e2')
    wb.save(filename)
    print(f"Workbook {filename} saved successfully.")

filename = input("Enter the file name:")
process_workbook(filename)
```

## Example

```text
Enter the file name: sample.xlsx
Row 2: 100 -> 90.0
Row 3: 200 -> 180.0
Row 4: Non-numeric data in cell C4
...
Workbook sample.xlsx saved successfully.
```

## Contributing

Contributions are welcome! Please feel free to submit a pull request or open an issue if you have any suggestions or improvements.

---
