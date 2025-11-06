# Excel Material Cost Calculator

The Material Cost Calculator is an Excel-based tool designed to estimate total project or part costs based on material quantities and unit prices.
It automatically aggregates costs by material type and supplier, featuring dynamic tables, charts, and VBA automation for real-time updates.

## Features

- Automatic cost calculation by material and supplier.
- Use of core Excel functions (```XLOOKUP, SUM, PRODUCT```, etc.).
- Refresh with a button via VBA when data changes.
- Interactive dashboards with slicers for Material Type.
- Cost distribution charts for clear visual analysis.
- Structured and modular sheet design (Prices, Detail, Summary, Notes).

## Structure

- Prices --> Material database including code, name, type, unit price, and supplier. You can add or edit items freely.
- Detail --> Project input: material codes and quantities. Only modify the Quantity column and select material type; other fields update automatically.
- Summary --> Dynamic table summarizing total costs and graphical insights. Filters available via slicers (Material Type, Supplier).
- Notes --> Documentation and version information. Internal README for quick reference.

## Usage

1. Open the workbook and go to the Detail sheet.
2. Select a material description (dropdown list).
3. Enter the quantity required.
4. The total cost updates automatically.
5. Check the Summary sheet for charts and breakdowns. Push the Update button to propagate changes to prices or quantity in the Prices sheet throughout the workbook.

## Example Data

<img width="751" height="438" alt="Price table" src="https://github.com/user-attachments/assets/f4907f38-9970-4ca7-8bc0-30323caff364" />


## Future Improvements

- Add cost-per-project summary by category.
- Include a form-based data entry system (VBA UserForm).
- Add automatic export to PDF for reporting.
- Optional password protection for summary formulas.

## Author & Version

Created by: Sandra Martín

Version: 1.0

Date: November 4, 2025

## License

This project is released under the [MIT License](https://choosealicense.com/licenses/mit/). Feel free to modify and reuse for personal or educational purposes.
