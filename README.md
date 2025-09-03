# Excel-Exporter
An exporter for generating github compatible files from excel files.

# Current Features

This exports the following:
- Sheets displayed values
- Sheets formula values
- Sheet vba files
- ThisWorkbook vba file
- Module vba files
- Class vba files
- Custom Ribbon XML

> [!NOTE]
> 
> This program **does not** fire off vba code in `ThisWorkbook` > `Workbook_Open()` (and similar) subroutines

> [!WARNING]
> 
> I usually use [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) for creating my custom ribbons
> 
> As such, I haven't tested on any other custom ribbons

# Notes

I have been looking for something that I could run on my vba/excel projects to keep track of changes in both the worksheets and the VBA backend. As I work with the VBA code, the displayed values _and_ the formulas, I need something to track all of it. Here it is!

Feel free to fork and push your suggestions!

Happy to take any suggestions, and interested in any issues you have.

## Roadmap
<img width="1117" height="890" alt="image" src="https://github.com/user-attachments/assets/69c0f6ba-5006-46bc-ab13-257e78cdc1c2" />
