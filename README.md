# csv2clgp
Fills Class List Generator Pro PDFs with Skyward CSV class rosters

[Class List Generator Pro](https://www.teacherspayteachers.com/Product/Class-List-Generator-PRO-1997070) is a PDF with a series of templates for generating class rosters.  Users fill in a master roster, and the rest of the pages are automatically filled from the master.

[Skyward Student Information System](https://www.skyward.com/k-12/student-information-system) is a web-based application popular among schools for managing student data.

Skyward SIS provides an option to export class rosters as CSV files.  The csv2clgp PowerShell script takes those CSV files as input and outputs the roster information into the Class List Generator Pro PDF.  Each CSV roster is output to its own CLGP PDF.

# Use
1. Place your Class List Generator Pro PDF in the same directory as csv2clgp, and name it `clgp_empty.pdf`.
1. In the `input` directory, place your Skyward class roster CSV files.
1. Run `csv2clgp.ps1`.  Optionally, use `csv2clgp.bat`.  It will simply run the PowerShell script and wait for you to press a key before closing the window.
1. View your generated PDF files in the `output` directory.

# Technical Details
csv2clgp is a PowerShell script that uses the [iText 7 .NET library](https://github.com/itext/itext7-dotnet) to fill the CLGP PDF.  iText 7 is made available under the AGPL.

You must purchase a license to use the CLGP PDF, so it is not included with this script.
