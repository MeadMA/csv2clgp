# csv2clgp: Fills Class List Generator Pro PDFs from Skyward CSV exports
# Copyright (C) 2020  Michael A. Mead
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

## Program starts below here.

# Finds a specific form field in the list of all fields.
Function Get-Field {
	param(
		[Object]$Fields,
		[String]$StudentAttribute,
		[Int]$Count
	)

	# Convert single-digit count to double-digit
	If ($Count -lt 10) {
		$StringCount = "0$($Count)"
	} Else {
		$StringCount = "$($Count)"
	}

	# Try with one space
	$Field = ($Fields | Where-Object {$_.key -eq "$($StudentAttribute) $($StringCount)"}).Value
	# If not found, try with two spaces
	If (!$Field) {
		$Field = ($Fields | Where-Object {$_.key -eq "$($StudentAttribute)  $($StringCount)"}).Value
	}

	Return $Field
}

# Finds a specific form field in the list of all fields, and sets it to a given
# value.
Function Set-Field {
	param(
		[Object]$Fields,
		[String]$StudentAttribute,
		[Int]$Count,
		[String]$NewValue
	)

	$Field = Get-Field -Fields $Fields -StudentAttribute $StudentAttribute -Count $Count
	$Discard = $Field.SetValue($NewValue)
}

# Fills Class List Generator Pro PDF with values in a given Skyward CSV
Function Import-Csv-to-Clgp {
	param(
		[String]$CsvPath,
		[String]$ClgpPath
	)

	# Generate destination file name
	$csvName = (Get-ChildItem $CsvPath).BaseName
	$destPdf = "$PSScriptRoot\output\$($csvName).pdf"

	# Read CSV file, and strip off non-CSV header lines Skyward likes to
	# include
	$rawcsv = Get-Content $CsvPath
	$rawcsv = $rawcsv[6..$rawcsv.length]

	# Process remainder of file as a CSV
	$tempfile = "$($env:TEMP)\csv2clgp_$(Get-Random).csv"
	Set-Content -Path $tempfile -Value $rawcsv
	$csv = Import-Csv $tempfile
	Remove-Item $tempfile

	# Open CLGP template with iText, and ignore owner password
	$reader = [iText.Kernel.Pdf.PdfReader]::new($ClgpPath)
	$Discard = $reader.SetUnethicalReading($True)

	# Open PDF writer for destination PDF
	$writer = [iText.Kernel.Pdf.PdfWriter]::new($destPdf)

	# Initiate PDF work with iText
	$doc = [iText.Kernel.Pdf.PdfDocument]::new($reader, $writer)

	# Get form fields from PDF
	$form = [iText.Forms.PdfAcroForm]::getAcroForm($doc, $True)
	$fields = $form.getFormFields()

	# Update form fields for each student's attributes from CSV
	$StudentCount = 1
	ForEach ($Student in $csv) {
		$NameSplit = $Student."Last, First MI" -Split ", "
		$LastName = $NameSplit[0]
		$FirstName = $NameSplit[1].Substring(0, $NameSplit[1].Length-3)
		$FullName = "$($FirstName) $($LastName)"

		Set-Field -Fields $Fields -StudentAttribute "STUDENT" -Count $StudentCount -NewValue $FullName
		Set-Field -Fields $Fields -StudentAttribute "FIRST" -Count $StudentCount -NewValue $FirstName
		Set-Field -Fields $Fields -StudentAttribute "LAST" -Count $StudentCount -NewValue $LastName
		Set-Field -Fields $Fields -StudentAttribute "PHONE" -Count $StudentCount -NewValue $Student."Phone Number"
		Set-Field -Fields $Fields -StudentAttribute "E-MAIL" -Count $StudentCount -NewValue $Student."Email"

		$StudentCount = $StudentCount + 1
		If ($StudentCount -gt 40) {
			Break
		}
	}

	# Close document, saving changes
	$doc.Close()
}

# Ensure input and output directories exist
If (!(Test-Path "$PSScriptRoot\input")) {
	New-Item -Path "$PSScriptRoot\input" -ItemType Directory
}
If (!(Test-Path "$PSScriptRoot\output")) {
	New-Item -Path "$PSScriptRoot\output" -ItemType Directory
}

# Load iText 7 and its dependencies
Add-Type -Path "$PSScriptRoot\lib\Common.Logging.Core.dll"
Add-Type -Path "$PSScriptRoot\lib\Common.Logging.dll"
Add-Type -Path "$PSScriptRoot\lib\itext.io.dll"
Add-Type -Path "$PSScriptRoot\lib\itext.kernel.dll"
Add-Type -Path "$PSScriptRoot\lib\itext.forms.dll"
Add-Type -Path "$PSScriptRoot\lib\itext.layout.dll"
Add-Type -Path "$PSScriptRoot\lib\BouncyCastle.Crypto.dll"

# Get source file (empty Class List Generator Pro PDF) and list of CSVs
$ClgpTemplate = "$PSScriptRoot\clgp_empty.pdf"
If (!(Test-Path $ClgpTemplate)) {
	Write-Error "clgp_empty.pdf is missing!"
	Exit 1
}
$csvList = Get-ChildItem "$PSScriptRoot\input\*.csv"
If ($csvList.Count -eq 0) {
	Write-Error "No CSV files found!"
	Exit 2
}

# Process CSV files
ForEach ($csv in $csvList) {
	Write-Output "Processing $($csv.Name)..."
	Import-Csv-to-Clgp -CsvPath $csv.FullName -ClgpPath $ClgpTemplate
}

# Notify we're done
Write-Output "All files processed."
Exit 0
