<#
    .SYNOPSIS
        Generating HTML-code

    .DESCRIPTION
        A longer description.

    .PARAMETER ScriptEnviroment
        Configures a script where to get html tags.
        0 - from a sheet with index '2' in an editable Excel file
        1 - from hard-coded values in script it self
        Default value - 1

    .PARAMETER FilePath
        Path to the edited Excel file
        Default value - "%some_path_to_file"

    .EXAMPLE
        Example of how to run the script:
        powershell -NoProfile -executionpolicy bypass .\gen_html_from_xlsx1.ps1 -ScriptEnviroment 0
        where gen_html_from_xlsx1.ps1 name of the script

    .LINK
        Links to further documentation.

    .NOTES
        Detail on what the script does, if this is needed.

    #>
[CmdletBinding()]
param
(
    [Parameter(Mandatory=$false)][bool]$ScriptEnviroment=$True,
    [Parameter(Mandatory=$false)][string]$FilePath="%some_path_to_file"
)

# Create an Object Excel.Application using Com interface
$objExcel = New-Object -ComObject Excel.Application

# Disable the 'visible' property so the document won't open in excel
$objExcel.Visible = $false
$objExcel.DisplayAlerts = $false

# Open the Excel file and save it in $WorkBook
$WorkBook = $objExcel.Workbooks.Open($FilePath)

$html_tags = @('empty')

if ($ScriptEnviroment -ne $True)
{
    # Specify the index of the WorkSheet in the Excel file. Starts from '1'
    $SheetIndex = 2
    # Load the WorkSheet
    $WorkSheet = $WorkBook.sheets.item($SheetIndex)
    # Count number of rows on selected sheet
    $Sheet_Rows = $WorkSheet.UsedRange.Rows.Count
    for ($i=1; $i -le $Sheet_Rows; $i++)
    {
        $html_tags += $WorkSheet.Range("A$i").Text
        Write-Progress -Activity 'Read HTML from file' -Status "Read-> $FilePath" -PercentComplete ($i / $Sheet_Rows * 100)
    }
}

# Specify the index of the WorkSheet in the Excel file. Starts from '1'
$SheetIndex = 1
# Load the WorkSheet
$WorkSheet = $WorkBook.sheets.item($SheetIndex)
$WorkSheet.activate()

# Count number of rows on selected sheet
$Sheet_Rows = $WorkSheet.UsedRange.Rows.Count
$Count_Rows = $Sheet_Rows - 1

for ($i=2; $i -le $Sheet_Rows; $i++)
{
    $name = $WorkSheet.Range("A$i").Text.Trim()
    $version = $WorkSheet.Range("B$i").Text.Trim()
    $desc = $WorkSheet.Range("D$i").Text.Trim()
    $price = $WorkSheet.Range("F$i").Text.Trim()
    $tags = $WorkSheet.Range("E$i").Text.Trim()
    $test_res = $WorkSheet.Range("G$i").Text.Trim()
    
    $doc_num = $WorkSheet.Range("H$i").Text.Trim()
    $manufactured = $WorkSheet.Range("C$i").Text.Trim()
    $prereq = $WorkSheet.Range("O$i").Text.Trim()
    $web_site = $WorkSheet.Range("K$i").Text.Trim()
    $web_download = $WorkSheet.Range("L$i").Text.Trim()

    $download_special = $WorkSheet.Range("N$i").Text.Trim()
    $proxy_work = $WorkSheet.Range("Q$i").Text.Trim()
    $tester_name = $WorkSheet.Range("J$i").Text.Trim()
    $customer = $WorkSheet.Range("I$i").Text.Trim()
    $comment = $WorkSheet.Range("P$i").Text.Trim()

    $html = ''
    if ($ScriptEnviroment)
    {
        $html += "<h2><span style='color: #000000;'>" + $name + '&nbsp;' + $version
        $html += "</span></h2><p><span style='color: #000000;'>&nbsp;</span></p><p><span style='color: #000000;'>"
        $html += $desc
        $html += "</span></p><p><span style='color: #000000;'>1%:&nbsp;"
        $html += $price
        $html += "</span></p><p><span style='color: #000000;'>2%:&nbsp;"
        $html += $tags  + '...'
    }
    else
    {
        $html += $html_tags[1] + $name + $html_tags[2] + $version
        $html += $html_tags[3]
        $html += $desc
        $html += $html_tags[4]
        $html += $price
        $html += $html_tags[5]
        $html += $tags
    }
        
    $WorkSheet.Cells.Item($i,20) = $html
    $WorkBook.Save()
    Write-Progress -Activity 'Generating HTML' -Status "Generating-> $FilePath" -PercentComplete ($i / $Count_Rows * 100)
}
$WorkBook.Close()
$objExcel.Quit()
Write-Output 'Done.'