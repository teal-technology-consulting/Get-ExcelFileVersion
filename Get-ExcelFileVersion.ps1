# Get-ExcelFileVersion.ps1

<#
.SYNOPSIS
   This script retrieves the version of an Excel file.
.PARAMETER Path
   Specifies the path to the Excel file. Be careful to specify only Excel file since the script does not perform any checks on the file type.
.EXAMPLE
   Get-ExcelFileVersion.ps1 -Path .\MyExcelFile.xls
   Retrieves the Excel version of the file MyExcelFile.xls in the current directory.
.EXAMPLE
   dir "$HOME\Documents" -Recurse -Filter *.xls | Get-ExcelFileVersion.ps1
   Retrieves the Excel version of all files with an .XLS extension underneath the "$HOME\Documents" directory.
.NOTES
   See https://docs.microsoft.com/en-us/office/vba/api/excel.xlfileformat for the meaning of the XlFileFormat property.
#>

#Requires -PSEdition Desktop
#Requires -Version 3.0

[CmdletBinding()]
Param
(
    [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
    [Alias('FullName')]
    [string]$Path
)

Begin {
    $Excel = New-Object -ComObject Excel.Application
}

Process {
    Resolve-Path -Path $Path | Select-Object -ExpandProperty Path | ForEach-Object {
        Write-Verbose "Opening workbook '$($_)'."
        $Workbook = $Excel.Workbooks.Open($_, $false, $true)
        if ($Workbook) {
            [int]$XlFileFormatValue = $Workbook.FileFormat
            $XlFileFormatName = ([Microsoft.Office.Interop.Excel.XlFileFormat]).GetEnumName($XlFileFormatValue)

            Get-Item -Path $_ | 
                Select-Object -Property FullName,
                                        @{Name='XlFileFormatName';Expression={$XlFileFormatName}},
                                        @{Name='XlFileFormatValue';Expression={$XlFileFormatValue}}
            $Workbook.Close($false)
            $XlFileFormat = $null
        }
    }
}

End {
    if ($Excel) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    }
}
