#################################################CHECKLIST COM#################################################
            #Specify the path of the excel file
            $FilePath = "F:\■■■■■■■■\■■■■■■■■\baseFileToReadFrom.xlsx"
            #Specify the Sheet name
            $SheetName = "■■■■■■■■" 
            <#
            Function checks for base file and appends a number to it
            example: 11-xxxx_CheckList1
            #>
            Function Get-NewFileName {
                param(
                    # Path to file
                    [Parameter(Mandatory=$true)]
                    [string]
                    $FilePath, 
                    # Base file name (without number or extension)
                    [Parameter(Mandatory=$true)]
                    [string]
                    $FileName,
                    # File extension (without the .)
                    [Parameter(Mandatory=$true)]
                    [string]
                    $FileExt
                )
                if (Test-Path "$FilePath\$FileName.$FileExt") {
                    # File exists, append number
                    $fileSeed = 0
                    do {
                        $fileSeed++
                        $newFullPath = "$FilePath\$FileName$fileSeed.$FileExt"
                    } until ( ! (Test-Path $newFullPath) )
                } else {
                    return "$FilePath\$FileName.$FileExt"
                }
                return "$newFullPath"
            }

            $final = Get-NewFileName -FilePath "F:\checklist\DB" -FileName "gold" -FileExt "xlsx"
            Copy-Item -Path "F:\■■■■■■■■\.xlsx" -Destination '$final'

            #Write-Host $final
            # Create an Object Excel.Application using Com interface
            $objExcel = New-Object -ComObject Excel.Application
            # Disable the 'visible' property so the document won't open in excel
            $objExcel.Visible = $false
            # Open the Excel file and set Read-Only to FALSE
            $WorkBook = $objExcel.Workbooks.Open($FilePath, $null, $false)
            <#
            Take the active worksheet, reference the cell by text, obtain the value
            replace the value with the variable given and save it
            This should be unique file generated during runtime / temp file
            #>
            $WorkSheet = $WorkBook.sheets.item($SheetName)

            #ENTER NEW INPUT HERE FOR EXCEL#
            #SYSINFO SECTION#
            $WorkSheet.Range("E4").Value2 = $orderBox.Text
            $WorkSheet.Range("L4").Value2 = $env:COMPUTERNAME
            $WorkSheet.Range("E5").Value2 = $caseNBox.Text
            $WorkSheet.Range("L5").Value2 = $makeBox.Text, $modelBox.Text
            $WorkSheet.Range("L6").Value2 = $serialBox.Text
            #TASKS TO PERFORM SECTION#
            <#
            $WorkSheet.Range("B17", "C17").Value2 = 
            $WorkSheet.Range("B18", "C18").Value2 =
            $WorkSheet.Range("B19", "C19").Value2 =
            $WorkSheet.Range("B20", "C20").Value2 =
            $WorkSheet.Range("B21", "C21").Value2 =
            $WorkSheet.Range("B22", "C22").Value2 =
            #>
            $Guid = [Guid]::newGuid()
            #Saves the changes made by the program to the execel file
            $WorkBook.SaveAs($final)

            #Closes the Excel sheet
            $objExcel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
            spps -n Excel
