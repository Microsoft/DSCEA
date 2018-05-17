function Convert-DSCEAresultsToExcel {
<#   
.SYNOPSIS   
Creates Excel sheet file from DSCEA results

.DESCRIPTION 
Used to convert the raw XML output from a DSCEA scan to a table in Excel with coloring or TRUE and FALSE

.PARAMETER InputXML
The file name (full file path) of the XML file you would like to convert to XLSX.  If one is not provided, Convert-DSCEAresultsToExcel looks to the current directory for the most recently created XML file.

.PARAMETER OutFile
The new file name (full file path) of the output XLSX file.  The file path must include a pre-existing folder.  If one is not provided, Convert-DSCEAresultsToExcel will create output.xlsx in the current directory.

.LINK
https://microsoft.github.io/DSCEA

.EXAMPLE
Convert-DSCEAresultsToExcel

Description
-----------
Looks in the current directory for the most recently created XML file, parses the file and converts it to a XLSX file in the same directory

.EXAMPLE
Convert-DSCEAresultsToExcel -InputXML C:\Users\username\Documents\DSCEA\results.20170311-2307-31.xml

Description
-----------
Converts C:\Users\username\Documents\DSCEA\results.20170311-2307-31.xml to a XLSX file, saved in the current directory

.EXAMPLE
Convert-DSCEAresultsToExcel -InputXML C:\Users\username\Documents\DSCEA\results.20170311-2307-31.xml -OutFile C:\Users\username\Documents\DSCEA\output.XLSX

Description
-----------
Converts C:\Users\username\Documents\DSCEA\results.20170311-2307-31.xml to a XLSX file saved at C:\Users\username\Documents\DSCEA\output.XLSX
#>
    [CmdLetBinding()]
    param
    (
    [String]$InputXML = '.',
    [String]$OutFile = '.\output.xlsx',
    [String]$ExceptionsFile = $null
    )
    begin
    {
    if(!$(get-module ImportExcel))
    { 
        if(!$(get-module -ListAvailable | Where-Object name -eq "ImportExcel"))
        {
            Write-Verbose "You need to install the PowerShell module ImportExcel found in the PSGallery"  
            break
        }
        else
        {
            Import-Module ImportExcel
        }

    }
    }
    Process
    {
    (Get-ChildItem -Path $InputXML -Filter '*.xml' | Sort-Object LastWriteTime -Descending)[0] | ForEach-Object {
        $results = Import-CliXml $_
        $newdata = $results | ForEach-Object {
            $_.Compliance | ForEach-Object {
                if($_.ResourcesInDesiredState) {
                    $_.ResourcesInDesiredState | ForEach-Object {
                        $_
                    }
                }
                if($_.ResourcesNotInDesiredState) {
                    $_.ResourcesNotInDesiredState | ForEach-Object {
                        $_
                    }
                }
            }
        }
    }
    $parsedData = $newdata | Select-Object PSComputerName, ResourceName, InstanceName, InDesiredState, ConfigurationName, StartDate
    if(Test-Path $OutFile ){Remove-Item -Path $OutFile -Force}
    $parsedData | Export-Excel -path $OutFile -WorkSheetname "Results" -BoldTopRow -TableStyle Medium2 -TableName "securitypolicyresults" -NoLegend -AutoSize -FreezeTopRow -ConditionalText $( 
    New-ConditionalText FALSE -Range "D:D" -BackgroundColor Red -ConditionalTextColor Black
    New-ConditionalText TRUE -Range "D:D" -BackgroundColor Green -ConditionalTextColor Black
    )
            


    }
    END
    {
        return Get-ItemProperty $OutFile
    }
}