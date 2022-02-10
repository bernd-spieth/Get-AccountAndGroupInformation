<#
    This script collects the files, created by the Get-AccountAndGroupInformation.ps1 script and merges the data of the newest files into one file
#>

#region variables

# Check which PowerShell version we are running on because we collect service information differntly in versions
$powerShellMajorVersion = ($PSVersionTable).PSVersion.Major

$currentDate = Get-Date -Format "yyyy-MM-dd--HH-mm-ss"

$exportRootFolder = "C:\Temp"

$accountInformationRootFolder = Join-Path $exportRootFolder "Accounts"
$accountInformationExportFileName = ("Merged-Accounts-{0}.csv" -f $currentDate)
$accountInformationExportPath = Join-Path $exportRootFolder $accountInformationExportFileName

$serviceInformationRootFolder = Join-Path $exportRootFolder "Services"
$serviceInformationExportFileName = ("Merged-Services-{0}.csv" -f $currentDate)
$serviceInformationExportPath = Join-Path $exportRootFolder $serviceInformationExportFileName
#endregion

#region functions
function New-MergedFile
{
    <#
        Merge information files to a single file
    #>
    param
    (
        [parameter(Mandatory=$true)]$FilePattern,
        [parameter(Mandatory=$true)]$SourcePath,
        [parameter(Mandatory=$true)]$ExportPath
    )

    $files = Get-ChildItem -Path $SourcePath -Filter $FilePattern

    if(Test-Path $ExportPath){Remove-Item -Path $ExportPath}

    # Merge the data to one csv file
    foreach($file in $files)
    {
        try{
            if(!(Test-Path $ExportPath))
            {
                Get-Content -Path $file.FullName | Out-File -Path $ExportPath
            }
            else {
                Get-Content -Path $file.FullName | Select-Object -Skip 1 | Out-File -Path $ExportPath -Append
            }
        }
        catch{
            $e = $_.Exception
            $message = $e.Message
        
            while ($e.InnerException) {
                $e = $e.InnerException
                $message += "`n" + $e.Message
            }
            $message
        }
    }
}

#endregions

# First merge all account files
New-MergedFile -FilePattern "*-Accounts*.csv" -SourcePath $accountInformationRootFolder -ExportPath $accountInformationExportPath

# Then merge all services files
New-MergedFile -FilePattern "*-Services*.csv" -SourcePath $serviceInformationRootFolder -ExportPath $serviceInformationExportPath
