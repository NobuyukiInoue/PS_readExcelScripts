param($outputFile)

if (-Not $outputFile) {
    Write-Host "Usage : $MyInvocation.MyCommand.Name <output-file>"
    exit
}

.\read_CenterSW_Cells.ps1 | Out-File $outputFile -Encoding default
