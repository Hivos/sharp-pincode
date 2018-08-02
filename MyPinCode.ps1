Add-Type -AssemblyName PresentationFramework
. C:\Windows\scripts\sharp.ps1

$docs = [Environment]::GetFolderPath('Personal')
$pinFile = Join-Path -Path $docs -ChildPath $('sharp-pins')
$pinCode = (Get-Content $pinFile).Split(';')[0]
[string]$today =  Get-Date -Format yyyyMMdd
[string]$dateLastReset = (Get-Content $pinFile).Split(';')[2]
[int]$pinTries = (Get-Content $pinFile).Split(';')[3]



If (-not (Test-Path $pinFilePath)) {
    Set-Pin
    $pinCode = (Get-Content $pinFile).Split(';')[0]
}



$msgBoxInput =  [System.Windows.MessageBox]::Show("Your pincode is: $pinCode`nDo you want to generate a new PIN-code?",'Printer pincode','YesNo')

switch  ($msgBoxInput) {

    'Yes' {
        If ($today -ne $dateLastReset ) {
            $pinTries = 0
        }
                
        If ($pinTries -gt 2) {
            $pinCode = (Get-Content $pinFile).Split(';')[0]
            [System.Windows.MessageBox]::Show("Maximum amount of resets exceeded for today!`nYour pincode is: $pinCode`n")
            Break
        }
        $pinTries++
        Set-Pin -pinTries $pinTries -Today $today
        $pinCode = (Get-Content $pinFile).Split(';')[0]
        [System.Windows.MessageBox]::Show("Your new pincode is: $pinCode")

    }

      'No' {
    }
}