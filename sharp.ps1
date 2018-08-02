$MinPin = 10000
$MaxPin = 99999
$sqlLite3 = 'C:\Windows\tools\sqlite3.exe'
$sqlLiteDB = 'C:\Windows\Tools\sharp.db'
$pinCodeBytes=""
$Printers = 'MX4141-160', 'MX4070-151', 'MX4141-152', 'MX4070-153'

$zeroPadding = ",0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00," +
"0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00," +
"0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00," + 
"0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00," + 
"0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00," + 
"0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00," +
"0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00," + 
"0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00," +
"0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00"

$docs = [Environment]::GetFolderPath('Personal')
$pinFilePath = Join-Path -Path $docs -ChildPath $('sharp-pins')
#$pinFilePathOthers = Join-Path -Path $docs -ChildPath 'sharp-pins-*'
#Remove-Item $pinFilePathOthers -Exclude *$('sharp-pins-' +$thisWeek), *$('sharp-pins-' + $previousWeek) -Force -Confirm:$False

$disallowedPins = @("11111","22222","33333","44444","55555","66666","77777","88888","99999","10000","20000",
"30000","40000","50000","60000","70000","80000","90000","12345","23456","34567","45678","56789","54321","12121")

$pinCodeHashed = (Get-Content $pinFilePath).Split(';')[1]
$pinCodeBytes = foreach ($b in (0..15)){"0x" + $pinCodeHashed.Substring(2*$b,2)}
$pinCodeBytes = $pinCodeBytes -Join ","
$pinCodeBytes = $pinCodeBytes + $zeroPadding
#$pinCodeBytes|Out-File $pinFilePath -Append
## https://stackoverflow.com/questions/6551224/how-to-set-a-binary-registry-value-reg-binary-with-powershell
## https://stackoverflow.com/a/18092826
$array = $pinCodeBytes -split ","

Function Set-Pin {
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$False,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$pinTries = 0,
        [Parameter(Mandatory=$False,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [string]$today = (Get-Date -Format yyyyMMdd)
    )
    [string]$pin = Get-Random -Minimum $MinPin -Maximum $MaxPin
    While ($disallowedPins.Contains($pin)) { 
        [string]$pin = Get-Random -Minimum $MinPin -Maximum $MaxPin
    }
    [string]$hash = ( &$sqlLite3 $sqlLiteDB  "SELECT hash FROM pins WHERE pin = $pin;")
    [string]$key = $pin + ';' + $hash + ';' + $today + ';' +$pinTries

    If (Test-Path $pinFilePath) {
        $fileObject = Get-Item($pinFilePath) -Force
        $fileObject.Attributes = ""
    }

    Write-Output $key |Out-File $pinFilePath -Force -Confirm:$False
    $fileObject = Get-Item($pinFilePath) -Force
    $fileObject.Attributes = ""
    $fileObject.Attributes = "Hidden","System"
}

Function Set-RegistryJobControl {
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $PrinterName
    )
    
    New-Item -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Force -ErrorAction SilentlyContinue
    
    
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'account_number' -Value ([byte[]](0x00))
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'folder_index' -Value 0 -Type Dword
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'folder_pass' -Value ([byte[]](0x00))
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'job_name' -Value '' -Type String
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'login_name' -Value '' -Type String
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'login_pass' -Value ([byte[]](0x00))
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'pin' -Value ([byte[]]($array))
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'set_login_name' -Value 0 -Type Dword
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'set_login_pass' -Value 0 -Type Dword
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'use_account_number' -Value 0 -Type Dword
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'use_job_name' -Value 0 -Type Dword
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'use_PIN' -Value 1
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'use_user_name' -Value 0 -Type Dword
    Set-ItemProperty -Path "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control" -Name 'user_name' -Value '' -Type String
}


Function Remove-RegistryPermissions {
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $PrinterName
    )
 
    $username = "Hivos\" + $env:USERNAME
    $path = "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control"
    $Acl = Get-ACL $path
    $acl.SetAccessRuleProtection($true, $true)
    Set-Acl $path $Acl

    $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("SOFTWARE\Sharp\$PrinterName\printer_ui\job_control",[Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree,[System.Security.AccessControl.RegistryRights]::ChangePermissions)
    $acl = $key.GetAccessControl()
    $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("hivos\app_print_without_pin","FullControl","Allow")
    $acl.SetAccessRule($rule)
    $key.SetAccessControl($acl)

    $acl = $key.GetAccessControl()
    $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("$username","ReadPermissions, QueryValues,Notify,EnumerateSubKeys","Allow")
    $acl.SetAccessRule($rule)
    $key.SetAccessControl($acl)
}

Function Add-RegistryPermissions {
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $PrinterName
    )
 
    $username = "Hivos\" + $env:USERNAME

    $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("SOFTWARE\Sharp\$PrinterName\printer_ui\job_control",[Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree,[System.Security.AccessControl.RegistryRights]::ChangePermissions)

    $acl = $key.GetAccessControl()
    $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("$username","FullControl","Allow")
    $acl.SetAccessRule($rule)
    $key.SetAccessControl($acl)
}

If (-not (Test-Path $pinFilePath)) {
    Set-Pin
}

foreach ($printer in $Printers) { Add-RegistryPermissions -Printername $printer }
foreach ($printer in $Printers) { Set-RegistryJobControl -Printername $printer }
foreach ($printer in $Printers) { Remove-RegistryPermissions -Printername $printer }

$pinCode = (Get-Content $pinFilePath).Split(';')[0]
New-Item -Path 'HKCU:\Software\Hivos\Sharp' -Force -ErrorAction SilentlyContinue
Set-ItemProperty -Path 'HKCU:\Software\Hivos\Sharp' -Name 'PIN' -Value $pinCode