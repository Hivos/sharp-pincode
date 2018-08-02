## Install or use portable version of AutoIt before using
## Go to "Devices and Printers" and right-click a Sharp printer >> "Printer Preferences" >> "Job Handling", under "Document Filing" select "Hold Only
## Place a checkmark at "PIN Code" and use 10000 as a first pincode, then click "Apply". Make sure to leave the Window open!!!

#include <FileConstants.au3>
#include <MsgBoxConstants.au3>



## Filename to write pin codes to
$sFileName = "c:\tmp\sharp-pins.txt"
## 
$hFilehandle = FileOpen($sFileName, $FO_OVERWRITE)

## Open Au3Info.exe to discover the window Class of the "Job Handling" Window, then bring that window to front

WinActivate("[CLASS:#32770]", "")

Local $hWnd = WinWait("[CLASS:#32770]", "", 10)

For $i = 10000 To 999999 Step 1
    ## Use Au3Info.exe to discover the class of the PIN infut field (i.e. "Edit3")
    ControlSetText($hWnd, "", "Edit3", $i)
    ## As a precaution we wait for 100 ms, not sure if this is needed
    Sleep(100)
    ## Click the "Apply button; Use Au3Info.exe to discover the Instance of the "Apply" button (i.e. "50")
    ControlClick($hWnd, "", "[Class:Button;Instance:50]")
    
	## Get the hashed hexed value of the pincode from the registry and remove the first 2 characters (=0x), so we easily import those values later with "reg add" or powershell
    Local $sSharpPinHex = StringTrimLeft (RegRead("HKEY_CURRENT_USER\Software\SHARP\MX4070-151\printer_ui\job_control", "pin"), 2)
    
    ## Concatenate our output in the form <pincode>;<hexified-pincode> on each line
	Local $sLineToWrite = String($i) & ";" & $sSharpPinHex 
    FileWrite($hFilehandle, @CRLF & $sLineToWrite)
Next


