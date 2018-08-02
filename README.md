# Sharp pincode

Docs on wiki:

https://helpdesk.hivos.nl/start/manuals/secure_printing

## Scripts
sharp-pincode-bruteforcer.au   # AutoIt script that creates a rainbow table of sorts for 5-digit pincodes
MyPincode.ps1                  # App that presents user with pincode and optionally change it (max. 3 changes per day allowed)
sharp.ps1                      # script that runs via scheduled task, generates the pincode, sets pincode and other settings in registry

Scripts are stored for deployment in E:\Applications\FilesForWorkstations\Scripts and ultimately end up in c:\windows\scripts
