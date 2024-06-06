# NOTE: no need to test connection for LastLoggedOnUser
# NOTE: checkin filepath is not the best way to verify if something is installed
# file containing computer names
$computerNamesFile = "C:\temp\banana.txt"

# if the file exists
if (-not (Test-Path $computerNamesFile)) {
    Write-Host "Serial Number of ${computerName}: ${serialNumber}"
    exit
}

# Read computer names from the file
$computerNames = Get-Content $computerNamesFile

# Loopy loop
foreach ($computerName in $computerNames) {
    # Check if the computer is reachable
    if (Test-Connection -ComputerName $computerName -Count 1 -Quiet) {

	
	
	# Gets the OUs for the PCs in the list
	$ou = Add-Content -Path ".\OUs.txt" ((Get-EMComputerDetails $computerName | find "ECN Computers"))
 	# Gets the Service Tags for the PCs in the list
	$serialNumber = Add-Content -Path ".\OUs.txt" ((Get-WmiObject Win32_BIOS -ComputerName $computerName).SerialNumber)

	# If the PC has VISUAL STUDIO (Office 2021, I think)
	$visio = Invoke-Command -ComputerName $computerName -ScriptBlock { Test-Path "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe" }
	# OR has VISUAL STUDIO (Office 365, I think)
	$visio = Invoke-Command -ComputerName $computerName -ScriptBlock { Test-Path "C:\Program Files\Microsoft Office\root\Office16\VISIO.exe" }
        Add-Content -Path ".\VISIO.txt" -Value "${visio}"

	# If the PC has MSProject (for Office 2019 install, I think)
	$project = Invoke-Command -ComputerName $computerName -ScriptBlock { Test-Path "C:\Program Files\Microsoft Office\root\Office16\WINPROJ.exe" } 
	Add-Content -Path ".\MSproject.txt" -Value "${MSproject}"

	

	# Gets the very last last logon users from a list of users
	$line0 = (Get-AMLastLogin $computerName | Out-String -Stream | Select-String -Pattern 'BOILERAD' | ForEach-Object { $_.Line.Trim() -split "\r?\n" })
	$line1 = ($line[0] -split '\\')[1]
	$line2 = ($line2 -split ' ')[0]	
	Add-Content -Path ".\userss.txt" (($line2))

	

	$visio = Invoke-Command -ComputerName $computerName -ScriptBlock { Test-Path "C:\Program Files\Microsoft Office\root\Office16\VISIO.exe" } # VISIO Office 2019 I think
        Add-Content -Path ".\VISIO.txt" -Value "${visio}"

    } else {

	# used to place gaps for PCs that cannot be pinged
 	# uncomment the one's being used I guess

        #Add-Content -Path ".\OUs.txt" -Value "NULL"
        #Add-Content -Path ".\SERIALs.txt" -Value "NULL"
	Add-Content -Path ".\VISIO.txt" -Value "NULL"
	#Add-Content -Path ".\MSproject.txt" -Value "NULL"
	#Add-Content -Path ".\userss.txt" -Value "NULL"
    }
}
