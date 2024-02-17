# Registry setting to "Trust access to the VBA project object model" in Word
$registryKey = "HKCU:Software\Microsoft\Office\16.0\Word\Security"
$registryValue = "AccessVBOM"
$registryData1 = "1"
$registryData0 = "0"
# Defines the path each flag file created depending on the original registry state
$flagPath1 = "C:\Users\cortez\AppData\Roaming\Microsoft\Templates\T1137-001_Flag1.txt"
$flagPath2 = "C:\Users\cortez\AppData\Roaming\Microsoft\Templates\T1137-001_Flag2.txt"
# Define the path of copied normal template for restoral
$copyPath = "C:\Users\cortez\AppData\Roaming\Microsoft\Templates\Normal1.dotm"
# Define the path to the normal template
$docPath = "C:\Users\cortez\AppData\Roaming\Microsoft\Templates\Normal.dotm"

# Delete the scheduled task created by the Macro
schtasks /Delete /TN "OpenCalcTask" /F
# Delete the newly created template
Remove-Item -Force $docPath -ErrorAction SilentlyContinue
# Restore the original template
Rename-Item -Force -Path $copyPath -NewName $docPath -ErrorAction SilentlyContinue

#Restore the original state of the registry key
if (Test-Path $flagPath1) 
{
    # The value was originally 0, set back to 0
    New-ItemProperty -Path $registryKey -Name $registryValue -Value $registryData0 -PropertyType DWORD -Force
    Remove-Item -Force $flagPath1 -ErrorAction SilentlyContinue
 } 
 elseif (Test-Path $flagPath2)
 {
    #The value did not previously exist, delete the value
    Remove-ItemProperty -Path $registryKey -Name $registryValue
    Remove-Item -Force $flagPath2 -ErrorAction SilentlyContinue
 }
 else 
 {
     # The value was already 1, do nothing
     Write-Host "The value $registryValue already existed in $registryKey."
 }