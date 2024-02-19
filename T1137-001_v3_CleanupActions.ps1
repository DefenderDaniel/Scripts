# Registry setting to "Trust access to the VBA project object model" in Word
$registryKey = "HKCU:Software\Microsoft\Office\16.0\Word\Security"
$registryValue = "AccessVBOM"
$registryData1 = "1"
$registryData0 = "0"
# Defines the path each flag file created depending on the original registry state
$flagPath1 = "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\T1137-001_Flag1.txt"
$flagPath2 = "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\T1137-001_Flag2.txt"
# Define the path of copied normal template for restoral
$copyPath = "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\Normal1.dotm"
# Define the path to the normal template
$docPath = "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\Normal.dotm"

# Delete the scheduled task created by the Macro
schtasks /Delete /TN "OpenCalcTask" /F

#Restore the orginal template if the backup copy exists
if (Test-Path $copyPath)
{
    #Delete the injected template
    Remove-Item -Force $docPath -ErrorAction SilentlyContinue
    # Restore the original template
    Rename-Item -Force -Path $copyPath -NewName $docPath -ErrorAction SilentlyContinue
    Write-Host "The original template has been restored"
}
else
{
    Write-Host "The original template is present"
}
     
#Restore the original state of the registry key
if (Test-Path $flagPath1) 
{
    # The value was originally 0, set back to 0
    New-ItemProperty -Path $registryKey -Name $registryValue -Value $registryData0 -PropertyType DWORD -Force
    Remove-Item -Force $flagPath1 -ErrorAction SilentlyContinue
    Write-Host "The original registry state has been restored"
 } 
 elseif (Test-Path $flagPath2)
 {
    #The value did not previously exist, delete the value
    Remove-ItemProperty -Path $registryKey -Name $registryValue
    Remove-Item -Force $flagPath2 -ErrorAction SilentlyContinue
    Write-Host "The original registry state has been restored"
 }
 else 
 {
     # The value was already 1, do nothing
     Write-Host "The value $registryValue already existed in $registryKey."
 }