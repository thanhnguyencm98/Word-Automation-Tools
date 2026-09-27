<#
.SYNOPSIS
    Rebuilds the VBA modules in WordAutomationTools.docm from Macros\*.bas.

.DESCRIPTION
    Removes every standard module from the .docm, imports each .bas file in
    Macros\, and saves the document, so the committed .docm matches the source.

    Requires desktop Word for Windows and the Trust Center option
    "Trust access to the VBA project object model"
    (File > Options > Trust Center > Trust Center Settings > Macro Settings).
    Turn that option off again when finished.

    After running, open the .docm, press Alt+F11, and select
    Debug > Compile VBAProject before committing.
#>
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$docmPath = Join-Path $root 'WordAutomationTools.docm'
$macroDir = Join-Path $root 'Macros'
$vbextStdModule = 1

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($docmPath)
    try {
        $components = $doc.VBProject.VBComponents
    } catch {
        throw "Word refused access to the VBA project. Enable 'Trust access to the VBA project object model' and run again."
    }

    # Remove all standard modules first so renamed or deleted .bas files do not linger.
    foreach ($component in @($components | Where-Object { $_.Type -eq $vbextStdModule })) {
        Write-Host "Removing $($component.Name)"
        $components.Remove($component)
    }
    foreach ($bas in Get-ChildItem -Path $macroDir -Filter '*.bas') {
        Write-Host "Importing $($bas.Name)"
        [void]$components.Import($bas.FullName)
    }

    $doc.Save()
    $doc.Close()
    Write-Host "Updated $docmPath"
} finally {
    $word.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
