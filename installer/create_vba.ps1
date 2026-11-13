# Approche VBA : Creer un .xlsm avec macro qui ouvre le taskpane
# Cette approche fonctionne sur TOUTES les versions d'Excel

$excel = $null
try {
    # Fermer Excel existant
    Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3

    $outputPath = Join-Path $env:USERPROFILE 'Desktop\AssistantIA_VBA.xlsm'
    if (Test-Path $outputPath) { Remove-Item $outputPath -Force }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Assistant IA"

    # Ajouter du contenu
    $ws.Cells.Item(1,1).Value2 = "ASSISTANT IA EXCEL - AFRIQUE FRANCOPHONE"
    $ws.Cells.Item(1,1).Font.Bold = $true
    $ws.Cells.Item(1,1).Font.Size = 16
    $ws.Cells.Item(1,1).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(16, 185, 129))

    $ws.Cells.Item(3,1).Value2 = "Cliquez sur le bouton ci-dessous pour ouvrir l'Assistant IA :"
    $ws.Cells.Item(3,1).Font.Size = 12

    # Ajouter le code VBA
    $vbaCode = @'
Sub OuvrirAssistantIA()
    ' Ouvrir l'Assistant IA dans le navigateur par defaut
    Dim url As String
    url = "https://afriexcel.netlify.app/"
    
    ' Methode 1: Essayer d'ouvrir comme Web Add-in via Shell
    On Error Resume Next
    
    ' Ouvrir dans le navigateur
    CreateObject("WScript.Shell").Run url, 1, False
    
    MsgBox "L'Assistant IA s'ouvre dans votre navigateur." & vbCrLf & vbCrLf & _
           "Vous pouvez l'utiliser en parallele avec Excel :" & vbCrLf & _
           "- Rapprochement bancaire (Wave, Orange Money, MTN MoMo)" & vbCrLf & _
           "- Bilan SYSCOHADA" & vbCrLf & _
           "- Calcul de paie locale" & vbCrLf & _
           "- Formules et nettoyage de donnees", _
           vbInformation, "Assistant IA Excel - Afrique"
End Sub

Sub Auto_Open()
    ' Message d'accueil a l'ouverture
    MsgBox "Bienvenue dans l'Assistant IA Excel - Afrique Francophone !" & vbCrLf & vbCrLf & _
           "Cliquez sur le bouton 'Ouvrir Assistant IA' dans la feuille," & vbCrLf & _
           "ou utilisez Alt+F8 > OuvrirAssistantIA", _
           vbInformation, "Assistant IA Excel"
End Sub
'@

    # Ajouter le module VBA
    $vbProject = $wb.VBProject
    $vbModule = $vbProject.VBComponents.Add(1) # vbext_ct_StdModule
    $vbModule.Name = "AssistantIA"
    $vbModule.CodeModule.AddFromString($vbaCode)

    # Ajouter un bouton
    $btn = $ws.Buttons.Add(20, 80, 250, 40)
    $btn.Caption = "Ouvrir l'Assistant IA"
    $btn.OnAction = "OuvrirAssistantIA"
    $btn.Font.Size = 14
    $btn.Font.Bold = $true

    # Ajuster les colonnes
    $ws.Columns.Item("A").ColumnWidth = 60

    # Sauvegarder en .xlsm
    $wb.SaveAs($outputPath, 52) # xlOpenXMLWorkbookMacroEnabled
    Write-Host "Fichier cree: $outputPath"

    $wb.Close($false)
    $excel.Quit()
    Write-Host "Excel ferme."
} catch {
    Write-Host "ERREUR: $($_.Exception.Message)"
    if ($excel) {
        try { $excel.Quit() } catch {}
    }
} finally {
    if ($excel) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
