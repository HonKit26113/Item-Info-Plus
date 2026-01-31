Sub ExportAllLanguages_Debug()
    Dim ws As Worksheet
    Dim i As Long
    Dim LanguageIDCell As String
    Dim TotalLanguages As Long
    
    ' ==========================================
    '               SETUP AREA
    ' ==========================================
    Set ws = ActiveSheet
    
    ' 1. CRITICAL: The cell that changes the number
    LanguageIDCell = "B1"
    
    ' 2. Total languages to export
    TotalLanguages = 29
    
    ' ==========================================
    '            END SETUP AREA
    ' ==========================================

    Dim fName As String
    Dim folderPath As String
    Dim filePath As String
    Dim cell As Range
    Dim lastRow As Long
    Dim utf8Stream As Object

    ' Get folder path
    folderPath = Application.ActiveWorkbook.Path
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' START MAIN LOOP
    For i = 1 To TotalLanguages
        
        ' 1. Change ID and Force Update
        ws.Range(LanguageIDCell).Value = i
        Application.CalculateFull
        DoEvents
        
        ' 2. Clean Filename
        fName = ws.Range("C1").Value
        If Trim(fName) = "" Then fName = "Language_" & i
        
        If LCase(Right(fName, 4)) = ".txt" Then fName = Left(fName, Len(fName) - 4)
        If LCase(Right(fName, 5)) = ".lang" Then fName = Left(fName, Len(fName) - 5)
        
        fName = fName & ".lang"
        filePath = folderPath & "\exported\" & fName
        
        ' 3. EXPORT DATA (With Error Trap)
        lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        
        If lastRow > 0 Then
            Set utf8Stream = CreateObject("ADODB.Stream")
            utf8Stream.Type = 2
            utf8Stream.Charset = "utf-8"
            utf8Stream.Open
            
            ' ENABLE ERROR TRAPPING
            On Error GoTo RowCrashHandler
            
            For Each cell In ws.Range("D1:D" & lastRow)
                ' If this line fails, code jumps to RowCrashHandler below
                utf8Stream.WriteText cell.Value & vbCrLf
            Next cell
            
            ' DISABLE ERROR TRAPPING (so we don't trap non-loop errors)
            On Error GoTo 0
            
            utf8Stream.SaveToFile filePath, 2
            utf8Stream.Close
            Set utf8Stream = Nothing
        End If
        
    Next i

    MsgBox "Done! All files exported successfully."
    Exit Sub

' ==========================================
'           ERROR HANDLER
' ==========================================
RowCrashHandler:
    ' Close the file so it isn't locked
    If Not utf8Stream Is Nothing Then
        utf8Stream.Close
        Set utf8Stream = Nothing
    End If
    
    ' Show the popup with the specific Row Number
    MsgBox "CRITICAL ERROR FOUND!" & vbCrLf & vbCrLf & _
           "Language Number: " & i & vbCrLf & _
           "Excel Row: " & cell.Row & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical
           
    ' Stop the script so you can fix it
    End
    
End Sub

