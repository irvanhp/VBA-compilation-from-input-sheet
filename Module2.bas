Attribute VB_Name = "Module2"
Option Explicit

Function FindHeaderColumn(ws As Worksheet, headerName As String, maxRow As Long, ByRef headerRow As Long) As Long
    Dim r As Long, c As Long
    Dim lastCol As Long
    
    FindHeaderColumn = 0
    headerRow = 0
    
    For r = 1 To maxRow
        lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        For c = 1 To lastCol
            If Trim(LCase(ws.Cells(r, c).Value)) = LCase(headerName) Then
                FindHeaderColumn = c
                headerRow = r
                Exit Function
            End If
        Next c
    Next r
End Function

Function FindNoBAP(ws As Worksheet) As String
    Dim c As Range
    Dim txtLower As String
    Dim txtOrig As String
    Dim parts() As String
    Dim cnt As Variant
    Dim lastUsedCol As Long
    Dim col As Long
    Dim candidate As String
    
    If ws Is Nothing Then
        FindNoBAP = ""
        Exit Function
    End If
    
    ' Cek cepat apakah sheet benar-benar kosong
    On Error Resume Next
    cnt = Application.CountA(ws.Cells)
    On Error GoTo 0
    If IsError(cnt) Then cnt = 0
    If cnt = 0 Then
        FindNoBAP = ""
        Exit Function
    End If
    
    ' Tentukan last used column
    If ws.UsedRange Is Nothing Then
        FindNoBAP = ""
        Exit Function
    End If
    lastUsedCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
    
    ' Loop setiap sel
    For Each c In ws.UsedRange
        If c.MergeCells Then
            txtOrig = Trim(CStr(c.MergeArea.Cells(1, 1).Text))
        Else
            txtOrig = Trim(CStr(c.Text))
        End If
        
        txtLower = LCase(txtOrig)
        ' toleransi penulisan bap
        If txtLower Like "*no.*bap*" Or txtLower Like "*no bap*" Or txtLower Like "*no.bap*" Then
            ' 1) Jika di sel yang sama ada ":", ambil bagian setelah ":"
            If InStr(txtOrig, ":") > 0 Then
                parts = Split(txtOrig, ":")
                If UBound(parts) >= 1 Then
                    candidate = Trim(parts(1))
                    If candidate <> "" Then
                        FindNoBAP = candidate
                        Exit Function
                    End If
                End If
            End If
            
            ' 2) Jika tidak ada ":", cari ke kanan di baris yang sama sampai ujung
            For col = c.Column + 1 To lastUsedCol
                If ws.Cells(c.Row, col).MergeCells Then
                    candidate = Trim(CStr(ws.Cells(c.Row, col).MergeArea.Cells(1, 1).Text))
                Else
                    candidate = Trim(CStr(ws.Cells(c.Row, col).Text))
                End If
                If Len(candidate) > 0 And candidate <> ":" Then
                    ' jika di sel kanan juga ada ':', ambil bagian setelahnya
                    If InStr(candidate, ":") > 0 Then
                        parts = Split(candidate, ":")
                        If UBound(parts) >= 1 Then
                            candidate = Trim(parts(1))
                        End If
                    End If
                    FindNoBAP = candidate
                    Exit Function
                End If
            Next col
            
            ' jika tidak menemukan apa-apa, done blank
            FindNoBAP = ""
            Exit Function
        End If
    Next c
    
    FindNoBAP = ""
End Function

Sub Compile_List_FromFolder()
    Dim fd As FileDialog
    Dim folderPath As String, fileName As String
    Dim wbSource As Workbook, wbDest As Workbook
    Dim wsSource As Worksheet, wsDest As Worksheet, wsLog As Worksheet
    Dim lastRow As Long, lastCol As Long, destRow As Long
    Dim KodeCol As Long, headerRow As Long
    Dim copyRange As Range
    Dim noBAP As String
    Dim savePath As String
    Dim saveName As String
    
    ' Pilih folder
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Pilih Folder List"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Buat workbook baru untuk hasil compile
    Set wbDest = Workbooks.Add
    On Error Resume Next
    Application.DisplayAlerts = False
    wbDest.Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsDest = wbDest.Sheets.Add
    wsDest.Name = "Compiled_List"
    destRow = 1
    ActiveWindow.DisplayGridlines = False
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    ' Buat sheet log
    Set wsLog = wbDest.Sheets.Add
    wsLog.Name = "Log_File"
    wsLog.Range("A1:D1").Value = Array("File Name", "Status", "Keterangan", "No.BAP")
    
    Application.ScreenUpdating = False
    
    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        Set wbSource = Nothing
        On Error Resume Next
        Set wbSource = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
        On Error GoTo 0
        
        If Not wbSource Is Nothing Then
            On Error Resume Next
            Set wsSource = wbSource.Sheets("List")
            If wsSource Is Nothing Then Set wsSource = wbSource.Sheets("List ")
            On Error GoTo 0
            
            If Not wsSource Is Nothing Then
                ' Cari No.BAP
                noBAP = FindNoBAP(wsSource)
                
                ' Cari kolom kode item
                KodeCol = FindHeaderColumn(wsSource, "kode item", 20, headerRow)
                
                If KodeCol > 0 Then
                    lastRow = wsSource.Cells(wsSource.Rows.Count, KodeCol).End(xlUp).Row
                    lastCol = wsSource.Cells(headerRow, wsSource.Columns.Count).End(xlToLeft).Column
                    If destRow = 1 Then
                        wsDest.Cells(1, 1).Value = "Source_File"
                        wsDest.Cells(1, 2).Value = "No.BAP"
                        wsSource.Range(wsSource.Cells(headerRow, KodeCol), wsSource.Cells(headerRow, lastCol)).Copy _
                            wsDest.Cells(1, 3)
                        destRow = 2
                    End If
                    
                    ' Copy data baris setelah header
                    If lastRow > headerRow Then
                        Set copyRange = wsSource.Range(wsSource.Cells(headerRow + 1, KodeCol), wsSource.Cells(lastRow, lastCol))
                        wsDest.Cells(destRow, 1).Resize(copyRange.Rows.Count).Value = fileName
                        wsDest.Cells(destRow, 2).Resize(copyRange.Rows.Count).Value = noBAP
                        copyRange.Copy wsDest.Cells(destRow, 3)
                        destRow = destRow + copyRange.Rows.Count
                    End If
                    
                    wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(1, 4).Value = Array(fileName, "OK", "Data berhasil diambil", noBAP)
                Else
                    wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(1, 4).Value = Array(fileName, "Gagal", "Kolom 'kode item' tidak ditemukan", noBAP)
                End If
            Else
                wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(1, 4).Value = Array(fileName, "Gagal", "Sheet 'List' tidak ditemukan", "")
            End If
            wbSource.Close SaveChanges:=False
        Else
            wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(1, 4).Value = Array(fileName, "Gagal", "File tidak bisa dibuka", "")
        End If
        
        fileName = Dir()
    Loop
    
    With Sheets("Compiled_List")
    
    'AutoFit semua kolom
    .Cells.EntireColumn.AutoFit
    
    'Format header A1:B1
    With .Range("A1:J1")
        .Interior.Pattern = xlSolid
        .Interior.ThemeColor = xlThemeColorAccent6
        .Interior.TintAndShade = 0.599993896298105
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With
    ActiveWindow.DisplayGridlines = False
    With .Range("A1").CurrentRegion.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .Weight = xlThin
    End With

End With
    

'=====Auto save source file============

saveName = "Compiled_" & Format(Now, "yyyy-mm-dd_hh-nn-ss") & ".xlsx"
savePath = folderPath & saveName
Application.DisplayAlerts = False
wbDest.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
wbDest.Close SaveChanges:=False
Application.DisplayAlerts = True

Application.ScreenUpdating = True

MsgBox "Done!" & vbCrLf & _
       "File hasil sudah disimpan." & vbCrLf & _
       savePath

End Sub


