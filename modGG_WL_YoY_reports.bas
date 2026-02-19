Attribute VB_Name = "modGG_WL_YoY_reports"
Option Explicit

'====================================================
' Main Macro: Build YOY ZIPs, multiple temp files per ZIP (Column H only) with logging
'====================================================
Sub Build_GG_WL_YOY_Zips()

    Dim wbCA As Workbook
    Dim wsCA As Worksheet
    Dim wsBrand As Worksheet
    Dim wbBrandReport As Workbook
    Dim wsWorking As Worksheet
    
    Dim logWS As Worksheet
    Dim brandData As Variant
    Dim rngBrand As Range
    Dim rngHeader As Range
    Dim rngBrandReport As Range
    
    Dim lastCAColumn As Long
    Dim totalsRow As Long
    Dim i As Long, j As Long
    Dim startRow As Long, endRow As Long
    
    Dim tempFilePath As String, zipPath As String
    Dim logRow As Long
    Dim currentZIP As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wbCA = ActiveWorkbook
    Set wsCA = wbCA.Worksheets("Expo Broker Brands")
    Set wsBrand = wbCA.Worksheets("Brands2Send")
    
    '--------------------------------------------
    ' Sort Brands2Send by Column H (ZIP path)
    '--------------------------------------------
    wsBrand.Sort.SortFields.Clear
    wsBrand.Sort.SortFields.Add Key:=wsBrand.columns(8), SortOn:=xlSortOnValues
    With wsBrand.Sort
        .SetRange wsBrand.UsedRange
        .Header = xlYes
        .Apply
    End With
    
    '--------------------------------------------
    ' Create or clear log sheet
    '--------------------------------------------
    On Error Resume Next
    Set logWS = wbCA.Worksheets("YOY_Log")
    On Error GoTo 0
    If logWS Is Nothing Then
        Set logWS = wbCA.Worksheets.Add
        logWS.Name = "YOY_Log"
        logWS.Range("A1:E1").Value = Array("Brand", "Temp File", "ZIP Path", "Status", "Timestamp")
    Else
        logWS.Rows("2:" & logWS.Rows.Count).Clear
    End If
    logRow = 2
    
    '--------------------------------------------
    ' CA sheet setup
    '--------------------------------------------
    lastCAColumn = wsCA.Range("A2").End(xlToRight).Column
    Set rngHeader = wsCA.Range(wsCA.Cells(1, 1), wsCA.Cells(1, lastCAColumn))
    
    '--------------------------------------------
    ' Load Brands2Send table
    '--------------------------------------------
    Set rngBrand = wsBrand.Range("A1").CurrentRegion
    brandData = rngBrand.Value
    
    '--------------------------------------------
    ' Initialize progress bar
    '--------------------------------------------
    ProgressBar.Show vbModeless
    UpdateProgressBar 0
    
    '--------------------------------------------
    ' Loop through brands
    '--------------------------------------------
    i = 2
    Do While i <= UBound(brandData, 1)
        zipPath = Trim(CStr(brandData(i, 8)))
        If zipPath = "" Then
            ' Log skipped brand
            logWS.Cells(logRow, 1).Value = CStr(brandData(i, 1))
            logWS.Cells(logRow, 2).Value = ""
            logWS.Cells(logRow, 3).Value = ""
            logWS.Cells(logRow, 4).Value = "Skipped - Missing ZIP path"
            logWS.Cells(logRow, 5).Value = Now
            logRow = logRow + 1
            i = i + 1
            GoTo NextIteration
        End If
        
        currentZIP = zipPath
        
        ' Loop through consecutive rows with same ZIP
        j = i
        Do While j <= UBound(brandData, 1) And Trim(CStr(brandData(j, 8))) = currentZIP
            
            ' Update progress bar
            UpdateProgressBar (j - 1) / (UBound(brandData, 1) - 1)
            
            ' Determine start/end rows for brand
            startRow = CLng(brandData(j, 6))
            endRow = CLng(brandData(j, 7))
            
            ' Build brand range
            Set rngBrandReport = wsCA.Range(wsCA.Cells(startRow, 1), wsCA.Cells(endRow, lastCAColumn))
            
            ' Create temp workbook
            Set wbBrandReport = Workbooks.Add
            Set wsWorking = wbBrandReport.Worksheets(1)
            
            ' Copy header + data
            rngHeader.Copy wsWorking.Cells(1, 1)
            rngBrandReport.Copy wsWorking.Cells(2, 1)
            
            ' Add totals row
            totalsRow = wsWorking.Cells(wsWorking.Rows.Count, 1).End(xlUp).Row + 1
            With wsWorking
                .Cells(totalsRow, 1).Value = "TOTALS"
                .Cells(totalsRow, 16).Formula = "=SUM(" & .Range(.Cells(2, 16), .Cells(totalsRow - 1, 16)).Address & ")"
                .Cells(totalsRow, 17).Formula = "=SUM(" & .Range(.Cells(2, 17), .Cells(totalsRow - 1, 17)).Address & ")"
                .Cells(totalsRow, 19).Formula = "=SUM(" & .Range(.Cells(2, 19), .Cells(totalsRow - 1, 19)).Address & ")"
                .Cells(totalsRow, 22).Formula = "=SUM(" & .Range(.Cells(2, 22), .Cells(totalsRow - 1, 22)).Address & ")"
                .Range("U:U,X:X").NumberFormat = "0.00%"
                .Range("A" & totalsRow & ":Y" & totalsRow).Font.Bold = True
                .columns.AutoFit
            End With
            
            ' Save temp file
            tempFilePath = wbCA.path & "\" & CleanFileName(CStr(brandData(j, 1))) & ".xlsx"
            wbBrandReport.SaveAs tempFilePath
            wbBrandReport.Close False
            
            ' Add temp file to ZIP with error handling
            On Error Resume Next
            AddFileToZip tempFilePath, currentZIP
            If Err.Number <> 0 Then
                logWS.Cells(logRow, 1).Value = CStr(brandData(j, 1))
                logWS.Cells(logRow, 2).Value = tempFilePath
                logWS.Cells(logRow, 3).Value = currentZIP
                logWS.Cells(logRow, 4).Value = "ERROR adding to ZIP: " & Err.Description
                logWS.Cells(logRow, 5).Value = Now
                Err.Clear
            Else
                ' Log success
                logWS.Cells(logRow, 1).Value = CStr(brandData(j, 1))
                logWS.Cells(logRow, 2).Value = tempFilePath
                logWS.Cells(logRow, 3).Value = currentZIP
                logWS.Cells(logRow, 4).Value = "Success"
                logWS.Cells(logRow, 5).Value = Now
            End If
            logRow = logRow + 1
            On Error GoTo 0
            
            ' Delete temp file
            If Dir(tempFilePath) <> "" Then Kill tempFilePath
            
            j = j + 1
        Loop
        
        i = j ' move to next ZIP
NextIteration:
    Loop
    
    ' Complete progress bar
    UpdateProgressBar 1
    Application.Wait Now + TimeValue("0:00:01")
    Unload ProgressBar
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "YOY Brand ZIP files completed successfully. See YOY_Log for details.", vbInformation

End Sub

'====================================================
' Create empty ZIP
'====================================================
Sub CreateEmptyZip(ByVal zipPath As Variant)
    If Dir(zipPath) <> "" Then Exit Sub
    Open zipPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

'====================================================
' Add a single file to ZIP
'====================================================
Sub AddFileToZip(ByVal filePath As Variant, ByVal zipPath As Variant)
    Dim ShellApp As Object
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(zipPath).CopyHere filePath
    Application.Wait Now + TimeValue("0:00:02")
End Sub

'====================================================
' Clean filename
'====================================================
Function CleanFileName(ByVal s As Variant) As String
    Dim badChars As Variant, c As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each c In badChars
        s = Replace(s, c, "")
    Next c
    CleanFileName = Trim(CStr(s))
End Function

'====================================================
' Update filling progress bar
'====================================================
Sub UpdateProgressBar(ByVal pct As Double)
    ' pct = 0 to 1
    With ProgressBar
        .lblBarFill.Width = .FrameBarBackground.Width * pct
        .lblPercent.Caption = Format(pct, "0%") & " Complete"
        DoEvents
    End With
End Sub

