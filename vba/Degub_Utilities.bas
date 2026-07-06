Attribute VB_Name = "Degub_Utilities"
Option Explicit

Sub Debug_ListWorksheetNames()

    ' ---------------------------------------------------
    ' PURPOSE:
    ' Print all worksheet names to the Immediate Window.
    '
    ' How to view results:
    '   - Press Ctrl + G in the VBA Editor
    '   - Look at the Immediate Window
    ' ---------------------------------------------------
    
    Dim ws As Worksheet
    
    Debug.Print "---------------------------------"
    Debug.Print "Worksheet names in this workbook:"
    Debug.Print "---------------------------------"
    
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.Index & ": " & ws.Name
    Next ws
    
End Sub

Sub Debug_CheckRequiredTabs()

    ' ---------------------------------------------------
    ' PURPOSE:
    ' Check whether all required JCAHO tabs exist.
    ' Results print to the Immediate Window.
    ' ---------------------------------------------------
    
    Dim requiredTabs As Variant
    Dim tabName As Variant
    
    requiredTabs = Array( _
        "Orders", _
        "Pathology", _
        "Stats", _
        "Suspended", _
        "All Deficiencies > 30 Days", _
        "All Deficiencies < 30 Days", _
        "> 30 Days Total Patients" _
    )

    Debug.Print "---------------------------------"
    Debug.Print "Required tab check:"
    Debug.Print "---------------------------------"

    For Each tabName In requiredTabs
        If WorksheetExists(CStr(tabName)) Then
            Debug.Print "FOUND:     " & CStr(tabName)
        Else
            Debug.Print "MISSING:   " & CStr(tabName)
        End If
    Next tabName
    
End Sub

Sub Debug_CheckRequiredHeaders()

    ' ---------------------------------------------------
    ' PURPOSE:
    ' Check whether key headers exist on Sheet 1.
    ' Results print to the Immediate Window.
    ' ---------------------------------------------------

    Dim rawSheet As Worksheet
    Dim requiredHeaders As Variant
    Dim headerName As Variant
    Dim colNumber As Long
    
    Set rawSheet = ThisWorkbook.Worksheets(1)
    
    requiredHeaders = Array( _
        "DEF_TYPE_NM", _
        "LST_ACT_AGE_BKT", _
        "CUR_STATUS_NM", _
        "PAT_NAME", _
        "MRN" _
    )
    
    Debug.Print "-----------------------------------------------"
    Debug.Print "Required header check on " & rawSheet.Name & ":"
    Debug.Print "-----------------------------------------------"
    
    For Each headerName In requiredHeaders
    
        colNumber = FindColumnByHeader(rawSheet, CStr(headerName))
        
        If colNumber > 0 Then
            Debug.Print "FOUND:     " & CStr(headerName) & " in columns " & colNumber
        Else
            Debug.Print "MISSING:   " & CStr(headerName)
        End If
        
    Next headerName

End Sub

Sub Debug_PrintSheetRowCounts()

    ' ---------------------------------------------------
    ' PURPOSE:
    ' Print the last used row for each worksheet.
    ' Useful for checking whether rows moved correctly.
    ' Results print to the Immediate Window.
    ' ---------------------------------------------------
    
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Debug.Print "------------------------"
    Debug.Print "Worksheet row counts:"
    Debug.Print "------------------------"
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        Debug.Print ws.Name & ": last used row in Column A = " & lastRow
    Next ws
    
End Sub

Sub Debug_PrintDefTypeCountsOnSheet1()

    ' ------------------------------------------------------
    ' PURPOSE:
    ' Count DEF_TYPE_NM values on Sheet 1.
    ' Useful before and after moving Pathology/Orders rows.
    ' Results print to the Immediate Window.
    ' ------------------------------------------------------
    
    Dim rawSheet As Worksheet
    Dim defTypeCol As Long
    Dim lastRow As Long
    Dim i As Long
    Dim value As String
    
    Dim pathologyCount As Long
    Dim verbalTreatmentPlanCount As Long
    Dim verbalTreatmentOrderCount As Long
    Dim otherCount As Long
    
    Set rawSheet = ThisWorkbook.Worksheets(1)
    
    defTypeCol = FindColumnByHeader(rawSheet, "DEF_TYPE_NM")
    
    If defTypeCol = 0 Then
        Debug.Print "DEF_TYPE_NM column not found on " & rawSheet.Name
        Exit Sub
    End If
    
    lastRow = rawSheet.Cells(rawSheet.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
    
        value = Trim(rawSheet.Cells(i, defTypeCol).value)
        
        If value = "Pathology Report" Then
            pathologyCount = pathologyCount + 1
        ElseIf value = "Value Treatment Plan" Then
            verbalTreatmentPlanCount = verbalTreatmentPlanCount + 1
        ElseIf value = "Verbal/Cosign Order" Then
            verbalCosignOrderCount = verbalCosignOrderCount + 1
        Else
            otherCount = otherCount + 1
        End If
        
    Next i
    
    Debug.Print "-------------------------------------------"
    Debug.Print "DEF_TYPE_NM count on " & rawSheet.Name & ":"
    Debug.Print "-------------------------------------------"
    Debug.Print "Pathology Report: " & pathologyCount
    Debug.Print "Verbal Treatment Plan: " & verbalTreatmentPlanCount
    Debug.Print "Verbal/Cosign Order: " & verbalCosignOrderCount
    Debug.Print "Other: " & otherCount

End Sub

Function WorksheetExists(sheetName As String) As Boolean

    ' ------------------------------------
    ' PURPOSE:
    ' Return TRUE if a worksheet exists.
    ' Return FALSE if it does not exist.
    ' ------------------------------------
    
    Dim ws As Worksheet
    
    WorksheetExists = False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws

End Function
