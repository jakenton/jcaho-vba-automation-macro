Attribute VB_Name = "JCAHO_Report_Automation"
' ------------------------------------------------------------
' JCAHO Phase 1 Macro
' ------------------------------------------------------------

Option Explicit

Sub PrepareJCAHOPhase1()

    ' ----------------------------------------------------------------
    ' PURPOSE:
    ' Prepare the weekly JCAHO report before Epic review.
    '
    ' This macro does the following:
    ' 1. Removes the report title row.
    ' 2. Removes the footer row.
    ' 3. Changes the workbook font to Calibri.
    ' 4. Creates the standard report tabs.
    ' 5. Copies raw data to the >30 day sheet.
    ' 6. Moves <30 day records to the <30 day sheet.
    ' 7. Removes "Pathology Report" rows from the >30 day count sheet.
    ' 8. Moves "Pathology Report" rows from 'Sheet 1' to 'Pathology.'
    ' 9. Deletes unnecessary columns from 'Sheet 1.'
    ' 10. Deletes unnecessary columns from "Pathology," but keeps CUR_STATUS_NM.
    ' 11. Moves verbal order rows from 'Sheet 1' to 'Orders.'
    ' ----------------------------------------------------------------

    Dim rawSheet As Worksheet
    
    ' Assumes the raw Epic export is the first worksheet in the workbook.
    Set rawSheet = ThisWorkbook.Worksheets(1)

    Application.ScreenUpdating = False

    CleanRawExport rawSheet

    ' Set workbook font to Calibri for readability.
    ' This avoids selecting all worksheets at once, which can cause errors.
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Name = "Calibri"
    Next ws

    CreateJCAHOTabs

    BuildThirtyDaySheets rawSheet

    ' Move Pathology rows BEFORE deleting columns so Pathology can keep CUR_STATUS_NM.
    MoveRowsByDefType rawSheet, "Pathology", Array("Pathology Report")

    ' Delete unnecessary columns from 'Sheet 1.'
    DeleteUnneededColumns rawSheet

    'Delete unnecessary columns from 'Pathology,' but keep CUR_STATUS_NM.
    DeleteUnneededColumnsKeepCurStatus ThisWorkbook.Worksheets("Pathology")

    ' Move verbal order rows from 'Sheet 1' to 'Orders.'
    MoveRowsByDefType rawSheet, "Orders", Array("Verbal Treatment Plan", "Verbal/Cosign Order")

    Application.ScreenUpdating = True

    MsgBox "Phase 1 JCAHO preparation is complete.", vbInformation

End Sub


Sub CleanRawExport(rawSheet As Worksheet)

    ' Removes the report title row and footer from the raw Epic export.

    Dim lastRow As Long
    
    ' Delete Row 1 because it contains the report title, not column headers.
    rawSheet.Rows(1).Delete

    ' Find the last used row in Column A.
    lastRow = rawSheet.Cells(rawSheet.Rows.Count, "A").End(xlUp).Row

    ' Delete the footer only if Column A says "JCAHO Details."
    If rawSheet.Cells(lastRow, "A").Value = "JCAHO Details" Then
        rawSheet.Rows(lastRow).Delete
    End If

End Sub


Sub CreateJCAHOTabs()

    ' Creates the standard report tabs if they do not already exist.

    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim sheetExists As Boolean

    sheetNames = Array( _
        "Orders", _
        "Pathology", _
        "Stats", _
        "Suspended", _
        "All Deficiencies > 30 Days", _
        "All Deficiencies < 30 Days", _
        "> 30 Days Total Patients" _
    )

    For Each sheetName In sheetNames

        sheetExists = False
        
        ' Check wheterh this sheet already exists.
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = CStr(sheetName) Then
                sheetExists = True
                Exit For
            End If
        Next ws

        ' If it does not exist, create new sheet and rename THAT specific new sheet.
        If sheetExists = False Then
            Set newSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            newSheet.Name = CStr(sheetName)
        End If

    Next sheetName

End Sub


Sub BuildThirtyDaySheets(rawSheet As Worksheet)

    ' Copies all raw data to the > 30 Day sheet.
    ' Moves 0-15 day and 15-30 day rows into the < 30 day sheet.
    ' Removes "Pathology Report" rows from the > 30 day count sheet.

    Dim over30Sheet As Worksheet
    Dim under30sheet As Worksheet
    Dim ageCol As Long
    Dim defTypeCol As Long
    Dim lastRow As Long
    Dim i As Long
    Dim nextRow As Long
    Dim ageValue As String
    Dim defTypeValue As String

    Set over30Sheet = ThisWorkbook.Worksheets("All Deficiencies > 30 Days")
    Set under30sheet = ThisWorkbook.Worksheets("All Deficiencies < 30 Days")

    over30Sheet.Cells.Clear
    under30sheet.Cells.Clear

    rawSheet.UsedRange.Copy Destination:=over30Sheet.Range("A1")

    ageCol = FindColumnByHeader(over30Sheet, "LST_ACT_AGE_BKT")
    defTypeCol = FindColumnByHeader(over30Sheet, "DEF_TYPE_NM")

    If ageCol = 0 Or defTypeCol = 0 Then
        MsgBox "Could not find LST_ACT_AGE_BKT or DEF_TYPE_NM column.", vbCritical
        Exit Sub
    End If

    over30Sheet.Rows(1).Copy Destination:=under30sheet.Rows(1)

    lastRow = over30Sheet.Cells(over30Sheet.Rows.Count, "A").End(xlUp).Row

    For i = lastRow To 2 Step -1

        ageValue = Trim(over30Sheet.Cells(i, ageCol).Value)
        defTypeValue = Trim(over30Sheet.Cells(i, defTypeCol).Value)

        If defTypeValue = "Pathology Report" Then
            over30Sheet.Rows(i).Delete

        ElseIf ageValue = "0-15 days" Or ageValue = "0-15 days" _
            Or ageValue = "15-30 days" Or ageValue = "15-30 days" Then
        
            nextRow = under30sheet.Cells(under30sheet.Rows.Count, "A").End(xlUp).Row + 1
            over30Sheet.Rows(i).Copy Destination:=under30sheet.Rows(nextRow)
            over30Sheet.Rows(i).Delete

        End If

    Next i

End Sub


Sub DeleteUnneededColumns(rawSheet As Worksheet)

    ' Deletes unnecessary columns from Sheet 1
    ' This version deletes CUR_STATUS_NM.

    Dim columnsToDelete As Variant
    Dim colName As Variant
    Dim colNumber As Long

    columnsToDelete = Array( _
        "DFI_ID", _
        "DEPARTMENT_ID", _
        "PAT_ID", _
        "LOC_ID", _
        "DEF_ID", _
        "DEF_TYPE_C", _
        "ORG_USER_ID", _
        "ORG_STATUS", _
        "CUR_USER_ID", _
        "LST_ACTION_AGE", _
        "CUR_STATUS", _
        "CUR_STATUS_NM", _
        "AUDIT_ASGN_PROV_ID", _
        "PROV_NAME", _
        "DELINQUENCY_DATE", _
        "CHART_ID", _
        "CONTACT_DATE_REAL", _
        "PAT_CLASS_C", _
        "PAT_CLASS", _
        "DISCHARGE_DATE", _
        "COMPLETED_DATE", _
        "ADMIN_CLOSE_DTTM", _
        "END_DATE" _
    )

    For Each colName In columnsToDelete

        colNumber = FindColumnByHeader(rawSheet, CStr(colName))

        If colNumber > 0 Then
            rawSheet.Columns(colNumber).Delete
        End If

    Next colName

End Sub


Sub DeleteUnneededColumnsKeepCurStatus(targetSheet As Worksheet)

    ' Deletes unnecessary columns from Pathology worksheet, but keeps CUR_STATUS_NM for Pathology Report review

    Dim columnsToDelete As Variant
    Dim colName As Variant
    Dim colNumber As Long

    columnsToDelete = Array( _
        "DFI_ID", _
        "DEPARTMENT_ID", _
        "PAT_ID", _
        "LOC_ID", _
        "DEF_ID", _
        "DEF_TYPE_C", _
        "ORG_USER_ID", _
        "ORG_STATUS", _
        "CUR_USER_ID", _
        "LST_ACTION_AGE", _
        "CUR_STATUS", _
        "AUDIT_ASGN_PROV_ID", _
        "PROV_NAME", _
        "DELINQUENCY_DATE", _
        "CHART_ID", _
        "CONTACT_DATE_REAL", _
        "PAT_CLASS_C", _
        "PAT_CLASS", _
        "DISCHARGE_DATE", _
        "COMPLETED_DATE", _
        "ADMIN_CLOSE_DTTM", _
        "END_DATE" _
    )

    For Each colName In columnsToDelete
    
        colNumber = FindColumnByHeader(targetSheet, CStr(colName))

        If colNumber > 0 Then
            targetSheet.Columns(colNumber).Delete
        End If

    Next colName

End Sub


Sub MoveRowsByDefType(rawSheet As Worksheet, destinationSheetName As String, defTypes As Variant)

    ' Moves rows from 'Sheet 1' to a destination sheet based on DEF_TYPE_NM.

    Dim destinationSheet As Worksheet
    Dim defTypeCol As Long
    Dim lastRow As Long
    Dim nextRow As Long
    Dim i As Long
    Dim defType As Variant
    Dim currentDefType As String
    Dim shouldMove As Boolean

    Set destinationSheet = ThisWorkbook.Worksheets(destinationSheetName)

    destinationSheet.Cells.Clear

    defTypeCol = FindColumnByHeader(rawSheet, "DEF_TYPE_NM")

    If defTypeCol = 0 Then
        MsgBox "Could not find DEF_TYPE_NM column on Sheet 1.", vbCritical
        Exit Sub
    End If

    rawSheet.Rows(1).Copy Destination:=destinationSheet.Rows(1)

    lastRow = rawSheet.Cells(rawSheet.Rows.Count, "A").End(xlUp).Row

    For i = lastRow To 2 Step -1

        currentDefType = Trim(rawSheet.Cells(i, defTypeCol).Value)
        shouldMove = False

        For Each defType In defTypes
            If currentDefType = CStr(defType) Then
                shouldMove = True
                Exit For
            End If
        Next defType
        
        If shouldMove = True Then
            nextRow = destinationSheet.Cells(destinationSheet.Rows.Count, "A").End(xlUp).Row + 1
            rawSheet.Rows(i).Copy Destination:=destinationSheet.Rows(nextRow)
            rawSheet.Rows(i).Delete
        End If
        
    Next i
    
End Sub



