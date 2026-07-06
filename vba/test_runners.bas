Attribute VB_Name = "test_runners"
Option Explicit

Sub Test_01_CreateJCAHOTabs()

    ' --------------------------------------
    ' TEST PURPOSE:
    ' Test only the worksheet creation logic.
    '
    ' This macro does NOT:
    '   - clean the raw report
    '   - move rows
    '   - delete columns
    '
    ' Expected result:
    ' The workbook should contain these tabs:
    '   - Orders
    '   - Pathology
    '   - Stats
    '   - Suspended
    '   - All Deficiencies > 30 Days
    '   - All Deficiencies < 30 Days
    '   - > 30 Days Total Patients
    ' --------------------------------------

    MsgBox "Starting Test 01: Create JCAHO tabs.", vbInformation
    
    CreateJCAHOTabs
    
    MsgBox "Test 01 complete. Check the worksheet tabs at the bottom of Excel.", vbInformation
    
End Sub

Sub Test_02_CleanRawExport()

    ' ----------------------------------------------------------
    ' TEST PURPOSE:
    ' Test only the raw report cleanup logic.
    '
    ' This macro does NOT:
    '   - create worksheets
    '   - move rows
    '   - delete columns
    '
    ' Expected result:
    '   - Report title (Row 1) is removed.
    '   - Column headers move to Row 1.
    '   - Footer row beginning with "JCAHO Details" is removed.
    ' ----------------------------------------------------------
    
    Dim rawSheet As Worksheet
    Set rawSheet = ThisWorkbook.Worksheets(1)
    
    MsgBox "Starting Test 02: Remove title row and footer row.", vbInformation
    
    CleanRawExport rawSheet

    MsgBox "Test 02 complete. Confirm Row 1 now contains column headers and footer is gone."

End Sub

Sub Test_03_BuildThirtyDaySheet()

    ' -----------------------------------------------------------------------------------------------
    ' TEST PURPOSE:
    ' Test only the creation of the >30 Day and <30 Day worksheets.
    '
    ' This macro does NOT:
    '   - clean the raw report
    '   - move Pathology rows
    '   - move Orders rows
    '   - delete columns
    '
    ' Expected result:
    '   - "All Decifiencies > 30 Days" contains only deficiencies older than 30 days.
    '   - "All Deficiencies < 30 Days" contains deficiencies in the 0-15 and 15-30 day age buckets.
    '   - "Pathology Report" rows are excluded form the >30 Day worksheet.
    ' -----------------------------------------------------------------------------------------------

    Dim rawSheet As Worksheet
    Set rawSheet = ThisWorkbook.Worksheets(1)
    
    MsgBox "Starting Test 03: Build >30 and <30 day sheets.", vbInformation
    
    CreateJCAHOTabs
    BuildThirtyDaySheets rawSheet
    
    MsgBox "Test 03 complete. Check the >30 and <30 day worksheets.", vbInformation
    
End Sub

Sub Test_04_DeleteUnneededColumns()

    ' -----------------------------------------------------------------------------------------------
    ' TEST PURPOSE:
    ' Test only the removal of unnecessary columns from Sheet 1.
    '
    ' This macro does NOT:
    '   - clean the raw report
    '   - create worksheets
    '   - move rows
    '
    ' Expected result:
    '   - Only the columns required for Epic review remain on Sheet 1.
    '   - CUR_STATUS_NM is removed.
    ' -----------------------------------------------------------------------------------------------

    Dim rawSheet As Worksheet
    Set rawSheet = ThisWorkbook.Worksheets(1)
    
    MsgBox "Starting Test 04: Delete unnecessary columns from Sheet 1.", vbInformation
    
    DeletedUnneededColumns rawSheet
    
    MsgBox "Test 04 complete. Confirm only the needed columns remain on Sheet 1.", vbInformation

End Sub

Sub Test_05_MovePathologyRows()

    ' -----------------------------------------------------------------------------------------------
    ' TEST PURPOSE:
    ' Test only the movement of "Pathology Report" rows from Sheet1 to the Pathology worksheet.
    '
    ' This macro does NOT:
    '   - clean the raw report
    '   - delete worksheets
    '   - move Orders rows
    '
    ' Expected result:
    '   - All "Pathology Report" rows are moved to the Pathology worksheet.
    '   - Those rows are removed from Sheet 1.
    ' -----------------------------------------------------------------------------------------------

    Dim rawSheet As Worksheet
    Set rawSheet = ThisWorkbook.Worksheets(1)
    
    MsgBox "Starting Test 05: Move Pathology Report rows.", vbInformation
    
    CreateJCAHOTabs
    MoveRowsByDefType rawSheet, "Pathology", Array("Pathology Report")
    
    MsgBox "Test 05 complete. Check Pathology sheet and confirm rows were removed from Sheet 1.", vbInformation
    
End Sub

Sub Test_06_DeleteUnneededColumnsKeepCurStatus()

    ' -----------------------------------------------------------------------------------------------
    ' TEST PURPOSE:
    ' Test only the removal of unnecessary columns from the Pathology worksheet.
    '
    ' This macro does NOT:
    '   - clean the raw report
    '   - move rows
    '   - modify Sheet 1
    '
    ' Expected result:
    '   - Unnecessary columns are removed.
    '   - CUR_STATUS_NM remains available for Pathology review.
    ' -----------------------------------------------------------------------------------------------

    Dim pathologySheet As Worksheet
    
    MsgBox "Starting Test 06: Delete Pathology columns but keep CUR_STATUS_NM.", vbInformation
    
    CreateJCAHOTabs
    Set pathologySheet = ThisWorkbook.Worksheets("Pathology")
    
    DeleteUnneededColumnsKeepCurStatus pathologySheet
    
    MsgBox "Test 06 complete. Confirm CUR_STATUS_NM remains on Pathology sheet.", vbInformation
    
End Sub

Sub Test_07_MoveOrderRows()

    ' -----------------------------------------------------------------------------------------------
    ' TEST PURPOSE:
    ' Test only the movement of verbal order deficiencies from Sheet 1 to Orders.
    '
    ' This macro does NOT:
    '   - clean the raw report
    '   - delete columns
    '   - move Pathology rows
    '
    ' Expected result:
    '   - "Verbal Treatment Plan" rows are moved to Orders.
    '   - "Verbal/Cosign Order" rows are moved to Orders.
    '   - Those rows are removed from Sheet 1.
    ' -----------------------------------------------------------------------------------------------

    Dim rawSheet As Workheet
    Set rawSheet = ThisWorkbook.Worksheets(1)
    
    MsgBox "Starting Test 07: Move verbal order rows.", vbInformation
    
    MsgBox "Test 07 complete. Check Orders sheet and confirm rows were removed from Sheet 1."
    
End Sub

Sub Test_08_FullPhase1Workflow()

    ' -----------------------------------------------------------------------------------------------
    ' TEST PURPOSE:
    ' Test the complete Phase 1 workflow.
    '
    ' This macro performs the following:
    '   - report cleanup
    '   - worksheet creation
    '   - >30/<30 day separation
    '   - Pathology row movement
    '   - column removal
    '   - Orders row movement
    '
    ' Expected result:
    '   - The workbook is fully prepared for Epic review using the standard Phase 1 workflow.
    ' -----------------------------------------------------------------------------------------------

    MsgBox "Starting Test 08: Full Phase 1 workflow.", vbInformation
    
    PrepareJCAHOPhase1
    
    MsgBox "Test 08 complete. Review all tabs and row counts.", vbInformation
    
End Sub