Sub DeleteUnneededColumnsKeepCurStatus(targetSheet As Worksheet)

    ' Deletes unnecessary columns from a worksheet, but keeps CUR_STATUS_NM for Pathology Report review.

    Dim columnsToDelete As Variant
    Dim colName As Variant
    Dim colNumer As Long

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
        
    )