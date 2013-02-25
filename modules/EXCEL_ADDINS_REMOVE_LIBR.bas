Attribute VB_Name = "EXCEL_ADDINS_REMOVE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Sub REMOVE_OLD_ADDINS_LINKS_FUNC()

Dim SRC_WSHEET As Excel.Worksheet
Const PUB_ADDINS1_STR As String = "'C:\Users\rafael_nicolas\AppData\Roaming\Microsoft\AddIns\NF_BLACK_BOX.xlam'!"
Const PUB_ADDINS2_STR As String = "'C:\Users\nfermincota\AppData\Roaming\Microsoft\AddIns\NF_BLACK_BOX.xlam'!"

On Error GoTo ERROR_LABEL

Call EXCEL_TURN_OFF_EVENTS_FUNC
For Each SRC_WSHEET In ActiveWorkbook.Worksheets
    SRC_WSHEET.Cells.Replace _
        What:=PUB_ADDINS1_STR, _
        Replacement:="", _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        MatchCase:=False, _
        SearchFormat:=False, _
        ReplaceFormat:=False
    SRC_WSHEET.Cells.Replace _
        What:=PUB_ADDINS2_STR, _
        Replacement:="", _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        MatchCase:=False, _
        SearchFormat:=False, _
        ReplaceFormat:=False
Next SRC_WSHEET
Call EXCEL_TURN_ON_EVENTS_FUNC

Exit Sub
ERROR_LABEL:
Call EXCEL_TURN_ON_EVENTS_FUNC
End Sub


