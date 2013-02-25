Attribute VB_Name = "FINAN_DERIV_BID_ASK_AVG_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'************************************************************************************
'************************************************************************************
'FUNCTION      : AVERAGE_OPTION_BID_ASK_TABLE_FUNC
'DESCRIPTION   : Avg. Bid Ask Table
'LIBRARY       : DERIVATIVES
'GROUP         : BID_ASK_AVG
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************
'
Function AVERAGE_OPTION_BID_ASK_TABLE_FUNC(ByRef DATA_MATRIX As Variant)

Dim ii As Long
Dim jj As Long
  
Dim TEMP_ARR As Variant

Dim FIRST_GROUP As Variant
Dim SECOND_GROUP As Variant

Dim TEMP_MATRIX As Variant
'Dim DATA_MATRIX As Variant
  
Dim CALL_BID_VALUE As Double
Dim CALL_ASK_VALUE As Double
Dim PUT_BID_VALUE As Double
Dim PUT_ASK_VALUE As Double
    
Dim CALL_BID_TOTAL As Double
Dim CALL_ASK_TOTAL As Double
Dim PUT_BID_TOTAL As Double
Dim PUT_ASK_TOTAL As Double
  
Dim CALL_BID_COUNTER As Long
Dim CALL_ASK_COUNTER As Long
Dim PUT_BID_COUNTER As Long
Dim PUT_ASK_COUNTER As Long
  
On Error GoTo ERROR_LABEL

'DATA_MATRIX = DATA_RNG
DATA_MATRIX = MATRIX_DOUBLE_SORT_FUNC(DATA_MATRIX)
SECOND_GROUP = AGGREGATE_OPTION_BID_ASK_QUOTES_FUNC(DATA_MATRIX)

ReDim TEMP_ARR(1 To UBound(SECOND_GROUP), 1 To 6)

For ii = 1 To UBound(SECOND_GROUP)
    
    FIRST_GROUP = SECOND_GROUP(ii)
    TEMP_ARR(ii, 1) = FIRST_GROUP(1)
    TEMP_ARR(ii, 2) = FIRST_GROUP(2)
    
    TEMP_MATRIX = FIRST_GROUP(3)
    
    CALL_BID_VALUE = 0: CALL_ASK_VALUE = 0: PUT_BID_VALUE = 0: PUT_ASK_VALUE = 0
    CALL_BID_TOTAL = 0: CALL_ASK_TOTAL = 0: PUT_BID_TOTAL = 0: PUT_ASK_TOTAL = 0
    CALL_BID_COUNTER = 0: CALL_ASK_COUNTER = 0: PUT_BID_COUNTER = 0: PUT_ASK_COUNTER = 0
    
    For jj = 1 To UBound(TEMP_MATRIX, 1)
        If TEMP_MATRIX(jj, 1) > 0 Then
            CALL_BID_TOTAL = CALL_BID_TOTAL + TEMP_MATRIX(jj, 1)
            CALL_BID_COUNTER = CALL_BID_COUNTER + 1
        End If
        If TEMP_MATRIX(jj, 2) > 0 Then
            CALL_ASK_TOTAL = CALL_ASK_TOTAL + TEMP_MATRIX(jj, 2)
            CALL_ASK_COUNTER = CALL_ASK_COUNTER + 1
        End If
        If TEMP_MATRIX(jj, 3) > 0 Then
            PUT_BID_TOTAL = PUT_BID_TOTAL + TEMP_MATRIX(jj, 3)
            PUT_BID_COUNTER = PUT_BID_COUNTER + 1
        End If
        If TEMP_MATRIX(jj, 4) > 0 Then
            PUT_ASK_TOTAL = PUT_ASK_TOTAL + TEMP_MATRIX(jj, 4)
            PUT_ASK_COUNTER = PUT_ASK_COUNTER + 1
        End If
    Next jj
    
    If CALL_BID_COUNTER > 0 Then: CALL_BID_VALUE = CALL_BID_TOTAL / CALL_BID_COUNTER
    If CALL_ASK_COUNTER > 0 Then: CALL_ASK_VALUE = CALL_ASK_TOTAL / CALL_ASK_COUNTER
    If PUT_BID_COUNTER > 0 Then: PUT_BID_VALUE = PUT_BID_TOTAL / PUT_BID_COUNTER
    If PUT_ASK_COUNTER > 0 Then: PUT_ASK_VALUE = PUT_ASK_TOTAL / PUT_ASK_COUNTER
    
    TEMP_ARR(ii, 3) = CALL_BID_VALUE
    TEMP_ARR(ii, 4) = CALL_ASK_VALUE
    TEMP_ARR(ii, 5) = PUT_BID_VALUE
    TEMP_ARR(ii, 6) = PUT_ASK_VALUE
Next ii

AVERAGE_OPTION_BID_ASK_TABLE_FUNC = TEMP_ARR

Exit Function
ERROR_LABEL:
AVERAGE_OPTION_BID_ASK_TABLE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : AGGREGATE_OPTION_BID_ASK_QUOTES_FUNC
'DESCRIPTION   : Aggregate Quotes
'LIBRARY       : DERIVATIVES
'GROUP         : BID_ASK_AVG
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function AGGREGATE_OPTION_BID_ASK_QUOTES_FUNC(ByRef DATA_MATRIX As Variant)
  
Dim i As Long
Dim j As Long
  
Dim ii As Long
Dim jj As Long
  
Dim SROW As Long
Dim NROWS As Long
  
Dim STRIKE As Double
Dim MATURITY As Date
  
Dim TEMP_ARR As Variant
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
'Dim DATA_MATRIX As Variant
  
Dim FIRST_GROUP As Variant
Dim SECOND_GROUP As Variant
  
On Error GoTo ERROR_LABEL
  
'DATA_MATRIX = DATA_RNG
  
SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)
  
MATURITY = DATA_MATRIX(1, 1)
STRIKE = DATA_MATRIX(1, 2)
  
ReDim FIRST_GROUP(1 To 3)
FIRST_GROUP(1) = MATURITY
FIRST_GROUP(2) = STRIKE
  
ReDim TEMP_MATRIX(1 To 1, 1 To 4)
TEMP_MATRIX(1, 1) = DATA_MATRIX(1, 3)
TEMP_MATRIX(1, 2) = DATA_MATRIX(1, 4)
TEMP_MATRIX(1, 3) = DATA_MATRIX(1, 5)
TEMP_MATRIX(1, 4) = DATA_MATRIX(1, 6)
FIRST_GROUP(3) = TEMP_MATRIX

ReDim SECOND_GROUP(1 To 1) 'Aggregate
SECOND_GROUP(1) = FIRST_GROUP
j = 1

For i = (SROW + 1) To NROWS 'update the option quotes mat
    MATURITY = DATA_MATRIX(i, 1)
    STRIKE = DATA_MATRIX(i, 2)
    
    If ((MATURITY = DATA_MATRIX(i - 1, 1)) And (STRIKE = DATA_MATRIX(i - 1, 2))) Then
      FIRST_GROUP = SECOND_GROUP(j)
      TEMP_MATRIX = FIRST_GROUP(3)
      
      ReDim TEMP_VECTOR(1 To 4, 1 To 1)
      TEMP_VECTOR(1, 1) = DATA_MATRIX(i, 3)
      TEMP_VECTOR(2, 1) = DATA_MATRIX(i, 4)
      TEMP_VECTOR(3, 1) = DATA_MATRIX(i, 5)
      TEMP_VECTOR(4, 1) = DATA_MATRIX(i, 6)
    
      ReDim TEMP_ARR(1 To UBound(TEMP_MATRIX, 1) + 1, 1 To UBound(TEMP_MATRIX, 2))
      'Insert Matrix Row
        
      For ii = 1 To UBound(TEMP_MATRIX, 1)
        For jj = 1 To UBound(TEMP_MATRIX, 2)
            TEMP_ARR(ii, jj) = TEMP_MATRIX(ii, jj)
        Next jj
      Next ii
        
      For jj = 1 To UBound(TEMP_MATRIX, 2)
        TEMP_ARR(ii, jj) = TEMP_VECTOR(jj, 1)
      Next jj
        
      TEMP_MATRIX = TEMP_ARR
      FIRST_GROUP(3) = TEMP_MATRIX
      SECOND_GROUP(j) = FIRST_GROUP
    Else
      ReDim FIRST_GROUP(1 To 3)
      FIRST_GROUP(1) = MATURITY
      FIRST_GROUP(2) = STRIKE
      
      ReDim TEMP_MATRIX(1 To 1, 1 To 4)
      TEMP_MATRIX(1, 1) = DATA_MATRIX(i, 3)
      TEMP_MATRIX(1, 2) = DATA_MATRIX(i, 4)
      TEMP_MATRIX(1, 3) = DATA_MATRIX(i, 5)
      TEMP_MATRIX(1, 4) = DATA_MATRIX(i, 6)
          
      FIRST_GROUP(3) = TEMP_MATRIX
      j = j + 1
      ReDim Preserve SECOND_GROUP(1 To j)
      SECOND_GROUP(j) = FIRST_GROUP
    End If
Next i
  
AGGREGATE_OPTION_BID_ASK_QUOTES_FUNC = SECOND_GROUP

Exit Function
ERROR_LABEL:
AGGREGATE_OPTION_BID_ASK_QUOTES_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GET_OPTION_BID_ASK_VALUES_FUNC
'DESCRIPTION   : Extract Bid Ask Values
'LIBRARY       : DERIVATIVES
'GROUP         : BID_ASK_AVG
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function GET_OPTION_BID_ASK_VALUES_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal COL_CALL_BID As Long = 4, _
Optional ByVal COL_CALL_ASK As Long = 5, _
Optional ByVal COL_PUT_BID As Long = 11, _
Optional ByVal COL_PUT_ASK As Long = 12, _
Optional ByVal COL_MATURITY As Long = 2, _
Optional ByVal COL_STRIKE As Long = 8, _
Optional ByVal SROW As Long = 3)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim STRIKE As Double
Dim MATURITY As Date

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)

'----------------------------------------------------------------------------------------------
Select Case VERSION
'----------------------------------------------------------------------------------------------
Case 0 'CBOE
'----------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS - SROW, 1 To 6)
    
    j = 1
    For i = (SROW + 1) To NROWS
      TEMP_VECTOR = PARSE_OPTION_BID_ASK_DATE_FUNC(DATA_MATRIX(i, 1))
      
      MATURITY = TEMP_VECTOR(1, 1)
      STRIKE = TEMP_VECTOR(2, 1)
      
      TEMP_MATRIX(j, 1) = MATURITY
      TEMP_MATRIX(j, 2) = STRIKE
      TEMP_MATRIX(j, 3) = CDec(DATA_MATRIX(i, COL_CALL_BID))
      TEMP_MATRIX(j, 4) = CDec(DATA_MATRIX(i, COL_CALL_ASK))
      TEMP_MATRIX(j, 5) = CDec(DATA_MATRIX(i, COL_PUT_BID))
      TEMP_MATRIX(j, 6) = CDec(DATA_MATRIX(i, COL_PUT_ASK))
      
      j = j + 1
    Next i
'----------------------------------------------------------------------------------------------
Case Else 'Yahoo
'----------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To 6)
    
    j = 1
    For i = 1 To NROWS
      
      TEMP_MATRIX(j, 1) = CDate(DATA_MATRIX(i, COL_MATURITY))
      TEMP_MATRIX(j, 2) = CDec(DATA_MATRIX(i, COL_STRIKE))
      TEMP_MATRIX(j, 3) = CDec(DATA_MATRIX(i, COL_CALL_BID))
      TEMP_MATRIX(j, 4) = CDec(DATA_MATRIX(i, COL_CALL_ASK))
      TEMP_MATRIX(j, 5) = CDec(DATA_MATRIX(i, COL_PUT_BID))
      TEMP_MATRIX(j, 6) = CDec(DATA_MATRIX(i, COL_PUT_ASK))
      
      j = j + 1
    Next i
'----------------------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------------------

GET_OPTION_BID_ASK_VALUES_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
GET_OPTION_BID_ASK_VALUES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PARSE_OPTION_BID_ASK_DATE_FUNC
'DESCRIPTION   : Parse Third Friday Function (weekday of thursday is 5)
'LIBRARY       : DERIVATIVES
'GROUP         : BID_ASK_AVG
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function PARSE_OPTION_BID_ASK_DATE_FUNC(ByVal DATE_STR As String)
  
Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim MONTH_STR As String
Dim YEAR_STR As String
Dim TEMP_STR As String
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To 2, 1 To 1)
  
'------------------------------------------------------------------------------
MONTH_STR = Mid(DATE_STR, 4, 3)
YEAR_STR = "20" & Left(DATE_STR, 2)
For i = 1 To 7
    TEMP_STR = MONTH_STR & " 0" & i & "," & YEAR_STR
    NSIZE = Weekday(CDate(TEMP_STR))
    If NSIZE = 6 Then Exit For
Next i
j = 14 + i
TEMP_STR = MONTH_STR & " " & CStr(j) & "," & YEAR_STR
TEMP_VECTOR(1, 1) = CDate(TEMP_STR)
'------------------------------------------------------------------------------
k = InStr(1, DATE_STR, "(", vbTextCompare)
If k = 0 Then: GoTo ERROR_LABEL
TEMP_VECTOR(2, 1) = CDec(Mid(DATE_STR, 7, k - 7)) 'Out STRIKE
'------------------------------------------------------------------------------
PARSE_OPTION_BID_ASK_DATE_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
PARSE_OPTION_BID_ASK_DATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GET_OPTION_BID_ASK_PRICE_FUNC
'DESCRIPTION   : OPTION SPOT PRICE
'LIBRARY       : DERIVATIVES
'GROUP         : BID_ASK_AVG
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function GET_OPTION_BID_ASK_PRICE_FUNC(ByVal FULL_PATH_NAME As String)
On Error GoTo ERROR_LABEL
GET_OPTION_BID_ASK_PRICE_FUNC = _
    CDec(CONVERT_TEXT_FILE_MATRIX_FUNC(FULL_PATH_NAME, 5, 4, ",")(1, 2))
Exit Function
ERROR_LABEL:
GET_OPTION_BID_ASK_PRICE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GET_OPTION_BID_ASK_SETTLEMENT_FUNC
'DESCRIPTION   : OPT SETTLEMENT DATE
'LIBRARY       : DERIVATIVES
'GROUP         : BID_ASK_AVG
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function GET_OPTION_BID_ASK_SETTLEMENT_FUNC(ByVal FULL_PATH_NAME As String)
Dim DATE_STR As String

On Error GoTo ERROR_LABEL
DATE_STR = CONVERT_TEXT_FILE_MATRIX_FUNC(FULL_PATH_NAME, 5, 4, ",")(2, 1)
'Jun 07 2006 @ 19:24 ET (Data 20 Minutes Delayed)
DATE_STR = Left(DATE_STR, 6) & "," & Mid(DATE_STR, 8, 4)
GET_OPTION_BID_ASK_SETTLEMENT_FUNC = CDate(DATE_STR)

Exit Function
ERROR_LABEL:
GET_OPTION_BID_ASK_SETTLEMENT_FUNC = Err.number
End Function
