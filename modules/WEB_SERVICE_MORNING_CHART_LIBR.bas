Attribute VB_Name = "WEB_SERVICE_MORNING_CHART_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MORNINGSTAR_FUNDAMENTAL_CHARTS_FUNC
'DESCRIPTION   :
'LIBRARY       : WEB_SERVICE
'GROUP         : MORNINGSTAR_CHART
'ID            : 001
'LAST UPDATE   : 28/05/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function MORNINGSTAR_FUNDAMENTAL_CHARTS_FUNC( _
ByRef TICKERS_RNG As Variant, _
ByRef NAMES_RNG As Variant, _
Optional ByRef SUFFIX_STR As String = _
"PB,PC,PE,PS,RG,OIG,EPSG,EQG,CFO,EPS,ROEG10,ROAG10,PROA,ROEA,TOTR,CR,DE,DTC", _
Optional ByVal DELIM_CHR As String = ",")

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim TEMP_FLAG As Boolean
Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim NAMES_VECTOR As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

SUFFIX_STR = Trim(SUFFIX_STR)
MORNINGSTAR_FUNDAMENTAL_CHARTS_FUNC = False

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then: _
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    
    NAMES_VECTOR = NAMES_RNG
    If UBound(NAMES_VECTOR, 1) = 1 Then: _
        NAMES_VECTOR = MATRIX_TRANSPOSE_FUNC(NAMES_VECTOR)
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
    
    ReDim NAMES_VECTOR(1 To 1, 1 To 1)
    NAMES_VECTOR(1, 1) = NAMES_RNG
End If

If UBound(TICKERS_VECTOR, 1) <> UBound(NAMES_VECTOR, 1) Then: _
GoTo ERROR_LABEL

NROWS = UBound(TICKERS_VECTOR, 1)
NSIZE = COUNT_CHARACTERS_FUNC(SUFFIX_STR, DELIM_CHR) + 1

ReDim ATEMP_VECTOR(1 To NROWS * NSIZE, 1 To 1)
ReDim BTEMP_VECTOR(1 To NROWS * NSIZE, 1 To 1)

kk = 1
For ii = 1 To NROWS
    j = 0
    For jj = 1 To NSIZE
        If jj <> NSIZE Then
            i = j + 1
            j = InStr(i, SUFFIX_STR, DELIM_CHR)
        Else
            i = j + 1
            j = Len(SUFFIX_STR) + 1
        End If
        ATEMP_VECTOR(kk, 1) = _
        TICKERS_VECTOR(ii, 1) & DELIM_CHR & Mid(SUFFIX_STR, i, j - i)
        BTEMP_VECTOR(kk, 1) = NAMES_VECTOR(ii, 1)
        kk = kk + 1
    Next jj
Next ii

TEMP_FLAG = _
PRINT_WEB_FINANCIAL_CHARTS_FUNC(ATEMP_VECTOR, BTEMP_VECTOR, 5, NSIZE, "")
If TEMP_FLAG = False Then: GoTo ERROR_LABEL

MORNINGSTAR_FUNDAMENTAL_CHARTS_FUNC = True

Exit Function
ERROR_LABEL:
MORNINGSTAR_FUNDAMENTAL_CHARTS_FUNC = False
End Function


