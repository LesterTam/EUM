Attribute VB_Name = "MATRIX_GROUP_LIBR"

'-----------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-----------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_GROUP_FUNC
'DESCRIPTION   : Aggregate Vector & Matrix --> Creating a Group of Data
'LIBRARY       : MATRIX
'GROUP         : GROUP
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_GROUP_FUNC(ByRef KEY_VECTOR As Variant, _
ByRef DATA_MATRIX As Variant, _
Optional ByVal OUTPUT As Integer = 0)
  
Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long '--> Temp Index No.
Dim l As Long

Dim NSIZE As Long

Dim TEMP_VALUE As Variant 'key for indexing
  
Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant 'the row of data to be indexed

Dim TEMP_GROUP As Variant
Dim TEMP_MATRIX As Variant

Dim KEY_GROUP_ARR As Variant
Dim DATA_GROUP_ARR As Variant

On Error GoTo ERROR_LABEL
  
ReDim DATA_GROUP_ARR(0 To 0)
ReDim KEY_GROUP_ARR(0 To 0)

'------------------------------------------------------------------------------
For i = 1 To UBound(KEY_VECTOR, 1)
'------------------------------------------------------------------------------
    TEMP_VALUE = KEY_VECTOR(i) '---> ONE DIMENSION ARRAY
    
    ReDim ATEMP_VECTOR(1 To UBound(DATA_MATRIX, 2)) 'ReadMatrixRowIntoArray
    For j = 1 To UBound(DATA_MATRIX, 2)
      ATEMP_VECTOR(j) = DATA_MATRIX(i, j)
    Next j
    
    BTEMP_VECTOR = ATEMP_VECTOR
    k = -1
    For j = 1 To UBound(KEY_GROUP_ARR, 1) 'Find Group
      If TEMP_VALUE = KEY_GROUP_ARR(j) Then
        k = j
        GoTo 1983
      End If
    Next j
1983:
'------------------------------------------------------------------------------
    If k > 0 Then 'InsertRowInGroup
'------------------------------------------------------------------------------
      TEMP_GROUP = DATA_GROUP_ARR(k)
      If UBound(TEMP_GROUP, 1) > 0 Then 'InsertRowInMatrix
            ReDim ATEMP_VECTOR(1 To UBound(TEMP_GROUP, 1) + 1, 1 To UBound(TEMP_GROUP, 2))
      Else
            ReDim ATEMP_VECTOR(1 To 1, 1 To UBound(TEMP_GROUP, 2))
      End If
      
      For h = 1 To UBound(TEMP_GROUP, 1)
          For l = 1 To UBound(TEMP_GROUP, 2)
            ATEMP_VECTOR(h, l) = TEMP_GROUP(h, l)
          Next l
      Next h
      
      For l = 1 To UBound(BTEMP_VECTOR, 1)
            ATEMP_VECTOR(UBound(TEMP_GROUP, 1) + 1, l) = BTEMP_VECTOR(l)
      Next l
      
      TEMP_GROUP = ATEMP_VECTOR
      DATA_GROUP_ARR(k) = TEMP_GROUP
'------------------------------------------------------------------------------
    Else 'Add Group
'------------------------------------------------------------------------------
        NSIZE = UBound(DATA_GROUP_ARR, 1)
        
        ReDim Preserve DATA_GROUP_ARR(1 To NSIZE + 1)
        ReDim Preserve KEY_GROUP_ARR(1 To NSIZE + 1)
        ReDim TEMP_MATRIX(0 To 0, 1 To UBound(BTEMP_VECTOR, 1))
        
        KEY_GROUP_ARR(NSIZE + 1) = TEMP_VALUE
            
        If UBound(TEMP_MATRIX, 1) > 0 Then 'InsertRowInMatrix
              ReDim ATEMP_VECTOR(1 To UBound(TEMP_MATRIX, 1) + 1, 1 To UBound(TEMP_MATRIX, 2))
        Else
              ReDim ATEMP_VECTOR(1 To 1, 1 To UBound(TEMP_MATRIX, 2))
        End If
        
        For h = 1 To UBound(TEMP_MATRIX, 1)
          For l = 1 To UBound(TEMP_MATRIX, 2)
            ATEMP_VECTOR(h, l) = TEMP_MATRIX(h, l)
          Next l
        Next h
        
        For l = 1 To UBound(BTEMP_VECTOR, 1)
            ATEMP_VECTOR(UBound(TEMP_MATRIX, 1) + 1, l) = BTEMP_VECTOR(l)
        Next l
        
        DATA_GROUP_ARR(NSIZE + 1) = ATEMP_VECTOR
'------------------------------------------------------------------------------
    End If
'------------------------------------------------------------------------------
Next i
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------
    MATRIX_GROUP_FUNC = DATA_GROUP_ARR
'------------------------------------------------------------------------------
Case 1
'------------------------------------------------------------------------------
    MATRIX_GROUP_FUNC = KEY_GROUP_ARR
'------------------------------------------------------------------------------
Case Else
'------------------------------------------------------------------------------
    ReDim TEMP_GROUP(1 To 2)
    TEMP_GROUP(1) = DATA_GROUP_ARR
    TEMP_GROUP(2) = KEY_GROUP_ARR
    
    MATRIX_GROUP_FUNC = TEMP_GROUP
'------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
MATRIX_GROUP_FUNC = Err.number
End Function
