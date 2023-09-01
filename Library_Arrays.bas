Attribute VB_Name = "Library_Arrays"
Option Explicit

'FUNCTION:      Is2D
'==============================================================================
'RETURNS TRUE IF ARRAY IS 2 DIMENSIONAL
'==============================================================================

Function Is2D(ByRef vArray As Variant) As Boolean
    Dim a As Long
    On Error Resume Next
    a = LBound(vArray, 2)
    Is2D = Err = 0
    On Error GoTo 0
End Function

'FUNCTION:      To2D
'==============================================================================
'CONVERTS A 1 DIMENSIONAL ARRAY TO 2 DIMENSIONS
'
'[AsColumn]     FALSE: RESULT STRUCTURED AS ROW
'               TRUE: RESULT STRUCTURED AS COLUMN
'[InPlace]      FALSE: RETURNS NEW ARRAY
'               TRUE: OVERWRITES [vArray]
'==============================================================================

Function To2D( _
    ByRef vArray As Variant, _
    Optional ByVal AsColumn As Boolean = False, _
    Optional ByVal InPlace As Boolean = False _
    ) As Variant

'ITERATORS
Dim a As Long

'ARRAY DIMENSIONS
Dim L1, U1 As Long
L1 = LBound(vArray, 1)
U1 = UBound(vArray, 1)

'FAIL IF ALREADY 2D
If Is2D(vArray) Then
    Debug.Print "Array is already 2D"
    To2D = vArray
    Exit Function
Else

    'STRUCTURE AS COLUMN
    If AsColumn = True Then
        ReDim Result(L1 To U1, L1 To L1)
        For a = L1 To U1
            Result(a, L1) = vArray(a)
        Next a

    'STRUCTURE AS ROW
    Else
        ReDim Result(L1 To L1, L1 To U1)
        For a = L1 To U1
            Result(L1, a) = vArray(a)
        Next a
    End If

End If

'INPLACE
If InPlace = True Then
    vArray = Result
Else
    To2D = Result
End If

End Function

'FUNCTION:      ArrayPrint
'==============================================================================
'PRINTS ARRAY TO IMMEDIATE WINDOW
'
'[Specs]        PRINTS ROW AND COL BOUNDS
'==============================================================================

Function ArrayPrint( _
    ByRef vArray As Variant, _
    Optional ByVal specs As Boolean = False _
    )

'ITERATORS
Dim a, b, c As Long

'ARRAY DIMENSIONS
Dim L1, U1, L2, U2 As Long
L1 = LBound(vArray, 1)
U1 = UBound(vArray, 1)

'VARS
Dim printString As String
Dim PrintArray As Variant:  ReDim PrintArray(L1 To U1)
Dim Seperator As String:    Seperator = ", "

'HANDLE FOR 1D ARRAY
If Is2D(vArray) = False Then

    'CONCAT ARRAY TO PRINT AS ONE ROW
    For a = L1 To U1
        printString = printString + CStr(vArray(a))
        If a <> U1 Then
            printString = printString + Seperator
        End If
    Next a
    
    'SPECS
    If specs Then
        Debug.Print "Rows: " & L1 & " to " & U1
    End If

    'PRINT ARRAY
    Debug.Print printString

'HANDLE FOR 2D ARRAY
Else

    L2 = LBound(vArray, 2)
    U2 = UBound(vArray, 2)
       
    For a = L1 To U1
        For b = L2 To U2
            printString = printString + CStr(vArray(a, b))
            If b <> U2 Then
                printString = printString + Seperator
            End If
        Next b
           
        'ADD STRING TO ARRAY
        PrintArray(a) = printString
        printString = ""
    
    Next a

    'SPACER
    Debug.Print ""
    
    'SPECS
    If specs Then
        Debug.Print "Rows: " & L1 & " to " & U1
        Debug.Print "Cols: " & L2 & " to " & U2
        Debug.Print String(12, "-")
    End If
    
    'PRINT ARRAY
    For c = LBound(PrintArray) To UBound(PrintArray)
        Debug.Print PrintArray(c)
    Next c

End If

End Function

'FUNCTION:      ArrayLen
'==============================================================================
'RETURNS COUNT OF ROWS OR COLUMNS IN ARRAY
'
'[Dimension]    1 = ROWS    (MATCHES UBOUND/LBOUND)
'               2 = COLS
'==============================================================================

Function ArrayLen( _
    ByRef vArray As Variant, _
    Optional ByVal Dimension As Long = 1 _
    ) As Long

If Dimension = 1 Then
    ArrayLen = UBound(vArray, 1) - LBound(vArray, 1) + 1
ElseIf Dimension = 2 Then
    ArrayLen = UBound(vArray, 2) - LBound(vArray, 2) + 1
End If

End Function

'FUNCTION:      ChangeBase
'==============================================================================
'RETURNS ARRAY WITH BASE 1 IF [vArray] BASE = 0 / BASE 1 IF [vArray] BASE = 0
'==============================================================================

Function ChangeBase( _
    ByRef vArray As Variant, _
    Optional ByVal InPlace As Boolean = False _
    ) As Variant

Dim Result As Variant

'ITERATORS
Dim a, b As Long

'ARRAY DIMENSIONS
Dim L1, U1, L2, U2 As Long
L1 = LBound(vArray, 1)
U1 = UBound(vArray, 1)
L2 = LBound(vArray, 2)
U2 = UBound(vArray, 2)


'CHANGE FROM BASE 0 TO BASE 1
If L1 = 0 Then

    ReDim Result(1 To U1 + 1, 1 To U2 + 1)
    For a = L1 To U1
        For b = L2 To U2
            Result(a + 1, b + 1) = vArray(a, b)
        Next b
    Next a

'CHANGE FROM BASE 1 TO BASE 0
ElseIf L1 = 1 Then

    ReDim Result(0 To U1 - 1, 0 To U2 - 1)
    For a = L1 To U1
        For b = L2 To U2
            Result(a - 1, b - 1) = vArray(a, b)
        Next b
    Next a

End If

'INPLACE
If InPlace = True Then
    vArray = Result
Else
    ChangeBase = Result
End If

End Function

'FUNCTION:      Row
'==============================================================================
'RETURNS SINGLE ROW FROM ARRAY
'
'[Row]          ROW TO BE RETURNED
'[As2D]         TRUE: RETURNS 2D ARRAY
'               FALSE: RETURNS 1D ARRAY
'[InPlace]      FALSE: RETURNS NEW ARRAY
'               TRUE: OVERWRITES [vArray]
'==============================================================================

Function Row( _
    ByRef vArray As Variant, _
    ByVal Index As Long, _
    Optional ByVal As2D As Boolean = False, _
    Optional ByVal InPlace As Boolean = False _
    ) As Variant

'ITERATOR
Dim a As Long

'ARRAY DIMENSIONS
Dim L1, U1, L2, U2 As Long
L1 = LBound(vArray, 1)
U1 = UBound(vArray, 1)
L2 = LBound(vArray, 2)
U2 = UBound(vArray, 2)

'BUILD 2D ROW
If As2D = True Then
    ReDim Result(L1 To L1, L2 To U2)
    For a = L2 To U2
        Result(L1, a) = vArray(Index, a)
    Next a

'BUILD 1D ARRAY
Else
    ReDim Result(L2 To U2)
    For a = L2 To U2
        Result(a) = vArray(Index, a)
    End If
End If

'INPLACE
If InPlace = True Then
    vArray = Reslt
Else
    Row = Result
End If

End Function

'FUNCTION:      Col
'==============================================================================
'RETURNS SINGLE COLUMN FROM ARRAY
'
'[Col]          COLUMN TO BE RETURNED
'[As2D]         TRUE: RETURNS 2D ARRAY
'               FALSE: RETURNS 1D ARRAY
'[InPlace]      FALSE: RETURNS NEW ARRAY
'               TRUE: OVERWRITES [vArray]
'==============================================================================

Function Col( _
    ByRef vArray As Variant, _
    ByVal Index As Long, _
    Optional ByVal As2D As Boolean = False, _
    Optional ByVal InPlace As Boolean = False _
    ) As Variant

'ITERATOR
Dim a As Long

'ARRAY DIMENSIONS
Dim L1, U1, L2, U2 As Long
L1 = LBound(vArray, 1)
U1 = UBound(vArray, 1)
L2 = LBound(vArray, 2)
U2 = UBound(vArray, 2)

'BUILD 2D COL
If As2D = True Then
    ReDim Result(L1 To U1, L1 To L1)
    For a = L1 To U1
        Result(a, L1) = vArray(a, Index)
    Next a

'BUILD 1D ARRAY
Else
    ReDim Result(L1 To U1)
    For a = L1 To U1
        Result(a) = vArray(a, Index)
    Next a
End If


'INPLACE
If InPlace = True Then
    vArray = Result
Else
    Col = Result
End If

End Function

'FUNCTION:      SliceRows
'==============================================================================
'RETURNS SLICED ARRAY CONTAINING SPECIFIED ROWS
'
'[Index]        ARRAY OF INDEX NUMBERS OF ROWS TO BE RETURNED
'               CAN ALSO BE PASSED AS INT OR LONG
'[InPlace]      FALSE: RETURNS NEW ARRAY
'               TRUE: OVERWRITES [vArray]
'==============================================================================

Function SliceRows( _
    ByRef vArray As Variant, _
    ByVal Index As Variant, _
    Optional ByVal InPlace As Boolean = False _
    ) As Variant

'ITERATORS
Dim a, b As Long

'ARRAY DIMENSIONS
Dim L1, U1, L2, U2 As Long
L1 = LBound(vArray, 1)
U1 = UBound(vArray, 1)
L2 = LBound(vArray, 2)
U2 = UBound(vArray, 2)

'VARS
Dim RowCount As Long
Dim IndexArray As Variant

'CONVERT INDEX IF PASSED AS INT/LONG
If IsArray(Index) = False Then
    IndexArray = Array(Index)
Else
    IndexArray = Index
End If

'GET SIZE OF RETURN ARRAY
RowCount = UBound(IndexArray) + L1

'BUILD OUTPUT
ReDim Result(L1 To RowCount, L2 To U2)
For a = L1 To RowCount
    For b = L2 To U2
        Result(a, b) = vArray(IndexArray(a - L1), b)
    Next b
Next a

'INPLACE
If InPlace = True Then
    vArray = Result
Else
    SliceRows = Result
End If

End Function

'FUNCTION:      SliceCols
'==============================================================================
'RETURNS SLICED ARRAY CONTAINING SPECIFIED COLUMNS
'
'[Index]        ARRAY OF INDEX NUMBERS OF COLUMNS TO BE RETURNED
'               CAN ALSO BE PASSED AS INT OR LONG
'[InPlace]      FALSE: RETURNS NEW ARRAY
'               TRUE: OVERWRITES [vArray]
'==============================================================================

Function SliceCols( _
    ByRef vArray As Variant, _
    ByVal Index As Variant, _
    Optional ByVal Exclusive As Boolean = False, _
    Optional ByVal InPlace As Boolean = False _
    ) As Variant

'ITERATORS
Dim a, b As Long

'ARRAY DIMENSIONS
Dim L1, U1, L2, U2 As Long
L1 = LBound(vArray, 1)
U1 = UBound(vArray, 1)
L2 = LBound(vArray, 2)
U2 = UBound(vArray, 2)

'VARS
Dim ColCount As Long
Dim IndexArray As Variant

'CONVERT INDEX IF PASSED AS INT/LONG
If IsArray(Index) = False Then
    IndexArray = Array(Index)
Else
    IndexArray = Index
End If

'GET SIZE OF RETURN ARRAY
ColCount = UBound(IndexArray) + L1

'BUILD OUTPUT
ReDim Result(L1 To U1, L2 To ColCount)
For a = L1 To U1
    For b = L2 To ColCount
        Result(a, b) = vArray(a, IndexArray(b - L1))
    Next b
Next a


'INPLACE
If InPlace = True Then
    vArray = Result
Else
    SliceCols = Result
End If

End Function

'FUNCTION:      AppendRows
'==============================================================================
'RETURNS [vArray1] WITH [vArray2] APPENDED AS ROWS
'BASE AGNOSTIC. RESULT WILL MATCH [vArray1] BASE
'
'[InPlace]      FALSE: RETURNS NEW ARRAY
'               TRUE: OVERWRITES [vArray1]
'==============================================================================

Function AppendRows( _
    ByRef vArray1 As Variant, _
    ByRef vArray2 As Variant, _
    Optional ByVal InPlace As Boolean = False _
    ) As Variant

Dim Result As Variant

'ITERATORS
Dim a, b As Long

'ARRAY 1 DIMENSIONS
Dim v1_L1, v1_U1, v1_L2, v1_U2, v1_Rows, v1_Cols As Long
v1_L1 = LBound(vArray1, 1)
v1_U1 = UBound(vArray1, 1)
v1_L2 = LBound(vArray1, 2)
v1_U2 = UBound(vArray1, 2)

v1_Rows = v1_U1 + 1 - v1_L1
v1_Cols = v1_U2 + 1 - v1_L2

'ARRAY 2 DIMENSIONS
Dim v2_L1, v2_U1, v2_L2, v2_U2, v2_Rows, v2_Cols  As Long
v2_L1 = LBound(vArray2, 1)
v2_U1 = UBound(vArray2, 1)
v2_L2 = LBound(vArray2, 2)
v2_U2 = UBound(vArray2, 2)

v2_Rows = v2_U1 + 1 - v2_L1
v2_Cols = v2_U2 + 1 - v2_L2

'EXIT IF ARRAYS ARENT SAME WIDTH
If v1_Cols <> v2_Cols Then
    Debug.Print "AppendRows Failed: Arrays have different number of Columns"
    Debug.Print "vArray1: " & v1_Cols
    Debug.Print "vArray2: " & v2_Cols
    AppendRows = False
    Exit Function
End If

'VAR TO ACCOMODATE DIFFERENT ARRAY BASES
Dim Adjust As Long
Adjust = v1_L1 - v2_L1

'SIZE RESULT ARRAY
Dim Rows As Long:   Rows = v1_U1 + v2_U1 + 1 - v2_L2
ReDim Result(v1_L1 To Rows, v1_L2 To v1_U2)

'ADD [vArray1] TO [Result]
For a = v1_L1 To v1_U1
    For b = v1_L2 To v1_U2
        Result(a, b) = vArray1(a, b)
    Next b
Next a

'ADD [vArray2] TO [Result]
For a = v2_L1 To v2_U1
    For b = v2_L2 To v2_U2
        Result(a + Adjust + v1_Rows, b + Adjust) = vArray2(a, b)
    Next b
Next a

'INPLACE
If InPlace = True Then
    vArray1 = Result
Else
    AppendRows = Result
End If

End Function

'FUNCTION:      AppendCols
'==============================================================================
'RETURNS [vArray1] WITH [vArray2] APPENDED AS COLUMNS
'BASE AGNOSTIC. RESULT WILL MATCH [vArray1] BASE
'
'[InPlace]      FALSE: RETURNS NEW ARRAY
'               TRUE: OVERWRITES [vArray1]
'==============================================================================

Function AppendCols( _
    ByRef vArray1 As Variant, _
    ByRef vArray2 As Variant, _
    Optional ByVal InPlace As Boolean = False _
    ) As Variant

Dim Result As Variant

'ITERATORS
Dim a, b As Long

'ARRAY 1 DIMENSIONS
Dim v1_L1, v1_U1, v1_L2, v1_U2, v1_Rows, v1_Cols As Long
v1_L1 = LBound(vArray1, 1)
v1_U1 = UBound(vArray1, 1)
v1_L2 = LBound(vArray1, 2)
v1_U2 = UBound(vArray1, 2)

v1_Rows = v1_U1 + 1 - v1_L1
v1_Cols = v1_U2 + 1 - v1_L2

'ARRAY 2 DIMENSIONS
Dim v2_L1, v2_U1, v2_L2, v2_U2, v2_Rows, v2_Cols  As Long
v2_L1 = LBound(vArray2, 1)
v2_U1 = UBound(vArray2, 1)
v2_L2 = LBound(vArray2, 2)
v2_U2 = UBound(vArray2, 2)

v2_Rows = v2_U1 + 1 - v2_L1
v2_Cols = v2_U2 + 1 - v2_L2

'EXIT IF ARRAYS ARENT SAME WIDTH
If v1_Rows <> v2_Rows Then
    Debug.Print "AppendRows Failed: Arrays have different number of Rows"
    Debug.Print "vArray1: " & v1_Rows
    Debug.Print "vArray2: " & v2_Rows
    AppendCols = False
    Exit Function
End If

'VAR TO ACCOMODATE DIFFERENT ARRAY BASES
Dim Adjust As Long
Adjust = v1_L1 - v2_L1

'SIZE RESULT ARRAY
Dim Cols As Long:   Cols = v1_U2 + v2_U2 + 1 - v2_L1
ReDim Result(v1_L1 To v1_U1, v1_L2 To Cols)

'ADD [vArray1] TO [Result]
For a = v1_L1 To v1_U1
    For b = v1_L2 To v1_U2
        Result(a, b) = vArray1(a, b)
    Next b
Next a

'ADD [vArray2] TO [Result]
For a = v2_L1 To v2_U1
    For b = v2_L2 To v2_U2
        Result(a + Adjust, b + Adjust + v1_Cols) = vArray2(a, b)
    Next b
Next a

'INPLACE
If InPlace = True Then
    vArray1 = Result
Else
    AppendCols = Result
End If

End Function

'FUNCTION:      AppendFree
'==============================================================================
'RETURNS [vArray1] WITH [vArray2] APPENDED AS EITHER COLS OR ROWS
'IGNORES DIFFERENT ARRAY SIZES LEAVING EMPTY ITEMS AS vbNullString
'SLOWER THAN AppendRows or AppendCols ON LARGE ARRAYS
'BASE AGNOSTIC. RESULT WILL MATCH [vArray1] BASE
'
'[AsCols]       TRUE: APPENDS [vArray2] AS COLUMNS
'               FALSE: APPENDS [vArray2] AS ROWS
'[InPlace]      FALSE: RETURNS NEW ARRAY
'               TRUE: OVERWRITES [vArray1]
'==============================================================================

Function AppendFree( _
    ByRef vArray1 As Variant, _
    ByRef vArray2 As Variant, _
    Optional ByVal AsCols As Boolean = False, _
    Optional ByVal InPlace As Boolean = False _
    ) As Variant

Dim Result As Variant

'ITERATORS
Dim a, b As Long

'ARRAY 1 DIMENSIONS
Dim v1_L1, v1_U1, v1_L2, v1_U2, v1_Rows, v1_Cols As Long
v1_L1 = LBound(vArray1, 1)
v1_U1 = UBound(vArray1, 1)
v1_L2 = LBound(vArray1, 2)
v1_U2 = UBound(vArray1, 2)

v1_Rows = v1_U1 + 1 - v1_L1
v1_Cols = v1_U2 + 1 - v1_L2

'ARRAY 2 DIMENSIONS
Dim v2_L1, v2_U1, v2_L2, v2_U2, v2_Rows, v2_Cols  As Long
v2_L1 = LBound(vArray2, 1)
v2_U1 = UBound(vArray2, 1)
v2_L2 = LBound(vArray2, 2)
v2_U2 = UBound(vArray2, 2)

v2_Rows = v2_U1 + 1 - v2_L1
v2_Cols = v2_U2 + 1 - v2_L2

'VAR TO ACCOMODATE DIFFERENT ARRAY BASES

Dim Base As Long:   Base = v1_L1
Dim R_Rows As Long
Dim R_Cols As Long

Dim Offset As Long: Offset = v1_L1 - v2_L1

'APPEND AS COLUMNS
If AsCols = True Then

    'SIZE [Result]
    R_Rows = Application.WorksheetFunction.Max(v1_Rows, v2_Rows) - 1 + Base
    R_Cols = v1_Cols + v2_Cols - 1 + Base
    ReDim Result(Base To R_Rows, Base To R_Cols)

    'ADD [vArray1] to [Result]
    For a = v1_L1 To v1_U1
        For b = v1_L2 To v1_U2
            Result(a, b) = vArray1(a, b)
        Next b
    Next a

    'ADD [vArray2] to [Result]
    For a = v2_L1 To v2_U1
        For b = v2_L2 To v2_U2
            Result(a + Offset, b + Offset + v1_Cols) = vArray2(a, b)
        Next b
    Next a

'APPEND AS ROWS
ElseIf AsCols = False Then

    'SIZE [Result]
    R_Rows = v1_Rows + v2_Rows - 1 + Base
    R_Cols = Application.WorksheetFunction.Max(v1_Cols, v2_Cols) - 1 + Base
    ReDim Result(Base To R_Rows, Base To R_Cols)

    'ADD [vArray1] to [Result]
    For a = v1_L1 To v1_U1
        For b = v1_L2 To v1_U2
            Result(a, b) = vArray1(a, b)
        Next b
    Next a
    
    'ADD [vArray2] to [Result]
    For a = v2_L1 To v2_U1
        For b = v2_L2 To v2_U2
            Result(a + Offset + v1_Rows, b + Offset) = vArray2(a, b)
        Next b
    Next a

End If

'INPLACE
If InPlace = True Then
    vArray1 = Result
Else
    AppendFree = Result
End If

End Function

'FUNCTION:      Filter
'==============================================================================
'RETURNS ARRAY WHERE VALUE IN [FilterCol] SATISFIES THE SELECTED OPERATOR FOR
'[FilterValue]
'
'[FilterCol]    COLUMN TO BE COMPARED FOR FILTER
'[Operation]    ACCEPTS =, <>, >, <, >=, <=, Like
'[FilterValue]  VALUE TO BE COMPARED AGAINST
'[HasHeader]    TRUE: INCLUDES FIRST ROW IN [Result] AND EXCLUDES FROM FILTER
'               OPERATION
'[InPlace]      FALSE: RETURNS NEW ARRAY
'               TRUE: OVERWRITES [vArray]
'==============================================================================

Function Filter( _
    ByRef vArray As Variant, _
    ByVal FilterCol As Long, _
    ByVal Operation As String, _
    ByVal FilterValue As Variant, _
    Optional HasHeader As Boolean = False, _
    Optional ByVal InPlace As Variant = False _
    ) As Variant

'ITERATORS
Dim a, b As Long

'ARRAY DIMENSIONS
Dim L1, U1, L2, U2 As Long
L1 = LBound(vArray, 1)
U1 = UBound(vArray, 1)
L2 = LBound(vArray, 2)
U2 = UBound(vArray, 2)

'MATCH DICTIONARY
Dim MatchDict As Object
Set MatchDict = CreateObject("Scripting.Dictionary")

'TEMP ARRAY CONTAINING INDEX OF MATCHING ROWS
Dim Matches As Variant

'ADD HEADER ROW TO RESULTS
Dim H As Long: H = 0
If HasHeader Then
    H = 1
    MatchDict.Add L1, L1
End If

'FILTER ROWS FOR SELECTED OPERATION
Select Case Operation

    Case "="
        For a = L1 + H To U1
            If vArray(a, FilterCol) = FilterValue Then
                MatchDict.Add a, a
            End If
        Next a
    
    Case "<>"
        For a = L1 + H To U1
            If vArray(a, FilterCol) <> FilterValue Then
                MatchDict.Add a, a
            End If
        Next a

    Case ">"
        For a = L1 + H To U1
            If vArray(a, FilterCol) > FilterValue Then
                MatchDict.Add a, a
            End If
        Next a
    
    Case "<"
        For a = L1 + H To U1
            If vArray(a, FilterCol) < FilterValue Then
                MatchDict.Add a, a
            End If
        Next a

    Case ">="
        For a = L1 + H To U1
            If vArray(a, FilterCol) >= FilterValue Then
                MatchDict.Add a, a
            End If
        Next a
    
    Case "<="
        For a = L1 + H To U1
            If vArray(a, FilterCol) <= FilterValue Then
                MatchDict.Add a, a
            End If
        Next a

    Case "Like"
        For a = L1 + H To U1
            If vArray(a, FilterCol) Like FilterValue Then
                MatchDict.Add a, a
            End If
        Next a

End Select

'IF RESULTS FOUND
If MatchDict.Count > 0 Then

    'SIZE RESULTS ARRAY
    Matches = MatchDict.Keys
    ReDim Result(L1 To MatchDict.Count - 1 + L1, L2 To U2)

    'ADD MATCHED ROWS TO RESULT
    For a = LBound(Matches) To UBound(Matches)
        For b = L2 To U2
            Result(a + L1, b) = vArray(Matches(a), b)
        Next b
    Next a

'IF NO RESULTS FOUND
Else
    Debug.Print "Filter Found 0 Matching Rows"
    Filter = False
    Exit Function
End If

'INPLACE
If InPlace = True Then
    vArray = Result
Else
    Filter = Result
End If

End Function

'FUNCTION:      UniqueRows
'==============================================================================
'RETURNS ARRAY CONTAINING ONLY ROWS WITH UNIQUE VALUES IN SPECIFIED COLUMNS
'
'[Index]        COMPARES ALL COLUMNS IF MISSING
'               ARRAY OF INDEX NUMBERS OF COLUMNS TO BE COMPARED FOR UNIQUENESS
'               CAN ALSO BE PASSED AS INT OR LONG FOR SINGLE COLUMN
'[InPlace]      FALSE: RETURNS NEW ARRAY
'               TRUE: OVERWRITES [vArray]
'==============================================================================

Function UniqueRows( _
    ByRef vArray As Variant, _
    Optional ByVal Index As Variant, _
    Optional ByVal InPlace As Boolean = False _
    ) As Variant

'ITERATORS
Dim a, b As Long

'ARRAY DIMENSIONS
Dim L1, U1, L2, U2 As Long
L1 = LBound(vArray, 1)
U1 = UBound(vArray, 1)
L2 = LBound(vArray, 2)
U2 = UBound(vArray, 2)

'VARS
Dim IndexArray As Variant
Dim UniqueArray As Variant
Dim CompareString As String

'MATCH DICTIONARY
Dim UniqueDict As Object
Set UniqueDict = CreateObject("Scripting.Dictionary")

'SET [IndexArray] TO ALL COLS IF MISSING
If IsMissing(Index) Then
    ReDim IndexArray(0 To U2 - L1)
    For a = L2 To U2
        IndexArray(a - L1) = a
    Next a
Else
    'CONVERT INDEX IF PASSED AS INT/LONG
    If IsArray(Index) = False Then
        IndexArray = Array(Index)
    Else
        IndexArray = Index
    End If
End If

'FIND UNIQUE ROWS BY ADDING VALUES TO DICTIONARY AS KEYS
'DICT CANNOT HAVE DUPLICATE KEYS
For a = L1 To U1
    
    'CONCAT ROW AS SINGLE STRING
    For b = LBound(IndexArray) To UBound(IndexArray)
        CompareString = CompareString + CStr(vArray(a, IndexArray(b)))
    Next b

    'ADD STRING TO DICT
    On Error Resume Next
        UniqueDict.Add CompareString, a
    On Error GoTo 0
    'RESET
    CompareString = vbNullString

Next a

'GET ARRAY CONTAINING INDEX OF UNIQUE ROWS
UniqueArray = UniqueDict.Items

'BUILD RESULT
ReDim Result(L1 To UBound(UniqueArray) + L1, L2 To U2)
For a = LBound(UniqueArray) To UBound(UniqueArray)
    For b = L2 To U2
        Result(a + L1, b) = vArray(UniqueArray(a), b)
    Next b
Next a

'INPLACE
If InPlace = True Then
    vArray = Result
Else
    UniqueRows = Result
End If

End Function

Function ArrayContains( _
    ByVal vArray As Variant, _
    ByRef Search As Variant _
    ) As Variant

'ITERATORS
Dim a, b, Location1, Location2 As Long

Dim Result As Variant
Result = False

'ARRAY DIMENSIONS
Dim L1, U1, L2, U2 As Long
L1 = LBound(vArray, 1)
U1 = UBound(vArray, 1)

'SEARCH 1D ARRAY
If Is2D(vArray) = False Then

    For a = L1 To U1
        If vArray(a) = Search Then
            Result = a
            Exit For
        End If
    Next a

'SEARCH 2D ARRAY
Else

    L2 = LBound(vArray, 2)
    U2 = UBound(vArray, 2)

    For a = L1 To U1
        For b = L2 To U2
            If vArray(a, b) = Search Then
                Result = Array(a, b)
                Exit For
            End If
        Next b
    Next a
     
End If

ArrayContains = Result

End Function

Function NumRange(ByVal First As Long, ByVal Last As Long) As Variant
Dim a, Result() As Long
ReDim Result(0 To Last - First)
For a = First To Last
    Result(a - First) = a
Next a
NumRange = Result
End Function

Function ColLetter(number As Long) As String
    ColLetter = Split(Cells(1, number).Address, "$")(1)
End Function
















