Attribute VB_Name = "Module1"
' Copyrights © 2012 - finaquant.com (Tunc A. Kutukcuoglu)
' Website: http://finaquant.com/
'
' Central download page (check for updates):
' http://finaquant.com/download
'
' Copyright disclaimer:
' You can use this software (i.e. functions) for non-commercial and commercial
' purposes provided that you leave the credits for the publisher "finaquant"
' in place for each function included.
'
' Responsibility of use belongs 100% to user:
' The author and publisher of this software assumes no responsibility for errors or omissions,
' or for damages resulting from the use of this code. In no event shall the
' author be liable for any loss of profit or any other commercial damage caused or
' alleged to have been caused directly or indirectly by this material.

' FQ prefix: finaquant
' declare data type for all variables
Option Explicit
Option Base 1

' Global constants

' set FQ_DEBUG = True for debug modus
' In debug modus:
' MessageBox error notifications are disabled
' Error notifications are directly printed to immediate window
' Public Const FQ_DEBUG As Boolean = True ' debug mode

' Version history
' V1.1      First original version,  June 2012
' V1.2      July 2012
'           Matrix determinant function is added (Module1)
'           Procedures for test data generation are added to Module3

Public Const FQ_ErrorNum As Byte = 17   ' standard error: can't perform requested operation

Enum MatrixDirection
    nRowByRow = 1           ' row-wise
    nColByCol = 2           ' column-wise
    nAllElements = 3
End Enum

Enum MatrixAlignment
    nHorizontal = 1
    nVertical = 2
End Enum
 
Enum SortOption
    nAscending = 1
    nDescending = 2
End Enum
'******************************************************************
' Customized message box
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_MessageBox(msg As String)
Attribute FQ_MessageBox.VB_Description = "Customized message box; prints to immediate window if the global parameter FQ_DEBUG is true."
Attribute FQ_MessageBox.VB_ProcData.VB_Invoke_Func = " \n14"
If FQ_DEBUG Then
    Debug.Print (msg)
Else
    FQ_MessageBox = MsgBox(Prompt:=msg, Title:="Elementary Matrix Functions - finaquant.com")
End If
End Function
'******************************************************************
' Returns the  number of dimensions in an array
' returns 0 for scalar arguments
'******************************************************************
Function FQ_ArrayDimension(Arr As Variant) As Integer
Attribute FQ_ArrayDimension.VB_Description = "Check if array; returns True if the argument is an array; otherwise False."
Attribute FQ_ArrayDimension.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ndim As Integer, ub As Integer
ndim = 0
On Error Resume Next
Do
    ndim = ndim + 1
    ub = UBound(Arr, ndim)
Loop Until Err.Number <> 0
FQ_ArrayDimension = ndim - 1
End Function
'******************************************************************
' Returns True if the argument is an array; otherwise False
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_CheckIfArray(Arr As Variant) As Boolean
If IsArray(Arr) Then
    FQ_CheckIfArray = True
Else
    FQ_CheckIfArray = False
End If
End Function
'******************************************************************
' Checks if empty array; returns True if array is empty
' - error if input argument Arr is not an array
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_CheckIfEmptyArray(Arr As Variant) As Boolean
Attribute FQ_CheckIfEmptyArray.VB_Description = "Check if empty array; returns True if array is empty."
Attribute FQ_CheckIfEmptyArray.VB_ProcData.VB_Invoke_Func = " \n14"
If Not IsArray(Arr) Then
    FQ_MessageBox ("Error in FQ_CheckIfEmptyArray: Input argument Arr is not an array!")
End If
' check if empty
On Error GoTo ErrorHandler
If UBound(Arr) = 0 Then
    FQ_CheckIfEmptyArray = True
Else
    FQ_CheckIfEmptyArray = False
End If
Exit Function
ErrorHandler:
    FQ_CheckIfEmptyArray = True
End Function
'******************************************************************
' Returns True if the argument is an matrix; otherwise False
' A matrix must be:
' - 2-dimensional array of data type Double
' - lowest element index must be 1 (lower bound)
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_CheckIfMatrix(Arr As Variant) As Boolean
Attribute FQ_CheckIfMatrix.VB_Description = "Check if matrix; returns True if the argument is an matrix; otherwise False."
Attribute FQ_CheckIfMatrix.VB_ProcData.VB_Invoke_Func = " \n14"
If IsArray(Arr) And (FQ_ArrayDimension(Arr) = 2) And _
    (VarType(Arr) = vbArray + vbDouble) Then
    
    If LBound(Arr, 1) = 1 And LBound(Arr, 2) = 1 Then
        FQ_CheckIfMatrix = True
    Else
        FQ_CheckIfMatrix = False
    End If
Else
    FQ_CheckIfMatrix = False
End If
End Function
'******************************************************************
' Returns True if the argument is a vector; otherwise False
' - A vector must be:
' - 1-dimensional array of data type Double or Long
' - lowest element index must be 1 (lower bound)
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_CheckIfVector(Arr As Variant) As Boolean
Attribute FQ_CheckIfVector.VB_Description = "Check if vector; returns True if the argument is a vector; otherwise False."
Attribute FQ_CheckIfVector.VB_ProcData.VB_Invoke_Func = " \n14"
If IsArray(Arr) And (FQ_ArrayDimension(Arr) = 1) And _
    ((VarType(Arr) = vbArray + vbDouble) Or _
    (VarType(Arr) = vbArray + vbLong)) Then
    
    If LBound(Arr, 1) = 1 Then
        FQ_CheckIfVector = True
    Else
        FQ_CheckIfVector = False
    End If
Else
    FQ_CheckIfVector = False
End If
End Function
'******************************************************************
' Converts a variant array with numeric elements into a matrix
' - 1-dimensional variant array is converted into a 1xN horizontal matrix
' - example 2x3 variant array: [{1,2,3; 4,5,6}]
' - error if dimension of variant array is not in set {0,1,2}
' - error if there is a non-numeric element in variant array
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_var_to_matrix(Arr As Variant) As Double()
Attribute FQ_var_to_matrix.VB_Description = "Converts a variant array with numeric elements into a matrix."
Attribute FQ_var_to_matrix.VB_ProcData.VB_Invoke_Func = " \n14"
'check dimension of Arr
Dim ndim As Integer, nrow As Integer, ncol As Integer, nlen As Integer
Dim i As Integer, j As Integer
Dim Mres() As Double

ndim = FQ_ArrayDimension(Arr)

Select Case ndim
    Case 0
        If IsNumeric(Arr) Then
            ReDim Mres(1 To 1, 1 To 1)
            Mres(1, 1) = Arr
        Else
            FQ_MessageBox ("Error in FQ_var_to_matrix: All element values must be numeric (ndim 0)!")
            Err.Raise (FQ_ErrorNum)
            Exit Function
        End If
        
    Case 1
        nlen = UBound(Arr, 1) - LBound(Arr, 1) + 1
        ReDim Mres(1 To 1, 1 To nlen)
        
        For i = 1 To nlen
            If IsNumeric(Arr(i + LBound(Arr, 1) - 1)) Then
                Mres(1, i) = Arr(i + LBound(Arr, 1) - 1)
            Else
                FQ_MessageBox ("Error in FQ_var_to_matrix: All element values must be numeric (ndim 1)!")
                Err.Raise (FQ_ErrorNum)
                Exit Function
            End If
        Next i
        
    Case 2
        nrow = UBound(Arr, 1) - LBound(Arr, 1) + 1
        ncol = UBound(Arr, 2) - LBound(Arr, 2) + 1
        ReDim Mres(1 To nrow, 1 To ncol)
        
        For i = 1 To nrow
            For j = 1 To ncol
                If IsNumeric(Arr(i + LBound(Arr, 1) - 1, j + LBound(Arr, 2) - 1)) Then
                    Mres(i, j) = Arr(i + LBound(Arr, 1) - 1, j + LBound(Arr, 2) - 1)
                Else
                    FQ_MessageBox ("Error in FQ_var_to_matrix: All element values must be numeric (ndim 2)!")
                    Err.Raise (FQ_ErrorNum)
                    Exit Function
                End If
            Next j
        Next i
        
    Case Else
        FQ_MessageBox ("Error in FQ_var_to_matrix: Improper array dimension (ndim > 2)")
        Err.Raise (FQ_ErrorNum)
        Exit Function
End Select
FQ_var_to_matrix = Mres
End Function
'******************************************************************
' Converts a variant array with numeric elements into a vector
' - example variant array: [{1,2,3,4,5,6}]
' - error if dimension of variant array is not 0 or 1
' - error if there is a non-numeric element in variant array
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_var_to_vector(Arr As Variant) As Double()
Attribute FQ_var_to_vector.VB_Description = "Converts a variant array with numeric elements into a vector."
Attribute FQ_var_to_vector.VB_ProcData.VB_Invoke_Func = " \n14"
'check dimension of Arr
Dim ndim As Integer, nlen As Integer
Dim i As Integer
Dim Vres() As Double

ndim = FQ_ArrayDimension(Arr)

Select Case ndim
    Case 0
        If IsNumeric(Arr) Then
            ReDim Vres(1 To 1)
            Vres(1) = Arr
        Else
            FQ_MessageBox ("Error in FQ_var_to_vector: All element values must be numeric (ndim 0)!")
            Err.Raise (FQ_ErrorNum)
            Exit Function
        End If
        
    Case 1
        nlen = UBound(Arr, 1) - LBound(Arr, 1) + 1
        ReDim Vres(1 To nlen)
        
        For i = 1 To nlen
            If IsNumeric(Arr(i + LBound(Arr, 1) - 1)) Then
                Vres(i) = Arr(i + LBound(Arr, 1) - 1)
            Else
                FQ_MessageBox ("Error in FQ_var_to_vector: All element values must be numeric (ndim 1)!")
                Err.Raise (FQ_ErrorNum)
                Exit Function
            End If
        Next i
        
    Case Else
        FQ_MessageBox ("Error in FQ_var_to_vector: Improper array dimension (ndim > 1)")
        Err.Raise (FQ_ErrorNum)
        Exit Function
End Select
FQ_var_to_vector = Vres
End Function
'******************************************************************
' Converts a matrix into a printable formatted string
' for displaying matrices to user
' - returns "ERROR" if the input argument is not a matrix
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_format(M As Variant) As String
Attribute FQ_matrix_format.VB_Description = "Converts a matrix into a printable formatted string."
Attribute FQ_matrix_format.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i As Integer, j As Integer
Dim SepStr As String        ' element seperator
Dim LineStr As String       ' line seperator
Dim SpaceStr As String      ' spaces
Dim NumberStr As String     ' formatted number
Dim MaxNumberLength As Byte

' SepStr = Chr(9)            ' horizontal tab \t
SepStr = " "
LineStr = Chr(10)           ' line feed \n
MaxNumberLength = 15

SpaceStr = String(MaxNumberLength, " ")  ' Space characters for fixed length number display

' check if M is a matrix
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_matrix_format: Input argument must be a matrix!")
    FQ_matrix_format = "ERROR"
    Exit Function
End If

' initiate return value
FQ_matrix_format = LineStr

For i = 1 To UBound(M, 1)
    For j = 1 To UBound(M, 2)
        ' see vba format() function in excel help for formatting numbers
        NumberStr = Right(Format(M(i, j), "###########0.00"), MaxNumberLength)
        FQ_matrix_format = FQ_matrix_format & SepStr & Left(SpaceStr, MaxNumberLength - Len(NumberStr)) & NumberStr
    Next j
    FQ_matrix_format = FQ_matrix_format & LineStr
Next i
End Function
'******************************************************************
' Converts a vector into a printable formatted string
' for displaying vectors to user
' - returns "ERROR" if the input argument is not a matrix
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_format(V As Variant) As String
Attribute FQ_vector_format.VB_Description = "Converts a vector into a printable formatted string."
Attribute FQ_vector_format.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i As Integer
Dim SepStr As String        ' element seperator
Dim SpaceStr As String      ' spaces
Dim NumberStr As String     ' formatted number
Dim MaxNumberLength As Byte
' SepStr = Chr(9)           ' horizontal tab \t
SepStr = " "
MaxNumberLength = 15
SpaceStr = String(MaxNumberLength, " ")  ' Space characters for fixed length number display

' check if M is a matrix
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_vector_format: Input argument must be a vector!")
    FQ_vector_format = "ERROR"
    Exit Function
End If

' initiate return value
FQ_vector_format = ""

For i = 1 To UBound(V, 1)
    ' see vba format() function in excel help for formatting numbers
    NumberStr = Right(Format(V(i), "###########0.00"), MaxNumberLength)
    FQ_vector_format = FQ_vector_format & SepStr & Left(SpaceStr, MaxNumberLength - Len(NumberStr)) & NumberStr
Next i
End Function
'******************************************************************
' Converts a 1-dimensional (1xN or Nx1) matrix to a vector
' - returns the same vector if the input M is a vector
' - error if M is not a 1-dimensional matrix
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_1DimMatrix_to_Vector(M() As Double) As Double()
Attribute FQ_1DimMatrix_to_Vector.VB_Description = "Converts a 1-dimensional (1xN or Nx1) matrix to a vector."
Attribute FQ_1DimMatrix_to_Vector.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nrow As Long, ncol As Long, Vres() As Double
Dim i As Integer, j As Integer

If FQ_CheckIfVector(M) Then
    FQ_1DimMatrix_to_Vector = M
    Exit Function
End If
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_1DimMatrix_to_Vector: Input argument M is not a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' get row and column sizes
nrow = UBound(M, 1)
ncol = UBound(M, 2)
If nrow <> 1 And ncol <> 1 Then
    FQ_MessageBox ("Error in FQ_1DimMatrix_to_Vector: Input argument M is not a 1-dimensional matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
ReDim Vres(1 To nrow * ncol)
For i = 1 To nrow
    For j = 1 To ncol
        Vres((i - 1) * ncol + j) = M(i, j)
    Next j
Next i
FQ_1DimMatrix_to_Vector = Vres
End Function
'******************************************************************
' Converts Vector do 1-dimensional matrix, either
' vertical or horizontal depending on matrix alignment argument
' - returns the same matrix (aligned as requested) if input V is a 1-dim matrix
' - error if M is not a vector or 1-dim matrix
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_Vector_to_1DimMatrix(V() As Double, malign As MatrixAlignment) As Double()
Attribute FQ_Vector_to_1DimMatrix.VB_Description = "Converts Vector to a 1-dimensional matrix, either vertical or horizontal depending on matrix alignment argument."
Attribute FQ_Vector_to_1DimMatrix.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nlen As Long, Mres() As Double
Dim i As Integer, j As Integer
If FQ_CheckIfMatrix(V) Then
    If UBound(V, 1) = 1 Or UBound(V, 2) = 1 Then ' 1-dim matrix
        If (UBound(V, 1) > 1 And malign = nHorizontal) Or _
            (UBound(V, 2) > 1 And malign = nVertical) Then
            FQ_Vector_to_1DimMatrix = FQ_matrix_transpose(V)
        Else
            FQ_Vector_to_1DimMatrix = V
        End If
        Exit Function
    Else
        FQ_MessageBox ("Error in FQ_Vector_to_1DimMatrix: Input argument V is a 2-dimensional matrix!")
        Err.Raise (FQ_ErrorNum)
        Exit Function
    End If
End If
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_Vector_to_1DimMatrix: Input argument V is a not a vector or 1-dimensional matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
Select Case malign
    Case MatrixAlignment.nHorizontal
        nlen = UBound(V, 1)
        ReDim Mres(1 To 1, 1 To nlen)
        For i = 1 To nlen
            Mres(1, i) = V(i)
        Next i
    Case MatrixAlignment.nVertical
        nlen = UBound(V, 1)
        ReDim Mres(1 To nlen, 1 To 1)
        For i = 1 To nlen
            Mres(i, 1) = V(i)
        Next i
    Case Else
        FQ_MessageBox ("Error in FQ_Vector_to_1DimMatrix: Invalid matrix alignment option!")
        Err.Raise (FQ_ErrorNum)
        Exit Function
End Select
FQ_Vector_to_1DimMatrix = Mres
End Function
'******************************************************************
' Transpose matrix
' M2 = transpose(M1)  such that M1(j, i) = M2(i, j)
' - error if M is not a matrix
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_transpose(M() As Double) As Double()
Attribute FQ_matrix_transpose.VB_Description = "Transpose matrix; M2 = transpose(M1)."
Attribute FQ_matrix_transpose.VB_ProcData.VB_Invoke_Func = " \n14"
' Check if matrix
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_matrix_transpose: Input argument M must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' worksheet function returns variant array
FQ_matrix_transpose = FQ_var_to_matrix(Application.WorksheetFunction.Transpose(M))
End Function
'******************************************************************
' Inverse matrix: Y = inv(M)
' - error if M is not a square matrix (ncol = nrow)
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_inverse(M() As Double) As Double()
Attribute FQ_matrix_inverse.VB_Description = "Inverse matrix: Y = inv(M)."
Attribute FQ_matrix_inverse.VB_ProcData.VB_Invoke_Func = " \n14"
On Error GoTo EH1
' Check if square matrix
If Not (FQ_CheckIfMatrix(M) And UBound(M, 1) = UBound(M, 2)) Then
    FQ_MessageBox ("Error in FQ_matrix_inverse: Input argument M must be a square matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' worksheet function returns variant array
FQ_matrix_inverse = FQ_var_to_matrix(Application.WorksheetFunction.MInverse(M))
Exit Function
EH1:
FQ_MessageBox ("Error in FQ_matrix_inverse: " & Err.Number & " - " & Err.Description)
Err.Raise (Err.Number)
End Function
'******************************************************************
' Writes the values of a range of cells into a 2-dimensional
' NxM variant array (1 to N, 1 to M) or
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_range_to_variant(Rn As Range) As Variant
Attribute FQ_range_to_variant.VB_Description = "Writes the numeric values of a worksheet range into a 2-dimensional NxM variant array."
Attribute FQ_range_to_variant.VB_ProcData.VB_Invoke_Func = " \n14"
Dim trn As Range
If Rn.Areas.Count > 1 Then
    FQ_MessageBox ("Error in FQ_range_to_variant: Selected range must be a single contiguous area; multiple range selection is not permitted!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
Set trn = Intersect(Rn.Parent.UsedRange, Rn)
FQ_range_to_variant = trn.Value
End Function
'******************************************************************
' Writes the values of a 2-dimensional variant array into a range in excel,
' starting from the upper left corner of the range (Cells(1,1))
' - error if array is empty or not initialized, or not 2-dim
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Sub FQ_variant_to_range(Arr As Variant, Rn As Range)
Attribute FQ_variant_to_range.VB_Description = "Writes the values of a 2-dimensional variant array into a range in excel, starting from the upper left corner of the range (Cells(1,1))."
Attribute FQ_variant_to_range.VB_ProcData.VB_Invoke_Func = " \n14"
Dim RangeData As Range, nrow As Long, ncol As Long
On Error GoTo ErrorHandler
nrow = UBound(Arr, 1)
ncol = UBound(Arr, 2)
Set RangeData = Range(Rn.Cells(1, 1), Rn.Cells(nrow, ncol))
RangeData.Value = Arr
Exit Sub
ErrorHandler:
FQ_MessageBox ("Error in FQ_variant_to_range: " & Err.Number & " - " & Err.Description)
Err.Raise (Err.Number)
End Sub
'******************************************************************
' Writes the values of a matrix (2-dim double) into a range in excel,
' starting from the upper left corner of the range (Cells(1,1))
' - error if M is not matrix
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Sub FQ_matrix_to_range(M() As Double, Rn As Range)
Attribute FQ_matrix_to_range.VB_Description = " Writes the values of a matrix (2-dim double) into a worksheet range in excel, starting from the upper left corner of the range."
Attribute FQ_matrix_to_range.VB_ProcData.VB_Invoke_Func = " \n14"
' Check if matrix
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_matrix_to_range: Input argument M must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Sub
End If
Call FQ_variant_to_range(M, Rn)
End Sub
'******************************************************************
' Writes the values of a vector (1-dim double) into a range in excel,
' starting from the upper left corner of the range (Cells(1,1))
' - error if V is not vector
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Sub FQ_vector_to_range(V() As Double, Rn As Range, Direction As MatrixAlignment)
Attribute FQ_vector_to_range.VB_Description = "Writes the values of a vector (1-dim double) into a range in excel, starting from the upper left corner of the range."
Attribute FQ_vector_to_range.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i As Long, M() As Double, vlen As Long, nrow As Long, ncol As Long
Dim j As Long
' Check if vector
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_vector_to_range: Input argument V must be a vector!")
    Exit Sub
End If
vlen = UBound(V, 1)
Select Case Direction
    Case nHorizontal
        nrow = 1
        ncol = vlen
    Case nVertical
        nrow = vlen
        ncol = 1
    Case Else
        FQ_MessageBox ("Error in FQ_vector_to_range: Invalid Direction option!")
        Exit Sub
End Select
ReDim M(1 To nrow, 1 To ncol)
For i = 1 To nrow
    For j = 1 To ncol
        M(i, j) = V(ncol * (i - 1) + j)
    Next j
Next i
Call FQ_variant_to_range(M, Rn)
End Sub
'******************************************************************
' Reads a numeric range and converts it into a matrix with the same
' row and column size
' - error if any range value is not numeric; all cells in the range
' must be numeric and non-empty
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_range_to_matrix(R As Range) As Double()
Attribute FQ_range_to_matrix.VB_Description = "Reads a numeric range and converts it into a matrix with the same row and column size."
Attribute FQ_range_to_matrix.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i As Long, j As Long, nrow As Long, ncol As Long
Dim M() As Double
Dim MyCell As Range
' get size of the range
nrow = R.rows.CountLarge
ncol = R.CountLarge / nrow
ReDim M(1 To nrow, 1 To ncol)

' read range cell by cell
For i = 1 To nrow
    For j = 1 To ncol
        Set MyCell = R.Cells(i, j)
        If IsEmpty(MyCell) Then
            FQ_MessageBox ("Error in FQ_range_to_matrix: Empty cell (" & i & ", " & j & ")!")
            Err.Raise (FQ_ErrorNum)
            Exit Function
        End If
        If Not IsNumeric(MyCell) Then
            FQ_MessageBox ("Error in FQ_range_to_matrix: Non-numeric cell (" & i & ", " & j & ")!")
            Err.Raise (FQ_ErrorNum)
            Exit Function
        End If
        M(i, j) = MyCell.Value
    Next j
Next i
FQ_range_to_matrix = M
End Function
'******************************************************************
' Reads a numeric range row by row and writes the values into a vector
' - error if any range value is not numeric; all cells in the range
' must be numeric and non-empty
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_range_to_vector(R As Range) As Double()
Attribute FQ_range_to_vector.VB_Description = "Reads a numeric range row by row and writes the values into a vector."
Attribute FQ_range_to_vector.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i As Long, j As Long, nrow As Long, ncol As Long
Dim V() As Double, ElementCount As Long
Dim MyCell As Range
' get size of the range
nrow = R.rows.CountLarge
ncol = R.CountLarge / nrow

ElementCount = nrow * ncol
ReDim V(1 To ElementCount)

' read range cell by cell
For i = 1 To nrow
    For j = 1 To ncol
        Set MyCell = R.Cells(i, j)
        If IsEmpty(MyCell) Then
            FQ_MessageBox ("Error in FQ_range_to_vector: Empty cell (" & i & ", " & j & ")!")
            Err.Raise (FQ_ErrorNum)
            Exit Function
        End If
        If Not IsNumeric(MyCell) Then
            FQ_MessageBox ("Error in FQ_range_to_vector: Non-numeric cell (" & i & ", " & j & ")!")
            Err.Raise (FQ_ErrorNum)
            Exit Function
        End If
        V((i - 1) * ncol + j) = MyCell.Value
    Next j
Next i
FQ_range_to_vector = V
End Function
'******************************************************************
' Writes the values of a vector into a matrix either row by row, or
' column by column.
' - matrix row/column size is given by parameters nrow and ncol
' - vector elements are recycled with mod() function, if matrix have
' more elements than the vector
' - error if V is not a vector
' - error of nrow < 1 or ncol < 1
' - error if invalid fill option
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_to_matrix(V() As Double, FillOption, nrow As Long, ncol As Long) As Double()
Attribute FQ_vector_to_matrix.VB_Description = "Writes the values of a vector into a matrix either row by row, or column by column."
Attribute FQ_vector_to_matrix.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ElementCountVector As Long, ElementCountMatrix As Long, i As Long, j As Long, ctr As Long
Dim Mres() As Double, iLim As Long, jLim As Long
' Check if vector
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_vector_to_matrix: Input argument V must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' check matrix row/col size
If nrow < 1 Or ncol < 1 Then
    FQ_MessageBox ("Error in FQ_vector_to_matrix: nrow and ncol must be positive integers!")
    Err.Raise (FQ_ErrorNum)
End If
' check fill option
If FillOption <> nRowByRow And FillOption <> nColByCol Then
    FQ_MessageBox ("Error in FQ_vector_to_matrix: Improper fill option!")
    Err.Raise (FQ_ErrorNum)
End If

ReDim Mres(1 To nrow, 1 To ncol)

ElementCountVector = UBound(V, 1)
ElementCountMatrix = nrow * ncol

If FillOption = nRowByRow Then
    iLim = nrow
    jLim = ncol
Else
    iLim = ncol
    jLim = nrow
End If

For i = 1 To iLim
    For j = 1 To jLim
        ctr = jLim * (i - 1) + j
        If (ctr Mod ElementCountVector) <> 0 Then
            If FillOption = nRowByRow Then
                Mres(i, j) = V(ctr Mod ElementCountVector)
            Else
                Mres(j, i) = V(ctr Mod ElementCountVector)
            End If
        Else
            If FillOption = nRowByRow Then
                Mres(i, j) = V(ElementCountVector)
            Else
                Mres(j, i) = V(ElementCountVector)
            End If
        End If
    Next j
Next i
FQ_vector_to_matrix = Mres
End Function
'******************************************************************
' Reads the elements of a matrix either row by row, or column by column,
' and writes their values into a vector
' - error if input argument M is not a matrix
' - error if invalid fill option
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_to_vector(M() As Double, FillOption) As Double()
Attribute FQ_matrix_to_vector.VB_Description = "Reads the elements of a matrix either row by row, or column by column, and writes their values into a vector."
Attribute FQ_matrix_to_vector.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ElementCountMatrix As Long, i As Long, j As Long, ctr As Long
Dim Vres() As Double, iLim As Long, jLim As Long
Dim nrow As Long, ncol As Long
' Check if matrix
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_matrix_to_vector: Input argument M must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' check fill option
If FillOption <> nRowByRow And FillOption <> nColByCol Then
    FQ_MessageBox ("Error in FQ_matrix_to_vector: Improper fill option!")
    Err.Raise (FQ_ErrorNum)
End If
' get row and column sizes
nrow = UBound(M, 1)
ncol = UBound(M, 2)
ElementCountMatrix = nrow * ncol
ReDim Vres(1 To ElementCountMatrix)

If FillOption = nRowByRow Then
    iLim = nrow
    jLim = ncol
Else
    iLim = ncol
    jLim = nrow
End If

For i = 1 To iLim
    For j = 1 To jLim
        ctr = jLim * (i - 1) + j
        
        If FillOption = nRowByRow Then
            Vres(ctr) = M(i, j)
        Else
            Vres(ctr) = M(j, i)
        End If
    Next j
Next i
FQ_matrix_to_vector = Vres
End Function
'******************************************************************
' Creates matrix with sequential element values with given row and
' column sizes. Fills matrix row-wise with numbers.
' - set Interval = 0 for constant element values
' - error input arguments nrow and ncol are not positive integers
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_create(StartValue As Double, Interval As Double, nrow As Long, ncol As Long) As Double()
Attribute FQ_matrix_create.VB_Description = "Creates matrix with sequential element values with given row and column sizes. Fills matrix row-wise with sequential numbers."
Attribute FQ_matrix_create.VB_ProcData.VB_Invoke_Func = " \n14"
Dim M() As Double
Dim i As Long, j As Long
' check row and column sizes
If Not (nrow >= 0 And ncol > 0) Then
    FQ_MessageBox ("Error in FQ_create_matrix: nrow and ncol must be positive integers!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' create matrix
ReDim M(1 To nrow, 1 To ncol)
' fill all elements row-wise
For i = 1 To nrow
    For j = 1 To ncol
        M(i, j) = StartValue + Interval * ((i - 1) * ncol + j - 1)
    Next j
Next i
FQ_matrix_create = M
End Function
'******************************************************************
' Creates matrix with random element values between 0 and 1
' with given row and column sizes. Fills matrix row-wise with numbers.
' - error input arguments nrow and ncol are not positive integers
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_rand(nrow As Long, ncol As Long) As Double()
Attribute FQ_matrix_rand.VB_Description = "Creates matrix with random element values between 0 and 1 with given row and column sizes. Fills matrix row-wise with random numbers."
Attribute FQ_matrix_rand.VB_ProcData.VB_Invoke_Func = " \n14"
Dim M() As Double
Dim i As Long, j As Long
' check row and column sizes
If Not (nrow >= 0 And ncol > 0) Then
    FQ_MessageBox ("Error in FQ_matrix_rand: nrow and ncol must be positive integers!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' create matrix
ReDim M(1 To nrow, 1 To ncol)
' fill all elements row-wise
For i = 1 To nrow
    For j = 1 To ncol
        M(i, j) = Rnd
    Next j
Next i
FQ_matrix_rand = M
End Function
'******************************************************************
' Creates a vector with given length (ElementCount), start value
' and interval between subsequent elements.
' - error if ElementCount is not a positive integer
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_sequence(StartValue As Double, Interval As Double, ElementCount As Long) As Double()
Attribute FQ_vector_sequence.VB_Description = "Creates a vector with given length (ElementCount), start value and interval between subsequent elements."
Attribute FQ_vector_sequence.VB_ProcData.VB_Invoke_Func = " \n14"
Dim V() As Double
Dim i As Long
' check ElementCount
If Not (ElementCount > 0) Then
    FQ_MessageBox ("Error in FQ_vector_sequence: ElementCount must be a positive integer!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' create vector
ReDim V(1 To ElementCount)
For i = 1 To ElementCount
    V(i) = StartValue + Interval * (i - 1)
Next i
FQ_vector_sequence = V
End Function
'******************************************************************
' Creates a vector with random element values between 0 and 1
' with given vector length (ElementCount)
' and interval between subsequent elements.
' - error if ElementCount is not a positive integer
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_rand(ElementCount As Long) As Double()
Attribute FQ_vector_rand.VB_Description = "Creates a vector with random element values between 0 and 1 with given vector length (ElementCount)."
Attribute FQ_vector_rand.VB_ProcData.VB_Invoke_Func = " \n14"
Dim V() As Double
Dim i As Long
' check ElementCount
If Not (ElementCount > 0) Then
    FQ_MessageBox ("Error in FQ_vector_rand: ElementCount must be a positive integer!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' create vector
ReDim V(1 To ElementCount)
For i = 1 To ElementCount
    V(i) = Rnd
Next i
FQ_vector_rand = V
End Function
'******************************************************************
' Check if all vector values are positive integers within limits
' i.e. return True if:
' 1) LowerLimit <= all values <= UpperLimit, and
' 2) all values are whole numbers
' - often used to check the validity of matrix or vector indices
' - error if input argument V is not a vector
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_check_index_values(V As Variant, UpperLimit As Long, LowerLimit As Long) As Boolean
Attribute FQ_check_index_values.VB_Description = "Check if all vector values are positive integers within limits."
Attribute FQ_check_index_values.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ElementCount As Long, i As Long
' Check if vector
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_vector_check_values: Input argument V must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
FQ_check_index_values = True
ElementCount = UBound(V)
For i = 1 To ElementCount
    If V(i) < LowerLimit Or V(i) > UpperLimit Or V(i) <> Int(V(i)) Then
        FQ_check_index_values = False
        Exit Function
    End If
Next i
End Function
'******************************************************************
' returns partition of a matrix indicated by column and row index vectors
' - no row/column index vector means, all rows/columns are selected
' - error if any index element is not a positive integer larger than 0
' - error if any index element is larger than row/col size of matrix
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_partition(M() As Double, Optional RowInd As Variant, Optional Colind As Variant) As Double()
Attribute FQ_matrix_partition.VB_Description = "Returns partition of a matrix indicated by column and row index vectors."
Attribute FQ_matrix_partition.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nrow As Long, ncol As Long, i As Long, j As Long
Dim rowlen As Long, collen As Long
Dim Mres() As Double
Dim rowindx() As Double, colindx() As Double
' Check if matrix
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_matrix_partition: Input argument M must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' Get matrix size
nrow = UBound(M, 1)
ncol = UBound(M, 2)

If Not IsMissing(RowInd) Then
    ' Check if vector
    If Not FQ_CheckIfVector(RowInd) Then
        FQ_MessageBox ("Error in FQ_matrix_partition: Input argument rowind must be a vector!")
        Err.Raise (FQ_ErrorNum)
        Exit Function
    End If
    ' Check index values
    If Not FQ_check_index_values(RowInd, nrow, 1) Then
        FQ_MessageBox ("Error in FQ_matrix_partition: Improper row indices!")
        Err.Raise (FQ_ErrorNum)
        Exit Function
    End If
    rowindx = RowInd
Else
    rowindx = FQ_vector_sequence(1, 1, nrow)
End If

If Not IsMissing(Colind) Then
    ' Check if vector
    If Not FQ_CheckIfVector(Colind) Then
        FQ_MessageBox ("Error in FQ_matrix_partition: Input argument colind must be a vector!")
        Err.Raise (FQ_ErrorNum)
        Exit Function
    End If
    ' Check index values
    If Not FQ_check_index_values(Colind, ncol, 1) Then
        FQ_MessageBox ("Error in FQ_matrix_partition: Improper column indices!")
        Err.Raise (FQ_ErrorNum)
        Exit Function
    End If
    colindx = Colind
Else
    colindx = FQ_vector_sequence(1, 1, ncol)
End If

' Get vector sizes
rowlen = UBound(rowindx)
collen = UBound(colindx)
ReDim Mres(1 To rowlen, 1 To collen)
' fill values
For i = 1 To rowlen
    For j = 1 To collen
        Mres(i, j) = M(rowindx(i), colindx(j))
    Next j
Next i
FQ_matrix_partition = Mres
End Function
'******************************************************************
' returns partition of a vector indicated index vector ind
' - error if any index element is not a positive integer larger than 0
' - error if any index element is larger than vector length
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_partition(V() As Double, ind() As Double) As Double()
Attribute FQ_vector_partition.VB_Description = "Returns partition of a vector indicated index vector ind."
Attribute FQ_vector_partition.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nlen As Long, i As Long, indlen As Long
Dim Vres() As Double
' Check if vector
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_vector_partition: Input argument V must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' Check if vector
If Not FQ_CheckIfVector(ind) Then
    FQ_MessageBox ("Error in FQ_vector_partition: Input argument ind must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' Get vector size
nlen = UBound(V, 1)
' Check index values
If Not FQ_check_index_values(ind, nlen, 1) Then
    FQ_MessageBox ("Error in FQ_vector_partition: Improper index value!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' Get vector sizes
indlen = UBound(ind)
ReDim Vres(1 To indlen)
' fill values
For i = 1 To indlen
        Vres(i) = V(ind(i))
Next i
FQ_vector_partition = Vres
End Function
'******************************************************************
' Returns the sum of elements of matrix M, either row or column wise
' - Rowwise sum returns a horizontal 1xNcol matrix
' - Columnwise sum returns a vertical 1 xNrow matrix
' - Element sum (all elements) returns a 1x1 matrix
' - error if M is not a matrix
' - error if SumOption is not 1 (nRowWiseSum) or 2 (nColWiseSum) or 3 (nElementSum)
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_element_sum(M() As Double, SumOption As MatrixDirection) As Double()
Attribute FQ_matrix_element_sum.VB_Description = "Returns the sum of elements of matrix M, either row or column wise."
Attribute FQ_matrix_element_sum.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Msum() As Double, xsum As Double
Dim nrow As Integer, ncol As Integer
Dim i As Integer, j As Integer

If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_matrix_element_sum: Input argument M must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
    
' get row and column sizes
nrow = UBound(M, 1)
ncol = UBound(M, 2)
    
    Select Case SumOption
        Case nRowByRow
            ReDim Msum(1 To 1, 1 To ncol)   ' 1 x ncol horizontal matrix
            For i = 1 To ncol
                xsum = 0    ' initiate sum of a column
                For j = 1 To nrow
                    xsum = xsum + M(j, i)
                Next j
                Msum(1, i) = xsum
             Next i
        Case nColByCol
            ReDim Msum(1 To nrow, 1 To 1)   ' nrow x 1 vectical matrix
            For i = 1 To nrow
                xsum = 0    ' initiate sum of a row
                For j = 1 To ncol
                    xsum = xsum + M(i, j)
                Next j
                Msum(i, 1) = xsum
             Next i
        Case nAllElements
            ReDim Msum(1 To 1, 1 To 1)   ' 1x1 unity matrix
            xsum = 0        ' initiate sum of all elements
            For i = 1 To nrow
                For j = 1 To ncol
                    xsum = xsum + M(i, j)
                Next j
             Next i
             Msum(1, 1) = xsum
        Case Else
            FQ_MessageBox ("Error in FQ_matrix_element_sum: Unknown sum option!")
            Err.Raise (FQ_ErrorNum)
            Exit Function
    End Select
FQ_matrix_element_sum = Msum
End Function
'******************************************************************
' Applies the given aggregation function (sum, min, max, avg, median)
' on the matrix, and returns a scalar number.
' - AggregateOption: nRowByRow, nColByCol or nAllElements
' -  nColByCol aggregation returns a 1xN horizontal matrix
' -  nRowByRow aggregation returns a Nx1 vertical matrix
' - nAllElements aggregation returns a 1x1 unity matrix
' - error if M is not a vector
' - error if unknown aggregation function
' - error if unknown aggregation option
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_aggregate(M() As Double, AggregateFunct As String, _
        AggregateOption As MatrixDirection) As Double()
Attribute FQ_matrix_aggregate.VB_Description = "Applies the given aggregation function (sum, min, max, avg, median) on the matrix."
Attribute FQ_matrix_aggregate.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Mres() As Double, x As Double, Msub() As Double
Dim nrow As Integer, ncol As Integer
Dim i As Integer, j As Integer
' Check if matrix
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_matrix_aggregate: Input argument M must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' get row and column sizes
nrow = UBound(M, 1)
ncol = UBound(M, 2)
    
Select Case AggregateOption
    Case nColByCol
        ReDim Mres(1 To 1, 1 To ncol)   ' 1 x ncol horizontal matrix
        For i = 1 To ncol
            ' sub-matrix with i'th column of M
            Msub = FQ_matrix_partition(M, Colind:=FQ_var_to_vector(Array(i)))
            Select Case LCase(AggregateFunct)
                Case "sum"
                    Mres(1, i) = Application.WorksheetFunction.Sum(Msub)
                Case "avg", "average", "mean"
                    Mres(1, i) = Application.WorksheetFunction.Average(Msub)
                Case "min", "minimum"
                    Mres(1, i) = Application.WorksheetFunction.Min(Msub)
                Case "max", "maximum"
                    Mres(1, i) = Application.WorksheetFunction.Max(Msub)
                Case "med", "median"
                    Mres(1, i) = Application.WorksheetFunction.Median(Msub)
                Case Else
                    FQ_MessageBox ("Error in FQ_matrix_aggregate: Invalid aggregation function!")
                    Exit Function
            End Select
         Next i
         
     Case nRowByRow
        ReDim Mres(1 To nrow, 1 To 1)   ' nrow x 1 vertical matrix
        For i = 1 To nrow
            ' sub-matrix with i'th row of M
            Msub = FQ_matrix_partition(M, RowInd:=FQ_var_to_vector(Array(i)))
            Select Case LCase(AggregateFunct)
                Case "sum"
                    Mres(i, 1) = Application.WorksheetFunction.Sum(Msub)
                Case "avg", "average", "mean"
                    Mres(i, 1) = Application.WorksheetFunction.Average(Msub)
                Case "min", "minimum"
                    Mres(i, 1) = Application.WorksheetFunction.Min(Msub)
                Case "max", "maximum"
                    Mres(i, 1) = Application.WorksheetFunction.Max(Msub)
                Case "med", "median"
                    Mres(i, 1) = Application.WorksheetFunction.Median(Msub)
                Case Else
                    FQ_MessageBox ("Error in FQ_matrix_aggregate: Invalid aggregation function!")
                    Exit Function
            End Select
         Next i
     
    Case nAllElements
        ReDim Mres(1 To 1, 1 To 1)   ' 1 x 1 unity matrix
        Select Case LCase(AggregateFunct)
            Case "sum"
                Mres(1, 1) = Application.WorksheetFunction.Sum(M)
            Case "avg", "average", "mean"
                Mres(1, 1) = Application.WorksheetFunction.Average(M)
            Case "min", "minimum"
                Mres(1, 1) = Application.WorksheetFunction.Min(M)
            Case "max", "maximum"
                Mres(1, 1) = Application.WorksheetFunction.Max(M)
            Case "med", "median"
                Mres(1, 1) = Application.WorksheetFunction.Median(M)
            Case Else
                FQ_MessageBox ("Error in FQ_matrix_aggregate: Invalid aggregation function!")
                Exit Function
        End Select

    Case Else
        FQ_MessageBox ("Error in FQ_matrix_aggregate: Unknown sum option!")
        Err.Raise (FQ_ErrorNum)
        Exit Function
End Select
FQ_matrix_aggregate = Mres
End Function
'******************************************************************
' Adds a scalar number to all elements of matrix
' - error if M is not a matrix
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_scalar_add(M() As Double, x As Double) As Double()
Attribute FQ_matrix_scalar_add.VB_Description = "Adds a scalar number to all elements of matrix."
Attribute FQ_matrix_scalar_add.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nrow As Integer, ncol As Integer
Dim Mres() As Double
Dim i As Integer, j As Integer

' Check if matrix
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in matrix_scalar_add: Input argument M must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' get row and column sizes
nrow = UBound(M, 1)
ncol = UBound(M, 2)
ReDim Mres(1 To nrow, 1 To ncol)

For i = 1 To nrow
    For j = 1 To ncol
        Mres(i, j) = M(i, j) + x
    Next j
Next i
FQ_matrix_scalar_add = Mres
End Function
'******************************************************************
' Spreadsheet version of the function FQ_matrix_scalar_add
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQS_matrix_scalar_add(M As Range, x As Variant)
Attribute FQS_matrix_scalar_add.VB_Description = "Spreadsheet version of the function FQ_matrix_scalar_add: Adds a scalar number to all elements of matrix."
Attribute FQS_matrix_scalar_add.VB_ProcData.VB_Invoke_Func = " \n14"
Dim M1() As Double, x1 As Double
Dim Result() As Double
On Error GoTo EH1
M1 = FQ_range_to_matrix(M)
On Error GoTo EH2
x1 = CDbl(x)
On Error GoTo 0
FQS_matrix_scalar_add = FQ_matrix_scalar_add(M1, x1)
Exit Function
EH1:
FQ_MessageBox ("Error in FQS_matrix_scalar_add: Matrix M could not be read! Please select an area with all cells filled with numeric values.")
Err.Raise (FQ_ErrorNum)
Exit Function
EH2:
FQ_MessageBox ("Error in FQS_matrix_scalar_add: Scalar X could not be read! Please select a single numeric cell for X.")
Err.Raise (FQ_ErrorNum)
End Function
'******************************************************************
' Adds a scalar number to all elements of vector
' - error if V is not a vector
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_scalar_add(V() As Double, x As Double) As Double()
Attribute FQ_vector_scalar_add.VB_Description = "Adds a scalar number to all elements of vector."
Attribute FQ_vector_scalar_add.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nlen As Integer
Dim Vres() As Double
Dim i As Integer

' Check if vector
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_vector_scalar_add: Input argument V must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' get vector length
nlen = UBound(V, 1)
ReDim Vres(1 To nlen)

For i = 1 To nlen
        Vres(i) = V(i) + x
Next i
FQ_vector_scalar_add = Vres
End Function
'******************************************************************
' Multiplies all elements of vector with a scalar number x
' - error if V is not a vector
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_scalar_multiply(V() As Double, x As Double) As Double()
Attribute FQ_vector_scalar_multiply.VB_Description = "Multiplies all elements of vector with a scalar number x."
Attribute FQ_vector_scalar_multiply.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nlen As Integer
Dim Vres() As Double
Dim i As Integer

' Check if vector
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_vector_scalar_multiply: Input argument V must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If

' get vector length
nlen = UBound(V, 1)
ReDim Vres(1 To nlen)

For i = 1 To nlen
        Vres(i) = V(i) * x
Next i
 
FQ_vector_scalar_multiply = Vres
End Function
'******************************************************************
' Applies the given aggregation function (sum, min, max, avg, median)
' on the vector, and returns a scalar number.
' - error if V is not a vector
' - error if unknown aggregation function
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_aggregate(V() As Double, AggregateFunct As String) As Double
Attribute FQ_vector_aggregate.VB_Description = "Applies the given aggregation function (sum, min, max, avg, median) on the vector, and returns a scalar number."
Attribute FQ_vector_aggregate.VB_ProcData.VB_Invoke_Func = " \n14"
Dim x As Double
' Check if vector
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_vector_aggregate: Input argument V must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If

Select Case LCase(AggregateFunct)
    Case "sum"
        x = Application.WorksheetFunction.Sum(V)
    Case "avg", "average", "mean"
        x = Application.WorksheetFunction.Average(V)
    Case "min", "minimum"
        x = Application.WorksheetFunction.Min(V)
    Case "max", "maximum"
        x = Application.WorksheetFunction.Max(V)
    Case "med", "median"
        x = Application.WorksheetFunction.Median(V)
    Case Else
        FQ_MessageBox ("Error in FQ_vector_aggregate: Invalid aggregation function!")
        Exit Function
End Select
FQ_vector_aggregate = x
End Function
'******************************************************************
' Returns the number of elements in matrix M
' - error if M is not a matrix
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_element_count(M() As Double)
Attribute FQ_matrix_element_count.VB_Description = "Returns the total number of elements in matrix M."
Attribute FQ_matrix_element_count.VB_ProcData.VB_Invoke_Func = " \n14"
' Check if matrix
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_matrix_scalar_multiply: Input argument M must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
FQ_matrix_element_count = UBound(M, 1) * UBound(M, 2)
End Function
'******************************************************************
' Multiplies all elements of matrix with a scalar number
' - error if M is not a matrix
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_scalar_multiply(M() As Double, x As Double) As Double()
Attribute FQ_matrix_scalar_multiply.VB_Description = "Multiplies all elements of matrix with a scalar number x."
Attribute FQ_matrix_scalar_multiply.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nrow As Integer, ncol As Integer
Dim Mres() As Double
Dim i As Integer, j As Integer

' Check if matrix
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_matrix_scalar_multiply: Input argument M must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If

' get row and column sizes
nrow = UBound(M, 1)
ncol = UBound(M, 2)
ReDim Mres(1 To nrow, 1 To ncol)

For i = 1 To nrow
    For j = 1 To ncol
        Mres(i, j) = M(i, j) * x
    Next j
Next i
 
FQ_matrix_scalar_multiply = Mres
End Function
'******************************************************************
' Multiplies rows or columns of matrix M with corresponding elements
' of vector V
' MultiplyOption: nRowByRow, nColByCol
' - error if M is not a matrix, or if V is not a vector
' - error if vector size is not equal to row/column size of matrix
' depending on MultiplyOption
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_vector_multiply(M() As Double, V() As Double, MultiplyOption As Byte) As Double()
Attribute FQ_matrix_vector_multiply.VB_Description = "Multiplies rows or columns of matrix M with corresponding elements of vector V."
Attribute FQ_matrix_vector_multiply.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nrow As Integer, ncol As Integer, vlen As Long
Dim i As Long, j As Long, Mres() As Double
' Check if matrix
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_matrix_vector_multiply: Input argument M must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' Check if vector
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_matrix_vector_multiply: Input argument V must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' get row and column sizes
nrow = UBound(M, 1)
ncol = UBound(M, 2)
vlen = UBound(V, 1)
If MultiplyOption = nRowByRow And vlen <> nrow Then
    FQ_MessageBox ("Error in FQ_matrix_vector_multiply: Vector length must be equal to row size of matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
If MultiplyOption = nColByCol And vlen <> ncol Then
    FQ_MessageBox ("Error in FQ_matrix_vector_multiply: Vector length must be equal to column size of matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If

ReDim Mres(1 To nrow, 1 To ncol)

Select Case MultiplyOption
    Case nRowByRow
        For i = 1 To nrow
            For j = 1 To ncol
                Mres(i, j) = M(i, j) * V(i)
            Next j
        Next i
    Case nColByCol
        For i = 1 To nrow
            For j = 1 To ncol
                Mres(i, j) = M(i, j) * V(j)
            Next j
        Next i
    Case Else
        FQ_MessageBox ("Error in FQ_matrix_vector_multiply: Invalid multiply option!")
        Err.Raise (FQ_ErrorNum)
        Exit Function
End Select
FQ_matrix_vector_multiply = Mres
End Function
'******************************************************************
' Adds up the elements of two matrices with identical row/column sizes
' i.e. element-wise sum of two matrices
' - error if M1 and/or M2 is not a matrix
' - error if row and column sizes of matrices M1 and M2 are not identical
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_matrix_sum(M1() As Double, M2() As Double) As Double()
Attribute FQ_matrix_matrix_sum.VB_Description = "Adds up the elements of two matrices with identical row/column sizes."
Attribute FQ_matrix_matrix_sum.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nrow As Integer, ncol As Integer
Dim i As Integer, j As Integer
Dim Mres() As Double

' Check if matrix
If Not FQ_CheckIfMatrix(M1) Or Not FQ_CheckIfMatrix(M2) Then
    FQ_MessageBox ("Error in FQ_matrix_matrix_sum: Input arguments M1 and M2 must be matrices!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' check if matrix sizes are identical
nrow = UBound(M1, 1)
ncol = UBound(M1, 2)
If UBound(M2, 1) <> nrow Or UBound(M2, 2) <> ncol Then
    FQ_MessageBox ("Error in FQ_matrix_matrix_sum: Size of matrices M1 and M2 must be identical!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If

' sum matrix elements
ReDim Mres(1 To nrow, 1 To ncol)
For i = 1 To nrow
    For j = 1 To ncol
        Mres(i, j) = M1(i, j) + M2(i, j)
    Next j
Next i
FQ_matrix_matrix_sum = Mres
End Function
'******************************************************************
' Adds up the elements of two vectors with identical lengths
' i.e. element-wise sum of two equal-sized vectors
' - error if V1 and/or V2 is not a vector
' - error if lengths of vectors V1 and V2 are not identical
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_vector_sum(V1() As Double, V2() As Double) As Double()
Attribute FQ_vector_vector_sum.VB_Description = "Adds up the elements of two vectors with identical lengths."
Attribute FQ_vector_vector_sum.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nlen As Integer
Dim i As Integer
Dim Vres() As Double

' Check if vector
If Not FQ_CheckIfVector(V1) Or Not FQ_CheckIfVector(V2) Then
    FQ_MessageBox ("Error in FQ_vector_vector_sum: Input arguments V1 and V2 must be vectors!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' check if vector sizes are identical
nlen = UBound(V1, 1)

If UBound(V2, 1) <> nlen Then
    FQ_MessageBox ("Error in FQ_vector_vector_sum: Sizes of vectors V1 and V2 must be identical!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If

' sum vector elements
ReDim Vres(1 To nlen)
For i = 1 To nlen
        Vres(i) = V1(i) + V2(i)
Next i
FQ_vector_vector_sum = Vres
End Function
'******************************************************************
' Elementwise multiplication of two equal-sized matrices
' R = M1 .* M2 (matlab notation)
' - error if M1 and/or M2 is not a matrix
' - error if row and column sizes of matrices M1 and M2 are not identical
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_elementwise_multiply(M1() As Double, M2() As Double) As Double()
Attribute FQ_matrix_elementwise_multiply.VB_Description = "Elementwise multiplication of two equal-sized matrices; R = M1 .* M2 (matlab notation)."
Attribute FQ_matrix_elementwise_multiply.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nrow As Integer, ncol As Integer
Dim i As Integer, j As Integer
Dim Mres() As Double

' Check if matrix
If Not FQ_CheckIfMatrix(M1) Or Not FQ_CheckIfMatrix(M2) Then
    FQ_MessageBox ("Error in FQ_matrix_elementwise_multiply: Input arguments M1 and M2 must be matrices!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' check if matrix sizes are identical
nrow = UBound(M1, 1)
ncol = UBound(M1, 2)
If UBound(M2, 1) <> nrow Or UBound(M2, 2) <> ncol Then
    FQ_MessageBox ("Error in FQ_matrix_elementwise_multiply: Size of matrices M1 and M2 must be identical!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If

' sum matrix elements
ReDim Mres(1 To nrow, 1 To ncol)
For i = 1 To nrow
    For j = 1 To ncol
        Mres(i, j) = M1(i, j) * M2(i, j)
    Next j
 Next i
FQ_matrix_elementwise_multiply = Mres
End Function
'******************************************************************
' Elementwise division of two equal-sized matrices
' R = M1 ./ M2 (matlab notation)
' - error if M1 and/or M2 is not a matrix
' - error if row and column sizes of matrices M1 and M2 are not identical
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_elementwise_divide(M1() As Double, M2() As Double) As Double()
Attribute FQ_matrix_elementwise_divide.VB_Description = "Elementwise division of two equal-sized matrices; R = M1 ./ M2 (matlab notation)."
Attribute FQ_matrix_elementwise_divide.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nrow As Integer, ncol As Integer
Dim i As Integer, j As Integer
Dim Mres() As Double

' Check if matrix
If Not FQ_CheckIfMatrix(M1) Or Not FQ_CheckIfMatrix(M2) Then
    FQ_MessageBox ("Error in FQ_matrix_elementwise_divide: Input arguments M1 and M2 must be matrices!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' check if matrix sizes are identical
nrow = UBound(M1, 1)
ncol = UBound(M1, 2)
If UBound(M2, 1) <> nrow Or UBound(M2, 2) <> ncol Then
    FQ_MessageBox ("Error in FQ_matrix_elementwise_divide: Size of matrices M1 and M2 must be identical!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If

' sum matrix elements
ReDim Mres(1 To nrow, 1 To ncol)
For i = 1 To nrow
    For j = 1 To ncol
        Mres(i, j) = M1(i, j) / M2(i, j)
    Next j
 Next i
FQ_matrix_elementwise_divide = Mres
End Function
'******************************************************************
' Reverse the order of vector elements; f.e. [1 2 3] --> [3 2 1]
' - error if V is not a vector
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_reverse(V() As Double) As Double()
Attribute FQ_vector_reverse.VB_Description = "Reverse the order of vector elements; f.e. [1 4 3] --> [3 4 1]."
Attribute FQ_vector_reverse.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i As Long, vlen As Long, Vres() As Double
' Check if vector
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_vector_reverse: Input argument V must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' get vector size
vlen = UBound(V, 1)
ReDim Vres(1 To vlen)
For i = 1 To vlen
    Vres(i) = V(vlen - i + 1)
Next i
FQ_vector_reverse = Vres
End Function
'******************************************************************
' Matrix multiplication in linear algebra,  C = A x B
' - error of M1 and/or M2 is not a matrix
' - error if matrix sizes don't match: ncol of M1 must be equal to nrow of B
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_multiplication(M1() As Double, M2() As Double) As Double()
Attribute FQ_matrix_multiplication.VB_Description = "Matrix multiplication in linear algebra,  C = A x B."
Attribute FQ_matrix_multiplication.VB_ProcData.VB_Invoke_Func = " \n14"
' Check if matrix
If Not FQ_CheckIfMatrix(M1) Or Not FQ_CheckIfMatrix(M2) Then
    FQ_MessageBox ("Error in FQ_matrix_multiplication: Input arguments M1 and M2 must be matrices!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' check if matrix sizes match
If Not (UBound(M1, 2) = UBound(M2, 1)) Then
    FQ_MessageBox ("Error in FQ_matrix_multiplication: Matrix sizes don't match for multiplication!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' worksheet function returns variant array
FQ_matrix_multiplication = FQ_var_to_matrix(Application.WorksheetFunction.MMult(M1, M2))
End Function
'******************************************************************
' Appends matrix M2 to M1 either vertically or horizontally
' AppendOption: AppendVertically or AppendHorizontally
' - returns empty array if both matrices are empty
' - returns M2 if M1 is empty
' - returns M1 if M2 is empty
' - error if M1 is not a matrix unless it is empty
' - error if M2 is not a matrix unless it is empty
' - error if matrix sizes don't match for an append operation
'   row sizes must be equal for horizontal append
'   column sizes must be equal for vertical append
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_append(M1() As Double, M2() As Double, AppendOption As MatrixAlignment) As Double()
Attribute FQ_matrix_append.VB_Description = "Appends matrix M2 to M1 either vertically or horizontally."
Attribute FQ_matrix_append.VB_ProcData.VB_Invoke_Func = " \n14"
Dim M() As Double
Dim nrow1 As Long, ncol1 As Long, nrow2 As Long, ncol2 As Long
Dim i As Long, j As Long
' check if both matrices are empty
If FQ_ArrayDimension(M1) = 0 And FQ_ArrayDimension(M2) = 0 Then
    FQ_matrix_append = M
    Exit Function
End If
' check if a matrix is empty
If FQ_ArrayDimension(M1) = 0 Then
    ' Check if matrix
    If Not (FQ_CheckIfMatrix(M2)) Then
        FQ_MessageBox ("Error in FQ_matrix_append: Input argument M2 must be a matrix! (M1 is empty)")
        Err.Raise (FQ_ErrorNum)
        Exit Function
    End If
    FQ_matrix_append = M2
    Exit Function
End If
If FQ_ArrayDimension(M2) = 0 Then
    ' Check if matrix
    If Not (FQ_CheckIfMatrix(M1)) Then
        FQ_MessageBox ("Error in FQ_matrix_append: Input argument M1 must be a matrix! (M2 is empty)")
        Err.Raise (FQ_ErrorNum)
        Exit Function
    End If
    FQ_matrix_append = M1
    Exit Function
End If
' Check if matrix (both non-empty)
If Not (FQ_CheckIfMatrix(M1) And FQ_CheckIfMatrix(M2)) Then
    FQ_MessageBox ("Error in FQ_matrix_append: Input arguments M1 and M2 must be matrices!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' get matrix sizes
nrow1 = UBound(M1, 1)
ncol1 = UBound(M1, 2)
nrow2 = UBound(M2, 1)
ncol2 = UBound(M2, 2)

Select Case AppendOption
    Case nHorizontal
        If nrow1 <> nrow2 Then
            FQ_MessageBox ("Error in FQ_matrix_append: Row sizes of M1 and M2 must be equal for horizontal append!")
            Err.Raise (FQ_ErrorNum)
            Exit Function
        End If
        ReDim M(1 To nrow1, 1 To (ncol1 + ncol2))
        ' fill values of M1
        For i = 1 To nrow1
            For j = 1 To ncol1
                M(i, j) = M1(i, j)
            Next j
        Next i
        ' fill values of M2
        For i = 1 To nrow2
            For j = 1 To ncol2
                M(i, j + ncol1) = M2(i, j)
            Next j
        Next i
        
    Case nVertical
        If ncol1 <> ncol2 Then
            FQ_MessageBox ("Error in FQ_matrix_append: Column sizes of M1 and M2 must be equal for vertical append!")
            Err.Raise (FQ_ErrorNum)
            Exit Function
        End If
        ReDim M(1 To (nrow1 + nrow2), 1 To ncol1)
        ' fill values of M1
        For i = 1 To nrow1
            For j = 1 To ncol1
                M(i, j) = M1(i, j)
            Next j
        Next i
        ' fill values of M2
        For i = 1 To nrow2
            For j = 1 To ncol2
                M(i + nrow1, j) = M2(i, j)
            Next j
        Next i
        
    Case Else
        FQ_MessageBox ("Error in FQ_matrix_append: Invalid append option!")
        Err.Raise (FQ_ErrorNum)
        Exit Function
End Select
FQ_matrix_append = M
End Function
'******************************************************************
' Appends vector V2 to V1 such that result vector = [V1, V2]
' - returns empty array if both vectors are empty
' - returns V2 if V1 is empty
' - returns V1 if V2 is empty
' - error if V1 is not a vector unless it is empty
' - error if V2 is not a vector unless it is empty
' - error if V1 and/or V2 is not a vector unless they are empty
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_vector_append(V1() As Double, V2() As Double) As Double()
Attribute FQ_vector_vector_append.VB_Description = "Appends vector V2 to V1 such that result vector = [V1, V2]."
Attribute FQ_vector_vector_append.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Vres() As Double, i As Long, nlen1 As Long, nlen2 As Long
' check if both vectors are empty
If FQ_ArrayDimension(V1) = 0 And FQ_ArrayDimension(V2) = 0 Then
    FQ_vector_vector_append = Vres
    Exit Function
End If
' check if a vector is empty
If FQ_ArrayDimension(V1) = 0 Then
    ' Check if vector
    If Not (FQ_CheckIfVector(V2)) Then
        FQ_MessageBox ("Error in FQ_vector_vector_append: Input argument V2 must be a vector! (V1 is empty)")
        Err.Raise (FQ_ErrorNum)
        Exit Function
    End If
    FQ_vector_vector_append = V2
    Exit Function
End If
If FQ_ArrayDimension(V2) = 0 Then
    ' Check if vector
    If Not (FQ_CheckIfVector(V1)) Then
        FQ_MessageBox ("Error in FQ_vector_vector_append: Input argument V1 must be a vector! (V2 is empty)")
        Err.Raise (FQ_ErrorNum)
        Exit Function
    End If
    FQ_vector_vector_append = V1
    Exit Function
End If
' Check if vector
If Not FQ_CheckIfVector(V1) Then
    FQ_MessageBox ("Error in FQ_vector_vector_append: Input argument V1 must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
If Not FQ_CheckIfVector(V2) Then
    FQ_MessageBox ("Error in FQ_vector_vector_append: Input argument V2 must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
'Get vector sizes
nlen1 = UBound(V1, 1)
nlen2 = UBound(V2, 1)

Vres = V1
ReDim Preserve Vres(1 To (nlen1 + nlen2))
For i = 1 To nlen2
    Vres(nlen1 + i) = V2(i)
Next i
FQ_vector_vector_append = Vres
End Function
'******************************************************************
' Applies the given mathematical operation like abs(), fix(), sin() etc.
' (all available single-argument VBA functions) on all elements of the matrix M
' - error if input argument M is not a matrix
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_operation(M() As Double, OperationName As String) As Double()
Attribute FQ_matrix_operation.VB_Description = "Applies the given mathematical operation like abs(), fix(), sin() etc. (all available single-argument VBA functions) on all elements of the matrix M."
Attribute FQ_matrix_operation.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Mres() As Double
Dim i As Long, j As Long, nrow As Long, ncol As Long
' Check if matrix
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_matrix_operation: Input argument M must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' get row and column sizes
nrow = UBound(M, 1)
ncol = UBound(M, 2)
ReDim Mres(1 To nrow, 1 To ncol)

' apply operation on each element
On Error GoTo ErrorHandler
For i = 1 To nrow
    For j = 1 To ncol
        ' Str() converts number to string
        Mres(i, j) = Evaluate(OperationName & "(" & str(M(i, j)) & ")")
    Next j
 Next i
FQ_matrix_operation = Mres
Exit Function
ErrorHandler:
FQ_MessageBox ("Error in FQ_matrix_operation: Unknown operation name!")
Err.Raise (Err.Number)
End Function
'******************************************************************
' Applies the given mathematical operation like abs(), fix(), sin() etc.
' (all available single-argument VBA functions) on all elements of the vector V
' - error if input argument V is not a vector
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_operation(V() As Double, OperationName As String) As Double()
Attribute FQ_vector_operation.VB_Description = "Applies the given mathematical operation like abs(), fix(), sin() etc. (all available single-argument VBA functions) on all elements of the vector V."
Attribute FQ_vector_operation.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Vres() As Double
Dim i As Long, vlen As Long
' Check if vector
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_vector_operation: Input argument V must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' get vector size
vlen = UBound(V, 1)
ReDim Vres(1 To vlen)
' apply operation on each element
On Error GoTo ErrorHandler
For i = 1 To vlen
    ' Str() converts number to string
    Vres(i) = Evaluate(OperationName & "(" & str(V(i)) & ")")
Next i
FQ_vector_operation = Vres
Exit Function
ErrorHandler:
FQ_MessageBox ("Error in FQ_vector_operation: Unknown operation name!")
Err.Raise (Err.Number)
End Function
'******************************************************************
' Assigns values of vector V2 to the partition of vector V1 selected
' by the index vector ind1. i.e. V1(ind1) = V2
' - error if V1 and/or V2 and/or ind1 are not vectors
' - error of length(V2) is not equal to length(ind1)
' - error if an index value in ind1 is not a positive integer between
'   1 and length of V1
' - error if index vector ind1 is not unique with distinct index values
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_partition_assign(V1() As Double, ind1() As Double, V2() As Double) As Double()
Attribute FQ_vector_partition_assign.VB_Description = "Assigns values of vector V2 to the partition of vector V1 selected by the index vector ind1. i.e. V1(ind1) = V2."
Attribute FQ_vector_partition_assign.VB_ProcData.VB_Invoke_Func = " \n14"
Dim vlen1 As Long, vlen2 As Long, i As Long, Vres() As Double
' Check if vector
If Not FQ_CheckIfVector(V1) Then
    FQ_MessageBox ("Error in FQ_vector_partition_assign: Input argument V1 must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
If Not FQ_CheckIfVector(V2) Then
    FQ_MessageBox ("Error in FQ_vector_partition_assign: Input argument V2 must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
If Not FQ_CheckIfVector(ind1) Then
    FQ_MessageBox ("Error in FQ_vector_partition_assign: Input argument ind1 must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' get vector lengths
vlen1 = UBound(V1, 1)
vlen2 = UBound(V2, 1)

If vlen2 <> UBound(ind1, 1) Then
    FQ_MessageBox ("Error in FQ_vector_partition_assign: Length of V2 must be equal to length of index vector ind1!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
If Not FQ_check_index_values(ind1, vlen1, 1) Then
    FQ_MessageBox ("Error in FQ_vector_partition_assign: Improper index values in ind1!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
If Not FQ_vector_if_unique(ind1) Then
    FQ_MessageBox ("Error in FQ_vector_partition_assign: ind1 must have unique index values!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If

Vres = V1
For i = 1 To vlen2
    Vres(ind1(i)) = V2(i)
Next i
FQ_vector_partition_assign = Vres
End Function
'******************************************************************
' Assigns values of matrix M2 to the partition of matrix M1 selected
' by the index vectors rowind1 and colind1: M1(rowind1, colind1) = M2
' - error if M1 and/or M2 are not matrices
' - error if rowind1 and/or rowind2 are not vectors
' - error if length of vectors rowind1/colind1 are not equal to row/column
' size of M2
' - error if index value in rowind1/colind1 are not a positive integers
' between 1 and nrow1/ncol1
' - error if index vectors rowind1/colind1 are not unique with distinct index values
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_matrix_partition_assign(M1() As Double, rowind1() As Double, colind1() As Double, M2() As Double) As Double()
Attribute FQ_matrix_partition_assign.VB_Description = "Assigns values of matrix M2 to the partition of matrix M1 selected by the index vectors rowind1 and colind1: M1(rowind1, colind1) = M2."
Attribute FQ_matrix_partition_assign.VB_ProcData.VB_Invoke_Func = " \n14"
Dim nrow1 As Long, ncol1 As Long, nrow2 As Long, ncol2 As Long
Dim i As Long, j As Long, Mres() As Double
' Check if matrix
If Not FQ_CheckIfMatrix(M1) Then
    FQ_MessageBox ("Error in FQ_matrix_partition_assign: Input argument M1 must be a matrix!")
    Err.Raise (Err.Number)
    Exit Function
End If
If Not FQ_CheckIfMatrix(M2) Then
    FQ_MessageBox ("Error in FQ_matrix_partition_assign: Input argument M2 must be a matrix!")
    Err.Raise (Err.Number)
    Exit Function
End If
' Check if vector
If Not FQ_CheckIfVector(rowind1) Then
    FQ_MessageBox ("Error in FQ_matrix_partition_assign: Input argument rowind1 must be a vector!")
    Err.Raise (Err.Number)
    Exit Function
End If
' Check if vector
If Not FQ_CheckIfVector(colind1) Then
    FQ_MessageBox ("Error in FQ_matrix_partition_assign: Input argument colind1 must be a vector!")
    Err.Raise (Err.Number)
    Exit Function
End If

'get matrix sizes
nrow1 = UBound(M1, 1)
ncol1 = UBound(M1, 2)
nrow2 = UBound(M2, 1)
ncol2 = UBound(M2, 2)

' check vector lengths
If nrow2 <> UBound(rowind1, 1) Then
    FQ_MessageBox ("Error in FQ_matrix_partition_assign: Length of rowind1 must be equal to row size of matrix M2!")
    Err.Raise (Err.Number)
    Exit Function
End If
If ncol2 <> UBound(colind1, 1) Then
    FQ_MessageBox ("Error in FQ_matrix_partition_assign: Length of colind1 must be equal to column size of matrix M2!")
    Err.Raise (Err.Number)
    Exit Function
End If
' check values in index vectors
If Not FQ_check_index_values(rowind1, nrow1, 1) Then
    FQ_MessageBox ("Error in FQ_matrix_partition_assign: Improper index values in rowind1!")
    Err.Raise (Err.Number)
    Exit Function
End If
If Not FQ_check_index_values(colind1, ncol1, 1) Then
    FQ_MessageBox ("Error in FQ_matrix_partition_assign: Improper index values in colind1!")
    Err.Raise (Err.Number)
    Exit Function
End If
' check if index vectors have distinct values
If Not FQ_vector_if_unique(rowind1) Then
    FQ_MessageBox ("Error in FQ_matrix_partition_assign: rowind1 must have unique index values!")
    Err.Raise (Err.Number)
    Exit Function
End If
If Not FQ_vector_if_unique(colind1) Then
    FQ_MessageBox ("Error in FQ_matrix_partition_assign: colind1 must have unique index values!")
    Err.Raise (Err.Number)
    Exit Function
End If

Mres = M1
For i = 1 To nrow2
    For j = 1 To ncol2
        Mres(rowind1(i), colind1(j)) = M2(i, j)
    Next j
Next i
FQ_matrix_partition_assign = Mres
End Function
'******************************************************************
' Quick sort function
' sources:
' http://stackoverflow.com/questions/152319/vba-array-sort-function
' http://en.allexperts.com/q/Visual-Basic-1048/string-manipulation.htm
' Minor adjustments by Tunc A. Kütükcüoglu for producing index vector ind
' such that Vout = Vin(ind)
'******************************************************************
Sub QuickSort(vArray As Variant, IndArr As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  Dim tmpSwap As Variant, tmpSwapInd As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        tmpSwapInd = IndArr(tmpLow)     ' added by Tunc
        
        vArray(tmpLow) = vArray(tmpHi)
        IndArr(tmpLow) = IndArr(tmpHi)  ' added by Tunc
        
        vArray(tmpHi) = tmpSwap
        IndArr(tmpHi) = tmpSwapInd      ' added by Tunc
        
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, IndArr, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, IndArr, tmpLow, inHi

End Sub
'******************************************************************
' Sorts the elements of input vector Vin in ascending or descending
' order depending on SortOption
' - ind: index vector such that Vout = Vin(ind)
' - error if Vin is not a vector
' - assumption: minimum absolute difference between all element values:
'   MinElementDiff = 1/100000 (to save time)
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Sub FQ_vector_sort(Vin() As Double, Vout() As Double, ind() As Double, SortOpt As SortOption)
Attribute FQ_vector_sort.VB_Description = "Sorts the elements of input vector Vin in ascending or descending order."
Attribute FQ_vector_sort.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ElementCount As Long
Dim i As Long, j As Long
Dim MinElementDiff As Double, Vs() As Double, Delta As Double
' minimum difference
MinElementDiff = 0.000001
' Check if vector
If Not FQ_CheckIfVector(Vin) Then
    FQ_MessageBox ("Error in FQ_vector_sort: Input argument Vin must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Sub
End If
' get length of vector
ElementCount = UBound(Vin)
If ElementCount = 1 Then
    Vout = Vin
    ReDim ind(1 To 1)
    ind(1) = 1
    Exit Sub
End If
' initiate vectors
Vs = Vin

' adjust values of vector elements (without changing orders) so that
' elements of equal values are sorted properly:
' V = ([1 1 1]) --> ind = [1 2 3] for ascending order
Delta = MinElementDiff / ElementCount
For i = 1 To ElementCount
    Vs(i) = Vs(i) + i * Delta
Next i
ind = FQ_vector_sequence(1, 1, ElementCount)
Call QuickSort(Vs, ind, inLow:=1, inHi:=ElementCount)
Select Case SortOpt
    Case nAscending
        ' do nothing
    Case nDescending
        ind = FQ_vector_reverse(ind)
    Case Else
        FQ_MessageBox ("Error in FQ_vector_sort: Invalid sort option!")
        Err.Raise (FQ_ErrorNum)
        Exit Sub
End Select
Vout = FQ_vector_partition(Vin, ind)
End Sub
'******************************************************************
' Sorted output vector Vout contains distinct (unique) elements of the
' input vector Vin in ascending order.
' - ind: index vector such that Vout = Vin(ind)
' - error if Vin is not a vector
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Sub FQ_vector_unique(Vin() As Double, Vout() As Double, ind() As Double)
Attribute FQ_vector_unique.VB_Description = "Sorted output vector Vout contains distinct (unique) elements of the input vector Vin in ascending order."
Attribute FQ_vector_unique.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ElementCount As Long, i As Long, j As Long
Dim Vsort() As Double, Vind() As Double
' Check if vector
If Not FQ_CheckIfVector(Vin) Then
    FQ_MessageBox ("Error in FQ_vector_unique: Input argument Vin must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Sub
End If
' get length of vector
ElementCount = UBound(Vin)
' initiate vectors
Vsort = Vin
Vind = FQ_vector_sequence(1, 1, ElementCount)
Call QuickSort(Vsort, Vind, inLow:=1, inHi:=ElementCount)

' get unique vector
ReDim Vout(1 To ElementCount)
ReDim ind(1 To ElementCount)
Vout(1) = Vsort(1)
ind(1) = Vind(1)

j = 1
For i = 2 To ElementCount
    If Vsort(i) <> Vsort(i - 1) Then
        j = j + 1
        Vout(j) = Vsort(i)
        ind(j) = Vind(i)
    End If
Next i
ReDim Preserve Vout(1 To j)     ' unique vector
ReDim Preserve ind(1 To j)      ' Vout = Vin(ind)
End Sub
'******************************************************************
' Returns True if vector Vin contains distinct (unique) element values;
' otherwise returns false
' - error if Vin is not a vector
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_vector_if_unique(V() As Double) As Boolean
Attribute FQ_vector_if_unique.VB_Description = "Returns True if vector Vin contains distinct (unique) element values; otherwise returns false."
Attribute FQ_vector_if_unique.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ElementCount As Long, i As Long, j As Long
' Check if vector
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_vector_if_unique: Input argument V must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' get length of vector
ElementCount = UBound(V)
FQ_vector_if_unique = True
For i = 1 To ElementCount - 1
    For j = i + 1 To ElementCount
        If V(i) = V(j) Then
            FQ_vector_if_unique = False
            Exit Function
        End If
    Next j
Next i
End Function
'******************************************************************
' Finds the positions (indices) of vector elements that satisfy the
' search criterion.
' Returns IfFound = True if there is at least a single match;
' otherwise returns false and empty index vector ind.
' - Example: Return all indices ind of vector elements such that
'   V(i) >= SearchValue
' - ComparisonOperator is a string from set {'=',  '<>', '>', '>=', '<', '<='}
' - error if V is not a vector
' - error of undefined comparison operator not in the set above
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Sub FQ_vector_find_elements(V() As Double, ComparisonOperator As String, _
    SearchValue As Double, ind() As Double, IfFound As Boolean)
Attribute FQ_vector_find_elements.VB_Description = "Finds the positions (indices) of vector elements that satisfy the search criterion."
Attribute FQ_vector_find_elements.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i As Long, j As Long, ElementCount As Long, EmptyArr() As Double
' Check if vector
If Not FQ_CheckIfVector(V) Then
    FQ_MessageBox ("Error in FQ_vector_find_elements: Input argument V must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Sub
End If
'Get vector size
ElementCount = UBound(V)
ReDim ind(1 To ElementCount)

j = 0
IfFound = False

For i = 1 To ElementCount
    Select Case ComparisonOperator
        Case "="
            If V(i) = SearchValue Then
                j = j + 1
                ind(j) = i
            End If
            
        Case "<>"
            If V(i) <> SearchValue Then
                j = j + 1
                ind(j) = i
            End If
            
        Case ">"
            If V(i) > SearchValue Then
                j = j + 1
                ind(j) = i
            End If
            
         Case "<"
            If V(i) < SearchValue Then
                j = j + 1
                ind(j) = i
            End If
            
        Case ">="
            If V(i) >= SearchValue Then
                j = j + 1
                ind(j) = i
            End If
            
        Case "<="
            If V(i) <= SearchValue Then
                j = j + 1
                ind(j) = i
            End If
            
        Case Else
            FQ_MessageBox ("Error in FQ_vector_find_elements: Invalid comparison operator!")
            Err.Raise (FQ_ErrorNum)
            Exit Sub
    End Select
Next i

If j > 0 Then
    IfFound = True
    ReDim Preserve ind(1 To j)
    Else
    ind = EmptyArr
End If
End Sub
'******************************************************************
' Finds the positions (row and column indices) of matrix elements that
' satisfy search criterion
' Matching indices are returned with:
'   a) vector index Vind of RowByRow serialized matrix, such that
'      a matching element M(i,j) maps to V(MatchCount) = nrow * (i-1) + j
'   b) matching matrix Mmatch, such that Mmatch(i,j) = 1 for a match; otherwise 0
' Returns IfFound = True if there is at least a single match; othewise False.
' otherwise returns false.
' - Example: Return all row/col indices Mind of matrix elements such that
'   M(i,j) >= SearchValue
' - ComparisonOperator is a string from set {'=',  '<>', '>', '>=', '<', '<='}
' - error if M is not a matrix
' - error of undefined comparison operator not in the set above
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Sub FQ_matrix_find_elements(M() As Double, ComparisonOperator As String, _
    SearchValue As Double, Vind() As Double, _
    Mmatch() As Double, IfFound As Boolean)
Attribute FQ_matrix_find_elements.VB_Description = "Finds the positions (row and column indices) of matrix elements that satisfy search criterion."
Attribute FQ_matrix_find_elements.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i As Long, j As Long, nrow As Long, ncol As Long, MatchCount As Long
Dim EmptyArr() As Double
' Check if matrix
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_matrix_find_elements: Input argument M must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Sub
End If
' get row and column sizes
nrow = UBound(M, 1)
ncol = UBound(M, 2)
ReDim Mmatch(1 To nrow, 1 To ncol)
ReDim Vind(1 To (nrow * ncol))
MatchCount = 0

' find matching elements
For i = 1 To nrow
    For j = 1 To ncol
        Select Case ComparisonOperator
            Case "="
                If M(i, j) = SearchValue Then
                    MatchCount = MatchCount + 1
                    Vind(MatchCount) = nrow * (i - 1) + j
                    Mmatch(i, j) = 1
                End If
                
            Case "<>"
                If M(i, j) <> SearchValue Then
                    MatchCount = MatchCount + 1
                    Vind(MatchCount) = nrow * (i - 1) + j
                    Mmatch(i, j) = 1
                End If
                
            Case ">"
                If M(i, j) > SearchValue Then
                    MatchCount = MatchCount + 1
                    Vind(MatchCount) = nrow * (i - 1) + j
                    Mmatch(i, j) = 1
                End If
                
             Case "<"
                If M(i, j) < SearchValue Then
                    MatchCount = MatchCount + 1
                    Vind(MatchCount) = nrow * (i - 1) + j
                    Mmatch(i, j) = 1
                End If
                
            Case ">="
                If M(i, j) >= SearchValue Then
                    MatchCount = MatchCount + 1
                    Vind(MatchCount) = nrow * (i - 1) + j
                    Mmatch(i, j) = 1
                End If
                
            Case "<="
                If M(i, j) <= SearchValue Then
                    MatchCount = MatchCount + 1
                    Vind(MatchCount) = nrow * (i - 1) + j
                    Mmatch(i, j) = 1
                End If
                
            Case Else
                FQ_MessageBox ("Error in FQ_matrix_find_elements: Invalid comparison operator!")
                Err.Raise (FQ_ErrorNum)
                Exit Sub
        End Select
    Next j
Next i

If MatchCount > 0 Then
    IfFound = True
    ReDim Preserve Vind(1 To MatchCount)
    Else
    Vind = EmptyArr
End If
End Sub
'******************************************************************
' Find vector indices such that V1 = V2(ind)
' V2containsV1 = True: V2 contains all element values in V1; i.e.
'   V1 is a subset of V2
' - returned vector ind can be an empty vector
' - index value is 0 for an element of V1 not found in V2, for example:
'   V1 = [1 2 3], V2 = [5 2 3 6] --> ind = [0 1 2]
' - error if V1 and/or V2 is not a vector
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Sub FQ_vector_find_indices(V1() As Double, V2() As Double, ind() As Double, _
    V2containsV1 As Boolean)
Attribute FQ_vector_find_indices.VB_Description = "Find vector indices such that V1 = V2(ind); V2containsV1 = True: V2 contains all element values in V1; i.e. V1 is a subset of V2."
Attribute FQ_vector_find_indices.VB_ProcData.VB_Invoke_Func = " \n14"
Dim vlen1 As Long, vlen2 As Long, i As Long, j As Long
Dim Vind() As Double, ElementFound As Boolean
' Check if vector
If Not FQ_CheckIfVector(V1) Then
    FQ_MessageBox ("Error in FQ_vector_find_indices: Input argument V1 must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Sub
End If
If Not FQ_CheckIfVector(V2) Then
    FQ_MessageBox ("Error in FQ_vector_find_indices: Input argument V2 must be a vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Sub
End If
' get vector sizes
vlen1 = UBound(V1, 1)
vlen2 = UBound(V2, 1)
ReDim Vind(1 To vlen1)
V2containsV1 = True

For i = 1 To vlen1
    ElementFound = False
    For j = 1 To vlen2
        If V1(i) = V2(j) Then
            ElementFound = True
            Vind(i) = j
            Exit For
        End If
    Next j
    If ElementFound = False Then
        Vind(i) = 0
        V2containsV1 = False
    End If
Next i
ind = Vind
End Sub
'******************************************************************
' more intuitive alias name for FQ_vector_find_indices
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Sub FQ_vector_map_to_set_elements(V1() As Double, Vset() As Double, ind() As Double, _
    VsetContainsV1 As Boolean)
Attribute FQ_vector_map_to_set_elements.VB_Description = "More intuitive alias name for FQ_vector_find_indices."
Attribute FQ_vector_map_to_set_elements.VB_ProcData.VB_Invoke_Func = " \n14"
Call FQ_vector_find_indices(V1, Vset, ind, VsetContainsV1)
End Sub
'******************************************************************
' Sort rows of matrix in either ascending or descending order,
' w.r.t. given column indices with vector ColInd
' ColInd = [1 3 -2] means order by 1., 3. and 2. columns, 2. in descending order
' - returns empty matrix Ms if M1 is empty
' - returns row indices with RowInd, such that Ms = M1(RowInd, :)
' - error if the absolute value of a column index in ColInd is not a
'   positive integer between 1 and column size of M1
' - error if abs(ColInd) is not a unique vector
' - error if M1 is not a matrix
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Sub FQ_matrix_sort(M1() As Double, Colind() As Double, Ms() As Double, _
    RowInd() As Double)
Attribute FQ_matrix_sort.VB_Description = "Sort rows of matrix in either ascending or descending order."
Attribute FQ_matrix_sort.VB_ProcData.VB_Invoke_Func = " \n14"
Dim EmptyArr() As Double, nrow As Long, ncol As Long, i As Long, j As Long
Dim AbsColInd() As Double, LenColInd As Long, Msub() As Double
Dim Vcol() As Double, VcolSorted() As Double, Vind() As Double
' check if empty matrix
If FQ_ArrayDimension(M1) = 0 Then
    Ms = EmptyArr
    RowInd = EmptyArr
    Exit Sub
End If
' Check if matrix
If Not FQ_CheckIfMatrix(M1) Then
    FQ_MessageBox ("Error in FQ_matrix_sort: Input argument M1 must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Sub
End If
' Check if single-row matrix (shortcut)
If UBound(M1, 1) = 1 Then
    Ms = M1
    ReDim RowInd(1 To 1)
    RowInd(1) = 1
End If
' check column indices
AbsColInd = FQ_vector_operation(Colind, "abs")
LenColInd = UBound(Colind, 1)
If Not FQ_check_index_values(AbsColInd, LenColInd, 1) Then
    FQ_MessageBox ("Error in FQ_matrix_sort: Improper column indices in ColInd!")
    Err.Raise (FQ_ErrorNum)
    Exit Sub
End If
' check if unique
If Not FQ_vector_if_unique(AbsColInd) Then
    FQ_MessageBox ("Error in FQ_matrix_sort: ABS(ColInd) must be a unique vector!")
    Err.Raise (FQ_ErrorNum)
    Exit Sub
End If
' get row and column sizes
nrow = UBound(M1, 1)
ncol = UBound(M1, 2)
RowInd = FQ_vector_sequence(1, 1, nrow)
ReDim Vcol(1 To nrow)

' begin sorting
Msub = FQ_matrix_partition(M1, Colind:=AbsColInd)
For i = 1 To LenColInd
    ' read column of Msub into a vector
    For j = 1 To nrow
        Vcol(j) = Msub(j, LenColInd - i + 1)
    Next j
    ' sort vector
    If Colind(LenColInd - i + 1) > 0 Then
        Call FQ_vector_sort(Vcol, VcolSorted, Vind, nAscending)
    Else
        Call FQ_vector_sort(Vcol, VcolSorted, Vind, nDescending)
    End If
    RowInd = FQ_vector_partition(RowInd, Vind)
    Msub = FQ_matrix_partition(Msub, RowInd:=Vind)
Next i
Ms = FQ_matrix_partition(M1, RowInd:=RowInd)
End Sub
'******************************************************************
' Creates an embedded chart on the given sheet
' type "MyChart.ChartType =" in VBE to see the available chart types
' - default values for size parameters nLeft, nTop, nWidth, nHeight:
'       50, 50, 300, 200
' - error if a worksheet with the given name doesnot exist
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Sub FQ_create_embedded_chart(SheetName As String, DataRange As Range, _
    Optional nLeft As Variant, Optional nTop As Variant, _
    Optional nWidth As Variant, Optional nHeight As Variant, _
    Optional ChartType As Variant)
Attribute FQ_create_embedded_chart.VB_Description = "Creates an embedded chart on the given sheet."
Attribute FQ_create_embedded_chart.VB_ProcData.VB_Invoke_Func = " \n14"
Dim MyChart As Chart
Dim xLeft As Long, xTop As Long, xWidth As Long, xHeight As Long
' set default values
If IsMissing(nLeft) Then
    xLeft = 50
End If
If IsMissing(nTop) Then
    xTop = 50
End If
If IsMissing(nWidth) Then
    xWidth = 300
End If
If IsMissing(nHeight) Then
    xHeight = 200
End If

On Error GoTo EH1
Set MyChart = ThisWorkbook.Sheets(SheetName).ChartObjects. _
    Add(CDbl(xLeft), CDbl(xTop), CDbl(xWidth), CDbl(xHeight)).Chart
On Error GoTo EH2
MyChart.SetSourceData Source:=DataRange
If Not IsMissing(ChartType) Then
    On Error GoTo EH3
    MyChart.ChartType = ChartType
End If
Exit Sub
EH1:
FQ_MessageBox ("Error in FQ_create_embedded_chart: Chart could not be created! Check the size parameters, and check if a worksheet with the given name exists.")
Exit Sub
EH2:
FQ_MessageBox ("Error in FQ_create_embedded_chart: Soource data for the chart could not be set! Check the range for source data.")
Exit Sub
EH3:
FQ_MessageBox ("Error in FQ_create_embedded_chart: Invalid chart type!")
End Sub
'******************************************************************
' Converts the 1 or 2 dimensional variant array into a string array
' - error if array dimension arrdim = 0 or arrdim > 2
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_var_to_str(Arr As Variant) As String()
Attribute FQ_var_to_str.VB_Description = "Converts the 1 or 2 dimensional variant array into a string array."
Attribute FQ_var_to_str.VB_ProcData.VB_Invoke_Func = " \n14"
Dim arrdim As Byte, i As Long, j As Long, str() As String, x As Double
On Error GoTo EH1
' check dimension
arrdim = FQ_ArrayDimension(Arr)
If arrdim = 0 Or arrdim > 2 Then
    FQ_MessageBox ("Error in FQ_var_to_str: Improper array dimension " & arrdim)
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
If arrdim = 1 Then
    ReDim str(1 To (UBound(Arr) - LBound(Arr) + 1))
    For i = LBound(Arr) To UBound(Arr)
        str(i - LBound(Arr) + 1) = CStr(Arr(i))
    Next i
End If
If arrdim = 2 Then
    ReDim str(1 To (UBound(Arr, 1) - LBound(Arr, 1) + 1), 1 To (UBound(Arr, 2) - LBound(Arr, 2) + 1))
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        For j = LBound(Arr, 2) To UBound(Arr, 2)
        str(i - LBound(Arr, 1) + 1, j - LBound(Arr, 2) + 1) = CStr(Arr(i, j))
        Next j
    Next i
End If
FQ_var_to_str = str
Exit Function
EH1:
Debug.Print "in EH1"
FQ_MessageBox ("Error in FQ_var_to_str: " & Err.Number & " - " & Err.Description)
    Err.Raise (Err.Number)
End Function
'******************************************************************
' Creates a mixed table with variant array as data set
' - returns empty table if there was an error
' - error if the size of DataSet doesn't match given nrows and ncols
' - error if other arrays row or column names, row or column
'   descriptions doesn't match corresponding nrows/
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_create_data_table(DataSet As Variant, nrows As Long, ncols As Long, _
    row_names() As String, col_names() As String, row_desc() As String, _
    col_desc() As String) As DataTable
Attribute FQ_create_data_table.VB_Description = "Creates a mixed table with variant array as data set."
Attribute FQ_create_data_table.VB_ProcData.VB_Invoke_Func = " \n14"
Dim dt As DataTable, EmptyDt As DataTable
'On Error GoTo EH1
' check DataSet size
If FQ_ArrayDimension(DataSet) <> 2 Then
    FQ_MessageBox ("Error in FQ_create_data_table: DataSet must be a 2-dimensional array!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
ElseIf UBound(DataSet, 1) <> nrows Or UBound(DataSet, 2) <> ncols Then
    FQ_MessageBox ("Error in create_mixed_table: DataSet size does not match given nrows and/or ncols!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' check array sizes
If Not FQ_CheckIfEmptyArray(row_names) And FQ_ArrayDimension(row_names) <> 1 Then
    FQ_MessageBox ("Error in FQ_create_data_table: row_names must be an empty or 1-dimensional array!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
ElseIf FQ_CheckIfEmptyArray(row_names) Then
    ' do nothing
ElseIf UBound(row_names) <> nrows Then
    FQ_MessageBox ("Error in FQ_create_data_table: Size of string array row_names is neither 0 or nrows!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
If Not FQ_CheckIfEmptyArray(col_names) And FQ_ArrayDimension(col_names) <> 1 Then
    FQ_MessageBox ("Error in FQ_create_data_table: col_names must be an empty or 1-dimensional array!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
ElseIf FQ_CheckIfEmptyArray(col_names) Then
    ' do nothing
ElseIf UBound(col_names) <> ncols Then
    FQ_MessageBox ("Error in FQ_create_data_table: Size of string array col_names is neither 0 or ncols!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
If Not FQ_CheckIfEmptyArray(row_desc) And FQ_ArrayDimension(row_desc) <> 1 Then
    FQ_MessageBox ("Error in FQ_create_data_table: row_desc must be an empty or 1-dimensional array!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
ElseIf FQ_CheckIfEmptyArray(row_desc) Then
    ' do nothing
ElseIf UBound(row_desc) <> nrows Then
    FQ_MessageBox ("Error in FQ_create_data_table: Size of string array row_desc is neither 0 or nrows!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
If Not FQ_CheckIfEmptyArray(col_desc) And FQ_ArrayDimension(col_desc) <> 1 Then
    FQ_MessageBox ("Error in FQ_create_data_table: row_desc must be an empty or 1-dimensional array!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
ElseIf FQ_CheckIfEmptyArray(col_desc) Then
    ' do nothing
ElseIf UBound(col_desc) <> ncols Then
    FQ_MessageBox ("Error in FQ_create_data_table: Size of string array col_desc is neither 0 or ncols!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' checks OK, create table
dt.nrows = nrows
dt.ncols = ncols
dt.DataSet = DataSet
dt.row_names = row_names
dt.column_names = col_names
dt.row_descriptions = row_desc
dt.column_descriptions = col_desc
FQ_create_data_table = dt
Exit Function
EH1:
FQ_MessageBox ("Error in FQ_create_data_table: " & Err.Number & " - " & Err.Description)
FQ_create_data_table = EmptyDt
End Function
'******************************************************************
' Checks the consistence of a data table; returns true if it is consistent
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_check_if_table_consistent(DTable As DataTable) As Boolean
Attribute FQ_check_if_table_consistent.VB_Description = "Checks the consistence of a data table; returns true if it is consistent."
Attribute FQ_check_if_table_consistent.VB_ProcData.VB_Invoke_Func = " \n14"
Dim dt As DataTable, EmptyDt As DataTable
Dim nrows As Long, ncols As Long, row_names() As String, col_names() As String
Dim row_desc() As String, col_desc() As String, DataSet As Variant
On Error GoTo EH1
DataSet = DTable.DataSet
nrows = DTable.nrows
ncols = DTable.ncols
row_names = DTable.row_names
row_desc = DTable.row_descriptions
col_names = DTable.column_names
col_desc = DTable.column_descriptions

FQ_check_if_table_consistent = True
' check DataSet size
If FQ_ArrayDimension(DataSet) <> 2 Then
    FQ_check_if_table_consistent = False
    Exit Function
ElseIf UBound(DataSet, 1) <> nrows Or UBound(DataSet, 2) <> ncols Then
    FQ_check_if_table_consistent = False
    Exit Function
End If
' check array sizes
If Not FQ_CheckIfEmptyArray(row_names) And FQ_ArrayDimension(row_names) <> 1 Then
    FQ_check_if_table_consistent = False
    Exit Function
ElseIf FQ_CheckIfEmptyArray(row_names) Then
    ' do nothing
ElseIf UBound(row_names) <> nrows Then
    FQ_check_if_table_consistent = False
    Exit Function
End If
If Not FQ_CheckIfEmptyArray(col_names) And FQ_ArrayDimension(col_names) <> 1 Then
    FQ_check_if_table_consistent = False
    Exit Function
ElseIf FQ_CheckIfEmptyArray(col_names) Then
    ' do nothing
ElseIf UBound(col_names) <> ncols Then
    FQ_check_if_table_consistent = False
    Exit Function
End If
If Not FQ_CheckIfEmptyArray(row_desc) And FQ_ArrayDimension(row_desc) <> 1 Then
    FQ_check_if_table_consistent = False
    Exit Function
ElseIf FQ_CheckIfEmptyArray(row_desc) Then
    ' do nothing
ElseIf UBound(row_desc) <> nrows Then
    FQ_check_if_table_consistent = False
    Exit Function
End If
If Not FQ_CheckIfEmptyArray(col_desc) And FQ_ArrayDimension(col_desc) <> 1 Then
    FQ_check_if_table_consistent = False
    Exit Function
ElseIf FQ_CheckIfEmptyArray(col_desc) Then
    ' do nothing
ElseIf UBound(col_desc) <> ncols Then
    FQ_check_if_table_consistent = False
    Exit Function
End If
Exit Function
EH1:
FQ_MessageBox ("Error in FQ_check_if_consistent_table: " & Err.Number & " - " & Err.Description)
FQ_check_if_table_consistent = False
End Function
'******************************************************************
' Adds a comment to given cell (upper-left cell of a range)
' - overwrites an existing comment
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Sub FQ_add_comment_to_cell(Rn As Range, Comment As String)
Attribute FQ_add_comment_to_cell.VB_Description = "Adds a comment to given cell (upper-left cell of a range)."
Attribute FQ_add_comment_to_cell.VB_ProcData.VB_Invoke_Func = " \n14"
Dim UpperLeft As Range
On Error GoTo EH1
Set UpperLeft = Range(Rn.Cells(1, 1), Rn.Cells(1, 1))
If Not UpperLeft.Comment Is Nothing Then
    UpperLeft.Comment.Delete
End If
UpperLeft.AddComment (Comment)
Exit Sub
EH1:
FQ_MessageBox ("Error in FQ_add_comment_to_cell: " & Err.Number & " - " & Err.Description)
Err.Raise (Err.Number)
End Sub
'******************************************************************
' Inserts table into the given range, starting from top-left cell
' - Inserts table only if it is consistent
' - Adds comments to row/column names if IfAddComments = True
' - Checks consistency of description arrays even if IfAddComments = False
' - error if size of DataSet doesnot match nrows/ncols
' - error if other arrays row or column names, row or column
'   descriptions doesn't match corresponding nrows/ncols
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Sub FQ_table_to_range(DTable As DataTable, Rn As Range, IfAddComments As Boolean)
Attribute FQ_table_to_range.VB_Description = "Inserts table into the given range, starting from top-left cell."
Attribute FQ_table_to_range.VB_ProcData.VB_Invoke_Func = " \n14"
Dim RangeRowNames As Range, RangeColNames As Range, RangeDataSet As Range
Dim nrows As Long, ncols As Long, MyCell As Range
Dim IfColNamesExist As Boolean, IfRowNamesExist
Dim dt As DataTable, EmptyDt As DataTable, i As Long
Dim HorizontalShift As Long, VerticalShift As Long
' check if table is consistent
If Not FQ_check_if_table_consistent(DTable) Then
    FQ_MessageBox ("Error in FQ_table_to_range: Inconsistent data table! Check sizes of table arrays compared to nrows and ncols")
    Err.Raise (FQ_ErrorNum)
    Exit Sub
End If
'On Error GoTo EH1
nrows = DTable.nrows
ncols = DTable.ncols
If FQ_CheckIfEmptyArray(DTable.row_names) Then
    IfRowNamesExist = False
    HorizontalShift = 0
Else
    IfRowNamesExist = True
    Set RangeRowNames = Range(Rn.Cells(2, 1), Rn.Cells(nrows + 1, 1))
    HorizontalShift = 1
End If
If FQ_CheckIfEmptyArray(DTable.column_names) Then
    IfColNamesExist = False
    VerticalShift = 0
Else
    IfColNamesExist = True
    Set RangeColNames = Range(Rn.Cells(1, 1 + HorizontalShift), _
        Rn.Cells(1, ncols + HorizontalShift))
    VerticalShift = 1
End If
Set RangeDataSet = Range(Rn.Cells(1 + VerticalShift, 1 + HorizontalShift), _
        Rn.Cells(nrows + VerticalShift, ncols + HorizontalShift))
        
' checks OK, insert table into range
If IfRowNamesExist Then
    RangeRowNames.Value = Application.WorksheetFunction.Transpose(DTable.row_names)
End If
If IfColNamesExist Then
    RangeColNames.Value = DTable.column_names
End If
RangeDataSet.Value = DTable.DataSet
' add comments
If IfAddComments And IfRowNamesExist And FQ_ArrayDimension(DTable.row_descriptions) = 1 Then
    For i = 1 To nrows
        If Len(DTable.row_descriptions(i)) > 0 Then
            Call FQ_add_comment_to_cell(RangeRowNames.Cells(i, 1), _
                DTable.row_descriptions(i))
        End If
    Next
End If
If IfAddComments And IfColNamesExist And FQ_ArrayDimension(DTable.column_descriptions) = 1 Then
    For i = 1 To ncols
        If Len(DTable.column_descriptions(i)) > 0 Then
            Call FQ_add_comment_to_cell(RangeColNames.Cells(1, i), _
                DTable.column_descriptions(i))
        End If
    Next
End If
Exit Sub
EH1:
FQ_MessageBox ("Error in FQ_table_to_range: " & Err.Number & " - " & Err.Description)
Err.Raise (Err.Number)
End Sub
'******************************************************************
' Determinant of a square matrix
' - error if M is not a matrix
' Author: Finaquant Analytics Ltd. - www.finaquant.com
'******************************************************************
Function FQ_determinant(M() As Double) As Double
Dim s As Long, k As Long, p As Long, i As Long, j As Long
Dim rows As Long, cols As Long
Dim Mx() As Double, sv As Double, sk As Double, st As Double
' Check if matrix
If Not FQ_CheckIfMatrix(M) Then
    FQ_MessageBox ("Error in FQ_determinant: Input argument M must be a matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If
' Check if square matrix
If UBound(M, 1) <> UBound(M, 2) Then
    FQ_MessageBox ("Error in FQ_determinant: Input argument M must be a NxN square matrix!")
    Err.Raise (FQ_ErrorNum)
    Exit Function
End If

rows = UBound(M, 1)
' copy matrix M to Mx
ReDim Mx(1 To rows, 1 To rows)
For i = 1 To rows
    For j = 1 To rows
        Mx(i, j) = M(i, j)
    Next j
Next i

st = 1      ' temp determinant

For k = 1 To rows
    If Mx(k, k) = 0 Then
        j = k
        While j < rows And Mx(k, j) = 0
            j = j + 1
        Wend
        If Mx(k, j) = 0 Then
            FQ_determinant = 0
            Exit Function
        Else
            For i = k To rows
                sv = Mx(i, j)
                Mx(i, j) = Mx(i, k)
                Mx(i, k) = sv
            Next i
        End If
        st = -st
    End If
    
    sk = Mx(k, k)
    st = st * sk
    
    If k < rows Then
        p = k + 1
        For i = p To rows
            For j = p To rows
                Mx(i, j) = Mx(i, j) - Mx(i, k) * (Mx(k, j) / sk)
            Next j
        Next i
    End If
Next k
FQ_determinant = st
End Function

' Copyrights © 2012 - finaquant.com (Tunc A. Kutukcuoglu)
' Website: http://finaquant.com/
Sub SetProcedureInfo()
Application.MacroOptions Macro:="FQ_MessageBox", _
Description:="Customized message box; prints to immediate window if the global parameter FQ_DEBUG is true."

Application.MacroOptions Macro:="FQ_ArrayDimension", _
Description:="Get array dimension; returns the  number of dimensions in an array."

Application.MacroOptions Macro:="FQ_ArrayDimension", _
Description:="Check if array; returns True if the argument is an array; otherwise False."

Application.MacroOptions Macro:="FQ_CheckIfEmptyArray", _
Description:="Check if empty array; returns True if array is empty."

Application.MacroOptions Macro:="FQ_CheckIfMatrix", _
Description:="Check if matrix; returns True if the argument is an matrix; otherwise False."

Application.MacroOptions Macro:="FQ_CheckIfVector", _
Description:="Check if vector; returns True if the argument is a vector; otherwise False."

Application.MacroOptions Macro:="FQ_var_to_matrix", _
Description:="Converts a variant array with numeric elements into a matrix."

Application.MacroOptions Macro:="FQ_var_to_vector", _
Description:="Converts a variant array with numeric elements into a vector."

Application.MacroOptions Macro:="FQ_matrix_format", _
Description:="Converts a matrix into a printable formatted string."

Application.MacroOptions Macro:="FQ_vector_format", _
Description:="Converts a vector into a printable formatted string."

Application.MacroOptions Macro:="FQ_1DimMatrix_to_Vector", _
Description:="Converts a 1-dimensional (1xN or Nx1) matrix to a vector."

Application.MacroOptions Macro:="FQ_Vector_to_1DimMatrix", _
Description:="Converts Vector to a 1-dimensional matrix, either vertical or horizontal depending on matrix alignment argument."

Application.MacroOptions Macro:="FQ_matrix_transpose", _
Description:="Transpose matrix; M2 = transpose(M1)."

Application.MacroOptions Macro:="FQ_matrix_inverse", _
Description:="Inverse matrix: Y = inv(M)."

Application.MacroOptions Macro:="FQ_range_to_variant", _
Description:="Writes the numeric values of a worksheet range into a 2-dimensional NxM variant array."

Application.MacroOptions Macro:="FQ_variant_to_range", _
Description:="Writes the values of a 2-dimensional variant array into a range in excel, starting from the upper left corner of the range (Cells(1,1))."

Application.MacroOptions Macro:="FQ_matrix_to_range", _
Description:=" Writes the values of a matrix (2-dim double) into a worksheet range in excel, starting from the upper left corner of the range."

Application.MacroOptions Macro:="FQ_vector_to_range", _
Description:="Writes the values of a vector (1-dim double) into a range in excel, starting from the upper left corner of the range."

Application.MacroOptions Macro:="FQ_range_to_matrix", _
Description:="Reads a numeric range and converts it into a matrix with the same row and column size."

Application.MacroOptions Macro:="FQ_range_to_vector", _
Description:="Reads a numeric range row by row and writes the values into a vector."

Application.MacroOptions Macro:="FQ_vector_to_matrix", _
Description:="Writes the values of a vector into a matrix either row by row, or column by column."

Application.MacroOptions Macro:="FQ_matrix_to_vector", _
Description:="Reads the elements of a matrix either row by row, or column by column, and writes their values into a vector."

Application.MacroOptions Macro:="FQ_matrix_create", _
Description:="Creates matrix with sequential element values with given row and column sizes. Fills matrix row-wise with sequential numbers."

Application.MacroOptions Macro:="FQ_matrix_rand", _
Description:="Creates matrix with random element values between 0 and 1 with given row and column sizes. Fills matrix row-wise with random numbers."

Application.MacroOptions Macro:="FQ_vector_sequence", _
Description:="Creates a vector with given length (ElementCount), start value and interval between subsequent elements."

Application.MacroOptions Macro:="FQ_vector_rand", _
Description:="Creates a vector with random element values between 0 and 1 with given vector length (ElementCount)."

Application.MacroOptions Macro:="FQ_check_index_values", _
Description:="Check if all vector values are positive integers within limits."

Application.MacroOptions Macro:="FQ_matrix_partition", _
Description:="Returns partition of a matrix indicated by column and row index vectors."

Application.MacroOptions Macro:="FQ_vector_partition", _
Description:="Returns partition of a vector indicated index vector ind."

Application.MacroOptions Macro:="FQ_matrix_element_sum", _
Description:="Returns the sum of elements of matrix M, either row or column wise."

Application.MacroOptions Macro:="FQ_matrix_aggregate", _
Description:="Applies the given aggregation function (sum, min, max, avg, median) on the matrix."

Application.MacroOptions Macro:="FQ_matrix_scalar_add", _
Description:="Adds a scalar number to all elements of matrix."

Application.MacroOptions Macro:="FQS_matrix_scalar_add", _
Description:="Spreadsheet version of the function FQ_matrix_scalar_add: Adds a scalar number to all elements of matrix."

Application.MacroOptions Macro:="FQ_vector_scalar_add", _
Description:="Adds a scalar number to all elements of vector."

Application.MacroOptions Macro:="FQ_vector_scalar_multiply", _
Description:="Multiplies all elements of vector with a scalar number x."

Application.MacroOptions Macro:="FQ_vector_aggregate", _
Description:="Applies the given aggregation function (sum, min, max, avg, median) on the vector, and returns a scalar number."

Application.MacroOptions Macro:="FQ_matrix_element_count", _
Description:="Returns the total number of elements in matrix M."

Application.MacroOptions Macro:="FQ_matrix_scalar_multiply", _
Description:="Multiplies all elements of matrix with a scalar number x."

Application.MacroOptions Macro:="FQ_matrix_vector_multiply", _
Description:="Multiplies rows or columns of matrix M with corresponding elements of vector V."

Application.MacroOptions Macro:="FQ_matrix_matrix_sum", _
Description:="Adds up the elements of two matrices with identical row/column sizes."

Application.MacroOptions Macro:="FQ_vector_vector_sum", _
Description:="Adds up the elements of two vectors with identical lengths."

Application.MacroOptions Macro:="FQ_matrix_elementwise_multiply", _
Description:="Elementwise multiplication of two equal-sized matrices; R = M1 .* M2 (matlab notation)."

Application.MacroOptions Macro:="FQ_matrix_elementwise_divide", _
Description:="Elementwise division of two equal-sized matrices; R = M1 ./ M2 (matlab notation)."

Application.MacroOptions Macro:="FQ_vector_reverse", _
Description:="Reverse the order of vector elements; f.e. [1 4 3] --> [3 4 1]."

Application.MacroOptions Macro:="FQ_matrix_multiplication", _
Description:="Matrix multiplication in linear algebra,  C = A x B."

Application.MacroOptions Macro:="FQ_matrix_append", _
Description:="Appends matrix M2 to M1 either vertically or horizontally."

Application.MacroOptions Macro:="FQ_vector_vector_append", _
Description:="Appends vector V2 to V1 such that result vector = [V1, V2]."

Application.MacroOptions Macro:="FQ_matrix_operation", _
Description:="Applies the given mathematical operation like abs(), fix(), sin() etc. (all available single-argument VBA functions) on all elements of the matrix M."

Application.MacroOptions Macro:="FQ_vector_operation", _
Description:="Applies the given mathematical operation like abs(), fix(), sin() etc. (all available single-argument VBA functions) on all elements of the vector V."

Application.MacroOptions Macro:="FQ_vector_partition_assign", _
Description:="Assigns values of vector V2 to the partition of vector V1 selected by the index vector ind1. i.e. V1(ind1) = V2."

Application.MacroOptions Macro:="FQ_matrix_partition_assign", _
Description:="Assigns values of matrix M2 to the partition of matrix M1 selected by the index vectors rowind1 and colind1: M1(rowind1, colind1) = M2."

Application.MacroOptions Macro:="FQ_vector_sort", _
Description:="Sorts the elements of input vector Vin in ascending or descending order."

Application.MacroOptions Macro:="FQ_vector_unique", _
Description:="Sorted output vector Vout contains distinct (unique) elements of the input vector Vin in ascending order."

Application.MacroOptions Macro:="FQ_vector_if_unique", _
Description:="Returns True if vector Vin contains distinct (unique) element values; otherwise returns false."

Application.MacroOptions Macro:="FQ_vector_find_elements", _
Description:="Finds the positions (indices) of vector elements that satisfy the search criterion."

Application.MacroOptions Macro:="FQ_matrix_find_elements", _
Description:="Finds the positions (row and column indices) of matrix elements that satisfy search criterion."

Application.MacroOptions Macro:="FQ_vector_find_indices", _
Description:="Find vector indices such that V1 = V2(ind); V2containsV1 = True: V2 contains all element values in V1; i.e. V1 is a subset of V2."

Application.MacroOptions Macro:="FQ_vector_map_to_set_elements", _
Description:="More intuitive alias name for FQ_vector_find_indices."

Application.MacroOptions Macro:="FQ_matrix_sort", _
Description:="Sort rows of matrix in either ascending or descending order."

Application.MacroOptions Macro:="FQ_create_embedded_chart", _
Description:="Creates an embedded chart on the given sheet."

Application.MacroOptions Macro:="FQ_var_to_str", _
Description:="Converts the 1 or 2 dimensional variant array into a string array."

Application.MacroOptions Macro:="FQ_create_data_table", _
Description:="Creates a mixed table with variant array as data set."

Application.MacroOptions Macro:="FQ_check_if_table_consistent", _
Description:="Checks the consistence of a data table; returns true if it is consistent."

Application.MacroOptions Macro:="FQ_add_comment_to_cell", _
Description:="Adds a comment to given cell (upper-left cell of a range)."

Application.MacroOptions Macro:="FQ_table_to_range", _
Description:="Inserts table into the given range, starting from top-left cell."
End Sub
