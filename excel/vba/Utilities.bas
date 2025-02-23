Attribute VB_Name = "Utilities"
Option Explicit
Option Private Module

Public Const two32 As LongLong = 2 ^ 32
Public Const two31 As LongLong = 2 ^ 31



Function CInt32__Scalar(x As LongLong) As Long
  ' Mimics C's two's complement overflow of signed Int32
  x = x Mod two32
  If Sgn(x) = -1 Then
    x = x + two32
  End If
  If x >= two31 Then
    x = x - two32
  End If
  CInt32__Scalar = CLng(x)
End Function



Function CInt32__Array(x() As LongLong) As Long()
  Dim i As Integer
  Dim result() As Long
  ReDim result(LBound(x) To UBound(x)) As Long
  For i = LBound(x) To UBound(x)
    result(i) = CInt32__Scalar(x(i))
  Next
  CInt32__Array = result
End Function



Sub CHECK__CInt32()
  
  Dim x(10 To 15) As LongLong
  Dim y() As Long
  Dim i As Integer
  
  x(10) = 129347812983#
  x(11) = -987917263
  x(12) = 19273981729381#
  x(13) = -9879879878768#
  x(14) = 1287
  x(15) = -98798576
  
  y = CInt32__Array(x)
  
  For i = LBound(y) To UBound(y)
    Debug.Print y(i)
  Next

End Sub



Function CDblArray__Range(r As Range) As Double()

  Dim M As Integer
  Dim N As Integer
  Dim i As Integer
  Dim j As Integer
  
  N = r.Rows.Count
  M = r.Columns.Count
  
  Dim x() As Double
  
  ReDim x(1 To N * M) As Double
  
  For i = 1 To N
    For j = 1 To M
      x((j - 1) * N + i) = CDbl(r.Cells(i, j).Value)
    Next
  Next
  
  CDblArray__Range = x
  
End Function



Public Function ShiftRight32(x As Long, s As Byte) As Long
  Dim result As Long
  
  ' s is a Byte and cannot be negative
  If s = 0 Then
    ShiftRight32 = x
    Exit Function
  End If
  
  If s >= 32 Then
    ShiftRight32 = 0
    Exit Function
  End If
  
  ' s is between 1 and 31
  ' First deal with MSB
  ' result = x
  ' If Sgn(x) = -1 Then
  '   result = (x And &H7FFFFFFF) \ 2 Or &H40000000
  '   s = s - 1
  ' End If
  
  ' Handle MSB only for negative values
  If x < 0 Then
    result = (x And &H7FFFFFFF) \ 2 Or &H40000000
    s = s - 1
  Else
    result = x
  End If
  
  
  Do While s > 0
    result = result \ 2
    s = s - 1
  Loop
  
  ShiftRight32 = result
  
End Function


Public Function ShiftLeft32(x As Long, s As Byte) As Long
  ' https://www.excely.com/excel-vba/bit-shifting-function.shtml
  Dim result As Long
  Dim M As Long ' placeholder for bit in 31st position
  
  If s = 0 Then
    ShiftLeft32 = x
    Exit Function
  End If
  
  If s >= 32 Then
    ShiftLeft32 = 0
    Exit Function
  End If
  
  result = x
  Do While s > 0
    M = result And &H40000000
    result = (result And &H3FFFFFFF) * 2
    If M <> 0 Then
      result = result Or &H80000000
    End If
    s = s - 1
  Loop
  
  ShiftLeft32 = result
  
End Function




