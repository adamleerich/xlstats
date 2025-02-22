Attribute VB_Name = "Utilities"
Option Explicit

Public Const two32 As LongLong = 2 ^ 32
Public Const two31 As LongLong = 2 ^ 31



Function CInt32__Scalar(x As LongLong) As Long
  ' Mimics C's two's complement overflow of signed Int32
  x = x Mod two32
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



