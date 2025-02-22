Attribute VB_Name = "Int32Examples"
Option Explicit

Sub IntegerSizes()
  Dim N As LongLong
  Dim p As Integer
  
  N = 1
  For p = 1 To 64
    N = 2 ^ p - 1
    N = N + 1
  Next
  
End Sub
