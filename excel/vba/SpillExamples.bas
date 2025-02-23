Attribute VB_Name = "SpillExamples"
Option Explicit

Public Function SpillTest_NxM(r As Range) As Variant()

  Dim N As Integer
  Dim m As Integer
  Dim x() As Variant
  Dim i As Integer
  Dim j As Integer
  
  N = r.Rows.Count
  m = r.Columns.Count
  
  ReDim x(1 To N, 1 To m) As Variant
  
  For i = 1 To N
    For j = 1 To m
      x(i, j) = r.Cells(i, j).Value
    Next
  Next
  
  SpillTest_NxM = x
  
End Function


Public Function SpillTest_Nx1(r As Range, Optional down_columns As Boolean = True) As Variant

  Dim N As Integer
  Dim m As Integer
  Dim x() As Variant
  Dim i As Integer
  Dim j As Integer
  
  N = r.Rows.Count
  m = r.Columns.Count
  
  ReDim x(1 To N * m, 1 To 1) As Variant
  
  For i = 1 To N
    For j = 1 To m
      If down_columns Then
        x((j - 1) * m + i, 1) = r.Cells(i, j).Value
      Else
        x((i - 1) * N + j, 1) = r.Cells(i, j).Value
      End If
    Next
  Next
  
  SpillTest_Nx1 = x
  
End Function



Public Function SpillTest_1xM(r As Range, Optional down_columns As Boolean = True) As Variant

  Dim N As Integer
  Dim m As Integer
  Dim x() As Variant
  Dim i As Integer
  Dim j As Integer
  
  N = r.Rows.Count
  m = r.Columns.Count
  
  ReDim x(1 To 1, 1 To N * m) As Variant
  
  For i = 1 To N
    For j = 1 To m
      If down_columns Then
        x(1, (j - 1) * m + i) = r.Cells(i, j).Value
      Else
        x(1, (i - 1) * N + j) = r.Cells(i, j).Value
      End If
    Next
  Next
  
  SpillTest_1xM = x
  
End Function


