Attribute VB_Name = "SpillExamples"
Option Explicit

Public Function SpillTest_NxM(r As Range) As Long()

  Dim N As Long       ' 1048576  = 2^20 rows max
  Dim M As Integer    ' 16384    = 2^14 columns max
  Dim x() As Long
  Dim i As Long
  Dim j As Integer
  
  N = r.Rows.Count
  M = r.Columns.Count
  
  ReDim x(1 To N, 1 To M) As Long
  
  For i = 1 To N
    For j = 1 To M
      x(i, j) = r.Cells(i, j).Value
    Next
  Next
  
  SpillTest_NxM = x
  
End Function


Public Function SpillTest_Nx1(r As Range, Optional down_columns As Boolean = True) As Variant

  Dim N As Integer
  Dim M As Integer
  Dim x() As Variant
  Dim i As Integer
  Dim j As Integer
  
  N = r.Rows.Count
  M = r.Columns.Count
  
  ReDim x(1 To N * M, 1 To 1) As Variant
  
  For i = 1 To N
    For j = 1 To M
      If down_columns Then
        x((j - 1) * M + i, 1) = r.Cells(i, j).Value
      Else
        x((i - 1) * N + j, 1) = r.Cells(i, j).Value
      End If
    Next
  Next
  
  SpillTest_Nx1 = x
  
End Function



Public Function SpillTest_1xM(r As Range, Optional down_columns As Boolean = True) As Variant

  Dim N As Integer
  Dim M As Integer
  Dim x() As Variant
  Dim i As Integer
  Dim j As Integer
  
  N = r.Rows.Count
  M = r.Columns.Count
  
  ReDim x(1 To 1, 1 To N * M) As Variant
  
  For i = 1 To N
    For j = 1 To M
      If down_columns Then
        x(1, (j - 1) * M + i) = r.Cells(i, j).Value
      Else
        x(1, (i - 1) * N + j) = r.Cells(i, j).Value
      End If
    Next
  Next
  
  SpillTest_1xM = x
  
End Function


