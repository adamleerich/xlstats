Attribute VB_Name = "Moments"
Option Explicit

Private Function ToArray__Range(r As Range) As Double()

  Dim m As Integer
  Dim N As Integer
  Dim i As Integer
  Dim j As Integer
  
  N = r.Rows.Count
  m = r.Columns.Count
  
  Dim x() As Double
  
  ReDim x(1 To N * m) As Double
  
  For i = 1 To N
    For j = 1 To m
      x((j - 1) * N + i) = CDbl(r.Cells(i, j).Value)
    Next
  Next
  
  ToArray__Range = x
  
End Function


Public Function XLStatsMean(r As Range) As Double
  Dim x() As Double
  x = ToArray__Range(r)
  XLStatsMean = Mean__Array(x)
End Function

Public Function XLStatsVariance(r As Range, Optional pMethod As String = "Population") As Double
  Dim x() As Double
  x = ToArray__Range(r)
  XLStatsVariance = Variance__Array(x, pMethod)
End Function


Private Function Mean__Array(x() As Double) As Double

  Dim mSum As Double
  Dim i As Integer
  
  For i = LBound(x) To UBound(x)
    mSum = mSum + x(i)
  Next
  
  Mean__Array = mSum / (UBound(x) - LBound(x) + 1)
  
End Function



Private Function Variance__Array(x() As Double, pMethod) As Double

  Dim mSum As Double
  Dim mMean As Double
  Dim i As Integer
  Dim N As Integer
  
  pMethod = LCase(pMethod)
  
  If pMethod <> "population" And _
      pMethod <> "sample" Then
    Err.Raise 1022, "Variance__Array", "Unexpected argument value: pMethod = '" & pMethod & "'"
  End If
  
  mMean = Mean__Array(x)
  
  For i = LBound(x) To UBound(x)
    mSum = mSum + (x(i) - mMean) ^ 2
  Next
  
  If pMethod = "population" Then
    N = UBound(x) - LBound(x) + 1
  Else
    N = UBound(x) - LBound(x)
  End If
  
  Variance__Array = mSum / N
  
End Function




Private Function FitDistMME__Array(x() As Double, pDistr As String) As Variant()
  
  Dim m As Double 'Mean
  Dim v As Double 'Population variance
  
  pDistr = LCase(pDistr)
  
  If pDistr <> "norm" And _
      pDistr <> "lnorm" And _
      pDistr <> "pois" And _
      pDistr <> "exp" And _
      pDistr <> "gamma" And _
      pDistr <> "nbinom" And _
      pDistr <> "geom" And _
      pDistr <> "beta" And _
      pDistr <> "unif" And _
      pDistr <> "logis" Then
    Err.Raise 1022, "FitDistMME__Array", "Unexpected argument value: pDistr = '" & pDistr & "'"
  End If
  
  m = Mean__Array(x)
  
  
  
End Function



