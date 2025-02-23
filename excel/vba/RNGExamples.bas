Attribute VB_Name = "RNGExamples"
Option Explicit

Public Function XLStats_runif(N As Integer, pMin As Double, pMax As Double) As Double()

  Dim mRng As New XLStatsRNG
  Dim y() As Double
  Dim z() As Double
  Dim i As Integer
  
  ReDim z(1 To N, 1 To 1) As Double
  XLStats_runif = z
  
  ' Debug.Print mRNG.RandomSeed
  mRng.seed = 123
  y = mRng.runif(N, pMin, pMax)
  
  ' For i = 1 To N
  '   z(i, 1) = y(i)
  ' Next
  
End Function


Public Function XLStatsRNG_runif_internal(seed As Long) As Double
  Dim mRng As New XLStatsRNG
  XLStatsRNG_runif_internal = mRng.runif_internal(seed)
End Function
