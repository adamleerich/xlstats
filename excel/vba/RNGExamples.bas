Attribute VB_Name = "RNGExamples"
Option Explicit

Public Function CHECK__CursorProperty()

  Dim mRNG As New XLStatsRNG
  Dim y() As Long
  Dim z(1 To 624, 1 To 1) As Long
  Dim i As Integer
  
  ' Debug.Print mRNG.RandomSeed
  mRNG.Seed = 123
  mRNG.mt_twist
  y = mRNG.get_randomseeds
  
  For i = 1 To 624
    z(i, 1) = y(i)
  Next
  
  CHECK__CursorProperty = z
  
End Function


