VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XLStatsRNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mt_seeds(1 To 624) As LongLong
Private mt_cursor As Integer


Private Sub Class_Initialize()
  mt_cursor = 1
End Sub


Public Function IncrementCursor() As Integer
  mt_cursor = mt_cursor + 1
  IncrementCursor = mt_cursor
End Function


Private Sub Class_Terminate()


End Sub






