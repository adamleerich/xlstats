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


' When to use SUB, FUNCTION, LET, GET & SET
' https://stackoverflow.com/a/75615408
'
' > has _return?
'   |_ yes > has parameter?
'   |        |_ yes: Function
'   |        |_ no > likely verb?
'   |                |_ yes: Function
'   |                |_ no: Get
'   |_ no > has parameter?
'           |_ yes > likely verb?
'           |        |_ yes: Sub
'           |        |_ no > uses _object?
'           |                |_ yes: Set
'           |                |_ no: Let
'           |_ no: Sub
'


Private latest_seed As LongLong
Private counter_uses As Long
Private counter_twists As Long
Private mt_initialized As Boolean
Private mt_seeds(1 To 624) As Long
Private mt_cursor As Integer

Private Const two32 As LongLong = 2 ^ 32
Private Const two31 As LongLong = 2 ^ 31
  

Private Sub Class_Initialize()
  mt_initialized = False
End Sub


Property Get Cursor() As Integer
  Cursor = mt_cursor
End Property


Private Sub mt_init(Optional pSeed As Variant)
  
  Dim pSeed64 As LongLong
  
  ' Get seed if missing
  If IsMissing(pSeed) Then
    pSeed64 = RandomSeed
  Else
    pSeed64 = CLngLng(pSeed)
  End If
  
  Dim i As Integer
  Dim mSeeds(1 To 676) As LongLong '676 = 624+52
  
  mSeeds(1) = CInt32__Scalar(pSeed64)
  For i = 2 To 676
    pSeed64 = mSeeds(i - 1) * 69069 + 1
    mSeeds(i) = CInt32__Scalar(pSeed64)
  Next
  
  For i = 1 To 624
    mt_seeds(i) = CLng(mSeeds(i + 52))
  Next
  
  mt_cursor = 624
  latest_seed = pSeed
  counter_uses = 0
  counter_twists = 0
  mt_initialized = True
  
End Sub


Property Let Seed(x As LongLong)
  mt_init x
End Property


Private Function RandomSeed() As Long
  RandomSeed = Int(two32 * Rnd - two31)
End Function


Function get_randomseeds() As Long()
  get_randomseeds = mt_seeds
End Function

