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

Public Function mt_twist()
  
  If Not mt_initialized Then
    mt_init
  End If

  Dim N As Integer
  Dim m As Integer
    
  Dim MATRIX_A As Long
  Dim LOWER_MASK As Long
  Dim UPPER_MASK As Long
  Dim mag01(1 To 2) As Long
  Dim k As Integer
  Dim y As Long
  
  Dim seeds(1 To 624) As Long
  For k = 1 To 624
    seeds(k) = mt_seeds(k)
  Next
  
  N = 624
  m = 397
  MATRIX_A = &H9908B0DF      '10011001000010001011000011011111
  LOWER_MASK = &H7FFFFFFF    '01111111111111111111111111111111
  UPPER_MASK = &H80000000    '10000000000000000000000000000000
  mag01(1) = 0
  mag01(2) = MATRIX_A
    
  ' 1:227 references 1:228,398:624
  For k = 1 To (N - m)
    y = (seeds(k) And UPPER_MASK) Or (seeds(k + 1) And LOWER_MASK)
    seeds(k) = seeds(k + m) Xor ShiftRight32(y, 1) Xor mag01((y And &H1) + 1)
  Next k
    
  ' 228:454 references 228:455,1:227
  For k = (N - m + 1) To (2 * (N - m))
    y = (seeds(k) And UPPER_MASK) Or (seeds(k + 1) And LOWER_MASK)
    seeds(k) = seeds(k + (m - N)) Xor ShiftRight32(y, 1) Xor mag01((y And &H1) + 1)
  Next k
    
  ' 455:623 references 455:624,228:396
  For k = (2 * (N - m) + 1) To (N - 1)
    y = (seeds(k) And UPPER_MASK) Or (seeds(k + 1) And LOWER_MASK)
    seeds(k) = seeds(k + (m - N)) Xor ShiftRight32(y, 1) Xor mag01((y And &H1) + 1)
  Next k
    
  ' k = 624 references 1,624,397
  y = (seeds(N) And UPPER_MASK) Or (seeds(1) And LOWER_MASK)
  seeds(N) = seeds(m) Xor ShiftRight32(y, 1) Xor mag01((y And &H1) + 1)
  
  For k = 1 To 624
    mt_seeds(k) = seeds(k)
  Next
  
  counter_twists = counter_twists + 1
  mt_cursor = 0
  
End Function


