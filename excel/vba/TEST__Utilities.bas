Attribute VB_Name = "TEST__Utilities"
Option Explicit

Sub Test_CInt32__Scalar()
    Debug.Print CInt32__Scalar(0)           ' Expected: 0
    Debug.Print CInt32__Scalar(2147483647)  ' Expected: 2147483647 (Max CInt32__Scalar)
    Debug.Print CInt32__Scalar(2147483648^) ' Expected: -2147483648 (Overflow wraps around)
    Debug.Print CInt32__Scalar(-1)          ' Expected: -1
    Debug.Print CInt32__Scalar(-2147483648^) ' Expected: -2147483648 (Min CInt32__Scalar)
    Debug.Print CInt32__Scalar(-2147483649^) ' Expected: 2147483647 (Underflow wraps around)
    Debug.Print CInt32__Scalar(4294967295^) ' Expected: -1 (Unsigned max wraps to -1)
    Debug.Print CInt32__Scalar(4294967296^) ' Expected: 0 (Wraps back to zero)
End Sub

Sub Test_ShiftRight32()
    Debug.Print ShiftRight32(8, 1)      ' Expected: 4 (0000 1000 >> 1 = 0000 0100)
    Debug.Print ShiftRight32(8, 3)      ' Expected: 1 (0000 1000 >> 3 = 0000 0001)
    Debug.Print ShiftRight32(-8, 1)     ' Expected: 2147483644 (Logical shift removes sign)
    Debug.Print ShiftRight32(-1, 1)     ' Expected: 2147483647 (0xFFFFFFFF >> 1 = 0x7FFFFFFF)
    Debug.Print ShiftRight32(-2147483648#, 1) ' Expected: 1073741824 (0x80000000 >> 1)
    Debug.Print ShiftRight32(1, 32)     ' Expected: 0 (Shifting beyond 31 bits should be 0)
End Sub

Sub Test_ShiftLeft32()
    Debug.Print ShiftLeft32(1, 1)      ' Expected: 2 (0000 0001 << 1 = 0000 0010)
    Debug.Print ShiftLeft32(1, 3)      ' Expected: 8 (0000 0001 << 3 = 0000 1000)
    Debug.Print ShiftLeft32(1073741824, 1) ' Expected: -2147483648 (Sign bit is hit)
    Debug.Print ShiftLeft32(-1073741824, 1) ' Expected: 2147483648 (Overflows to 0)
    Debug.Print ShiftLeft32(1, 32)     ' Expected: 0 (Shifting beyond 31 bits should be 0)
End Sub


Sub CHECK__CInt32()
  
  Dim x(10 To 15) As LongLong
  Dim y() As Long
  Dim i As Long
  
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

