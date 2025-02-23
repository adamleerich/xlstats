Attribute VB_Name = "XLStats"
Option Explicit




Public Function XLStatsMean(r As Range) As Double
  Dim x() As Double
  x = CDblArray__Range(r)
  XLStatsMean = Mean__Array(x)
End Function

Public Function XLStatsVariance(r As Range, Optional pMethod As String = "Population") As Double
  Dim x() As Double
  x = CDblArray__Range(r)
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



Public Function XLStatsFitDistMME(r As Range, pDistr As String, Optional pVerbose As Boolean = True) As Variant()
  Dim x() As Double
  x = CDblArray__Range(r)
  XLStatsFitDistMME = FitDistMME__Array(x, pDistr, pVerbose)
End Function



Private Function FitDistMME__Array(x() As Double, pDistr As String, pVerbose As Boolean) As Variant()
  
  Dim m As Double 'Mean
  Dim v As Double 'Population variance
  
  Dim p1 As Double 'Parameter 1
  Dim p2 As Double 'Parameter 2
  Dim p1_name As String
  Dim p2_name As String
  Dim m_form As String 'Formula for mean in terms of parameters for reference
  Dim v_form As String 'Formula for variance
  
  Dim result_verbose(1 To 4, 1 To 2) As Variant
  Dim result_base(1 To 2) As Variant
  
  pDistr = LCase(pDistr)
  
  ' TODO
  ' Add pareto, pois, exp, gamma, nbinom, geom, beta, logis, etc.?
  If pDistr <> "norm" And _
      pDistr <> "lnorm" And _
      pDistr <> "unif" Then
    Err.Raise 1022, "FitDistMME__Array", "Unexpected argument value: pDistr = '" & pDistr & "'"
  End If
  
  m = Mean__Array(x)
  v = Variance__Array(x, "Population")
  
  
  If pDistr = "unif" Then
    Dim r As Double
    r = Sqr(3 * v) ' half-width of interval, with mean as center
    p1 = m - r
    p2 = m + r
    p1_name = "a: minimum"
    p2_name = "b: maximum"
    m_form = "1/2*(a+b)"
    v_form = "1/12*(b-a)^2"
  End If
  
  
  If pDistr = "norm" Then
    p1 = m
    p2 = Sqr(v)
    p1_name = "m: mean"
    p2_name = "s: sd"
    m_form = "m"
    v_form = "s^2"
  End If
  
  
  If pDistr = "lnorm" Then
    ' https://en.wikipedia.org/wiki/Log-normal_distribution#Method_of_moments
    p1 = Log(m / Sqr(1 + v / m ^ 2))
    p2 = Sqr(Log(1 + v / m ^ 2))
    p1_name = "m: mu"
    p2_name = "s: sigma"
    m_form = "exp(m + s^2/2)"
    v_form = "[exp(s^2)-1]*exp(2m+s^2)"
  End If
  
  
  
  If pVerbose Then
    result_verbose(1, 1) = p1_name
    result_verbose(1, 2) = p1
    result_verbose(2, 1) = p2_name
    result_verbose(2, 2) = p2
    result_verbose(3, 1) = "Mean = "
    result_verbose(3, 2) = m_form
    result_verbose(4, 1) = "Variance = "
    result_verbose(4, 2) = v_form
    FitDistMME__Array = result_verbose
  Else
    result_base(1) = p1
    result_base(2) = p2
    FitDistMME__Array = result_base
  End If
  
End Function



