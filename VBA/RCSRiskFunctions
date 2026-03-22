' -----------------------------------------------------------------------------------
' Author: Aaron Moynahan
' Copyright (c) 2026 Aaron Moynahan
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software to use, copy, modify, and distribute it, subject to the
' following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND.
' -----------------------------------------------------------------------------------

Option Explicit

'PI_VAL: Value of p used in normal density calculations. Required for standard normal PDF (phi_pdf)
Private Const PI_VAL As Double = 3.14159265358979
'IG_EPS: Small tolerance for incomplete gamma calculations. Controls convergence and avoids numerical instability
Private Const IG_EPS As Double = 0.000000000001
'IG_FPMIN: Very small positive number to avoid division underflow. Prevents divide-by-zero in continued fraction algorithms
Private Const IG_FPMIN As Double = 1E-300
'IG_MAXIT: Maximum number of iterations for incomplete gamma routines. Limits computation time and ensures convergence termination
Private Const IG_MAXIT As Long = 200

'BIG_POS: Large positive sentinel value (~infinity). Used to cap overflow or represent very large results
Private Const BIG_POS As Double = 1E+300
'BIG_NEG: Large negative sentinel value (~-infinity). Used for log-scale outputs when values are zero or invalid
Private Const BIG_NEG As Double = -1E+300
'DEFAULT_TOL: Default tolerance for numerical solvers. Controls accuracy of root-finding algorithms
Private Const DEFAULT_TOL As Double = 0.0000001
'DEFAULT_MAXITER: Default maximum iterations for numerical solvers. Prevents infinite loops in root-finding
Private Const DEFAULT_MAXITER As Long = 200

'FLEX_GG_Q_EPS: Small threshold for generalized gamma shape parameter (q ˜ 0). Triggers fallback to log-normal limit for stability
Private Const FLEX_GG_Q_EPS As Double = 0.000000000001
'FLEX_GF_P_EPS: Small threshold for generalized F parameter (p ˜ 0). Triggers fallback to generalized gamma for stability
Private Const FLEX_GF_P_EPS As Double = 0.000000000001

'DIST_EXP: Code identifier for exponential distribution. Used in dispatch logic for distribution selection
Private Const DIST_EXP As Long = 1
'DIST_WEIBULL: Code identifier for Weibull (AFT parameterization). Used in distribution dispatch routines
Private Const DIST_WEIBULL As Long = 2
'DIST_WEIBULLPH: Code identifier for Weibull proportional hazards parameterization. Supports PH-form Weibull models
Private Const DIST_WEIBULLPH As Long = 3
'DIST_GOMPERTZ: Code identifier for Gompertz distribution. Used in distribution dispatch routines
Private Const DIST_GOMPERTZ As Long = 4
'DIST_LNORM: Code identifier for log-normal distribution. Used in distribution dispatch routines
Private Const DIST_LNORM As Long = 5
'DIST_LLOGIS: Code identifier for log-logistic distribution. Used in distribution dispatch routines
Private Const DIST_LLOGIS As Long = 6
'DIST_GAMMA: Code identifier for gamma distribution. Used in distribution dispatch routines
Private Const DIST_GAMMA As Long = 7


Private Const RMST_MIN_STEPS As Long = 256
Private Const RMST_MAX_STEPS As Long = 16384
Private Const RMST_REL_TOL As Double = 0.0000001

'==========================================================
' Basic helpers
'==========================================================

' MaxDbl
' Returns the larger of two double values
' Used in numerical routines and truncation logic
Private Function MaxDbl(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then
        MaxDbl = a
    Else
        MaxDbl = b
    End If
End Function
' SafeExp

' Computes exponential with overflow protection
' Prevents numerical errors for very large or small inputs
Private Function SafeExp(ByVal x As Double) As Double
    If x > 700# Then
        SafeExp = Exp(700#)
    ElseIf x < -700# Then
        SafeExp = Exp(-700#)
    Else
        SafeExp = Exp(x)
    End If
End Function

' SafeLog
' Computes log safely, returning a large negative value for nonpositive inputs
' Avoids runtime errors in log-scale calculations
Private Function SafeLog(ByVal x As Double) As Double
    If x <= 0# Then
        SafeLog = BIG_NEG
    Else
        SafeLog = Log(x)
    End If
End Function

' StdNormCDF
' Returns standard normal cumulative probability
' Used for log-normal and other normal-based calculations
Private Function StdNormCDF(ByVal z As Double) As Double
    StdNormCDF = Application.WorksheetFunction.Norm_S_Dist(z, True)
End Function

' StdNormPDF
' Returns standard normal density
' Used for density and hazard calculations involving the normal distribution
Private Function StdNormPDF(ByVal z As Double) As Double
    StdNormPDF = Exp(-0.5 * z * z) / Sqr(2# * PI_VAL)
End Function

' StdNormInv
' Returns inverse standard normal cumulative probability
' Used for quantile calculations under the log-normal distribution
Private Function StdNormInv(ByVal p As Double) As Double
    StdNormInv = Application.WorksheetFunction.Norm_S_Inv(p)
End Function

' UniformOpen01
' Generates a random uniform value strictly between 0 and 1
' Avoids boundary problems in inverse-CDF random generation
Private Function UniformOpen01() As Double
    Dim u As Double
    u = Rnd
    If u <= 0# Then u = 0.0000000000001
    If u >= 1# Then u = 0.9999999999999
    UniformOpen01 = u
End Function

' FinishProb
' Applies lower-tail or upper-tail and optional log transformation to a probability
' Standardizes probability output formatting across distributions
Private Function FinishProb(ByVal pVal As Double, _
                            ByVal lowerTail As Boolean, _
                            ByVal logP As Boolean) As Double
    Dim outVal As Double
    
    If lowerTail Then
        outVal = pVal
    Else
        outVal = 1# - pVal
    End If
    
    If logP Then
        FinishProb = SafeLog(outVal)
    Else
        FinishProb = outVal
    End If
End Function

' DecodeProb
' Converts input probability from log and/or upper-tail form to lower-tail probability
' Standardizes probability input before quantile calculations
Private Function DecodeProb(ByVal p As Double, _
                            ByVal lowerTail As Boolean, _
                            ByVal logP As Boolean) As Double
    Dim pp As Double
    
    If logP Then
        pp = Exp(p)
    Else
        pp = p
    End If
    
    If Not lowerTail Then
        pp = 1# - pp
    End If
    
    DecodeProb = pp
End Function

' ScalarOrArray
' Applies a distribution function to a scalar, array, or Excel range
' Lets worksheet functions work with single values or spilled ranges
Private Function ScalarOrArray(ByVal x As Variant, _
                               ByVal distCode As Long, _
                               ByVal mode As String, _
                               ByVal p1 As Double, _
                               Optional ByVal p2 As Double = 0#, _
                               Optional ByVal lowerTail As Boolean = True, _
                               Optional ByVal logP As Boolean = False, _
                               Optional ByVal logFlag As Boolean = False) As Variant
    Dim vals As Variant
    Dim out() As Variant
    Dim r As Long, c As Long
    Dim rng As Range
    
    On Error GoTo EH
    
    If IsObject(x) Then
        If TypeName(x) = "Range" Then
            Set rng = x
            
            If rng.CountLarge = 1 Then
                If IsNumeric(rng.Value2) Then
                    ScalarOrArray = EvalOne(CDbl(rng.Value2), distCode, mode, p1, p2, lowerTail, logP, logFlag)
                Else
                    ScalarOrArray = CVErr(xlErrValue)
                End If
                Exit Function
            End If
            
            vals = rng.Value2
            ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
            
            For r = 1 To UBound(vals, 1)
                For c = 1 To UBound(vals, 2)
                    If IsNumeric(vals(r, c)) Then
                        out(r, c) = EvalOne(CDbl(vals(r, c)), distCode, mode, p1, p2, lowerTail, logP, logFlag)
                    Else
                        out(r, c) = CVErr(xlErrValue)
                    End If
                Next c
            Next r
            
            ScalarOrArray = out
            Exit Function
        End If
    End If
    
    If IsArray(x) Then
        vals = x
        ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
        
        For r = 1 To UBound(vals, 1)
            For c = 1 To UBound(vals, 2)
                If IsNumeric(vals(r, c)) Then
                    out(r, c) = EvalOne(CDbl(vals(r, c)), distCode, mode, p1, p2, lowerTail, logP, logFlag)
                Else
                    out(r, c) = CVErr(xlErrValue)
                End If
            Next c
        Next r
        
        ScalarOrArray = out
    ElseIf IsNumeric(x) Then
        ScalarOrArray = EvalOne(CDbl(x), distCode, mode, p1, p2, lowerTail, logP, logFlag)
    Else
        ScalarOrArray = CVErr(xlErrValue)
    End If
    
    Exit Function
    
EH:
    ScalarOrArray = CVErr(xlErrValue)
End Function


' RandomVector
' Generates one or more random draws from the selected distribution
' Provides vectorized random sampling for worksheet use
Private Function RandomVector(ByVal n As Long, _
                              ByVal distCode As Long, _
                              ByVal p1 As Double, _
                              Optional ByVal p2 As Double = 0#) As Variant
    Dim out() As Variant
    Dim i As Long
    
    If n <= 0 Then
        RandomVector = CVErr(xlErrValue)
        Exit Function
    End If
    
    If n = 1 Then
        RandomVector = EvalRandom(distCode, p1, p2)
        Exit Function
    End If
    
    ReDim out(1 To n, 1 To 1)
    For i = 1 To n
        out(i, 1) = EvalRandom(distCode, p1, p2)
    Next i
    
    RandomVector = out
End Function



'==========================================================
' Incomplete gamma / incomplete beta for VBA / Excel
'
' Provided functions:
'
' LogGammaVBA(x)               = log Gamma(x)
' LogBetaVBA(a, b)             = log Beta(a,b)
'
' GammaP(a, x)                 = regularized lower incomplete gamma
' GammaQ(a, x)                 = regularized upper incomplete gamma
' LowerIncGamma(a, x)          = lower incomplete gamma
' UpperIncGamma(a, x)          = upper incomplete gamma
'
' BetaI(x, a, b)               = regularized incomplete beta
' LowerIncBeta(x, a, b)        = incomplete beta from 0 to x
' UpperIncBeta(x, a, b)        = incomplete beta from x to 1
'
' Notes:
' - GammaP(a,x) + GammaQ(a,x) = 1
' - LowerIncGamma(a,x) + UpperIncGamma(a,x) = Gamma(a)
' - LowerIncBeta(x,a,b) + UpperIncBeta(x,a,b) = Beta(a,b)
'
'==========================================================

'Some funcitons constants deleted because they already existed.

'==========================================================
' Basic helpers
'==========================================================

' MinDbl
' Returns the smaller of two double values
' Used in numerical helper logic where lower bounds matter
Private Function MinDbl(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then
        MinDbl = a
    Else
        MinDbl = b
    End If
End Function


'==========================================================
' log-gamma and log-beta
'==========================================================

' LogGammaVBA
' Computes log-gamma using Excel functions with fallback support
' Provides a stable gamma log needed by special functions and distributions
Public Function LogGammaVBA(ByVal x As Double) As Double
    On Error GoTo Fallback
    
    If x <= 0# Then
        Err.Raise vbObjectError + 5000, , "x must be > 0"
    End If
    
    ' Preferred if available
    LogGammaVBA = Application.WorksheetFunction.GammaLn_Precise(x)
    Exit Function

Fallback:
    On Error GoTo EH
    
    If x <= 0# Then
        Err.Raise vbObjectError + 5001, , "x must be > 0"
    End If
    
    ' Older Excel fallback
    LogGammaVBA = Application.WorksheetFunction.GammaLn(x)
    Exit Function

EH:
    LogGammaVBA = CVErr(xlErrNum)
End Function

' LogBetaVBA
' Computes log-beta from log-gamma values
' Used in incomplete beta and related distribution calculations
Public Function LogBetaVBA(ByVal a As Double, ByVal b As Double) As Double
    On Error GoTo EH
    
    If a <= 0# Or b <= 0# Then
        Err.Raise vbObjectError + 5002, , "a and b must be > 0"
    End If
    
    LogBetaVBA = LogGammaVBA(a) + LogGammaVBA(b) - LogGammaVBA(a + b)
    Exit Function

EH:
    LogBetaVBA = CVErr(xlErrNum)
End Function

' BetaFuncVBA
' Computes beta function from its log form
' Avoids overflow/underflow in beta evaluation
Private Function BetaFuncVBA(ByVal a As Double, ByVal b As Double) As Double
    BetaFuncVBA = SafeExp(LogBetaVBA(a, b))
End Function

'==========================================================
' Incomplete gamma: internal series for P(a,x)
'
' Returns regularized lower incomplete gamma P(a,x)
' Best when x < a + 1
'==========================================================

' GammaPSeries
' Computes regularized lower incomplete gamma using a series expansion
' Best for cases where x is smaller than a + 1
Private Function GammaPSeries(ByVal a As Double, ByVal x As Double) As Double
    Dim gln As Double
    Dim sum_ As Double
    Dim del As Double
    Dim ap As Double
    Dim n As Long
    
    If a <= 0# Then Err.Raise vbObjectError + 5100, , "a must be > 0"
    If x < 0# Then Err.Raise vbObjectError + 5101, , "x must be >= 0"
    
    If x = 0# Then
        GammaPSeries = 0#
        Exit Function
    End If
    
    gln = LogGammaVBA(a)
    ap = a
    del = 1# / a
    sum_ = del
    
    For n = 1 To IG_MAXIT
        ap = ap + 1#
        del = del * x / ap
        sum_ = sum_ + del
        
        If Abs(del) < Abs(sum_) * IG_EPS Then
            GammaPSeries = sum_ * SafeExp(-x + a * Log(x) - gln)
            Exit Function
        End If
    Next n
    
    GammaPSeries = sum_ * SafeExp(-x + a * Log(x) - gln)
End Function

'==========================================================
' Incomplete gamma: internal continued fraction for Q(a,x)
'
' Returns regularized upper incomplete gamma Q(a,x)
' Best when x >= a + 1
'==========================================================

' GammaQCF
' Computes regularized upper incomplete gamma using a continued fraction
' Best for cases where x is at least a + 1
Private Function GammaQCF(ByVal a As Double, ByVal x As Double) As Double
    Dim gln As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim h As Double
    Dim an As Double
    Dim del As Double
    Dim i As Long
    
    If a <= 0# Then Err.Raise vbObjectError + 5200, , "a must be > 0"
    If x < 0# Then Err.Raise vbObjectError + 5201, , "x must be >= 0"
    
    If x = 0# Then
        GammaQCF = 1#
        Exit Function
    End If
    
    gln = LogGammaVBA(a)
    b = x + 1# - a
    c = 1# / IG_FPMIN
    d = 1# / b
    h = d
    
    For i = 1 To IG_MAXIT
        an = -CDbl(i) * (CDbl(i) - a)
        b = b + 2#
        
        d = an * d + b
        If Abs(d) < IG_FPMIN Then d = IG_FPMIN
        
        c = b + an / c
        If Abs(c) < IG_FPMIN Then c = IG_FPMIN
        
        d = 1# / d
        del = d * c
        h = h * del
        
        If Abs(del - 1#) < IG_EPS Then
            GammaQCF = h * SafeExp(-x + a * Log(x) - gln)
            Exit Function
        End If
    Next i
    
    GammaQCF = h * SafeExp(-x + a * Log(x) - gln)
End Function

'==========================================================
' Public incomplete gamma
'==========================================================

' GammaP
' Returns regularized lower incomplete gamma P(a, x)
' Public wrapper choosing the numerically appropriate method
Public Function GammaP(ByVal a As Double, ByVal x As Double) As Variant
    On Error GoTo EH
    
    If a <= 0# Then Err.Raise vbObjectError + 5300, , "a must be > 0"
    If x < 0# Then Err.Raise vbObjectError + 5301, , "x must be >= 0"
    
    If x = 0# Then
        GammaP = 0#
    ElseIf x < a + 1# Then
        GammaP = GammaPSeries(a, x)
    Else
        GammaP = 1# - GammaQCF(a, x)
    End If
    
    Exit Function

EH:
    GammaP = CVErr(xlErrNum)
End Function

' GammaQ
' Returns regularized upper incomplete gamma Q(a, x)
' Public wrapper choosing the numerically appropriate method
Public Function GammaQ(ByVal a As Double, ByVal x As Double) As Variant
    On Error GoTo EH
    
    If a <= 0# Then Err.Raise vbObjectError + 5302, , "a must be > 0"
    If x < 0# Then Err.Raise vbObjectError + 5303, , "x must be >= 0"
    
    If x = 0# Then
        GammaQ = 1#
    ElseIf x < a + 1# Then
        GammaQ = 1# - GammaPSeries(a, x)
    Else
        GammaQ = GammaQCF(a, x)
    End If
    
    Exit Function

EH:
    GammaQ = CVErr(xlErrNum)
End Function

' LowerIncGamma
' Returns unregularized lower incomplete gamma
' Converts regularized gamma probability back to integral scale
Public Function LowerIncGamma(ByVal a As Double, ByVal x As Double) As Variant
    On Error GoTo EH
    
    If a <= 0# Then Err.Raise vbObjectError + 5304, , "a must be > 0"
    If x < 0# Then Err.Raise vbObjectError + 5305, , "x must be >= 0"
    
    LowerIncGamma = CDbl(GammaP(a, x)) * SafeExp(LogGammaVBA(a))
    Exit Function

EH:
    LowerIncGamma = CVErr(xlErrNum)
End Function

' UpperIncGamma
' Returns unregularized upper incomplete gamma
' Converts regularized gamma probability back to integral scale
Public Function UpperIncGamma(ByVal a As Double, ByVal x As Double) As Variant
    On Error GoTo EH
    
    If a <= 0# Then Err.Raise vbObjectError + 5306, , "a must be > 0"
    If x < 0# Then Err.Raise vbObjectError + 5307, , "x must be >= 0"
    
    UpperIncGamma = CDbl(GammaQ(a, x)) * SafeExp(LogGammaVBA(a))
    Exit Function

EH:
    UpperIncGamma = CVErr(xlErrNum)
End Function

'==========================================================
' Incomplete beta: continued fraction
'
' Internal routine for regularized incomplete beta
'==========================================================

' BetaCF
' Computes continued fraction for incomplete beta evaluation
' Core numerical routine for regularized incomplete beta
Private Function BetaCF(ByVal a As Double, ByVal b As Double, ByVal x As Double) As Double
    Dim qab As Double
    Dim qap As Double
    Dim qam As Double
    Dim c As Double
    Dim d As Double
    Dim h As Double
    Dim aa As Double
    Dim del As Double
    Dim m As Long
    Dim m2 As Long
    
    If a <= 0# Or b <= 0# Then Err.Raise vbObjectError + 5400, , "a and b must be > 0"
    If x < 0# Or x > 1# Then Err.Raise vbObjectError + 5401, , "x must be in [0,1]"
    
    qab = a + b
    qap = a + 1#
    qam = a - 1#
    
    c = 1#
    d = 1# - qab * x / qap
    If Abs(d) < IG_FPMIN Then d = IG_FPMIN
    d = 1# / d
    h = d
    
    For m = 1 To IG_MAXIT
        m2 = 2 * m
        
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1# + aa * d
        If Abs(d) < IG_FPMIN Then d = IG_FPMIN
        c = 1# + aa / c
        If Abs(c) < IG_FPMIN Then c = IG_FPMIN
        d = 1# / d
        h = h * d * c
        
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1# + aa * d
        If Abs(d) < IG_FPMIN Then d = IG_FPMIN
        c = 1# + aa / c
        If Abs(c) < IG_FPMIN Then c = IG_FPMIN
        d = 1# / d
        del = d * c
        h = h * del
        
        If Abs(del - 1#) < IG_EPS Then
            BetaCF = h
            Exit Function
        End If
    Next m
    
    BetaCF = h
End Function

'==========================================================
' Regularized incomplete beta I_x(a,b)
'==========================================================

' BetaI
' Returns regularized incomplete beta I_x(a, b)
' Public beta probability function used by flexible distributions
Public Function BetaI(ByVal x As Double, ByVal a As Double, ByVal b As Double) As Variant
    Dim bt As Double
    
    On Error GoTo EH
    
    If a <= 0# Or b <= 0# Then Err.Raise vbObjectError + 5500, , "a and b must be > 0"
    If x < 0# Or x > 1# Then Err.Raise vbObjectError + 5501, , "x must be in [0,1]"
    
    If x = 0# Then
        BetaI = 0#
        Exit Function
    End If
    
    If x = 1# Then
        BetaI = 1#
        Exit Function
    End If
    
    bt = SafeExp(LogGammaVBA(a + b) - LogGammaVBA(a) - LogGammaVBA(b) + a * Log(x) + b * Log(1# - x))
    
    If x < (a + 1#) / (a + b + 2#) Then
        BetaI = bt * BetaCF(a, b, x) / a
    Else
        BetaI = 1# - bt * BetaCF(b, a, 1# - x) / b
    End If
    
    Exit Function

EH:
    BetaI = CVErr(xlErrNum)
End Function

'==========================================================
' Unregularized incomplete beta
'
' B_x(a,b)     = integral from 0 to x
' Upper part   = integral from x to 1
'==========================================================

' LowerIncBeta
' Returns unregularized lower incomplete beta
' Converts beta probability back to integral scale
Public Function LowerIncBeta(ByVal x As Double, ByVal a As Double, ByVal b As Double) As Variant
    On Error GoTo EH
    
    If a <= 0# Or b <= 0# Then Err.Raise vbObjectError + 5502, , "a and b must be > 0"
    If x < 0# Or x > 1# Then Err.Raise vbObjectError + 5503, , "x must be in [0,1]"
    
    LowerIncBeta = CDbl(BetaI(x, a, b)) * SafeExp(LogBetaVBA(a, b))
    Exit Function

EH:
    LowerIncBeta = CVErr(xlErrNum)
End Function

' UpperIncBeta
' Returns unregularized upper incomplete beta
' Returns the upper beta integral from x to 1
Public Function UpperIncBeta(ByVal x As Double, ByVal a As Double, ByVal b As Double) As Variant
    On Error GoTo EH
    
    If a <= 0# Or b <= 0# Then Err.Raise vbObjectError + 5504, , "a and b must be > 0"
    If x < 0# Or x > 1# Then Err.Raise vbObjectError + 5505, , "x must be in [0,1]"
    
    UpperIncBeta = (1# - CDbl(BetaI(x, a, b))) * SafeExp(LogBetaVBA(a, b))
    Exit Function

EH:
    UpperIncBeta = CVErr(xlErrNum)
End Function


'==========================================================
' Exponential
'==========================================================

' EvalExp
' Evaluates exponential density, probability, quantile, hazard, cumulative hazard, or survival
' Core engine for exponential worksheet functions
Private Function EvalExp(ByVal x As Double, ByVal mode As String, _
                         ByVal rate_ As Double, _
                         ByVal lowerTail As Boolean, _
                         ByVal logP As Boolean, _
                         ByVal logFlag As Boolean) As Variant
    Dim dens As Double, cdf As Double, surv As Double
    Dim haz As Double, cumHaz As Double, p As Double
    
    If rate_ <= 0# Then Err.Raise vbObjectError + 2000, , "rate must be > 0"
    
    Select Case LCase$(mode)
        Case "d"
            If x < 0# Then
                dens = 0#
            Else
                dens = rate_ * Exp(-rate_ * x)
            End If
            If logFlag Then
                EvalExp = SafeLog(dens)
            Else
                EvalExp = dens
            End If
        
        Case "p"
            If x <= 0# Then
                cdf = 0#
            Else
                cdf = 1# - Exp(-rate_ * x)
            End If
            EvalExp = FinishProb(cdf, lowerTail, logP)
        
        Case "q"
            p = DecodeProb(x, lowerTail, logP)
            If p < 0# Or p > 1# Then Err.Raise vbObjectError + 2001, , "p must be in [0,1]"
            If p <= 0# Then
                EvalExp = 0#
            ElseIf p >= 1# Then
                EvalExp = BIG_POS
            Else
                EvalExp = -Log(1# - p) / rate_
            End If
        
        Case "h"
            If x < 0# Then
                haz = 0#
            Else
                haz = rate_
            End If
            If logFlag Then
                EvalExp = SafeLog(haz)
            Else
                EvalExp = haz
            End If
        
        Case "ch"
            If x <= 0# Then
                cumHaz = 0#
            Else
                cumHaz = rate_ * x
            End If
            If logFlag Then
                EvalExp = SafeLog(cumHaz)
            Else
                EvalExp = cumHaz
            End If
        
        Case "s"
            If x <= 0# Then
                surv = 1#
            Else
                surv = Exp(-rate_ * x)
            End If
            If logFlag Then
                EvalExp = SafeLog(surv)
            Else
                EvalExp = surv
            End If
        
        Case Else
            EvalExp = CVErr(xlErrValue)
    End Select
End Function

'==========================================================
' Weibull (R-style / AFT): shape, scale
'==========================================================

' EvalExp
' Evaluates exponential density, probability, quantile, hazard, cumulative hazard, or survival
' Core engine for exponential worksheet functions
Private Function EvalWeibull(ByVal x As Double, ByVal mode As String, _
                             ByVal shape_ As Double, ByVal scale_ As Double, _
                             ByVal lowerTail As Boolean, _
                             ByVal logP As Boolean, _
                             ByVal logFlag As Boolean) As Variant
    Dim z As Double, dens As Double, cdf As Double, surv As Double
    Dim haz As Double, cumHaz As Double, p As Double
    
    If shape_ <= 0# Then Err.Raise vbObjectError + 2100, , "shape must be > 0"
    If scale_ <= 0# Then Err.Raise vbObjectError + 2101, , "scale must be > 0"
    
    Select Case LCase$(mode)
        Case "d"
            If x <= 0# Then
                If x = 0# And shape_ = 1# Then
                    dens = 1# / scale_
                Else
                    dens = 0#
                End If
            Else
                z = x / scale_
                dens = (shape_ / scale_) * z ^ (shape_ - 1#) * Exp(-(z ^ shape_))
            End If
            If logFlag Then
                EvalWeibull = SafeLog(dens)
            Else
                EvalWeibull = dens
            End If
        
        Case "p"
            If x <= 0# Then
                cdf = 0#
            Else
                z = x / scale_
                cdf = 1# - Exp(-(z ^ shape_))
            End If
            EvalWeibull = FinishProb(cdf, lowerTail, logP)
        
        Case "q"
            p = DecodeProb(x, lowerTail, logP)
            If p < 0# Or p > 1# Then Err.Raise vbObjectError + 2102, , "p must be in [0,1]"
            If p <= 0# Then
                EvalWeibull = 0#
            ElseIf p >= 1# Then
                EvalWeibull = BIG_POS
            Else
                EvalWeibull = scale_ * (-Log(1# - p)) ^ (1# / shape_)
            End If
        
        Case "h"
            If x <= 0# Then
                haz = 0#
            Else
                z = x / scale_
                haz = (shape_ / scale_) * z ^ (shape_ - 1#)
            End If
            If logFlag Then
                EvalWeibull = SafeLog(haz)
            Else
                EvalWeibull = haz
            End If
        
        Case "ch"
            If x <= 0# Then
                cumHaz = 0#
            Else
                z = x / scale_
                cumHaz = z ^ shape_
            End If
            If logFlag Then
                EvalWeibull = SafeLog(cumHaz)
            Else
                EvalWeibull = cumHaz
            End If
        
        Case "s"
            If x <= 0# Then
                surv = 1#
            Else
                z = x / scale_
                surv = Exp(-(z ^ shape_))
            End If
            If logFlag Then
                EvalWeibull = SafeLog(surv)
            Else
                EvalWeibull = surv
            End If
        
        Case Else
            EvalWeibull = CVErr(xlErrValue)
    End Select
End Function

'==========================================================
' WeibullPH: shape, scale
' hazard = shape * scale * x^(shape-1)
' cumhaz = scale * x^shape
'==========================================================

' EvalWeibullPH
' Evaluates Weibull proportional hazards parameterization
' Supports PH-form Weibull calculations for survival modeling
Private Function EvalWeibullPH(ByVal x As Double, ByVal mode As String, _
                               ByVal shape_ As Double, ByVal scale_ As Double, _
                               ByVal lowerTail As Boolean, _
                               ByVal logP As Boolean, _
                               ByVal logFlag As Boolean) As Variant
    Dim dens As Double, cdf As Double, surv As Double
    Dim haz As Double, cumHaz As Double, p As Double
    
    If shape_ <= 0# Then Err.Raise vbObjectError + 2200, , "shape must be > 0"
    If scale_ <= 0# Then Err.Raise vbObjectError + 2201, , "scale must be > 0"
    
    Select Case LCase$(mode)
        Case "d"
            If x <= 0# Then
                dens = 0#
            Else
                dens = shape_ * scale_ * x ^ (shape_ - 1#) * Exp(-scale_ * x ^ shape_)
            End If
            If logFlag Then
                EvalWeibullPH = SafeLog(dens)
            Else
                EvalWeibullPH = dens
            End If
        
        Case "p"
            If x <= 0# Then
                cdf = 0#
            Else
                cdf = 1# - Exp(-scale_ * x ^ shape_)
            End If
            EvalWeibullPH = FinishProb(cdf, lowerTail, logP)
        
        Case "q"
            p = DecodeProb(x, lowerTail, logP)
            If p < 0# Or p > 1# Then Err.Raise vbObjectError + 2202, , "p must be in [0,1]"
            If p <= 0# Then
                EvalWeibullPH = 0#
            ElseIf p >= 1# Then
                EvalWeibullPH = BIG_POS
            Else
                EvalWeibullPH = (-Log(1# - p) / scale_) ^ (1# / shape_)
            End If
        
        Case "h"
            If x <= 0# Then
                haz = 0#
            Else
                haz = shape_ * scale_ * x ^ (shape_ - 1#)
            End If
            If logFlag Then
                EvalWeibullPH = SafeLog(haz)
            Else
                EvalWeibullPH = haz
            End If
        
        Case "ch"
            If x <= 0# Then
                cumHaz = 0#
            Else
                cumHaz = scale_ * x ^ shape_
            End If
            If logFlag Then
                EvalWeibullPH = SafeLog(cumHaz)
            Else
                EvalWeibullPH = cumHaz
            End If
        
        Case "s"
            If x <= 0# Then
                surv = 1#
            Else
                surv = Exp(-scale_ * x ^ shape_)
            End If
            If logFlag Then
                EvalWeibullPH = SafeLog(surv)
            Else
                EvalWeibullPH = surv
            End If
        
        Case Else
            EvalWeibullPH = CVErr(xlErrValue)
    End Select
End Function

'==========================================================
' Gompertz: shape, rate
' h(x) = rate * exp(shape*x)
' H(x) = rate/shape * (exp(shape*x)-1), if shape <> 0
'==========================================================

' GompertzCumHaz
' Computes Gompertz cumulative hazard
' Central helper used by Gompertz survival functions
Private Function GompertzCumHaz(ByVal x As Double, ByVal shape_ As Double, ByVal rate_ As Double) As Double
    If x <= 0# Then
        GompertzCumHaz = 0#
    ElseIf shape_ = 0# Then
        GompertzCumHaz = rate_ * x
    Else
        GompertzCumHaz = (rate_ / shape_) * (Exp(shape_ * x) - 1#)
    End If
End Function

' EvalGompertz
' Evaluates Gompertz density, probability, quantile, hazard, cumulative hazard, or survival
' Core engine for Gompertz worksheet functions
Private Function EvalGompertz(ByVal x As Double, ByVal mode As String, _
                              ByVal shape_ As Double, ByVal rate_ As Double, _
                              ByVal lowerTail As Boolean, _
                              ByVal logP As Boolean, _
                              ByVal logFlag As Boolean) As Variant
    Dim dens As Double, cdf As Double, surv As Double
    Dim haz As Double, cumHaz As Double, p As Double
    Dim pmax As Double, argVal As Double
    
    If rate_ <= 0# Then Err.Raise vbObjectError + 2300, , "rate must be > 0"
    
    Select Case LCase$(mode)
        Case "d"
            If x <= 0# Then
                dens = 0#
            Else
                haz = rate_ * Exp(shape_ * x)
                cumHaz = GompertzCumHaz(x, shape_, rate_)
                dens = haz * Exp(-cumHaz)
            End If
            If logFlag Then
                EvalGompertz = SafeLog(dens)
            Else
                EvalGompertz = dens
            End If
        
        Case "p"
            If x <= 0# Then
                cdf = 0#
            Else
                cumHaz = GompertzCumHaz(x, shape_, rate_)
                cdf = 1# - Exp(-cumHaz)
            End If
            EvalGompertz = FinishProb(cdf, lowerTail, logP)
        
        Case "q"
            p = DecodeProb(x, lowerTail, logP)
            If p < 0# Or p > 1# Then Err.Raise vbObjectError + 2301, , "p must be in [0,1]"
            
            If shape_ < 0# Then
                pmax = 1# - Exp(rate_ / shape_)
                If p >= pmax Then
                    EvalGompertz = BIG_POS
                    Exit Function
                End If
            End If
            
            If p <= 0# Then
                EvalGompertz = 0#
            ElseIf p >= 1# Then
                EvalGompertz = BIG_POS
            ElseIf shape_ = 0# Then
                EvalGompertz = -Log(1# - p) / rate_
            Else
                argVal = 1# + (shape_ / rate_) * (-Log(1# - p))
                If argVal <= 0# Then
                    EvalGompertz = BIG_POS
                Else
                    EvalGompertz = Log(argVal) / shape_
                End If
            End If
        
        Case "h"
            If x <= 0# Then
                haz = 0#
            Else
                haz = rate_ * Exp(shape_ * x)
            End If
            If logFlag Then
                EvalGompertz = SafeLog(haz)
            Else
                EvalGompertz = haz
            End If
        
        Case "ch"
            cumHaz = GompertzCumHaz(x, shape_, rate_)
            If logFlag Then
                EvalGompertz = SafeLog(cumHaz)
            Else
                EvalGompertz = cumHaz
            End If
        
        Case "s"
            If x <= 0# Then
                surv = 1#
            Else
                cumHaz = GompertzCumHaz(x, shape_, rate_)
                surv = Exp(-cumHaz)
            End If
            If logFlag Then
                EvalGompertz = SafeLog(surv)
            Else
                EvalGompertz = surv
            End If
        
        Case Else
            EvalGompertz = CVErr(xlErrValue)
    End Select
End Function

'==========================================================
' Log-normal: meanlog, sdlog
'==========================================================

' EvalLNorm
' Evaluates log-normal density, probability, quantile, hazard, cumulative hazard, or survival
' Core engine for log-normal worksheet functions
Private Function EvalLNorm(ByVal x As Double, ByVal mode As String, _
                           ByVal meanlog_ As Double, ByVal sdlog_ As Double, _
                           ByVal lowerTail As Boolean, _
                           ByVal logP As Boolean, _
                           ByVal logFlag As Boolean) As Variant
    Dim z As Double, dens As Double, cdf As Double, surv As Double
    Dim haz As Double, cumHaz As Double, p As Double
    
    If sdlog_ <= 0# Then Err.Raise vbObjectError + 2400, , "sdlog must be > 0"
    
    Select Case LCase$(mode)
        Case "d"
            If x <= 0# Then
                dens = 0#
            Else
                z = (Log(x) - meanlog_) / sdlog_
                dens = StdNormPDF(z) / (x * sdlog_)
            End If
            If logFlag Then
                EvalLNorm = SafeLog(dens)
            Else
                EvalLNorm = dens
            End If
        
        Case "p"
            If x <= 0# Then
                cdf = 0#
            Else
                z = (Log(x) - meanlog_) / sdlog_
                cdf = StdNormCDF(z)
            End If
            EvalLNorm = FinishProb(cdf, lowerTail, logP)
        
        Case "q"
            p = DecodeProb(x, lowerTail, logP)
            If p < 0# Or p > 1# Then Err.Raise vbObjectError + 2401, , "p must be in [0,1]"
            If p <= 0# Then
                EvalLNorm = 0#
            ElseIf p >= 1# Then
                EvalLNorm = BIG_POS
            Else
                EvalLNorm = Exp(meanlog_ + sdlog_ * StdNormInv(p))
            End If
        
        Case "h"
            If x <= 0# Then
                haz = 0#
            Else
                z = (Log(x) - meanlog_) / sdlog_
                surv = 1# - StdNormCDF(z)
                If surv <= 0# Then
                    haz = BIG_POS
                Else
                    dens = StdNormPDF(z) / (x * sdlog_)
                    haz = dens / surv
                End If
            End If
            If logFlag Then
                EvalLNorm = SafeLog(haz)
            Else
                EvalLNorm = haz
            End If
        
        Case "ch"
            If x <= 0# Then
                cumHaz = 0#
            Else
                z = (Log(x) - meanlog_) / sdlog_
                surv = 1# - StdNormCDF(z)
                If surv <= 0# Then
                    cumHaz = BIG_POS
                Else
                    cumHaz = -Log(surv)
                End If
            End If
            If logFlag Then
                EvalLNorm = SafeLog(cumHaz)
            Else
                EvalLNorm = cumHaz
            End If
        
        Case "s"
            If x <= 0# Then
                surv = 1#
            Else
                z = (Log(x) - meanlog_) / sdlog_
                surv = 1# - StdNormCDF(z)
            End If
            If logFlag Then
                EvalLNorm = SafeLog(surv)
            Else
                EvalLNorm = surv
            End If
        
        Case Else
            EvalLNorm = CVErr(xlErrValue)
    End Select
End Function

'==========================================================
' Log-logistic: shape, scale
'==========================================================

' EvalLLogis
' Evaluates log-logistic density, probability, quantile, hazard, cumulative hazard, or survival
' Core engine for log-logistic worksheet functions
Private Function EvalLLogis(ByVal x As Double, ByVal mode As String, _
                            ByVal shape_ As Double, ByVal scale_ As Double, _
                            ByVal lowerTail As Boolean, _
                            ByVal logP As Boolean, _
                            ByVal logFlag As Boolean) As Variant
    Dim z As Double, dens As Double, cdf As Double, surv As Double
    Dim haz As Double, cumHaz As Double, p As Double
    
    If shape_ <= 0# Then Err.Raise vbObjectError + 2500, , "shape must be > 0"
    If scale_ <= 0# Then Err.Raise vbObjectError + 2501, , "scale must be > 0"
    
    Select Case LCase$(mode)
        Case "d"
            If x <= 0# Then
                dens = 0#
            Else
                z = x / scale_
                dens = (shape_ / scale_) * z ^ (shape_ - 1#) / ((1# + z ^ shape_) ^ 2)
            End If
            If logFlag Then
                EvalLLogis = SafeLog(dens)
            Else
                EvalLLogis = dens
            End If
        
        Case "p"
            If x <= 0# Then
                cdf = 0#
            Else
                z = x / scale_
                cdf = (z ^ shape_) / (1# + z ^ shape_)
            End If
            EvalLLogis = FinishProb(cdf, lowerTail, logP)
        
        Case "q"
            p = DecodeProb(x, lowerTail, logP)
            If p < 0# Or p > 1# Then Err.Raise vbObjectError + 2502, , "p must be in [0,1]"
            If p <= 0# Then
                EvalLLogis = 0#
            ElseIf p >= 1# Then
                EvalLLogis = BIG_POS
            Else
                EvalLLogis = scale_ * (p / (1# - p)) ^ (1# / shape_)
            End If
        
        Case "h"
            If x <= 0# Then
                haz = 0#
            Else
                z = x / scale_
                haz = (shape_ / scale_) * z ^ (shape_ - 1#) / (1# + z ^ shape_)
            End If
            If logFlag Then
                EvalLLogis = SafeLog(haz)
            Else
                EvalLLogis = haz
            End If
        
        Case "ch"
            If x <= 0# Then
                cumHaz = 0#
            Else
                z = x / scale_
                cumHaz = Log(1# + z ^ shape_)
            End If
            If logFlag Then
                EvalLLogis = SafeLog(cumHaz)
            Else
                EvalLLogis = cumHaz
            End If
        
        Case "s"
            If x <= 0# Then
                surv = 1#
            Else
                z = x / scale_
                surv = 1# / (1# + z ^ shape_)
            End If
            If logFlag Then
                EvalLLogis = SafeLog(surv)
            Else
                EvalLLogis = surv
            End If
        
        Case Else
            EvalLLogis = CVErr(xlErrValue)
    End Select
End Function

'==========================================================
' Gamma: shape, rate
' Uses Excel GAMMA.DIST and GAMMA.INV
'==========================================================

' EvalGammaDist
' Evaluates gamma density, probability, quantile, hazard, cumulative hazard, or survival
' Uses Excel gamma functions for core gamma distribution calculations
Private Function EvalGammaDist(ByVal x As Double, ByVal mode As String, _
                               ByVal shape_ As Double, ByVal rate_ As Double, _
                               ByVal lowerTail As Boolean, _
                               ByVal logP As Boolean, _
                               ByVal logFlag As Boolean) As Variant
    Dim dens As Double, cdf As Double, surv As Double
    Dim haz As Double, cumHaz As Double, p As Double
    Dim scale_ As Double
    
    If shape_ <= 0# Then Err.Raise vbObjectError + 2600, , "shape must be > 0"
    If rate_ <= 0# Then Err.Raise vbObjectError + 2601, , "rate must be > 0"
    
    scale_ = 1# / rate_
    
    Select Case LCase$(mode)
        Case "d"
            If x <= 0# Then
                dens = 0#
            Else
                dens = Application.WorksheetFunction.Gamma_Dist(x, shape_, scale_, False)
            End If
            If logFlag Then
                EvalGammaDist = SafeLog(dens)
            Else
                EvalGammaDist = dens
            End If
        
        Case "p"
            If x <= 0# Then
                cdf = 0#
            Else
                cdf = Application.WorksheetFunction.Gamma_Dist(x, shape_, scale_, True)
            End If
            EvalGammaDist = FinishProb(cdf, lowerTail, logP)
        
        Case "q"
            p = DecodeProb(x, lowerTail, logP)
            If p < 0# Or p > 1# Then Err.Raise vbObjectError + 2602, , "p must be in [0,1]"
            If p <= 0# Then
                EvalGammaDist = 0#
            ElseIf p >= 1# Then
                EvalGammaDist = BIG_POS
            Else
                EvalGammaDist = Application.WorksheetFunction.Gamma_Inv(p, shape_, scale_)
            End If
        
        Case "h"
            If x <= 0# Then
                haz = 0#
            Else
                dens = Application.WorksheetFunction.Gamma_Dist(x, shape_, scale_, False)
                surv = 1# - Application.WorksheetFunction.Gamma_Dist(x, shape_, scale_, True)
                If surv <= 0# Then
                    haz = BIG_POS
                Else
                    haz = dens / surv
                End If
            End If
            If logFlag Then
                EvalGammaDist = SafeLog(haz)
            Else
                EvalGammaDist = haz
            End If
        
        Case "ch"
            If x <= 0# Then
                cumHaz = 0#
            Else
                surv = 1# - Application.WorksheetFunction.Gamma_Dist(x, shape_, scale_, True)
                If surv <= 0# Then
                    cumHaz = BIG_POS
                Else
                    cumHaz = -Log(surv)
                End If
            End If
            If logFlag Then
                EvalGammaDist = SafeLog(cumHaz)
            Else
                EvalGammaDist = cumHaz
            End If
        
        Case "s"
            If x <= 0# Then
                surv = 1#
            Else
                surv = 1# - Application.WorksheetFunction.Gamma_Dist(x, shape_, scale_, True)
            End If
            If logFlag Then
                EvalGammaDist = SafeLog(surv)
            Else
                EvalGammaDist = surv
            End If
        
        Case Else
            EvalGammaDist = CVErr(xlErrValue)
    End Select
End Function

'==========================================================
' Dispatcher
'==========================================================

' EvalOne
' Dispatches evaluation to the correct distribution engine
' Central routing function for all scalar distribution calculations
Private Function EvalOne(ByVal x As Double, _
                         ByVal distCode As Long, _
                         ByVal mode As String, _
                         ByVal p1 As Double, _
                         ByVal p2 As Double, _
                         ByVal lowerTail As Boolean, _
                         ByVal logP As Boolean, _
                         ByVal logFlag As Boolean) As Variant
    On Error GoTo EH
    
    Select Case distCode
        Case DIST_EXP
            EvalOne = EvalExp(x, mode, p1, lowerTail, logP, logFlag)
        Case DIST_WEIBULL
            EvalOne = EvalWeibull(x, mode, p1, p2, lowerTail, logP, logFlag)
        Case DIST_WEIBULLPH
            EvalOne = EvalWeibullPH(x, mode, p1, p2, lowerTail, logP, logFlag)
        Case DIST_GOMPERTZ
            EvalOne = EvalGompertz(x, mode, p1, p2, lowerTail, logP, logFlag)
        Case DIST_LNORM
            EvalOne = EvalLNorm(x, mode, p1, p2, lowerTail, logP, logFlag)
        Case DIST_LLOGIS
            EvalOne = EvalLLogis(x, mode, p1, p2, lowerTail, logP, logFlag)
        Case DIST_GAMMA
            EvalOne = EvalGammaDist(x, mode, p1, p2, lowerTail, logP, logFlag)
        Case Else
            EvalOne = CVErr(xlErrValue)
    End Select
    
    Exit Function
    
EH:
    EvalOne = CVErr(xlErrValue)
End Function

' EvalRandom
' Generates one random draw from the selected distribution using inverse-CDF logic
' Central routing function for random sampling
Private Function EvalRandom(ByVal distCode As Long, _
                            ByVal p1 As Double, _
                            ByVal p2 As Double) As Variant
    Dim u As Double
    
    u = UniformOpen01()
    
    Select Case distCode
        Case DIST_EXP
            EvalRandom = EvalExp(u, "q", p1, True, False, False)
        Case DIST_WEIBULL
            EvalRandom = EvalWeibull(u, "q", p1, p2, True, False, False)
        Case DIST_WEIBULLPH
            EvalRandom = EvalWeibullPH(u, "q", p1, p2, True, False, False)
        Case DIST_GOMPERTZ
            EvalRandom = EvalGompertz(u, "q", p1, p2, True, False, False)
        Case DIST_LNORM
            EvalRandom = EvalLNorm(u, "q", p1, p2, True, False, False)
        Case DIST_LLOGIS
            EvalRandom = EvalLLogis(u, "q", p1, p2, True, False, False)
        Case DIST_GAMMA
            EvalRandom = EvalGammaDist(u, "q", p1, p2, True, False, False)
        Case Else
            EvalRandom = CVErr(xlErrValue)
    End Select
End Function

'==========================================================
' Public worksheet functions: Exponential
'==========================================================

Public Function dexp(ByVal x As Variant, ByVal rate_ As Double, _
                     Optional ByVal logFlag As Boolean = False) As Variant
    dexp = ScalarOrArray(x, DIST_EXP, "d", rate_, 0#, True, False, logFlag)
End Function

Public Function pexp(ByVal q As Variant, ByVal rate_ As Double, _
                     Optional ByVal lowerTail As Boolean = True, _
                     Optional ByVal logP As Boolean = False) As Variant
    pexp = ScalarOrArray(q, DIST_EXP, "p", rate_, 0#, lowerTail, logP, False)
End Function

Public Function qexp(ByVal p As Variant, ByVal rate_ As Double, _
                     Optional ByVal lowerTail As Boolean = True, _
                     Optional ByVal logP As Boolean = False) As Variant
    qexp = ScalarOrArray(p, DIST_EXP, "q", rate_, 0#, lowerTail, logP, False)
End Function

Public Function hexp(ByVal x As Variant, ByVal rate_ As Double, _
                     Optional ByVal logFlag As Boolean = False) As Variant
    hexp = ScalarOrArray(x, DIST_EXP, "h", rate_, 0#, True, False, logFlag)
End Function

Public Function chexp(ByVal x As Variant, ByVal rate_ As Double, _
                      Optional ByVal logFlag As Boolean = False) As Variant
    chexp = ScalarOrArray(x, DIST_EXP, "ch", rate_, 0#, True, False, logFlag)
End Function

Public Function Sexp(ByVal x As Variant, ByVal rate_ As Double, _
                     Optional ByVal logFlag As Boolean = False) As Variant
    Sexp = ScalarOrArray(x, DIST_EXP, "s", rate_, 0#, True, False, logFlag)
End Function

Public Function rexp(Optional ByVal n As Long = 1, _
                     Optional ByVal rate_ As Double = 1#) As Variant
    rexp = RandomVector(n, DIST_EXP, rate_, 0#)
End Function

'==========================================================
' Public worksheet functions: Weibull
'==========================================================

Public Function dweibull(ByVal x As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                         Optional ByVal logFlag As Boolean = False) As Variant
    dweibull = ScalarOrArray(x, DIST_WEIBULL, "d", shape_, scale_, True, False, logFlag)
End Function

Public Function pweibull(ByVal q As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                         Optional ByVal lowerTail As Boolean = True, _
                         Optional ByVal logP As Boolean = False) As Variant
    pweibull = ScalarOrArray(q, DIST_WEIBULL, "p", shape_, scale_, lowerTail, logP, False)
End Function

Public Function qweibull(ByVal p As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                         Optional ByVal lowerTail As Boolean = True, _
                         Optional ByVal logP As Boolean = False) As Variant
    qweibull = ScalarOrArray(p, DIST_WEIBULL, "q", shape_, scale_, lowerTail, logP, False)
End Function

Public Function hweibull(ByVal x As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                         Optional ByVal logFlag As Boolean = False) As Variant
    hweibull = ScalarOrArray(x, DIST_WEIBULL, "h", shape_, scale_, True, False, logFlag)
End Function

Public Function chweibull(ByVal x As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    chweibull = ScalarOrArray(x, DIST_WEIBULL, "ch", shape_, scale_, True, False, logFlag)
End Function

Public Function Sweibull(ByVal x As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                         Optional ByVal logFlag As Boolean = False) As Variant
    Sweibull = ScalarOrArray(x, DIST_WEIBULL, "s", shape_, scale_, True, False, logFlag)
End Function

Public Function rweibull(Optional ByVal n As Long = 1, _
                         Optional ByVal shape_ As Double = 1#, _
                         Optional ByVal scale_ As Double = 1#) As Variant
    rweibull = RandomVector(n, DIST_WEIBULL, shape_, scale_)
End Function

'==========================================================
' Public worksheet functions: WeibullPH
'==========================================================

Public Function dweibullPH(ByVal x As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                           Optional ByVal logFlag As Boolean = False) As Variant
    dweibullPH = ScalarOrArray(x, DIST_WEIBULLPH, "d", shape_, scale_, True, False, logFlag)
End Function

Public Function pweibullPH(ByVal q As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                           Optional ByVal lowerTail As Boolean = True, _
                           Optional ByVal logP As Boolean = False) As Variant
    pweibullPH = ScalarOrArray(q, DIST_WEIBULLPH, "p", shape_, scale_, lowerTail, logP, False)
End Function

Public Function qweibullPH(ByVal p As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                           Optional ByVal lowerTail As Boolean = True, _
                           Optional ByVal logP As Boolean = False) As Variant
    qweibullPH = ScalarOrArray(p, DIST_WEIBULLPH, "q", shape_, scale_, lowerTail, logP, False)
End Function

Public Function hweibullPH(ByVal x As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                           Optional ByVal logFlag As Boolean = False) As Variant
    hweibullPH = ScalarOrArray(x, DIST_WEIBULLPH, "h", shape_, scale_, True, False, logFlag)
End Function

Public Function chweibullPH(ByVal x As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                            Optional ByVal logFlag As Boolean = False) As Variant
    chweibullPH = ScalarOrArray(x, DIST_WEIBULLPH, "ch", shape_, scale_, True, False, logFlag)
End Function

Public Function SweibullPH(ByVal x As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                           Optional ByVal logFlag As Boolean = False) As Variant
    SweibullPH = ScalarOrArray(x, DIST_WEIBULLPH, "s", shape_, scale_, True, False, logFlag)
End Function

Public Function rweibullPH(Optional ByVal n As Long = 1, _
                           Optional ByVal shape_ As Double = 1#, _
                           Optional ByVal scale_ As Double = 1#) As Variant
    rweibullPH = RandomVector(n, DIST_WEIBULLPH, shape_, scale_)
End Function

'==========================================================
' Public worksheet functions: Gompertz
'==========================================================

Public Function dgompertz(ByVal x As Variant, ByVal shape_ As Double, ByVal rate_ As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    dgompertz = ScalarOrArray(x, DIST_GOMPERTZ, "d", shape_, rate_, True, False, logFlag)
End Function

Public Function pgompertz(ByVal q As Variant, ByVal shape_ As Double, ByVal rate_ As Double, _
                          Optional ByVal lowerTail As Boolean = True, _
                          Optional ByVal logP As Boolean = False) As Variant
    pgompertz = ScalarOrArray(q, DIST_GOMPERTZ, "p", shape_, rate_, lowerTail, logP, False)
End Function

Public Function qgompertz(ByVal p As Variant, ByVal shape_ As Double, ByVal rate_ As Double, _
                          Optional ByVal lowerTail As Boolean = True, _
                          Optional ByVal logP As Boolean = False) As Variant
    qgompertz = ScalarOrArray(p, DIST_GOMPERTZ, "q", shape_, rate_, lowerTail, logP, False)
End Function

Public Function hgompertz(ByVal x As Variant, ByVal shape_ As Double, ByVal rate_ As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    hgompertz = ScalarOrArray(x, DIST_GOMPERTZ, "h", shape_, rate_, True, False, logFlag)
End Function

Public Function chgompertz(ByVal x As Variant, ByVal shape_ As Double, ByVal rate_ As Double, _
                           Optional ByVal logFlag As Boolean = False) As Variant
    chgompertz = ScalarOrArray(x, DIST_GOMPERTZ, "ch", shape_, rate_, True, False, logFlag)
End Function

Public Function Sgompertz(ByVal x As Variant, ByVal shape_ As Double, ByVal rate_ As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    Sgompertz = ScalarOrArray(x, DIST_GOMPERTZ, "s", shape_, rate_, True, False, logFlag)
End Function

Public Function rgompertz(Optional ByVal n As Long = 1, _
                          Optional ByVal shape_ As Double = 0.1, _
                          Optional ByVal rate_ As Double = 1#) As Variant
    rgompertz = RandomVector(n, DIST_GOMPERTZ, shape_, rate_)
End Function

'==========================================================
' Public worksheet functions: Log-normal
'==========================================================

Public Function dlnorm(ByVal x As Variant, ByVal meanlog_ As Double, ByVal sdlog_ As Double, _
                       Optional ByVal logFlag As Boolean = False) As Variant
    dlnorm = ScalarOrArray(x, DIST_LNORM, "d", meanlog_, sdlog_, True, False, logFlag)
End Function

Public Function plnorm(ByVal q As Variant, ByVal meanlog_ As Double, ByVal sdlog_ As Double, _
                       Optional ByVal lowerTail As Boolean = True, _
                       Optional ByVal logP As Boolean = False) As Variant
    plnorm = ScalarOrArray(q, DIST_LNORM, "p", meanlog_, sdlog_, lowerTail, logP, False)
End Function

Public Function qlnorm(ByVal p As Variant, ByVal meanlog_ As Double, ByVal sdlog_ As Double, _
                       Optional ByVal lowerTail As Boolean = True, _
                       Optional ByVal logP As Boolean = False) As Variant
    qlnorm = ScalarOrArray(p, DIST_LNORM, "q", meanlog_, sdlog_, lowerTail, logP, False)
End Function

Public Function hlnorm(ByVal x As Variant, ByVal meanlog_ As Double, ByVal sdlog_ As Double, _
                       Optional ByVal logFlag As Boolean = False) As Variant
    hlnorm = ScalarOrArray(x, DIST_LNORM, "h", meanlog_, sdlog_, True, False, logFlag)
End Function

Public Function chlnorm(ByVal x As Variant, ByVal meanlog_ As Double, ByVal sdlog_ As Double, _
                        Optional ByVal logFlag As Boolean = False) As Variant
    chlnorm = ScalarOrArray(x, DIST_LNORM, "ch", meanlog_, sdlog_, True, False, logFlag)
End Function

Public Function Slnorm(ByVal x As Variant, ByVal meanlog_ As Double, ByVal sdlog_ As Double, _
                       Optional ByVal logFlag As Boolean = False) As Variant
    Slnorm = ScalarOrArray(x, DIST_LNORM, "s", meanlog_, sdlog_, True, False, logFlag)
End Function

Public Function rlnorm(Optional ByVal n As Long = 1, _
                       Optional ByVal meanlog_ As Double = 0#, _
                       Optional ByVal sdlog_ As Double = 1#) As Variant
    rlnorm = RandomVector(n, DIST_LNORM, meanlog_, sdlog_)
End Function

'==========================================================
' Public worksheet functions: Log-logistic
'==========================================================

Public Function dllogis(ByVal x As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                        Optional ByVal logFlag As Boolean = False) As Variant
    dllogis = ScalarOrArray(x, DIST_LLOGIS, "d", shape_, scale_, True, False, logFlag)
End Function

Public Function pllogis(ByVal q As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                        Optional ByVal lowerTail As Boolean = True, _
                        Optional ByVal logP As Boolean = False) As Variant
    pllogis = ScalarOrArray(q, DIST_LLOGIS, "p", shape_, scale_, lowerTail, logP, False)
End Function

Public Function qllogis(ByVal p As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                        Optional ByVal lowerTail As Boolean = True, _
                        Optional ByVal logP As Boolean = False) As Variant
    qllogis = ScalarOrArray(p, DIST_LLOGIS, "q", shape_, scale_, lowerTail, logP, False)
End Function

Public Function hllogis(ByVal x As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                        Optional ByVal logFlag As Boolean = False) As Variant
    hllogis = ScalarOrArray(x, DIST_LLOGIS, "h", shape_, scale_, True, False, logFlag)
End Function

Public Function chllogis(ByVal x As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                         Optional ByVal logFlag As Boolean = False) As Variant
    chllogis = ScalarOrArray(x, DIST_LLOGIS, "ch", shape_, scale_, True, False, logFlag)
End Function

Public Function Sllogis(ByVal x As Variant, ByVal shape_ As Double, ByVal scale_ As Double, _
                        Optional ByVal logFlag As Boolean = False) As Variant
    Sllogis = ScalarOrArray(x, DIST_LLOGIS, "s", shape_, scale_, True, False, logFlag)
End Function

Public Function rllogis(Optional ByVal n As Long = 1, _
                        Optional ByVal shape_ As Double = 1#, _
                        Optional ByVal scale_ As Double = 1#) As Variant
    rllogis = RandomVector(n, DIST_LLOGIS, shape_, scale_)
End Function

'==========================================================
' Public worksheet functions: Gamma
'==========================================================

Public Function dgamma(ByVal x As Variant, ByVal shape_ As Double, ByVal rate_ As Double, _
                       Optional ByVal logFlag As Boolean = False) As Variant
    dgamma = ScalarOrArray(x, DIST_GAMMA, "d", shape_, rate_, True, False, logFlag)
End Function

Public Function pgamma(ByVal q As Variant, ByVal shape_ As Double, ByVal rate_ As Double, _
                       Optional ByVal lowerTail As Boolean = True, _
                       Optional ByVal logP As Boolean = False) As Variant
    pgamma = ScalarOrArray(q, DIST_GAMMA, "p", shape_, rate_, lowerTail, logP, False)
End Function

Public Function qgamma(ByVal p As Variant, ByVal shape_ As Double, ByVal rate_ As Double, _
                       Optional ByVal lowerTail As Boolean = True, _
                       Optional ByVal logP As Boolean = False) As Variant
    qgamma = ScalarOrArray(p, DIST_GAMMA, "q", shape_, rate_, lowerTail, logP, False)
End Function

Public Function hgamma(ByVal x As Variant, ByVal shape_ As Double, ByVal rate_ As Double, _
                       Optional ByVal logFlag As Boolean = False) As Variant
    hgamma = ScalarOrArray(x, DIST_GAMMA, "h", shape_, rate_, True, False, logFlag)
End Function

Public Function chgamma(ByVal x As Variant, ByVal shape_ As Double, ByVal rate_ As Double, _
                        Optional ByVal logFlag As Boolean = False) As Variant
    chgamma = ScalarOrArray(x, DIST_GAMMA, "ch", shape_, rate_, True, False, logFlag)
End Function

Public Function Sgamma(ByVal x As Variant, ByVal shape_ As Double, ByVal rate_ As Double, _
                       Optional ByVal logFlag As Boolean = False) As Variant
    Sgamma = ScalarOrArray(x, DIST_GAMMA, "s", shape_, rate_, True, False, logFlag)
End Function

Public Function rgamma(Optional ByVal n As Long = 1, _
                       Optional ByVal shape_ As Double = 1#, _
                       Optional ByVal rate_ As Double = 1#) As Variant
    rgamma = RandomVector(n, DIST_GAMMA, shape_, rate_)
End Function

'Functions that require incomplete gamma, incomplete beta, or both

'=========================================================================
' Assumes these helpers already exist in your project:
'
'   SafeExp(x As Double) As Double
'   SafeLog(x As Double) As Double
'   MaxDbl(a As Double, b As Double) As Double
'   UniformOpen01() As Double
'   BIG_POS, BIG_NEG
'   DEFAULT_TOL, DEFAULT_MAXITER
'
'   LogGammaVBA(x) As Double
'   LogBetaVBA(a, b) As Double
'   GammaP(a, x) As Variant
'   GammaQ(a, x) As Variant
'   BetaI(x, a, b) As Variant
'
'   dlnorm / plnorm / qlnorm / hlnorm / chlnorm / Slnorm
'
'=========================================================================
' Naming notes:
'   - cumulative hazard uses "ch" because VBA is case-insensitive
'   - ".orig" becomes "_orig" because VBA identifiers cannot contain "."
'=========================================================================

'==========================================================
' Helpers
'==========================================================

' DecodeProbFlex
' Converts flexible-model probability input from log and/or upper-tail form
' Standardizes probability input for flexible quantile calculations
Private Function DecodeProbFlex(ByVal probVal As Double, _
                                ByVal lowerTail As Boolean, _
                                ByVal logP As Boolean) As Double
    Dim pWork As Double

    If logP Then
        pWork = Exp(probVal)
    Else
        pWork = probVal
    End If

    If Not lowerTail Then pWork = 1# - pWork
    DecodeProbFlex = pWork
End Function

' FinishProbFlex
' Applies lower-tail or upper-tail and optional log transformation to a flexible-model probability
' Standardizes probability output for flexible distributions
Private Function FinishProbFlex(ByVal cdfVal As Double, _
                                ByVal lowerTail As Boolean, _
                                ByVal logP As Boolean) As Double
    Dim outVal As Double

    If lowerTail Then
        outVal = cdfVal
    Else
        outVal = 1# - cdfVal
    End If

    If logP Then
        FinishProbFlex = SafeLog(outVal)
    Else
        FinishProbFlex = outVal
    End If
End Function

' Clamp01
' Restricts a numeric value to the interval [0, 1]
' Protects probabilities from small numerical drift
Private Function Clamp01(ByVal xVal As Double) As Double
    If xVal < 0# Then
        Clamp01 = 0#
    ElseIf xVal > 1# Then
        Clamp01 = 1#
    Else
        Clamp01 = xVal
    End If
End Function

'==========================================================
' gengamma  (Prentice parameterization)
' Parameters: muVal, sigmaVal, qParam
'==========================================================

' GenGammaY_Core
' Computes transformed generalized gamma Y value
' Used as the core argument in generalized gamma formulas
Private Function GenGammaY_Core(ByVal xVal As Double, _
                                ByVal muVal As Double, _
                                ByVal sigmaVal As Double, _
                                ByVal qParam As Double) As Double
    Dim wVal As Double

    If xVal <= 0# Then
        GenGammaY_Core = 0#
        Exit Function
    End If

    wVal = (Log(xVal) - muVal) / sigmaVal
    GenGammaY_Core = SafeExp(qParam * wVal) / (qParam * qParam)
End Function

' GenGammaCDF_Core
' Computes generalized gamma cumulative distribution function
' Handles both general case and log-normal limit when q approaches zero
Private Function GenGammaCDF_Core(ByVal xVal As Double, _
                                  ByVal muVal As Double, _
                                  ByVal sigmaVal As Double, _
                                  ByVal qParam As Double) As Double
    Dim kVal As Double
    Dim yVal As Double

    If sigmaVal <= 0# Then Err.Raise vbObjectError + 6100, , "sigma must be > 0"

    If xVal <= 0# Then
        GenGammaCDF_Core = 0#
        Exit Function
    End If

    If Abs(qParam) < FLEX_GG_Q_EPS Then
        GenGammaCDF_Core = CDbl(plnorm(xVal, muVal, sigmaVal, True, False))
        Exit Function
    End If

    kVal = 1# / (qParam * qParam)
    yVal = GenGammaY_Core(xVal, muVal, sigmaVal, qParam)

    If qParam > 0# Then
        GenGammaCDF_Core = Clamp01(CDbl(GammaP(kVal, yVal)))
    Else
        GenGammaCDF_Core = Clamp01(CDbl(GammaQ(kVal, yVal)))
    End If
End Function

' GenGammaSurv_Core
' Computes generalized gamma survival function
' Converts generalized gamma CDF to survival scale
Private Function GenGammaSurv_Core(ByVal xVal As Double, _
                                   ByVal muVal As Double, _
                                   ByVal sigmaVal As Double, _
                                   ByVal qParam As Double) As Double
    If xVal <= 0# Then
        GenGammaSurv_Core = 1#
    Else
        GenGammaSurv_Core = 1# - GenGammaCDF_Core(xVal, muVal, sigmaVal, qParam)
    End If
End Function

' GenGammaSurv_Core
' Computes generalized gamma survival function
' Converts generalized gamma CDF to survival scale
Private Function GenGammaDensity_Core(ByVal xVal As Double, _
                                      ByVal muVal As Double, _
                                      ByVal sigmaVal As Double, _
                                      ByVal qParam As Double) As Double
    Dim kVal As Double
    Dim yVal As Double
    Dim logfVal As Double

    If sigmaVal <= 0# Then Err.Raise vbObjectError + 6101, , "sigma must be > 0"

    If xVal <= 0# Then
        GenGammaDensity_Core = 0#
        Exit Function
    End If

    If Abs(qParam) < FLEX_GG_Q_EPS Then
        GenGammaDensity_Core = CDbl(dlnorm(xVal, muVal, sigmaVal, False))
        Exit Function
    End If

    kVal = 1# / (qParam * qParam)
    yVal = GenGammaY_Core(xVal, muVal, sigmaVal, qParam)

    logfVal = Log(Abs(qParam)) + kVal * Log(yVal) - yVal - Log(sigmaVal) - Log(xVal) - LogGammaVBA(kVal)
    GenGammaDensity_Core = SafeExp(logfVal)
End Function

' GenGammaHaz_Core
' Computes generalized gamma hazard
' Combines density and survival with protection against zero survival
Private Function GenGammaHaz_Core(ByVal xVal As Double, _
                                  ByVal muVal As Double, _
                                  ByVal sigmaVal As Double, _
                                  ByVal qParam As Double) As Double
    Dim survVal As Double

    If xVal <= 0# Then
        GenGammaHaz_Core = 0#
        Exit Function
    End If

    If Abs(qParam) < FLEX_GG_Q_EPS Then
        GenGammaHaz_Core = CDbl(hlnorm(xVal, muVal, sigmaVal, False))
        Exit Function
    End If

    survVal = GenGammaSurv_Core(xVal, muVal, sigmaVal, qParam)

    If survVal <= 0# Then
        GenGammaHaz_Core = BIG_POS
    Else
        GenGammaHaz_Core = GenGammaDensity_Core(xVal, muVal, sigmaVal, qParam) / survVal
    End If
End Function

' GenGammaCumHaz_Core
' Computes generalized gamma cumulative hazard
' Converts generalized gamma survival to cumulative hazard scale
Private Function GenGammaCumHaz_Core(ByVal xVal As Double, _
                                     ByVal muVal As Double, _
                                     ByVal sigmaVal As Double, _
                                     ByVal qParam As Double) As Double
    Dim survVal As Double

    If xVal <= 0# Then
        GenGammaCumHaz_Core = 0#
        Exit Function
    End If

    If Abs(qParam) < FLEX_GG_Q_EPS Then
        GenGammaCumHaz_Core = CDbl(chlnorm(xVal, muVal, sigmaVal, False))
        Exit Function
    End If

    survVal = GenGammaSurv_Core(xVal, muVal, sigmaVal, qParam)

    If survVal <= 0# Then
        GenGammaCumHaz_Core = BIG_POS
    Else
        GenGammaCumHaz_Core = -Log(survVal)
    End If
End Function

' GenGammaQuantile_Core
' Computes generalized gamma quantile using bisection search
' Finds time corresponding to a target probability
Private Function GenGammaQuantile_Core(ByVal probVal As Double, _
                                       ByVal muVal As Double, _
                                       ByVal sigmaVal As Double, _
                                       ByVal qParam As Double, _
                                       Optional ByVal lowerTail As Boolean = True, _
                                       Optional ByVal logP As Boolean = False, _
                                       Optional ByVal tol As Double = DEFAULT_TOL, _
                                       Optional ByVal maxiter As Long = DEFAULT_MAXITER) As Double
    Dim targetVal As Double
    Dim loVal As Double, hiVal As Double, midVal As Double
    Dim fmidVal As Double
    Dim iterVal As Long

    If sigmaVal <= 0# Then Err.Raise vbObjectError + 6102, , "sigma must be > 0"

    If Abs(qParam) < FLEX_GG_Q_EPS Then
        GenGammaQuantile_Core = CDbl(qlnorm(probVal, muVal, sigmaVal, lowerTail, logP))
        Exit Function
    End If

    targetVal = DecodeProbFlex(probVal, lowerTail, logP)

    If targetVal < 0# Or targetVal > 1# Then Err.Raise vbObjectError + 6103, , "p must be in [0,1]"

    If targetVal <= 0# Then
        GenGammaQuantile_Core = 0#
        Exit Function
    End If

    If targetVal >= 1# Then
        GenGammaQuantile_Core = BIG_POS
        Exit Function
    End If

    loVal = 0#
    hiVal = MaxDbl(1#, SafeExp(muVal))

    Do While GenGammaCDF_Core(hiVal, muVal, sigmaVal, qParam) < targetVal
        hiVal = hiVal * 2#
        If hiVal > 1000000000000# Then Exit Do
    Loop

    For iterVal = 1 To maxiter
        midVal = 0.5 * (loVal + hiVal)
        fmidVal = GenGammaCDF_Core(midVal, muVal, sigmaVal, qParam) - targetVal

        If Abs(fmidVal) < tol Then
            GenGammaQuantile_Core = midVal
            Exit Function
        End If

        If fmidVal >= 0# Then
            hiVal = midVal
        Else
            loVal = midVal
        End If

        If Abs(hiVal - loVal) < tol * (1# + midVal) Then
            GenGammaQuantile_Core = 0.5 * (loVal + hiVal)
            Exit Function
        End If
    Next iterVal

    GenGammaQuantile_Core = 0.5 * (loVal + hiVal)
End Function

'==========================================================
' gengamma_orig  (Stacy parameterization)
' Parameters: shapeVal, scaleVal, kVal
'==========================================================

' GenGammaOrigZ_Core
' Computes Stacy generalized gamma transformed z value
' Forms the core argument for gamma-based calculations
Private Function GenGammaOrigZ_Core(ByVal xVal As Double, _
                                    ByVal shapeVal As Double, _
                                    ByVal scaleVal As Double) As Double
    If xVal <= 0# Then
        GenGammaOrigZ_Core = 0#
    Else
        GenGammaOrigZ_Core = (xVal / scaleVal) ^ shapeVal
    End If
End Function

' GenGammaOrigCDF_Core
' Computes Stacy generalized gamma cumulative distribution function
' Uses regularized incomplete gamma for probability evaluation
Private Function GenGammaOrigCDF_Core(ByVal xVal As Double, _
                                      ByVal shapeVal As Double, _
                                      ByVal scaleVal As Double, _
                                      ByVal kVal As Double) As Double
    Dim zVal As Double

    If shapeVal <= 0# Then Err.Raise vbObjectError + 6200, , "shape must be > 0"
    If scaleVal <= 0# Then Err.Raise vbObjectError + 6201, , "scale must be > 0"
    If kVal <= 0# Then Err.Raise vbObjectError + 6202, , "k must be > 0"

    If xVal <= 0# Then
        GenGammaOrigCDF_Core = 0#
        Exit Function
    End If

    zVal = GenGammaOrigZ_Core(xVal, shapeVal, scaleVal)
    GenGammaOrigCDF_Core = Clamp01(CDbl(GammaP(kVal, zVal)))
End Function

' GenGammaOrigSurv_Core
' Computes Stacy generalized gamma survival function
' Converts CDF to survival scale
Private Function GenGammaOrigSurv_Core(ByVal xVal As Double, _
                                       ByVal shapeVal As Double, _
                                       ByVal scaleVal As Double, _
                                       ByVal kVal As Double) As Double
    If xVal <= 0# Then
        GenGammaOrigSurv_Core = 1#
    Else
        GenGammaOrigSurv_Core = 1# - GenGammaOrigCDF_Core(xVal, shapeVal, scaleVal, kVal)
    End If
End Function

' GenGammaOrigDensity_Core
' Computes Stacy generalized gamma density
' Uses log-scale formulation for numerical stability
Private Function GenGammaOrigDensity_Core(ByVal xVal As Double, _
                                          ByVal shapeVal As Double, _
                                          ByVal scaleVal As Double, _
                                          ByVal kVal As Double) As Double
    Dim zVal As Double
    Dim logfVal As Double

    If shapeVal <= 0# Then Err.Raise vbObjectError + 6203, , "shape must be > 0"
    If scaleVal <= 0# Then Err.Raise vbObjectError + 6204, , "scale must be > 0"
    If kVal <= 0# Then Err.Raise vbObjectError + 6205, , "k must be > 0"

    If xVal <= 0# Then
        GenGammaOrigDensity_Core = 0#
        Exit Function
    End If

    zVal = GenGammaOrigZ_Core(xVal, shapeVal, scaleVal)

    logfVal = Log(shapeVal) + (shapeVal * kVal - 1#) * Log(xVal) _
            - (shapeVal * kVal) * Log(scaleVal) - zVal - LogGammaVBA(kVal)

    GenGammaOrigDensity_Core = SafeExp(logfVal)
End Function

' GenGammaOrigHaz_Core
' Computes Stacy generalized gamma hazard
' Combines density and survival with protection against zero survival
Private Function GenGammaOrigHaz_Core(ByVal xVal As Double, _
                                      ByVal shapeVal As Double, _
                                      ByVal scaleVal As Double, _
                                      ByVal kVal As Double) As Double
    Dim survVal As Double

    If xVal <= 0# Then
        GenGammaOrigHaz_Core = 0#
        Exit Function
    End If

    survVal = GenGammaOrigSurv_Core(xVal, shapeVal, scaleVal, kVal)

    If survVal <= 0# Then
        GenGammaOrigHaz_Core = BIG_POS
    Else
        GenGammaOrigHaz_Core = GenGammaOrigDensity_Core(xVal, shapeVal, scaleVal, kVal) / survVal
    End If
End Function

' GenGammaOrigCumHaz_Core
' Computes Stacy generalized gamma cumulative hazard
' Converts survival to cumulative hazard scale
Private Function GenGammaOrigCumHaz_Core(ByVal xVal As Double, _
                                         ByVal shapeVal As Double, _
                                         ByVal scaleVal As Double, _
                                         ByVal kVal As Double) As Double
    Dim survVal As Double

    If xVal <= 0# Then
        GenGammaOrigCumHaz_Core = 0#
        Exit Function
    End If

    survVal = GenGammaOrigSurv_Core(xVal, shapeVal, scaleVal, kVal)

    If survVal <= 0# Then
        GenGammaOrigCumHaz_Core = BIG_POS
    Else
        GenGammaOrigCumHaz_Core = -Log(survVal)
    End If
End Function

' GenGammaOrigQuantile_Core
' Computes Stacy generalized gamma quantile using bisection search
' Finds time corresponding to a target probability
Private Function GenGammaOrigQuantile_Core(ByVal probVal As Double, _
                                           ByVal shapeVal As Double, _
                                           ByVal scaleVal As Double, _
                                           ByVal kVal As Double, _
                                           Optional ByVal lowerTail As Boolean = True, _
                                           Optional ByVal logP As Boolean = False, _
                                           Optional ByVal tol As Double = DEFAULT_TOL, _
                                           Optional ByVal maxiter As Long = DEFAULT_MAXITER) As Double
    Dim targetVal As Double
    Dim loVal As Double, hiVal As Double, midVal As Double
    Dim fmidVal As Double
    Dim iterVal As Long

    If shapeVal <= 0# Then Err.Raise vbObjectError + 6206, , "shape must be > 0"
    If scaleVal <= 0# Then Err.Raise vbObjectError + 6207, , "scale must be > 0"
    If kVal <= 0# Then Err.Raise vbObjectError + 6208, , "k must be > 0"

    targetVal = DecodeProbFlex(probVal, lowerTail, logP)

    If targetVal < 0# Or targetVal > 1# Then Err.Raise vbObjectError + 6209, , "p must be in [0,1]"

    If targetVal <= 0# Then
        GenGammaOrigQuantile_Core = 0#
        Exit Function
    End If

    If targetVal >= 1# Then
        GenGammaOrigQuantile_Core = BIG_POS
        Exit Function
    End If

    loVal = 0#
    hiVal = MaxDbl(1#, scaleVal)

    Do While GenGammaOrigCDF_Core(hiVal, shapeVal, scaleVal, kVal) < targetVal
        hiVal = hiVal * 2#
        If hiVal > 1000000000000# Then Exit Do
    Loop

    For iterVal = 1 To maxiter
        midVal = 0.5 * (loVal + hiVal)
        fmidVal = GenGammaOrigCDF_Core(midVal, shapeVal, scaleVal, kVal) - targetVal

        If Abs(fmidVal) < tol Then
            GenGammaOrigQuantile_Core = midVal
            Exit Function
        End If

        If fmidVal >= 0# Then
            hiVal = midVal
        Else
            loVal = midVal
        End If

        If Abs(hiVal - loVal) < tol * (1# + midVal) Then
            GenGammaOrigQuantile_Core = 0.5 * (loVal + hiVal)
            Exit Function
        End If
    Next iterVal

    GenGammaOrigQuantile_Core = 0.5 * (loVal + hiVal)
End Function

'==========================================================
' genf_orig
' Parameters: muVal, sigmaVal, s1Val, s2Val
'==========================================================

' GenFOrigY_Core
' Computes transformed y term for original generalized F
' Used as an intermediate value in genf_orig formulas
Private Function GenFOrigY_Core(ByVal xVal As Double, _
                                ByVal muVal As Double, _
                                ByVal sigmaVal As Double) As Double
    If xVal <= 0# Then
        GenFOrigY_Core = 0#
    Else
        GenFOrigY_Core = SafeExp((Log(xVal) - muVal) / sigmaVal)
    End If
End Function

' GenFOrigZ_Core
' Computes transformed z term for original generalized F
' Maps x to the incomplete beta probability scale
Private Function GenFOrigZ_Core(ByVal xVal As Double, _
                                ByVal muVal As Double, _
                                ByVal sigmaVal As Double, _
                                ByVal s1Val As Double, _
                                ByVal s2Val As Double) As Double
    Dim yVal As Double

    If xVal <= 0# Then
        GenFOrigZ_Core = 0#
        Exit Function
    End If

    yVal = GenFOrigY_Core(xVal, muVal, sigmaVal)
    GenFOrigZ_Core = (s1Val * yVal) / (s2Val + s1Val * yVal)
End Function

' GenFOrigCDF_Core
' Computes original generalized F cumulative distribution function
' Uses regularized incomplete beta for probability evaluation
Private Function GenFOrigCDF_Core(ByVal xVal As Double, _
                                  ByVal muVal As Double, _
                                  ByVal sigmaVal As Double, _
                                  ByVal s1Val As Double, _
                                  ByVal s2Val As Double) As Double
    Dim zVal As Double

    If sigmaVal <= 0# Then Err.Raise vbObjectError + 6300, , "sigma must be > 0"
    If s1Val <= 0# Then Err.Raise vbObjectError + 6301, , "s1 must be > 0"
    If s2Val <= 0# Then Err.Raise vbObjectError + 6302, , "s2 must be > 0"

    If xVal <= 0# Then
        GenFOrigCDF_Core = 0#
        Exit Function
    End If

    zVal = GenFOrigZ_Core(xVal, muVal, sigmaVal, s1Val, s2Val)
    GenFOrigCDF_Core = Clamp01(CDbl(BetaI(zVal, s1Val, s2Val)))
End Function

' GenFOrigSurv_Core
' Computes original generalized F survival function
' Converts CDF to survival scale
Private Function GenFOrigSurv_Core(ByVal xVal As Double, _
                                   ByVal muVal As Double, _
                                   ByVal sigmaVal As Double, _
                                   ByVal s1Val As Double, _
                                   ByVal s2Val As Double) As Double
    If xVal <= 0# Then
        GenFOrigSurv_Core = 1#
    Else
        GenFOrigSurv_Core = 1# - GenFOrigCDF_Core(xVal, muVal, sigmaVal, s1Val, s2Val)
    End If
End Function

' GenFOrigDensity_Core
' Computes original generalized F density
' Uses log-scale formulation for numerical stability
Private Function GenFOrigDensity_Core(ByVal xVal As Double, _
                                      ByVal muVal As Double, _
                                      ByVal sigmaVal As Double, _
                                      ByVal s1Val As Double, _
                                      ByVal s2Val As Double) As Double
    Dim wVal As Double
    Dim logfVal As Double

    If sigmaVal <= 0# Then Err.Raise vbObjectError + 6303, , "sigma must be > 0"
    If s1Val <= 0# Then Err.Raise vbObjectError + 6304, , "s1 must be > 0"
    If s2Val <= 0# Then Err.Raise vbObjectError + 6305, , "s2 must be > 0"

    If xVal <= 0# Then
        GenFOrigDensity_Core = 0#
        Exit Function
    End If

    wVal = (Log(xVal) - muVal) / sigmaVal

    logfVal = s1Val * Log(s1Val / s2Val) + s1Val * wVal _
            - Log(sigmaVal) - Log(xVal) _
            - (s1Val + s2Val) * Log(1# + (s1Val / s2Val) * SafeExp(wVal)) _
            - LogBetaVBA(s1Val, s2Val)

    GenFOrigDensity_Core = SafeExp(logfVal)
End Function

' GenFOrigHaz_Core
' Computes original generalized F hazard
' Combines density and survival with protection against zero survival
Private Function GenFOrigHaz_Core(ByVal xVal As Double, _
                                  ByVal muVal As Double, _
                                  ByVal sigmaVal As Double, _
                                  ByVal s1Val As Double, _
                                  ByVal s2Val As Double) As Double
    Dim survVal As Double

    If xVal <= 0# Then
        GenFOrigHaz_Core = 0#
        Exit Function
    End If

    survVal = GenFOrigSurv_Core(xVal, muVal, sigmaVal, s1Val, s2Val)

    If survVal <= 0# Then
        GenFOrigHaz_Core = BIG_POS
    Else
        GenFOrigHaz_Core = GenFOrigDensity_Core(xVal, muVal, sigmaVal, s1Val, s2Val) / survVal
    End If
End Function


' GenFOrigCumHaz_Core
' Computes original generalized F cumulative hazard
' Converts survival to cumulative hazard scale
Private Function GenFOrigCumHaz_Core(ByVal xVal As Double, _
                                     ByVal muVal As Double, _
                                     ByVal sigmaVal As Double, _
                                     ByVal s1Val As Double, _
                                     ByVal s2Val As Double) As Double
    Dim survVal As Double

    If xVal <= 0# Then
        GenFOrigCumHaz_Core = 0#
        Exit Function
    End If

    survVal = GenFOrigSurv_Core(xVal, muVal, sigmaVal, s1Val, s2Val)

    If survVal <= 0# Then
        GenFOrigCumHaz_Core = BIG_POS
    Else
        GenFOrigCumHaz_Core = -Log(survVal)
    End If
End Function

' GenFOrigQuantile_Core
' Computes original generalized F quantile using bisection search
' Finds time corresponding to a target probability
Private Function GenFOrigQuantile_Core(ByVal probVal As Double, _
                                       ByVal muVal As Double, _
                                       ByVal sigmaVal As Double, _
                                       ByVal s1Val As Double, _
                                       ByVal s2Val As Double, _
                                       Optional ByVal lowerTail As Boolean = True, _
                                       Optional ByVal logP As Boolean = False, _
                                       Optional ByVal tol As Double = DEFAULT_TOL, _
                                       Optional ByVal maxiter As Long = DEFAULT_MAXITER) As Double
    Dim targetVal As Double
    Dim loVal As Double, hiVal As Double, midVal As Double
    Dim fmidVal As Double
    Dim iterVal As Long

    If sigmaVal <= 0# Then Err.Raise vbObjectError + 6306, , "sigma must be > 0"
    If s1Val <= 0# Then Err.Raise vbObjectError + 6307, , "s1 must be > 0"
    If s2Val <= 0# Then Err.Raise vbObjectError + 6308, , "s2 must be > 0"

    targetVal = DecodeProbFlex(probVal, lowerTail, logP)

    If targetVal < 0# Or targetVal > 1# Then Err.Raise vbObjectError + 6309, , "p must be in [0,1]"

    If targetVal <= 0# Then
        GenFOrigQuantile_Core = 0#
        Exit Function
    End If

    If targetVal >= 1# Then
        GenFOrigQuantile_Core = BIG_POS
        Exit Function
    End If

    loVal = 0#
    hiVal = MaxDbl(1#, SafeExp(muVal))

    Do While GenFOrigCDF_Core(hiVal, muVal, sigmaVal, s1Val, s2Val) < targetVal
        hiVal = hiVal * 2#
        If hiVal > 1000000000000# Then Exit Do
    Loop

    For iterVal = 1 To maxiter
        midVal = 0.5 * (loVal + hiVal)
        fmidVal = GenFOrigCDF_Core(midVal, muVal, sigmaVal, s1Val, s2Val) - targetVal

        If Abs(fmidVal) < tol Then
            GenFOrigQuantile_Core = midVal
            Exit Function
        End If

        If fmidVal >= 0# Then
            hiVal = midVal
        Else
            loVal = midVal
        End If

        If Abs(hiVal - loVal) < tol * (1# + midVal) Then
            GenFOrigQuantile_Core = 0.5 * (loVal + hiVal)
            Exit Function
        End If
    Next iterVal

    GenFOrigQuantile_Core = 0.5 * (loVal + hiVal)
End Function

'==========================================================
' genf  (Prentice parameterization)
' Parameters: muVal, sigmaVal, qParam, pParam
'==========================================================

' GenF_Params
' Converts Prentice generalized F parameters into original F parameters
' Prepares delta, s1, and s2 for downstream genf calculations
Private Sub GenF_Params(ByVal qParam As Double, _
                        ByVal pParam As Double, _
                        ByRef deltaVal As Double, _
                        ByRef s1Val As Double, _
                        ByRef s2Val As Double)
    Dim d2Val As Double

    If pParam < 0# Then Err.Raise vbObjectError + 6400, , "P must be >= 0"

    d2Val = qParam * qParam + 2# * pParam
    If d2Val <= 0# Then Err.Raise vbObjectError + 6401, , "Q^2 + 2*P must be > 0"

    deltaVal = Sqr(d2Val)

    If pParam <= FLEX_GF_P_EPS Then
        s1Val = 0#
        s2Val = 0#
    Else
        s1Val = 2# / (qParam * qParam + 2# * pParam + qParam * deltaVal)
        s2Val = 2# / (qParam * qParam + 2# * pParam - qParam * deltaVal)
    End If
End Sub


' GenFCDF_Core
' Computes Prentice generalized F cumulative distribution function
' Falls back to generalized gamma when p is near zero
Private Function GenFCDF_Core(ByVal xVal As Double, _
                              ByVal muVal As Double, _
                              ByVal sigmaVal As Double, _
                              ByVal qParam As Double, _
                              ByVal pParam As Double) As Double
    Dim deltaVal As Double, s1Val As Double, s2Val As Double

    If sigmaVal <= 0# Then Err.Raise vbObjectError + 6402, , "sigma must be > 0"

    If pParam <= FLEX_GF_P_EPS Then
        GenFCDF_Core = GenGammaCDF_Core(xVal, muVal, sigmaVal, qParam)
        Exit Function
    End If

    Call GenF_Params(qParam, pParam, deltaVal, s1Val, s2Val)
    GenFCDF_Core = GenFOrigCDF_Core(xVal, muVal, sigmaVal / deltaVal, s1Val, s2Val)
End Function


' GenFSurv_Core
' Computes Prentice generalized F survival function
' Converts CDF to survival scal
Private Function GenFSurv_Core(ByVal xVal As Double, _
                               ByVal muVal As Double, _
                               ByVal sigmaVal As Double, _
                               ByVal qParam As Double, _
                               ByVal pParam As Double) As Double
    If xVal <= 0# Then
        GenFSurv_Core = 1#
    Else
        GenFSurv_Core = 1# - GenFCDF_Core(xVal, muVal, sigmaVal, qParam, pParam)
    End If
End Function

' GenFDensity_Core
' Computes Prentice generalized F density
' Falls back to generalized gamma when p is near zero
Private Function GenFDensity_Core(ByVal xVal As Double, _
                                  ByVal muVal As Double, _
                                  ByVal sigmaVal As Double, _
                                  ByVal qParam As Double, _
                                  ByVal pParam As Double) As Double
    Dim deltaVal As Double, s1Val As Double, s2Val As Double

    If sigmaVal <= 0# Then Err.Raise vbObjectError + 6403, , "sigma must be > 0"

    If pParam <= FLEX_GF_P_EPS Then
        GenFDensity_Core = GenGammaDensity_Core(xVal, muVal, sigmaVal, qParam)
        Exit Function
    End If

    Call GenF_Params(qParam, pParam, deltaVal, s1Val, s2Val)
    GenFDensity_Core = GenFOrigDensity_Core(xVal, muVal, sigmaVal / deltaVal, s1Val, s2Val)
End Function

' GenFHaz_Core
' Computes Prentice generalized F hazard
' Combines density and survival with protection against zero survival
Private Function GenFHaz_Core(ByVal xVal As Double, _
                              ByVal muVal As Double, _
                              ByVal sigmaVal As Double, _
                              ByVal qParam As Double, _
                              ByVal pParam As Double) As Double
    Dim survVal As Double

    If xVal <= 0# Then
        GenFHaz_Core = 0#
        Exit Function
    End If

    survVal = GenFSurv_Core(xVal, muVal, sigmaVal, qParam, pParam)

    If survVal <= 0# Then
        GenFHaz_Core = BIG_POS
    Else
        GenFHaz_Core = GenFDensity_Core(xVal, muVal, sigmaVal, qParam, pParam) / survVal
    End If
End Function

' GenFCumHaz_Core
' Computes Prentice generalized F cumulative hazard
' Converts survival to cumulative hazard scale
Private Function GenFCumHaz_Core(ByVal xVal As Double, _
                                 ByVal muVal As Double, _
                                 ByVal sigmaVal As Double, _
                                 ByVal qParam As Double, _
                                 ByVal pParam As Double) As Double
    Dim survVal As Double

    If xVal <= 0# Then
        GenFCumHaz_Core = 0#
        Exit Function
    End If

    survVal = GenFSurv_Core(xVal, muVal, sigmaVal, qParam, pParam)

    If survVal <= 0# Then
        GenFCumHaz_Core = BIG_POS
    Else
        GenFCumHaz_Core = -Log(survVal)
    End If
End Function

' GenFQuantile_Core
' Computes Prentice generalized F quantile using bisection search
' Falls back to generalized gamma when p is near zero
Private Function GenFQuantile_Core(ByVal probVal As Double, _
                                   ByVal muVal As Double, _
                                   ByVal sigmaVal As Double, _
                                   ByVal qParam As Double, _
                                   ByVal pParam As Double, _
                                   Optional ByVal lowerTail As Boolean = True, _
                                   Optional ByVal logP As Boolean = False, _
                                   Optional ByVal tol As Double = DEFAULT_TOL, _
                                   Optional ByVal maxiter As Long = DEFAULT_MAXITER) As Double
    Dim targetVal As Double
    Dim loVal As Double, hiVal As Double, midVal As Double
    Dim fmidVal As Double
    Dim iterVal As Long

    If sigmaVal <= 0# Then Err.Raise vbObjectError + 6404, , "sigma must be > 0"

    If pParam <= FLEX_GF_P_EPS Then
        GenFQuantile_Core = GenGammaQuantile_Core(probVal, muVal, sigmaVal, qParam, lowerTail, logP, tol, maxiter)
        Exit Function
    End If

    targetVal = DecodeProbFlex(probVal, lowerTail, logP)

    If targetVal < 0# Or targetVal > 1# Then Err.Raise vbObjectError + 6405, , "p must be in [0,1]"

    If targetVal <= 0# Then
        GenFQuantile_Core = 0#
        Exit Function
    End If

    If targetVal >= 1# Then
        GenFQuantile_Core = BIG_POS
        Exit Function
    End If

    loVal = 0#
    hiVal = MaxDbl(1#, SafeExp(muVal))

    Do While GenFCDF_Core(hiVal, muVal, sigmaVal, qParam, pParam) < targetVal
        hiVal = hiVal * 2#
        If hiVal > 1000000000000# Then Exit Do
    Loop

    For iterVal = 1 To maxiter
        midVal = 0.5 * (loVal + hiVal)
        fmidVal = GenFCDF_Core(midVal, muVal, sigmaVal, qParam, pParam) - targetVal

        If Abs(fmidVal) < tol Then
            GenFQuantile_Core = midVal
            Exit Function
        End If

        If fmidVal >= 0# Then
            hiVal = midVal
        Else
            loVal = midVal
        End If

        If Abs(hiVal - loVal) < tol * (1# + midVal) Then
            GenFQuantile_Core = 0.5 * (loVal + hiVal)
            Exit Function
        End If
    Next iterVal

    GenFQuantile_Core = 0.5 * (loVal + hiVal)
End Function

'==========================================================
' Generic wrappers for 3-parameter families
'==========================================================

' Eval3ParamFlex
' Evaluates 3-parameter flexible distributions for density, CDF, quantile, hazard, cumulative hazard, or survival
' Dispatches generalized gamma family calculations by mode
Private Function Eval3ParamFlex(ByVal xVal As Double, _
                                ByVal modeKey As String, _
                                ByVal familyKey As String, _
                                ByVal par1 As Double, _
                                ByVal par2 As Double, _
                                ByVal par3 As Double, _
                                ByVal lowerTail As Boolean, _
                                ByVal logP As Boolean, _
                                ByVal logFlag As Boolean) As Variant
    Dim densVal As Double, cdfVal As Double, survVal As Double
    Dim hazVal As Double, cumHazVal As Double

    On Error GoTo EH

    Select Case LCase$(Trim$(familyKey))

        Case "gengamma"
            Select Case LCase$(Trim$(modeKey))
                Case "d"
                    densVal = GenGammaDensity_Core(xVal, par1, par2, par3)
                    Eval3ParamFlex = IIf(logFlag, SafeLog(densVal), densVal)
                Case "p"
                    cdfVal = GenGammaCDF_Core(xVal, par1, par2, par3)
                    Eval3ParamFlex = FinishProbFlex(cdfVal, lowerTail, logP)
                Case "q"
                    Eval3ParamFlex = GenGammaQuantile_Core(xVal, par1, par2, par3, lowerTail, logP)
                Case "h"
                    hazVal = GenGammaHaz_Core(xVal, par1, par2, par3)
                    Eval3ParamFlex = IIf(logFlag, SafeLog(hazVal), hazVal)
                Case "ch"
                    cumHazVal = GenGammaCumHaz_Core(xVal, par1, par2, par3)
                    Eval3ParamFlex = IIf(logFlag, SafeLog(cumHazVal), cumHazVal)
                Case "s"
                    survVal = GenGammaSurv_Core(xVal, par1, par2, par3)
                    Eval3ParamFlex = IIf(logFlag, SafeLog(survVal), survVal)
                Case Else
                    Eval3ParamFlex = CVErr(xlErrValue)
            End Select

        Case "gengamma_orig"
            Select Case LCase$(Trim$(modeKey))
                Case "d"
                    densVal = GenGammaOrigDensity_Core(xVal, par1, par2, par3)
                    Eval3ParamFlex = IIf(logFlag, SafeLog(densVal), densVal)
                Case "p"
                    cdfVal = GenGammaOrigCDF_Core(xVal, par1, par2, par3)
                    Eval3ParamFlex = FinishProbFlex(cdfVal, lowerTail, logP)
                Case "q"
                    Eval3ParamFlex = GenGammaOrigQuantile_Core(xVal, par1, par2, par3, lowerTail, logP)
                Case "h"
                    hazVal = GenGammaOrigHaz_Core(xVal, par1, par2, par3)
                    Eval3ParamFlex = IIf(logFlag, SafeLog(hazVal), hazVal)
                Case "ch"
                    cumHazVal = GenGammaOrigCumHaz_Core(xVal, par1, par2, par3)
                    Eval3ParamFlex = IIf(logFlag, SafeLog(cumHazVal), cumHazVal)
                Case "s"
                    survVal = GenGammaOrigSurv_Core(xVal, par1, par2, par3)
                    Eval3ParamFlex = IIf(logFlag, SafeLog(survVal), survVal)
                Case Else
                    Eval3ParamFlex = CVErr(xlErrValue)
            End Select

        Case Else
            Eval3ParamFlex = CVErr(xlErrValue)
    End Select

    Exit Function

EH:
    Eval3ParamFlex = CVErr(xlErrValue)
End Function

' Run3ParamFlex
' Applies 3-parameter flexible distribution evaluation to scalars, arrays, or Excel ranges
' Provides vectorized worksheet support for generalized gamma families
Private Function Run3ParamFlex(ByVal xInput As Variant, _
                               ByVal modeKey As String, _
                               ByVal familyKey As String, _
                               ByVal par1 As Double, _
                               ByVal par2 As Double, _
                               ByVal par3 As Double, _
                               Optional ByVal lowerTail As Boolean = True, _
                               Optional ByVal logP As Boolean = False, _
                               Optional ByVal logFlag As Boolean = False) As Variant
    Dim vals As Variant
    Dim out() As Variant
    Dim rVal As Long, cVal As Long
    Dim rngObj As Range

    On Error GoTo EH

    If IsObject(xInput) Then
        If TypeName(xInput) = "Range" Then
            Set rngObj = xInput

            If rngObj.CountLarge = 1 Then
                If IsNumeric(rngObj.Value2) Then
                    Run3ParamFlex = Eval3ParamFlex(CDbl(rngObj.Value2), modeKey, familyKey, par1, par2, par3, lowerTail, logP, logFlag)
                Else
                    Run3ParamFlex = CVErr(xlErrValue)
                End If
                Exit Function
            End If

            vals = rngObj.Value2
            ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))

            For rVal = 1 To UBound(vals, 1)
                For cVal = 1 To UBound(vals, 2)
                    If IsNumeric(vals(rVal, cVal)) Then
                        out(rVal, cVal) = Eval3ParamFlex(CDbl(vals(rVal, cVal)), modeKey, familyKey, par1, par2, par3, lowerTail, logP, logFlag)
                    Else
                        out(rVal, cVal) = CVErr(xlErrValue)
                    End If
                Next cVal
            Next rVal

            Run3ParamFlex = out
            Exit Function
        End If
    End If

    If IsArray(xInput) Then
        vals = xInput
        ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))

        For rVal = 1 To UBound(vals, 1)
            For cVal = 1 To UBound(vals, 2)
                If IsNumeric(vals(rVal, cVal)) Then
                    out(rVal, cVal) = Eval3ParamFlex(CDbl(vals(rVal, cVal)), modeKey, familyKey, par1, par2, par3, lowerTail, logP, logFlag)
                Else
                    out(rVal, cVal) = CVErr(xlErrValue)
                End If
            Next cVal
        Next rVal

        Run3ParamFlex = out
    ElseIf IsNumeric(xInput) Then
        Run3ParamFlex = Eval3ParamFlex(CDbl(xInput), modeKey, familyKey, par1, par2, par3, lowerTail, logP, logFlag)
    Else
        Run3ParamFlex = CVErr(xlErrValue)
    End If

    Exit Function

EH:
    Run3ParamFlex = CVErr(xlErrValue)
End Function

'==========================================================
' Generic wrappers for 4-parameter families
'==========================================================

' Eval4ParamFlex
' Evaluates 4-parameter flexible distributions for density, CDF, quantile, hazard, cumulative hazard, or survival
' Dispatches generalized F family calculations by mode
Private Function Eval4ParamFlex(ByVal xVal As Double, _
                                ByVal modeKey As String, _
                                ByVal familyKey As String, _
                                ByVal par1 As Double, _
                                ByVal par2 As Double, _
                                ByVal par3 As Double, _
                                ByVal par4 As Double, _
                                ByVal lowerTail As Boolean, _
                                ByVal logP As Boolean, _
                                ByVal logFlag As Boolean) As Variant
    Dim densVal As Double, cdfVal As Double, survVal As Double
    Dim hazVal As Double, cumHazVal As Double

    On Error GoTo EH

    Select Case LCase$(Trim$(familyKey))

        Case "genf"
            Select Case LCase$(Trim$(modeKey))
                Case "d"
                    densVal = GenFDensity_Core(xVal, par1, par2, par3, par4)
                    Eval4ParamFlex = IIf(logFlag, SafeLog(densVal), densVal)
                Case "p"
                    cdfVal = GenFCDF_Core(xVal, par1, par2, par3, par4)
                    Eval4ParamFlex = FinishProbFlex(cdfVal, lowerTail, logP)
                Case "q"
                    Eval4ParamFlex = GenFQuantile_Core(xVal, par1, par2, par3, par4, lowerTail, logP)
                Case "h"
                    hazVal = GenFHaz_Core(xVal, par1, par2, par3, par4)
                    Eval4ParamFlex = IIf(logFlag, SafeLog(hazVal), hazVal)
                Case "ch"
                    cumHazVal = GenFCumHaz_Core(xVal, par1, par2, par3, par4)
                    Eval4ParamFlex = IIf(logFlag, SafeLog(cumHazVal), cumHazVal)
                Case "s"
                    survVal = GenFSurv_Core(xVal, par1, par2, par3, par4)
                    Eval4ParamFlex = IIf(logFlag, SafeLog(survVal), survVal)
                Case Else
                    Eval4ParamFlex = CVErr(xlErrValue)
            End Select

        Case "genf_orig"
            Select Case LCase$(Trim$(modeKey))
                Case "d"
                    densVal = GenFOrigDensity_Core(xVal, par1, par2, par3, par4)
                    Eval4ParamFlex = IIf(logFlag, SafeLog(densVal), densVal)
                Case "p"
                    cdfVal = GenFOrigCDF_Core(xVal, par1, par2, par3, par4)
                    Eval4ParamFlex = FinishProbFlex(cdfVal, lowerTail, logP)
                Case "q"
                    Eval4ParamFlex = GenFOrigQuantile_Core(xVal, par1, par2, par3, par4, lowerTail, logP)
                Case "h"
                    hazVal = GenFOrigHaz_Core(xVal, par1, par2, par3, par4)
                    Eval4ParamFlex = IIf(logFlag, SafeLog(hazVal), hazVal)
                Case "ch"
                    cumHazVal = GenFOrigCumHaz_Core(xVal, par1, par2, par3, par4)
                    Eval4ParamFlex = IIf(logFlag, SafeLog(cumHazVal), cumHazVal)
                Case "s"
                    survVal = GenFOrigSurv_Core(xVal, par1, par2, par3, par4)
                    Eval4ParamFlex = IIf(logFlag, SafeLog(survVal), survVal)
                Case Else
                    Eval4ParamFlex = CVErr(xlErrValue)
            End Select

        Case Else
            Eval4ParamFlex = CVErr(xlErrValue)
    End Select

    Exit Function

EH:
    Eval4ParamFlex = CVErr(xlErrValue)
End Function

' Run4ParamFlex
' Applies 4-parameter flexible distribution evaluation to scalars, arrays, or Excel ranges
' Provides vectorized worksheet support for generalized F families
Private Function Run4ParamFlex(ByVal xInput As Variant, _
                               ByVal modeKey As String, _
                               ByVal familyKey As String, _
                               ByVal par1 As Double, _
                               ByVal par2 As Double, _
                               ByVal par3 As Double, _
                               ByVal par4 As Double, _
                               Optional ByVal lowerTail As Boolean = True, _
                               Optional ByVal logP As Boolean = False, _
                               Optional ByVal logFlag As Boolean = False) As Variant
    Dim vals As Variant
    Dim out() As Variant
    Dim rVal As Long, cVal As Long
    Dim rngObj As Range

    On Error GoTo EH

    If IsObject(xInput) Then
        If TypeName(xInput) = "Range" Then
            Set rngObj = xInput

            If rngObj.CountLarge = 1 Then
                If IsNumeric(rngObj.Value2) Then
                    Run4ParamFlex = Eval4ParamFlex(CDbl(rngObj.Value2), modeKey, familyKey, par1, par2, par3, par4, lowerTail, logP, logFlag)
                Else
                    Run4ParamFlex = CVErr(xlErrValue)
                End If
                Exit Function
            End If

            vals = rngObj.Value2
            ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))

            For rVal = 1 To UBound(vals, 1)
                For cVal = 1 To UBound(vals, 2)
                    If IsNumeric(vals(rVal, cVal)) Then
                        out(rVal, cVal) = Eval4ParamFlex(CDbl(vals(rVal, cVal)), modeKey, familyKey, par1, par2, par3, par4, lowerTail, logP, logFlag)
                    Else
                        out(rVal, cVal) = CVErr(xlErrValue)
                    End If
                Next cVal
            Next rVal

            Run4ParamFlex = out
            Exit Function
        End If
    End If

    If IsArray(xInput) Then
        vals = xInput
        ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))

        For rVal = 1 To UBound(vals, 1)
            For cVal = 1 To UBound(vals, 2)
                If IsNumeric(vals(rVal, cVal)) Then
                    out(rVal, cVal) = Eval4ParamFlex(CDbl(vals(rVal, cVal)), modeKey, familyKey, par1, par2, par3, par4, lowerTail, logP, logFlag)
                Else
                    out(rVal, cVal) = CVErr(xlErrValue)
                End If
            Next cVal
        Next rVal

        Run4ParamFlex = out
    ElseIf IsNumeric(xInput) Then
        Run4ParamFlex = Eval4ParamFlex(CDbl(xInput), modeKey, familyKey, par1, par2, par3, par4, lowerTail, logP, logFlag)
    Else
        Run4ParamFlex = CVErr(xlErrValue)
    End If

    Exit Function

EH:
    Run4ParamFlex = CVErr(xlErrValue)
End Function

'==========================================================
' Random wrappers
'==========================================================

' RandomVector3Flex
' Generates random draws for 3-parameter flexible distributions using inverse-CDF sampling
' Supports random simulation for generalized gamma families
Private Function RandomVector3Flex(ByVal nVal As Long, _
                                   ByVal familyKey As String, _
                                   ByVal par1 As Double, _
                                   ByVal par2 As Double, _
                                   ByVal par3 As Double) As Variant
    Dim out() As Variant
    Dim iVal As Long
    Dim uVal As Double

    If nVal <= 0 Then
        RandomVector3Flex = CVErr(xlErrValue)
        Exit Function
    End If

    If nVal = 1 Then
        uVal = UniformOpen01()
        Select Case LCase$(familyKey)
            Case "gengamma"
                RandomVector3Flex = GenGammaQuantile_Core(uVal, par1, par2, par3, True, False)
            Case "gengamma_orig"
                RandomVector3Flex = GenGammaOrigQuantile_Core(uVal, par1, par2, par3, True, False)
            Case Else
                RandomVector3Flex = CVErr(xlErrValue)
        End Select
        Exit Function
    End If

    ReDim out(1 To nVal, 1 To 1)
    For iVal = 1 To nVal
        uVal = UniformOpen01()
        Select Case LCase$(familyKey)
            Case "gengamma"
                out(iVal, 1) = GenGammaQuantile_Core(uVal, par1, par2, par3, True, False)
            Case "gengamma_orig"
                out(iVal, 1) = GenGammaOrigQuantile_Core(uVal, par1, par2, par3, True, False)
            Case Else
                out(iVal, 1) = CVErr(xlErrValue)
        End Select
    Next iVal

    RandomVector3Flex = out
End Function

' RandomVector4Flex
' Generates random draws for 4-parameter flexible distributions using inverse-CDF sampling
' Supports random simulation for generalized F families
Private Function RandomVector4Flex(ByVal nVal As Long, _
                                   ByVal familyKey As String, _
                                   ByVal par1 As Double, _
                                   ByVal par2 As Double, _
                                   ByVal par3 As Double, _
                                   ByVal par4 As Double) As Variant
    Dim out() As Variant
    Dim iVal As Long
    Dim uVal As Double

    If nVal <= 0 Then
        RandomVector4Flex = CVErr(xlErrValue)
        Exit Function
    End If

    If nVal = 1 Then
        uVal = UniformOpen01()
        Select Case LCase$(familyKey)
            Case "genf"
                RandomVector4Flex = GenFQuantile_Core(uVal, par1, par2, par3, par4, True, False)
            Case "genf_orig"
                RandomVector4Flex = GenFOrigQuantile_Core(uVal, par1, par2, par3, par4, True, False)
            Case Else
                RandomVector4Flex = CVErr(xlErrValue)
        End Select
        Exit Function
    End If

    ReDim out(1 To nVal, 1 To 1)
    For iVal = 1 To nVal
        uVal = UniformOpen01()
        Select Case LCase$(familyKey)
            Case "genf"
                out(iVal, 1) = GenFQuantile_Core(uVal, par1, par2, par3, par4, True, False)
            Case "genf_orig"
                out(iVal, 1) = GenFOrigQuantile_Core(uVal, par1, par2, par3, par4, True, False)
            Case Else
                out(iVal, 1) = CVErr(xlErrValue)
        End Select
    Next iVal

    RandomVector4Flex = out
End Function

'==========================================================
' Public functions: gengamma
'==========================================================

Public Function dgengamma(ByVal xInput As Variant, _
                          ByVal muVal As Double, _
                          ByVal sigmaVal As Double, _
                          ByVal qParam As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    dgengamma = Run3ParamFlex(xInput, "d", "gengamma", muVal, sigmaVal, qParam, True, False, logFlag)
End Function

Public Function pgengamma(ByVal xInput As Variant, _
                          ByVal muVal As Double, _
                          ByVal sigmaVal As Double, _
                          ByVal qParam As Double, _
                          Optional ByVal lowerTail As Boolean = True, _
                          Optional ByVal logP As Boolean = False) As Variant
    pgengamma = Run3ParamFlex(xInput, "p", "gengamma", muVal, sigmaVal, qParam, lowerTail, logP, False)
End Function

Public Function qgengamma(ByVal probInput As Variant, _
                          ByVal muVal As Double, _
                          ByVal sigmaVal As Double, _
                          ByVal qParam As Double, _
                          Optional ByVal lowerTail As Boolean = True, _
                          Optional ByVal logP As Boolean = False) As Variant
    qgengamma = Run3ParamFlex(probInput, "q", "gengamma", muVal, sigmaVal, qParam, lowerTail, logP, False)
End Function

Public Function hgengamma(ByVal xInput As Variant, _
                          ByVal muVal As Double, _
                          ByVal sigmaVal As Double, _
                          ByVal qParam As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    hgengamma = Run3ParamFlex(xInput, "h", "gengamma", muVal, sigmaVal, qParam, True, False, logFlag)
End Function

Public Function chgengamma(ByVal xInput As Variant, _
                           ByVal muVal As Double, _
                           ByVal sigmaVal As Double, _
                           ByVal qParam As Double, _
                           Optional ByVal logFlag As Boolean = False) As Variant
    chgengamma = Run3ParamFlex(xInput, "ch", "gengamma", muVal, sigmaVal, qParam, True, False, logFlag)
End Function

Public Function Sgengamma(ByVal xInput As Variant, _
                          ByVal muVal As Double, _
                          ByVal sigmaVal As Double, _
                          ByVal qParam As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    Sgengamma = Run3ParamFlex(xInput, "s", "gengamma", muVal, sigmaVal, qParam, True, False, logFlag)
End Function

Public Function rgengamma(Optional ByVal nVal As Long = 1, _
                          Optional ByVal muVal As Double = 0#, _
                          Optional ByVal sigmaVal As Double = 1#, _
                          Optional ByVal qParam As Double = 1#) As Variant
    rgengamma = RandomVector3Flex(nVal, "gengamma", muVal, sigmaVal, qParam)
End Function

'==========================================================
' Public functions: gengamma_orig
'==========================================================

Public Function dgengamma_orig(ByVal xInput As Variant, _
                               ByVal shapeVal As Double, _
                               ByVal scaleVal As Double, _
                               ByVal kVal As Double, _
                               Optional ByVal logFlag As Boolean = False) As Variant
    dgengamma_orig = Run3ParamFlex(xInput, "d", "gengamma_orig", shapeVal, scaleVal, kVal, True, False, logFlag)
End Function

Public Function pgengamma_orig(ByVal xInput As Variant, _
                               ByVal shapeVal As Double, _
                               ByVal scaleVal As Double, _
                               ByVal kVal As Double, _
                               Optional ByVal lowerTail As Boolean = True, _
                               Optional ByVal logP As Boolean = False) As Variant
    pgengamma_orig = Run3ParamFlex(xInput, "p", "gengamma_orig", shapeVal, scaleVal, kVal, lowerTail, logP, False)
End Function

Public Function qgengamma_orig(ByVal probInput As Variant, _
                               ByVal shapeVal As Double, _
                               ByVal scaleVal As Double, _
                               ByVal kVal As Double, _
                               Optional ByVal lowerTail As Boolean = True, _
                               Optional ByVal logP As Boolean = False) As Variant
    qgengamma_orig = Run3ParamFlex(probInput, "q", "gengamma_orig", shapeVal, scaleVal, kVal, lowerTail, logP, False)
End Function

Public Function hgengamma_orig(ByVal xInput As Variant, _
                               ByVal shapeVal As Double, _
                               ByVal scaleVal As Double, _
                               ByVal kVal As Double, _
                               Optional ByVal logFlag As Boolean = False) As Variant
    hgengamma_orig = Run3ParamFlex(xInput, "h", "gengamma_orig", shapeVal, scaleVal, kVal, True, False, logFlag)
End Function

Public Function chgengamma_orig(ByVal xInput As Variant, _
                                ByVal shapeVal As Double, _
                                ByVal scaleVal As Double, _
                                ByVal kVal As Double, _
                                Optional ByVal logFlag As Boolean = False) As Variant
    chgengamma_orig = Run3ParamFlex(xInput, "ch", "gengamma_orig", shapeVal, scaleVal, kVal, True, False, logFlag)
End Function

Public Function Sgengamma_orig(ByVal xInput As Variant, _
                               ByVal shapeVal As Double, _
                               ByVal scaleVal As Double, _
                               ByVal kVal As Double, _
                               Optional ByVal logFlag As Boolean = False) As Variant
    Sgengamma_orig = Run3ParamFlex(xInput, "s", "gengamma_orig", shapeVal, scaleVal, kVal, True, False, logFlag)
End Function

Public Function rgengamma_orig(Optional ByVal nVal As Long = 1, _
                               Optional ByVal shapeVal As Double = 1#, _
                               Optional ByVal scaleVal As Double = 1#, _
                               Optional ByVal kVal As Double = 1#) As Variant
    rgengamma_orig = RandomVector3Flex(nVal, "gengamma_orig", shapeVal, scaleVal, kVal)
End Function

'==========================================================
' Public functions: genf
'==========================================================

Public Function dgenf(ByVal xInput As Variant, _
                      ByVal muVal As Double, _
                      ByVal sigmaVal As Double, _
                      ByVal qParam As Double, _
                      ByVal pParam As Double, _
                      Optional ByVal logFlag As Boolean = False) As Variant
    dgenf = Run4ParamFlex(xInput, "d", "genf", muVal, sigmaVal, qParam, pParam, True, False, logFlag)
End Function

Public Function pgenf(ByVal xInput As Variant, _
                      ByVal muVal As Double, _
                      ByVal sigmaVal As Double, _
                      ByVal qParam As Double, _
                      ByVal pParam As Double, _
                      Optional ByVal lowerTail As Boolean = True, _
                      Optional ByVal logP As Boolean = False) As Variant
    pgenf = Run4ParamFlex(xInput, "p", "genf", muVal, sigmaVal, qParam, pParam, lowerTail, logP, False)
End Function

Public Function qgenf(ByVal probInput As Variant, _
                      ByVal muVal As Double, _
                      ByVal sigmaVal As Double, _
                      ByVal qParam As Double, _
                      ByVal pParam As Double, _
                      Optional ByVal lowerTail As Boolean = True, _
                      Optional ByVal logP As Boolean = False) As Variant
    qgenf = Run4ParamFlex(probInput, "q", "genf", muVal, sigmaVal, qParam, pParam, lowerTail, logP, False)
End Function

Public Function hgenf(ByVal xInput As Variant, _
                      ByVal muVal As Double, _
                      ByVal sigmaVal As Double, _
                      ByVal qParam As Double, _
                      ByVal pParam As Double, _
                      Optional ByVal logFlag As Boolean = False) As Variant
    hgenf = Run4ParamFlex(xInput, "h", "genf", muVal, sigmaVal, qParam, pParam, True, False, logFlag)
End Function

Public Function chgenf(ByVal xInput As Variant, _
                       ByVal muVal As Double, _
                       ByVal sigmaVal As Double, _
                       ByVal qParam As Double, _
                       ByVal pParam As Double, _
                       Optional ByVal logFlag As Boolean = False) As Variant
    chgenf = Run4ParamFlex(xInput, "ch", "genf", muVal, sigmaVal, qParam, pParam, True, False, logFlag)
End Function

Public Function Sgenf(ByVal xInput As Variant, _
                      ByVal muVal As Double, _
                      ByVal sigmaVal As Double, _
                      ByVal qParam As Double, _
                      ByVal pParam As Double, _
                      Optional ByVal logFlag As Boolean = False) As Variant
    Sgenf = Run4ParamFlex(xInput, "s", "genf", muVal, sigmaVal, qParam, pParam, True, False, logFlag)
End Function

Public Function rgenf(Optional ByVal nVal As Long = 1, _
                      Optional ByVal muVal As Double = 0#, _
                      Optional ByVal sigmaVal As Double = 1#, _
                      Optional ByVal qParam As Double = 0#, _
                      Optional ByVal pParam As Double = 1#) As Variant
    rgenf = RandomVector4Flex(nVal, "genf", muVal, sigmaVal, qParam, pParam)
End Function

'==========================================================
' Public functions: genf_orig
'==========================================================

Public Function dgenf_orig(ByVal xInput As Variant, _
                           ByVal muVal As Double, _
                           ByVal sigmaVal As Double, _
                           ByVal s1Val As Double, _
                           ByVal s2Val As Double, _
                           Optional ByVal logFlag As Boolean = False) As Variant
    dgenf_orig = Run4ParamFlex(xInput, "d", "genf_orig", muVal, sigmaVal, s1Val, s2Val, True, False, logFlag)
End Function

Public Function pgenf_orig(ByVal xInput As Variant, _
                           ByVal muVal As Double, _
                           ByVal sigmaVal As Double, _
                           ByVal s1Val As Double, _
                           ByVal s2Val As Double, _
                           Optional ByVal lowerTail As Boolean = True, _
                           Optional ByVal logP As Boolean = False) As Variant
    pgenf_orig = Run4ParamFlex(xInput, "p", "genf_orig", muVal, sigmaVal, s1Val, s2Val, lowerTail, logP, False)
End Function

Public Function qgenf_orig(ByVal probInput As Variant, _
                           ByVal muVal As Double, _
                           ByVal sigmaVal As Double, _
                           ByVal s1Val As Double, _
                           ByVal s2Val As Double, _
                           Optional ByVal lowerTail As Boolean = True, _
                           Optional ByVal logP As Boolean = False) As Variant
    qgenf_orig = Run4ParamFlex(probInput, "q", "genf_orig", muVal, sigmaVal, s1Val, s2Val, lowerTail, logP, False)
End Function

Public Function hgenf_orig(ByVal xInput As Variant, _
                           ByVal muVal As Double, _
                           ByVal sigmaVal As Double, _
                           ByVal s1Val As Double, _
                           ByVal s2Val As Double, _
                           Optional ByVal logFlag As Boolean = False) As Variant
    hgenf_orig = Run4ParamFlex(xInput, "h", "genf_orig", muVal, sigmaVal, s1Val, s2Val, True, False, logFlag)
End Function

Public Function chgenf_orig(ByVal xInput As Variant, _
                            ByVal muVal As Double, _
                            ByVal sigmaVal As Double, _
                            ByVal s1Val As Double, _
                            ByVal s2Val As Double, _
                            Optional ByVal logFlag As Boolean = False) As Variant
    chgenf_orig = Run4ParamFlex(xInput, "ch", "genf_orig", muVal, sigmaVal, s1Val, s2Val, True, False, logFlag)
End Function

Public Function Sgenf_orig(ByVal xInput As Variant, _
                           ByVal muVal As Double, _
                           ByVal sigmaVal As Double, _
                           ByVal s1Val As Double, _
                           ByVal s2Val As Double, _
                           Optional ByVal logFlag As Boolean = False) As Variant
    Sgenf_orig = Run4ParamFlex(xInput, "s", "genf_orig", muVal, sigmaVal, s1Val, s2Val, True, False, logFlag)
End Function

Public Function rgenf_orig(Optional ByVal nVal As Long = 1, _
                           Optional ByVal muVal As Double = 0#, _
                           Optional ByVal sigmaVal As Double = 1#, _
                           Optional ByVal s1Val As Double = 1#, _
                           Optional ByVal s2Val As Double = 1#) As Variant
    rgenf_orig = RandomVector4Flex(nVal, "genf_orig", muVal, sigmaVal, s1Val, s2Val)
End Function

''' MIXTURE CODE BELOW

'==========================================================
' MIXTURE CURE RISK FUNCTIONS FOR FLEXSURVCURE
'
' This block is intended to live inside the same RiskFunctions
' module as the existing flexsurv risk functions above.
'
' Assumption in this version:
'   thetaInput is already the cure fraction
'   and does NOT require logistic / probit / cloglog /
'   identity transformation.
'
' Cure-fraction model:
'   S_mix(t)  = theta + (1-theta) * S0(t)
'   F_mix(t)  = (1-theta) * F0(t)
'   f_mix(t)  = (1-theta) * f0(t)
'   h_mix(t)  = ((1-theta) * f0(t)) / S_mix(t)
'   H_mix(t)  = -log(S_mix(t))
'
' Public worksheet functions follow naming like:
'   Smixweibull, hmixweibull, chmixweibull, pmixweibull,
'   dmixweibull, qmixweibull
'
'==========================================================

'==========================================================
' Mixture helpers
'==========================================================

Private Function ClampThetaMix(ByVal thetaVal As Double) As Double
    If thetaVal < 0# Then
        ClampThetaMix = 0#
    ElseIf thetaVal >= 1# Then
        ClampThetaMix = 1# - 0.000000000001
    Else
        ClampThetaMix = thetaVal
    End If
End Function

Private Function SmixCore(ByVal surv0 As Double, ByVal thetaVal As Double) As Double
    thetaVal = ClampThetaMix(thetaVal)
    SmixCore = thetaVal + (1# - thetaVal) * surv0
End Function

Private Function pmixCore(ByVal cdf0 As Double, ByVal thetaVal As Double) As Double
    thetaVal = ClampThetaMix(thetaVal)
    pmixCore = (1# - thetaVal) * cdf0
End Function

Private Function dmixCore(ByVal dens0 As Double, ByVal thetaVal As Double) As Double
    thetaVal = ClampThetaMix(thetaVal)
    dmixCore = (1# - thetaVal) * dens0
End Function

Private Function hmixCore(ByVal dens0 As Double, ByVal surv0 As Double, ByVal thetaVal As Double) As Double
    Dim denomVal As Double
    
    thetaVal = ClampThetaMix(thetaVal)
    denomVal = thetaVal + (1# - thetaVal) * surv0
    
    If denomVal <= 0# Then
        hmixCore = BIG_POS
    Else
        hmixCore = ((1# - thetaVal) * dens0) / denomVal
    End If
End Function

Private Function chmixCore(ByVal surv0 As Double, ByVal thetaVal As Double) As Double
    Dim sMixVal As Double
    
    sMixVal = SmixCore(surv0, thetaVal)
    
    If sMixVal <= 0# Then
        chmixCore = BIG_POS
    Else
        chmixCore = -Log(sMixVal)
    End If
End Function

Private Function qmixProbCore(ByVal pVal As Double, ByVal thetaVal As Double) As Double
    thetaVal = ClampThetaMix(thetaVal)
    
    If pVal < 0# Or pVal > 1# Then
        qmixProbCore = BIG_POS
        Exit Function
    End If
    
    If pVal <= 0# Then
        qmixProbCore = 0#
    ElseIf pVal >= 1# - thetaVal Then
        qmixProbCore = BIG_POS
    Else
        qmixProbCore = pVal / (1# - thetaVal)
    End If
End Function

'==========================================================
' 1-2 parameter mixture evaluator
'==========================================================

' EvalMixOne
' Evaluates one mixture-cure value for density, CDF, quantile, hazard, cumulative hazard, or survival
' Wraps standard distributions with cure-fraction logic
Private Function EvalMixOne(ByVal xVal As Double, _
                            ByVal distCode As Long, _
                            ByVal modeKey As String, _
                            ByVal thetaVal As Double, _
                            ByVal p1 As Double, _
                            ByVal p2 As Double, _
                            ByVal lowerTail As Boolean, _
                            ByVal logP As Boolean, _
                            ByVal logFlag As Boolean) As Variant
    Dim dens0 As Double, cdf0 As Double, surv0 As Double
    Dim outVal As Double, pWork As Double, pAdj As Double
    
    On Error GoTo EH
    
    thetaVal = ClampThetaMix(thetaVal)
    
    Select Case LCase$(modeKey)
        Case "d"
            dens0 = CDbl(EvalOne(xVal, distCode, "d", p1, p2, True, False, False))
            outVal = dmixCore(dens0, thetaVal)
            EvalMixOne = IIf(logFlag, SafeLog(outVal), outVal)
        
        Case "p"
            cdf0 = CDbl(EvalOne(xVal, distCode, "p", p1, p2, True, False, False))
            outVal = pmixCore(cdf0, thetaVal)
            EvalMixOne = FinishProb(outVal, lowerTail, logP)
        
        Case "q"
            pWork = DecodeProb(xVal, lowerTail, logP)
            If pWork < 0# Or pWork > 1# Then Err.Raise vbObjectError + 7000, , "p must be in [0,1]"
            
            pAdj = qmixProbCore(pWork, thetaVal)
            If pAdj >= BIG_POS Then
                EvalMixOne = BIG_POS
            ElseIf pAdj <= 0# Then
                EvalMixOne = 0#
            Else
                EvalMixOne = EvalOne(pAdj, distCode, "q", p1, p2, True, False, False)
            End If
        
        Case "h"
            dens0 = CDbl(EvalOne(xVal, distCode, "d", p1, p2, True, False, False))
            surv0 = CDbl(EvalOne(xVal, distCode, "s", p1, p2, True, False, False))
            outVal = hmixCore(dens0, surv0, thetaVal)
            EvalMixOne = IIf(logFlag, SafeLog(outVal), outVal)
        
        Case "ch"
            surv0 = CDbl(EvalOne(xVal, distCode, "s", p1, p2, True, False, False))
            outVal = chmixCore(surv0, thetaVal)
            EvalMixOne = IIf(logFlag, SafeLog(outVal), outVal)
        
        Case "s"
            surv0 = CDbl(EvalOne(xVal, distCode, "s", p1, p2, True, False, False))
            outVal = SmixCore(surv0, thetaVal)
            EvalMixOne = IIf(logFlag, SafeLog(outVal), outVal)
        
        Case Else
            EvalMixOne = CVErr(xlErrValue)
    End Select
    
    Exit Function
    
EH:
    EvalMixOne = CVErr(xlErrValue)
End Function

' MixScalarOrArray
' Applies mixture-cure evaluation to scalars, arrays, or Excel ranges
' Provides vectorized worksheet support for mixture-cure functions
Private Function MixScalarOrArray(ByVal xInput As Variant, _
                                  ByVal distCode As Long, _
                                  ByVal modeKey As String, _
                                  ByVal thetaVal As Double, _
                                  ByVal p1 As Double, _
                                  Optional ByVal p2 As Double = 0#, _
                                  Optional ByVal lowerTail As Boolean = True, _
                                  Optional ByVal logP As Boolean = False, _
                                  Optional ByVal logFlag As Boolean = False) As Variant
    Dim vals As Variant
    Dim outArr() As Variant
    Dim rVal As Long, cVal As Long
    Dim rngObj As Range
    
    On Error GoTo EH
    
    If IsObject(xInput) Then
        If TypeName(xInput) = "Range" Then
            Set rngObj = xInput
            
            If rngObj.CountLarge = 1 Then
                If IsNumeric(rngObj.Value2) Then
                    MixScalarOrArray = EvalMixOne(CDbl(rngObj.Value2), distCode, modeKey, thetaVal, p1, p2, lowerTail, logP, logFlag)
                Else
                    MixScalarOrArray = CVErr(xlErrValue)
                End If
                Exit Function
            End If
            
            vals = rngObj.Value2
            ReDim outArr(1 To UBound(vals, 1), 1 To UBound(vals, 2))
            
            For rVal = 1 To UBound(vals, 1)
                For cVal = 1 To UBound(vals, 2)
                    If IsNumeric(vals(rVal, cVal)) Then
                        outArr(rVal, cVal) = EvalMixOne(CDbl(vals(rVal, cVal)), distCode, modeKey, thetaVal, p1, p2, lowerTail, logP, logFlag)
                    Else
                        outArr(rVal, cVal) = CVErr(xlErrValue)
                    End If
                Next cVal
            Next rVal
            
            MixScalarOrArray = outArr
            Exit Function
        End If
    End If
    
    If IsArray(xInput) Then
        vals = xInput
        ReDim outArr(1 To UBound(vals, 1), 1 To UBound(vals, 2))
        
        For rVal = 1 To UBound(vals, 1)
            For cVal = 1 To UBound(vals, 2)
                If IsNumeric(vals(rVal, cVal)) Then
                    outArr(rVal, cVal) = EvalMixOne(CDbl(vals(rVal, cVal)), distCode, modeKey, thetaVal, p1, p2, lowerTail, logP, logFlag)
                Else
                    outArr(rVal, cVal) = CVErr(xlErrValue)
                End If
            Next cVal
        Next rVal
        
        MixScalarOrArray = outArr
    ElseIf IsNumeric(xInput) Then
        MixScalarOrArray = EvalMixOne(CDbl(xInput), distCode, modeKey, thetaVal, p1, p2, lowerTail, logP, logFlag)
    Else
        MixScalarOrArray = CVErr(xlErrValue)
    End If
    
    Exit Function
    
EH:
    MixScalarOrArray = CVErr(xlErrValue)
End Function

'==========================================================
' 3-parameter flex mixture evaluator
'==========================================================

' Eval3ParamFlexMix
' Evaluates 3-parameter flexible mixture-cure distributions by mode
' Dispatches generalized gamma mixture calculations
Private Function Eval3ParamFlexMix(ByVal xVal As Double, _
                                   ByVal modeKey As String, _
                                   ByVal familyKey As String, _
                                   ByVal thetaVal As Double, _
                                   ByVal par1 As Double, _
                                   ByVal par2 As Double, _
                                   ByVal par3 As Double, _
                                   ByVal lowerTail As Boolean, _
                                   ByVal logP As Boolean, _
                                   ByVal logFlag As Boolean) As Variant
    Dim dens0 As Double, cdf0 As Double, surv0 As Double
    Dim outVal As Double, pWork As Double, pAdj As Double
    
    On Error GoTo EH
    
    thetaVal = ClampThetaMix(thetaVal)
    
    Select Case LCase$(Trim$(familyKey))
    
        Case "gengamma"
            Select Case LCase$(Trim$(modeKey))
                Case "d"
                    dens0 = GenGammaDensity_Core(xVal, par1, par2, par3)
                    outVal = dmixCore(dens0, thetaVal)
                    Eval3ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case "p"
                    cdf0 = GenGammaCDF_Core(xVal, par1, par2, par3)
                    outVal = pmixCore(cdf0, thetaVal)
                    Eval3ParamFlexMix = FinishProbFlex(outVal, lowerTail, logP)
                Case "q"
                    pWork = DecodeProbFlex(xVal, lowerTail, logP)
                    pAdj = qmixProbCore(pWork, thetaVal)
                    If pAdj >= BIG_POS Then
                        Eval3ParamFlexMix = BIG_POS
                    ElseIf pAdj <= 0# Then
                        Eval3ParamFlexMix = 0#
                    Else
                        Eval3ParamFlexMix = GenGammaQuantile_Core(pAdj, par1, par2, par3, True, False)
                    End If
                Case "h"
                    dens0 = GenGammaDensity_Core(xVal, par1, par2, par3)
                    surv0 = GenGammaSurv_Core(xVal, par1, par2, par3)
                    outVal = hmixCore(dens0, surv0, thetaVal)
                    Eval3ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case "ch"
                    surv0 = GenGammaSurv_Core(xVal, par1, par2, par3)
                    outVal = chmixCore(surv0, thetaVal)
                    Eval3ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case "s"
                    surv0 = GenGammaSurv_Core(xVal, par1, par2, par3)
                    outVal = SmixCore(surv0, thetaVal)
                    Eval3ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case Else
                    Eval3ParamFlexMix = CVErr(xlErrValue)
            End Select
        
        Case "gengamma_orig"
            Select Case LCase$(Trim$(modeKey))
                Case "d"
                    dens0 = GenGammaOrigDensity_Core(xVal, par1, par2, par3)
                    outVal = dmixCore(dens0, thetaVal)
                    Eval3ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case "p"
                    cdf0 = GenGammaOrigCDF_Core(xVal, par1, par2, par3)
                    outVal = pmixCore(cdf0, thetaVal)
                    Eval3ParamFlexMix = FinishProbFlex(outVal, lowerTail, logP)
                Case "q"
                    pWork = DecodeProbFlex(xVal, lowerTail, logP)
                    pAdj = qmixProbCore(pWork, thetaVal)
                    If pAdj >= BIG_POS Then
                        Eval3ParamFlexMix = BIG_POS
                    ElseIf pAdj <= 0# Then
                        Eval3ParamFlexMix = 0#
                    Else
                        Eval3ParamFlexMix = GenGammaOrigQuantile_Core(pAdj, par1, par2, par3, True, False)
                    End If
                Case "h"
                    dens0 = GenGammaOrigDensity_Core(xVal, par1, par2, par3)
                    surv0 = GenGammaOrigSurv_Core(xVal, par1, par2, par3)
                    outVal = hmixCore(dens0, surv0, thetaVal)
                    Eval3ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case "ch"
                    surv0 = GenGammaOrigSurv_Core(xVal, par1, par2, par3)
                    outVal = chmixCore(surv0, thetaVal)
                    Eval3ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case "s"
                    surv0 = GenGammaOrigSurv_Core(xVal, par1, par2, par3)
                    outVal = SmixCore(surv0, thetaVal)
                    Eval3ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case Else
                    Eval3ParamFlexMix = CVErr(xlErrValue)
            End Select
        
        Case Else
            Eval3ParamFlexMix = CVErr(xlErrValue)
    End Select
    
    Exit Function
    
EH:
    Eval3ParamFlexMix = CVErr(xlErrValue)
End Function

' Run3ParamFlexMix
' Applies 3-parameter mixture-cure evaluation to scalars, arrays, or Excel ranges
' Provides vectorized worksheet support for generalized gamma mixture functions
Private Function Run3ParamFlexMix(ByVal xInput As Variant, _
                                  ByVal modeKey As String, _
                                  ByVal familyKey As String, _
                                  ByVal thetaVal As Double, _
                                  ByVal par1 As Double, _
                                  ByVal par2 As Double, _
                                  ByVal par3 As Double, _
                                  Optional ByVal lowerTail As Boolean = True, _
                                  Optional ByVal logP As Boolean = False, _
                                  Optional ByVal logFlag As Boolean = False) As Variant
    Dim vals As Variant
    Dim outArr() As Variant
    Dim rVal As Long, cVal As Long
    Dim rngObj As Range
    
    On Error GoTo EH
    
    If IsObject(xInput) Then
        If TypeName(xInput) = "Range" Then
            Set rngObj = xInput
            
            If rngObj.CountLarge = 1 Then
                If IsNumeric(rngObj.Value2) Then
                    Run3ParamFlexMix = Eval3ParamFlexMix(CDbl(rngObj.Value2), modeKey, familyKey, thetaVal, par1, par2, par3, lowerTail, logP, logFlag)
                Else
                    Run3ParamFlexMix = CVErr(xlErrValue)
                End If
                Exit Function
            End If
            
            vals = rngObj.Value2
            ReDim outArr(1 To UBound(vals, 1), 1 To UBound(vals, 2))
            
            For rVal = 1 To UBound(vals, 1)
                For cVal = 1 To UBound(vals, 2)
                    If IsNumeric(vals(rVal, cVal)) Then
                        outArr(rVal, cVal) = Eval3ParamFlexMix(CDbl(vals(rVal, cVal)), modeKey, familyKey, thetaVal, par1, par2, par3, lowerTail, logP, logFlag)
                    Else
                        outArr(rVal, cVal) = CVErr(xlErrValue)
                    End If
                Next cVal
            Next rVal
            
            Run3ParamFlexMix = outArr
            Exit Function
        End If
    End If
    
    If IsArray(xInput) Then
        vals = xInput
        ReDim outArr(1 To UBound(vals, 1), 1 To UBound(vals, 2))
        
        For rVal = 1 To UBound(vals, 1)
            For cVal = 1 To UBound(vals, 2)
                If IsNumeric(vals(rVal, cVal)) Then
                    outArr(rVal, cVal) = Eval3ParamFlexMix(CDbl(vals(rVal, cVal)), modeKey, familyKey, thetaVal, par1, par2, par3, lowerTail, logP, logFlag)
                Else
                    outArr(rVal, cVal) = CVErr(xlErrValue)
                End If
            Next cVal
        Next rVal
        
        Run3ParamFlexMix = outArr
    ElseIf IsNumeric(xInput) Then
        Run3ParamFlexMix = Eval3ParamFlexMix(CDbl(xInput), modeKey, familyKey, thetaVal, par1, par2, par3, lowerTail, logP, logFlag)
    Else
        Run3ParamFlexMix = CVErr(xlErrValue)
    End If
    
    Exit Function
    
EH:
    Run3ParamFlexMix = CVErr(xlErrValue)
End Function

'==========================================================
' 4-parameter flex mixture evaluator
'==========================================================

' Eval4ParamFlexMix
' Evaluates 4-parameter flexible mixture-cure distributions by mode
' Dispatches generalized F mixture calculations
Private Function Eval4ParamFlexMix(ByVal xVal As Double, _
                                   ByVal modeKey As String, _
                                   ByVal familyKey As String, _
                                   ByVal thetaVal As Double, _
                                   ByVal par1 As Double, _
                                   ByVal par2 As Double, _
                                   ByVal par3 As Double, _
                                   ByVal par4 As Double, _
                                   ByVal lowerTail As Boolean, _
                                   ByVal logP As Boolean, _
                                   ByVal logFlag As Boolean) As Variant
    Dim dens0 As Double, cdf0 As Double, surv0 As Double
    Dim outVal As Double, pWork As Double, pAdj As Double
    
    On Error GoTo EH
    
    thetaVal = ClampThetaMix(thetaVal)
    
    Select Case LCase$(Trim$(familyKey))
    
        Case "genf"
            Select Case LCase$(Trim$(modeKey))
                Case "d"
                    dens0 = GenFDensity_Core(xVal, par1, par2, par3, par4)
                    outVal = dmixCore(dens0, thetaVal)
                    Eval4ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case "p"
                    cdf0 = GenFCDF_Core(xVal, par1, par2, par3, par4)
                    outVal = pmixCore(cdf0, thetaVal)
                    Eval4ParamFlexMix = FinishProbFlex(outVal, lowerTail, logP)
                Case "q"
                    pWork = DecodeProbFlex(xVal, lowerTail, logP)
                    pAdj = qmixProbCore(pWork, thetaVal)
                    If pAdj >= BIG_POS Then
                        Eval4ParamFlexMix = BIG_POS
                    ElseIf pAdj <= 0# Then
                        Eval4ParamFlexMix = 0#
                    Else
                        Eval4ParamFlexMix = GenFQuantile_Core(pAdj, par1, par2, par3, par4, True, False)
                    End If
                Case "h"
                    dens0 = GenFDensity_Core(xVal, par1, par2, par3, par4)
                    surv0 = GenFSurv_Core(xVal, par1, par2, par3, par4)
                    outVal = hmixCore(dens0, surv0, thetaVal)
                    Eval4ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case "ch"
                    surv0 = GenFSurv_Core(xVal, par1, par2, par3, par4)
                    outVal = chmixCore(surv0, thetaVal)
                    Eval4ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case "s"
                    surv0 = GenFSurv_Core(xVal, par1, par2, par3, par4)
                    outVal = SmixCore(surv0, thetaVal)
                    Eval4ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case Else
                    Eval4ParamFlexMix = CVErr(xlErrValue)
            End Select
        
        Case "genf_orig"
            Select Case LCase$(Trim$(modeKey))
                Case "d"
                    dens0 = GenFOrigDensity_Core(xVal, par1, par2, par3, par4)
                    outVal = dmixCore(dens0, thetaVal)
                    Eval4ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case "p"
                    cdf0 = GenFOrigCDF_Core(xVal, par1, par2, par3, par4)
                    outVal = pmixCore(cdf0, thetaVal)
                    Eval4ParamFlexMix = FinishProbFlex(outVal, lowerTail, logP)
                Case "q"
                    pWork = DecodeProbFlex(xVal, lowerTail, logP)
                    pAdj = qmixProbCore(pWork, thetaVal)
                    If pAdj >= BIG_POS Then
                        Eval4ParamFlexMix = BIG_POS
                    ElseIf pAdj <= 0# Then
                        Eval4ParamFlexMix = 0#
                    Else
                        Eval4ParamFlexMix = GenFOrigQuantile_Core(pAdj, par1, par2, par3, par4, True, False)
                    End If
                Case "h"
                    dens0 = GenFOrigDensity_Core(xVal, par1, par2, par3, par4)
                    surv0 = GenFOrigSurv_Core(xVal, par1, par2, par3, par4)
                    outVal = hmixCore(dens0, surv0, thetaVal)
                    Eval4ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case "ch"
                    surv0 = GenFOrigSurv_Core(xVal, par1, par2, par3, par4)
                    outVal = chmixCore(surv0, thetaVal)
                    Eval4ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case "s"
                    surv0 = GenFOrigSurv_Core(xVal, par1, par2, par3, par4)
                    outVal = SmixCore(surv0, thetaVal)
                    Eval4ParamFlexMix = IIf(logFlag, SafeLog(outVal), outVal)
                Case Else
                    Eval4ParamFlexMix = CVErr(xlErrValue)
            End Select
        
        Case Else
            Eval4ParamFlexMix = CVErr(xlErrValue)
    End Select
    
    Exit Function
    
EH:
    Eval4ParamFlexMix = CVErr(xlErrValue)
End Function


' Run4ParamFlexMix
' Applies 4-parameter mixture-cure evaluation to scalars, arrays, or Excel ranges
' Provides vectorized worksheet support for generalized F mixture functions
Private Function Run4ParamFlexMix(ByVal xInput As Variant, _
                                  ByVal modeKey As String, _
                                  ByVal familyKey As String, _
                                  ByVal thetaVal As Double, _
                                  ByVal par1 As Double, _
                                  ByVal par2 As Double, _
                                  ByVal par3 As Double, _
                                  ByVal par4 As Double, _
                                  Optional ByVal lowerTail As Boolean = True, _
                                  Optional ByVal logP As Boolean = False, _
                                  Optional ByVal logFlag As Boolean = False) As Variant
    Dim vals As Variant
    Dim outArr() As Variant
    Dim rVal As Long, cVal As Long
    Dim rngObj As Range
    
    On Error GoTo EH
    
    If IsObject(xInput) Then
        If TypeName(xInput) = "Range" Then
            Set rngObj = xInput
            
            If rngObj.CountLarge = 1 Then
                If IsNumeric(rngObj.Value2) Then
                    Run4ParamFlexMix = Eval4ParamFlexMix(CDbl(rngObj.Value2), modeKey, familyKey, thetaVal, par1, par2, par3, par4, lowerTail, logP, logFlag)
                Else
                    Run4ParamFlexMix = CVErr(xlErrValue)
                End If
                Exit Function
            End If
            
            vals = rngObj.Value2
            ReDim outArr(1 To UBound(vals, 1), 1 To UBound(vals, 2))
            
            For rVal = 1 To UBound(vals, 1)
                For cVal = 1 To UBound(vals, 2)
                    If IsNumeric(vals(rVal, cVal)) Then
                        outArr(rVal, cVal) = Eval4ParamFlexMix(CDbl(vals(rVal, cVal)), modeKey, familyKey, thetaVal, par1, par2, par3, par4, lowerTail, logP, logFlag)
                    Else
                        outArr(rVal, cVal) = CVErr(xlErrValue)
                    End If
                Next cVal
            Next rVal
            
            Run4ParamFlexMix = outArr
            Exit Function
        End If
    End If
    
    If IsArray(xInput) Then
        vals = xInput
        ReDim outArr(1 To UBound(vals, 1), 1 To UBound(vals, 2))
        
        For rVal = 1 To UBound(vals, 1)
            For cVal = 1 To UBound(vals, 2)
                If IsNumeric(vals(rVal, cVal)) Then
                    outArr(rVal, cVal) = Eval4ParamFlexMix(CDbl(vals(rVal, cVal)), modeKey, familyKey, thetaVal, par1, par2, par3, par4, lowerTail, logP, logFlag)
                Else
                    outArr(rVal, cVal) = CVErr(xlErrValue)
                End If
            Next cVal
        Next rVal
        
        Run4ParamFlexMix = outArr
    ElseIf IsNumeric(xInput) Then
        Run4ParamFlexMix = Eval4ParamFlexMix(CDbl(xInput), modeKey, familyKey, thetaVal, par1, par2, par3, par4, lowerTail, logP, logFlag)
    Else
        Run4ParamFlexMix = CVErr(xlErrValue)
    End If
    
    Exit Function
    
EH:
    Run4ParamFlexMix = CVErr(xlErrValue)
End Function

'==========================================================
' Public worksheet functions: Exponential mixture
'==========================================================

Public Function dmixexp(ByVal x As Variant, ByVal thetaInput As Double, ByVal rate_ As Double, _
                        Optional ByVal logFlag As Boolean = False) As Variant
    dmixexp = MixScalarOrArray(x, DIST_EXP, "d", ClampThetaMix(thetaInput), rate_, 0#, True, False, logFlag)
End Function

Public Function pmixexp(ByVal q As Variant, ByVal thetaInput As Double, ByVal rate_ As Double, _
                        Optional ByVal lowerTail As Boolean = True, _
                        Optional ByVal logP As Boolean = False) As Variant
    pmixexp = MixScalarOrArray(q, DIST_EXP, "p", ClampThetaMix(thetaInput), rate_, 0#, lowerTail, logP, False)
End Function

Public Function qmixexp(ByVal p As Variant, ByVal thetaInput As Double, ByVal rate_ As Double, _
                        Optional ByVal lowerTail As Boolean = True, _
                        Optional ByVal logP As Boolean = False) As Variant
    qmixexp = MixScalarOrArray(p, DIST_EXP, "q", ClampThetaMix(thetaInput), rate_, 0#, lowerTail, logP, False)
End Function

Public Function hmixexp(ByVal x As Variant, ByVal thetaInput As Double, ByVal rate_ As Double, _
                        Optional ByVal logFlag As Boolean = False) As Variant
    hmixexp = MixScalarOrArray(x, DIST_EXP, "h", ClampThetaMix(thetaInput), rate_, 0#, True, False, logFlag)
End Function

Public Function chmixexp(ByVal x As Variant, ByVal thetaInput As Double, ByVal rate_ As Double, _
                         Optional ByVal logFlag As Boolean = False) As Variant
    chmixexp = MixScalarOrArray(x, DIST_EXP, "ch", ClampThetaMix(thetaInput), rate_, 0#, True, False, logFlag)
End Function

Public Function Smixexp(ByVal x As Variant, ByVal thetaInput As Double, ByVal rate_ As Double, _
                        Optional ByVal logFlag As Boolean = False) As Variant
    Smixexp = MixScalarOrArray(x, DIST_EXP, "s", ClampThetaMix(thetaInput), rate_, 0#, True, False, logFlag)
End Function

'==========================================================
' Public worksheet functions: Weibull mixture
'==========================================================

Public Function dmixweibull(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                            Optional ByVal logFlag As Boolean = False) As Variant
    dmixweibull = MixScalarOrArray(x, DIST_WEIBULL, "d", ClampThetaMix(thetaInput), shape_, scale_, True, False, logFlag)
End Function

Public Function pmixweibull(ByVal q As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                            Optional ByVal lowerTail As Boolean = True, _
                            Optional ByVal logP As Boolean = False) As Variant
    pmixweibull = MixScalarOrArray(q, DIST_WEIBULL, "p", ClampThetaMix(thetaInput), shape_, scale_, lowerTail, logP, False)
End Function

Public Function qmixweibull(ByVal p As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                            Optional ByVal lowerTail As Boolean = True, _
                            Optional ByVal logP As Boolean = False) As Variant
    qmixweibull = MixScalarOrArray(p, DIST_WEIBULL, "q", ClampThetaMix(thetaInput), shape_, scale_, lowerTail, logP, False)
End Function

Public Function hmixweibull(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                            Optional ByVal logFlag As Boolean = False) As Variant
    hmixweibull = MixScalarOrArray(x, DIST_WEIBULL, "h", ClampThetaMix(thetaInput), shape_, scale_, True, False, logFlag)
End Function

Public Function chmixweibull(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                             Optional ByVal logFlag As Boolean = False) As Variant
    chmixweibull = MixScalarOrArray(x, DIST_WEIBULL, "ch", ClampThetaMix(thetaInput), shape_, scale_, True, False, logFlag)
End Function

Public Function Smixweibull(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                            Optional ByVal logFlag As Boolean = False) As Variant
    Smixweibull = MixScalarOrArray(x, DIST_WEIBULL, "s", ClampThetaMix(thetaInput), shape_, scale_, True, False, logFlag)
End Function

'==========================================================
' Public worksheet functions: WeibullPH mixture
'==========================================================

Public Function dmixweibullPH(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                              Optional ByVal logFlag As Boolean = False) As Variant
    dmixweibullPH = MixScalarOrArray(x, DIST_WEIBULLPH, "d", ClampThetaMix(thetaInput), shape_, scale_, True, False, logFlag)
End Function

Public Function pmixweibullPH(ByVal q As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                              Optional ByVal lowerTail As Boolean = True, _
                              Optional ByVal logP As Boolean = False) As Variant
    pmixweibullPH = MixScalarOrArray(q, DIST_WEIBULLPH, "p", ClampThetaMix(thetaInput), shape_, scale_, lowerTail, logP, False)
End Function

Public Function qmixweibullPH(ByVal p As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                              Optional ByVal lowerTail As Boolean = True, _
                              Optional ByVal logP As Boolean = False) As Variant
    qmixweibullPH = MixScalarOrArray(p, DIST_WEIBULLPH, "q", ClampThetaMix(thetaInput), shape_, scale_, lowerTail, logP, False)
End Function

Public Function hmixweibullPH(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                              Optional ByVal logFlag As Boolean = False) As Variant
    hmixweibullPH = MixScalarOrArray(x, DIST_WEIBULLPH, "h", ClampThetaMix(thetaInput), shape_, scale_, True, False, logFlag)
End Function

Public Function chmixweibullPH(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                               Optional ByVal logFlag As Boolean = False) As Variant
    chmixweibullPH = MixScalarOrArray(x, DIST_WEIBULLPH, "ch", ClampThetaMix(thetaInput), shape_, scale_, True, False, logFlag)
End Function

Public Function SmixweibullPH(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                              Optional ByVal logFlag As Boolean = False) As Variant
    SmixweibullPH = MixScalarOrArray(x, DIST_WEIBULLPH, "s", ClampThetaMix(thetaInput), shape_, scale_, True, False, logFlag)
End Function

'==========================================================
' Public worksheet functions: Gompertz mixture
'==========================================================

Public Function dmixgompertz(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal rate_ As Double, _
                             Optional ByVal logFlag As Boolean = False) As Variant
    dmixgompertz = MixScalarOrArray(x, DIST_GOMPERTZ, "d", ClampThetaMix(thetaInput), shape_, rate_, True, False, logFlag)
End Function

Public Function pmixgompertz(ByVal q As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal rate_ As Double, _
                             Optional ByVal lowerTail As Boolean = True, _
                             Optional ByVal logP As Boolean = False) As Variant
    pmixgompertz = MixScalarOrArray(q, DIST_GOMPERTZ, "p", ClampThetaMix(thetaInput), shape_, rate_, lowerTail, logP, False)
End Function

Public Function qmixgompertz(ByVal p As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal rate_ As Double, _
                             Optional ByVal lowerTail As Boolean = True, _
                             Optional ByVal logP As Boolean = False) As Variant
    qmixgompertz = MixScalarOrArray(p, DIST_GOMPERTZ, "q", ClampThetaMix(thetaInput), shape_, rate_, lowerTail, logP, False)
End Function

Public Function hmixgompertz(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal rate_ As Double, _
                             Optional ByVal logFlag As Boolean = False) As Variant
    hmixgompertz = MixScalarOrArray(x, DIST_GOMPERTZ, "h", ClampThetaMix(thetaInput), shape_, rate_, True, False, logFlag)
End Function

Public Function chmixgompertz(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal rate_ As Double, _
                              Optional ByVal logFlag As Boolean = False) As Variant
    chmixgompertz = MixScalarOrArray(x, DIST_GOMPERTZ, "ch", ClampThetaMix(thetaInput), shape_, rate_, True, False, logFlag)
End Function

Public Function Smixgompertz(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal rate_ As Double, _
                             Optional ByVal logFlag As Boolean = False) As Variant
    Smixgompertz = MixScalarOrArray(x, DIST_GOMPERTZ, "s", ClampThetaMix(thetaInput), shape_, rate_, True, False, logFlag)
End Function

'==========================================================
' Public worksheet functions: Log-normal mixture
'==========================================================

Public Function dmixlnorm(ByVal x As Variant, ByVal thetaInput As Double, ByVal meanlog_ As Double, ByVal sdlog_ As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    dmixlnorm = MixScalarOrArray(x, DIST_LNORM, "d", ClampThetaMix(thetaInput), meanlog_, sdlog_, True, False, logFlag)
End Function

Public Function pmixlnorm(ByVal q As Variant, ByVal thetaInput As Double, ByVal meanlog_ As Double, ByVal sdlog_ As Double, _
                          Optional ByVal lowerTail As Boolean = True, _
                          Optional ByVal logP As Boolean = False) As Variant
    pmixlnorm = MixScalarOrArray(q, DIST_LNORM, "p", ClampThetaMix(thetaInput), meanlog_, sdlog_, lowerTail, logP, False)
End Function

Public Function qmixlnorm(ByVal p As Variant, ByVal thetaInput As Double, ByVal meanlog_ As Double, ByVal sdlog_ As Double, _
                          Optional ByVal lowerTail As Boolean = True, _
                          Optional ByVal logP As Boolean = False) As Variant
    qmixlnorm = MixScalarOrArray(p, DIST_LNORM, "q", ClampThetaMix(thetaInput), meanlog_, sdlog_, lowerTail, logP, False)
End Function

Public Function hmixlnorm(ByVal x As Variant, ByVal thetaInput As Double, ByVal meanlog_ As Double, ByVal sdlog_ As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    hmixlnorm = MixScalarOrArray(x, DIST_LNORM, "h", ClampThetaMix(thetaInput), meanlog_, sdlog_, True, False, logFlag)
End Function

Public Function chmixlnorm(ByVal x As Variant, ByVal thetaInput As Double, ByVal meanlog_ As Double, ByVal sdlog_ As Double, _
                           Optional ByVal logFlag As Boolean = False) As Variant
    chmixlnorm = MixScalarOrArray(x, DIST_LNORM, "ch", ClampThetaMix(thetaInput), meanlog_, sdlog_, True, False, logFlag)
End Function

Public Function Smixlnorm(ByVal x As Variant, ByVal thetaInput As Double, ByVal meanlog_ As Double, ByVal sdlog_ As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    Smixlnorm = MixScalarOrArray(x, DIST_LNORM, "s", ClampThetaMix(thetaInput), meanlog_, sdlog_, True, False, logFlag)
End Function

'==========================================================
' Public worksheet functions: Log-logistic mixture
'==========================================================

Public Function dmixllogis(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                           Optional ByVal logFlag As Boolean = False) As Variant
    dmixllogis = MixScalarOrArray(x, DIST_LLOGIS, "d", ClampThetaMix(thetaInput), shape_, scale_, True, False, logFlag)
End Function

Public Function pmixllogis(ByVal q As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                           Optional ByVal lowerTail As Boolean = True, _
                           Optional ByVal logP As Boolean = False) As Variant
    pmixllogis = MixScalarOrArray(q, DIST_LLOGIS, "p", ClampThetaMix(thetaInput), shape_, scale_, lowerTail, logP, False)
End Function

Public Function qmixllogis(ByVal p As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                           Optional ByVal lowerTail As Boolean = True, _
                           Optional ByVal logP As Boolean = False) As Variant
    qmixllogis = MixScalarOrArray(p, DIST_LLOGIS, "q", ClampThetaMix(thetaInput), shape_, scale_, lowerTail, logP, False)
End Function

Public Function hmixllogis(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                           Optional ByVal logFlag As Boolean = False) As Variant
    hmixllogis = MixScalarOrArray(x, DIST_LLOGIS, "h", ClampThetaMix(thetaInput), shape_, scale_, True, False, logFlag)
End Function

Public Function chmixllogis(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                            Optional ByVal logFlag As Boolean = False) As Variant
    chmixllogis = MixScalarOrArray(x, DIST_LLOGIS, "ch", ClampThetaMix(thetaInput), shape_, scale_, True, False, logFlag)
End Function

Public Function Smixllogis(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal scale_ As Double, _
                           Optional ByVal logFlag As Boolean = False) As Variant
    Smixllogis = MixScalarOrArray(x, DIST_LLOGIS, "s", ClampThetaMix(thetaInput), shape_, scale_, True, False, logFlag)
End Function

'==========================================================
' Public worksheet functions: Gamma mixture
'==========================================================

Public Function dmixgamma(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal rate_ As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    dmixgamma = MixScalarOrArray(x, DIST_GAMMA, "d", ClampThetaMix(thetaInput), shape_, rate_, True, False, logFlag)
End Function

Public Function pmixgamma(ByVal q As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal rate_ As Double, _
                          Optional ByVal lowerTail As Boolean = True, _
                          Optional ByVal logP As Boolean = False) As Variant
    pmixgamma = MixScalarOrArray(q, DIST_GAMMA, "p", ClampThetaMix(thetaInput), shape_, rate_, lowerTail, logP, False)
End Function

Public Function qmixgamma(ByVal p As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal rate_ As Double, _
                          Optional ByVal lowerTail As Boolean = True, _
                          Optional ByVal logP As Boolean = False) As Variant
    qmixgamma = MixScalarOrArray(p, DIST_GAMMA, "q", ClampThetaMix(thetaInput), shape_, rate_, lowerTail, logP, False)
End Function

Public Function hmixgamma(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal rate_ As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    hmixgamma = MixScalarOrArray(x, DIST_GAMMA, "h", ClampThetaMix(thetaInput), shape_, rate_, True, False, logFlag)
End Function

Public Function chmixgamma(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal rate_ As Double, _
                           Optional ByVal logFlag As Boolean = False) As Variant
    chmixgamma = MixScalarOrArray(x, DIST_GAMMA, "ch", ClampThetaMix(thetaInput), shape_, rate_, True, False, logFlag)
End Function

Public Function Smixgamma(ByVal x As Variant, ByVal thetaInput As Double, ByVal shape_ As Double, ByVal rate_ As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    Smixgamma = MixScalarOrArray(x, DIST_GAMMA, "s", ClampThetaMix(thetaInput), shape_, rate_, True, False, logFlag)
End Function

'==========================================================
' Public worksheet functions: gengamma mixture
'==========================================================

Public Function dmixgengamma(ByVal xInput As Variant, _
                             ByVal thetaInput As Double, _
                             ByVal muVal As Double, _
                             ByVal sigmaVal As Double, _
                             ByVal qParam As Double, _
                             Optional ByVal logFlag As Boolean = False) As Variant
    dmixgengamma = Run3ParamFlexMix(xInput, "d", "gengamma", ClampThetaMix(thetaInput), muVal, sigmaVal, qParam, True, False, logFlag)
End Function

Public Function pmixgengamma(ByVal xInput As Variant, _
                             ByVal thetaInput As Double, _
                             ByVal muVal As Double, _
                             ByVal sigmaVal As Double, _
                             ByVal qParam As Double, _
                             Optional ByVal lowerTail As Boolean = True, _
                             Optional ByVal logP As Boolean = False) As Variant
    pmixgengamma = Run3ParamFlexMix(xInput, "p", "gengamma", ClampThetaMix(thetaInput), muVal, sigmaVal, qParam, lowerTail, logP, False)
End Function

Public Function qmixgengamma(ByVal probInput As Variant, _
                             ByVal thetaInput As Double, _
                             ByVal muVal As Double, _
                             ByVal sigmaVal As Double, _
                             ByVal qParam As Double, _
                             Optional ByVal lowerTail As Boolean = True, _
                             Optional ByVal logP As Boolean = False) As Variant
    qmixgengamma = Run3ParamFlexMix(probInput, "q", "gengamma", ClampThetaMix(thetaInput), muVal, sigmaVal, qParam, lowerTail, logP, False)
End Function

Public Function hmixgengamma(ByVal xInput As Variant, _
                             ByVal thetaInput As Double, _
                             ByVal muVal As Double, _
                             ByVal sigmaVal As Double, _
                             ByVal qParam As Double, _
                             Optional ByVal logFlag As Boolean = False) As Variant
    hmixgengamma = Run3ParamFlexMix(xInput, "h", "gengamma", ClampThetaMix(thetaInput), muVal, sigmaVal, qParam, True, False, logFlag)
End Function

Public Function chmixgengamma(ByVal xInput As Variant, _
                              ByVal thetaInput As Double, _
                              ByVal muVal As Double, _
                              ByVal sigmaVal As Double, _
                              ByVal qParam As Double, _
                              Optional ByVal logFlag As Boolean = False) As Variant
    chmixgengamma = Run3ParamFlexMix(xInput, "ch", "gengamma", ClampThetaMix(thetaInput), muVal, sigmaVal, qParam, True, False, logFlag)
End Function

Public Function Smixgengamma(ByVal xInput As Variant, _
                             ByVal thetaInput As Double, _
                             ByVal muVal As Double, _
                             ByVal sigmaVal As Double, _
                             ByVal qParam As Double, _
                             Optional ByVal logFlag As Boolean = False) As Variant
    Smixgengamma = Run3ParamFlexMix(xInput, "s", "gengamma", ClampThetaMix(thetaInput), muVal, sigmaVal, qParam, True, False, logFlag)
End Function

'==========================================================
' Public worksheet functions: gengamma_orig mixture
'==========================================================

Public Function dmixgengamma_orig(ByVal xInput As Variant, _
                                  ByVal thetaInput As Double, _
                                  ByVal shapeVal As Double, _
                                  ByVal scaleVal As Double, _
                                  ByVal kVal As Double, _
                                  Optional ByVal logFlag As Boolean = False) As Variant
    dmixgengamma_orig = Run3ParamFlexMix(xInput, "d", "gengamma_orig", ClampThetaMix(thetaInput), shapeVal, scaleVal, kVal, True, False, logFlag)
End Function

Public Function pmixgengamma_orig(ByVal xInput As Variant, _
                                  ByVal thetaInput As Double, _
                                  ByVal shapeVal As Double, _
                                  ByVal scaleVal As Double, _
                                  ByVal kVal As Double, _
                                  Optional ByVal lowerTail As Boolean = True, _
                                  Optional ByVal logP As Boolean = False) As Variant
    pmixgengamma_orig = Run3ParamFlexMix(xInput, "p", "gengamma_orig", ClampThetaMix(thetaInput), shapeVal, scaleVal, kVal, lowerTail, logP, False)
End Function

Public Function qmixgengamma_orig(ByVal probInput As Variant, _
                                  ByVal thetaInput As Double, _
                                  ByVal shapeVal As Double, _
                                  ByVal scaleVal As Double, _
                                  ByVal kVal As Double, _
                                  Optional ByVal lowerTail As Boolean = True, _
                                  Optional ByVal logP As Boolean = False) As Variant
    qmixgengamma_orig = Run3ParamFlexMix(probInput, "q", "gengamma_orig", ClampThetaMix(thetaInput), shapeVal, scaleVal, kVal, lowerTail, logP, False)
End Function

Public Function hmixgengamma_orig(ByVal xInput As Variant, _
                                  ByVal thetaInput As Double, _
                                  ByVal shapeVal As Double, _
                                  ByVal scaleVal As Double, _
                                  ByVal kVal As Double, _
                                  Optional ByVal logFlag As Boolean = False) As Variant
    hmixgengamma_orig = Run3ParamFlexMix(xInput, "h", "gengamma_orig", ClampThetaMix(thetaInput), shapeVal, scaleVal, kVal, True, False, logFlag)
End Function

Public Function chmixgengamma_orig(ByVal xInput As Variant, _
                                   ByVal thetaInput As Double, _
                                   ByVal shapeVal As Double, _
                                   ByVal scaleVal As Double, _
                                   ByVal kVal As Double, _
                                   Optional ByVal logFlag As Boolean = False) As Variant
    chmixgengamma_orig = Run3ParamFlexMix(xInput, "ch", "gengamma_orig", ClampThetaMix(thetaInput), shapeVal, scaleVal, kVal, True, False, logFlag)
End Function

Public Function Smixgengamma_orig(ByVal xInput As Variant, _
                                  ByVal thetaInput As Double, _
                                  ByVal shapeVal As Double, _
                                  ByVal scaleVal As Double, _
                                  ByVal kVal As Double, _
                                  Optional ByVal logFlag As Boolean = False) As Variant
    Smixgengamma_orig = Run3ParamFlexMix(xInput, "s", "gengamma_orig", ClampThetaMix(thetaInput), shapeVal, scaleVal, kVal, True, False, logFlag)
End Function

'==========================================================
' Public worksheet functions: genf mixture
'==========================================================

Public Function dmixgenf(ByVal xInput As Variant, _
                         ByVal thetaInput As Double, _
                         ByVal muVal As Double, _
                         ByVal sigmaVal As Double, _
                         ByVal qParam As Double, _
                         ByVal pParam As Double, _
                         Optional ByVal logFlag As Boolean = False) As Variant
    dmixgenf = Run4ParamFlexMix(xInput, "d", "genf", ClampThetaMix(thetaInput), muVal, sigmaVal, qParam, pParam, True, False, logFlag)
End Function

Public Function pmixgenf(ByVal xInput As Variant, _
                         ByVal thetaInput As Double, _
                         ByVal muVal As Double, _
                         ByVal sigmaVal As Double, _
                         ByVal qParam As Double, _
                         ByVal pParam As Double, _
                         Optional ByVal lowerTail As Boolean = True, _
                         Optional ByVal logP As Boolean = False) As Variant
    pmixgenf = Run4ParamFlexMix(xInput, "p", "genf", ClampThetaMix(thetaInput), muVal, sigmaVal, qParam, pParam, lowerTail, logP, False)
End Function

Public Function qmixgenf(ByVal probInput As Variant, _
                         ByVal thetaInput As Double, _
                         ByVal muVal As Double, _
                         ByVal sigmaVal As Double, _
                         ByVal qParam As Double, _
                         ByVal pParam As Double, _
                         Optional ByVal lowerTail As Boolean = True, _
                         Optional ByVal logP As Boolean = False) As Variant
    qmixgenf = Run4ParamFlexMix(probInput, "q", "genf", ClampThetaMix(thetaInput), muVal, sigmaVal, qParam, pParam, lowerTail, logP, False)
End Function

Public Function hmixgenf(ByVal xInput As Variant, _
                         ByVal thetaInput As Double, _
                         ByVal muVal As Double, _
                         ByVal sigmaVal As Double, _
                         ByVal qParam As Double, _
                         ByVal pParam As Double, _
                         Optional ByVal logFlag As Boolean = False) As Variant
    hmixgenf = Run4ParamFlexMix(xInput, "h", "genf", ClampThetaMix(thetaInput), muVal, sigmaVal, qParam, pParam, True, False, logFlag)
End Function

Public Function chmixgenf(ByVal xInput As Variant, _
                          ByVal thetaInput As Double, _
                          ByVal muVal As Double, _
                          ByVal sigmaVal As Double, _
                          ByVal qParam As Double, _
                          ByVal pParam As Double, _
                          Optional ByVal logFlag As Boolean = False) As Variant
    chmixgenf = Run4ParamFlexMix(xInput, "ch", "genf", ClampThetaMix(thetaInput), muVal, sigmaVal, qParam, pParam, True, False, logFlag)
End Function

Public Function Smixgenf(ByVal xInput As Variant, _
                         ByVal thetaInput As Double, _
                         ByVal muVal As Double, _
                         ByVal sigmaVal As Double, _
                         ByVal qParam As Double, _
                         ByVal pParam As Double, _
                         Optional ByVal logFlag As Boolean = False) As Variant
    Smixgenf = Run4ParamFlexMix(xInput, "s", "genf", ClampThetaMix(thetaInput), muVal, sigmaVal, qParam, pParam, True, False, logFlag)
End Function

'==========================================================
' Public worksheet functions: genf_orig mixture
'==========================================================

Public Function dmixgenf_orig(ByVal xInput As Variant, _
                              ByVal thetaInput As Double, _
                              ByVal muVal As Double, _
                              ByVal sigmaVal As Double, _
                              ByVal s1Val As Double, _
                              ByVal s2Val As Double, _
                              Optional ByVal logFlag As Boolean = False) As Variant
    dmixgenf_orig = Run4ParamFlexMix(xInput, "d", "genf_orig", ClampThetaMix(thetaInput), muVal, sigmaVal, s1Val, s2Val, True, False, logFlag)
End Function

Public Function pmixgenf_orig(ByVal xInput As Variant, _
                              ByVal thetaInput As Double, _
                              ByVal muVal As Double, _
                              ByVal sigmaVal As Double, _
                              ByVal s1Val As Double, _
                              ByVal s2Val As Double, _
                              Optional ByVal lowerTail As Boolean = True, _
                              Optional ByVal logP As Boolean = False) As Variant
    pmixgenf_orig = Run4ParamFlexMix(xInput, "p", "genf_orig", ClampThetaMix(thetaInput), muVal, sigmaVal, s1Val, s2Val, lowerTail, logP, False)
End Function

Public Function qmixgenf_orig(ByVal probInput As Variant, _
                              ByVal thetaInput As Double, _
                              ByVal muVal As Double, _
                              ByVal sigmaVal As Double, _
                              ByVal s1Val As Double, _
                              ByVal s2Val As Double, _
                              Optional ByVal lowerTail As Boolean = True, _
                              Optional ByVal logP As Boolean = False) As Variant
    qmixgenf_orig = Run4ParamFlexMix(probInput, "q", "genf_orig", ClampThetaMix(thetaInput), muVal, sigmaVal, s1Val, s2Val, lowerTail, logP, False)
End Function

Public Function hmixgenf_orig(ByVal xInput As Variant, _
                              ByVal thetaInput As Double, _
                              ByVal muVal As Double, _
                              ByVal sigmaVal As Double, _
                              ByVal s1Val As Double, _
                              ByVal s2Val As Double, _
                              Optional ByVal logFlag As Boolean = False) As Variant
    hmixgenf_orig = Run4ParamFlexMix(xInput, "h", "genf_orig", ClampThetaMix(thetaInput), muVal, sigmaVal, s1Val, s2Val, True, False, logFlag)
End Function

Public Function chmixgenf_orig(ByVal xInput As Variant, _
                               ByVal thetaInput As Double, _
                               ByVal muVal As Double, _
                               ByVal sigmaVal As Double, _
                               ByVal s1Val As Double, _
                               ByVal s2Val As Double, _
                               Optional ByVal logFlag As Boolean = False) As Variant
    chmixgenf_orig = Run4ParamFlexMix(xInput, "ch", "genf_orig", ClampThetaMix(thetaInput), muVal, sigmaVal, s1Val, s2Val, True, False, logFlag)
End Function

Public Function Smixgenf_orig(ByVal xInput As Variant, _
                              ByVal thetaInput As Double, _
                              ByVal muVal As Double, _
                              ByVal sigmaVal As Double, _
                              ByVal s1Val As Double, _
                              ByVal s2Val As Double, _
                              Optional ByVal logFlag As Boolean = False) As Variant
    Smixgenf_orig = Run4ParamFlexMix(xInput, "s", "genf_orig", ClampThetaMix(thetaInput), muVal, sigmaVal, s1Val, s2Val, True, False, logFlag)
End Function

'==========================================================
' RMST FUNCTIONS
'
' This module assumes your existing RiskFunctions module
' already contains:
'
'   DIST_EXP, DIST_WEIBULL, DIST_WEIBULLPH, DIST_GOMPERTZ,
'   DIST_LNORM, DIST_LLOGIS, DIST_GAMMA
'
'   BIG_POS
'   Sexp, Sweibull, SweibullPH, Sgompertz, Slnorm, Sllogis, Sgamma
'   Sgengamma, Sgengamma_orig, Sgenf, Sgenf_orig
'   Smixexp, Smixweibull, SmixweibullPH, Smixgompertz, Smixlnorm,
'   Smixllogis, Smixgamma
'   Smixgengamma, Smixgengamma_orig, Smixgenf, Smixgenf_orig
'   ClampThetaMix
'
' RMST(tau) = integral from 0 to tau of S(t) dt
'
' The implementation uses adaptive trapezoidal integration and
' supports scalar values, Excel ranges, and arrays for tau.
'==========================================================

'==========================================================
' Basic RMST helpers
'==========================================================

Private Function RmstClamp01(ByVal x As Double) As Double
    If x < 0# Then
        RmstClamp01 = 0#
    ElseIf x > 1# Then
        RmstClamp01 = 1#
    Else
        RmstClamp01 = x
    End If
End Function

Private Function RmstAsDouble(ByVal v As Variant) As Variant
    If IsError(v) Then
        RmstAsDouble = CVErr(xlErrValue)
    ElseIf IsNumeric(v) Then
        RmstAsDouble = CDbl(v)
    Else
        RmstAsDouble = CVErr(xlErrValue)
    End If
End Function

'==========================================================
' Standard 1-2 parameter distributions
'==========================================================

Private Function RmstSurvStd(ByVal t As Double, _
                             ByVal distCode As Long, _
                             ByVal p1 As Double, _
                             Optional ByVal p2 As Double = 0#) As Variant
    Dim v As Variant
    
    If t <= 0# Then
        RmstSurvStd = 1#
        Exit Function
    End If
    
    Select Case distCode
        Case DIST_EXP
            v = Sexp(t, p1, False)
        Case DIST_WEIBULL
            v = Sweibull(t, p1, p2, False)
        Case DIST_WEIBULLPH
            v = SweibullPH(t, p1, p2, False)
        Case DIST_GOMPERTZ
            v = Sgompertz(t, p1, p2, False)
        Case DIST_LNORM
            v = Slnorm(t, p1, p2, False)
        Case DIST_LLOGIS
            v = Sllogis(t, p1, p2, False)
        Case DIST_GAMMA
            v = Sgamma(t, p1, p2, False)
        Case Else
            RmstSurvStd = CVErr(xlErrValue)
            Exit Function
    End Select
    
    If IsError(v) Then
        RmstSurvStd = CVErr(xlErrValue)
    Else
        RmstSurvStd = RmstClamp01(CDbl(v))
    End If
End Function

Private Function RmstTrapzStd(ByVal tau As Double, _
                              ByVal distCode As Long, _
                              ByVal p1 As Double, _
                              Optional ByVal p2 As Double = 0#, _
                              Optional ByVal nSteps As Long = RMST_MIN_STEPS) As Variant
    Dim h As Double
    Dim i As Long
    Dim s0 As Variant, s1 As Variant
    Dim area As Double
    Dim t1 As Double
    
    If tau < 0# Then
        RmstTrapzStd = CVErr(xlErrNum)
        Exit Function
    End If
    
    If tau = 0# Then
        RmstTrapzStd = 0#
        Exit Function
    End If
    
    If nSteps < 1 Then nSteps = 1
    
    h = tau / nSteps
    s0 = RmstSurvStd(0#, distCode, p1, p2)
    
    If IsError(s0) Then
        RmstTrapzStd = s0
        Exit Function
    End If
    
    area = 0#
    
    For i = 1 To nSteps
        t1 = i * h
        s1 = RmstSurvStd(t1, distCode, p1, p2)
        
        If IsError(s1) Then
            RmstTrapzStd = s1
            Exit Function
        End If
        
        area = area + 0.5 * (CDbl(s0) + CDbl(s1)) * h
        s0 = s1
    Next i
    
    RmstTrapzStd = area
End Function

Private Function RmstOneStd(ByVal tau As Double, _
                            ByVal distCode As Long, _
                            ByVal p1 As Double, _
                            Optional ByVal p2 As Double = 0#) As Variant
    Dim n As Long
    Dim a0 As Variant, a1 As Variant
    Dim errAbs As Double
    Dim scaleVal As Double
    
    If tau < 0# Then
        RmstOneStd = CVErr(xlErrNum)
        Exit Function
    End If
    
    If tau = 0# Then
        RmstOneStd = 0#
        Exit Function
    End If
    
    n = RMST_MIN_STEPS
    a0 = RmstTrapzStd(tau, distCode, p1, p2, n)
    
    If IsError(a0) Then
        RmstOneStd = a0
        Exit Function
    End If
    
    Do While n < RMST_MAX_STEPS
        n = n * 2
        a1 = RmstTrapzStd(tau, distCode, p1, p2, n)
        
        If IsError(a1) Then
            RmstOneStd = a1
            Exit Function
        End If
        
        errAbs = Abs(CDbl(a1) - CDbl(a0))
        scaleVal = 1# + Abs(CDbl(a1))
        
        If errAbs <= RMST_REL_TOL * scaleVal Then
            RmstOneStd = a1
            Exit Function
        End If
        
        a0 = a1
    Loop
    
    RmstOneStd = a0
End Function

Private Function RmstScalarOrArrayStd(ByVal tauInput As Variant, _
                                      ByVal distCode As Long, _
                                      ByVal p1 As Double, _
                                      Optional ByVal p2 As Double = 0#) As Variant
    Dim vals As Variant
    Dim out() As Variant
    Dim r As Long, c As Long
    Dim rng As Range
    
    On Error GoTo EH
    
    If IsObject(tauInput) Then
        If TypeName(tauInput) = "Range" Then
            Set rng = tauInput
            
            If rng.CountLarge = 1 Then
                If IsNumeric(rng.Value2) Then
                    RmstScalarOrArrayStd = RmstOneStd(CDbl(rng.Value2), distCode, p1, p2)
                Else
                    RmstScalarOrArrayStd = CVErr(xlErrValue)
                End If
                Exit Function
            End If
            
            vals = rng.Value2
            ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
            
            For r = 1 To UBound(vals, 1)
                For c = 1 To UBound(vals, 2)
                    If IsNumeric(vals(r, c)) Then
                        out(r, c) = RmstOneStd(CDbl(vals(r, c)), distCode, p1, p2)
                    Else
                        out(r, c) = CVErr(xlErrValue)
                    End If
                Next c
            Next r
            
            RmstScalarOrArrayStd = out
            Exit Function
        End If
    End If
    
    If IsArray(tauInput) Then
        vals = tauInput
        ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
        
        For r = 1 To UBound(vals, 1)
            For c = 1 To UBound(vals, 2)
                If IsNumeric(vals(r, c)) Then
                    out(r, c) = RmstOneStd(CDbl(vals(r, c)), distCode, p1, p2)
                Else
                    out(r, c) = CVErr(xlErrValue)
                End If
            Next c
        Next r
        
        RmstScalarOrArrayStd = out
    ElseIf IsNumeric(tauInput) Then
        RmstScalarOrArrayStd = RmstOneStd(CDbl(tauInput), distCode, p1, p2)
    Else
        RmstScalarOrArrayStd = CVErr(xlErrValue)
    End If
    
    Exit Function
    
EH:
    RmstScalarOrArrayStd = CVErr(xlErrValue)
End Function

'==========================================================
' Flexible 3-parameter distributions
'==========================================================

Private Function RmstSurvFlex3(ByVal t As Double, _
                               ByVal distName As String, _
                               ByVal p1 As Double, _
                               ByVal p2 As Double, _
                               ByVal p3 As Double) As Variant
    Dim v As Variant
    
    If t <= 0# Then
        RmstSurvFlex3 = 1#
        Exit Function
    End If
    
    Select Case LCase$(Trim$(distName))
        Case "gengamma"
            v = Sgengamma(t, p1, p2, p3, False)
        Case "gengamma_orig"
            v = Sgengamma_orig(t, p1, p2, p3, False)
        Case Else
            RmstSurvFlex3 = CVErr(xlErrValue)
            Exit Function
    End Select
    
    If IsError(v) Then
        RmstSurvFlex3 = CVErr(xlErrValue)
    Else
        RmstSurvFlex3 = RmstClamp01(CDbl(v))
    End If
End Function

Private Function RmstTrapzFlex3(ByVal tau As Double, _
                                ByVal distName As String, _
                                ByVal p1 As Double, _
                                ByVal p2 As Double, _
                                ByVal p3 As Double, _
                                Optional ByVal nSteps As Long = RMST_MIN_STEPS) As Variant
    Dim h As Double
    Dim i As Long
    Dim s0 As Variant, s1 As Variant
    Dim area As Double
    Dim t1 As Double
    
    If tau < 0# Then
        RmstTrapzFlex3 = CVErr(xlErrNum)
        Exit Function
    End If
    
    If tau = 0# Then
        RmstTrapzFlex3 = 0#
        Exit Function
    End If
    
    If nSteps < 1 Then nSteps = 1
    
    h = tau / nSteps
    s0 = RmstSurvFlex3(0#, distName, p1, p2, p3)
    
    If IsError(s0) Then
        RmstTrapzFlex3 = s0
        Exit Function
    End If
    
    area = 0#
    
    For i = 1 To nSteps
        t1 = i * h
        s1 = RmstSurvFlex3(t1, distName, p1, p2, p3)
        
        If IsError(s1) Then
            RmstTrapzFlex3 = s1
            Exit Function
        End If
        
        area = area + 0.5 * (CDbl(s0) + CDbl(s1)) * h
        s0 = s1
    Next i
    
    RmstTrapzFlex3 = area
End Function

Private Function RmstOneFlex3(ByVal tau As Double, _
                              ByVal distName As String, _
                              ByVal p1 As Double, _
                              ByVal p2 As Double, _
                              ByVal p3 As Double) As Variant
    Dim n As Long
    Dim a0 As Variant, a1 As Variant
    Dim errAbs As Double
    Dim scaleVal As Double
    
    If tau < 0# Then
        RmstOneFlex3 = CVErr(xlErrNum)
        Exit Function
    End If
    
    If tau = 0# Then
        RmstOneFlex3 = 0#
        Exit Function
    End If
    
    n = RMST_MIN_STEPS
    a0 = RmstTrapzFlex3(tau, distName, p1, p2, p3, n)
    
    If IsError(a0) Then
        RmstOneFlex3 = a0
        Exit Function
    End If
    
    Do While n < RMST_MAX_STEPS
        n = n * 2
        a1 = RmstTrapzFlex3(tau, distName, p1, p2, p3, n)
        
        If IsError(a1) Then
            RmstOneFlex3 = a1
            Exit Function
        End If
        
        errAbs = Abs(CDbl(a1) - CDbl(a0))
        scaleVal = 1# + Abs(CDbl(a1))
        
        If errAbs <= RMST_REL_TOL * scaleVal Then
            RmstOneFlex3 = a1
            Exit Function
        End If
        
        a0 = a1
    Loop
    
    RmstOneFlex3 = a0
End Function

Private Function RmstScalarOrArrayFlex3(ByVal tauInput As Variant, _
                                        ByVal distName As String, _
                                        ByVal p1 As Double, _
                                        ByVal p2 As Double, _
                                        ByVal p3 As Double) As Variant
    Dim vals As Variant
    Dim out() As Variant
    Dim r As Long, c As Long
    Dim rng As Range
    
    On Error GoTo EH
    
    If IsObject(tauInput) Then
        If TypeName(tauInput) = "Range" Then
            Set rng = tauInput
            
            If rng.CountLarge = 1 Then
                If IsNumeric(rng.Value2) Then
                    RmstScalarOrArrayFlex3 = RmstOneFlex3(CDbl(rng.Value2), distName, p1, p2, p3)
                Else
                    RmstScalarOrArrayFlex3 = CVErr(xlErrValue)
                End If
                Exit Function
            End If
            
            vals = rng.Value2
            ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
            
            For r = 1 To UBound(vals, 1)
                For c = 1 To UBound(vals, 2)
                    If IsNumeric(vals(r, c)) Then
                        out(r, c) = RmstOneFlex3(CDbl(vals(r, c)), distName, p1, p2, p3)
                    Else
                        out(r, c) = CVErr(xlErrValue)
                    End If
                Next c
            Next r
            
            RmstScalarOrArrayFlex3 = out
            Exit Function
        End If
    End If
    
    If IsArray(tauInput) Then
        vals = tauInput
        ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
        
        For r = 1 To UBound(vals, 1)
            For c = 1 To UBound(vals, 2)
                If IsNumeric(vals(r, c)) Then
                    out(r, c) = RmstOneFlex3(CDbl(vals(r, c)), distName, p1, p2, p3)
                Else
                    out(r, c) = CVErr(xlErrValue)
                End If
            Next c
        Next r
        
        RmstScalarOrArrayFlex3 = out
    ElseIf IsNumeric(tauInput) Then
        RmstScalarOrArrayFlex3 = RmstOneFlex3(CDbl(tauInput), distName, p1, p2, p3)
    Else
        RmstScalarOrArrayFlex3 = CVErr(xlErrValue)
    End If
    
    Exit Function
    
EH:
    RmstScalarOrArrayFlex3 = CVErr(xlErrValue)
End Function

'==========================================================
' Flexible 4-parameter distributions
'==========================================================

Private Function RmstSurvFlex4(ByVal t As Double, _
                               ByVal distName As String, _
                               ByVal p1 As Double, _
                               ByVal p2 As Double, _
                               ByVal p3 As Double, _
                               ByVal p4 As Double) As Variant
    Dim v As Variant
    
    If t <= 0# Then
        RmstSurvFlex4 = 1#
        Exit Function
    End If
    
    Select Case LCase$(Trim$(distName))
        Case "genf"
            v = Sgenf(t, p1, p2, p3, p4, False)
        Case "genf_orig"
            v = Sgenf_orig(t, p1, p2, p3, p4, False)
        Case Else
            RmstSurvFlex4 = CVErr(xlErrValue)
            Exit Function
    End Select
    
    If IsError(v) Then
        RmstSurvFlex4 = CVErr(xlErrValue)
    Else
        RmstSurvFlex4 = RmstClamp01(CDbl(v))
    End If
End Function

Private Function RmstTrapzFlex4(ByVal tau As Double, _
                                ByVal distName As String, _
                                ByVal p1 As Double, _
                                ByVal p2 As Double, _
                                ByVal p3 As Double, _
                                ByVal p4 As Double, _
                                Optional ByVal nSteps As Long = RMST_MIN_STEPS) As Variant
    Dim h As Double
    Dim i As Long
    Dim s0 As Variant, s1 As Variant
    Dim area As Double
    Dim t1 As Double
    
    If tau < 0# Then
        RmstTrapzFlex4 = CVErr(xlErrNum)
        Exit Function
    End If
    
    If tau = 0# Then
        RmstTrapzFlex4 = 0#
        Exit Function
    End If
    
    If nSteps < 1 Then nSteps = 1
    
    h = tau / nSteps
    s0 = RmstSurvFlex4(0#, distName, p1, p2, p3, p4)
    
    If IsError(s0) Then
        RmstTrapzFlex4 = s0
        Exit Function
    End If
    
    area = 0#
    
    For i = 1 To nSteps
        t1 = i * h
        s1 = RmstSurvFlex4(t1, distName, p1, p2, p3, p4)
        
        If IsError(s1) Then
            RmstTrapzFlex4 = s1
            Exit Function
        End If
        
        area = area + 0.5 * (CDbl(s0) + CDbl(s1)) * h
        s0 = s1
    Next i
    
    RmstTrapzFlex4 = area
End Function

Private Function RmstOneFlex4(ByVal tau As Double, _
                              ByVal distName As String, _
                              ByVal p1 As Double, _
                              ByVal p2 As Double, _
                              ByVal p3 As Double, _
                              ByVal p4 As Double) As Variant
    Dim n As Long
    Dim a0 As Variant, a1 As Variant
    Dim errAbs As Double
    Dim scaleVal As Double
    
    If tau < 0# Then
        RmstOneFlex4 = CVErr(xlErrNum)
        Exit Function
    End If
    
    If tau = 0# Then
        RmstOneFlex4 = 0#
        Exit Function
    End If
    
    n = RMST_MIN_STEPS
    a0 = RmstTrapzFlex4(tau, distName, p1, p2, p3, p4, n)
    
    If IsError(a0) Then
        RmstOneFlex4 = a0
        Exit Function
    End If
    
    Do While n < RMST_MAX_STEPS
        n = n * 2
        a1 = RmstTrapzFlex4(tau, distName, p1, p2, p3, p4, n)
        
        If IsError(a1) Then
            RmstOneFlex4 = a1
            Exit Function
        End If
        
        errAbs = Abs(CDbl(a1) - CDbl(a0))
        scaleVal = 1# + Abs(CDbl(a1))
        
        If errAbs <= RMST_REL_TOL * scaleVal Then
            RmstOneFlex4 = a1
            Exit Function
        End If
        
        a0 = a1
    Loop
    
    RmstOneFlex4 = a0
End Function

Private Function RmstScalarOrArrayFlex4(ByVal tauInput As Variant, _
                                        ByVal distName As String, _
                                        ByVal p1 As Double, _
                                        ByVal p2 As Double, _
                                        ByVal p3 As Double, _
                                        ByVal p4 As Double) As Variant
    Dim vals As Variant
    Dim out() As Variant
    Dim r As Long, c As Long
    Dim rng As Range
    
    On Error GoTo EH
    
    If IsObject(tauInput) Then
        If TypeName(tauInput) = "Range" Then
            Set rng = tauInput
            
            If rng.CountLarge = 1 Then
                If IsNumeric(rng.Value2) Then
                    RmstScalarOrArrayFlex4 = RmstOneFlex4(CDbl(rng.Value2), distName, p1, p2, p3, p4)
                Else
                    RmstScalarOrArrayFlex4 = CVErr(xlErrValue)
                End If
                Exit Function
            End If
            
            vals = rng.Value2
            ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
            
            For r = 1 To UBound(vals, 1)
                For c = 1 To UBound(vals, 2)
                    If IsNumeric(vals(r, c)) Then
                        out(r, c) = RmstOneFlex4(CDbl(vals(r, c)), distName, p1, p2, p3, p4)
                    Else
                        out(r, c) = CVErr(xlErrValue)
                    End If
                Next c
            Next r
            
            RmstScalarOrArrayFlex4 = out
            Exit Function
        End If
    End If
    
    If IsArray(tauInput) Then
        vals = tauInput
        ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
        
        For r = 1 To UBound(vals, 1)
            For c = 1 To UBound(vals, 2)
                If IsNumeric(vals(r, c)) Then
                    out(r, c) = RmstOneFlex4(CDbl(vals(r, c)), distName, p1, p2, p3, p4)
                Else
                    out(r, c) = CVErr(xlErrValue)
                End If
            Next c
        Next r
        
        RmstScalarOrArrayFlex4 = out
    ElseIf IsNumeric(tauInput) Then
        RmstScalarOrArrayFlex4 = RmstOneFlex4(CDbl(tauInput), distName, p1, p2, p3, p4)
    Else
        RmstScalarOrArrayFlex4 = CVErr(xlErrValue)
    End If
    
    Exit Function
    
EH:
    RmstScalarOrArrayFlex4 = CVErr(xlErrValue)
End Function

'==========================================================
' Mixture 1-2 parameter distributions
'==========================================================

Private Function RmstSurvMixStd(ByVal t As Double, _
                                ByVal distCode As Long, _
                                ByVal thetaInput As Double, _
                                ByVal p1 As Double, _
                                Optional ByVal p2 As Double = 0#) As Variant
    Dim v As Variant
    Dim thetaVal As Double
    
    thetaVal = ClampThetaMix(thetaInput)
    
    If t <= 0# Then
        RmstSurvMixStd = 1#
        Exit Function
    End If
    
    Select Case distCode
        Case DIST_EXP
            v = Smixexp(t, thetaVal, p1, False)
        Case DIST_WEIBULL
            v = Smixweibull(t, thetaVal, p1, p2, False)
        Case DIST_WEIBULLPH
            v = SmixweibullPH(t, thetaVal, p1, p2, False)
        Case DIST_GOMPERTZ
            v = Smixgompertz(t, thetaVal, p1, p2, False)
        Case DIST_LNORM
            v = Smixlnorm(t, thetaVal, p1, p2, False)
        Case DIST_LLOGIS
            v = Smixllogis(t, thetaVal, p1, p2, False)
        Case DIST_GAMMA
            v = Smixgamma(t, thetaVal, p1, p2, False)
        Case Else
            RmstSurvMixStd = CVErr(xlErrValue)
            Exit Function
    End Select
    
    If IsError(v) Then
        RmstSurvMixStd = CVErr(xlErrValue)
    Else
        RmstSurvMixStd = RmstClamp01(CDbl(v))
    End If
End Function

Private Function RmstTrapzMixStd(ByVal tau As Double, _
                                 ByVal distCode As Long, _
                                 ByVal thetaInput As Double, _
                                 ByVal p1 As Double, _
                                 Optional ByVal p2 As Double = 0#, _
                                 Optional ByVal nSteps As Long = RMST_MIN_STEPS) As Variant
    Dim h As Double
    Dim i As Long
    Dim s0 As Variant, s1 As Variant
    Dim area As Double
    Dim t1 As Double
    
    If tau < 0# Then
        RmstTrapzMixStd = CVErr(xlErrNum)
        Exit Function
    End If
    
    If tau = 0# Then
        RmstTrapzMixStd = 0#
        Exit Function
    End If
    
    If nSteps < 1 Then nSteps = 1
    
    h = tau / nSteps
    s0 = RmstSurvMixStd(0#, distCode, thetaInput, p1, p2)
    
    If IsError(s0) Then
        RmstTrapzMixStd = s0
        Exit Function
    End If
    
    area = 0#
    
    For i = 1 To nSteps
        t1 = i * h
        s1 = RmstSurvMixStd(t1, distCode, thetaInput, p1, p2)
        
        If IsError(s1) Then
            RmstTrapzMixStd = s1
            Exit Function
        End If
        
        area = area + 0.5 * (CDbl(s0) + CDbl(s1)) * h
        s0 = s1
    Next i
    
    RmstTrapzMixStd = area
End Function

Private Function RmstOneMixStd(ByVal tau As Double, _
                               ByVal distCode As Long, _
                               ByVal thetaInput As Double, _
                               ByVal p1 As Double, _
                               Optional ByVal p2 As Double = 0#) As Variant
    Dim n As Long
    Dim a0 As Variant, a1 As Variant
    Dim errAbs As Double
    Dim scaleVal As Double
    
    If tau < 0# Then
        RmstOneMixStd = CVErr(xlErrNum)
        Exit Function
    End If
    
    If tau = 0# Then
        RmstOneMixStd = 0#
        Exit Function
    End If
    
    n = RMST_MIN_STEPS
    a0 = RmstTrapzMixStd(tau, distCode, thetaInput, p1, p2, n)
    
    If IsError(a0) Then
        RmstOneMixStd = a0
        Exit Function
    End If
    
    Do While n < RMST_MAX_STEPS
        n = n * 2
        a1 = RmstTrapzMixStd(tau, distCode, thetaInput, p1, p2, n)
        
        If IsError(a1) Then
            RmstOneMixStd = a1
            Exit Function
        End If
        
        errAbs = Abs(CDbl(a1) - CDbl(a0))
        scaleVal = 1# + Abs(CDbl(a1))
        
        If errAbs <= RMST_REL_TOL * scaleVal Then
            RmstOneMixStd = a1
            Exit Function
        End If
        
        a0 = a1
    Loop
    
    RmstOneMixStd = a0
End Function

Private Function RmstScalarOrArrayMixStd(ByVal tauInput As Variant, _
                                         ByVal distCode As Long, _
                                         ByVal thetaInput As Double, _
                                         ByVal p1 As Double, _
                                         Optional ByVal p2 As Double = 0#) As Variant
    Dim vals As Variant
    Dim out() As Variant
    Dim r As Long, c As Long
    Dim rng As Range
    
    On Error GoTo EH
    
    If IsObject(tauInput) Then
        If TypeName(tauInput) = "Range" Then
            Set rng = tauInput
            
            If rng.CountLarge = 1 Then
                If IsNumeric(rng.Value2) Then
                    RmstScalarOrArrayMixStd = RmstOneMixStd(CDbl(rng.Value2), distCode, thetaInput, p1, p2)
                Else
                    RmstScalarOrArrayMixStd = CVErr(xlErrValue)
                End If
                Exit Function
            End If
            
            vals = rng.Value2
            ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
            
            For r = 1 To UBound(vals, 1)
                For c = 1 To UBound(vals, 2)
                    If IsNumeric(vals(r, c)) Then
                        out(r, c) = RmstOneMixStd(CDbl(vals(r, c)), distCode, thetaInput, p1, p2)
                    Else
                        out(r, c) = CVErr(xlErrValue)
                    End If
                Next c
            Next r
            
            RmstScalarOrArrayMixStd = out
            Exit Function
        End If
    End If
    
    If IsArray(tauInput) Then
        vals = tauInput
        ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
        
        For r = 1 To UBound(vals, 1)
            For c = 1 To UBound(vals, 2)
                If IsNumeric(vals(r, c)) Then
                    out(r, c) = RmstOneMixStd(CDbl(vals(r, c)), distCode, thetaInput, p1, p2)
                Else
                    out(r, c) = CVErr(xlErrValue)
                End If
            Next c
        Next r
        
        RmstScalarOrArrayMixStd = out
    ElseIf IsNumeric(tauInput) Then
        RmstScalarOrArrayMixStd = RmstOneMixStd(CDbl(tauInput), distCode, thetaInput, p1, p2)
    Else
        RmstScalarOrArrayMixStd = CVErr(xlErrValue)
    End If
    
    Exit Function
    
EH:
    RmstScalarOrArrayMixStd = CVErr(xlErrValue)
End Function

'==========================================================
' Mixture 3-parameter flexible distributions
'==========================================================

Private Function RmstSurvMixFlex3(ByVal t As Double, _
                                  ByVal distName As String, _
                                  ByVal thetaInput As Double, _
                                  ByVal p1 As Double, _
                                  ByVal p2 As Double, _
                                  ByVal p3 As Double) As Variant
    Dim v As Variant
    Dim thetaVal As Double
    
    thetaVal = ClampThetaMix(thetaInput)
    
    If t <= 0# Then
        RmstSurvMixFlex3 = 1#
        Exit Function
    End If
    
    Select Case LCase$(Trim$(distName))
        Case "gengamma"
            v = Smixgengamma(t, thetaVal, p1, p2, p3, False)
        Case "gengamma_orig"
            v = Smixgengamma_orig(t, thetaVal, p1, p2, p3, False)
        Case Else
            RmstSurvMixFlex3 = CVErr(xlErrValue)
            Exit Function
    End Select
    
    If IsError(v) Then
        RmstSurvMixFlex3 = CVErr(xlErrValue)
    Else
        RmstSurvMixFlex3 = RmstClamp01(CDbl(v))
    End If
End Function

Private Function RmstTrapzMixFlex3(ByVal tau As Double, _
                                   ByVal distName As String, _
                                   ByVal thetaInput As Double, _
                                   ByVal p1 As Double, _
                                   ByVal p2 As Double, _
                                   ByVal p3 As Double, _
                                   Optional ByVal nSteps As Long = RMST_MIN_STEPS) As Variant
    Dim h As Double
    Dim i As Long
    Dim s0 As Variant, s1 As Variant
    Dim area As Double
    Dim t1 As Double
    
    If tau < 0# Then
        RmstTrapzMixFlex3 = CVErr(xlErrNum)
        Exit Function
    End If
    
    If tau = 0# Then
        RmstTrapzMixFlex3 = 0#
        Exit Function
    End If
    
    If nSteps < 1 Then nSteps = 1
    
    h = tau / nSteps
    s0 = RmstSurvMixFlex3(0#, distName, thetaInput, p1, p2, p3)
    
    If IsError(s0) Then
        RmstTrapzMixFlex3 = s0
        Exit Function
    End If
    
    area = 0#
    
    For i = 1 To nSteps
        t1 = i * h
        s1 = RmstSurvMixFlex3(t1, distName, thetaInput, p1, p2, p3)
        
        If IsError(s1) Then
            RmstTrapzMixFlex3 = s1
            Exit Function
        End If
        
        area = area + 0.5 * (CDbl(s0) + CDbl(s1)) * h
        s0 = s1
    Next i
    
    RmstTrapzMixFlex3 = area
End Function

Private Function RmstOneMixFlex3(ByVal tau As Double, _
                                 ByVal distName As String, _
                                 ByVal thetaInput As Double, _
                                 ByVal p1 As Double, _
                                 ByVal p2 As Double, _
                                 ByVal p3 As Double) As Variant
    Dim n As Long
    Dim a0 As Variant, a1 As Variant
    Dim errAbs As Double
    Dim scaleVal As Double
    
    If tau < 0# Then
        RmstOneMixFlex3 = CVErr(xlErrNum)
        Exit Function
    End If
    
    If tau = 0# Then
        RmstOneMixFlex3 = 0#
        Exit Function
    End If
    
    n = RMST_MIN_STEPS
    a0 = RmstTrapzMixFlex3(tau, distName, thetaInput, p1, p2, p3, n)
    
    If IsError(a0) Then
        RmstOneMixFlex3 = a0
        Exit Function
    End If
    
    Do While n < RMST_MAX_STEPS
        n = n * 2
        a1 = RmstTrapzMixFlex3(tau, distName, thetaInput, p1, p2, p3, n)
        
        If IsError(a1) Then
            RmstOneMixFlex3 = a1
            Exit Function
        End If
        
        errAbs = Abs(CDbl(a1) - CDbl(a0))
        scaleVal = 1# + Abs(CDbl(a1))
        
        If errAbs <= RMST_REL_TOL * scaleVal Then
            RmstOneMixFlex3 = a1
            Exit Function
        End If
        
        a0 = a1
    Loop
    
    RmstOneMixFlex3 = a0
End Function

Private Function RmstScalarOrArrayMixFlex3(ByVal tauInput As Variant, _
                                           ByVal distName As String, _
                                           ByVal thetaInput As Double, _
                                           ByVal p1 As Double, _
                                           ByVal p2 As Double, _
                                           ByVal p3 As Double) As Variant
    Dim vals As Variant
    Dim out() As Variant
    Dim r As Long, c As Long
    Dim rng As Range
    
    On Error GoTo EH
    
    If IsObject(tauInput) Then
        If TypeName(tauInput) = "Range" Then
            Set rng = tauInput
            
            If rng.CountLarge = 1 Then
                If IsNumeric(rng.Value2) Then
                    RmstScalarOrArrayMixFlex3 = RmstOneMixFlex3(CDbl(rng.Value2), distName, thetaInput, p1, p2, p3)
                Else
                    RmstScalarOrArrayMixFlex3 = CVErr(xlErrValue)
                End If
                Exit Function
            End If
            
            vals = rng.Value2
            ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
            
            For r = 1 To UBound(vals, 1)
                For c = 1 To UBound(vals, 2)
                    If IsNumeric(vals(r, c)) Then
                        out(r, c) = RmstOneMixFlex3(CDbl(vals(r, c)), distName, thetaInput, p1, p2, p3)
                    Else
                        out(r, c) = CVErr(xlErrValue)
                    End If
                Next c
            Next r
            
            RmstScalarOrArrayMixFlex3 = out
            Exit Function
        End If
    End If
    
    If IsArray(tauInput) Then
        vals = tauInput
        ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
        
        For r = 1 To UBound(vals, 1)
            For c = 1 To UBound(vals, 2)
                If IsNumeric(vals(r, c)) Then
                    out(r, c) = RmstOneMixFlex3(CDbl(vals(r, c)), distName, thetaInput, p1, p2, p3)
                Else
                    out(r, c) = CVErr(xlErrValue)
                End If
            Next c
        Next r
        
        RmstScalarOrArrayMixFlex3 = out
    ElseIf IsNumeric(tauInput) Then
        RmstScalarOrArrayMixFlex3 = RmstOneMixFlex3(CDbl(tauInput), distName, thetaInput, p1, p2, p3)
    Else
        RmstScalarOrArrayMixFlex3 = CVErr(xlErrValue)
    End If
    
    Exit Function
    
EH:
    RmstScalarOrArrayMixFlex3 = CVErr(xlErrValue)
End Function

'==========================================================
' Mixture 4-parameter flexible distributions
'==========================================================

Private Function RmstSurvMixFlex4(ByVal t As Double, _
                                  ByVal distName As String, _
                                  ByVal thetaInput As Double, _
                                  ByVal p1 As Double, _
                                  ByVal p2 As Double, _
                                  ByVal p3 As Double, _
                                  ByVal p4 As Double) As Variant
    Dim v As Variant
    Dim thetaVal As Double
    
    thetaVal = ClampThetaMix(thetaInput)
    
    If t <= 0# Then
        RmstSurvMixFlex4 = 1#
        Exit Function
    End If
    
    Select Case LCase$(Trim$(distName))
        Case "genf"
            v = Smixgenf(t, thetaVal, p1, p2, p3, p4, False)
        Case "genf_orig"
            v = Smixgenf_orig(t, thetaVal, p1, p2, p3, p4, False)
        Case Else
            RmstSurvMixFlex4 = CVErr(xlErrValue)
            Exit Function
    End Select
    
    If IsError(v) Then
        RmstSurvMixFlex4 = CVErr(xlErrValue)
    Else
        RmstSurvMixFlex4 = RmstClamp01(CDbl(v))
    End If
End Function

Private Function RmstTrapzMixFlex4(ByVal tau As Double, _
                                   ByVal distName As String, _
                                   ByVal thetaInput As Double, _
                                   ByVal p1 As Double, _
                                   ByVal p2 As Double, _
                                   ByVal p3 As Double, _
                                   ByVal p4 As Double, _
                                   Optional ByVal nSteps As Long = RMST_MIN_STEPS) As Variant
    Dim h As Double
    Dim i As Long
    Dim s0 As Variant, s1 As Variant
    Dim area As Double
    Dim t1 As Double
    
    If tau < 0# Then
        RmstTrapzMixFlex4 = CVErr(xlErrNum)
        Exit Function
    End If
    
    If tau = 0# Then
        RmstTrapzMixFlex4 = 0#
        Exit Function
    End If
    
    If nSteps < 1 Then nSteps = 1
    
    h = tau / nSteps
    s0 = RmstSurvMixFlex4(0#, distName, thetaInput, p1, p2, p3, p4)
    
    If IsError(s0) Then
        RmstTrapzMixFlex4 = s0
        Exit Function
    End If
    
    area = 0#
    
    For i = 1 To nSteps
        t1 = i * h
        s1 = RmstSurvMixFlex4(t1, distName, thetaInput, p1, p2, p3, p4)
        
        If IsError(s1) Then
            RmstTrapzMixFlex4 = s1
            Exit Function
        End If
        
        area = area + 0.5 * (CDbl(s0) + CDbl(s1)) * h
        s0 = s1
    Next i
    
    RmstTrapzMixFlex4 = area
End Function

Private Function RmstOneMixFlex4(ByVal tau As Double, _
                                 ByVal distName As String, _
                                 ByVal thetaInput As Double, _
                                 ByVal p1 As Double, _
                                 ByVal p2 As Double, _
                                 ByVal p3 As Double, _
                                 ByVal p4 As Double) As Variant
    Dim n As Long
    Dim a0 As Variant, a1 As Variant
    Dim errAbs As Double
    Dim scaleVal As Double
    
    If tau < 0# Then
        RmstOneMixFlex4 = CVErr(xlErrNum)
        Exit Function
    End If
    
    If tau = 0# Then
        RmstOneMixFlex4 = 0#
        Exit Function
    End If
    
    n = RMST_MIN_STEPS
    a0 = RmstTrapzMixFlex4(tau, distName, thetaInput, p1, p2, p3, p4, n)
    
    If IsError(a0) Then
        RmstOneMixFlex4 = a0
        Exit Function
    End If
    
    Do While n < RMST_MAX_STEPS
        n = n * 2
        a1 = RmstTrapzMixFlex4(tau, distName, thetaInput, p1, p2, p3, p4, n)
        
        If IsError(a1) Then
            RmstOneMixFlex4 = a1
            Exit Function
        End If
        
        errAbs = Abs(CDbl(a1) - CDbl(a0))
        scaleVal = 1# + Abs(CDbl(a1))
        
        If errAbs <= RMST_REL_TOL * scaleVal Then
            RmstOneMixFlex4 = a1
            Exit Function
        End If
        
        a0 = a1
    Loop
    
    RmstOneMixFlex4 = a0
End Function

Private Function RmstScalarOrArrayMixFlex4(ByVal tauInput As Variant, _
                                           ByVal distName As String, _
                                           ByVal thetaInput As Double, _
                                           ByVal p1 As Double, _
                                           ByVal p2 As Double, _
                                           ByVal p3 As Double, _
                                           ByVal p4 As Double) As Variant
    Dim vals As Variant
    Dim out() As Variant
    Dim r As Long, c As Long
    Dim rng As Range
    
    On Error GoTo EH
    
    If IsObject(tauInput) Then
        If TypeName(tauInput) = "Range" Then
            Set rng = tauInput
            
            If rng.CountLarge = 1 Then
                If IsNumeric(rng.Value2) Then
                    RmstScalarOrArrayMixFlex4 = RmstOneMixFlex4(CDbl(rng.Value2), distName, thetaInput, p1, p2, p3, p4)
                Else
                    RmstScalarOrArrayMixFlex4 = CVErr(xlErrValue)
                End If
                Exit Function
            End If
            
            vals = rng.Value2
            ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
            
            For r = 1 To UBound(vals, 1)
                For c = 1 To UBound(vals, 2)
                    If IsNumeric(vals(r, c)) Then
                        out(r, c) = RmstOneMixFlex4(CDbl(vals(r, c)), distName, thetaInput, p1, p2, p3, p4)
                    Else
                        out(r, c) = CVErr(xlErrValue)
                    End If
                Next c
            Next r
            
            RmstScalarOrArrayMixFlex4 = out
            Exit Function
        End If
    End If
    
    If IsArray(tauInput) Then
        vals = tauInput
        ReDim out(1 To UBound(vals, 1), 1 To UBound(vals, 2))
        
        For r = 1 To UBound(vals, 1)
            For c = 1 To UBound(vals, 2)
                If IsNumeric(vals(r, c)) Then
                    out(r, c) = RmstOneMixFlex4(CDbl(vals(r, c)), distName, thetaInput, p1, p2, p3, p4)
                Else
                    out(r, c) = CVErr(xlErrValue)
                End If
            Next c
        Next r
        
        RmstScalarOrArrayMixFlex4 = out
    ElseIf IsNumeric(tauInput) Then
        RmstScalarOrArrayMixFlex4 = RmstOneMixFlex4(CDbl(tauInput), distName, thetaInput, p1, p2, p3, p4)
    Else
        RmstScalarOrArrayMixFlex4 = CVErr(xlErrValue)
    End If
    
    Exit Function
    
EH:
    RmstScalarOrArrayMixFlex4 = CVErr(xlErrValue)
End Function

'==========================================================
' Public worksheet functions: standard RMST
'==========================================================

Public Function rmstexp(ByVal tau As Variant, ByVal rate_ As Double) As Variant
    rmstexp = RmstScalarOrArrayStd(tau, DIST_EXP, rate_, 0#)
End Function

Public Function rmstweibull(ByVal tau As Variant, ByVal shape_ As Double, ByVal scale_ As Double) As Variant
    rmstweibull = RmstScalarOrArrayStd(tau, DIST_WEIBULL, shape_, scale_)
End Function

Public Function rmstweibullPH(ByVal tau As Variant, ByVal shape_ As Double, ByVal scale_ As Double) As Variant
    rmstweibullPH = RmstScalarOrArrayStd(tau, DIST_WEIBULLPH, shape_, scale_)
End Function

Public Function rmstgompertz(ByVal tau As Variant, ByVal shape_ As Double, ByVal rate_ As Double) As Variant
    rmstgompertz = RmstScalarOrArrayStd(tau, DIST_GOMPERTZ, shape_, rate_)
End Function

Public Function rmstlnorm(ByVal tau As Variant, ByVal meanlog_ As Double, ByVal sdlog_ As Double) As Variant
    rmstlnorm = RmstScalarOrArrayStd(tau, DIST_LNORM, meanlog_, sdlog_)
End Function

Public Function rmstllogis(ByVal tau As Variant, ByVal shape_ As Double, ByVal scale_ As Double) As Variant
    rmstllogis = RmstScalarOrArrayStd(tau, DIST_LLOGIS, shape_, scale_)
End Function

Public Function rmstgamma(ByVal tau As Variant, ByVal shape_ As Double, ByVal rate_ As Double) As Variant
    rmstgamma = RmstScalarOrArrayStd(tau, DIST_GAMMA, shape_, rate_)
End Function

Public Function rmstgengamma(ByVal tau As Variant, _
                             ByVal muVal As Double, _
                             ByVal sigmaVal As Double, _
                             ByVal qParam As Double) As Variant
    rmstgengamma = RmstScalarOrArrayFlex3(tau, "gengamma", muVal, sigmaVal, qParam)
End Function

Public Function rmstgengamma_orig(ByVal tau As Variant, _
                                  ByVal shapeVal As Double, _
                                  ByVal scaleVal As Double, _
                                  ByVal kVal As Double) As Variant
    rmstgengamma_orig = RmstScalarOrArrayFlex3(tau, "gengamma_orig", shapeVal, scaleVal, kVal)
End Function

Public Function rmstgenf(ByVal tau As Variant, _
                         ByVal muVal As Double, _
                         ByVal sigmaVal As Double, _
                         ByVal qParam As Double, _
                         ByVal pParam As Double) As Variant
    rmstgenf = RmstScalarOrArrayFlex4(tau, "genf", muVal, sigmaVal, qParam, pParam)
End Function

Public Function rmstgenf_orig(ByVal tau As Variant, _
                              ByVal muVal As Double, _
                              ByVal sigmaVal As Double, _
                              ByVal s1Val As Double, _
                              ByVal s2Val As Double) As Variant
    rmstgenf_orig = RmstScalarOrArrayFlex4(tau, "genf_orig", muVal, sigmaVal, s1Val, s2Val)
End Function

'==========================================================
' Public worksheet functions: mixture-cure RMST
'==========================================================

Public Function rmstmixexp(ByVal tau As Variant, _
                           ByVal thetaInput As Double, _
                           ByVal rate_ As Double) As Variant
    rmstmixexp = RmstScalarOrArrayMixStd(tau, DIST_EXP, thetaInput, rate_, 0#)
End Function

Public Function rmstmixweibull(ByVal tau As Variant, _
                               ByVal thetaInput As Double, _
                               ByVal shape_ As Double, _
                               ByVal scale_ As Double) As Variant
    rmstmixweibull = RmstScalarOrArrayMixStd(tau, DIST_WEIBULL, thetaInput, shape_, scale_)
End Function

Public Function rmstmixweibullPH(ByVal tau As Variant, _
                                 ByVal thetaInput As Double, _
                                 ByVal shape_ As Double, _
                                 ByVal scale_ As Double) As Variant
    rmstmixweibullPH = RmstScalarOrArrayMixStd(tau, DIST_WEIBULLPH, thetaInput, shape_, scale_)
End Function

Public Function rmstmixgompertz(ByVal tau As Variant, _
                                ByVal thetaInput As Double, _
                                ByVal shape_ As Double, _
                                ByVal rate_ As Double) As Variant
    rmstmixgompertz = RmstScalarOrArrayMixStd(tau, DIST_GOMPERTZ, thetaInput, shape_, rate_)
End Function

Public Function rmstmixlnorm(ByVal tau As Variant, _
                             ByVal thetaInput As Double, _
                             ByVal meanlog_ As Double, _
                             ByVal sdlog_ As Double) As Variant
    rmstmixlnorm = RmstScalarOrArrayMixStd(tau, DIST_LNORM, thetaInput, meanlog_, sdlog_)
End Function

Public Function rmstmixllogis(ByVal tau As Variant, _
                              ByVal thetaInput As Double, _
                              ByVal shape_ As Double, _
                              ByVal scale_ As Double) As Variant
    rmstmixllogis = RmstScalarOrArrayMixStd(tau, DIST_LLOGIS, thetaInput, shape_, scale_)
End Function

Public Function rmstmixgamma(ByVal tau As Variant, _
                             ByVal thetaInput As Double, _
                             ByVal shape_ As Double, _
                             ByVal rate_ As Double) As Variant
    rmstmixgamma = RmstScalarOrArrayMixStd(tau, DIST_GAMMA, thetaInput, shape_, rate_)
End Function

Public Function rmstmixgengamma(ByVal tau As Variant, _
                                ByVal thetaInput As Double, _
                                ByVal muVal As Double, _
                                ByVal sigmaVal As Double, _
                                ByVal qParam As Double) As Variant
    rmstmixgengamma = RmstScalarOrArrayMixFlex3(tau, "gengamma", thetaInput, muVal, sigmaVal, qParam)
End Function

Public Function rmstmixgengamma_orig(ByVal tau As Variant, _
                                     ByVal thetaInput As Double, _
                                     ByVal shapeVal As Double, _
                                     ByVal scaleVal As Double, _
                                     ByVal kVal As Double) As Variant
    rmstmixgengamma_orig = RmstScalarOrArrayMixFlex3(tau, "gengamma_orig", thetaInput, shapeVal, scaleVal, kVal)
End Function

Public Function rmstmixgenf(ByVal tau As Variant, _
                            ByVal thetaInput As Double, _
                            ByVal muVal As Double, _
                            ByVal sigmaVal As Double, _
                            ByVal qParam As Double, _
                            ByVal pParam As Double) As Variant
    rmstmixgenf = RmstScalarOrArrayMixFlex4(tau, "genf", thetaInput, muVal, sigmaVal, qParam, pParam)
End Function

Public Function rmstmixgenf_orig(ByVal tau As Variant, _
                                 ByVal thetaInput As Double, _
                                 ByVal muVal As Double, _
                                 ByVal sigmaVal As Double, _
                                 ByVal s1Val As Double, _
                                 ByVal s2Val As Double) As Variant
    rmstmixgenf_orig = RmstScalarOrArrayMixFlex4(tau, "genf_orig", thetaInput, muVal, sigmaVal, s1Val, s2Val)
End Function
 
 

