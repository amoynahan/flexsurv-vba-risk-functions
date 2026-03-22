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

'Small numerical tolerance (1e-10) used to handle floating-point precision issues
' Intended for safe comparisons (e.g., equality checks) and to avoid divide-by-zero errors
Private Const KM_EPS As Double = 0.0000000001

' KMTableFast
' Computes Kaplan–Meier table (time, risk, events, censoring, step and cumulative survival)
' Used for fast, unweighted survival estimation directly in Excel
Public Function KMTableFast(ByVal timeRange As Range, ByVal eventRange As Range) As Variant

    On Error GoTo ErrHandler
    KMTableFast = KMFastCore(timeRange, eventRange, False, Nothing)
    Exit Function
ErrHandler:
    KMTableFast = CVErr(xlErrValue)
End Function

' KMTableFastW
' Computes weighted Kaplan–Meier table using observation weights
' Used for analyses requiring weights (e.g., MAIC or adjusted IPD)
Public Function KMTableFastW(ByVal timeRange As Range, ByVal eventRange As Range, ByVal weightRange As Range) As Variant
    On Error GoTo ErrHandler
    KMTableFastW = KMFastCore(timeRange, eventRange, True, weightRange)
    Exit Function
ErrHandler:
    KMTableFastW = CVErr(xlErrValue)
End Function

' KMFastCore
' Core KM engine: aggregates times, counts events/censoring, and calculates survival stepwise
' Centralizes logic for both weighted and unweighted KM for efficiency and consistency
Private Function KMFastCore( _
    ByVal timeRange As Range, _
    ByVal eventRange As Range, _
    ByVal useWeights As Boolean, _
    ByVal weightRange As Range) As Variant

    Dim tData As Variant, eData As Variant, wData As Variant
    Dim n As Long, i As Long, m As Long
    Dim t As Double, e As Double, w As Double
    Dim key As String
    
    Dim dictEvent As Object
    Dim dictCensor As Object
    Dim dictTime As Object
    
    Dim keys As Variant
    Dim times() As Double
    Dim outArr() As Variant
    
    Dim totalRisk As Double
    Dim nRisk As Double
    Dim nEvent As Double
    Dim nCensor As Double
    Dim stepSurv As Double
    Dim kmSurv As Double
    
    If timeRange.Columns.Count <> 1 Or eventRange.Columns.Count <> 1 Then
        KMFastCore = CVErr(xlErrRef)
        Exit Function
    End If
    
    If timeRange.Rows.Count <> eventRange.Rows.Count Then
        KMFastCore = CVErr(xlErrRef)
        Exit Function
    End If
    
    If useWeights Then
        If weightRange Is Nothing Then
            KMFastCore = CVErr(xlErrRef)
            Exit Function
        End If
        If weightRange.Columns.Count <> 1 Then
            KMFastCore = CVErr(xlErrRef)
            Exit Function
        End If
        If weightRange.Rows.Count <> timeRange.Rows.Count Then
            KMFastCore = CVErr(xlErrRef)
            Exit Function
        End If
    End If
    
    tData = RangeTo2DColumn(timeRange)
    eData = RangeTo2DColumn(eventRange)
    If useWeights Then wData = RangeTo2DColumn(weightRange)
    
    n = UBound(tData, 1)
    
    Set dictEvent = CreateObject("Scripting.Dictionary")
    Set dictCensor = CreateObject("Scripting.Dictionary")
    Set dictTime = CreateObject("Scripting.Dictionary")
    
    totalRisk = 0#
    
    For i = 1 To n
        If Not IsError(tData(i, 1)) And Not IsError(eData(i, 1)) Then
            If IsNumeric(tData(i, 1)) And IsNumeric(eData(i, 1)) Then
                t = CDbl(tData(i, 1))
                e = CDbl(eData(i, 1))
                
                If useWeights Then
                    If Not IsError(wData(i, 1)) And IsNumeric(wData(i, 1)) Then
                        w = CDbl(wData(i, 1))
                    Else
                        w = 0#
                    End If
                Else
                    w = 1#
                End If
                
                If t >= 0# And (e = 0# Or e = 1#) And w > 0# Then
                    key = TimeKey(t)
                    
                    If Not dictTime.Exists(key) Then
                        dictTime.Add key, t
                        dictEvent.Add key, 0#
                        dictCensor.Add key, 0#
                    End If
                    
                    If e = 1# Then
                        dictEvent(key) = dictEvent(key) + w
                    Else
                        dictCensor(key) = dictCensor(key) + w
                    End If
                    
                    totalRisk = totalRisk + w
                End If
            End If
        End If
    Next i
    
    If dictTime.Count = 0 Then
        KMFastCore = CVErr(xlErrNA)
        Exit Function
    End If
    
    keys = dictTime.keys
    m = dictTime.Count
    
    ReDim times(1 To m)
    For i = 0 To m - 1
        times(i + 1) = CDbl(dictTime(keys(i)))
    Next i
    
    QuickSortDbl times, 1, m
    
    ReDim outArr(1 To m + 1, 1 To 6)
    
    outArr(1, 1) = "time"
    outArr(1, 2) = "n_risk"
    outArr(1, 3) = "n_event"
    outArr(1, 4) = "n_censor"
    outArr(1, 5) = "step_surv"
    outArr(1, 6) = "km_surv"
    
    kmSurv = 1#
    nRisk = totalRisk
    
    For i = 1 To m
        key = TimeKey(times(i))
        nEvent = CDbl(dictEvent(key))
        nCensor = CDbl(dictCensor(key))
        
        If nEvent > 0# Then
            stepSurv = 1# - nEvent / nRisk
            kmSurv = kmSurv * stepSurv
        Else
            stepSurv = 1#
        End If
        
        outArr(i + 1, 1) = times(i)
        outArr(i + 1, 2) = nRisk
        outArr(i + 1, 3) = nEvent
        outArr(i + 1, 4) = nCensor
        outArr(i + 1, 5) = stepSurv
        outArr(i + 1, 6) = kmSurv
        
        nRisk = nRisk - nEvent - nCensor
    Next i
    
    KMFastCore = outArr
End Function

' TimeKey
' Rounds and formats time values into stable string keys
' Ensures correct grouping of times despite floating-point precision issues
Private Function TimeKey(ByVal t As Double) As String
    TimeKey = Format$(Round(t, 10), "0.0000000000")
End Function

' RangeTo2DColumn
' Converts Excel range into a consistent 2D array format
' Standardizes input handling for single cells and multi-row ranges
Private Function RangeTo2DColumn(ByVal rng As Range) As Variant
    Dim v As Variant
    Dim arr(1 To 1, 1 To 1) As Variant
    
    v = rng.Value2
    
    If rng.Cells.Count = 1 Then
        arr(1, 1) = v
        RangeTo2DColumn = arr
    Else
        RangeTo2DColumn = v
    End If
End Function

' QuickSortDbl
' Sorts array of time values in ascending order
' Required for correct ordering of event times in KM calculations
Private Sub QuickSortDbl(ByRef x() As Double, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As Double, tmp As Double
    
    i = first
    j = last
    pivot = x((first + last) \ 2)
    
    Do While i <= j
        Do While x(i) < pivot
            i = i + 1
        Loop
        Do While x(j) > pivot
            j = j - 1
        Loop
        
        If i <= j Then
            tmp = x(i)
            x(i) = x(j)
            x(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If first < j Then QuickSortDbl x, first, j
    If i < last Then QuickSortDbl x, i, last
End Sub

'STEP PLOTS

' KMStepPlot
' Converts KM output into step-function coordinates (horizontal + vertical segments)
' Used to generate properly formatted KM curves for Excel plotting
Public Function KMStepPlot(ByVal timeRange As Range, ByVal survRange As Range, _
                           Optional ByVal startTime As Double = 0#, _
                           Optional ByVal startSurv As Double = 1#) As Variant
    On Error GoTo ErrHandler
    
    Dim tData As Variant, sData As Variant
    Dim times() As Double, survs() As Double
    Dim outArr() As Variant
    
    Dim n As Long, i As Long, validN As Long
    Dim t As Double, s As Double
    Dim r As Long
    
    If timeRange Is Nothing Or survRange Is Nothing Then
        KMStepPlot = CVErr(xlErrRef)
        Exit Function
    End If
    
    If timeRange.Columns.Count <> 1 Or survRange.Columns.Count <> 1 Then
        KMStepPlot = CVErr(xlErrRef)
        Exit Function
    End If
    
    If timeRange.Rows.Count <> survRange.Rows.Count Then
        KMStepPlot = CVErr(xlErrRef)
        Exit Function
    End If
    
    tData = RangeTo2DColumnStep(timeRange)
    sData = RangeTo2DColumnStep(survRange)
    n = UBound(tData, 1)
    
    ReDim times(1 To n)
    ReDim survs(1 To n)
    
    validN = 0
    For i = 1 To n
        If Not IsError(tData(i, 1)) And Not IsError(sData(i, 1)) Then
            If IsNumeric(tData(i, 1)) And IsNumeric(sData(i, 1)) Then
                t = CDbl(tData(i, 1))
                s = CDbl(sData(i, 1))
                
                If t >= startTime And s >= 0# And s <= 1# Then
                    validN = validN + 1
                    times(validN) = t
                    survs(validN) = s
                End If
            End If
        End If
    Next i
    
    If validN = 0 Then
        KMStepPlot = CVErr(xlErrNA)
        Exit Function
    End If
    
    ReDim Preserve times(1 To validN)
    ReDim Preserve survs(1 To validN)
    
    ' rows = 1 header + 1 start + 2*validN
    ReDim outArr(1 To 2 * validN + 2, 1 To 2)
    
    outArr(1, 1) = "plot_time"
    outArr(1, 2) = "plot_surv"
    
    outArr(2, 1) = startTime
    outArr(2, 2) = startSurv
    
    r = 2
    
    For i = 1 To validN
        ' horizontal segment to current time
        r = r + 1
        outArr(r, 1) = times(i)
        If i = 1 Then
            outArr(r, 2) = startSurv
        Else
            outArr(r, 2) = survs(i - 1)
        End If
        
        ' vertical drop at current time
        r = r + 1
        outArr(r, 1) = times(i)
        outArr(r, 2) = survs(i)
    Next i
    
    KMStepPlot = outArr
    Exit Function

ErrHandler:
    KMStepPlot = CVErr(xlErrValue)
End Function


' RangeTo2DColumnStep
' Converts range to 2D array for step plot functions
' Keeps plotting input handling consistent and isolated
Private Function RangeTo2DColumnStep(ByVal rng As Range) As Variant
    Dim v As Variant
    Dim arr(1 To 1, 1 To 1) As Variant
    
    v = rng.Value2
    
    If rng.Cells.Count = 1 Then
        arr(1, 1) = v
        RangeTo2DColumnStep = arr
    Else
        RangeTo2DColumnStep = v
    End If
End Function


' KMRiskAtIntervalStart
' Computes number at risk at fixed interval start times (e.g., every X units)
' Used to create risk tables aligned with reporting intervals in plots
Public Function KMRiskAtIntervalStart(ByVal timeRange As Range, _
                                      ByVal riskRange As Range, _
                                      ByVal intervalWidth As Double) As Variant
    On Error GoTo ErrHandler
    
    Dim tData As Variant, rData As Variant
    Dim times() As Double, risks() As Double
    Dim outArr() As Variant
    
    Dim n As Long, validN As Long
    Dim i As Long, j As Long, outN As Long
    Dim t As Double, r As Double
    Dim maxTime As Double
    Dim gridTime As Double
    Dim baseRisk As Double
    
    If intervalWidth <= 0# Then
        KMRiskAtIntervalStart = CVErr(xlErrNum)
        Exit Function
    End If
    
    If timeRange.Columns.Count <> 1 Or riskRange.Columns.Count <> 1 Then
        KMRiskAtIntervalStart = CVErr(xlErrRef)
        Exit Function
    End If
    
    If timeRange.Rows.Count <> riskRange.Rows.Count Then
        KMRiskAtIntervalStart = CVErr(xlErrRef)
        Exit Function
    End If
    
    tData = RangeTo2DColumnRiskStart(timeRange)
    rData = RangeTo2DColumnRiskStart(riskRange)
    n = UBound(tData, 1)
    
    ReDim times(1 To n)
    ReDim risks(1 To n)
    
    validN = 0
    maxTime = 0#
    
    For i = 1 To n
        If Not IsError(tData(i, 1)) And Not IsError(rData(i, 1)) Then
            If IsNumeric(tData(i, 1)) And IsNumeric(rData(i, 1)) Then
                t = CDbl(tData(i, 1))
                r = CDbl(rData(i, 1))
                
                If t >= 0# And r >= 0# Then
                    validN = validN + 1
                    times(validN) = t
                    risks(validN) = r
                    If t > maxTime Then maxTime = t
                End If
            End If
        End If
    Next i
    
    If validN = 0 Then
        KMRiskAtIntervalStart = CVErr(xlErrNA)
        Exit Function
    End If
    
    ReDim Preserve times(1 To validN)
    ReDim Preserve risks(1 To validN)
    
    QuickSortTimeRiskStart times, risks, 1, validN
    
    baseRisk = risks(1)
    
    outN = Int(maxTime / intervalWidth) + 1
    
    ReDim outArr(1 To outN + 1, 1 To 2)
    outArr(1, 1) = "time"
    outArr(1, 2) = "n_risk"
    
    For i = 0 To outN - 1
        gridTime = i * intervalWidth
        
        If gridTime = 0# Then
            outArr(i + 2, 1) = 0#
            outArr(i + 2, 2) = baseRisk
        Else
            j = FirstIndexGE(times, gridTime)
            
            outArr(i + 2, 1) = gridTime
            
            If j > 0 Then
                outArr(i + 2, 2) = risks(j)
            Else
                outArr(i + 2, 2) = 0#
            End If
        End If
    Next i
    
    KMRiskAtIntervalStart = outArr
    Exit Function

ErrHandler:
    KMRiskAtIntervalStart = CVErr(xlErrValue)
End Function

' FirstIndexGE
' Finds first index where time >= specified value (binary search)
' Efficiently maps interval grid points to observed KM times
Private Function FirstIndexGE(ByRef times() As Double, ByVal x As Double) As Long
    Dim lo As Long, hi As Long, mid As Long
    Dim ans As Long
    
    lo = LBound(times)
    hi = UBound(times)
    ans = 0
    
    Do While lo <= hi
        mid = (lo + hi) \ 2
        
        If times(mid) >= x Then
            ans = mid
            hi = mid - 1
        Else
            lo = mid + 1
        End If
    Loop
    
    FirstIndexGE = ans
End Function

' RangeTo2DColumnRiskStart
' Converts range to 2D array for risk table calculations
' Ensures consistent structure for interval-based risk extraction
Private Function RangeTo2DColumnRiskStart(ByVal rng As Range) As Variant
    Dim v As Variant
    Dim arr(1 To 1, 1 To 1) As Variant
    
    v = rng.Value2
    
    If rng.Cells.Count = 1 Then
        arr(1, 1) = v
        RangeTo2DColumnRiskStart = arr
    Else
        RangeTo2DColumnRiskStart = v
    End If
End Function

' QuickSortTimeRiskStart
' Sorts times while preserving alignment with corresponding risk values
' Maintains correct pairing of time and risk during ordering
Private Sub QuickSortTimeRiskStart(ByRef times() As Double, _
                                   ByRef risks() As Double, _
                                   ByVal first As Long, _
                                   ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As Double
    Dim tmpT As Double, tmpR As Double
    
    i = first
    j = last
    pivot = times((first + last) \ 2)
    
    Do While i <= j
        Do While times(i) < pivot
            i = i + 1
        Loop
        Do While times(j) > pivot
            j = j - 1
        Loop
        
        If i <= j Then
            tmpT = times(i)
            times(i) = times(j)
            times(j) = tmpT
            
            tmpR = risks(i)
            risks(i) = risks(j)
            risks(j) = tmpR
            
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If first < j Then QuickSortTimeRiskStart times, risks, first, j
    If i < last Then QuickSortTimeRiskStart times, risks, i, last
End Sub

' KMStepPlotTrim
' Builds KM step plot using only event times and extends to max observed time
' Produces cleaner, publication-ready KM curves without redundant points
Public Function KMStepPlotTrim(ByVal timeRange As Range, ByVal survRange As Range, _
                               Optional ByVal startTime As Double = 0#, _
                               Optional ByVal startSurv As Double = 1#) As Variant
    On Error GoTo ErrHandler
    
    Dim tData As Variant, sData As Variant
    Dim times() As Double, survs() As Double
    Dim eventTimes() As Double, eventSurvs() As Double
    Dim outArr() As Variant, finalArr() As Variant
    
    Dim n As Long, i As Long
    Dim validN As Long, eventN As Long
    Dim t As Double, s As Double
    Dim maxTime As Double
    Dim prevSurv As Double
    Dim r As Long
    Dim needExtend As Boolean
    Dim finalRows As Long
    
    If timeRange Is Nothing Or survRange Is Nothing Then
        KMStepPlotTrim = CVErr(xlErrRef)
        Exit Function
    End If
    
    If timeRange.Columns.Count <> 1 Or survRange.Columns.Count <> 1 Then
        KMStepPlotTrim = CVErr(xlErrRef)
        Exit Function
    End If
    
    If timeRange.Rows.Count <> survRange.Rows.Count Then
        KMStepPlotTrim = CVErr(xlErrRef)
        Exit Function
    End If
    
    tData = RangeTo2DColumnStep(timeRange)
    sData = RangeTo2DColumnStep(survRange)
    n = UBound(tData, 1)
    
    ReDim times(1 To n)
    ReDim survs(1 To n)
    
    validN = 0
    maxTime = startTime
    
    For i = 1 To n
        If Not IsError(tData(i, 1)) And Not IsError(sData(i, 1)) Then
            If IsNumeric(tData(i, 1)) And IsNumeric(sData(i, 1)) Then
                t = CDbl(tData(i, 1))
                s = CDbl(sData(i, 1))
                
                If t >= startTime And s >= 0# And s <= 1# Then
                    validN = validN + 1
                    times(validN) = t
                    survs(validN) = s
                    If t > maxTime Then maxTime = t
                End If
            End If
        End If
    Next i
    
    If validN = 0 Then
        KMStepPlotTrim = CVErr(xlErrNA)
        Exit Function
    End If
    
    ReDim Preserve times(1 To validN)
    ReDim Preserve survs(1 To validN)
    
    ReDim eventTimes(1 To validN)
    ReDim eventSurvs(1 To validN)
    
    prevSurv = startSurv
    eventN = 0
    
    For i = 1 To validN
        ' Keep only rows where the KM curve changes
        If survs(i) <> prevSurv Then
            eventN = eventN + 1
            eventTimes(eventN) = times(i)
            eventSurvs(eventN) = survs(i)
            prevSurv = survs(i)
        End If
    Next i
    
    If eventN > 0 Then
        ReDim Preserve eventTimes(1 To eventN)
        ReDim Preserve eventSurvs(1 To eventN)
    End If
    
    needExtend = False
    If maxTime > startTime Then
        If eventN = 0 Then
            needExtend = (maxTime > startTime)
        Else
            needExtend = (maxTime > eventTimes(eventN))
        End If
    End If
    
    ' Allocate maximum possible size, then trim to actual rows used
    ReDim outArr(1 To 2 + 2 * IIf(eventN > 0, eventN, 0) + IIf(needExtend, 1, 0), 1 To 2)
    
    outArr(1, 1) = "plot_time"
    outArr(1, 2) = "plot_surv"
    
    outArr(2, 1) = startTime
    outArr(2, 2) = startSurv
    
    r = 2
    
    For i = 1 To eventN
        ' horizontal segment to event time
        r = r + 1
        outArr(r, 1) = eventTimes(i)
        If i = 1 Then
            outArr(r, 2) = startSurv
        Else
            outArr(r, 2) = eventSurvs(i - 1)
        End If
        
        ' vertical drop at event time
        r = r + 1
        outArr(r, 1) = eventTimes(i)
        outArr(r, 2) = eventSurvs(i)
    Next i
    
    ' Extend horizontally to the largest observed time,
    ' even if the largest time is only a censor
    If needExtend Then
        r = r + 1
        outArr(r, 1) = maxTime
        If eventN = 0 Then
            outArr(r, 2) = startSurv
        Else
            outArr(r, 2) = eventSurvs(eventN)
        End If
    End If
    
    ' Trim to actual used rows so no extra default (0,0) row remains
    finalRows = r
    ReDim finalArr(1 To finalRows, 1 To 2)
    
    For i = 1 To finalRows
        finalArr(i, 1) = outArr(i, 1)
        finalArr(i, 2) = outArr(i, 2)
    Next i
    
    KMStepPlotTrim = finalArr
    Exit Function

ErrHandler:
    KMStepPlotTrim = CVErr(xlErrValue)
End Function


