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

'==========================================================
' Guyot KM reconstruction in VBA
'
' Public functions:
'   GuyotReconstructIPD(...)
'   GuyotReconstructDiagnostics(...)
'
' Inputs:
'   kmTimeRange   digitized KM times
'   kmSurvRange   digitized KM survivals
'   riskTimeRange reported numbers-at-risk times
'   riskNRange    reported numbers at risk
'   totalN        initial sample size
'   totalEvents   optional total number of events
'   kmMaxY        1 if KM survivals are proportions, 100 if percentages
'
' Output 1: pseudo-IPD
'   col1 = time
'   col2 = status (1=event, 0=censor)
'
' Output 2: diagnostics
'   interval
'   risk_time_start
'   risk_time_next
'   km_idx_start
'   km_idx_end
'   n_risk_reported_start
'   n_risk_target_next
'   n_censor_est
'   n_event_est
'   n_risk_reconstructed_end
'
' Notes:
' - Single-arm reconstruction
' - Preprocessing:
'     * remove non-finite rows
'     * normalize survival to [0,1]
'     * remove bad zero-reset tail
'     * sort KM by time
'     * collapse duplicate KM times keeping minimum survival
'     * clamp survival to [0,1]
'     * enforce non-increasing KM
'     * flatten tail after the last true drop
'     * ensure KM starts at (0,1)
'     * sort risk table by time
'     * remove duplicate risk times keeping last
'     * clamp risk table to KM range
'     * clamp risk counts to [0,totalN]
'     * enforce non-increasing risk counts
'     * ensure risk table starts at time 0 with totalN
' - Optional totalEvents constraint rescales interval events
' - Final reconciliation guarantees:
'       number of output IPD rows = totalN
'==========================================================

Private Const GUYOT_TOL As Double = 0.0000001
Private Const GUYOT_EVENT_TOL As Double = 0.000001

Private Type GuyotIntervalResult
    riskTimeStart As Double
    riskTimeNext As Double
    kmIdxStart As Long
    kmIdxEnd As Long
    nStart As Double
    nTargetNext As Double
    nCensor As Long
    nEvent As Long
    nEnd As Double
End Type

'==========================================================
' Public wrappers
'==========================================================

Private Sub SortIPD(ByRef ipd As Variant)
    Dim i As Long, j As Long
    Dim n As Long
    Dim tmpTime As Variant, tmpEvent As Variant

    If IsEmpty(ipd) Then Exit Sub
    If Not IsArray(ipd) Then Exit Sub

    On Error GoTo SafeExit
    n = UBound(ipd, 1)

    For i = 1 To n - 1
        For j = i + 1 To n
            If (CDbl(ipd(i, 1)) > CDbl(ipd(j, 1))) _
               Or ((CDbl(ipd(i, 1)) = CDbl(ipd(j, 1))) And _
                   (CDbl(ipd(i, 2)) > CDbl(ipd(j, 2)))) Then

                tmpTime = ipd(i, 1)
                tmpEvent = ipd(i, 2)

                ipd(i, 1) = ipd(j, 1)
                ipd(i, 2) = ipd(j, 2)

                ipd(j, 1) = tmpTime
                ipd(j, 2) = tmpEvent
            End If
        Next j
    Next i

SafeExit:
End Sub

Public Function GuyotReconstructIPD(ByVal kmTimeRange As Range, _
                                    ByVal kmSurvRange As Range, _
                                    ByVal riskTimeRange As Range, _
                                    ByVal riskNRange As Range, _
                                    ByVal totalN As Long, _
                                    Optional ByVal totalEvents As Variant, _
                                    Optional ByVal kmMaxY As Double = 1#) As Variant
    Dim kmTime() As Double
    Dim kmSurv() As Double
    Dim riskTime() As Double
    Dim riskN() As Double

    kmTime = RangeToVector(kmTimeRange)
    kmSurv = RangeToVector(kmSurvRange)
    riskTime = RangeToVector(riskTimeRange)
    riskN = RangeToVector(riskNRange)

    GuyotReconstructIPD = GuyotReconstructIPD_FromArrays( _
        kmTime, kmSurv, riskTime, riskN, totalN, totalEvents, kmMaxY)
End Function

Public Function GuyotReconstructDiagnostics(ByVal kmTimeRange As Range, _
                                            ByVal kmSurvRange As Range, _
                                            ByVal riskTimeRange As Range, _
                                            ByVal riskNRange As Range, _
                                            ByVal totalN As Long, _
                                            Optional ByVal totalEvents As Variant, _
                                            Optional ByVal kmMaxY As Double = 1#) As Variant
    Dim kmTime() As Double
    Dim kmSurv() As Double
    Dim riskTime() As Double
    Dim riskN() As Double

    kmTime = RangeToVector(kmTimeRange)
    kmSurv = RangeToVector(kmSurvRange)
    riskTime = RangeToVector(riskTimeRange)
    riskN = RangeToVector(riskNRange)

    GuyotReconstructDiagnostics = GuyotDiagnostics_FromArrays( _
        kmTime, kmSurv, riskTime, riskN, totalN, totalEvents, kmMaxY)
End Function

'==========================================================
' Main IPD engine
'==========================================================

Public Function GuyotReconstructIPD_FromArrays(ByRef kmTimeIn() As Double, _
                                               ByRef kmSurvIn() As Double, _
                                               ByRef riskTimeIn() As Double, _
                                               ByRef riskNIn() As Double, _
                                               ByVal totalN As Long, _
                                               Optional ByVal totalEvents As Variant, _
                                               Optional ByVal kmMaxY As Double = 1#) As Variant
    Dim kmTime() As Double
    Dim kmSurv() As Double
    Dim riskTime() As Double
    Dim riskN() As Double
    Dim intervals() As GuyotIntervalResult
    Dim eventCount() As Long
    Dim censorCount() As Long
    Dim nRiskKM() As Double

    Dim outArr() As Variant
    Dim outRow As Long
    Dim nKM As Long
    Dim i As Long, j As Long

    Dim totalEventsRecon As Long
    Dim totalCensorsRecon As Long
    Dim totalCensorsTarget As Long
    Dim remainingSubjects As Long
    Dim lastTime As Double
    Dim lastSurv As Double

    Dim tailCensorsAtLast As Long
    Dim censorsWrittenBeforeLast As Long
    Dim censorsAllowedBeforeLast As Long
    Dim cUse As Long
    Dim ct As Long

    If Not PreprocessInputs(kmTimeIn, kmSurvIn, riskTimeIn, riskNIn, _
                            kmTime, kmSurv, riskTime, riskN, totalN, kmMaxY) Then
        GuyotReconstructIPD_FromArrays = CVErr(xlErrValue)
        Exit Function
    End If

    If Not ReconstructIntervals(kmTime, kmSurv, riskTime, riskN, totalN, _
                                intervals, eventCount, censorCount, nRiskKM) Then
        GuyotReconstructIPD_FromArrays = CVErr(xlErrValue)
        Exit Function
    End If

    If Not IsMissing(totalEvents) Then
        If IsNumeric(totalEvents) Then
            ApplyTotalEventsConstraint eventCount, CLng(totalEvents)
        End If
    End If

    nKM = UBound(kmTime)
    lastTime = kmTime(nKM)
    lastSurv = kmSurv(nKM)

    totalEventsRecon = SumLong(eventCount, LBound(eventCount), UBound(eventCount))
    totalCensorsRecon = SumLong(censorCount, LBound(censorCount), UBound(censorCount))
    remainingSubjects = totalN - totalEventsRecon - totalCensorsRecon

    If remainingSubjects < 0 Then
        ReconcileOverflow totalN, eventCount, censorCount
        totalEventsRecon = SumLong(eventCount, LBound(eventCount), UBound(eventCount))
        totalCensorsRecon = SumLong(censorCount, LBound(censorCount), UBound(censorCount))
        remainingSubjects = totalN - totalEventsRecon - totalCensorsRecon
    End If

    totalCensorsTarget = totalN - totalEventsRecon
    If totalCensorsTarget < 0 Then totalCensorsTarget = 0

    tailCensorsAtLast = EstimateTailCensors(totalN, lastSurv, totalCensorsTarget)
    censorsAllowedBeforeLast = totalCensorsTarget - tailCensorsAtLast
    If censorsAllowedBeforeLast < 0 Then censorsAllowedBeforeLast = 0

    ReDim outArr(1 To totalN, 1 To 2)
    outRow = 0
    censorsWrittenBeforeLast = 0

    For j = 2 To nKM
        Dim d As Long, c As Long
        Dim tPrev As Double, tNow As Double

        d = eventCount(j)
        c = censorCount(j)
        tPrev = kmTime(j - 1)
        tNow = kmTime(j)

        For i = 1 To d
            outRow = outRow + 1
            If outRow > totalN Then Exit For
            outArr(outRow, 1) = tNow
            outArr(outRow, 2) = 1
        Next i

        If outRow > totalN Then Exit For

        cUse = c
        If censorsWrittenBeforeLast + cUse > censorsAllowedBeforeLast Then
            cUse = censorsAllowedBeforeLast - censorsWrittenBeforeLast
            If cUse < 0 Then cUse = 0
        End If

        For ct = 1 To cUse
            outRow = outRow + 1
            If outRow > totalN Then Exit For
            outArr(outRow, 1) = tPrev + (ct / (cUse + 1)) * (tNow - tPrev)
            outArr(outRow, 2) = 0
        Next ct

        censorsWrittenBeforeLast = censorsWrittenBeforeLast + cUse

        If outRow > totalN Then Exit For
    Next j

    For i = 1 To tailCensorsAtLast
        outRow = outRow + 1
        If outRow > totalN Then Exit For
        outArr(outRow, 1) = lastTime
        outArr(outRow, 2) = 0
    Next i

    If outRow < totalN Then
        For i = outRow + 1 To totalN
            outArr(i, 1) = lastTime
            outArr(i, 2) = 0
        Next i
        outRow = totalN
    End If

    outArr = Resize2D(outArr, totalN, 2)
    Call SortIPD(outArr)

    GuyotReconstructIPD_FromArrays = outArr
End Function

'==========================================================
' Diagnostics engine
'==========================================================

Public Function GuyotDiagnostics_FromArrays(ByRef kmTimeIn() As Double, _
                                            ByRef kmSurvIn() As Double, _
                                            ByRef riskTimeIn() As Double, _
                                            ByRef riskNIn() As Double, _
                                            ByVal totalN As Long, _
                                            Optional ByVal totalEvents As Variant, _
                                            Optional ByVal kmMaxY As Double = 1#) As Variant
    Dim kmTime() As Double
    Dim kmSurv() As Double
    Dim riskTime() As Double
    Dim riskN() As Double
    Dim intervals() As GuyotIntervalResult
    Dim eventCount() As Long
    Dim censorCount() As Long
    Dim nRiskKM() As Double
    Dim outArr() As Variant
    Dim i As Long, nIntervals As Long
    Dim totalEventsRecon As Long
    Dim totalCensorsRecon As Long
    Dim remainingSubjects As Long
    Dim tailCensorsAtLast As Long

    If Not PreprocessInputs(kmTimeIn, kmSurvIn, riskTimeIn, riskNIn, _
                            kmTime, kmSurv, riskTime, riskN, totalN, kmMaxY) Then
        GuyotDiagnostics_FromArrays = CVErr(xlErrValue)
        Exit Function
    End If

    If Not ReconstructIntervals(kmTime, kmSurv, riskTime, riskN, totalN, _
                                intervals, eventCount, censorCount, nRiskKM) Then
        GuyotDiagnostics_FromArrays = CVErr(xlErrValue)
        Exit Function
    End If

    If Not IsMissing(totalEvents) Then
        If IsNumeric(totalEvents) Then
            ApplyTotalEventsConstraint eventCount, CLng(totalEvents)
        End If
    End If

    totalEventsRecon = SumLong(eventCount, LBound(eventCount), UBound(eventCount))
    totalCensorsRecon = SumLong(censorCount, LBound(censorCount), UBound(censorCount))
    remainingSubjects = totalN - totalEventsRecon - totalCensorsRecon
    tailCensorsAtLast = EstimateTailCensors(totalN, kmSurv(UBound(kmSurv)), totalN - totalEventsRecon)

    nIntervals = UBound(intervals)

    ReDim outArr(1 To nIntervals + 2, 1 To 11)

    outArr(1, 1) = "interval"
    outArr(1, 2) = "risk_time_start"
    outArr(1, 3) = "risk_time_next"
    outArr(1, 4) = "km_idx_start"
    outArr(1, 5) = "km_idx_end"
    outArr(1, 6) = "n_risk_reported_start"
    outArr(1, 7) = "n_risk_target_next"
    outArr(1, 8) = "n_censor_est"
    outArr(1, 9) = "n_event_est"
    outArr(1, 10) = "n_risk_reconstructed_end"
    outArr(1, 11) = "tail_censors_at_last"

    For i = 1 To nIntervals
        intervals(i).nEvent = SumLong(eventCount, intervals(i).kmIdxStart, intervals(i).kmIdxEnd)
        intervals(i).nEnd = ComputeEndRisk(intervals(i).nStart, intervals(i).nEvent, intervals(i).nCensor)

        outArr(i + 1, 1) = i
        outArr(i + 1, 2) = intervals(i).riskTimeStart
        outArr(i + 1, 3) = intervals(i).riskTimeNext
        outArr(i + 1, 4) = intervals(i).kmIdxStart
        outArr(i + 1, 5) = intervals(i).kmIdxEnd
        outArr(i + 1, 6) = intervals(i).nStart
        outArr(i + 1, 7) = intervals(i).nTargetNext
        outArr(i + 1, 8) = intervals(i).nCensor
        outArr(i + 1, 9) = intervals(i).nEvent
        outArr(i + 1, 10) = intervals(i).nEnd
        outArr(i + 1, 11) = ""
    Next i

    outArr(nIntervals + 2, 1) = "TOTAL"
    outArr(nIntervals + 2, 2) = ""
    outArr(nIntervals + 2, 3) = ""
    outArr(nIntervals + 2, 4) = ""
    outArr(nIntervals + 2, 5) = ""
    outArr(nIntervals + 2, 6) = totalN
    outArr(nIntervals + 2, 7) = ""
    outArr(nIntervals + 2, 8) = totalCensorsRecon
    outArr(nIntervals + 2, 9) = totalEventsRecon
    outArr(nIntervals + 2, 10) = remainingSubjects
    outArr(nIntervals + 2, 11) = tailCensorsAtLast

    GuyotDiagnostics_FromArrays = outArr
End Function

'==========================================================
' Core reconstruction
'==========================================================

Private Function ReconstructIntervals(ByRef kmTime() As Double, _
                                      ByRef kmSurv() As Double, _
                                      ByRef riskTime() As Double, _
                                      ByRef riskN() As Double, _
                                      ByVal totalN As Long, _
                                      ByRef intervals() As GuyotIntervalResult, _
                                      ByRef eventCount() As Long, _
                                      ByRef censorCount() As Long, _
                                      ByRef nRiskKM() As Double) As Boolean
    Dim nKM As Long, nRisk As Long
    Dim intervalStartIdx() As Long, intervalEndIdx() As Long
    Dim i As Long
    Dim idxStart As Long, idxEnd As Long
    Dim nStart As Double, targetNextRisk As Double
    Dim guessCensor As Long, bestCensor As Long
    Dim bestDiff As Double, diffVal As Double

    On Error GoTo EH

    nKM = UBound(kmTime)
    nRisk = UBound(riskTime)

    ReDim intervalStartIdx(1 To nRisk)
    ReDim intervalEndIdx(1 To nRisk)
    ReDim intervals(1 To nRisk)
    ReDim eventCount(1 To nKM)
    ReDim censorCount(1 To nKM)
    ReDim nRiskKM(1 To nKM)

    For i = 1 To nRisk
        intervalStartIdx(i) = FindFirstKMIndexAtOrAfter(kmTime, riskTime(i))
        If intervalStartIdx(i) = 0 Then intervalStartIdx(i) = nKM

        If i < nRisk Then
            intervalEndIdx(i) = FindLastKMIndexBefore(kmTime, riskTime(i + 1))
            If intervalEndIdx(i) < intervalStartIdx(i) Then intervalEndIdx(i) = intervalStartIdx(i)
        Else
            intervalEndIdx(i) = nKM
        End If
    Next i

    nRiskKM(intervalStartIdx(1)) = totalN

    For i = 1 To nRisk
        idxStart = intervalStartIdx(i)
        idxEnd = intervalEndIdx(i)
        nStart = riskN(i)

        If i < nRisk Then
            targetNextRisk = riskN(i + 1)
        Else
            targetNextRisk = nStart
        End If

        bestDiff = 1E+99
        bestCensor = 0

        For guessCensor = 0 To CLng(Application.WorksheetFunction.Max(0, nStart))
            diffVal = IntervalRiskDifference(kmTime, kmSurv, idxStart, idxEnd, _
                                             nStart, guessCensor, targetNextRisk, _
                                             eventCount, censorCount, nRiskKM)

            If Abs(diffVal) < bestDiff Then
                bestDiff = Abs(diffVal)
                bestCensor = guessCensor
            End If

            If bestDiff = 0# Then Exit For
        Next guessCensor

        FillInterval kmTime, kmSurv, idxStart, idxEnd, nStart, bestCensor, _
                     eventCount, censorCount, nRiskKM

        intervals(i).riskTimeStart = riskTime(i)
        If i < nRisk Then
            intervals(i).riskTimeNext = riskTime(i + 1)
        Else
            intervals(i).riskTimeNext = kmTime(nKM)
        End If
        intervals(i).kmIdxStart = idxStart
        intervals(i).kmIdxEnd = idxEnd
        intervals(i).nStart = nStart
        intervals(i).nTargetNext = targetNextRisk
        intervals(i).nCensor = bestCensor
        intervals(i).nEvent = SumLong(eventCount, idxStart, idxEnd)
        intervals(i).nEnd = ComputeEndRisk(nStart, intervals(i).nEvent, bestCensor)

        If i < nRisk Then
            nRiskKM(intervalStartIdx(i + 1)) = intervals(i).nEnd
        End If
    Next i

    ReconstructIntervals = True
    Exit Function

EH:
    ReconstructIntervals = False
End Function

Private Function IntervalRiskDifference(ByRef kmTime() As Double, _
                                        ByRef kmSurv() As Double, _
                                        ByVal idxStart As Long, _
                                        ByVal idxEnd As Long, _
                                        ByVal nStart As Double, _
                                        ByVal totalCensor As Long, _
                                        ByVal targetNextRisk As Double, _
                                        ByRef eventCount() As Long, _
                                        ByRef censorCount() As Long, _
                                        ByRef nRiskKM() As Double) As Double
    Dim tmpEvent() As Long, tmpCensor() As Long
    Dim tmpRisk() As Double
    Dim j As Long

    ReDim tmpEvent(LBound(eventCount) To UBound(eventCount))
    ReDim tmpCensor(LBound(censorCount) To UBound(censorCount))
    ReDim tmpRisk(LBound(nRiskKM) To UBound(nRiskKM))

    For j = LBound(eventCount) To UBound(eventCount)
        tmpEvent(j) = eventCount(j)
        tmpCensor(j) = censorCount(j)
        tmpRisk(j) = nRiskKM(j)
    Next j

    FillInterval kmTime, kmSurv, idxStart, idxEnd, nStart, totalCensor, _
                 tmpEvent, tmpCensor, tmpRisk

    IntervalRiskDifference = ComputeEndRisk(nStart, SumLong(tmpEvent, idxStart, idxEnd), totalCensor) - targetNextRisk
End Function

Private Sub FillInterval(ByRef kmTime() As Double, _
                         ByRef kmSurv() As Double, _
                         ByVal idxStart As Long, _
                         ByVal idxEnd As Long, _
                         ByVal nStart As Double, _
                         ByVal totalCensor As Long, _
                         ByRef eventCount() As Long, _
                         ByRef censorCount() As Long, _
                         ByRef nRiskKM() As Double)
    Dim j As Long
    Dim cPlaced() As Long
    Dim atRisk As Double
    Dim d As Long
    Dim survPrev As Double, survNow As Double
    Dim thisC As Long
    Dim rawDrop As Double

    If idxEnd < idxStart Then Exit Sub

    BuildEvenCensorPlacement idxStart, idxEnd, totalCensor, cPlaced

    atRisk = nStart

    For j = idxStart To idxEnd
        nRiskKM(j) = atRisk

        thisC = cPlaced(j)
        censorCount(j) = thisC

        If j = 1 Then
            survPrev = 1#
        Else
            survPrev = kmSurv(j - 1)
        End If
        survNow = kmSurv(j)

        If survPrev <= 0# Or atRisk <= 0# Then
            d = 0
        Else
            If Abs(survPrev - survNow) <= GUYOT_EVENT_TOL Then
                d = 0
            Else
                rawDrop = atRisk * (1# - survNow / survPrev)
                If rawDrop < 0# Then rawDrop = 0#
                d = Round(rawDrop, 0)
            End If

            If d < 0 Then d = 0
            If d > atRisk - thisC Then d = CLng(atRisk - thisC)
            If d < 0 Then d = 0
        End If

        eventCount(j) = d
        atRisk = atRisk - d - thisC
        If atRisk < 0# Then atRisk = 0#
    Next j
End Sub

'==========================================================
' Spread censors across interval instead of front-loading
'==========================================================

Private Sub BuildEvenCensorPlacement(ByVal idxStart As Long, _
                                     ByVal idxEnd As Long, _
                                     ByVal totalCensor As Long, _
                                     ByRef cPlaced() As Long)
    Dim nSeg As Long
    Dim baseC As Long, remC As Long
    Dim j As Long
    Dim k As Long
    Dim pos As Long

    nSeg = idxEnd - idxStart + 1
    ReDim cPlaced(idxStart To idxEnd)

    If nSeg < 1 Or totalCensor <= 0 Then Exit Sub

    baseC = totalCensor \ nSeg
    remC = totalCensor Mod nSeg

    For j = idxStart To idxEnd
        cPlaced(j) = baseC
    Next j

    If remC = 0 Then Exit Sub

    For k = 1 To remC
        pos = idxStart + CLng(Fix((k * (nSeg + 1)) / (remC + 1))) - 1
        If pos < idxStart Then pos = idxStart
        If pos > idxEnd Then pos = idxEnd
        cPlaced(pos) = cPlaced(pos) + 1
    Next k
End Sub

Private Sub ApplyTotalEventsConstraint(ByRef eventCount() As Long, ByVal targetEvents As Long)
    Dim currentEvents As Long
    Dim scaleFactor As Double
    Dim j As Long
    Dim diff As Long

    If targetEvents < 0 Then Exit Sub

    currentEvents = SumLong(eventCount, LBound(eventCount), UBound(eventCount))
    If currentEvents <= 0 Then Exit Sub

    scaleFactor = targetEvents / currentEvents

    For j = LBound(eventCount) To UBound(eventCount)
        eventCount(j) = Round(eventCount(j) * scaleFactor, 0)
        If eventCount(j) < 0 Then eventCount(j) = 0
    Next j

    diff = targetEvents - SumLong(eventCount, LBound(eventCount), UBound(eventCount))

    j = UBound(eventCount)
    Do While diff <> 0 And j >= LBound(eventCount)
        If diff > 0 Then
            eventCount(j) = eventCount(j) + 1
            diff = diff - 1
        ElseIf diff < 0 And eventCount(j) > 0 Then
            eventCount(j) = eventCount(j) - 1
            diff = diff + 1
        End If

        j = j - 1
        If j < LBound(eventCount) And diff <> 0 Then j = UBound(eventCount)
    Loop
End Sub

Private Sub ReconcileOverflow(ByVal totalN As Long, ByRef eventCount() As Long, ByRef censorCount() As Long)
    Dim totalAssigned As Long
    Dim overflow As Long
    Dim j As Long

    totalAssigned = SumLong(eventCount, LBound(eventCount), UBound(eventCount)) + _
                    SumLong(censorCount, LBound(censorCount), UBound(censorCount))

    overflow = totalAssigned - totalN
    If overflow <= 0 Then Exit Sub

    For j = UBound(censorCount) To LBound(censorCount) Step -1
        Do While overflow > 0 And censorCount(j) > 0
            censorCount(j) = censorCount(j) - 1
            overflow = overflow - 1
        Loop
        If overflow = 0 Then Exit Sub
    Next j

    For j = UBound(eventCount) To LBound(eventCount) Step -1
        Do While overflow > 0 And eventCount(j) > 0
            eventCount(j) = eventCount(j) - 1
            overflow = overflow - 1
        Loop
        If overflow = 0 Then Exit Sub
    Next j
End Sub

Private Function EstimateTailCensors(ByVal totalN As Long, _
                                     ByVal lastSurv As Double, _
                                     ByVal totalCensorsTarget As Long) As Long
    Dim n As Long
    Dim s As Double

    If totalN <= 0 Then
        EstimateTailCensors = 0
        Exit Function
    End If

    s = lastSurv
    If s <= 0# Then
        EstimateTailCensors = 0
        Exit Function
    End If

    If s > 1# Then s = 1#

    n = CLng(Application.WorksheetFunction.Round(totalN * s, 0))

    If n < 0 Then n = 0
    If n > totalCensorsTarget Then n = totalCensorsTarget

    EstimateTailCensors = n
End Function

'==========================================================
' Preprocessing
'==========================================================

Private Function PreprocessInputs(ByRef kmTimeIn() As Double, _
                                  ByRef kmSurvIn() As Double, _
                                  ByRef riskTimeIn() As Double, _
                                  ByRef riskNIn() As Double, _
                                  ByRef kmTimeOut() As Double, _
                                  ByRef kmSurvOut() As Double, _
                                  ByRef riskTimeOut() As Double, _
                                  ByRef riskNOut() As Double, _
                                  ByVal totalN As Long, _
                                  Optional ByVal kmMaxY As Double = 1#) As Boolean
    Dim kmTime() As Double, kmSurv() As Double
    Dim riskTime() As Double, riskN() As Double

    On Error GoTo EH

    If totalN <= 0 Then GoTo EH
    If UBound(kmTimeIn) <> UBound(kmSurvIn) Then GoTo EH
    If UBound(riskTimeIn) <> UBound(riskNIn) Then GoTo EH
    If UBound(kmTimeIn) < 2 Then GoTo EH
    If UBound(riskTimeIn) < 1 Then GoTo EH

    kmTime = kmTimeIn
    kmSurv = kmSurvIn
    riskTime = riskTimeIn
    riskN = riskNIn

    RemoveNonFinitePairs kmTime, kmSurv
    If UBoundSafe(kmTime) < 2 Then GoTo EH

    NormalizeKMScale kmSurv, kmMaxY
    RemoveNegativeKMTimeRows kmTime, kmSurv
    RemoveBadZeroResetTail kmTime, kmSurv
    If UBoundSafe(kmTime) < 2 Then GoTo EH

    SortPairsByFirst kmTime, kmSurv
    CollapseDuplicateKMKeepMinSurv kmTime, kmSurv
    ClampVector kmSurv, 0#, 1#
    EnforceKMMonotone kmSurv
    FlattenKMTail kmTime, kmSurv
    EnsureKMStartsAtZero kmTime, kmSurv

    If UBoundSafe(kmTime) < 2 Then GoTo EH
    If Not ValidateKMInput(kmTime, kmSurv) Then GoTo EH

    RemoveNonFinitePairs riskTime, riskN
    If UBoundSafe(riskTime) < 1 Then GoTo EH

    SortPairsByFirst riskTime, riskN
    RemoveDuplicateRiskTimesKeepLast riskTime, riskN
    ClampRiskTimesToKMRange riskTime, kmTime
    ClampRiskCounts riskN, totalN
    EnforceRiskNonIncreasing riskN
    EnsureRiskStartsAtZero riskTime, riskN, totalN

    If UBoundSafe(riskTime) < 1 Then GoTo EH
    If Not ValidateRiskInput(riskTime, riskN, kmTime, totalN) Then GoTo EH

    kmTimeOut = kmTime
    kmSurvOut = kmSurv
    riskTimeOut = riskTime
    riskNOut = riskN

    PreprocessInputs = True
    Exit Function

EH:
    PreprocessInputs = False
End Function

Private Sub NormalizeKMScale(ByRef kmSurv() As Double, ByVal kmMaxY As Double)
    Dim i As Long

    If Abs(kmMaxY - 100#) <= GUYOT_TOL Then
        For i = 1 To UBound(kmSurv)
            kmSurv(i) = kmSurv(i) / 100#
        Next i
    ElseIf Abs(kmMaxY - 1#) <= GUYOT_TOL Then
        ' do nothing
    Else
        If MaxDouble(kmSurv) > 1.0001 Then
            For i = 1 To UBound(kmSurv)
                kmSurv(i) = kmSurv(i) / 100#
            Next i
        End If
    End If
End Sub

Private Sub RemoveNegativeKMTimeRows(ByRef kmTime() As Double, ByRef kmSurv() As Double)
    Dim i As Long, m As Long
    Dim t() As Double, s() As Double
    Dim n As Long

    n = UBound(kmTime)
    ReDim t(1 To n)
    ReDim s(1 To n)

    m = 0
    For i = 1 To n
        If kmTime(i) >= 0# Then
            m = m + 1
            t(m) = kmTime(i)
            s(m) = kmSurv(i)
        End If
    Next i

    If m = 0 Then
        ReDim t(1 To 1)
        ReDim s(1 To 1)
        t(1) = 0#
        s(1) = 1#
    Else
        ReDim Preserve t(1 To m)
        ReDim Preserve s(1 To m)
    End If

    kmTime = t
    kmSurv = s
End Sub

Private Sub RemoveBadZeroResetTail(ByRef kmTime() As Double, ByRef kmSurv() As Double)
    Dim i As Long
    Dim resetPos As Long
    Dim maxSeen As Double
    Dim t() As Double, s() As Double
    Dim m As Long
    Dim n As Long

    n = UBound(kmTime)
    resetPos = 0
    maxSeen = kmTime(1)

    For i = 2 To n
        If kmTime(i) > maxSeen Then maxSeen = kmTime(i)
        If kmTime(i) <= GUYOT_TOL And maxSeen > GUYOT_TOL Then
            resetPos = i
            Exit For
        End If
    Next i

    If resetPos = 0 Then Exit Sub

    ReDim t(1 To n)
    ReDim s(1 To n)
    m = 0

    For i = 1 To resetPos - 1
        m = m + 1
        t(m) = kmTime(i)
        s(m) = kmSurv(i)
    Next i

    For i = resetPos + 1 To n
        If kmTime(i) > maxSeen + GUYOT_TOL Then
            m = m + 1
            t(m) = kmTime(i)
            s(m) = kmSurv(i)
        End If
    Next i

    If m >= 1 Then
        ReDim Preserve t(1 To m)
        ReDim Preserve s(1 To m)
        kmTime = t
        kmSurv = s
    End If
End Sub

Private Sub CollapseDuplicateKMKeepMinSurv(ByRef kmTime() As Double, ByRef kmSurv() As Double)
    Dim n As Long, i As Long, m As Long
    Dim t() As Double, s() As Double

    n = UBound(kmTime)
    ReDim t(1 To n)
    ReDim s(1 To n)

    m = 0
    For i = 1 To n
        If m = 0 Then
            m = 1
            t(m) = kmTime(i)
            s(m) = kmSurv(i)
        ElseIf Abs(kmTime(i) - t(m)) <= GUYOT_TOL Then
            If kmSurv(i) < s(m) Then s(m) = kmSurv(i)
        Else
            m = m + 1
            t(m) = kmTime(i)
            s(m) = kmSurv(i)
        End If
    Next i

    ReDim Preserve t(1 To m)
    ReDim Preserve s(1 To m)

    kmTime = t
    kmSurv = s
End Sub

Private Sub FlattenKMTail(ByRef kmTime() As Double, ByRef kmSurv() As Double)
    Dim n As Long
    Dim i As Long
    Dim minSurv As Double
    Dim firstMinIdx As Long

    n = UBound(kmSurv)
    If n < 2 Then Exit Sub

    minSurv = kmSurv(1)
    For i = 2 To n
        If kmSurv(i) < minSurv Then
            minSurv = kmSurv(i)
        End If
    Next i

    firstMinIdx = 1
    For i = 1 To n
        If Abs(kmSurv(i) - minSurv) <= GUYOT_EVENT_TOL Then
            firstMinIdx = i
            Exit For
        End If
    Next i

    For i = firstMinIdx To n
        kmSurv(i) = minSurv
    Next i
End Sub

Private Sub RemoveDuplicateRiskTimesKeepLast(ByRef riskTime() As Double, ByRef riskN() As Double)
    Dim n As Long, i As Long, m As Long
    Dim t() As Double, r() As Double

    n = UBound(riskTime)
    ReDim t(1 To n)
    ReDim r(1 To n)

    m = 0
    For i = 1 To n
        If m = 0 Then
            m = 1
            t(m) = riskTime(i)
            r(m) = riskN(i)
        ElseIf Abs(riskTime(i) - t(m)) <= GUYOT_TOL Then
            r(m) = riskN(i)
        Else
            m = m + 1
            t(m) = riskTime(i)
            r(m) = riskN(i)
        End If
    Next i

    ReDim Preserve t(1 To m)
    ReDim Preserve r(1 To m)

    riskTime = t
    riskN = r
End Sub

Private Sub ClampVector(ByRef x() As Double, ByVal xmin As Double, ByVal xmax As Double)
    Dim i As Long
    For i = 1 To UBound(x)
        If x(i) < xmin Then x(i) = xmin
        If x(i) > xmax Then x(i) = xmax
    Next i
End Sub

Private Sub ClampRiskCounts(ByRef x() As Double, ByVal totalN As Long)
    Dim i As Long
    For i = 1 To UBound(x)
        If x(i) < 0# Then x(i) = 0#
        If x(i) > totalN Then x(i) = totalN
        x(i) = Round(x(i), 0)
    Next i
End Sub

Private Sub EnforceKMMonotone(ByRef kmSurv() As Double)
    Dim i As Long
    For i = 2 To UBound(kmSurv)
        If kmSurv(i) > kmSurv(i - 1) Then
            kmSurv(i) = kmSurv(i - 1)
        End If
    Next i
End Sub

Private Sub EnforceRiskNonIncreasing(ByRef riskN() As Double)
    Dim i As Long
    For i = 2 To UBound(riskN)
        If riskN(i) > riskN(i - 1) Then
            riskN(i) = riskN(i - 1)
        End If
    Next i
End Sub

Private Sub EnsureKMStartsAtZero(ByRef kmTime() As Double, ByRef kmSurv() As Double)
    Dim n As Long
    Dim t() As Double, s() As Double
    Dim i As Long

    If Abs(kmTime(1)) <= GUYOT_TOL Then
        kmTime(1) = 0#
        If kmSurv(1) < 1# Then kmSurv(1) = 1#
        Exit Sub
    End If

    n = UBound(kmTime)
    ReDim t(1 To n + 1)
    ReDim s(1 To n + 1)

    t(1) = 0#
    s(1) = 1#

    For i = 1 To n
        t(i + 1) = kmTime(i)
        s(i + 1) = kmSurv(i)
    Next i

    kmTime = t
    kmSurv = s
End Sub

Private Sub EnsureRiskStartsAtZero(ByRef riskTime() As Double, ByRef riskN() As Double, ByVal totalN As Long)
    Dim n As Long
    Dim t() As Double, r() As Double
    Dim i As Long

    If Abs(riskTime(1)) <= GUYOT_TOL Then
        riskTime(1) = 0#
        riskN(1) = totalN
        Exit Sub
    End If

    n = UBound(riskTime)
    ReDim t(1 To n + 1)
    ReDim r(1 To n + 1)

    t(1) = 0#
    r(1) = totalN

    For i = 1 To n
        t(i + 1) = riskTime(i)
        r(i + 1) = riskN(i)
    Next i

    riskTime = t
    riskN = r
End Sub

Private Sub ClampRiskTimesToKMRange(ByRef riskTime() As Double, ByRef kmTime() As Double)
    Dim i As Long
    Dim tmin As Double, tmax As Double

    tmin = 0#
    tmax = kmTime(UBound(kmTime))

    For i = 1 To UBound(riskTime)
        If riskTime(i) < tmin Then riskTime(i) = tmin
        If riskTime(i) > tmax Then riskTime(i) = tmax
    Next i
End Sub

Private Function ValidateKMInput(ByRef kmTime() As Double, ByRef kmSurv() As Double) As Boolean
    Dim i As Long

    ValidateKMInput = False

    If UBoundSafe(kmTime) <> UBoundSafe(kmSurv) Then Exit Function
    If UBoundSafe(kmTime) < 2 Then Exit Function

    For i = 1 To UBound(kmTime)
        If kmTime(i) < -GUYOT_TOL Then Exit Function
        If kmSurv(i) < -GUYOT_TOL Or kmSurv(i) > 1# + GUYOT_TOL Then Exit Function
    Next i

    For i = 2 To UBound(kmTime)
        If kmTime(i) < kmTime(i - 1) - GUYOT_TOL Then Exit Function
        If kmSurv(i) > kmSurv(i - 1) + GUYOT_TOL Then Exit Function
    Next i

    ValidateKMInput = True
End Function

Private Function ValidateRiskInput(ByRef riskTime() As Double, _
                                   ByRef riskN() As Double, _
                                   ByRef kmTime() As Double, _
                                   ByVal totalN As Long) As Boolean
    Dim i As Long

    ValidateRiskInput = False

    If UBoundSafe(riskTime) <> UBoundSafe(riskN) Then Exit Function
    If UBoundSafe(riskTime) < 1 Then Exit Function

    For i = 1 To UBound(riskTime)
        If riskTime(i) < -GUYOT_TOL Then Exit Function
        If riskTime(i) > kmTime(UBound(kmTime)) + GUYOT_TOL Then Exit Function
        If riskN(i) < -GUYOT_TOL Then Exit Function
        If riskN(i) > totalN + GUYOT_TOL Then Exit Function
    Next i

    For i = 2 To UBound(riskTime)
        If riskTime(i) < riskTime(i - 1) - GUYOT_TOL Then Exit Function
        If riskN(i) > riskN(i - 1) + GUYOT_TOL Then Exit Function
    Next i

    ValidateRiskInput = True
End Function

Private Sub SortPairsByFirst(ByRef x() As Double, ByRef y() As Double)
    QuickSortPairs x, y, LBound(x), UBound(x)
End Sub

Private Sub QuickSortPairs(ByRef x() As Double, ByRef y() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As Double
    Dim tx As Double, ty As Double

    i = lo
    j = hi
    pivot = x((lo + hi) \ 2)

    Do While i <= j
        Do While x(i) < pivot
            i = i + 1
        Loop
        Do While x(j) > pivot
            j = j - 1
        Loop

        If i <= j Then
            tx = x(i): x(i) = x(j): x(j) = tx
            ty = y(i): y(i) = y(j): y(j) = ty
            i = i + 1
            j = j - 1
        End If
    Loop

    If lo < j Then QuickSortPairs x, y, lo, j
    If i < hi Then QuickSortPairs x, y, i, hi
End Sub

'==========================================================
' Helpers
'==========================================================

Private Function ComputeEndRisk(ByVal nStart As Double, ByVal nEvent As Long, ByVal nCensor As Long) As Double
    ComputeEndRisk = nStart - nEvent - nCensor
    If ComputeEndRisk < 0# Then ComputeEndRisk = 0#
End Function

Private Function SumLong(ByRef x() As Long, ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long, s As Long
    s = 0
    For i = i1 To i2
        s = s + x(i)
    Next i
    SumLong = s
End Function

Private Function FindFirstKMIndexAtOrAfter(ByRef x() As Double, ByVal t As Double) As Long
    Dim i As Long
    For i = LBound(x) To UBound(x)
        If x(i) >= t Then
            FindFirstKMIndexAtOrAfter = i
            Exit Function
        End If
    Next i
    FindFirstKMIndexAtOrAfter = 0
End Function

Private Function FindLastKMIndexBefore(ByRef x() As Double, ByVal t As Double) As Long
    Dim i As Long, outVal As Long
    outVal = 0
    For i = LBound(x) To UBound(x)
        If x(i) < t Then
            outVal = i
        Else
            Exit For
        End If
    Next i
    FindLastKMIndexBefore = outVal
End Function

Private Function RangeToVector(ByVal rng As Range) As Double()
    Dim arr As Variant
    Dim out() As Double
    Dim r As Long, c As Long, n As Long

    arr = rng.Value2
    ReDim out(1 To rng.CountLarge)

    n = 0
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If IsNumeric(arr(r, c)) Then
                n = n + 1
                out(n) = CDbl(arr(r, c))
            End If
        Next c
    Next r

    ReDim Preserve out(1 To n)
    RangeToVector = out
End Function

Private Function Resize2D(ByRef arr As Variant, ByVal nRows As Long, ByVal nCols As Long) As Variant
    Dim out() As Variant
    Dim r As Long, c As Long

    ReDim out(1 To nRows, 1 To nCols)
    For r = 1 To nRows
        For c = 1 To nCols
            out(r, c) = arr(r, c)
        Next c
    Next r

    Resize2D = out
End Function

Private Sub RemoveNonFinitePairs(ByRef x() As Double, ByRef y() As Double)
    Dim i As Long, m As Long, n As Long
    Dim xx() As Double, yy() As Double

    n = UBound(x)
    ReDim xx(1 To n)
    ReDim yy(1 To n)

    m = 0
    For i = 1 To n
        If IsFiniteDouble(x(i)) And IsFiniteDouble(y(i)) Then
            m = m + 1
            xx(m) = x(i)
            yy(m) = y(i)
        End If
    Next i

    If m = 0 Then
        ReDim xx(1 To 1)
        ReDim yy(1 To 1)
        xx(1) = 0#
        yy(1) = 1#
    Else
        ReDim Preserve xx(1 To m)
        ReDim Preserve yy(1 To m)
    End If

    x = xx
    y = yy
End Sub

Private Function IsFiniteDouble(ByVal x As Double) As Boolean
    IsFiniteDouble = True
    If x <> x Then IsFiniteDouble = False
End Function

Private Function UBoundSafe(ByRef x() As Double) As Long
    On Error GoTo EH
    UBoundSafe = UBound(x)
    Exit Function
EH:
    UBoundSafe = 0
End Function

Private Function MaxDouble(ByRef x() As Double) As Double
    Dim i As Long
    Dim m As Double

    m = x(LBound(x))
    For i = LBound(x) + 1 To UBound(x)
        If x(i) > m Then m = x(i)
    Next i
    MaxDouble = m
End Function

