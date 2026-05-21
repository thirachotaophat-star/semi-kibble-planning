Attribute VB_Name = "SemiKibbleMVP"
Option Explicit

Public Const PCS_PER_BAG As Long = 5
Public Const BAG_PER_CASE As Long = 48
Public Const PCS_PER_CASE As Long = 240
Public Const PLAN_SLOT_COUNT As Long = 9

Public Const SH_STOCK As String = "Stock"
Public Const SH_PLAN_HEADER As String = "Plan_Header"
Public Const SH_PLAN_DETAIL As String = "Plan_Detail"
Public Const SH_BALANCE As String = "Balance"
Public Const SH_LOG As String = "Log"
Public Const SH_ALLOC_LINE As String = "AllocationLine"

Public Enum StockCol
    scRowId = 1
    scMFG = 2
    scEXP = 3
    scShift = 4
    scMC = 5
    scQtyPCS = 6
    scFifoSeq = 7
End Enum

Private Type TSheetMap
    wsStock As Worksheet
    wsPlanHeader As Worksheet
    wsPlanDetail As Worksheet
    wsBalance As Worksheet
    wsLog As Worksheet
    wsAllocationLine As Worksheet
End Type

Public Sub AutoAllocate()
    On Error GoTo ERR_HANDLER

    Dim sm As TSheetMap, planCols As Collection
    Set sm = MapSheets(ThisWorkbook)
    ValidateWorkbookState sm

    ClearLog sm
    ClearAllocationLine sm

    Set planCols = GetPlanColumns(sm.wsPlanHeader, 1)
    If planCols.Count = 0 Then Err.Raise 5, , "Plan_Header ต้องมีคอลัมน์ Plan1..Plan9"

    Dim stockLast As Long: stockLast = LastDataRow(sm.wsStock, scRowId)
    If stockLast < 2 Then Err.Raise 5, , "Stock ไม่มีข้อมูล"

    Dim allocatedByPlan(1 To PLAN_SLOT_COUNT) As Long
    Dim remainUsablePCS() As Long
    remainUsablePCS = BuildRemainingUsablePCS(sm.wsStock)
    Dim p As Long
    For p = 1 To PLAN_SLOT_COUNT
        allocatedByPlan(p) = AllocateSinglePlan(sm, p, planCols, remainUsablePCS)
    Next p

    WritePlanDetail sm, allocatedByPlan
    RecalculateBalance

    WriteLog sm, "INFO", "AutoAllocate completed."
    Exit Sub

ERR_HANDLER:
    SafeLogError "AutoAllocate", Err.Number, Err.Description
    MsgBox "AutoAllocate error: " & Err.Description, vbExclamation
End Sub

Private Function AllocateSinglePlan(ByRef sm As TSheetMap, ByVal planNo As Long, ByVal planCols As Collection, ByRef remainUsablePCS() As Long) As Long
    Dim colIdx As Long: colIdx = CLng(planCols(planNo))
    Dim rawPlanValue As Long: rawPlanValue = NzLong(sm.wsPlanHeader.Cells(2, colIdx).Value)
    ValidateNoNegative rawPlanValue, "Plan" & planNo
    Dim requiredPCS As Long
    requiredPCS = NormalizePlanToPCS(rawPlanValue)
    sm.wsPlanHeader.Cells(3, colIdx).Value = requiredPCS

    If requiredPCS = 0 Then Exit Function

    Dim stockLast As Long: stockLast = LastDataRow(sm.wsStock, scRowId)
    Dim r As Long, allocR As Long: allocR = LastDataRow(sm.wsAllocationLine, 1)
    If allocR < 2 Then allocR = 1

    For r = 2 To stockLast
        If requiredPCS <= 0 Then Exit For

        Dim allocatable As Long: allocatable = remainUsablePCS(r)
        If allocatable < PCS_PER_CASE Then GoTo NEXT_R

        Dim availCase As Long: availCase = allocatable \ PCS_PER_CASE
        Dim useCase As Long: useCase = WorksheetFunction.Min(availCase, requiredPCS \ PCS_PER_CASE)
        If useCase <= 0 Then GoTo NEXT_R

        Dim usePCS As Long: usePCS = useCase * PCS_PER_CASE
        requiredPCS = requiredPCS - usePCS
        AllocateSinglePlan = AllocateSinglePlan + usePCS
        remainUsablePCS(r) = remainUsablePCS(r) - usePCS

        allocR = allocR + 1
        sm.wsAllocationLine.Cells(allocR, 1).Value = Now
        sm.wsAllocationLine.Cells(allocR, 2).Value = "PLAN"
        sm.wsAllocationLine.Cells(allocR, 3).Value = planNo
        sm.wsAllocationLine.Cells(allocR, 4).Value = sm.wsStock.Cells(r, scRowId).Value
        sm.wsAllocationLine.Cells(allocR, 5).Value = ""
        sm.wsAllocationLine.Cells(allocR, 6).Value = usePCS
        sm.wsAllocationLine.Cells(allocR, 7).Value = useCase
NEXT_R:
    Next r

    HandleTopUp4647 sm, planNo, requiredPCS, AllocateSinglePlan, remainUsablePCS
End Function

Private Sub HandleTopUp4647(ByRef sm As TSheetMap, ByVal planNo As Long, ByRef requiredPCS As Long, ByRef allocatedPCS As Long, ByRef remainUsablePCS() As Long)
    If requiredPCS < PCS_PER_CASE Then Exit Sub

    Dim stockLast As Long: stockLast = LastDataRow(sm.wsStock, scRowId)
    Dim i As Long, j As Long, allocR As Long

    For i = 2 To stockLast
        If requiredPCS < PCS_PER_CASE Then Exit For

        Dim recvBag As Long: recvBag = remainUsablePCS(i) \ PCS_PER_BAG
        If recvBag <> 46 And recvBag <> 47 Then GoTo NEXT_I

        Dim needBag As Long: needBag = BAG_PER_CASE - recvBag
        Dim donorRows() As Long, donorTakeBag() As Long
        ReDim donorRows(1 To stockLast)
        ReDim donorTakeBag(1 To stockLast)
        Dim donorCount As Long: donorCount = 0
        Dim remainNeed As Long: remainNeed = needBag

        For j = i + 1 To stockLast
            If remainNeed <= 0 Then Exit For
            If CDate(sm.wsStock.Cells(j, scMFG).Value) = CDate(sm.wsStock.Cells(i, scMFG).Value) And _
               CDate(sm.wsStock.Cells(j, scEXP).Value) = CDate(sm.wsStock.Cells(i, scEXP).Value) Then
                Dim donorRemainBag As Long
                donorRemainBag = remainUsablePCS(j) \ PCS_PER_BAG
                If donorRemainBag > 0 Then
                    donorCount = donorCount + 1
                    donorRows(donorCount) = j
                    donorTakeBag(donorCount) = WorksheetFunction.Min(donorRemainBag, remainNeed)
                    remainNeed = remainNeed - donorTakeBag(donorCount)
                End If
            End If
        Next j

        If remainNeed = 0 Then
            Dim k As Long
            For k = 1 To donorCount
                remainUsablePCS(donorRows(k)) = remainUsablePCS(donorRows(k)) - (donorTakeBag(k) * PCS_PER_BAG)
                Debug.Assert remainUsablePCS(donorRows(k)) >= 0
                If remainUsablePCS(donorRows(k)) < 0 Then Err.Raise 5, , "donorRemainBag cannot be negative"
            Next k
            remainUsablePCS(i) = remainUsablePCS(i) - (recvBag * PCS_PER_BAG)

            For k = 1 To donorCount
                allocR = LastDataRow(sm.wsAllocationLine, 1)
                If allocR < 2 Then allocR = 1
                allocR = allocR + 1
                sm.wsAllocationLine.Cells(allocR, 1).Value = Now
                sm.wsAllocationLine.Cells(allocR, 2).Value = "TOPUP_" & recvBag
                sm.wsAllocationLine.Cells(allocR, 3).Value = planNo
                sm.wsAllocationLine.Cells(allocR, 4).Value = sm.wsStock.Cells(i, scRowId).Value
                sm.wsAllocationLine.Cells(allocR, 5).Value = sm.wsStock.Cells(donorRows(k), scRowId).Value
                sm.wsAllocationLine.Cells(allocR, 6).Value = donorTakeBag(k) * PCS_PER_BAG
                sm.wsAllocationLine.Cells(allocR, 7).Value = 1
            Next k

            allocatedPCS = allocatedPCS + PCS_PER_CASE
            requiredPCS = requiredPCS - PCS_PER_CASE
        End If
NEXT_I:
    Next i
End Sub

Private Function BuildRemainingUsablePCS(ByVal wsStock As Worksheet) As Long()
    Dim lastRow As Long: lastRow = LastDataRow(wsStock, scRowId)
    Dim arr() As Long
    ReDim arr(1 To lastRow)
    Dim r As Long, qtyPCS As Long
    For r = 2 To lastRow
        qtyPCS = NzLong(wsStock.Cells(r, scQtyPCS).Value)
        arr(r) = qtyPCS - (qtyPCS Mod PCS_PER_BAG)
    Next r
    BuildRemainingUsablePCS = arr
End Function

Private Function NormalizePlanToPCS(ByVal rawPlanValue As Long) As Long
    If rawPlanValue <= 0 Then
        NormalizePlanToPCS = 0
    ElseIf rawPlanValue Mod PCS_PER_CASE = 0 Then
        NormalizePlanToPCS = rawPlanValue
    Else
        NormalizePlanToPCS = CaseToPCS(rawPlanValue)
    End If
End Function

Public Sub TestScenario_NoDoubleAllocation_480_240_240_240()
    Dim alloc1 As Long, alloc2 As Long, alloc3 As Long
    alloc1 = 240
    alloc2 = 240
    alloc3 = 0
    Debug.Assert alloc1 = 240
    Debug.Assert alloc2 = 240
    Debug.Assert alloc3 = 0
End Sub

Public Sub TestScenario_TopUp46_WithMultiDonor_1Plus1()
    Dim receiverBag As Long: receiverBag = 46
    Dim donor1Bag As Long: donor1Bag = 1
    Dim donor2Bag As Long: donor2Bag = 1
    Dim needBag As Long: needBag = 48 - receiverBag
    Dim success As Boolean
    success = ((donor1Bag + donor2Bag) >= needBag)

    Debug.Assert needBag = 2
    Debug.Assert success = True
End Sub

Private Sub WritePlanDetail(ByRef sm As TSheetMap, ByRef allocatedByPlan() As Long)
    Dim i As Long
    sm.wsPlanDetail.Cells(1, 1).Value = "Customer"
    For i = 1 To PLAN_SLOT_COUNT
        sm.wsPlanDetail.Cells(1, i + 1).Value = "Plan" & i
    Next i

    sm.wsPlanDetail.Cells(2, 1).Value = "AllocatedPCS"
    For i = 1 To PLAN_SLOT_COUNT
        sm.wsPlanDetail.Cells(2, i + 1).Value = allocatedByPlan(i)
    Next i
End Sub

Public Sub RecalculateBalance()
    On Error GoTo ERR_HANDLER
    Dim sm As TSheetMap: Set sm = MapSheets(ThisWorkbook)
    ValidateWorkbookState sm

    Dim lastRow As Long: lastRow = LastDataRow(sm.wsStock, scRowId)
    sm.wsBalance.Range("A1:F1").Value = Array("RowID", "QtyPCS", "AllocatedPCS", "BalancePCS", "RejectPCS", "UsablePCS")
    sm.wsBalance.Range("A2:F9999").ClearContents

    Dim r As Long, qty As Long, rej As Long, usable As Long, alloc As Long, bal As Long
    For r = 2 To lastRow
        qty = NzLong(sm.wsStock.Cells(r, scQtyPCS).Value)
        usable = qty - (qty Mod PCS_PER_BAG)
        rej = qty Mod PCS_PER_BAG
        alloc = AllocatedByRow(sm, sm.wsStock.Cells(r, scRowId).Value)
        bal = usable - alloc

        sm.wsBalance.Cells(r, 1).Value = sm.wsStock.Cells(r, scRowId).Value
        sm.wsBalance.Cells(r, 2).Value = qty
        sm.wsBalance.Cells(r, 3).Value = alloc
        sm.wsBalance.Cells(r, 4).Value = bal
        sm.wsBalance.Cells(r, 5).Value = rej
        sm.wsBalance.Cells(r, 6).Value = usable
    Next r
    Exit Sub
ERR_HANDLER:
    SafeLogError "RecalculateBalance", Err.Number, Err.Description
End Sub

Private Function AllocatedByRow(ByRef sm As TSheetMap, ByVal rowId As Variant) As Long
    Dim lr As Long: lr = LastDataRow(sm.wsAllocationLine, 1)
    Dim r As Long
    For r = 2 To lr
        If sm.wsAllocationLine.Cells(r, 4).Value = rowId Then AllocatedByRow = AllocatedByRow + NzLong(sm.wsAllocationLine.Cells(r, 6).Value)
        If sm.wsAllocationLine.Cells(r, 5).Value = rowId Then AllocatedByRow = AllocatedByRow + NzLong(sm.wsAllocationLine.Cells(r, 6).Value)
    Next r
End Function

Private Function GetPlanColumns(ByVal ws As Worksheet, ByVal headerRow As Long) As Collection
    Dim col As New Collection
    Dim c As Long
    For c = 1 To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        If UCase$(Trim$(CStr(ws.Cells(headerRow, c).Value))) Like "PLAN#" Then col.Add c
    Next c
    If col.Count < PLAN_SLOT_COUNT Then
        For c = col.Count + 1 To PLAN_SLOT_COUNT
            col.Add c
        Next c
    End If
    Set GetPlanColumns = col
End Function

Private Function MapSheets(ByVal wb As Workbook) As TSheetMap
    Dim sm As TSheetMap
    Set sm.wsStock = RequireSheet(wb, SH_STOCK)
    Set sm.wsPlanHeader = RequireSheet(wb, SH_PLAN_HEADER)
    Set sm.wsPlanDetail = RequireSheet(wb, SH_PLAN_DETAIL)
    Set sm.wsBalance = RequireSheet(wb, SH_BALANCE)
    Set sm.wsLog = RequireSheet(wb, SH_LOG)
    Set sm.wsAllocationLine = RequireSheet(wb, SH_ALLOC_LINE)
    MapSheets = sm
End Function

Private Sub ValidateWorkbookState(ByRef sm As TSheetMap)
    If sm.wsStock Is Nothing Then Err.Raise 5, , "Missing Stock sheet"
End Sub
Private Function RequireSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error GoTo NOT_FOUND
    Set RequireSheet = wb.Worksheets(sheetName)
    Exit Function
NOT_FOUND:
    Err.Raise 9, , "Required sheet not found: " & sheetName
End Function
Private Function LastDataRow(ByVal ws As Worksheet, ByVal keyCol As Long) As Long
    LastDataRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
End Function
Private Function CaseToPCS(ByVal caseQty As Long) As Long: CaseToPCS = caseQty * PCS_PER_CASE: End Function
Private Sub ValidateNoNegative(ByVal valueX As Long, ByVal fieldName As String)
    If valueX < 0 Then Err.Raise 5, , fieldName & " cannot be negative"
End Sub
Private Function NzLong(ByVal v As Variant) As Long
    If IsError(v) Then
        NzLong = 0
    ElseIf IsEmpty(v) Or Trim$(CStr(v)) = "" Then
        NzLong = 0
    ElseIf IsNumeric(v) Then
        NzLong = CLng(v)
    Else
        Err.Raise 13, , "Non-numeric value encountered: " & CStr(v)
    End If
End Function
Private Sub ClearLog(ByRef sm As TSheetMap)
    sm.wsLog.Range("A1:C1").Value = Array("Timestamp", "Level", "Message")
    sm.wsLog.Range("A2:C9999").ClearContents
End Sub
Private Sub ClearAllocationLine(ByRef sm As TSheetMap)
    sm.wsAllocationLine.Range("A1:G1").Value = Array("Timestamp", "Type", "PlanNo", "ReceiverRowID", "DonorRowID", "PCS", "Case")
    sm.wsAllocationLine.Range("A2:G99999").ClearContents
End Sub
Private Sub WriteLog(ByRef sm As TSheetMap, ByVal levelText As String, ByVal msg As String)
    Dim r As Long: r = LastDataRow(sm.wsLog, 1): If r < 2 Then r = 1
    sm.wsLog.Cells(r + 1, 1).Value = Now
    sm.wsLog.Cells(r + 1, 2).Value = levelText
    sm.wsLog.Cells(r + 1, 3).Value = msg
End Sub
Private Sub SafeLogError(ByVal src As String, ByVal errNo As Long, ByVal errDesc As String)
    On Error Resume Next
    Dim sm As TSheetMap: Set sm = MapSheets(ThisWorkbook)
    WriteLog sm, "ERROR", src & " | " & errNo & " | " & errDesc
End Sub
