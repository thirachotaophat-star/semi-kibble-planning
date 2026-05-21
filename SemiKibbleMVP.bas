Attribute VB_Name = "SemiKibbleMVP"
Option Explicit

' =============================
' Semi Kibble FIFO Allocation MVP
' Scope: core structure + declarations + mapping + helpers + validation
' =============================

' ----- Unit constants -----
Public Const PCS_PER_BAG As Long = 5
Public Const BAG_PER_CASE As Long = 48
Public Const PCS_PER_CASE As Long = 240
Public Const PLAN_SLOT_COUNT As Long = 9

' ----- Sheet names -----
Public Const SH_STOCK As String = "Stock"
Public Const SH_PLAN_HEADER As String = "Plan_Header"
Public Const SH_PLAN_DETAIL As String = "Plan_Detail"
Public Const SH_BALANCE As String = "Balance"
Public Const SH_LOG As String = "Log"

' ----- Column mapping (Stock) -----
Public Enum StockCol
    scRowId = 1
    scMFG = 2
    scEXP = 3
    scShift = 4
    scMC = 5
    scQtyPCS = 6
    scFifoSeq = 7
End Enum

' ----- Column mapping (Plan_Header) -----
Public Enum PlanHeaderCol
    phPlanQtyCase = 1
    phPlanQtyPCS = 2
End Enum

' ----- Column mapping (Plan_Detail Plan1..Plan9 in PCS) -----
Public Enum PlanDetailCol
    pdPlanNo = 1
    pdPlanPCS = 2
    pdAllocatedPCS = 3
End Enum

' ----- State container -----
Private Type TSheetMap
    wsStock As Worksheet
    wsPlanHeader As Worksheet
    wsPlanDetail As Worksheet
    wsBalance As Worksheet
    wsLog As Worksheet
End Type

Private Type TPlanContext
    PlanCaseTotal As Long
    PlanPCSTotal As Long
End Type

' =====================================================
' Public entry points
' =====================================================

Public Sub AutoAllocate()
    On Error GoTo ERR_HANDLER

    Dim sm As TSheetMap
    Dim ctx As TPlanContext

    Set sm = MapSheets(ThisWorkbook)

    ValidateWorkbookState sm
    ctx = BuildPlanContext(sm)
    ValidatePlanContext ctx

    ' NOTE: MVP รอบนี้เตรียมโครงสร้าง + validation + helper เท่านั้น
    ' Allocation engine เต็มรูปแบบจะเติมใน iteration ถัดไป

    ClearLog sm
    WriteLog sm, "INFO", "AutoAllocate started (MVP structure only)."
    WriteLog sm, "INFO", "Plan Case=" & ctx.PlanCaseTotal & ", Plan PCS=" & ctx.PlanPCSTotal

    RecalculateBalance

    WriteLog sm, "INFO", "AutoAllocate completed (no allocation transaction in MVP)."
    Exit Sub

ERR_HANDLER:
    SafeLogError "AutoAllocate", Err.Number, Err.Description
    MsgBox "AutoAllocate error: " & Err.Description, vbExclamation
End Sub

Public Sub ClearPlan()
    On Error GoTo ERR_HANDLER

    Dim sm As TSheetMap
    Set sm = MapSheets(ThisWorkbook)

    ValidateWorkbookState sm

    sm.wsPlanDetail.Range("B2:C" & (1 + PLAN_SLOT_COUNT)).ClearContents
    sm.wsPlanHeader.Cells(2, phPlanQtyPCS).ClearContents

    ClearLog sm
    WriteLog sm, "INFO", "Plan detail/header cleared."
    Exit Sub

ERR_HANDLER:
    SafeLogError "ClearPlan", Err.Number, Err.Description
    MsgBox "ClearPlan error: " & Err.Description, vbExclamation
End Sub

Public Sub RecalculateBalance()
    On Error GoTo ERR_HANDLER

    Dim sm As TSheetMap
    Set sm = MapSheets(ThisWorkbook)

    ValidateWorkbookState sm

    Dim lastRow As Long
    lastRow = LastDataRow(sm.wsStock, scRowId)

    Dim r As Long
    Dim qtyPCS As Long, usableBag As Long, usablePCS As Long, rejectPCS As Long

    ' Balance output columns (A:E): RowID, QtyPCS, UsablePCS, RejectPCS, BalancePCS
    sm.wsBalance.Range("A1:E1").Value = Array("RowID", "QtyPCS", "UsablePCS", "RejectPCS", "BalancePCS")

    If lastRow < 2 Then
        sm.wsBalance.Range("A2:E9999").ClearContents
        WriteLog sm, "WARN", "Stock has no data rows."
        Exit Sub
    End If

    sm.wsBalance.Range("A2:E9999").ClearContents

    For r = 2 To lastRow
        qtyPCS = NzLong(sm.wsStock.Cells(r, scQtyPCS).Value)
        ValidateNoNegative qtyPCS, "Stock row " & r & " QtyPCS"

        usableBag = qtyPCS \ PCS_PER_BAG
        usablePCS = usableBag * PCS_PER_BAG
        rejectPCS = qtyPCS Mod PCS_PER_BAG

        sm.wsBalance.Cells(r, 1).Value = sm.wsStock.Cells(r, scRowId).Value
        sm.wsBalance.Cells(r, 2).Value = qtyPCS
        sm.wsBalance.Cells(r, 3).Value = usablePCS
        sm.wsBalance.Cells(r, 4).Value = rejectPCS
        sm.wsBalance.Cells(r, 5).Value = usablePCS
    Next r

    WriteLog sm, "INFO", "Balance recalculated. Rows=" & (lastRow - 1)
    Exit Sub

ERR_HANDLER:
    SafeLogError "RecalculateBalance", Err.Number, Err.Description
    MsgBox "RecalculateBalance error: " & Err.Description, vbExclamation
End Sub

' =====================================================
' Mapping / context
' =====================================================

Private Function MapSheets(ByVal wb As Workbook) As TSheetMap
    Dim sm As TSheetMap

    Set sm.wsStock = RequireSheet(wb, SH_STOCK)
    Set sm.wsPlanHeader = RequireSheet(wb, SH_PLAN_HEADER)
    Set sm.wsPlanDetail = RequireSheet(wb, SH_PLAN_DETAIL)
    Set sm.wsBalance = RequireSheet(wb, SH_BALANCE)
    Set sm.wsLog = RequireSheet(wb, SH_LOG)

    MapSheets = sm
End Function

Private Sub ValidateWorkbookState(ByRef sm As TSheetMap)
    If sm.wsStock Is Nothing Then Err.Raise 5, , "Missing Stock sheet"
    If sm.wsPlanHeader Is Nothing Then Err.Raise 5, , "Missing Plan_Header sheet"
    If sm.wsPlanDetail Is Nothing Then Err.Raise 5, , "Missing Plan_Detail sheet"
    If sm.wsBalance Is Nothing Then Err.Raise 5, , "Missing Balance sheet"
    If sm.wsLog Is Nothing Then Err.Raise 5, , "Missing Log sheet"
End Sub

Private Function BuildPlanContext(ByRef sm As TSheetMap) As TPlanContext
    Dim ctx As TPlanContext

    ctx.PlanCaseTotal = NzLong(sm.wsPlanHeader.Cells(2, phPlanQtyCase).Value)
    ValidateNoNegative ctx.PlanCaseTotal, "Plan Qty Case"

    ctx.PlanPCSTotal = CaseToPCS(ctx.PlanCaseTotal)
    sm.wsPlanHeader.Cells(2, phPlanQtyPCS).Value = ctx.PlanPCSTotal

    BuildPlanContext = ctx
End Function

Private Sub ValidatePlanContext(ByRef ctx As TPlanContext)
    If ctx.PlanPCSTotal Mod PCS_PER_CASE <> 0 Then
        Err.Raise 5, , "Plan PCS must be divisible by 240"
    End If
End Sub

' =====================================================
' Helpers
' =====================================================

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

Private Function CaseToPCS(ByVal caseQty As Long) As Long
    CaseToPCS = caseQty * PCS_PER_CASE
End Function

Private Function PCSToBag(ByVal pcsQty As Long) As Long
    PCSToBag = pcsQty \ PCS_PER_BAG
End Function

Private Function BagToPCS(ByVal bagQty As Long) As Long
    BagToPCS = bagQty * PCS_PER_BAG
End Function

Private Function IsValidOuterBagMultiple(ByVal pcsQty As Long) As Boolean
    IsValidOuterBagMultiple = (pcsQty Mod PCS_PER_BAG = 0)
End Function

Private Function IsCaseMultiple(ByVal pcsQty As Long) As Boolean
    IsCaseMultiple = (pcsQty Mod PCS_PER_CASE = 0)
End Function

Private Sub ValidateNoNegative(ByVal valueX As Long, ByVal fieldName As String)
    If valueX < 0 Then
        Err.Raise 5, , fieldName & " cannot be negative"
    End If
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

Private Sub WriteLog(ByRef sm As TSheetMap, ByVal levelText As String, ByVal msg As String)
    Dim r As Long
    r = LastDataRow(sm.wsLog, 1)
    If r < 2 Then r = 1

    sm.wsLog.Cells(r + 1, 1).Value = Now
    sm.wsLog.Cells(r + 1, 2).Value = levelText
    sm.wsLog.Cells(r + 1, 3).Value = msg
End Sub

Private Sub SafeLogError(ByVal src As String, ByVal errNo As Long, ByVal errDesc As String)
    On Error Resume Next
    Dim sm As TSheetMap
    Set sm = MapSheets(ThisWorkbook)
    WriteLog sm, "ERROR", src & " | " & errNo & " | " & errDesc
End Sub
