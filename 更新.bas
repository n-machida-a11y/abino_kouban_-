Attribute VB_Name = "更新"
Option Explicit

' IS_TEST_MODE / TEST_FILE_PATH / SHEET_* / CELL_* は Config モジュールで一元管理。

'================================================================================
' 「工事番号一覧」と「依頼履歴」を両方更新するマクロ
'================================================================================
Sub UpdateAllSheets()
    Dim originalScreenUpdating As Boolean
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean

    originalScreenUpdating = Application.ScreenUpdating
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerUpdateAll

    Call UpdateKoujiBangoListSheet(False)
    Call UpdateIraiRirekiSheet(False)

    MsgBox "「工事番号一覧」と「依頼履歴」シートが正常に更新されました。", vbInformation, "更新完了"

FinalizeUpdateAll:
    Application.ScreenUpdating = originalScreenUpdating
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Exit Sub

ErrorHandlerUpdateAll:
    MsgBox "一括更新中にエラーが発生しました: " & Err.Description, vbCritical, "更新エラー"
    Resume FinalizeUpdateAll
End Sub

'================================================================================
' 「工事番号一覧」シートを最新の情報に更新するマクロ
'================================================================================
Sub UpdateKoujiBangoListSheet(Optional ByVal ShowMessage As Boolean = True)
    Dim wbTarget As Workbook
    Dim wsSource As Worksheet
    Dim wsMaster As Worksheet
    Dim wsDest As Worksheet
    Dim targetFilePath As String
    Dim destSheetName As String
    Dim lastRowSource As Long
    Dim copyRange As Range

    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean
    Dim originalScreenUpdating As Boolean

    originalScreenUpdating = Application.ScreenUpdating
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerUpdateList

    targetFilePath = GetTargetFilePath()

    If Dir(targetFilePath) = "" Then
        MsgBox "対象のファイルが見つかりません。" & vbCrLf & targetFilePath, vbCritical, "ファイルエラー"
        GoTo FinalizeUpdateList
    End If

    On Error Resume Next
    Set wbTarget = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo ErrorHandlerUpdateList

    If wbTarget Is Nothing Then
        MsgBox "対象のExcelファイルを開けませんでした。" & vbCrLf & _
               "他のユーザーが使用中である可能性があります。", vbCritical
        GoTo FinalizeUpdateList
    End If

    If Not SheetExists(wbTarget, SHEET_KOUJI_LIST) Or Not SheetExists(wbTarget, SHEET_KANRI_MASTER) Then
        MsgBox "外部ファイルに必要なシート「" & SHEET_KOUJI_LIST & "」または「" & SHEET_KANRI_MASTER & "」が見つかりません。", vbCritical, "シートエラー"
        GoTo FinalizeUpdateList
    End If
    Set wsSource = wbTarget.Sheets(SHEET_KOUJI_LIST)
    Set wsMaster = wbTarget.Sheets(SHEET_KANRI_MASTER)

    destSheetName = Trim(CStr(wsMaster.Range(CELL_LOCAL_COPY_SHEET).Value))
    If destSheetName = "" Then
        MsgBox "外部ファイルの「" & SHEET_KANRI_MASTER & "」シート" & CELL_LOCAL_COPY_SHEET & "セルにコピー先のシート名が指定されていません。", vbExclamation
        GoTo FinalizeUpdateList
    End If

    If Not SheetExists(ThisWorkbook, destSheetName) Then
        MsgBox "このファイルにコピー先のシート「" & destSheetName & "」が見つかりませんでした。", vbExclamation
        GoTo FinalizeUpdateList
    End If
    Set wsDest = ThisWorkbook.Sheets(destSheetName)

    Call SafeUnprotect(wsDest)
    wsDest.Range("A3:X" & wsDest.Rows.count).Clear

    lastRowSource = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row

    If lastRowSource >= 5 Then
        Set copyRange = wsSource.Range("A5:X" & lastRowSource)
        copyRange.Copy Destination:=wsDest.Range("A3")
    End If

    If ShowMessage Then
        MsgBox "「" & destSheetName & "」シートが正常に更新されました。", vbInformation, "更新完了"
    End If
    Call SafeProtectData(wsDest)

FinalizeUpdateList:
    Application.CutCopyMode = False
    If Not wsMaster Is Nothing Then Set wsMaster = Nothing
    If Not wsSource Is Nothing Then Set wsSource = Nothing
    If Not wsDest Is Nothing Then Set wsDest = Nothing
    If Not wbTarget Is Nothing Then wbTarget.Close SaveChanges:=False

    Application.ScreenUpdating = originalScreenUpdating
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Exit Sub

ErrorHandlerUpdateList:
    MsgBox "「" & SHEET_KOUJI_LIST & "」シートの更新中にエラーが発生しました: " & Err.Description, vbCritical, "更新エラー"
    Resume FinalizeUpdateList
End Sub

'================================================================================
' 「依頼履歴」シートを最新の情報に更新するマクロ
'================================================================================
Sub UpdateIraiRirekiSheet(Optional ByVal ShowMessage As Boolean = True)
    Dim wbTarget As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim targetFilePath As String
    Dim lastRowSource As Long
    Dim copyRange As Range

    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean
    Dim originalScreenUpdating As Boolean

    originalScreenUpdating = Application.ScreenUpdating
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerUpdateRireki

    targetFilePath = GetTargetFilePath()

    If Dir(targetFilePath) = "" Then
        MsgBox "対象のファイルが見つかりません。" & vbCrLf & targetFilePath, vbCritical, "ファイルエラー"
        GoTo FinalizeUpdateRireki
    End If

    On Error Resume Next
    Set wbTarget = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo ErrorHandlerUpdateRireki

    If wbTarget Is Nothing Then
        MsgBox "対象のExcelファイルを開けませんでした。" & vbCrLf & _
               "他のユーザーが使用中である可能性があります。", vbCritical
        GoTo FinalizeUpdateRireki
    End If

    If Not SheetExists(wbTarget, SHEET_IRAI_RIREKI) Then
        MsgBox "外部ファイルに必要なシート「" & SHEET_IRAI_RIREKI & "」が見つかりません。", vbCritical, "シートエラー"
        GoTo FinalizeUpdateRireki
    End If
    Set wsSource = wbTarget.Sheets(SHEET_IRAI_RIREKI)

    If Not SheetExists(ThisWorkbook, SHEET_IRAI_RIREKI) Then
        MsgBox "このファイルにコピー先のシート「" & SHEET_IRAI_RIREKI & "」が見つかりませんでした。", vbExclamation
        GoTo FinalizeUpdateRireki
    End If
    Set wsDest = ThisWorkbook.Sheets(SHEET_IRAI_RIREKI)

    Call SafeUnprotect(wsDest)
    wsDest.Range("A3:W" & wsDest.Rows.count).Clear

    lastRowSource = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row

    If lastRowSource >= 2 Then
        Set copyRange = wsSource.Range("A2:W" & lastRowSource)
        copyRange.Copy Destination:=wsDest.Range("A3")
    End If

    If ShowMessage Then
        MsgBox "「" & SHEET_IRAI_RIREKI & "」シートが正常に更新されました。", vbInformation, "更新完了"
    End If
    Call SafeProtectData(wsDest)

FinalizeUpdateRireki:
    Application.CutCopyMode = False
    If Not wsSource Is Nothing Then Set wsSource = Nothing
    If Not wsDest Is Nothing Then Set wsDest = Nothing
    If Not wbTarget Is Nothing Then wbTarget.Close SaveChanges:=False

    Application.ScreenUpdating = originalScreenUpdating
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Exit Sub

ErrorHandlerUpdateRireki:
    MsgBox "「" & SHEET_IRAI_RIREKI & "」シートの更新中にエラーが発生しました: " & Err.Description, vbCritical, "更新エラー"
    Resume FinalizeUpdateRireki
End Sub


