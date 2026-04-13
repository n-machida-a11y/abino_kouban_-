VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' ThisWorkbook イベント
' 複数人利用のためのシート保護・フィルタ管理
'================================================================================

' ファイルを開いた時に自動実行
' 全シートのフィルタをクリアし、適切な保護をかけ直す
Private Sub Workbook_Open()
    Dim ws As Worksheet
    Dim originalScreenUpdating As Boolean
    Dim originalDisplayAlerts As Boolean

    originalScreenUpdating = Application.ScreenUpdating
    originalDisplayAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo ErrorHandler

    For Each ws In ThisWorkbook.Worksheets
        ' 一旦保護を解除してフィルタをクリア
        Call SafeUnprotect(ws)
        Call ClearAllFilters(ws)

        ' シート名に応じて再保護
        Select Case ws.Name
            Case SHEET_IRAISHO
                ' 請求書提出依頼書：完全保護（マクロ専用）
                Call SafeProtectFull(ws)
            Case SHEET_KOUJI_LIST, SHEET_IRAI_RIREKI
                ' 工事番号一覧、依頼履歴：フィルタOK、編集NG
                Call SafeProtectData(ws)
            Case SHEET_KANRI_MASTER, SHEET_OTHER_MASTER
                ' 管理マスタ、その他マスタ：完全保護
                Call SafeProtectFull(ws)
            Case Else
                ' その他のシート（入力フォーム等）は保護しない
        End Select
    Next ws

Finalize:
    Application.ScreenUpdating = originalScreenUpdating
    Application.DisplayAlerts = originalDisplayAlerts
    Exit Sub

ErrorHandler:
    ' エラーが発生しても起動処理は継続（ユーザー体験優先）
    Resume Finalize
End Sub
