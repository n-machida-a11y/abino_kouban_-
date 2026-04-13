VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 依頼書作成 
   Caption         =   "依頼書作成フォーム"
   ClientHeight    =   13836
   ClientLeft      =   24
   ClientTop       =   72
   ClientWidth     =   11136
   OleObjectBlob   =   "全社用依頼書作成.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "依頼書作成"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' IS_TEST_MODE / TEST_FILE_PATH / SHEET_* / CELL_* / PATH_CELL は Config モジュールで一元管理。

'--- 他のフォームから値を受け取るための変数 ---
Public m_SearchedKoujiName As String ' 検索対象の「工事名称」
Public m_PassedTantousha As String  ' 検索対象の「担当者名」

'--- このフォーム内だけで使う変数 ---
' ドロップダウンリスト同士が無限に更新しあうのを防ぐためのスイッチ（フラグ）
Private m_isComboBoxUpdating As Boolean
' 更新対象のデータの「行番号」を保存しておく変数
Private m_TargetRow As Long
' フォームに表示しているデータの「工事番号」を保存しておく変数
Private m_KoujiBangou As String

'--- 高速化のため、各種リストを一時保存するキャッシュ変数 ---
Private m_CachedSeikyuusakiList As Variant      ' 「請求書提出先」リスト
Private m_CachedTeishutsuyoukouList As Variant ' 「提出要項」リスト
Private m_CachedDoufuubutsuList As Variant     ' 「同封物」リスト

'--- このフォームで処理する外部ファイルのパスを保存する変数 ---
Private m_TARGET_FILE_PATH As String

'--- プログラム内で固定的に使う文字（定数） ---
' 外部ファイル内のシート名（Config モジュールの定数を使用）
' SHEET_KOUJI_LIST   = "工事番号一覧"  (旧 SHEET_KOUJI_LIST)
' SHEET_KANRI_MASTER = "管理マスタ"    (旧 SHEET_KANRI_MASTER)
' SHEET_OTHER_MASTER = "その他マスタ"  (旧 SHEET_OTHER_MASTER)
' SHEET_IRAI_RIREKI  = "依頼履歴"      (旧 SHEET_IRAI_RIREKI)
' 「その他マスタ」シート内の列定義
Private Const MASTER_OTHER_SEIKYUUSAKI_COL As String = "A"      ' 請求書提出先
Private Const MASTER_OTHER_YUBIN_NO_COL As String = "B"         ' 郵便番号
Private Const MASTER_OTHER_JUSHO_COL As String = "C"            ' 住所
Private Const MASTER_OTHER_TEISHUTSUYOUKOU_COL As String = "G" ' 提出要項
Private Const MASTER_OTHER_DOUFUUBUTSU_COL As String = "I"      ' 同封物

' 同封物リストボックスの区切り文字（保存用）
Private Const DOUFUUBUTSU_DELIMITER_SAVE As String = ","
' 同封物リストボックスの区切り文字（依頼書表示用）
Private Const DOUFUUBUTSU_DELIMITER_DISPLAY As String = "、"


'================================================================================
' ★★★ 他のモジュールからこのフォームを呼び出すための専用プロシージャ ★★★
' 外部から .Show で直接表示するのではなく、このプロシージャを経由して呼び出します。
'================================================================================
Public Sub SetupAndShow(ByVal KoujiName As String, ByVal Tantousha As String)
    ' --- ① 他のモジュールから工事名と担当者名を受け取る ---
    Me.m_SearchedKoujiName = KoujiName
    Me.m_PassedTantousha = Tantousha

    ' --- ② フォームが表示される前に、初期値を設定する ---
    Me.作成日.Value = Format(Date, "yyyy/mm/dd")
    Me.担当者.Value = Me.m_PassedTantousha
    Me.工事名称.Value = Me.m_SearchedKoujiName
    Me.提出日付.Value = Format(Date, "yyyy/mm/dd")
    Me.数量1.Value = 1
    Me.単位1.Value = "式"
    Me.小計.Value = 0
    Me.消費税.Value = 0
    Me.請求金額.Value = 0

    ' --- ③ 初期値設定後に、フォームを表示する ---
    Me.Show
End Sub

Private Sub Label13_Click()

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Label18_Click()

End Sub

Private Sub Label35_Click()

End Sub

Private Sub Label36_Click()

End Sub

Private Sub Label41_Click()

End Sub

Private Sub Label45_Click()

End Sub

Private Sub Label46_Click()

End Sub

Private Sub Label47_Click()

End Sub

Private Sub Label57_Click()

End Sub

Private Sub Label58_Click()

End Sub

Private Sub Label59_Click()

End Sub

Private Sub Label62_Click()

End Sub

Private Sub Label67_Click()

End Sub

Private Sub Label68_Click()

End Sub

Private Sub Label73_Click()

End Sub

Private Sub TextBox14_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

'================================================================================
' フォームが開かれる瞬間の準備処理（１回だけ実行）
' ここでは、ドロップダウンリストの中身など、フォームの基本的な部品を準備します。
'================================================================================
Private Sub UserForm_Initialize()
    Dim wbTarget_Init As Workbook
    Dim wsKanriMaster As Worksheet, wsOtherMaster As Worksheet
    Dim originalDisplayAlerts As Boolean, originalEnableEvents As Boolean

    ' --- 処理中の画面のちらつきや不要なメッセージを抑制 ---
    Application.ScreenUpdating = False
    originalDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    originalEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    m_TargetRow = 0 ' 更新対象の行番号をリセット
    Me.請求書提出先.MatchEntry = fmMatchEntryNone ' ドロップダウンの入力補完を無効化（自前のロジックを使うため）

    On Error GoTo ErrorHandlerInit

    ' --- キャッシュデータがあれば、ファイルを開かずに高速化 ---
    If Not IsEmpty(m_CachedSeikyuusakiList) Then GoTo FinalizeInitWithoutFileOpen

    ' --- 外部ファイルを開き、ドロップダウンリスト用のマスタデータを読み込む ---
    m_TARGET_FILE_PATH = GetTargetFilePath()

    If Dir(m_TARGET_FILE_PATH) = "" Then GoTo FinalizeInit

    Set wbTarget_Init = Application.Workbooks.Open(fileName:=m_TARGET_FILE_PATH, ReadOnly:=True, UpdateLinks:=0)

    If Not SheetExists(wbTarget_Init, SHEET_KANRI_MASTER) Or Not SheetExists(wbTarget_Init, SHEET_OTHER_MASTER) Then GoTo FinalizeInit
    Set wsKanriMaster = wbTarget_Init.Sheets(SHEET_KANRI_MASTER)
    Set wsOtherMaster = wbTarget_Init.Sheets(SHEET_OTHER_MASTER)

    ' --- 各種マスタデータを読み込み、キャッシュに保存 ---
    m_CachedSeikyuusakiList = GetColumnData(wsOtherMaster, MASTER_OTHER_SEIKYUUSAKI_COL, 2)
    Me.請求書提出先.List = m_CachedSeikyuusakiList

    m_CachedTeishutsuyoukouList = GetColumnData(wsOtherMaster, MASTER_OTHER_TEISHUTSUYOUKOU_COL, 2)
    Call PopulateComboBoxFromCache(Me.提出要項, m_CachedTeishutsuyoukouList)

    m_CachedDoufuubutsuList = GetColumnData(wsOtherMaster, MASTER_OTHER_DOUFUUBUTSU_COL, 2)
    Call PopulateListBoxFromCache(Me.同封物, m_CachedDoufuubutsuList)

FinalizeInit: ' 終了処理（ファイルを開いた場合）
    If Not wbTarget_Init Is Nothing Then wbTarget_Init.Close SaveChanges:=False

FinalizeInitWithoutFileOpen: ' 終了処理（ファイルを開かなかった場合も含む）
    ' 作成日を編集不可にする
    Me.作成日.Enabled = False
    
    ' 小計と消費税は自動計算なので、手入力できないようにする
    Me.小計.Enabled = False
    Me.消費税.Enabled = False

    ' --- Excelの設定を元に戻す ---
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandlerInit: ' エラー発生時の処理
    Resume FinalizeInit
End Sub

'================================================================================
' フォームがアクティブになった時の処理（特定の工事データを検索し、フォームに表示）
'================================================================================
Private Sub UserForm_Activate()
    Dim wbTarget_Activate As Workbook
    Dim wsTarget_Activate As Worksheet
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean

    ' --- 処理中の画面のちらつきや不要なメッセージを抑制 ---
    Application.ScreenUpdating = False
    originalDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    originalEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerActivate

    ' --- 編集対象のデータが指定されているか確認 ---
    If Trim(Me.m_SearchedKoujiName) = "" Or Trim(Me.m_PassedTantousha) = "" Then
        Unload Me
        GoTo FinalizeActivate
    End If

    ' Initializeで準備したキャッシュデータがあるか確認
    If IsEmpty(m_CachedSeikyuusakiList) Then
        Unload Me
        GoTo FinalizeActivate
    End If

    ' --- 外部ファイルを読み取り専用で開く ---
    Dim targetFilePath As String
    targetFilePath = GetTargetFilePath()

    If Dir(targetFilePath) = "" Then
        Unload Me
        GoTo FinalizeActivate
    End If

    Set wbTarget_Activate = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=True, UpdateLinks:=0)

    If Not SheetExists(wbTarget_Activate, SHEET_KOUJI_LIST) Then
        GoTo FinalizeActivate
    End If
    Set wsTarget_Activate = wbTarget_Activate.Sheets(SHEET_KOUJI_LIST)

    ' --- データシートを１行ずつ調べ、工事名称と担当者の両方が一致する行を探す ---
    m_TargetRow = 0
    Dim r As Long
    Dim sheetKoujiName As String, sheetStaffName As String
    For r = 2 To wsTarget_Activate.Cells(wsTarget_Activate.Rows.count, "E").End(xlUp).Row
        sheetKoujiName = Trim(CStr(wsTarget_Activate.Cells(r, "E").Value))
        sheetStaffName = Trim(CStr(wsTarget_Activate.Cells(r, "C").Value))

        If sheetKoujiName = Trim(Me.m_SearchedKoujiName) And sheetStaffName = Trim(Me.m_PassedTantousha) Then
            m_TargetRow = r ' 一致する行が見つかった
            Exit For      ' ループを抜ける
        End If
    Next r

    ' --- 検索結果の処理 ---
    If m_TargetRow > 0 Then
        ' データが見つかった場合、工事番号を取得し、既存のデータをフォームに読み込む
        m_KoujiBangou = wsTarget_Activate.Cells(m_TargetRow, "D").Value
        Call LoadDataToForm(wsTarget_Activate, m_TargetRow)
    Else
        ' データが見つからなかった場合はフォームを閉じる
        Unload Me
        GoTo FinalizeActivate
    End If

    ' 金額の合計を再計算
    m_isComboBoxUpdating = False
    Call CalculateTotals

FinalizeActivate: ' 終了処理
    Me.小計.Enabled = False
    Me.消費税.Enabled = False

    If Not wbTarget_Activate Is Nothing Then wbTarget_Activate.Close SaveChanges:=False
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandlerActivate: ' エラー発生時の処理
    Resume FinalizeActivate
End Sub

'================================================================================
' 「依頼書作成」ボタンがクリックされたときの処理
'================================================================================
Private Sub 依頼書作成_Click()
    Dim wbTarget_Click As Workbook
    Dim wsRequest As Worksheet, wsSrc As Worksheet, wsMaster_Click As Worksheet
    Dim wsRireki As Worksheet
    Dim wsRirekiLocal As Worksheet
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean
    Dim isSuccess As Boolean

    isSuccess = False ' 処理成功フラグを初期化

    ' --- 処理中の画面のちらつきや不要なメッセージを抑制 ---
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerClick

    ' --- 事前チェック ---
    If m_TargetRow = 0 Then GoTo CleanUpClick ' 更新対象行が不明なら中断
    If MsgBox("依頼書シートを作成します。よろしいですか？", vbYesNo + vbQuestion, "確認") = vbNo Then GoTo CleanUpClick ' ユーザーがキャンセルしたら中断

    ' --- ① 外部ファイルを書き込みモードで開く ---
    Dim targetFilePath As String
    targetFilePath = GetTargetFilePath()

    ' ファイルが既に開かれていないかチェック（二重編集を防ぐため）
    Dim openedWbMain As Workbook
    For Each openedWbMain In Application.Workbooks
        If openedWbMain.FullName = targetFilePath Then
            MsgBox "対象のExcelファイルが既に開かれています。閉じてください。", vbCritical
            GoTo CleanUpClick
        End If
    Next openedWbMain

    On Error Resume Next
    Set wbTarget_Click = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=False, UpdateLinks:=0)
    On Error GoTo ErrorHandlerClick

    If wbTarget_Click Is Nothing Then
        MsgBox "対象のExcelファイルを開けませんでした。処理を中断します。", vbCritical
        GoTo CleanUpClick
    End If
    If wbTarget_Click.ReadOnly Then
        MsgBox "対象のExcelファイルは読み取り専用で開かれました。" & vbCrLf & _
               "他のユーザーが使用中の可能性があります。処理を中断します。", vbExclamation
        GoTo CleanUpClick
    End If

    ' 必要なシートが存在するかチェック
    If Not SheetExists(wbTarget_Click, SHEET_KOUJI_LIST) Or _
       Not SheetExists(wbTarget_Click, SHEET_OTHER_MASTER) Or _
       Not SheetExists(wbTarget_Click, SHEET_IRAI_RIREKI) Then
        MsgBox "必要なシート（工事番号一覧、その他マスタ、依頼履歴）のいずれかが見つかりません。", vbCritical
        GoTo CleanUpClick
    End If
    Set wsSrc = wbTarget_Click.Sheets(SHEET_KOUJI_LIST)
    Set wsMaster_Click = wbTarget_Click.Sheets(SHEET_OTHER_MASTER)
    Set wsRireki = wbTarget_Click.Sheets(SHEET_IRAI_RIREKI)

    ' --- ② フォームの入力内容で、外部ファイルのデータを更新 ---
    ' (2-1) 「その他マスタ」に、新しい請求先情報があれば追加・更新する
    Call UpdateAddressMaster(wsMaster_Click)
    
    ' (2-2) 「工事番号一覧」シートのN列以降を更新する
    Call UpdateExternalFile(wsSrc, m_TargetRow)

    ' (2-2) 「依頼履歴」シートに、今回の依頼内容を新しい行として追記する
    Call AddDataToIraiRireki(wsRireki)
    
    ' (2-4) 「工事番号一覧」の色付け処理を停止（ユーザーの要望により）
    ' With wsSrc.Rows(m_TargetRow).Interior
    '     .Pattern = xlSolid
    '     .Color = RGB(220, 220, 220)  ' 色を薄めのグレーに設定
    ' End With

    ' 更新内容を保存
    Application.EnableEvents = False ' イベントを一時的に停止
    wbTarget_Click.Save              ' 保存を実行
    Application.EnableEvents = True  ' イベントを元に戻す

    ' --- ③ このツール内の「請求書提出依頼書」シートに内容を転記 ---
    Set wsRequest = ThisWorkbook.Sheets("請求書提出依頼書")
    wsRequest.Unprotect

    With wsRequest
        .Range("F7").Value = Me.請求書提出先.Value
        .Range("G8").Value = Me.郵便番号.Value & "　" & Me.住所.Value
        .Range("M10").Value = Me.工事名称.Value
        .Range("F10").Value = "工事番号：" & m_KoujiBangou
        .Range("F13").Value = Me.提出要項.Value
        
        .Range("F12").Value = GetSelectedListBoxItems(Me.同封物, DOUFUUBUTSU_DELIMITER_DISPLAY)
        
        .Range("F14").Value = "着手：" & Me.着手.Value
        .Range("J14").Value = "完成：" & Me.完成.Value
        .Range("N14").Value = "引渡日：" & Me.引渡日.Value
        .Range("T14").Value = "提出日付：" & Me.提出日付.Value
        ' 金額や数量の転記（カンマを削除して数値として転記）
        .Range("M16").Value = Val(Me.数量1.Value)
        .Range("N16").Value = Me.単位1.Value
        .Range("P16").Value = Val(Replace(Me.金額1.Value, ",", ""))
        .Range("T16").Value = Val(Replace(Me.金額1.Value, ",", ""))
        .Range("M17").Value = Val(Me.数量2.Value)
        .Range("N17").Value = Me.単位2.Value
        .Range("P17").Value = Val(Replace(Me.金額2.Value, ",", ""))
        .Range("T17").Value = Val(Replace(Me.金額2.Value, ",", ""))
        .Range("M18").Value = Val(Me.数量3.Value)
        .Range("N18").Value = Me.単位3.Value
        .Range("P18").Value = Val(Replace(Me.金額3.Value, ",", ""))
        .Range("T18").Value = Val(Replace(Me.金額3.Value, ",", ""))
        .Range("R19").Value = Val(Replace(Me.小計.Value, ",", ""))
        .Range("R20").Value = Val(Replace(Me.消費税.Value, ",", ""))
        .Range("R21").Value = Val(Replace(Me.請求金額.Value, ",", ""))
        .Range("T19").Value = Val(Replace(Me.小計.Value, ",", ""))
        .Range("T20").Value = Val(Replace(Me.消費税.Value, ",", ""))
        .Range("T21").Value = Val(Replace(Me.請求金額.Value, ",", ""))
        .Range("F22").Value = Me.引継ぎコメント.Value

        ' 作成日を結合セルB31に「yyyy年m月d日」形式で転記
        If IsDate(Me.作成日.Value) Then
            .Range("B30").Value = Format(CDate(Me.作成日.Value), "yyyy年m月d日")
        End If
        
        .Range("Q33").Value = Me.担当者.Value
    End With

    ' --- ④ このツール内の工事一覧シートも最新の状態に更新 ---
    Call UpdateLocalListSheet(wsSrc, wbTarget_Click.Sheets(SHEET_KANRI_MASTER))

    ' --- ⑤ このツール内の「依頼履歴」シートも最新の状態に更新 ---
    On Error Resume Next
    Set wsRirekiLocal = ThisWorkbook.Sheets(SHEET_IRAI_RIREKI)
    On Error GoTo ErrorHandlerClick
    
    If wsRirekiLocal Is Nothing Then
        MsgBox "このファイルに「" & SHEET_IRAI_RIREKI & "」が見つかりませんでした。", vbExclamation
    Else
        Call UpdateLocalRirekiSheet(wsRireki, wsRirekiLocal)
    End If

    isSuccess = True ' 全ての処理が成功

CleanUpClick: ' 終了処理
    If Not wbTarget_Click Is Nothing Then wbTarget_Click.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    If isSuccess Then MsgBox "依頼書と一覧シートを更新しました。", vbInformation, "完了"
    Unload Me
    Exit Sub

ErrorHandlerClick: ' エラー発生時の処理
    Resume CleanUpClick
End Sub

Private Sub 契約無し_Click()

End Sub

Private Sub 請求書記載文言_Change()

End Sub

'================================================================================
' フォーム上のコントロールに対するイベント処理
'================================================================================

' 「税込み」ボタンが押されたときの処理
Private Sub 税込み_Click()
    Dim amount As Currency
    ' 金額1が有効な数値かチェック
    If Not IsNumeric(Replace(Me.金額1.Value, ",", "")) Or Trim(Me.金額1.Value) = "" Then
        MsgBox "金額1に有効な数値を入力してください。", vbExclamation
        Exit Sub
    End If

    ' 金額1を税込み合計として、小計と請求金額に反映する
    amount = CCur(Val(Replace(Me.金額1.Value, ",", "")))
    Me.小計.Value = Format(amount, "#,##0")
    Me.消費税.Value = 0 ' 消費税は0にする
    Me.請求金額.Value = Format(amount, "#,##0")

    ' 他の金額欄はクリアする
    Me.金額2.Value = "": Me.数量2.Value = "": Me.単位2.Value = ""
    Me.金額3.Value = "": Me.数量3.Value = "": Me.単位3.Value = ""
End Sub

' 「請求書提出先」ドロップダウンの入力が変更されたときの処理（オートコンプリート機能）
Private Sub 請求書提出先_Change()
    If m_isComboBoxUpdating Then Exit Sub ' 無限ループ防止
    Dim originalText As String, selStart As Long
    Dim filteredList As Object, item As Variant
    Dim wbTarget_AutoFill As Workbook, wsMaster_AutoFill As Worksheet
    On Error GoTo ErrorHandlerAutoFill

    ' --- ① 入力された文字を含む候補だけにリストを絞り込む ---
    originalText = Me.請求書提出先.Text
    selStart = Me.請求書提出先.selStart
    Set filteredList = CreateObject("Scripting.Dictionary")
    If originalText <> "" And Not IsEmpty(m_CachedSeikyuusakiList) Then
        For Each item In m_CachedSeikyuusakiList
            If InStr(1, item, originalText, vbTextCompare) > 0 Then ' 大文字小文字を区別せず検索
                If Not filteredList.Exists(item) Then filteredList.Add item, True
            End If
        Next item
    End If

    ' --- ② 絞り込んだリストをドロップダウンに再設定 ---
    m_isComboBoxUpdating = True ' これからリストを更新することを示すスイッチON
    Me.請求書提出先.Clear
    If filteredList.count > 0 Then Me.請求書提出先.List = filteredList.Keys
    Me.請求書提出先.Text = originalText ' 入力途中の文字を復元
    Me.請求書提出先.selStart = selStart ' カーソル位置を復元
    If filteredList.count > 0 Then Me.請求書提出先.DropDown ' 候補リストを表示
    m_isComboBoxUpdating = False ' スイッチOFF

    ' --- ③ リストから項目が完全に選択されたら、関連情報を自動入力 ---
    If Me.請求書提出先.ListIndex > -1 Then
        Dim targetFilePath As String
        targetFilePath = GetTargetFilePath()
        Set wbTarget_AutoFill = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=True, UpdateLinks:=0)
        Set wsMaster_AutoFill = wbTarget_AutoFill.Sheets(SHEET_OTHER_MASTER)
        ' マスタから郵便番号や住所などを探して自動入力する
        Call AutoFillFromMaster(Me.請求書提出先.Value, wsMaster_AutoFill)
        If Not wbTarget_AutoFill Is Nothing Then wbTarget_AutoFill.Close SaveChanges:=False
    End If
Exit Sub
ErrorHandlerAutoFill:
    MsgBox "請求書提出先自動入力中にエラー発生: " & Err.Description, vbCritical
End Sub

' 金額欄が変更されたら、合計を自動計算する
Private Sub 金額1_Change(): Call CalculateTotals: End Sub
Private Sub 金額2_Change(): Call CalculateTotals: End Sub
Private Sub 金額3_Change(): Call CalculateTotals: End Sub

' 日付欄からフォーカスが外れたら、書式をチェック・統一する
Private Sub 作成日_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ValidateDate(Me.作成日, "作成日", Cancel): End Sub
Private Sub 着手_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ValidateDate(Me.着手, "工期着手", Cancel): End Sub
Private Sub 完成_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ValidateDate(Me.完成, "工期完成", Cancel): End Sub
Private Sub 引渡日_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ValidateDate(Me.引渡日, "引渡日", Cancel): End Sub
Private Sub 提出日付_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ValidateDate(Me.提出日付, "提出日付", Cancel): End Sub

'================================================================================
' 補助的な関数やサブルーチン群 (Helper Functions)
'================================================================================

' フォームの入力内容を、外部ファイルの「依頼履歴」シートに追記する
Private Sub AddDataToIraiRireki(ByVal wsRireki As Worksheet)
    If wsRireki Is Nothing Then Exit Sub

    wsRireki.Unprotect
    
    Dim nextRow As Long, lastRow As Long
    Dim lastNoStr As String, nextNoNum As Long, newIraiNo As String
    
    ' --- 1. 追記する行番号を決定 (1行目はヘッダー) ---
    lastRow = wsRireki.Cells(wsRireki.Rows.count, "A").End(xlUp).Row
    If lastRow < 1 Then ' シートが完全に空の場合
        nextRow = 2 ' 2行目から書き込む
    ElseIf lastRow = 1 Then ' ヘッダーしかない場合
        nextRow = 2 ' 2行目から書き込む
    Else
        nextRow = lastRow + 1
    End If
    
    ' --- 2. 依頼NOを自動採番 ("0301", "0302"...) ---
    ' 全行を走査して最大の連番を取得する。
    ' Excelが先頭ゼロを自動削除して数値保存する場合（例: "0301"→301）も考慮し
    ' 301?399の数値も旧形式（"03XX"の先頭ゼロ落ち）として検出する。
    Dim i As Long, maxNum As Long, cellVal As String, cellNum As Long, numVal As Long
    maxNum = 0
    For i = 2 To lastRow
        cellVal = Trim(CStr(wsRireki.Cells(i, "A").Value))
        If Left(cellVal, 2) = "03" And IsNumeric(Mid(cellVal, 3)) Then
            ' テキスト形式: "0301", "0302" など
            cellNum = CLng(Mid(cellVal, 3))
            If cellNum > maxNum Then maxNum = cellNum
        ElseIf IsNumeric(cellVal) Then
            ' 数値として保存された旧形式: 301?399（="0301"?"0399"の先頭ゼロ落ち）
            numVal = CLng(cellVal)
            If numVal >= 301 And numVal <= 399 Then
                cellNum = numVal - 300
                If cellNum > maxNum Then maxNum = cellNum
            End If
        End If
    Next i
    nextNoNum = maxNum + 1
    newIraiNo = "03" & Format(nextNoNum, "00")

    ' --- 3. データを書き込む ---
    With wsRireki
        .Cells(nextRow, "A").Value = newIraiNo ' 依頼NO (A列)
        .Cells(nextRow, "B").Value = "" ' 請求書番号（経理用） (B列)
        .Cells(nextRow, "C").Value = "" ' 請求書発行日 (C列)
        .Cells(nextRow, "D").Value = Me.担当者.Value ' 1:担当者 (D列)
        .Cells(nextRow, "E").Value = m_KoujiBangou ' 2:工事番号 (E列)
        .Cells(nextRow, "F").Value = Me.工事名称.Value ' 3:工事名称 (F列)
        .Cells(nextRow, "G").Value = FormatIfDate(Me.着手.Value) ' 4:工期 着手 (G列)
        .Cells(nextRow, "H").Value = FormatIfDate(Me.完成.Value) ' 5:工期 完成 (H列)
        .Cells(nextRow, "I").Value = Val(Replace(Me.請求金額.Value, ",", "")) ' 6:請負金額 (I列)
        .Cells(nextRow, "J").Value = FormatIfDate(Me.作成日.Value) ' 7:依頼書作成日 (J列)
        .Cells(nextRow, "K").Value = Me.提出要項.Value ' 8:提出要項 (K列)
        
        .Cells(nextRow, "L").Value = Me.提出要項その他.Value ' 9:提出要項(その他) (L列)
        
        .Cells(nextRow, "M").Value = GetSelectedListBoxItems(Me.同封物, DOUFUUBUTSU_DELIMITER_SAVE) ' 10:同封物 (M列)
        .Cells(nextRow, "N").Value = FormatIfDate(Me.提出日付.Value) ' 11:提出日付 (N列)
        .Cells(nextRow, "O").Value = FormatIfDate(Me.引渡日.Value) ' 12:引渡日 (O列)
        .Cells(nextRow, "P").Value = Me.請求書提出先.Value ' 13:請求書提出先 (P列)
        .Cells(nextRow, "Q").Value = Me.郵便番号.Value ' 14:郵便番号 (Q列)
        .Cells(nextRow, "R").Value = Me.住所.Value ' 15:提出先住所 (R列)
        
        .Cells(nextRow, "S").Value = Me.請求書記載文言.Value ' 16:請求書記載文言 (S列)
        
        .Cells(nextRow, "T").Value = Me.引継ぎコメント.Value ' 17:担当者引継ぎコメント (T列)
        
        If Me.領収書注意文.Value = True Then ' 18:領収書注意文 (U列)
            .Cells(nextRow, "U").Value = "有"
        Else
            .Cells(nextRow, "U").Value = ""
        End If
        If Me.振込手数料注意文.Value = True Then ' 19:振込手数料注意文 (V列)
            .Cells(nextRow, "V").Value = "有"
        Else
            .Cells(nextRow, "V").Value = ""
        End If
        If Me.但陽信金口座指定.Value = True Then ' 20:但陽信金口座指定 (W列)
            .Cells(nextRow, "W").Value = "有"
        Else
            .Cells(nextRow, "W").Value = ""
        End If
    End With

    ' PDF作成時にR3へ依頼NOを転記できるよう、請求書提出依頼書シートに書き込む
    Dim wsReq As Worksheet
    On Error Resume Next
    Set wsReq = ThisWorkbook.Sheets("請求書提出依頼書")
    On Error GoTo 0
    If Not wsReq Is Nothing Then wsReq.Range("R3").Value = newIraiNo
End Sub


' 日付の形式が正しいかチェックし、自動で書式を整える
Private Sub ValidateDate(ByVal DateField As MSForms.Control, ByVal FieldName As String, ByRef Cancel As MSForms.ReturnBoolean)
    Dim inputText As String
    inputText = Trim(DateField.Value)
    If inputText = "" Then Exit Sub
    If IsDate(inputText) Then
        DateField.Value = Format(CDate(inputText), "yyyy/mm/dd")
    Else
        MsgBox FieldName & " は「YYYY/MM/DD」形式で入力してください。", vbExclamation, "入力エラー"
        Cancel = True
    End If
End Sub

' 金額1～3を合計し、小計・消費税・請求金額を計算する
Private Sub CalculateTotals()
    Dim val1 As Currency, val2 As Currency, val3 As Currency, subTotal As Currency, tax As Currency, grandTotal As Currency
    val1 = Val(Replace(Me.金額1.Value, ",", ""))
    val2 = Val(Replace(Me.金額2.Value, ",", ""))
    val3 = Val(Replace(Me.金額3.Value, ",", ""))
    subTotal = val1 + val2 + val3
    tax = Application.WorksheetFunction.Round(subTotal * 0.1, 0) ' 消費税10%を計算（四捨五入）
    grandTotal = subTotal + tax
    Me.小計.Value = Format(subTotal, "#,##0")
    Me.消費税.Value = Format(tax, "#,##0")
    Me.請求金額.Value = Format(grandTotal, "#,##0")
End Sub

' フォームの入力内容を、外部ファイルの「工事番号一覧」シートに書き込む
Private Sub UpdateExternalFile(ByVal wsTarget As Worksheet, ByVal rowToUpdate As Long)
    If wsTarget Is Nothing Or rowToUpdate = 0 Then Exit Sub

    wsTarget.Unprotect

    With wsTarget
        .Cells(rowToUpdate, "S").Value = Me.請求書提出先.Value
        .Cells(rowToUpdate, "G").Value = FormatIfDate(Me.着手.Value)
        .Cells(rowToUpdate, "H").Value = FormatIfDate(Me.完成.Value)
        .Cells(rowToUpdate, "K").Value = Val(Replace(Me.小計.Value, ",", ""))
        .Cells(rowToUpdate, "C").Value = Me.担当者.Value
        .Cells(rowToUpdate, "Q").Value = FormatIfDate(Me.提出日付.Value)
        .Cells(rowToUpdate, "O").Value = Me.提出要項.Value
        
        .Cells(rowToUpdate, "P").Value = GetSelectedListBoxItems(Me.同封物, DOUFUUBUTSU_DELIMITER_SAVE)
        
        .Cells(rowToUpdate, "R").Value = FormatIfDate(Me.引渡日.Value)
        .Cells(rowToUpdate, "T").Value = Me.郵便番号.Value
        .Cells(rowToUpdate, "U").Value = Me.住所.Value
        .Cells(rowToUpdate, "X").Value = Me.引継ぎコメント.Value
        .Cells(rowToUpdate, "N").Value = Date ' 最終更新日
    End With
End Sub

' フォームの入力内容で、外部ファイルの「その他マスタ」を更新する
Private Sub UpdateAddressMaster(ByVal wsMaster As Worksheet)
    If wsMaster Is Nothing Then Exit Sub

    wsMaster.Unprotect

    Dim searchVal As String, foundCell As Range
    searchVal = Trim(Me.請求書提出先.Value)
    If searchVal <> "" Then
        ' マスタに同じ請求先があるか探す
        Set foundCell = wsMaster.Columns(MASTER_OTHER_SEIKYUUSAKI_COL).Find(What:=searchVal, LookIn:=xlValues, LookAt:=xlWhole)
        If foundCell Is Nothing Then
            ' 見つからない場合：新しいデータとして最終行に追加
            Dim lastRow As Long
            lastRow = wsMaster.Cells(wsMaster.Rows.count, MASTER_OTHER_SEIKYUUSAKI_COL).End(xlUp).Row + 1
            wsMaster.Cells(lastRow, MASTER_OTHER_SEIKYUUSAKI_COL).Value = searchVal
            wsMaster.Cells(lastRow, MASTER_OTHER_YUBIN_NO_COL).Value = Me.郵便番号.Value
            wsMaster.Cells(lastRow, MASTER_OTHER_JUSHO_COL).Value = Me.住所.Value
        Else
            ' 見つかった場合：既存のデータを最新の情報で上書き
            wsMaster.Cells(foundCell.Row, MASTER_OTHER_YUBIN_NO_COL).Value = Me.郵便番号.Value
            wsMaster.Cells(foundCell.Row, MASTER_OTHER_JUSHO_COL).Value = Me.住所.Value
        
        End If
    End If

    ' 「その他マスタ」は重要なデータなので、最後に保護をかける
    wsMaster.Protect
End Sub

' 「その他マスタ」から選択された請求先に紐づく情報を探し、フォームに自動入力する
Private Sub AutoFillFromMaster(ByVal selectedValue As String, ByVal wsMaster As Worksheet)
    If selectedValue = "" Or wsMaster Is Nothing Then Exit Sub
    Dim foundCell As Range
    On Error GoTo CleanUpAutoFill
    Set foundCell = wsMaster.Columns(MASTER_OTHER_SEIKYUUSAKI_COL).Find(What:=selectedValue, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        m_isComboBoxUpdating = True ' 無限ループ防止
        Me.郵便番号.Value = wsMaster.Cells(foundCell.Row, MASTER_OTHER_YUBIN_NO_COL).Value
        Me.住所.Value = wsMaster.Cells(foundCell.Row, MASTER_OTHER_JUSHO_COL).Value
        m_isComboBoxUpdating = False
    End If
CleanUpAutoFill:
    Exit Sub
End Sub

' キャッシュデータ（配列）をドロップダウンリストに設定する
Private Sub PopulateComboBoxFromCache(ByVal cmb As MSForms.ComboBox, ByVal cacheList As Variant)
    If IsArray(cacheList) Then
        cmb.List = cacheList
    Else
        cmb.Clear
    End If
End Sub

' キャッシュデータ（配列）をリストボックスに設定する
Private Sub PopulateListBoxFromCache(ByVal lb As MSForms.ListBox, ByVal cacheList As Variant)
    If IsArray(cacheList) Then
        lb.List = cacheList
    Else
        lb.Clear
    End If
End Sub


' シートの指定された列のデータを配列として取得する
Private Function GetColumnData(ByVal ws As Worksheet, ByVal col As String, ByVal startRow As Long) As Variant
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
    If lastRow >= startRow Then
        GetColumnData = ws.Range(col & startRow & ":" & col & lastRow).Value
    Else
        GetColumnData = Array() ' データがない場合は空の配列を返す
    End If
End Function

' ドロップダウンリストに値を設定する（リストにない場合は追加してから設定）
Private Sub SetComboBoxValue(ByVal cmb As MSForms.ComboBox, ByVal valueToSet As String)
    Dim i As Long, found As Boolean
    If Trim(valueToSet) = "" Then Exit Sub
    found = False
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = valueToSet Then
            found = True
            Exit For
        End If
    Next i
    If found Then
        cmb.Value = valueToSet
    Else
        cmb.AddItem valueToSet
        cmb.Value = valueToSet
    End If
End Sub

' リストボックスの選択項目を、区切り文字で連結した1つの文字列として取得する
Private Function GetSelectedListBoxItems(ByVal lb As MSForms.ListBox, ByVal delimiter As String) As String
    Dim result As String
    Dim i As Long
    result = ""
    
    If lb.ListCount = 0 Then
        GetSelectedListBoxItems = ""
        Exit Function
    End If
    
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then
            If result = "" Then
                result = lb.List(i)
            Else
                result = result & delimiter & lb.List(i)
            End If
        End If
    Next i
    
    GetSelectedListBoxItems = result
End Function

' 区切り文字で連結された文字列を元に、リストボックスの項目を選択状態にする
Private Sub SetListBoxSelection(ByVal lb As MSForms.ListBox, ByVal savedString As String, ByVal delimiter As String)
    Dim i As Long, j As Long
    Dim selectedArray As Variant
    
    ' 0. 念のため、全ての選択を解除
    If lb.ListCount > 0 Then
        For i = 0 To lb.ListCount - 1
            lb.Selected(i) = False
        Next i
    End If
    
    If Trim(savedString) = "" Then Exit Sub
    
    ' 1. 保存されている文字列を配列に分割
    selectedArray = Split(savedString, delimiter)
    
    ' 2. リストボックスの各項目をチェック
    If lb.ListCount > 0 Then
        For i = 0 To lb.ListCount - 1
            ' 3. 分割した配列の各項目をチェック
            For j = 0 To UBound(selectedArray)
                If Trim(selectedArray(j)) = lb.List(i) Then
                    lb.Selected(i) = True ' 一致したら選択
                    Exit For ' 内側ループを抜ける
                End If
            Next j
        Next i
    End If
End Sub


' シートの特定行からデータを読み込み、フォームの各項目に表示する処理
Private Sub LoadDataToForm(ByVal ws As Worksheet, ByVal rowNum As Long)
    m_isComboBoxUpdating = True ' 自動入力中はChangeイベントを無効化
    With ws
        Call SetComboBoxValue(Me.請求書提出先, .Cells(rowNum, "S").Value)
        Me.着手.Value = FormatIfDate(.Cells(rowNum, "G").Value)
        Me.完成.Value = FormatIfDate(.Cells(rowNum, "H").Value)
        Me.金額1.Value = .Cells(rowNum, "K").Value
        Me.担当者.Value = .Cells(rowNum, "C").Value
        Me.提出日付.Value = FormatIfDate(.Cells(rowNum, "Q").Value)
        Call SetComboBoxValue(Me.提出要項, .Cells(rowNum, "O").Value)
        
        Call SetListBoxSelection(Me.同封物, CStr(.Cells(rowNum, "P").Value), DOUFUUBUTSU_DELIMITER_SAVE)
        
        Me.引渡日.Value = FormatIfDate(.Cells(rowNum, "R").Value)
        Me.郵便番号.Value = .Cells(rowNum, "T").Value
        Me.住所.Value = .Cells(rowNum, "U").Value
        Me.引継ぎコメント.Value = .Cells(rowNum, "X").Value
    End With
    Call CalculateTotals ' 金額を再計算
    m_isComboBoxUpdating = False ' Changeイベントを有効に戻す
End Sub

' 値が日付なら"yyyy/mm/dd"形式の文字列に、そうでなければ空文字に変換する関数
Private Function FormatIfDate(ByVal Value As Variant) As Variant
    If IsDate(Value) Then
        FormatIfDate = Format(CDate(Value), "yyyy/mm/dd")
    Else
        FormatIfDate = ""
    End If
End Function

' このツール内の「工事一覧」シートを、外部ファイルの最新情報に更新する処理
Private Sub UpdateLocalListSheet(ByVal wsSource As Worksheet, ByVal wsMaster As Worksheet)
    If wsSource Is Nothing Or wsMaster Is Nothing Then
        MsgBox "更新元またはマスタシートの参照が不正なため、ローカルシートの更新を中断しました。", vbCritical, "引数エラー"
        Exit Sub
    End If
    Dim wsDest As Worksheet, destSheetName As String, lastRowSource As Long
    Dim copyRange As Range
    ' Dim dataArray As Variant, c As Long, currentColumnWidth As Double ' ← 不要になるためコメントアウトまたは削除
    Dim originalDisplayAlerts As Boolean, originalEnableEvents As Boolean
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    On Error GoTo ErrorHandlerUpdateLocal
    destSheetName = Trim(CStr(wsMaster.Range(CELL_LOCAL_COPY_SHEET).Value))
    If destSheetName = "" Then
        MsgBox "「" & SHEET_KANRI_MASTER & "」" & CELL_LOCAL_COPY_SHEET & "セルにシート名が指定されていません。", vbExclamation
        GoTo FinalizeUpdateLocal
    End If
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets(destSheetName)
    On Error GoTo ErrorHandlerUpdateLocal
    If wsDest Is Nothing Then
        MsgBox "このファイルに「" & destSheetName & "」が見つかりませんでした。", vbExclamation
        GoTo FinalizeUpdateLocal
    End If

    ' --- データのコピー処理 ---
    wsDest.Unprotect
    ' .ClearContents から .Clear に変更し、古い書式もクリアする
    wsDest.Range("A3:X" & wsDest.Rows.count).Clear
    
    lastRowSource = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    If lastRowSource >= 5 Then
        Set copyRange = wsSource.Range("A5:X" & lastRowSource)
        
        '--- 新しい「書式ごと」コピーする処理 ---
        copyRange.Copy Destination:=wsDest.Range("A3")
        
    End If
FinalizeUpdateLocal: ' 終了処理
    Application.CutCopyMode = False
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Exit Sub
ErrorHandlerUpdateLocal: ' エラー発生時の処理
    MsgBox "ローカルシート更新中に予期せぬエラー発生: " & Err.Description, vbCritical
    Resume FinalizeUpdateLocal
End Sub


' このツール内の「依頼履歴」シートを、外部ファイルの最新情報に更新する処理
Private Sub UpdateLocalRirekiSheet(ByVal wsSource As Worksheet, ByVal wsDest As Worksheet)
    If wsSource Is Nothing Or wsDest Is Nothing Then
        MsgBox "依頼履歴の更新元または更新先シートの参照が不正なため、ローカルシートの更新を中断しました。", vbCritical, "引数エラー"
        Exit Sub
    End If
    
    Dim lastRowSource As Long
    Dim copyRange As Range
    Dim originalDisplayAlerts As Boolean, originalEnableEvents As Boolean
    
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandlerUpdateLocalRireki
    
    ' --- データのコピー処理 ---
    wsDest.Unprotect
    
    ' 3行目以降（ヘッダー(2行目)を除く）のデータをクリア（書式ごと）
    wsDest.Range("A3:W" & wsDest.Rows.count).Clear
    
    ' 外部ファイルのデータ最終行を取得（1行目はヘッダーと仮定）
    lastRowSource = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    
    ' 外部ファイルにデータが2行目以降にある場合のみコピー
    If lastRowSource >= 2 Then
        ' コピー範囲はA2からW列の最終行まで（列構成に合わせて"W"に変更）
        Set copyRange = wsSource.Range("A2:W" & lastRowSource)
        
        ' ローカルシートのA3に「書式ごと」コピーする
        copyRange.Copy Destination:=wsDest.Range("A3")
    End If
    
FinalizeUpdateLocalRireki: ' 終了処理
    Application.CutCopyMode = False
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Exit Sub
    
ErrorHandlerUpdateLocalRireki: ' エラー発生時の処理
    MsgBox "ローカルの依頼履歴シート更新中に予期せぬエラー発生: " & Err.Description, vbCritical
    Resume FinalizeUpdateLocalRireki
End Sub


' 補足：シート存在チェック用の関数
Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Sub 提出要項_Change()

End Sub

Private Sub 提出要項その他_Change()

End Sub


Private Sub 提出要項_Click()

End Sub
