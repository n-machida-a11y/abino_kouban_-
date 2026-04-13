Attribute VB_Name = "PDF作成"
Option Explicit

' ファイル名に使えない文字を置換する補助的な関数
Private Function SanitizeFileName(ByVal fileName As String) As String
    Dim illegalChars As String, i As Long
    illegalChars = "\/:*?""<>|"
    SanitizeFileName = fileName
    For i = 1 To Len(illegalChars)
        SanitizeFileName = Replace(SanitizeFileName, Mid(illegalChars, i, 1), "_")
    Next i
End Function

Public Sub SaveRequestFormAsPDF()
    Dim wsRequest As Worksheet
    Dim saveFolder As String
    Dim pdfFileName As String
    Dim fullPath As String
    Dim recipient As String
    Dim KoujiName As String
    
    '--- ① エクスポート対象のシートを指定 ---
    On Error Resume Next
    Set wsRequest = ThisWorkbook.Sheets("請求書提出依頼書")
    On Error GoTo 0
    
    If wsRequest Is Nothing Then
        MsgBox "「請求書提出依頼書」シートが見つかりません。", vbCritical
        Exit Sub
    End If
    
    '--- 保存先フォルダ ---
    ' USERPROFILE が実際のドライブと異なる環境に対応するため
    ' Z:\Users\ユーザー名\Downloads を優先して使用する
      Dim userName As String
      userName = Environ("USERNAME")
      If Dir("Z:\Users\" & userName & "\Downloads", vbDirectory) <> "" Then
          saveFolder = "Z:\Users\" & userName & "\Downloads"
      Else
          saveFolder = Environ("USERPROFILE") & "\Downloads"
      End If
    '--- ③ PDFのファイル名をシートのセルから作成 ---
    recipient = wsRequest.Range("F7").Value
    KoujiName = wsRequest.Range("M10").Value
    
    If Trim(recipient) = "" Or Trim(KoujiName) = "" Then
        MsgBox "ファイル名の作成に必要な情報（F7セル:請求書提出先、M10セル:工事名称）がシートに見つかりません。", vbExclamation
        Exit Sub
    End If
    
    pdfFileName = recipient & "_" & KoujiName & "_" & Format(Now, "yyyymmdd")
    pdfFileName = SanitizeFileName(pdfFileName)
    fullPath = saveFolder & "\" & pdfFileName & ".pdf"

    '--- ④ PDFとしてエクスポート ---
    Dim exportError As String
    
    wsRequest.Activate
    
    Application.ScreenUpdating = True
    On Error Resume Next
    wsRequest.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fullPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    If Err.Number <> 0 Then exportError = Err.Description
    On Error GoTo 0

    '--- ⑤ ファイルが実際に存在するか最終確認 ---
    If Dir(fullPath) <> "" Then
        ' 修正箇所: 保存先を示すメッセージを変更
        MsgBox "PDFをダウンロードフォルダに保存しました。"
    Else
        MsgBox "PDFの作成に失敗しました。" & vbCrLf & "エラー内容: " & exportError, vbCritical
    End If
    
End Sub

