Attribute VB_Name = "PDF文字抽出_基本"
Option Explicit

' =====================================================================
'  PDF文字抽出_基本
'  ExcelブックのVBAからPDFの文字を抽出する最小構成のサンプルです。
'  doc-check-tool（GUI版・PDF.js）と同じ「値を抜き出して照合する」考え方を
'  ExcelマクロだけでもできるようにしたVBA版の土台です。
'
'  事前準備:
'   ・PC本体にMicrosoft Wordがインストールされていること
'     （WordはPDFを開くと自動でテキストへ変換してくれる機能があるため、
'      それを使って抽出しています。追加ライブラリのインストールは不要）
'   ・参照設定は不要（CreateObjectによる遅延バインディングのため）
' =====================================================================


' ---------------------------------------------------------------------
' PDFファイル全体をテキストとして抽出する
' ---------------------------------------------------------------------
Function ExtractTextFromPDF(pdfPath As String) As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim resultText As String

    On Error GoTo ErrHandler

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

    ' WordはPDFを開くと自動的にテキストへ変換して開いてくれる
    Set wordDoc = wordApp.Documents.Open(FileName:=pdfPath, ReadOnly:=True, ConfirmConversions:=False)

    resultText = wordDoc.Content.Text

    wordDoc.Close SaveChanges:=False
    wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing

    ExtractTextFromPDF = resultText
    Exit Function

ErrHandler:
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    ExtractTextFromPDF = "ERROR: " & Err.Description
End Function


' ---------------------------------------------------------------------
' 抽出した全文から、キーワードの近くにある値だけを正規表現で抜き出す
' 例: pattern = "数量[：:]\s*([0-9,]+)"  → 「数量：」の直後の数字を取得
' ---------------------------------------------------------------------
Function ExtractValueByKeyword(sourceText As String, pattern As String) As String
    Dim regex As Object
    Dim matches As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.Global = False
    regex.IgnoreCase = True

    If regex.Test(sourceText) Then
        Set matches = regex.Execute(sourceText)
        If matches(0).SubMatches.Count > 0 Then
            ExtractValueByKeyword = Trim(matches(0).SubMatches(0))
        Else
            ExtractValueByKeyword = Trim(matches(0).Value)
        End If
    Else
        ExtractValueByKeyword = ""
    End If
End Function


' ---------------------------------------------------------------------
' 実行サンプル
' アクティブシートに以下の列を用意しておくと、一括で抽出できます。
'   A列: PDFファイルのフルパス（2行目から）
'   B列: 抽出したい項目の正規表現パターン
'   C列: （実行後）抽出結果が入る
' ---------------------------------------------------------------------
Sub Run_ExtractSample()
    Dim ws As Worksheet
    Dim i As Long
    Dim pdfPath As String
    Dim pattern As String
    Dim fullText As String

    Set ws = ActiveSheet
    i = 2 ' 1行目はヘッダー想定

    Do While ws.Cells(i, 1).Value <> ""
        pdfPath = ws.Cells(i, 1).Value
        pattern = ws.Cells(i, 2).Value

        fullText = ExtractTextFromPDF(pdfPath)

        If Left(fullText, 6) = "ERROR:" Then
            ws.Cells(i, 3).Value = fullText
        Else
            ws.Cells(i, 3).Value = ExtractValueByKeyword(fullText, pattern)
        End If

        i = i + 1
    Loop

    MsgBox "抽出処理が完了しました。"
End Sub


' ---------------------------------------------------------------------
' 2つの値を比較して判定を返す（doc-check-toolの判定ロジックと同じ考え方）
' compareType: "exact" / "numeric" / "contains"
' ---------------------------------------------------------------------
Function JudgeValues(valueA As String, valueB As String, compareType As String) As String
    Dim a As String, b As String
    a = Replace(Replace(Trim(valueA), " ", ""), "　", "")
    b = Replace(Replace(Trim(valueB), " ", ""), "　", "")

    If a = "" Or b = "" Then
        JudgeValues = "要確認"
        Exit Function
    End If

    Select Case compareType
        Case "numeric"
            If IsNumeric(a) And IsNumeric(b) Then
                JudgeValues = IIf(CDbl(a) = CDbl(b), "一致", "不一致")
            Else
                JudgeValues = "要確認"
            End If
        Case "contains"
            JudgeValues = IIf(InStr(a, b) > 0 Or InStr(b, a) > 0, "一致", "不一致")
        Case Else ' exact
            JudgeValues = IIf(a = b, "一致", "不一致")
    End Select
End Function
