Attribute VB_Name = "PDF文字抽出_Excel汎用版"
Option Explicit

' =====================================================================
'  PDF文字抽出_Excel汎用版
'  「設定」シートに書類の組み合わせを何行でも登録しておき、
'  Run_CheckAll を実行すると「結果」シートに一括で照合結果を出力します。
'
'  設定シートの列構成（2行目がヘッダー、3行目はサンプル、4行目以降が入力欄）:
'   A:グループ名  B:書類1名称  C:書類1パス  D:書類1抽出パターン
'   E:書類2名称  F:書類2パス  G:書類2抽出パターン
'   H:書類3名称  I:書類3パス  J:書類3抽出パターン（任意・空欄可）
'   K:照合タイプ（exact / numeric / date / contains）
'
'  事前準備: PCにMicrosoft Wordがインストールされていること（参照設定は不要）
' =====================================================================

Dim textCache As Object ' PDFパス→全文テキスト のキャッシュ（同じPDFを何度も開かないため）

' ---------------------------------------------------------------------
' メイン処理：設定シートを1行ずつ読み込み、結果シートに書き出す
' ---------------------------------------------------------------------
Sub Run_CheckAll()
    Dim wsSet As Worksheet, wsRes As Worksheet
    Dim r As Long, outRow As Long
    Dim groupName As String
    Dim docName(1 To 3) As String, docPath(1 To 3) As String, docPattern(1 To 3) As String
    Dim compareType As String
    Dim values() As String
    Dim slotCount As Integer, i As Integer
    Dim verdict As String

    Set wsSet = ThisWorkbook.Sheets("設定")
    Set wsRes = ThisWorkbook.Sheets("結果")
    Set textCache = CreateObject("Scripting.Dictionary")

    ' 結果シートをクリア（ヘッダー行は残す）
    Dim lastResRow As Long
    lastResRow = wsRes.Cells(wsRes.Rows.Count, 1).End(xlUp).Row
    If lastResRow > 1 Then wsRes.Range("A2:F" & lastResRow).ClearContents

    outRow = 2
    r = 4 ' 4行目から入力データ（2行目ヘッダー、3行目サンプル）

    Do While wsSet.Cells(r, 1).Value <> ""
        groupName = wsSet.Cells(r, 1).Value
        docName(1) = wsSet.Cells(r, 2).Value: docPath(1) = wsSet.Cells(r, 3).Value: docPattern(1) = wsSet.Cells(r, 4).Value
        docName(2) = wsSet.Cells(r, 5).Value: docPath(2) = wsSet.Cells(r, 6).Value: docPattern(2) = wsSet.Cells(r, 7).Value
        docName(3) = wsSet.Cells(r, 8).Value: docPath(3) = wsSet.Cells(r, 9).Value: docPattern(3) = wsSet.Cells(r, 10).Value
        compareType = LCase(Trim(wsSet.Cells(r, 11).Value))
        If compareType = "" Then compareType = "exact"

        ' 使用するスロット数（書類3が空なら2件で比較）
        slotCount = 2
        If Trim(docPath(3)) <> "" Then slotCount = 3

        ReDim values(1 To slotCount)
        For i = 1 To slotCount
            values(i) = GetValueForDoc(docPath(i), docPattern(i))
        Next i

        verdict = JudgeGroup(values, compareType)

        wsRes.Cells(outRow, 1).Value = groupName
        wsRes.Cells(outRow, 2).Value = values(1)
        wsRes.Cells(outRow, 3).Value = values(2)
        If slotCount = 3 Then wsRes.Cells(outRow, 4).Value = values(3)
        wsRes.Cells(outRow, 5).Value = verdict
        wsRes.Cells(outRow, 6).Value = Now

        ' 判定に応じて色付け
        Select Case verdict
            Case "一致": wsRes.Cells(outRow, 5).Interior.Color = RGB(224, 240, 226)
            Case "不一致": wsRes.Cells(outRow, 5).Interior.Color = RGB(248, 222, 218)
            Case Else: wsRes.Cells(outRow, 5).Interior.Color = RGB(250, 235, 204)
        End Select

        outRow = outRow + 1
        r = r + 1
    Loop

    MsgBox "照合が完了しました（" & (outRow - 2) & "件）。「結果」シートを確認してください。"
End Sub


' ---------------------------------------------------------------------
' PDFパス＋抽出パターンから値を1件取得する（全文はキャッシュして使い回す）
' ---------------------------------------------------------------------
Function GetValueForDoc(pdfPath As String, pattern As String) As String
    Dim fullText As String

    If Trim(pdfPath) = "" Then
        GetValueForDoc = ""
        Exit Function
    End If

    If textCache.Exists(pdfPath) Then
        fullText = textCache(pdfPath)
    Else
        fullText = ExtractTextFromPDF(pdfPath)
        textCache.Add pdfPath, fullText
    End If

    If Left(fullText, 6) = "ERROR:" Then
        GetValueForDoc = fullText
    Else
        GetValueForDoc = ExtractValueByKeyword(fullText, pattern)
    End If
End Function


' ---------------------------------------------------------------------
' PDFファイル全体をテキストとして抽出する（Wordの変換機能を利用）
' ---------------------------------------------------------------------
Function ExtractTextFromPDF(pdfPath As String) As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim resultText As String

    On Error GoTo ErrHandler

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

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
' 抽出した全文から、キーワード近くの値を正規表現で抜き出す
' ---------------------------------------------------------------------
Function ExtractValueByKeyword(sourceText As String, pattern As String) As String
    Dim regex As Object, matches As Object

    If Trim(pattern) = "" Then
        ExtractValueByKeyword = ""
        Exit Function
    End If

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
' 複数の値（2〜3件）をまとめて比較し、"一致"/"不一致"/"要確認" を返す
' ---------------------------------------------------------------------
Function JudgeGroup(values() As String, compareType As String) As String
    Dim i As Integer
    Dim base As String
    Dim result As String

    For i = LBound(values) To UBound(values)
        If Trim(values(i)) = "" Or Left(values(i), 6) = "ERROR:" Then
            JudgeGroup = "要確認"
            Exit Function
        End If
    Next i

    base = values(LBound(values))
    For i = LBound(values) + 1 To UBound(values)
        result = JudgeValues(base, values(i), compareType)
        If result = "要確認" Then
            JudgeGroup = "要確認"
            Exit Function
        ElseIf result = "不一致" Then
            JudgeGroup = "不一致"
            Exit Function
        End If
    Next i

    JudgeGroup = "一致"
End Function


' ---------------------------------------------------------------------
' 2つの値を比較して判定を返す
' compareType: "exact" / "numeric" / "date" / "contains"
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
        Case "date"
            a = Replace(Replace(a, "年", "/"), "月", "/")
            a = Replace(a, "日", "")
            b = Replace(Replace(b, "年", "/"), "月", "/")
            b = Replace(b, "日", "")
            JudgeValues = IIf(a = b, "一致", "不一致")
        Case "contains"
            JudgeValues = IIf(InStr(a, b) > 0 Or InStr(b, a) > 0, "一致", "不一致")
        Case Else ' exact
            JudgeValues = IIf(a = b, "一致", "不一致")
    End Select
End Function
