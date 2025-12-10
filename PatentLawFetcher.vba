' PatentLawFetcher.vba
' e-Gov法令APIから特許法（昭和三十四年法律第百二十一号）の条文を取得し、
' Microsoft Excelワークシートに書き出すためのVBAマクロです。
'
' 使用方法:
' 1. Excelを開き、Alt+F11でVBAエディタを開く
' 2. 挿入 > 標準モジュール を選択
' 3. このコードを貼り付ける
' 4. ツール > 参照設定 で「Microsoft XML, v6.0」にチェックを入れる
' 5. Alt+F8でマクロ一覧を開き、「FetchAndWritePatentLaw」を実行

Sub FetchAndWritePatentLaw()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim lawNum As String
    Dim url As String
    Dim xmlHttp As Object
    Dim xmlDoc As Object
    Dim lawBody As Object
    Dim mainProvision As Object
    Dim articles As Object
    Dim article As Object
    Dim rowNum As Long

    On Error GoTo ErrorHandler

    sheetName = "特許法"
    lawNum = "334AC0000000121" ' 特許法の法令番号
    url = "https://laws.e-gov.go.jp/api/1/lawdata/" & lawNum

    ' シートの準備
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
        Debug.Print "シート「" & sheetName & "」を新規作成しました。"
    Else
        ws.Cells.Clear
        Debug.Print "シート「" & sheetName & "」をクリアしました。"
    End If

    ' ステータスバーに表示
    Application.StatusBar = "法令APIからデータの取得を開始します..."

    ' HTTPリクエストを送信
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP.6.0")
    xmlHttp.Open "GET", url, False
    xmlHttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    xmlHttp.send

    If xmlHttp.Status <> 200 Then
        MsgBox "APIからのデータ取得に失敗しました。ステータスコード: " & xmlHttp.Status, vbCritical
        Exit Sub
    End If

    Debug.Print "APIからのデータ取得に成功しました。"
    Application.StatusBar = "データの解析と書き込みを開始します..."

    ' XMLデータを解析
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.LoadXML xmlHttp.responseText

    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "XMLの解析に失敗しました: " & xmlDoc.parseError.reason, vbCritical
        Exit Sub
    End If

    ' LawBody要素を取得
    Set lawBody = GetLawBody(xmlDoc)
    If lawBody Is Nothing Then
        MsgBox "LawBody要素が見つかりませんでした。", vbCritical
        Exit Sub
    End If

    ' MainProvision要素を取得
    Set mainProvision = lawBody.SelectSingleNode(".//*[local-name()='MainProvision']")
    If mainProvision Is Nothing Then
        MsgBox "MainProvision要素が見つかりませんでした。", vbCritical
        Exit Sub
    End If

    ' Article要素を全て取得
    Set articles = mainProvision.SelectNodes(".//*[local-name()='Article']")
    If articles.Length = 0 Then
        MsgBox "Article要素が見つかりませんでした。", vbCritical
        Exit Sub
    End If

    Debug.Print "取得した条文数: " & articles.Length

    ' ヘッダーを書き込み
    ws.Cells(1, 1).Value = "条番号"
    ws.Cells(1, 2).Value = "条文本文"
    ws.Range("A1:B1").Font.Bold = True

    ' 条文データを書き込み
    rowNum = 2
    For Each article In articles
        ws.Cells(rowNum, 1).Value = GetArticleNumber(article)
        ws.Cells(rowNum, 2).Value = GetArticleContent(article)
        ws.Cells(rowNum, 2).WrapText = True
        rowNum = rowNum + 1
    Next article

    ' 列幅を調整
    ws.Columns(1).AutoFit
    ws.Columns(2).ColumnWidth = 80

    Application.StatusBar = False
    Debug.Print "スプレッドシートへの書き込みが完了しました。"
    MsgBox "特許法の条文をシートに書き出しました。" & vbCrLf & "条文数: " & articles.Length, vbInformation

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Debug.Print "エラー: " & Err.Description
End Sub

' LawBody要素を取得する関数
Function GetLawBody(xmlDoc As Object) As Object
    Dim applData As Object
    Dim lawFullText As Object
    Dim law As Object
    Dim lawBody As Object

    ' DataRoot > ApplData > LawFullText > Law > LawBody の構造を試す
    Set applData = xmlDoc.SelectSingleNode("//*[local-name()='ApplData']")
    If Not applData Is Nothing Then
        Set lawFullText = applData.SelectSingleNode(".//*[local-name()='LawFullText']")
        If Not lawFullText Is Nothing Then
            Set law = lawFullText.SelectSingleNode(".//*[local-name()='Law']")
            If Not law Is Nothing Then
                Set lawBody = law.SelectSingleNode(".//*[local-name()='LawBody']")
                If Not lawBody Is Nothing Then
                    Set GetLawBody = lawBody
                    Exit Function
                End If
            End If
        End If
    End If

    ' 直接LawBody要素を探す
    Set lawBody = xmlDoc.SelectSingleNode("//*[local-name()='LawBody']")
    Set GetLawBody = lawBody
End Function

' 条番号を取得する関数
Function GetArticleNumber(article As Object) As String
    Dim numAttr As Object
    Set numAttr = article.Attributes.getNamedItem("Num")
    If Not numAttr Is Nothing Then
        GetArticleNumber = numAttr.Text
    Else
        GetArticleNumber = "不明"
    End If
End Function

' 条文内容を取得する関数
Function GetArticleContent(article As Object) As String
    Dim content As String
    Dim articleCaption As Object
    Dim articleTitle As Object
    Dim paragraphs As Object
    Dim paragraph As Object
    Dim paragraphNum As Object
    Dim paragraphSentence As Object
    Dim items As Object
    Dim item As Object
    Dim itemTitle As Object
    Dim itemSentence As Object
    Dim i As Long

    content = ""

    ' ArticleCaption（条文の見出し）を取得
    Set articleCaption = article.SelectSingleNode(".//*[local-name()='ArticleCaption']")
    If Not articleCaption Is Nothing Then
        content = content & GetTextRecursive(articleCaption) & vbLf
    End If

    ' ArticleTitle（条の名称）を取得
    Set articleTitle = article.SelectSingleNode(".//*[local-name()='ArticleTitle']")
    If Not articleTitle Is Nothing Then
        content = content & GetTextRecursive(articleTitle) & vbLf
    End If

    ' Paragraph（項）を全て取得
    Set paragraphs = article.SelectNodes(".//*[local-name()='Paragraph']")
    For i = 0 To paragraphs.Length - 1
        Set paragraph = paragraphs(i)

        ' ParagraphNum（項番号）を取得
        Set paragraphNum = paragraph.SelectSingleNode(".//*[local-name()='ParagraphNum']")
        If Not paragraphNum Is Nothing Then
            content = content & GetTextRecursive(paragraphNum) & " "
        End If

        ' ParagraphSentence（項の本文）を取得
        Set paragraphSentence = paragraph.SelectSingleNode(".//*[local-name()='ParagraphSentence']")
        If Not paragraphSentence Is Nothing Then
            content = content & GetTextRecursive(paragraphSentence) & vbLf
        End If

        ' Item（号）を全て取得
        Set items = paragraph.SelectNodes(".//*[local-name()='Item']")
        If Not items Is Nothing Then
            Dim j As Long
            For j = 0 To items.Length - 1
                Set item = items(j)

                ' ItemTitle（号番号）を取得
                Set itemTitle = item.SelectSingleNode(".//*[local-name()='ItemTitle']")
                If Not itemTitle Is Nothing Then
                    ' 全角スペースでインデント
                    content = content & "　" & GetTextRecursive(itemTitle) & " "
                End If

                ' ItemSentence（号の本文）を取得
                Set itemSentence = item.SelectSingleNode(".//*[local-name()='ItemSentence']")
                If Not itemSentence Is Nothing Then
                    content = content & GetTextRecursive(itemSentence) & vbLf
                End If
            Next j
        End If

        ' 項の間に空行を追加（最後の項以外）
        If i < paragraphs.Length - 1 Then
            content = content & vbLf
        End If
    Next i

    GetArticleContent = Trim(content)
End Function

' 再帰的にテキストを取得する関数
Function GetTextRecursive(element As Object) As String
    Dim text As String
    Dim childNode As Object
    Dim normalizedText As String

    text = ""

    If element Is Nothing Then
        GetTextRecursive = ""
        Exit Function
    End If

    For Each childNode In element.ChildNodes
        If childNode.NodeType = 3 Then ' TEXT_NODE
            ' 空白を正規化（連続する空白・改行を単一スペースに）
            normalizedText = childNode.Text
            normalizedText = Replace(normalizedText, vbCrLf, " ")
            normalizedText = Replace(normalizedText, vbCr, " ")
            normalizedText = Replace(normalizedText, vbLf, " ")
            normalizedText = Replace(normalizedText, vbTab, " ")
            ' 連続するスペースを単一スペースに
            Do While InStr(normalizedText, "  ") > 0
                normalizedText = Replace(normalizedText, "  ", " ")
            Loop
            text = text & normalizedText
        ElseIf childNode.NodeType = 1 Then ' ELEMENT_NODE
            text = text & GetTextRecursive(childNode)
        End If
    Next childNode

    GetTextRecursive = Trim(text)
End Function
