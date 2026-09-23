Attribute VB_Name = "Module3"
Option Explicit

'========================================================
' 消費税法基本通達
'
' B1 : 通達名
' B2 : 通達番号
' B4 : 結果
'
' 例
' B1 = 消基通
' B2 = 4-3-1
'========================================================


'========================================================
' 設定
'========================================================

Private Const SHOUHI_BASE_URL As String = _
    "https://www.nta.go.jp"

Private Const SHOUHI_TOC_PATH As String = _
    "/law/tsutatsu/kihon/shohi/20230930/01.htm"

Private Const SHOUHI_BASE_PATH As String = _
    "/law/tsutatsu/kihon/shohi/20230930/"

Private Const SHOUHI_MAX_DEPTH As Long = 5


'========================================================
' メイン
'========================================================

Public Sub 消費税通達取得()

    Dim ws As Worksheet
    Dim tsutatsuName As String
    Dim tsutatsuNo As String
    Dim normalizedNo As String
    Dim result As String

    Set ws = ActiveSheet

    tsutatsuName = Trim$(CStr(ws.Range("B1").Value))
    tsutatsuNo = Trim$(CStr(ws.Range("B2").Value))

    If tsutatsuName = "" Then
        tsutatsuName = "消基通"
        ws.Range("B1").Value = tsutatsuName
    End If

    If tsutatsuNo = "" Then
        MsgBox _
            "通達番号を入力してください。" & vbCrLf & _
            "例：4-3-1", _
            vbExclamation
        Exit Sub
    End If

    normalizedNo = ShouhiNormalizeTsutatsuNo(tsutatsuNo)

    If normalizedNo = "" Then
        MsgBox _
            "通達番号を正しく入力してください。" & vbCrLf & _
            "例：4-3-1", _
            vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    On Error GoTo ErrHandler

    result = ShouhiGetTsutatsu(normalizedNo)

    If result = "" Then

        ws.Range("B4").Value = _
            "通達「" & normalizedNo & _
            "」を取得できませんでした。"

    Else

        ws.Range("B4").Value = _
            "【消費税法基本通達 " & normalizedNo & "】" & _
            vbCrLf & vbCrLf & _
            result

    End If

    ws.Range("B4").WrapText = True

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:

    Application.ScreenUpdating = True

    MsgBox _
        "通達取得中にエラーが発生しました。" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical

End Sub


'========================================================
' 通達取得
'========================================================

Private Function ShouhiGetTsutatsu( _
    ByVal tsutatsuNo As String) As String

    Dim pageUrl As String
    Dim html As String
    Dim result As String

    '----------------------------------------------------
    ' 通達番号から直接ページを作成
    '
    ' 4-3-1
    '   ↓
    ' /04/03.htm
    '----------------------------------------------------

    pageUrl = ShouhiMakePageUrl(tsutatsuNo)

    If pageUrl <> "" Then

        html = ShouhiHttpGetText(pageUrl)

        If html <> "" Then

            result = _
                ShouhiExtractTsutatsuEntry( _
                    html, _
                    tsutatsuNo)

            If result <> "" Then

                ShouhiGetTsutatsu = _
                    result & _
                    vbCrLf & vbCrLf & _
                    "出典:" & vbCrLf & _
                    pageUrl

                Exit Function

            End If

        End If

    End If

    '----------------------------------------------------
    ' 直接取得できなかった場合
    ' TOCから探す
    '----------------------------------------------------

    result = ShouhiGetTsutatsuFromToc(tsutatsuNo)

    If result <> "" Then
        ShouhiGetTsutatsu = result
    End If

End Function


'========================================================
' 通達番号からページURL作成
'========================================================

Private Function ShouhiMakePageUrl( _
    ByVal tsutatsuNo As String) As String

    Dim parts() As String
    Dim chapterNo As String
    Dim sectionNo As String

    tsutatsuNo = _
        ShouhiNormalizeTsutatsuNo(tsutatsuNo)

    If tsutatsuNo = "" Then Exit Function

    parts = Split(tsutatsuNo, "-")

    If UBound(parts) < 2 Then Exit Function

    chapterNo = parts(0)
    sectionNo = parts(1)

    If Not IsNumeric(chapterNo) Then Exit Function
    If Not IsNumeric(sectionNo) Then Exit Function

    chapterNo = Format$(CLng(chapterNo), "00")
    sectionNo = Format$(CLng(sectionNo), "00")

    ShouhiMakePageUrl = _
        SHOUHI_BASE_URL & _
        SHOUHI_BASE_PATH & _
        chapterNo & "/" & _
        sectionNo & ".htm"

End Function


'========================================================
' TOCから取得
'========================================================

Private Function ShouhiGetTsutatsuFromToc( _
    ByVal tsutatsuNo As String) As String

    Dim tocUrl As String
    Dim html As String
    Dim pageUrl As String
    Dim result As String

    tocUrl = _
        SHOUHI_BASE_URL & _
        SHOUHI_TOC_PATH

    html = ShouhiHttpGetText(tocUrl)

    If html = "" Then Exit Function

    pageUrl = _
        ShouhiFindBestPageFromToc( _
            html, _
            tsutatsuNo)

    If pageUrl <> "" Then

        html = ShouhiHttpGetText(pageUrl)

        If html <> "" Then

            result = _
                ShouhiExtractTsutatsuEntry( _
                    html, _
                    tsutatsuNo)

            If result <> "" Then

                ShouhiGetTsutatsuFromToc = _
                    result & _
                    vbCrLf & vbCrLf & _
                    "出典:" & vbCrLf & _
                    pageUrl

                Exit Function

            End If

        End If

    End If

    '----------------------------------------------------
    ' 再帰検索
    '----------------------------------------------------

    result = _
        ShouhiSearchChildPages( _
            tocUrl, _
            tsutatsuNo, _
            0)

    If result <> "" Then
        ShouhiGetTsutatsuFromToc = result
    End If

End Function


'========================================================
' TOCから候補ページを探す
'========================================================

Private Function ShouhiFindBestPageFromToc( _
    ByVal html As String, _
    ByVal tsutatsuNo As String) As String

    Dim links As Collection
    Dim item As Variant
    Dim linkUrl As String
    Dim linkText As String

    Set links = ShouhiExtractLinks(html)

    '----------------------------------------------------
    ' 通達番号そのものがリンク文字列にある場合
    '----------------------------------------------------

    For Each item In links

        linkUrl = CStr(item(0))
        linkText = CStr(item(1))

        If ShouhiIsTsutatsuUrl(linkUrl) Then

            If ShouhiNumberMatches( _
                linkText, _
                tsutatsuNo) Then

                ShouhiFindBestPageFromToc = linkUrl
                Exit Function

            End If

        End If

    Next item

    '----------------------------------------------------
    ' 法第○条等のリンク
    '----------------------------------------------------

    For Each item In links

        linkUrl = CStr(item(0))
        linkText = CStr(item(1))

        If ShouhiIsTsutatsuUrl(linkUrl) Then

            If InStr( _
                1, _
                linkText, _
                "第", _
                vbTextCompare) > 0 Then

                ShouhiFindBestPageFromToc = linkUrl
                Exit Function

            End If

        End If

    Next item

End Function


'========================================================
' 子ページ再帰検索
'========================================================

Private Function ShouhiSearchChildPages( _
    ByVal pageUrl As String, _
    ByVal tsutatsuNo As String, _
    ByVal depth As Long) As String

    Dim html As String
    Dim links As Collection
    Dim item As Variant
    Dim childUrl As String
    Dim result As String
    Dim pageResult As String

    If depth > SHOUHI_MAX_DEPTH Then Exit Function

    If Not ShouhiIsTsutatsuUrl(pageUrl) Then Exit Function

    html = ShouhiHttpGetText(pageUrl)

    If html = "" Then Exit Function

    '----------------------------------------------------
    ' このページを確認
    '----------------------------------------------------

    pageResult = _
        ShouhiExtractTsutatsuEntry( _
            html, _
            tsutatsuNo)

    If pageResult <> "" Then

        ShouhiSearchChildPages = _
            pageResult & _
            vbCrLf & vbCrLf & _
            "出典:" & vbCrLf & _
            pageUrl

        Exit Function

    End If

    '----------------------------------------------------
    ' 子リンク
    '----------------------------------------------------

    Set links = ShouhiExtractLinks(html)

    For Each item In links

        childUrl = CStr(item(0))

        If ShouhiIsTsutatsuUrl(childUrl) Then

            If Not ShouhiIsNonHtmlFile(childUrl) Then

                If StrComp( _
                    childUrl, _
                    pageUrl, _
                    vbTextCompare) <> 0 Then

                    result = _
                        ShouhiSearchChildPages( _
                            childUrl, _
                            tsutatsuNo, _
                            depth + 1)

                    If result <> "" Then

                        ShouhiSearchChildPages = result
                        Exit Function

                    End If

                End If

            End If

        End If

    Next item

End Function


'========================================================
' 通達本文抽出
'
' ★今回の修正版
'
' 例：
'
' （権利の内容に応ずることの例示）
' 4－3－2　本文……
'
' の場合、
'
' 「（権利の内容に応ずることの例示）」
' から先を本文に含めない。
'========================================================

Private Function ShouhiExtractTsutatsuEntry( _
    ByVal html As String, _
    ByVal tsutatsuNo As String) As String

    Dim text As String

    Dim startPos As Long
    Dim nextNoPos As Long
    Dim endPos As Long

    Dim body As String

    '----------------------------------------------------
    ' HTML → テキスト
    '----------------------------------------------------

    text = ShouhiHtmlToText(html)

    If text = "" Then Exit Function

    '----------------------------------------------------
    ' 現在の通達番号
    '----------------------------------------------------

    startPos = _
        ShouhiFindNumberFlexible( _
            text, _
            tsutatsuNo)

    If startPos <= 0 Then Exit Function

    '----------------------------------------------------
    ' 通達番号直後から本文開始
    '----------------------------------------------------

    startPos = _
        ShouhiFindBodyStart( _
            text, _
            startPos)

    If startPos <= 0 Then Exit Function

    '----------------------------------------------------
    ' 次の通達番号
    '----------------------------------------------------

    nextNoPos = _
        ShouhiFindNextTsutatsuNumber( _
            text, _
            startPos)

    If nextNoPos > 0 Then

        '------------------------------------------------
        ' ★重要
        '
        ' 次の通達番号の直前にある
        ' 「（・・・・）」形式の見出しを探す。
        '
        ' 見つかったらそこから切る。
        '------------------------------------------------

        endPos = _
            ShouhiFindHeadingBefore( _
                text, _
                startPos, _
                nextNoPos)

        If endPos <= startPos Then
            endPos = nextNoPos
        End If

        body = _
            Mid$( _
                text, _
                startPos, _
                endPos - startPos)

    Else

        '------------------------------------------------
        ' 次の通達番号がない場合
        '------------------------------------------------

        endPos = _
            ShouhiFindFooterPosition( _
                text, _
                startPos)

        If endPos <= startPos Then

            body = Mid$(text, startPos)

        Else

            body = _
                Mid$( _
                    text, _
                    startPos, _
                    endPos - startPos)

        End If

    End If

    '----------------------------------------------------
    ' 本文整理
    '----------------------------------------------------

    body = ShouhiCleanText(body)

    If body = "" Then Exit Function

    ShouhiExtractTsutatsuEntry = body

End Function


'========================================================
' 通達番号検索
'========================================================

Private Function ShouhiFindNumberFlexible( _
    ByVal text As String, _
    ByVal tsutatsuNo As String) As Long

    Dim target As String
    Dim i As Long
    Dim j As Long
    Dim targetIndex As Long

    Dim c As String
    Dim normalizedChar As String

    target = _
        ShouhiNormalizeTsutatsuNo(tsutatsuNo)

    If target = "" Then Exit Function

    For i = 1 To Len(text)

        If ShouhiIsDigitCharacter( _
            Mid$(text, i, 1)) Then

            targetIndex = 1
            j = i

            Do While _
                j <= Len(text) And _
                targetIndex <= Len(target)

                c = Mid$(text, j, 1)

                normalizedChar = _
                    ShouhiNormalizeNumberCharacter(c)

                If normalizedChar = "" Then Exit Do

                If normalizedChar <> _
                    Mid$(target, targetIndex, 1) Then

                    Exit Do

                End If

                targetIndex = targetIndex + 1
                j = j + 1

            Loop

            If targetIndex > Len(target) Then

                If ShouhiValidNumberBoundary( _
                    text, _
                    i, _
                    j - 1) Then

                    ShouhiFindNumberFlexible = i
                    Exit Function

                End If

            End If

        End If

    Next i

End Function


'========================================================
' 通達番号の境界チェック
'========================================================

Private Function ShouhiValidNumberBoundary( _
    ByVal text As String, _
    ByVal startPos As Long, _
    ByVal endPos As Long) As Boolean

    Dim c As String

    If startPos > 1 Then

        c = Mid$(text, startPos - 1, 1)

        If ShouhiIsNumberCharacter(c) Then
            Exit Function
        End If

    End If

    If endPos < Len(text) Then

        c = Mid$(text, endPos + 1, 1)

        If ShouhiIsDigitCharacter(c) Then
            Exit Function
        End If

    End If

    ShouhiValidNumberBoundary = True

End Function


'========================================================
' 次の通達番号を探す
'========================================================

Private Function ShouhiFindNextTsutatsuNumber( _
    ByVal text As String, _
    ByVal startPos As Long) As Long

    Dim i As Long
    Dim j As Long

    Dim digitCount As Long
    Dim hyphenCount As Long

    Dim c As String
    Dim normalizedChar As String
    Dim numberText As String

    Dim lineStart As Boolean

    i = startPos + 1

    Do While i <= Len(text)

        c = Mid$(text, i, 1)

        '------------------------------------------------
        ' 行頭判定
        '------------------------------------------------

        lineStart = False

        If i = 1 Then

            lineStart = True

        ElseIf Mid$(text, i - 1, 1) = vbLf Then

            lineStart = True

        ElseIf Mid$(text, i - 1, 1) = vbCr Then

            lineStart = True

        End If

        '------------------------------------------------
        ' 行頭空白を飛ばす
        '------------------------------------------------

        If lineStart Then

            Do While i <= Len(text)

                c = Mid$(text, i, 1)

                If _
                    c = " " Or _
                    c = "　" Or _
                    c = vbTab Then

                    i = i + 1

                Else

                    Exit Do

                End If

            Loop

        End If

        If i > Len(text) Then Exit Do

        '------------------------------------------------
        ' 数字から始まるか
        '------------------------------------------------

        If ShouhiIsDigitCharacter( _
            Mid$(text, i, 1)) Then

            numberText = ""
            digitCount = 0
            hyphenCount = 0

            j = i

            Do While j <= Len(text)

                c = Mid$(text, j, 1)

                normalizedChar = _
                    ShouhiNormalizeNumberCharacter(c)

                If normalizedChar = "" Then Exit Do

                If _
                    normalizedChar >= "0" And _
                    normalizedChar <= "9" Then

                    numberText = _
                        numberText & normalizedChar

                    digitCount = digitCount + 1

                ElseIf normalizedChar = "-" Then

                    numberText = _
                        numberText & "-"

                    hyphenCount = hyphenCount + 1

                Else

                    Exit Do

                End If

                If _
                    digitCount >= 3 And _
                    hyphenCount >= 2 Then

                    Exit Do

                End If

                j = j + 1

            Loop

            If _
                digitCount >= 3 And _
                hyphenCount >= 2 Then

                If ShouhiLooksLikeTsutatsuNo(numberText) Then

                    ShouhiFindNextTsutatsuNumber = i
                    Exit Function

                End If

            End If

        End If

        i = i + 1

    Loop

End Function


'========================================================
' 通達番号らしいか
'========================================================

Private Function ShouhiLooksLikeTsutatsuNo( _
    ByVal numberText As String) As Boolean

    Dim parts() As String

    If numberText = "" Then Exit Function

    parts = Split(numberText, "-")

    If UBound(parts) <> 2 Then Exit Function

    If parts(0) = "" Then Exit Function
    If parts(1) = "" Then Exit Function
    If parts(2) = "" Then Exit Function

    If Not IsNumeric(parts(0)) Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function
    If Not IsNumeric(parts(2)) Then Exit Function

    ShouhiLooksLikeTsutatsuNo = True

End Function


'========================================================
' 数字1文字正規化
'========================================================

Private Function ShouhiNormalizeNumberCharacter( _
    ByVal c As String) As String

    If c = "" Then Exit Function

    ' 半角数字
    If _
        c >= "0" And _
        c <= "9" Then

        ShouhiNormalizeNumberCharacter = c
        Exit Function

    End If

    ' 全角数字
    If _
        c >= "０" And _
        c <= "９" Then

        ShouhiNormalizeNumberCharacter = _
            Chr$(AscW(c) - &HFEE0)

        Exit Function

    End If

    ' ハイフン類
    If _
        c = "-" Or _
        c = "－" Or _
        c = "‐" Or _
        c = "?" Or _
        c = "ー" Or _
        c = "―" Then

        ShouhiNormalizeNumberCharacter = "-"
        Exit Function

    End If

End Function


'========================================================
' 数字判定
'========================================================

Private Function ShouhiIsDigitCharacter( _
    ByVal c As String) As Boolean

    If c = "" Then Exit Function

    If _
        c >= "0" And _
        c <= "9" Then

        ShouhiIsDigitCharacter = True
        Exit Function

    End If

    If _
        c >= "０" And _
        c <= "９" Then

        ShouhiIsDigitCharacter = True
        Exit Function

    End If

End Function


'========================================================
' 数字・ハイフン判定
'========================================================

Private Function ShouhiIsNumberCharacter( _
    ByVal c As String) As Boolean

    If ShouhiIsDigitCharacter(c) Then

        ShouhiIsNumberCharacter = True
        Exit Function

    End If

    If _
        c = "-" Or _
        c = "－" Or _
        c = "‐" Or _
        c = "?" Or _
        c = "ー" Or _
        c = "―" Then

        ShouhiIsNumberCharacter = True
        Exit Function

    End If

End Function


'========================================================
' ★次の通達番号直前の見出しを探す
'
' 例：
'
' 4－3－1 本文
'
' （権利の内容に応ずることの例示）
' 4－3－2 本文
'
' ↓
'
' （権利の内容に応ずることの例示）
' の開始位置を返す
'========================================================

Private Function ShouhiFindHeadingBefore( _
    ByVal text As String, _
    ByVal startPos As Long, _
    ByVal nextNoPos As Long) As Long

    Dim i As Long
    Dim lineStart As Long
    Dim lineEnd As Long

    Dim lineText As String
    Dim trimmedLine As String

    Dim bestPos As Long

    i = startPos

    Do While i < nextNoPos

        lineStart = i

        lineEnd = _
            InStr(i, text, vbLf)

        If lineEnd = 0 Or _
           lineEnd >= nextNoPos Then

            lineEnd = nextNoPos

        End If

        lineText = _
            Mid$( _
                text, _
                lineStart, _
                lineEnd - lineStart)

        trimmedLine = Trim$(lineText)

        '------------------------------------------------
        ' >> を除去して判定
        '------------------------------------------------

        Do While Left$(trimmedLine, 2) = ">>"

            trimmedLine = _
                Trim$(Mid$(trimmedLine, 3))

        Loop

        '------------------------------------------------
        ' 括弧見出しなら記録
        '------------------------------------------------

        If ShouhiIsHeadingLine(trimmedLine) Then

            bestPos = lineStart

        End If

        If lineEnd >= nextNoPos Then Exit Do

        i = lineEnd + 1

    Loop

    '----------------------------------------------------
    ' 通常の改行ベースで見つかった場合
    '----------------------------------------------------

    If bestPos > 0 Then

        ShouhiFindHeadingBefore = bestPos
        Exit Function

    End If

    '----------------------------------------------------
    ' ★追加処理
    '
    ' HTML→テキスト化の都合で、
    '
    ' （見出し）4－3－2
    '
    ' のように同じ行になった場合に備える。
    '----------------------------------------------------

    ShouhiFindHeadingBefore = _
        ShouhiFindInlineHeadingBefore( _
            text, _
            startPos, _
            nextNoPos)

End Function


'========================================================
' 同一行にくっついた見出しを逆方向から探す
'========================================================

Private Function ShouhiFindInlineHeadingBefore( _
    ByVal text As String, _
    ByVal startPos As Long, _
    ByVal nextNoPos As Long) As Long

    Dim p As Long
    Dim q As Long

    Dim candidate As String

    p = nextNoPos - 1

    Do While p >= startPos

        If Mid$(text, p, 1) = "）" Then

            q = InStrRev( _
                text, _
                "（", _
                p)

            If q >= startPos Then

                candidate = _
                    Mid$( _
                        text, _
                        q, _
                        p - q + 1)

                If ShouhiIsHeadingLine(candidate) Then

                    ShouhiFindInlineHeadingBefore = q
                    Exit Function

                End If

            End If

        End If

        p = p - 1

    Loop

End Function


'========================================================
' 見出し行判定
'========================================================

Private Function ShouhiIsHeadingLine( _
    ByVal lineText As String) As Boolean

    Dim firstChar As String
    Dim lastChar As String

    lineText = Trim$(lineText)

    If lineText = "" Then Exit Function

    Do While Left$(lineText, 2) = ">>"

        lineText = _
            Trim$(Mid$(lineText, 3))

    Loop

    If Len(lineText) < 2 Then Exit Function

    firstChar = Left$(lineText, 1)
    lastChar = Right$(lineText, 1)

    If firstChar <> "（" Then Exit Function
    If lastChar <> "）" Then Exit Function

    ' 明らかな注記は除外
    If lineText = "（注）" Then Exit Function

    ShouhiIsHeadingLine = True

End Function


'========================================================
' フッター位置
'========================================================

Private Function ShouhiFindFooterPosition( _
    ByVal text As String, _
    ByVal startPos As Long) As Long

    Dim p1 As Long
    Dim p2 As Long

    p1 = _
        InStr( _
            startPos, _
            text, _
            "このページの先頭へ", _
            vbTextCompare)

    p2 = _
        InStr( _
            startPos, _
            text, _
            "法令等", _
            vbTextCompare)

    If p1 > 0 And p2 > 0 Then

        If p1 < p2 Then
            ShouhiFindFooterPosition = p1
        Else
            ShouhiFindFooterPosition = p2
        End If

        Exit Function

    End If

    If p1 > 0 Then

        ShouhiFindFooterPosition = p1
        Exit Function

    End If

    If p2 > 0 Then

        ShouhiFindFooterPosition = p2
        Exit Function

    End If

    ShouhiFindFooterPosition = Len(text) + 1

End Function


'========================================================
' 通達番号直後から本文開始
'========================================================

Private Function ShouhiFindBodyStart( _
    ByVal text As String, _
    ByVal numberPos As Long) As Long

    Dim p As Long
    Dim c As String

    p = numberPos

    '----------------------------------------------------
    ' 数字・ハイフンを飛ばす
    '----------------------------------------------------

    Do While p <= Len(text)

        c = Mid$(text, p, 1)

        If ShouhiIsNumberCharacter(c) Then

            p = p + 1

        Else

            Exit Do

        End If

    Loop

    '----------------------------------------------------
    ' 番号直後の空白・コロンを飛ばす
    '----------------------------------------------------

    Do While p <= Len(text)

        c = Mid$(text, p, 1)

        If _
            c = " " Or _
            c = "　" Or _
            c = vbTab Or _
            c = ":" Or _
            c = "：" Then

            p = p + 1

        Else

            Exit Do

        End If

    Loop

    ShouhiFindBodyStart = p

End Function


'========================================================
' 通達番号正規化
'========================================================

Private Function ShouhiNormalizeTsutatsuNo( _
    ByVal text As String) As String

    text = Trim$(text)

    On Error Resume Next

    text = StrConv(text, vbNarrow)

    On Error GoTo 0

    text = Replace(text, "－", "-")
    text = Replace(text, "‐", "-")
    text = Replace(text, "?", "-")
    text = Replace(text, "ー", "-")
    text = Replace(text, "―", "-")

    text = Replace(text, " ", "")
    text = Replace(text, "　", "")

    ShouhiNormalizeTsutatsuNo = text

End Function


'========================================================
' 通達番号一致
'========================================================

Private Function ShouhiNumberMatches( _
    ByVal text As String, _
    ByVal tsutatsuNo As String) As Boolean

    Dim normalizedText As String
    Dim normalizedNo As String

    normalizedText = _
        ShouhiNormalizeNumberText(text)

    normalizedNo = _
        ShouhiNormalizeTsutatsuNo(tsutatsuNo)

    If normalizedNo = "" Then Exit Function

    ShouhiNumberMatches = _
        InStr( _
            1, _
            normalizedText, _
            normalizedNo, _
            vbBinaryCompare) > 0

End Function


'========================================================
' 検索用文字列正規化
'========================================================

Private Function ShouhiNormalizeNumberText( _
    ByVal text As String) As String

    On Error Resume Next

    text = StrConv(text, vbNarrow)

    On Error GoTo 0

    text = Replace(text, "－", "-")
    text = Replace(text, "‐", "-")
    text = Replace(text, "?", "-")
    text = Replace(text, "ー", "-")
    text = Replace(text, "―", "-")

    text = Replace(text, " ", "")
    text = Replace(text, "　", "")

    ShouhiNormalizeNumberText = text

End Function


'========================================================
' HTML → テキスト
'
' ★修正
'
' h1～h6、section、article等にも改行を入れる。
'========================================================

Private Function ShouhiHtmlToText( _
    ByVal html As String) As String

    Dim text As String
    Dim re As Object

    text = html

    Set re = CreateObject("VBScript.RegExp")

    '----------------------------------------------------
    ' script / style / noscript
    '----------------------------------------------------

    With re

        .Global = True
        .IgnoreCase = True
        .MultiLine = True

        .pattern = "<script[\s\S]*?</script>"
        text = .Replace(text, "")

        .pattern = "<style[\s\S]*?</style>"
        text = .Replace(text, "")

        .pattern = "<noscript[\s\S]*?</noscript>"
        text = .Replace(text, "")

    End With

    '----------------------------------------------------
    ' ブロック開始タグ
    '----------------------------------------------------

    With re

        .Global = True
        .IgnoreCase = True
        .MultiLine = True

        .pattern = "<(h[1-6]|section|article|header|main|p|div|li|tr|table)[^>]*>"

        text = .Replace( _
            text, _
            vbCrLf)

    End With

    '----------------------------------------------------
    ' 改行タグ
    '----------------------------------------------------

    With re

        .Global = True
        .IgnoreCase = True
        .MultiLine = True

        .pattern = "<br\s*/?>"
        text = .Replace(text, vbCrLf)

        .pattern = "</(h[1-6]|section|article|header|main|p|div|li|tr|table)\s*>"
        text = .Replace(text, vbCrLf)

    End With

    '----------------------------------------------------
    ' HTMLタグ除去
    '----------------------------------------------------

    With re

        .Global = True
        .IgnoreCase = True
        .MultiLine = True

        .pattern = "<[^>]+>"
        text = .Replace(text, "")

    End With

    '----------------------------------------------------
    ' HTMLエンティティ
    '----------------------------------------------------

    text = ShouhiDecodeHtmlEntities(text)

    '----------------------------------------------------
    ' >> を改行
    '----------------------------------------------------

    text = _
        Replace( _
            text, _
            ">>", _
            vbCrLf)

    '----------------------------------------------------
    ' 改行統一
    '----------------------------------------------------

    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)

    '----------------------------------------------------
    ' 行末空白
    '----------------------------------------------------

    With re

        .Global = True
        .MultiLine = True

        .pattern = "[ \t]+$"

        text = .Replace(text, "")

    End With

    text = Replace(text, vbLf, vbCrLf)

    ShouhiHtmlToText = text

End Function


'========================================================
' HTMLエンティティ
'========================================================

Private Function ShouhiDecodeHtmlEntities( _
    ByVal text As String) As String

    text = Replace(text, "&nbsp;", " ")
    text = Replace(text, "&amp;", "&")
    text = Replace(text, "&lt;", "<")
    text = Replace(text, "&gt;", ">")
    text = Replace(text, "&quot;", """")
    text = Replace(text, "&#39;", "'")
    text = Replace(text, "&#x3000;", "　")
    text = Replace(text, "&#12288;", "　")

    ShouhiDecodeHtmlEntities = text

End Function


'========================================================
' 本文整理
'========================================================

Private Function ShouhiCleanText( _
    ByVal text As String) As String

    Dim re As Object

    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)

    Set re = CreateObject("VBScript.RegExp")

    '----------------------------------------------------
    ' >> 単独行削除
    '----------------------------------------------------

    With re

        .Global = True
        .MultiLine = True

        .pattern = "^[ \t]*>>[ \t]*$"

        text = .Replace(text, "")

    End With

    '----------------------------------------------------
    ' 行頭空白
    '----------------------------------------------------

    With re

        .Global = True
        .MultiLine = True

        .pattern = "^[ \t]+"

        text = .Replace(text, "")

    End With

    '----------------------------------------------------
    ' 行末空白
    '----------------------------------------------------

    With re

        .Global = True
        .MultiLine = True

        .pattern = "[ \t]+$"

        text = .Replace(text, "")

    End With

    '----------------------------------------------------
    ' 連続空行を整理
    '----------------------------------------------------

    With re

        .Global = True
        .MultiLine = True

        .pattern = "(\n[ \t]*){3,}"

        text = _
            .Replace( _
                text, _
                vbLf & vbLf)

    End With

    text = Trim$(text)

    text = Replace(text, vbLf, vbCrLf)

    ShouhiCleanText = text

End Function


'========================================================
' HTMLリンク抽出
'========================================================

Private Function ShouhiExtractLinks( _
    ByVal html As String) As Collection

    Dim result As New Collection

    Dim re As Object
    Dim matches As Object
    Dim m As Object

    Dim href As String
    Dim linkText As String

    Dim item(0 To 1) As String

    Set re = CreateObject("VBScript.RegExp")

    With re

        .Global = True
        .IgnoreCase = True
        .MultiLine = True

        .pattern = _
            "<a\b[^>]*href\s*=\s*[""" & "'" & _
            "]([^""" & "'" & _
            "]+)[""" & "'" & _
            "][^>]*>([\s\S]*?)</a>"

    End With

    Set matches = re.Execute(html)

    For Each m In matches

        href = Trim$(CStr(m.SubMatches(0)))

        linkText = CStr(m.SubMatches(1))

        linkText = _
            ShouhiHtmlToText(linkText)

        linkText = Trim$(linkText)

        href = ShouhiResolveUrl(href)

        If href <> "" Then

            item(0) = href
            item(1) = linkText

            result.Add item

        End If

    Next m

    Set ShouhiExtractLinks = result

End Function


'========================================================
' URL解決
'========================================================

Private Function ShouhiResolveUrl( _
    ByVal href As String) As String

    Dim baseUrl As String
    Dim path As String
    Dim lastSlash As Long

    href = Trim$(href)

    If href = "" Then Exit Function

    If Left$(href, 1) = "#" Then Exit Function

    If LCase$(Left$(href, 11)) = _
        "javascript:" Then Exit Function

    If LCase$(Left$(href, 7)) = _
        "mailto:" Then Exit Function

    If LCase$(Left$(href, 8)) = _
        "https://" Then

        ShouhiResolveUrl = href
        Exit Function

    End If

    If LCase$(Left$(href, 7)) = _
        "http://" Then

        ShouhiResolveUrl = href
        Exit Function

    End If

    If Left$(href, 2) = "//" Then

        ShouhiResolveUrl = _
            "https:" & href

        Exit Function

    End If

    If Left$(href, 1) = "/" Then

        ShouhiResolveUrl = _
            SHOUHI_BASE_URL & href

        Exit Function

    End If

    baseUrl = _
        SHOUHI_BASE_URL & _
        SHOUHI_BASE_PATH

    lastSlash = InStrRev(baseUrl, "/")

    If lastSlash > 0 Then

        path = _
            Left$(baseUrl, lastSlash) & _
            href

    Else

        path = baseUrl & href

    End If

    ShouhiResolveUrl = _
        ShouhiNormalizeUrl(path)

End Function


'========================================================
' URL正規化
'========================================================

Private Function ShouhiNormalizeUrl( _
    ByVal url As String) As String

    Dim parts() As String
    Dim i As Long

    Dim stack As Collection

    Dim part As String
    Dim result As String

    If url = "" Then Exit Function

    parts = Split(url, "/")

    Set stack = New Collection

    For i = _
        LBound(parts) To _
        UBound(parts)

        part = parts(i)

        If part = ".." Then

            If stack.Count > 0 Then
                stack.Remove stack.Count
            End If

        ElseIf _
            part <> "." And _
            part <> "" Then

            stack.Add part

        End If

    Next i

    result = "https://www.nta.go.jp"

    For i = 1 To stack.Count

        result = _
            result & "/" & _
            CStr(stack(i))

    Next i

    ShouhiNormalizeUrl = result

End Function


'========================================================
' 消費税基本通達配下か
'========================================================

Private Function ShouhiIsTsutatsuUrl( _
    ByVal url As String) As Boolean

    Dim normalized As String

    normalized = LCase$(Trim$(url))

    ShouhiIsTsutatsuUrl = _
        InStr( _
            1, _
            normalized, _
            LCase$( _
                SHOUHI_BASE_URL & _
                SHOUHI_BASE_PATH), _
            vbTextCompare) = 1

End Function


'========================================================
' HTMLではないファイルを除外
'========================================================

Private Function ShouhiIsNonHtmlFile( _
    ByVal url As String) As Boolean

    Dim lowerUrl As String

    lowerUrl = LCase$(url)

    If InStr(lowerUrl, ".xlsx") > 0 Then

        ShouhiIsNonHtmlFile = True
        Exit Function

    End If

    If InStr(lowerUrl, ".xls") > 0 Then

        ShouhiIsNonHtmlFile = True
        Exit Function

    End If

    If InStr(lowerUrl, ".pdf") > 0 Then

        ShouhiIsNonHtmlFile = True
        Exit Function

    End If

    If InStr(lowerUrl, ".doc") > 0 Then

        ShouhiIsNonHtmlFile = True
        Exit Function

    End If

    If InStr(lowerUrl, ".zip") > 0 Then

        ShouhiIsNonHtmlFile = True
        Exit Function

    End If

End Function


'========================================================
' HTTP GET
'========================================================

Private Function ShouhiHttpGetText( _
    ByVal url As String) As String

    Dim http As Object

    Dim bytes() As Byte
    Dim charset As String

    If Not ShouhiIsTsutatsuUrl(url) Then Exit Function

    If ShouhiIsNonHtmlFile(url) Then Exit Function

    On Error GoTo ErrHandler

    Set http = _
        CreateObject("MSXML2.XMLHTTP")

    http.Open _
        "GET", _
        url, _
        False

    http.setRequestHeader _
        "User-Agent", _
        "Mozilla/5.0"

    http.send

    If _
        http.Status < 200 Or _
        http.Status >= 300 Then

        Exit Function

    End If

    bytes = http.responseBody

    charset = _
        ShouhiDetectCharset(bytes)

    If charset = "" Then
        charset = "shift_jis"
    End If

    ShouhiHttpGetText = _
        ShouhiBytesToText( _
            bytes, _
            charset)

    Exit Function

ErrHandler:

    ShouhiHttpGetText = ""

End Function


'========================================================
' charset判定
'========================================================

Private Function ShouhiDetectCharset( _
    ByRef bytes() As Byte) As String

    Dim preview As String

    On Error GoTo ErrHandler

    preview = _
        ShouhiBytesToText( _
            bytes, _
            "iso-8859-1")

    preview = LCase$(preview)

    If InStr( _
        preview, _
        "charset=utf-8") > 0 Then

        ShouhiDetectCharset = "utf-8"
        Exit Function

    End If

    If InStr( _
        preview, _
        "charset=""utf-8""") > 0 Then

        ShouhiDetectCharset = "utf-8"
        Exit Function

    End If

    If InStr( _
        preview, _
        "charset=shift_jis") > 0 Then

        ShouhiDetectCharset = "shift_jis"
        Exit Function

    End If

    If InStr( _
        preview, _
        "charset=""shift_jis""") > 0 Then

        ShouhiDetectCharset = "shift_jis"
        Exit Function

    End If

    If InStr( _
        preview, _
        "charset=shift-jis") > 0 Then

        ShouhiDetectCharset = "shift_jis"
        Exit Function

    End If

    ShouhiDetectCharset = "shift_jis"
    Exit Function

ErrHandler:

    ShouhiDetectCharset = "shift_jis"

End Function


'========================================================
' Byte → String
'========================================================

Private Function ShouhiBytesToText( _
    ByRef bytes() As Byte, _
    ByVal charset As String) As String

    Dim stream As Object

    On Error GoTo ErrHandler

    Set stream = _
        CreateObject("ADODB.Stream")

    With stream

        .Type = 1
        .Open

        .Write bytes

        .Position = 0

        .Type = 2
        .charset = charset

        ShouhiBytesToText = _
            .ReadText

        .Close

    End With

    Exit Function

ErrHandler:

    ShouhiBytesToText = ""

End Function

