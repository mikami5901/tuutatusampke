Attribute VB_Name = "Module2"
Option Explicit

'========================================================
' 法人税法基本通達
'
' B1 : 通達名
' B2 : 通達番号
' B4 : 結果
'
' 例
' B1 = 法基通
' B2 = 1-2-1
'========================================================


'========================================================
' 設定
'========================================================

Private Const HOJIN_BASE_URL As String = _
    "https://www.nta.go.jp"

Private Const HOJIN_TOC_PATH As String = _
    "/law/tsutatsu/kihon/hojin/01.htm"

Private Const HOJIN_BASE_PATH As String = _
    "/law/tsutatsu/kihon/hojin/"

Private Const HOJIN_MAX_DEPTH As Long = 5


'========================================================
' メイン
'========================================================

Public Sub 法人税通達取得()

    Dim ws As Worksheet

    Dim tsutatsuName As String
    Dim tsutatsuNo As String
    Dim normalizedNo As String

    Dim result As String

    Set ws = ActiveSheet

    tsutatsuName = _
        Trim$(CStr(ws.Range("B1").Value))

    tsutatsuNo = _
        Trim$(CStr(ws.Range("B2").Value))

    If tsutatsuName = "" Then

        tsutatsuName = "法基通"

        ws.Range("B1").Value = _
            tsutatsuName

    End If

    If tsutatsuNo = "" Then

        MsgBox _
            "通達番号を入力してください。" & _
            vbCrLf & _
            "例：1-2-1", _
            vbExclamation

        Exit Sub

    End If

    normalizedNo = _
        HojinNormalizeTsutatsuNo(tsutatsuNo)

    If normalizedNo = "" Then

        MsgBox _
            "通達番号を正しく入力してください。" & _
            vbCrLf & _
            "例：1-2-1", _
            vbExclamation

        Exit Sub

    End If

    Application.ScreenUpdating = False

    On Error GoTo ErrHandler

    result = _
        HojinGetTsutatsu(normalizedNo)

    If result = "" Then

        ws.Range("B4").Value = _
            "通達「" & _
            normalizedNo & _
            "」を取得できませんでした。"

    Else

        ws.Range("B4").Value = _
            "【法人税法基本通達 " & _
            normalizedNo & _
            "】" & _
            vbCrLf & _
            vbCrLf & _
            result

    End If

    ws.Range("B4").WrapText = True

    Application.ScreenUpdating = True

    Exit Sub

ErrHandler:

    Application.ScreenUpdating = True

    MsgBox _
        "通達取得中にエラーが発生しました。" & _
        vbCrLf & _
        vbCrLf & _
        Err.Description, _
        vbCritical

End Sub


'========================================================
' 通達取得
'========================================================

Private Function HojinGetTsutatsu( _
    ByVal tsutatsuNo As String) As String

    Dim pageUrl As String
    Dim html As String
    Dim result As String

    '----------------------------------------------------
    ' まず通達番号から直接ページを作る
    '
    ' 1-2-1
    ' ↓
    ' /01/01_02.htm
    '----------------------------------------------------

    pageUrl = _
        HojinMakePageUrl(tsutatsuNo)

    If pageUrl <> "" Then

        html = _
            HojinHttpGetText(pageUrl)

        If html <> "" Then

            result = _
                HojinExtractTsutatsuEntry( _
                    html, _
                    tsutatsuNo)

            If result <> "" Then

                HojinGetTsutatsu = _
                    result & _
                    vbCrLf & _
                    vbCrLf & _
                    "出典:" & _
                    vbCrLf & _
                    pageUrl

                Exit Function

            End If

        End If

    End If

    '----------------------------------------------------
    ' 直接取得できなかった場合
    ' TOCから探す
    '----------------------------------------------------

    result = _
        HojinGetTsutatsuFromToc(tsutatsuNo)

    If result <> "" Then

        HojinGetTsutatsu = result

    End If

End Function


'========================================================
' 通達番号からページURL作成
'
' 1-2-1
' ↓
' /01/01_02.htm
'========================================================

Private Function HojinMakePageUrl( _
    ByVal tsutatsuNo As String) As String

    Dim parts() As String

    Dim chapterNo As String
    Dim sectionNo As String

    tsutatsuNo = _
        HojinNormalizeTsutatsuNo(tsutatsuNo)

    If tsutatsuNo = "" Then Exit Function

    parts = Split(tsutatsuNo, "-")

    If UBound(parts) < 2 Then Exit Function

    chapterNo = parts(0)
    sectionNo = parts(1)

    If Not IsNumeric(chapterNo) Then Exit Function
    If Not IsNumeric(sectionNo) Then Exit Function

    chapterNo = _
        Format$(CLng(chapterNo), "00")

    sectionNo = _
        Format$(CLng(sectionNo), "00")

    HojinMakePageUrl = _
        HOJIN_BASE_URL & _
        HOJIN_BASE_PATH & _
        chapterNo & "/" & _
        chapterNo & "_" & _
        sectionNo & ".htm"

End Function


'========================================================
' TOCから取得
'========================================================

Private Function HojinGetTsutatsuFromToc( _
    ByVal tsutatsuNo As String) As String

    Dim tocUrl As String
    Dim html As String
    Dim pageUrl As String
    Dim result As String

    tocUrl = _
        HOJIN_BASE_URL & _
        HOJIN_TOC_PATH

    html = _
        HojinHttpGetText(tocUrl)

    If html = "" Then Exit Function

    pageUrl = _
        HojinFindBestPageFromToc( _
            html, _
            tsutatsuNo)

    If pageUrl <> "" Then

        html = _
            HojinHttpGetText(pageUrl)

        If html <> "" Then

            result = _
                HojinExtractTsutatsuEntry( _
                    html, _
                    tsutatsuNo)

            If result <> "" Then

                HojinGetTsutatsuFromToc = _
                    result & _
                    vbCrLf & _
                    vbCrLf & _
                    "出典:" & _
                    vbCrLf & _
                    pageUrl

                Exit Function

            End If

        End If

    End If

    '----------------------------------------------------
    ' 再帰検索
    '----------------------------------------------------

    result = _
        HojinSearchChildPages( _
            tocUrl, _
            tsutatsuNo, _
            0)

    If result <> "" Then

        HojinGetTsutatsuFromToc = result

    End If

End Function


'========================================================
' TOCから候補ページを探す
'========================================================

Private Function HojinFindBestPageFromToc( _
    ByVal html As String, _
    ByVal tsutatsuNo As String) As String

    Dim links As Collection

    Dim item As Variant

    Dim linkUrl As String
    Dim linkText As String

    Set links = _
        HojinExtractLinks(html)

    '----------------------------------------------------
    ' 通達番号そのものがリンク文字列にある場合
    '----------------------------------------------------

    For Each item In links

        linkUrl = CStr(item(0))
        linkText = CStr(item(1))

        If HojinIsTsutatsuUrl(linkUrl) Then

            If HojinNumberMatches( _
                linkText, _
                tsutatsuNo) Then

                HojinFindBestPageFromToc = _
                    linkUrl

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

        If HojinIsTsutatsuUrl(linkUrl) Then

            If InStr( _
                1, _
                linkText, _
                "第", _
                vbTextCompare) > 0 Then

                HojinFindBestPageFromToc = _
                    linkUrl

                Exit Function

            End If

        End If

    Next item

End Function


'========================================================
' 子ページ再帰検索
'========================================================

Private Function HojinSearchChildPages( _
    ByVal pageUrl As String, _
    ByVal tsutatsuNo As String, _
    ByVal depth As Long) As String

    Dim html As String

    Dim links As Collection
    Dim item As Variant

    Dim childUrl As String

    Dim result As String
    Dim pageResult As String

    If depth > HOJIN_MAX_DEPTH Then Exit Function

    If Not HojinIsTsutatsuUrl(pageUrl) Then Exit Function

    html = _
        HojinHttpGetText(pageUrl)

    If html = "" Then Exit Function

    '----------------------------------------------------
    ' このページを確認
    '----------------------------------------------------

    pageResult = _
        HojinExtractTsutatsuEntry( _
            html, _
            tsutatsuNo)

    If pageResult <> "" Then

        HojinSearchChildPages = _
            pageResult & _
            vbCrLf & _
            vbCrLf & _
            "出典:" & _
            vbCrLf & _
            pageUrl

        Exit Function

    End If

    '----------------------------------------------------
    ' 子リンク
    '----------------------------------------------------

    Set links = _
        HojinExtractLinks(html)

    For Each item In links

        childUrl = CStr(item(0))

        If HojinIsTsutatsuUrl(childUrl) Then

            If Not HojinIsNonHtmlFile(childUrl) Then

                If StrComp( _
                    childUrl, _
                    pageUrl, _
                    vbTextCompare) <> 0 Then

                    result = _
                        HojinSearchChildPages( _
                            childUrl, _
                            tsutatsuNo, _
                            depth + 1)

                    If result <> "" Then

                        HojinSearchChildPages = _
                            result

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
' 例：
'
' 1-2-1 本文
'
' (組織変更等の場合の事業年度)
'
' 1-2-2 本文
'
' ↓
'
' 1-2-1 の本文だけ取得
'========================================================

Private Function HojinExtractTsutatsuEntry( _
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

    text = _
        HojinHtmlToText(html)

    If text = "" Then Exit Function

    '----------------------------------------------------
    ' 現在の通達番号
    '----------------------------------------------------

    startPos = _
        HojinFindNumberFlexible( _
            text, _
            tsutatsuNo)

    If startPos <= 0 Then Exit Function

    '----------------------------------------------------
    ' 通達番号直後から本文開始
    '----------------------------------------------------

    startPos = _
        HojinFindBodyStart( _
            text, _
            startPos)

    If startPos <= 0 Then Exit Function

    '----------------------------------------------------
    ' 次の通達番号
    '----------------------------------------------------

    nextNoPos = _
        HojinFindNextTsutatsuNumber( _
            text, _
            startPos)

    If nextNoPos > 0 Then

        '------------------------------------------------
        ' 次の通達番号直前の見出しを探す
        '------------------------------------------------

        endPos = _
            HojinFindHeadingBefore( _
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
            HojinFindFooterPosition( _
                text, _
                startPos)

        If endPos <= startPos Then

            body = _
                Mid$(text, startPos)

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

    body = _
        HojinCleanText(body)

    If body = "" Then Exit Function

    HojinExtractTsutatsuEntry = body

End Function


'========================================================
' 通達番号検索
'========================================================

Private Function HojinFindNumberFlexible( _
    ByVal text As String, _
    ByVal tsutatsuNo As String) As Long

    Dim target As String

    Dim i As Long
    Dim j As Long
    Dim targetIndex As Long

    Dim c As String
    Dim normalizedChar As String

    target = _
        HojinNormalizeTsutatsuNo(tsutatsuNo)

    If target = "" Then Exit Function

    For i = 1 To Len(text)

        If HojinIsDigitCharacter( _
            Mid$(text, i, 1)) Then

            targetIndex = 1
            j = i

            Do While _
                j <= Len(text) And _
                targetIndex <= Len(target)

                c = _
                    Mid$(text, j, 1)

                normalizedChar = _
                    HojinNormalizeNumberCharacter(c)

                If normalizedChar = "" Then Exit Do

                If normalizedChar <> _
                    Mid$(target, targetIndex, 1) Then

                    Exit Do

                End If

                targetIndex = _
                    targetIndex + 1

                j = j + 1

            Loop

            If targetIndex > Len(target) Then

                If HojinValidNumberBoundary( _
                    text, _
                    i, _
                    j - 1) Then

                    HojinFindNumberFlexible = i

                    Exit Function

                End If

            End If

        End If

    Next i

End Function


'========================================================
' 通達番号境界チェック
'========================================================

Private Function HojinValidNumberBoundary( _
    ByVal text As String, _
    ByVal startPos As Long, _
    ByVal endPos As Long) As Boolean

    Dim c As String

    If startPos > 1 Then

        c = _
            Mid$( _
                text, _
                startPos - 1, _
                1)

        If HojinIsNumberCharacter(c) Then
            Exit Function
        End If

    End If

    If endPos < Len(text) Then

        c = _
            Mid$( _
                text, _
                endPos + 1, _
                1)

        If HojinIsDigitCharacter(c) Then
            Exit Function
        End If

    End If

    HojinValidNumberBoundary = True

End Function


'========================================================
' 次の通達番号を探す
'========================================================

Private Function HojinFindNextTsutatsuNumber( _
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

                c = _
                    Mid$(text, i, 1)

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

        If HojinIsDigitCharacter( _
            Mid$(text, i, 1)) Then

            numberText = ""
            digitCount = 0
            hyphenCount = 0

            j = i

            Do While j <= Len(text)

                c = _
                    Mid$(text, j, 1)

                normalizedChar = _
                    HojinNormalizeNumberCharacter(c)

                If normalizedChar = "" Then Exit Do

                If _
                    normalizedChar >= "0" And _
                    normalizedChar <= "9" Then

                    numberText = _
                        numberText & _
                        normalizedChar

                    digitCount = _
                        digitCount + 1

                ElseIf normalizedChar = "-" Then

                    numberText = _
                        numberText & "-"

                    hyphenCount = _
                        hyphenCount + 1

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

                If HojinLooksLikeTsutatsuNo( _
                    numberText) Then

                    HojinFindNextTsutatsuNumber = i

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

Private Function HojinLooksLikeTsutatsuNo( _
    ByVal numberText As String) As Boolean

    Dim parts() As String

    If numberText = "" Then Exit Function

    parts = _
        Split(numberText, "-")

    If UBound(parts) <> 2 Then Exit Function

    If parts(0) = "" Then Exit Function
    If parts(1) = "" Then Exit Function
    If parts(2) = "" Then Exit Function

    If Not IsNumeric(parts(0)) Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function
    If Not IsNumeric(parts(2)) Then Exit Function

    HojinLooksLikeTsutatsuNo = True

End Function


'========================================================
' 数字1文字正規化
'========================================================

Private Function HojinNormalizeNumberCharacter( _
    ByVal c As String) As String

    If c = "" Then Exit Function

    ' 半角数字
    If _
        c >= "0" And _
        c <= "9" Then

        HojinNormalizeNumberCharacter = c

        Exit Function

    End If

    ' 全角数字
    If _
        c >= "０" And _
        c <= "９" Then

        HojinNormalizeNumberCharacter = _
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

        HojinNormalizeNumberCharacter = "-"

        Exit Function

    End If

End Function


'========================================================
' 数字判定
'========================================================

Private Function HojinIsDigitCharacter( _
    ByVal c As String) As Boolean

    If c = "" Then Exit Function

    If _
        c >= "0" And _
        c <= "9" Then

        HojinIsDigitCharacter = True

        Exit Function

    End If

    If _
        c >= "０" And _
        c <= "９" Then

        HojinIsDigitCharacter = True

        Exit Function

    End If

End Function


'========================================================
' 数字・ハイフン判定
'========================================================

Private Function HojinIsNumberCharacter( _
    ByVal c As String) As Boolean

    If HojinIsDigitCharacter(c) Then

        HojinIsNumberCharacter = True

        Exit Function

    End If

    If _
        c = "-" Or _
        c = "－" Or _
        c = "‐" Or _
        c = "?" Or _
        c = "ー" Or _
        c = "―" Then

        HojinIsNumberCharacter = True

    End If

End Function


'========================================================
' 次の通達番号直前の見出しを探す
'========================================================

Private Function HojinFindHeadingBefore( _
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
            InStr( _
                i, _
                text, _
                vbLf)

        If _
            lineEnd = 0 Or _
            lineEnd >= nextNoPos Then

            lineEnd = nextNoPos

        End If

        lineText = _
            Mid$( _
                text, _
                lineStart, _
                lineEnd - lineStart)

        trimmedLine = _
            Trim$(lineText)

        '------------------------------------------------
        ' >> を除去
        '------------------------------------------------

        Do While _
            Left$(trimmedLine, 2) = ">>"

            trimmedLine = _
                Trim$( _
                    Mid$( _
                        trimmedLine, _
                        3))

        Loop

        '------------------------------------------------
        ' 見出し判定
        '------------------------------------------------

        If HojinIsHeadingLine(trimmedLine) Then

            bestPos = lineStart

        End If

        If lineEnd >= nextNoPos Then Exit Do

        i = lineEnd + 1

    Loop

    If bestPos > 0 Then

        HojinFindHeadingBefore = bestPos

        Exit Function

    End If

    '----------------------------------------------------
    ' 同一行にくっついている場合
    '----------------------------------------------------

    HojinFindHeadingBefore = _
        HojinFindInlineHeadingBefore( _
            text, _
            startPos, _
            nextNoPos)

End Function


'========================================================
' 同一行の見出しを逆方向から探す
'
' 全角・半角両対応
'========================================================

Private Function HojinFindInlineHeadingBefore( _
    ByVal text As String, _
    ByVal startPos As Long, _
    ByVal nextNoPos As Long) As Long

    Dim p As Long
    Dim q As Long

    Dim candidate As String

    p = nextNoPos - 1

    Do While p >= startPos

        '------------------------------------------------
        ' 全角閉じ括弧
        '------------------------------------------------

        If Mid$(text, p, 1) = "）" Then

            q = _
                InStrRev( _
                    text, _
                    "（", _
                    p)

            If q >= startPos Then

                candidate = _
                    Mid$( _
                        text, _
                        q, _
                        p - q + 1)

                If HojinIsHeadingLine(candidate) Then

                    HojinFindInlineHeadingBefore = q

                    Exit Function

                End If

            End If

        End If

        '------------------------------------------------
        ' 半角閉じ括弧
        '------------------------------------------------

        If Mid$(text, p, 1) = ")" Then

            q = _
                InStrRev( _
                    text, _
                    "(", _
                    p)

            If q >= startPos Then

                candidate = _
                    Mid$( _
                        text, _
                        q, _
                        p - q + 1)

                If HojinIsHeadingLine(candidate) Then

                    HojinFindInlineHeadingBefore = q

                    Exit Function

                End If

            End If

        End If

        p = p - 1

    Loop

End Function


'========================================================
' 見出し行判定
'
' 全角
' （組織変更等の場合の事業年度）
'
' 半角
' (組織変更等の場合の事業年度)
'
' 混在
' （組織変更等の場合の事業年度)
' (組織変更等の場合の事業年度）
'
' すべて対応
'========================================================

Private Function HojinIsHeadingLine( _
    ByVal lineText As String) As Boolean

    Dim firstChar As String
    Dim lastChar As String

    lineText = _
        Trim$(lineText)

    If lineText = "" Then Exit Function

    '----------------------------------------------------
    ' >> を除去
    '----------------------------------------------------

    Do While _
        Left$(lineText, 2) = ">>"

        lineText = _
            Trim$( _
                Mid$( _
                    lineText, _
                    3))

    Loop

    If Len(lineText) < 2 Then Exit Function

    firstChar = _
        Left$(lineText, 1)

    lastChar = _
        Right$(lineText, 1)

    '----------------------------------------------------
    ' 全角・半角の開き括弧
    '----------------------------------------------------

    If _
        firstChar <> "（" And _
        firstChar <> "(" Then

        Exit Function

    End If

    '----------------------------------------------------
    ' 全角・半角の閉じ括弧
    '----------------------------------------------------

    If _
        lastChar <> "）" And _
        lastChar <> ")" Then

        Exit Function

    End If

    '----------------------------------------------------
    ' 注記は除外
    '----------------------------------------------------

    If lineText = "（注）" Then Exit Function

    If lineText = "(注)" Then Exit Function

    HojinIsHeadingLine = True

End Function


'========================================================
' フッター位置
'========================================================

Private Function HojinFindFooterPosition( _
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

    If _
        p1 > 0 And _
        p2 > 0 Then

        If p1 < p2 Then

            HojinFindFooterPosition = p1

        Else

            HojinFindFooterPosition = p2

        End If

        Exit Function

    End If

    If p1 > 0 Then

        HojinFindFooterPosition = p1

        Exit Function

    End If

    If p2 > 0 Then

        HojinFindFooterPosition = p2

        Exit Function

    End If

    HojinFindFooterPosition = _
        Len(text) + 1

End Function


'========================================================
' 通達番号直後から本文開始
'========================================================

Private Function HojinFindBodyStart( _
    ByVal text As String, _
    ByVal numberPos As Long) As Long

    Dim p As Long
    Dim c As String

    p = numberPos

    '----------------------------------------------------
    ' 数字・ハイフンを飛ばす
    '----------------------------------------------------

    Do While p <= Len(text)

        c = _
            Mid$(text, p, 1)

        If HojinIsNumberCharacter(c) Then

            p = p + 1

        Else

            Exit Do

        End If

    Loop

    '----------------------------------------------------
    ' 番号直後の空白・コロンを飛ばす
    '----------------------------------------------------

    Do While p <= Len(text)

        c = _
            Mid$(text, p, 1)

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

    HojinFindBodyStart = p

End Function


'========================================================
' 通達番号正規化
'========================================================

Private Function HojinNormalizeTsutatsuNo( _
    ByVal text As String) As String

    text = Trim$(text)

    On Error Resume Next

    text = _
        StrConv( _
            text, _
            vbNarrow)

    On Error GoTo 0

    text = _
        Replace(text, "－", "-")

    text = _
        Replace(text, "‐", "-")

    text = _
        Replace(text, "?", "-")

    text = _
        Replace(text, "ー", "-")

    text = _
        Replace(text, "―", "-")

    text = _
        Replace(text, " ", "")

    text = _
        Replace(text, "　", "")

    HojinNormalizeTsutatsuNo = text

End Function


'========================================================
' 通達番号一致
'========================================================

Private Function HojinNumberMatches( _
    ByVal text As String, _
    ByVal tsutatsuNo As String) As Boolean

    Dim normalizedText As String
    Dim normalizedNo As String

    normalizedText = _
        HojinNormalizeNumberText(text)

    normalizedNo = _
        HojinNormalizeTsutatsuNo(tsutatsuNo)

    If normalizedNo = "" Then Exit Function

    HojinNumberMatches = _
        InStr( _
            1, _
            normalizedText, _
            normalizedNo, _
            vbBinaryCompare) > 0

End Function


'========================================================
' 検索用文字列正規化
'========================================================

Private Function HojinNormalizeNumberText( _
    ByVal text As String) As String

    On Error Resume Next

    text = _
        StrConv( _
            text, _
            vbNarrow)

    On Error GoTo 0

    text = _
        Replace(text, "－", "-")

    text = _
        Replace(text, "‐", "-")

    text = _
        Replace(text, "?", "-")

    text = _
        Replace(text, "ー", "-")

    text = _
        Replace(text, "―", "-")

    text = _
        Replace(text, " ", "")

    text = _
        Replace(text, "　", "")

    HojinNormalizeNumberText = text

End Function


'========================================================
' HTML → テキスト
'========================================================

Private Function HojinHtmlToText( _
    ByVal html As String) As String

    Dim text As String

    Dim re As Object

    text = html

    Set re = _
        CreateObject( _
            "VBScript.RegExp")

    '----------------------------------------------------
    ' script / style / noscript
    '----------------------------------------------------

    With re

        .Global = True
        .IgnoreCase = True
        .MultiLine = True

        .pattern = _
            "<script[\s\S]*?</script>"

        text = _
            .Replace(text, "")

        .pattern = _
            "<style[\s\S]*?</style>"

        text = _
            .Replace(text, "")

        .pattern = _
            "<noscript[\s\S]*?</noscript>"

        text = _
            .Replace(text, "")

    End With

    '----------------------------------------------------
    ' ブロック開始タグ
    '----------------------------------------------------

    With re

        .Global = True
        .IgnoreCase = True
        .MultiLine = True

        .pattern = _
            "<(h[1-6]|section|article|header|main|p|div|li|tr|table)[^>]*>"

        text = _
            .Replace( _
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

        .pattern = _
            "<br\s*/?>"

        text = _
            .Replace( _
                text, _
                vbCrLf)

        .pattern = _
            "</(h[1-6]|section|article|header|main|p|div|li|tr|table)\s*>"

        text = _
            .Replace( _
                text, _
                vbCrLf)

    End With

    '----------------------------------------------------
    ' HTMLタグ除去
    '----------------------------------------------------

    With re

        .Global = True
        .IgnoreCase = True
        .MultiLine = True

        .pattern = _
            "<[^>]+>"

        text = _
            .Replace( _
                text, _
                "")

    End With

    '----------------------------------------------------
    ' HTMLエンティティ
    '----------------------------------------------------

    text = _
        HojinDecodeHtmlEntities(text)

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

    text = _
        Replace( _
            text, _
            vbCrLf, _
            vbLf)

    text = _
        Replace( _
            text, _
            vbCr, _
            vbLf)

    '----------------------------------------------------
    ' 行末空白
    '----------------------------------------------------

    With re

        .Global = True
        .MultiLine = True

        .pattern = _
            "[ \t]+$"

        text = _
            .Replace( _
                text, _
                "")

    End With

    text = _
        Replace( _
            text, _
            vbLf, _
            vbCrLf)

    HojinHtmlToText = text

End Function


'========================================================
' HTMLエンティティ
'========================================================

Private Function HojinDecodeHtmlEntities( _
    ByVal text As String) As String

    text = _
        Replace( _
            text, _
            "&nbsp;", _
            " ")

    text = _
        Replace( _
            text, _
            "&amp;", _
            "&")

    text = _
        Replace( _
            text, _
            "&lt;", _
            "<")

    text = _
        Replace( _
            text, _
            "&gt;", _
            ">")

    text = _
        Replace( _
            text, _
            "&quot;", _
            """")

    text = _
        Replace( _
            text, _
            "&#39;", _
            "'")

    text = _
        Replace( _
            text, _
            "&#x3000;", _
            "　")

    text = _
        Replace( _
            text, _
            "&#12288;", _
            "　")

    HojinDecodeHtmlEntities = text

End Function


'========================================================
' 本文整理
'========================================================

Private Function HojinCleanText( _
    ByVal text As String) As String

    Dim re As Object

    text = _
        Replace( _
            text, _
            vbCrLf, _
            vbLf)

    text = _
        Replace( _
            text, _
            vbCr, _
            vbLf)

    Set re = _
        CreateObject( _
            "VBScript.RegExp")

    '----------------------------------------------------
    ' >> 単独行削除
    '----------------------------------------------------

    With re

        .Global = True
        .MultiLine = True

        .pattern = _
            "^[ \t]*>>[ \t]*$"

        text = _
            .Replace( _
                text, _
                "")

    End With

    '----------------------------------------------------
    ' 行頭空白
    '----------------------------------------------------

    With re

        .Global = True
        .MultiLine = True

        .pattern = _
            "^[ \t]+"

        text = _
            .Replace( _
                text, _
                "")

    End With

    '----------------------------------------------------
    ' 行末空白
    '----------------------------------------------------

    With re

        .Global = True
        .MultiLine = True

        .pattern = _
            "[ \t]+$"

        text = _
            .Replace( _
                text, _
                "")

    End With

    '----------------------------------------------------
    ' 連続空行を整理
    '----------------------------------------------------

    With re

        .Global = True
        .MultiLine = True

        .pattern = _
            "(\n[ \t]*){3,}"

        text = _
            .Replace( _
                text, _
                vbLf & vbLf)

    End With

    text = _
        Trim$(text)

    text = _
        Replace( _
            text, _
            vbLf, _
            vbCrLf)

    HojinCleanText = text

End Function


'========================================================
' HTMLリンク抽出
'========================================================

Private Function HojinExtractLinks( _
    ByVal html As String) As Collection

    Dim result As New Collection

    Dim re As Object
    Dim matches As Object
    Dim m As Object

    Dim href As String
    Dim linkText As String

    Dim item(0 To 1) As String

    Set re = _
        CreateObject( _
            "VBScript.RegExp")

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

    Set matches = _
        re.Execute(html)

    For Each m In matches

        href = _
            Trim$( _
                CStr(m.SubMatches(0)))

        linkText = _
            CStr(m.SubMatches(1))

        linkText = _
            HojinHtmlToText(linkText)

        linkText = _
            Trim$(linkText)

        href = _
            HojinResolveUrl(href)

        If href <> "" Then

            item(0) = href
            item(1) = linkText

            result.Add item

        End If

    Next m

    Set HojinExtractLinks = result

End Function


'========================================================
' URL解決
'========================================================

Private Function HojinResolveUrl( _
    ByVal href As String) As String

    Dim baseUrl As String
    Dim path As String

    Dim lastSlash As Long

    href = _
        Trim$(href)

    If href = "" Then Exit Function

    If Left$(href, 1) = "#" Then Exit Function

    If LCase$(Left$(href, 11)) = _
        "javascript:" Then Exit Function

    If LCase$(Left$(href, 7)) = _
        "mailto:" Then Exit Function

    If LCase$(Left$(href, 8)) = _
        "https://" Then

        HojinResolveUrl = href

        Exit Function

    End If

    If LCase$(Left$(href, 7)) = _
        "http://" Then

        HojinResolveUrl = href

        Exit Function

    End If

    If Left$(href, 2) = "//" Then

        HojinResolveUrl = _
            "https:" & href

        Exit Function

    End If

    If Left$(href, 1) = "/" Then

        HojinResolveUrl = _
            HOJIN_BASE_URL & href

        Exit Function

    End If

    baseUrl = _
        HOJIN_BASE_URL & _
        HOJIN_BASE_PATH

    lastSlash = _
        InStrRev( _
            baseUrl, _
            "/")

    If lastSlash > 0 Then

        path = _
            Left$( _
                baseUrl, _
                lastSlash) & _
            href

    Else

        path = _
            baseUrl & href

    End If

    HojinResolveUrl = _
        HojinNormalizeUrl(path)

End Function


'========================================================
' URL正規化
'========================================================

Private Function HojinNormalizeUrl( _
    ByVal url As String) As String

    Dim parts() As String

    Dim i As Long

    Dim stack As Collection

    Dim part As String
    Dim result As String

    If url = "" Then Exit Function

    parts = _
        Split(url, "/")

    Set stack = New Collection

    For i = _
        LBound(parts) To _
        UBound(parts)

        part = parts(i)

        If part = ".." Then

            If stack.Count > 0 Then

                stack.Remove _
                    stack.Count

            End If

        ElseIf _
            part <> "." And _
            part <> "" Then

            stack.Add part

        End If

    Next i

    result = _
        "https://www.nta.go.jp"

    For i = 1 To stack.Count

        result = _
            result & _
            "/" & _
            CStr(stack(i))

    Next i

    HojinNormalizeUrl = result

End Function


'========================================================
' 法人税基本通達配下か
'========================================================

Private Function HojinIsTsutatsuUrl( _
    ByVal url As String) As Boolean

    Dim normalized As String

    normalized = _
        LCase$( _
            Trim$(url))

    HojinIsTsutatsuUrl = _
        InStr( _
            1, _
            normalized, _
            LCase$( _
                HOJIN_BASE_URL & _
                HOJIN_BASE_PATH), _
            vbTextCompare) = 1

End Function


'========================================================
' HTMLではないファイルを除外
'========================================================

Private Function HojinIsNonHtmlFile( _
    ByVal url As String) As Boolean

    Dim lowerUrl As String

    lowerUrl = _
        LCase$(url)

    If InStr(lowerUrl, ".xlsx") > 0 Then

        HojinIsNonHtmlFile = True

        Exit Function

    End If

    If InStr(lowerUrl, ".xls") > 0 Then

        HojinIsNonHtmlFile = True

        Exit Function

    End If

    If InStr(lowerUrl, ".pdf") > 0 Then

        HojinIsNonHtmlFile = True

        Exit Function

    End If

    If InStr(lowerUrl, ".doc") > 0 Then

        HojinIsNonHtmlFile = True

        Exit Function

    End If

    If InStr(lowerUrl, ".zip") > 0 Then

        HojinIsNonHtmlFile = True

        Exit Function

    End If

End Function


'========================================================
' HTTP GET
'========================================================

Private Function HojinHttpGetText( _
    ByVal url As String) As String

    Dim http As Object

    Dim bytes() As Byte

    Dim charset As String

    If Not HojinIsTsutatsuUrl(url) Then Exit Function

    If HojinIsNonHtmlFile(url) Then Exit Function

    On Error GoTo ErrHandler

    Set http = _
        CreateObject( _
            "MSXML2.XMLHTTP")

    http.Open _
        "GET", _
        url, _
        False

    http.SetRequestHeader _
        "User-Agent", _
        "Mozilla/5.0"

    http.Send

    If _
        http.Status < 200 Or _
        http.Status >= 300 Then

        Exit Function

    End If

    bytes = _
        http.ResponseBody

    charset = _
        HojinDetectCharset(bytes)

    If charset = "" Then

        charset = _
            "shift_jis"

    End If

    HojinHttpGetText = _
        HojinBytesToText( _
            bytes, _
            charset)

    Exit Function

ErrHandler:

    HojinHttpGetText = ""

End Function


'========================================================
' charset判定
'========================================================

Private Function HojinDetectCharset( _
    ByRef bytes() As Byte) As String

    Dim preview As String

    On Error GoTo ErrHandler

    preview = _
        HojinBytesToText( _
            bytes, _
            "iso-8859-1")

    preview = _
        LCase$(preview)

    If InStr( _
        preview, _
        "charset=utf-8") > 0 Then

        HojinDetectCharset = _
            "utf-8"

        Exit Function

    End If

    If InStr( _
        preview, _
        "charset=""utf-8""") > 0 Then

        HojinDetectCharset = _
            "utf-8"

        Exit Function

    End If

    If InStr( _
        preview, _
        "charset=shift_jis") > 0 Then

        HojinDetectCharset = _
            "shift_jis"

        Exit Function

    End If

    If InStr( _
        preview, _
        "charset=""shift_jis""") > 0 Then

        HojinDetectCharset = _
            "shift_jis"

        Exit Function

    End If

    If InStr( _
        preview, _
        "charset=shift-jis") > 0 Then

        HojinDetectCharset = _
            "shift_jis"

        Exit Function

    End If

    HojinDetectCharset = _
        "shift_jis"

    Exit Function

ErrHandler:

    HojinDetectCharset = _
        "shift_jis"

End Function


'========================================================
' Byte → String
'========================================================

Private Function HojinBytesToText( _
    ByRef bytes() As Byte, _
    ByVal charset As String) As String

    Dim stream As Object

    On Error GoTo ErrHandler

    Set stream = _
        CreateObject( _
            "ADODB.Stream")

    With stream

        .Type = 1

        .Open

        .Write bytes

        .position = 0

        .Type = 2

        .charset = charset

        HojinBytesToText = _
            .ReadText

        .Close

    End With

    Exit Function

ErrHandler:

    HojinBytesToText = ""

End Function

