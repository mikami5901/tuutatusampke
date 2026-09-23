Attribute VB_Name = "Module3"
Option Explicit

'============================================================
' 消費税法基本通達 取得モジュール
'
' Excelシート:
'
'   A1 : 通達名
'   B1 : 消基通
'
'   A2 : 通達番号
'   B2 : 1-1-1
'
' 実行:
'
'   消費税通達取得
'
' 結果:
'
'   B4 に取得した通達本文を表示
'
'============================================================


'============================================================
' 国税庁
'============================================================

Private Const SHOUHI_BASE_URL As String = _
    "https://www.nta.go.jp"

Private Const SHOUHI_TOC_PATH As String = _
    "/law/tsutatsu/kihon/shohi/20230930/01.htm"

Private Const SHOUHI_MAX_DEPTH As Long = 3


Private gShouhiVisited As Object

Private gShouhiFoundUrl As String


'============================================================
' Excelから実行
'============================================================

Public Sub 消費税通達取得()

    Dim ws As Worksheet

    Dim tsutatsuName As String
    Dim tsutatsuNumber As String

    Dim result As String

    Set ws = ActiveSheet

    '--------------------------------------------------------
    ' 入力
    '--------------------------------------------------------

    tsutatsuName = Trim$( _
        CStr(ws.Range("B1").Value))

    tsutatsuNumber = Trim$( _
        CStr(ws.Range("B2").Value))

    '--------------------------------------------------------
    ' 通達名
    '--------------------------------------------------------

    If Len(tsutatsuName) = 0 Then

        tsutatsuName = "消基通"

    End If

    '--------------------------------------------------------
    ' 通達番号
    '--------------------------------------------------------

    If Len(tsutatsuNumber) = 0 Then

        MsgBox _
            "B2に通達番号を入力してください。" & _
            vbCrLf & vbCrLf & _
            "例：1-1-1", _
            vbExclamation

        Exit Sub

    End If

    Application.ScreenUpdating = False

    Application.StatusBar = _
        "消費税法基本通達を取得しています..."

    On Error GoTo ErrorHandler

    '--------------------------------------------------------
    ' 結果消去
    '--------------------------------------------------------

    ws.Range("B4:B10000").ClearContents

    '--------------------------------------------------------
    ' 通達取得
    '--------------------------------------------------------

    result = ShouhiGetTsutatsu( _
                tsutatsuNumber)

    '--------------------------------------------------------
    ' Excelへ出力
    '--------------------------------------------------------

    ws.Range("B4").Value = result

    ws.Range("B4").WrapText = True

    ws.Columns("B").ColumnWidth = 100

    ws.Rows("4:10000").RowHeight = 18

    Application.StatusBar = False

    Application.ScreenUpdating = True

    MsgBox _
        "取得しました。", _
        vbInformation

    Exit Sub


ErrorHandler:

    Application.StatusBar = False

    Application.ScreenUpdating = True

    MsgBox _
        "消費税法基本通達の取得に失敗しました。" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbExclamation

End Sub


'============================================================
' メイン
'============================================================

Private Function ShouhiGetTsutatsu( _
    ByVal tsutatsuNumber As String) As String

    Dim tocHtml As String

    Dim candidateUrl As String

    Dim body As String

    Debug.Print String(70, "=")
    Debug.Print "SHOUHI TSUTATSU"
    Debug.Print "NUMBER = " & tsutatsuNumber
    Debug.Print String(70, "=")

    '--------------------------------------------------------
    ' TOC取得
    '--------------------------------------------------------

    Debug.Print "GET TOC..."

    tocHtml = ShouhiHttpGetText( _
                SHOUHI_BASE_URL & _
                SHOUHI_TOC_PATH)

    Debug.Print "TOC OK"
    Debug.Print _
        "TOC LENGTH = "; _
        Len(tocHtml)

    '--------------------------------------------------------
    ' まず通達番号からページを推定
    '
    ' 例:
    '
    ' 1-1-1
    ' ↓
    ' /01/01.htm
    '
    ' 5-1-1
    ' ↓
    ' /05/01.htm
    '
    ' 10-1-1
    ' ↓
    ' /10/01.htm
    '
    '--------------------------------------------------------

    candidateUrl = _
        ShouhiMakeCandidateUrl( _
            tsutatsuNumber)

    If Len(candidateUrl) > 0 Then

        Debug.Print
        Debug.Print "CANDIDATE:"
        Debug.Print candidateUrl

        body = _
            ShouhiGetEntryFromPage( _
                candidateUrl, _
                tsutatsuNumber)

        If Len(body) > 0 Then

            ShouhiGetTsutatsu = _
                ShouhiBuildResult( _
                    tsutatsuNumber, _
                    body, _
                    candidateUrl)

            Exit Function

        End If

    End If

    '--------------------------------------------------------
    ' TOCから候補ページを探す
    '--------------------------------------------------------

    Debug.Print
    Debug.Print "SEARCH TOC..."

    candidateUrl = _
        ShouhiFindPageFromToc( _
            tocHtml, _
            tsutatsuNumber)

    If Len(candidateUrl) > 0 Then

        Debug.Print "TOC CANDIDATE:"
        Debug.Print candidateUrl

        body = _
            ShouhiGetEntryFromPage( _
                candidateUrl, _
                tsutatsuNumber)

        If Len(body) > 0 Then

            ShouhiGetTsutatsu = _
                ShouhiBuildResult( _
                    tsutatsuNumber, _
                    body, _
                    candidateUrl)

            Exit Function

        End If

    End If

    '--------------------------------------------------------
    ' 階層探索
    '--------------------------------------------------------

    Debug.Print
    Debug.Print "SEARCH CHILD PAGES..."

    Set gShouhiVisited = _
        CreateObject("Scripting.Dictionary")

    gShouhiFoundUrl = vbNullString

    body = _
        ShouhiSearchChildPages( _
            tocHtml, _
            SHOUHI_TOC_PATH, _
            tsutatsuNumber, _
            0)

    If Len(body) > 0 Then

        ShouhiGetTsutatsu = _
            ShouhiBuildResult( _
                tsutatsuNumber, _
                body, _
                gShouhiFoundUrl)

        Exit Function

    End If

    Err.Raise _
        vbObjectError + 2000, , _
        "消費税法基本通達 " & _
        tsutatsuNumber & _
        " が見つかりませんでした。"

End Function


'============================================================
' 通達番号から候補URLを作る
'
' 1-1-1
' ↓
' 01/01.htm
'
' 5-1-1
' ↓
' 05/01.htm
'
' 10-1-1
' ↓
' 10/01.htm
'============================================================

Private Function ShouhiMakeCandidateUrl( _
    ByVal number As String) As String

    Dim parts() As String

    Dim chapterNo As Long
    Dim sectionNo As Long

    Dim chapterText As String
    Dim sectionText As String

    number = _
        ShouhiNormalizeNumber(number)

    parts = Split(number, "-")

    If UBound(parts) < 1 Then Exit Function

    chapterNo = _
        Val(parts(0))

    sectionNo = _
        Val(parts(1))

    If chapterNo <= 0 Then Exit Function

    If sectionNo <= 0 Then Exit Function

    chapterText = _
        Format$(chapterNo, "00")

    sectionText = _
        Format$(sectionNo, "00")

    ShouhiMakeCandidateUrl = _
        SHOUHI_BASE_URL & _
        "/law/tsutatsu/kihon/shohi/20230930/" & _
        chapterText & "/" & _
        sectionText & ".htm"

End Function


'============================================================
' 指定ページから取得
'============================================================

Private Function ShouhiGetEntryFromPage( _
    ByVal url As String, _
    ByVal number As String) As String

    Dim html As String

    Dim result As String

    On Error GoTo ErrorHandler

    Debug.Print "GET PAGE:"
    Debug.Print url

    html = _
        ShouhiHttpGetText(url)

    Debug.Print _
        "PAGE LENGTH = "; _
        Len(html)

    result = _
        ShouhiExtractTsutatsuEntry( _
            html, _
            number)

    If Len(result) > 0 Then

        Debug.Print "ENTRY FOUND"

        ShouhiGetEntryFromPage = _
            result

    Else

        Debug.Print "ENTRY NOT FOUND"

    End If

    Exit Function


ErrorHandler:

    Debug.Print _
        "PAGE ERROR: " & _
        Err.Description

End Function


'============================================================
' TOCからページを探す
'============================================================

Private Function ShouhiFindPageFromToc( _
    ByVal html As String, _
    ByVal number As String) As String

    Dim links As Collection

    Dim v As Variant

    Dim targetChapter As String
    Dim targetSection As String

    Dim text As String
    Dim href As String

    Dim score As Long
    Dim bestScore As Long

    Dim bestUrl As String

    targetChapter = _
        ShouhiGetNumberPart( _
            number, _
            1)

    targetSection = _
        ShouhiGetNumberPart( _
            number, _
            2)

    Set links = _
        ShouhiExtractLinkInfo( _
            html, _
            SHOUHI_TOC_PATH)

    Debug.Print _
        "TOC LINKS = "; _
        links.Count

    For Each v In links

        href = CStr(v(0))
        text = CStr(v(1))

        score = 0

        '--------------------------------------------
        ' 第X章
        '--------------------------------------------

        If InStr( _
            1, _
            text, _
            "第" & targetChapter & "章", _
            vbTextCompare) > 0 Then

            score = score + 30

        End If

        '--------------------------------------------
        ' 第X節
        '--------------------------------------------

        If InStr( _
            1, _
            text, _
            "第" & targetSection & "節", _
            vbTextCompare) > 0 Then

            score = score + 50

        End If

        '--------------------------------------------
        ' URL
        '--------------------------------------------

        If InStr( _
            1, _
            href, _
            "/" & _
            Format$( _
                Val(targetChapter), _
                "00") & "/", _
            vbTextCompare) > 0 Then

            score = score + 20

        End If

        If score > bestScore Then

            bestScore = score
            bestUrl = href

        End If

    Next v

    If bestScore > 0 Then

        ShouhiFindPageFromToc = _
            bestUrl

    End If

End Function


'============================================================
' 階層ページ探索
'============================================================

Private Function ShouhiSearchChildPages( _
    ByVal html As String, _
    ByVal currentPath As String, _
    ByVal number As String, _
    ByVal depth As Long) As String

    Dim links As Collection

    Dim v As Variant

    Dim href As String
    Dim childHtml As String

    Dim result As String

    If depth > SHOUHI_MAX_DEPTH Then Exit Function

    Set links = _
        ShouhiExtractLinks( _
            html, _
            currentPath)

    Debug.Print _
        String(depth * 2, " ") & _
        "DEPTH=" & depth & _
        " LINKS=" & links.Count

    For Each v In links

        href = CStr(v)

        If gShouhiVisited.Exists(href) Then

            GoTo ContinueLoop

        End If

        gShouhiVisited.Add _
            href, _
            True

        Debug.Print _
            String(depth * 2, " ") & _
            "TRY: " & href

        On Error Resume Next

        childHtml = _
            ShouhiHttpGetText(href)

        If Err.number <> 0 Then

            Err.Clear

            On Error GoTo 0

            GoTo ContinueLoop

        End If

        On Error GoTo 0

        '----------------------------------------------------
        ' ページ自身
        '----------------------------------------------------

        result = _
            ShouhiExtractTsutatsuEntry( _
                childHtml, _
                number)

        If Len(result) > 0 Then

            Debug.Print
            Debug.Print "***** FOUND *****"
            Debug.Print href
            Debug.Print

            gShouhiFoundUrl = href

            ShouhiSearchChildPages = _
                result

            Exit Function

        End If

        '----------------------------------------------------
        ' 下位ページ
        '----------------------------------------------------

        result = _
            ShouhiSearchChildPages( _
                childHtml, _
                href, _
                number, _
                depth + 1)

        If Len(result) > 0 Then

            ShouhiSearchChildPages = _
                result

            Exit Function

        End If

ContinueLoop:

    Next v

End Function


'============================================================
' ★ 通達本文抽出
'
' HTMLのタグ構造に依存しない。
'
' HTML
' ↓
' プレーンテキスト
' ↓
' 指定番号検索
' ↓
' 次の通達番号まで
'============================================================

Private Function ShouhiExtractTsutatsuEntry( _
    ByVal html As String, _
    ByVal targetNumber As String) As String

    Dim plain As String

    Dim p As Long
    Dim nextP As Long

    Dim target As String
    Dim nextNumber As String

    Debug.Print "TEXTIFY..."

    plain = _
        ShouhiHtmlToText(html)

    Debug.Print _
        "TEXT LENGTH = "; _
        Len(plain)

    target = _
        ShouhiNormalizeNumber( _
            targetNumber)

    Debug.Print _
        "TARGET=[" & target & "]"

    '--------------------------------------------------------
    ' 通常検索
    '--------------------------------------------------------

    p = InStr( _
            1, _
            plain, _
            target, _
            vbTextCompare)

    '--------------------------------------------------------
    ' 空白等が入るケース
    '--------------------------------------------------------

    If p = 0 Then

        p = _
            ShouhiFindNumberFlexible( _
                plain, _
                targetNumber)

    End If

    If p = 0 Then

        Debug.Print "NUMBER NOT FOUND"

        Exit Function

    End If

    Debug.Print _
        "NUMBER FOUND AT "; _
        p

    '--------------------------------------------------------
    ' 次の通達番号
    '--------------------------------------------------------

    nextNumber = _
        ShouhiGetNextNumber( _
            targetNumber)

    nextP = _
        ShouhiFindNumberFlexible( _
            Mid$( _
                plain, _
                p + Len(target) + 1), _
            nextNumber)

    If nextP > 0 Then

        nextP = _
            p + _
            Len(target) + _
            nextP

    Else

        nextP = _
            Len(plain) + 1

    End If

    '--------------------------------------------------------
    ' 本文
    '--------------------------------------------------------

    ShouhiExtractTsutatsuEntry = _
        ShouhiCleanBody( _
            Mid$( _
                plain, _
                p, _
                nextP - p))

End Function


'============================================================
' 柔軟な番号検索
'============================================================

Private Function ShouhiFindNumberFlexible( _
    ByVal text As String, _
    ByVal number As String) As Long

    Dim re As Object

    Dim matches As Object

    Dim parts() As String

    Dim pattern As String

    number = _
        ShouhiNormalizeNumber(number)

    parts = _
        Split(number, "-")

    If UBound(parts) < 1 Then Exit Function

    pattern = _
        parts(0) & _
        "\s*[-－??ー]\s*" & _
        parts(1)

    If UBound(parts) >= 2 Then

        pattern = _
            pattern & _
            "\s*[-－??ー]\s*" & _
            parts(2)

    End If

    Set re = _
        CreateObject( _
            "VBScript.RegExp")

    re.Global = False
    re.IgnoreCase = True
    re.MultiLine = True

    re.pattern = pattern

    Set matches = _
        re.Execute(text)

    If matches.Count > 0 Then

        ShouhiFindNumberFlexible = _
            matches(0).FirstIndex + 1

    End If

End Function


'============================================================
' 次の通達番号
'
' 1-1-1 → 1-1-2
' 5-1-1 → 5-1-2
'============================================================

Private Function ShouhiGetNextNumber( _
    ByVal number As String) As String

    Dim parts() As String

    Dim lastIndex As Long

    Dim n As Long

    number = _
        ShouhiNormalizeNumber( _
            number)

    parts = _
        Split(number, "-")

    lastIndex = _
        UBound(parts)

    n = _
        Val(parts(lastIndex))

    parts(lastIndex) = _
        CStr(n + 1)

    ShouhiGetNextNumber = _
        Join(parts, "-")

End Function


'============================================================
' 番号のn番目
'============================================================

Private Function ShouhiGetNumberPart( _
    ByVal number As String, _
    ByVal index As Long) As String

    Dim parts() As String

    number = _
        ShouhiNormalizeNumber( _
            number)

    parts = _
        Split(number, "-")

    If index <= _
        UBound(parts) + 1 Then

        ShouhiGetNumberPart = _
            parts(index - 1)

    End If

End Function


'============================================================
' HTML → TEXT
'============================================================

Private Function ShouhiHtmlToText( _
    ByVal html As String) As String

    Dim re As Object

    '--------------------------------------------------------
    ' 改行
    '--------------------------------------------------------

    html = Replace( _
        html, _
        "<br>", _
        vbCrLf, , , _
        vbTextCompare)

    html = Replace( _
        html, _
        "<br/>", _
        vbCrLf, , , _
        vbTextCompare)

    html = Replace( _
        html, _
        "<br />", _
        vbCrLf, , , _
        vbTextCompare)

    html = Replace( _
        html, _
        "</p>", _
        vbCrLf, , , _
        vbTextCompare)

    html = Replace( _
        html, _
        "</div>", _
        vbCrLf, , , _
        vbTextCompare)

    html = Replace( _
        html, _
        "</li>", _
        vbCrLf, , , _
        vbTextCompare)

    html = Replace( _
        html, _
        "</h1>", _
        vbCrLf, , , _
        vbTextCompare)

    html = Replace( _
        html, _
        "</h2>", _
        vbCrLf, , , _
        vbTextCompare)

    html = Replace( _
        html, _
        "</h3>", _
        vbCrLf, , , _
        vbTextCompare)

    '--------------------------------------------------------
    ' SCRIPT / STYLE
    '--------------------------------------------------------

    Set re = _
        CreateObject( _
            "VBScript.RegExp")

    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True

    re.pattern = _
        "<script[^>]*>[\s\S]*?</script>"

    html = _
        re.Replace( _
            html, _
            "")

    re.pattern = _
        "<style[^>]*>[\s\S]*?</style>"

    html = _
        re.Replace( _
            html, _
            "")

    '--------------------------------------------------------
    ' HTMLタグ
    '--------------------------------------------------------

    re.pattern = _
        "<[^>]*>"

    html = _
        re.Replace( _
            html, _
            "")

    '--------------------------------------------------------
    ' Entity
    '--------------------------------------------------------

    html = _
        ShouhiDecodeHtml(html)

    ShouhiHtmlToText = html

End Function


'============================================================
' HTML Entity
'============================================================

Private Function ShouhiDecodeHtml( _
    ByVal s As String) As String

    s = Replace(s, "&amp;", "&")
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quot;", """")
    s = Replace(s, "&#39;", "'")
    s = Replace(s, "&nbsp;", " ")
    s = Replace(s, "&emsp;", "　")
    s = Replace(s, "&ensp;", " ")

    ShouhiDecodeHtml = s

End Function


'============================================================
' HTTP取得
'
' UTF-8 / Shift_JISを両方試す
'============================================================

Private Function ShouhiHttpGetText( _
    ByVal url As String) As String

    Dim http As Object

    Dim bytes As Variant

    Set http = _
        CreateObject( _
            "WinHttp.WinHttpRequest.5.1")

    http.Open _
        "GET", _
        url, _
        False

    http.SetTimeouts _
        30000, _
        30000, _
        30000, _
        30000

    http.SetRequestHeader _
        "User-Agent", _
        "Mozilla/5.0"

    http.Send

    If http.Status < 200 Or _
       http.Status >= 300 Then

        Err.Raise _
            vbObjectError + 2100, , _
            "HTTP ERROR " & _
            http.Status & _
            vbCrLf & _
            url

    End If

    bytes = _
        http.ResponseBody

    ShouhiHttpGetText = _
        ShouhiDecodeResponse( _
            bytes)

End Function


'============================================================
' ResponseBodyデコード
'============================================================

Private Function ShouhiDecodeResponse( _
    ByVal bytes As Variant) As String

    Dim utf8Text As String

    Dim sjisText As String

    utf8Text = _
        ShouhiDecodeBytes( _
            bytes, _
            "utf-8")

    sjisText = _
        ShouhiDecodeBytes( _
            bytes, _
            "shift_jis")

    '--------------------------------------------------------
    ' 日本語ページとして自然な方
    '--------------------------------------------------------

    If ShouhiLooksLikeNta( _
        utf8Text) Then

        ShouhiDecodeResponse = _
            utf8Text

    ElseIf ShouhiLooksLikeNta( _
        sjisText) Then

        ShouhiDecodeResponse = _
            sjisText

    Else

        ShouhiDecodeResponse = _
            sjisText

    End If

End Function


'============================================================
' バイト列 → 文字列
'============================================================

Private Function ShouhiDecodeBytes( _
    ByVal bytes As Variant, _
    ByVal charset As String) As String

    Dim stm As Object

    Set stm = _
        CreateObject( _
            "ADODB.Stream")

    stm.Type = 1

    stm.Open

    stm.Write bytes

    stm.Position = 0

    stm.Type = 2

    stm.charset = charset

    ShouhiDecodeBytes = _
        stm.ReadText

    stm.Close

    Set stm = Nothing

End Function


'============================================================
' NTA HTML判定
'============================================================

Private Function ShouhiLooksLikeNta( _
    ByVal text As String) As Boolean

    Dim score As Long

    If InStr( _
        1, _
        text, _
        "国税庁", _
        vbTextCompare) > 0 Then

        score = score + 3

    End If

    If InStr( _
        1, _
        text, _
        "消費税", _
        vbTextCompare) > 0 Then

        score = score + 3

    End If

    If InStr( _
        1, _
        text, _
        "第1章", _
        vbTextCompare) > 0 Then

        score = score + 1

    End If

    If InStr( _
        1, _
        text, _
        "<html", _
        vbTextCompare) > 0 Then

        score = score + 1

    End If

    ShouhiLooksLikeNta = _
        (score >= 3)

End Function


'============================================================
' リンク抽出
'============================================================

Private Function ShouhiExtractLinks( _
    ByVal html As String, _
    ByVal basePath As String) As Collection

    Dim result As New Collection

    Dim info As Collection

    Dim v As Variant

    Set info = _
        ShouhiExtractLinkInfo( _
            html, _
            basePath)

    For Each v In info

        On Error Resume Next

        result.Add _
            CStr(v(0)), _
            CStr(v(0))

        On Error GoTo 0

    Next v

    Set ShouhiExtractLinks = _
        result

End Function


'============================================================
' href + text
'============================================================

Private Function ShouhiExtractLinkInfo( _
    ByVal html As String, _
    ByVal basePath As String) As Collection

    Dim result As New Collection

    Dim re As Object

    Dim matches As Object

    Dim m As Object

    Dim href As String

    Dim text As String

    Dim item(0 To 1) As Variant

    Set re = _
        CreateObject( _
            "VBScript.RegExp")

    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True

    re.pattern = _
        "<a[^>]*href\s*=\s*[""]([^""]+)[""][^>]*>" & _
        "([\s\S]*?)</a>"

    Set matches = _
        re.Execute(html)

    For Each m In matches

        href = _
            Trim$( _
                m.SubMatches(0))

        text = _
            ShouhiStripTags( _
                m.SubMatches(1))

        text = _
            ShouhiDecodeHtml( _
                text)

        text = _
            ShouhiCleanText( _
                text)

        If Len(href) = 0 Then
            GoTo ContinueLoop
        End If

        If Left$(href, 1) = "#" Then
            GoTo ContinueLoop
        End If

        href = _
            ShouhiResolveUrl( _
                href, _
                basePath)

        If Not ShouhiIsNtaUrl(href) Then
            GoTo ContinueLoop
        End If

        item(0) = href
        item(1) = text

        On Error Resume Next

        result.Add _
            item, _
            href

        On Error GoTo 0

ContinueLoop:

    Next m

    Set ShouhiExtractLinkInfo = _
        result

End Function


'============================================================
' URL解決
'============================================================

Private Function ShouhiResolveUrl( _
    ByVal href As String, _
    ByVal basePath As String) As String

    Dim baseDir As String

    Dim p As Long

    href = _
        Trim$(href)

    If LCase$(Left$(href, 4)) = _
        "http" Then

        ShouhiResolveUrl = _
            href

        Exit Function

    End If

    If Left$(href, 1) = "/" Then

        ShouhiResolveUrl = _
            SHOUHI_BASE_URL & _
            href

        Exit Function

    End If

    p = _
        InStrRev( _
            basePath, _
            "/")

    If p > 0 Then

        baseDir = _
            Left$( _
                basePath, _
                p)

    End If

    ShouhiResolveUrl = _
        SHOUHI_BASE_URL & _
        ShouhiNormalizeRelative( _
            baseDir & href)

End Function


'============================================================
' ../ 処理
'============================================================

Private Function ShouhiNormalizeRelative( _
    ByVal path As String) As String

    Dim parts() As String

    Dim stack() As String

    Dim i As Long

    Dim n As Long

    Dim part As String

    Dim result As String

    parts = _
        Split(path, "/")

    ReDim stack( _
        0 To UBound(parts))

    For i = _
        LBound(parts) To _
        UBound(parts)

        part = _
            parts(i)

        If part = ".." Then

            If n > 0 Then
                n = n - 1
            End If

        ElseIf part <> "" _
            And part <> "." Then

            stack(n) = part

            n = n + 1

        End If

    Next i

    result = "/"

    For i = 0 To n - 1

        result = _
            result & _
            stack(i)

        If i < n - 1 Then

            result = _
                result & "/"

        End If

    Next i

    ShouhiNormalizeRelative = _
        result

End Function


'============================================================
' NTA URL判定
'============================================================

Private Function ShouhiIsNtaUrl( _
    ByVal url As String) As Boolean

    ShouhiIsNtaUrl = _
        InStr( _
            1, _
            url, _
            "https://www.nta.go.jp/", _
            vbTextCompare) = 1

End Function


'============================================================
' 番号正規化
'============================================================

Private Function ShouhiNormalizeNumber( _
    ByVal s As String) As String

    On Error Resume Next

    s = _
        StrConv( _
            s, _
            vbNarrow)

    On Error GoTo 0

    s = Replace(s, "－", "-")
    s = Replace(s, "?", "-")
    s = Replace(s, "?", "-")
    s = Replace(s, "ー", "-")

    s = Replace(s, "　", "")

    ShouhiNormalizeNumber = _
        Trim$(s)

End Function


'============================================================
' 本文整理
'============================================================

Private Function ShouhiCleanBody( _
    ByVal s As String) As String

    s = Replace( _
        s, _
        ChrW(160), _
        " ")

    Do While InStr( _
        s, _
        vbCrLf & _
        vbCrLf & _
        vbCrLf) > 0

        s = Replace( _
            s, _
            vbCrLf & _
            vbCrLf & _
            vbCrLf, _
            vbCrLf & _
            vbCrLf)

    Loop

    ShouhiCleanBody = _
        Trim$(s)

End Function


'============================================================
' テキスト整理
'============================================================

Private Function ShouhiCleanText( _
    ByVal s As String) As String

    s = Replace( _
        s, _
        ChrW(160), _
        " ")

    Do While InStr( _
        s, _
        "  ") > 0

        s = Replace( _
            s, _
            "  ", _
            " ")

    Loop

    ShouhiCleanText = _
        Trim$(s)

End Function


'============================================================
' HTMLタグ除去
'============================================================

Private Function ShouhiStripTags( _
    ByVal html As String) As String

    Dim re As Object

    Set re = _
        CreateObject( _
            "VBScript.RegExp")

    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True

    re.pattern = _
        "<script[^>]*>[\s\S]*?</script>"

    html = _
        re.Replace( _
            html, _
            "")

    re.pattern = _
        "<style[^>]*>[\s\S]*?</style>"

    html = _
        re.Replace( _
            html, _
            "")

    re.pattern = _
        "<[^>]*>"

    ShouhiStripTags = _
        re.Replace( _
            html, _
            "")

End Function


'============================================================
' 結果
'============================================================

Private Function ShouhiBuildResult( _
    ByVal number As String, _
    ByVal body As String, _
    ByVal url As String) As String

    ShouhiBuildResult = _
        "【消費税法基本通達 " & _
        number & "】" & _
        vbCrLf & _
        body & _
        vbCrLf & _
        vbCrLf & _
        "出典:" & _
        vbCrLf & _
        url

End Function

