Attribute VB_Name = "Module4"
Option Explicit

'============================================================
' 所得税基本通達 取得モジュール
'
' Excel
'
'   A1：通達名
'   B1：所基通
'
'   A2：通達番号
'   B2：2-1
'
' 実行：
'
'   所得税通達取得
'
' 結果：
'
'   B4
'
'============================================================


'============================================================
' 国税庁
'============================================================

Private Const SHOTOKU_BASE_URL As String = _
    "https://www.nta.go.jp"

Private Const SHOTOKU_TOC_PATH As String = _
    "/law/tsutatsu/kihon/shotoku/01.htm"

Private Const SHOTOKU_MAX_DEPTH As Long = 2


Private gShotokuVisited As Object

Private gShotokuFoundUrl As String


'============================================================
' Excelから実行
'============================================================

Public Sub 所得税通達取得()

    Dim ws As Worksheet

    Dim tsutatsuName As String
    Dim tsutatsuNumber As String

    Dim result As String

    Set ws = ActiveSheet

    '--------------------------------------------------------
    ' 入力
    '--------------------------------------------------------

    tsutatsuName = _
        Trim$(CStr(ws.Range("B1").Value))

    tsutatsuNumber = _
        Trim$(CStr(ws.Range("B2").Value))

    '--------------------------------------------------------
    ' 通達名
    '--------------------------------------------------------

    If Len(tsutatsuName) = 0 Then

        tsutatsuName = "所基通"

    End If

    '--------------------------------------------------------
    ' 通達番号
    '--------------------------------------------------------

    If Len(tsutatsuNumber) = 0 Then

        MsgBox _
            "B2に通達番号を入力してください。" & _
            vbCrLf & vbCrLf & _
            "例：2-1", _
            vbExclamation

        Exit Sub

    End If

    Application.ScreenUpdating = False

    Application.StatusBar = _
        "所得税基本通達を取得しています..."

    On Error GoTo ErrorHandler

    '--------------------------------------------------------
    ' 結果消去
    '--------------------------------------------------------

    ws.Range("B4:B10000").ClearContents

    '--------------------------------------------------------
    ' 取得
    '--------------------------------------------------------

    result = _
        ShotokuGetTsutatsu( _
            tsutatsuNumber)

    '--------------------------------------------------------
    ' 出力
    '--------------------------------------------------------

    ws.Range("B4").Value = result

    ws.Range("B4").WrapText = True

    ws.Columns("B").ColumnWidth = 100

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
        "所得税基本通達の取得に失敗しました。" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbExclamation

End Sub


'============================================================
' メイン
'============================================================

Private Function ShotokuGetTsutatsu( _
    ByVal tsutatsuNumber As String) As String

    Dim tocHtml As String

    Dim candidateUrl As String

    Dim body As String

    Debug.Print String(70, "=")
    Debug.Print "SHOTOKU TSUTATSU"
    Debug.Print "NUMBER = " & tsutatsuNumber
    Debug.Print String(70, "=")

    '--------------------------------------------------------
    ' TOC
    '--------------------------------------------------------

    Debug.Print "GET TOC..."

    tocHtml = _
        ShotokuHttpGetText( _
            SHOTOKU_BASE_URL & _
            SHOTOKU_TOC_PATH)

    Debug.Print "TOC OK"

    Debug.Print _
        "TOC LENGTH = "; _
        Len(tocHtml)

    '--------------------------------------------------------
    ' 目次から候補ページを取得
    '--------------------------------------------------------

    Debug.Print
    Debug.Print "SEARCH TOC..."

    candidateUrl = _
        ShotokuFindBestPageFromToc( _
            tocHtml, _
            tsutatsuNumber)

    If Len(candidateUrl) > 0 Then

        Debug.Print
        Debug.Print "CANDIDATE:"
        Debug.Print candidateUrl

        body = _
            ShotokuGetEntryFromPage( _
                candidateUrl, _
                tsutatsuNumber)

        If Len(body) > 0 Then

            ShotokuGetTsutatsu = _
                ShotokuBuildResult( _
                    tsutatsuNumber, _
                    body, _
                    candidateUrl)

            Exit Function

        End If

    End If

    '--------------------------------------------------------
    ' TOCの候補で見つからなかった場合
    ' 全リンクを探索
    '--------------------------------------------------------

    Debug.Print
    Debug.Print "FULL TOC SEARCH..."

    Set gShotokuVisited = _
        CreateObject("Scripting.Dictionary")

    gShotokuFoundUrl = vbNullString

    body = _
        ShotokuSearchLinks( _
            tocHtml, _
            SHOTOKU_TOC_PATH, _
            tsutatsuNumber)

    If Len(body) > 0 Then

        ShotokuGetTsutatsu = _
            ShotokuBuildResult( _
                tsutatsuNumber, _
                body, _
                gShotokuFoundUrl)

        Exit Function

    End If

    '--------------------------------------------------------
    ' 見つからない
    '--------------------------------------------------------

    Err.Raise _
        vbObjectError + 4000, , _
        "所得税基本通達 " & _
        tsutatsuNumber & _
        " が見つかりませんでした。"

End Function


'============================================================
' TOCから候補ページを探す
'
' 例：
'
' 2-40
' ↓
' 法第2条《定義》関係
'
' 23-1
' ↓
' 法第23条《利子所得》関係
'
'============================================================

Private Function ShotokuFindBestPageFromToc( _
    ByVal html As String, _
    ByVal number As String) As String

    Dim links As Collection

    Dim v As Variant

    Dim href As String
    Dim text As String

    Dim firstPart As String

    Dim score As Long
    Dim bestScore As Long

    Dim bestUrl As String

    firstPart = _
        ShotokuGetNumberPart( _
            number, _
            1)

    Set links = _
        ShotokuExtractLinkInfo( _
            html, _
            SHOTOKU_TOC_PATH)

    Debug.Print _
        "TOC LINKS = "; _
        links.Count

    For Each v In links

        href = CStr(v(0))

        text = CStr(v(1))

        score = 0

        '----------------------------------------------------
        ' 法第○条
        '----------------------------------------------------

        If ShotokuTextHasArticle( _
            text, _
            firstPart) Then

            score = score + 100

        End If

        '----------------------------------------------------
        ' URLがshotoku配下
        '----------------------------------------------------

        If InStr( _
            1, _
            href, _
            "/law/tsutatsu/kihon/shotoku/", _
            vbTextCompare) > 0 Then

            score = score + 10

        End If

        If score > bestScore Then

            bestScore = score

            bestUrl = href

        End If

    Next v

    If bestScore > 0 Then

        ShotokuFindBestPageFromToc = _
            bestUrl

    End If

End Function


'============================================================
' 法第○条を判定
'
' 「法第2条」
' 「法第23条」
' 「法第23条から第35条まで」
' など
'============================================================

Private Function ShotokuTextHasArticle( _
    ByVal text As String, _
    ByVal articleNo As String) As Boolean

    Dim normalized As String

    normalized = _
        Replace(text, " ", "")

    normalized = _
        Replace(normalized, "　", "")

    If InStr( _
        1, _
        normalized, _
        "法第" & articleNo & "条", _
        vbTextCompare) > 0 Then

        ShotokuTextHasArticle = True

    End If

End Function


'============================================================
' 指定ページ取得
'============================================================

Private Function ShotokuGetEntryFromPage( _
    ByVal url As String, _
    ByVal number As String) As String

    Dim html As String

    Dim result As String

    On Error GoTo ErrorHandler

    Debug.Print "GET PAGE:"
    Debug.Print url

    html = _
        ShotokuHttpGetText(url)

    Debug.Print _
        "PAGE LENGTH = "; _
        Len(html)

    result = _
        ShotokuExtractTsutatsuEntry( _
            html, _
            number)

    If Len(result) > 0 Then

        Debug.Print "ENTRY FOUND"

        ShotokuGetEntryFromPage = _
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
' TOCの全リンクを探索
'============================================================

Private Function ShotokuSearchLinks( _
    ByVal html As String, _
    ByVal currentPath As String, _
    ByVal number As String) As String

    Dim links As Collection

    Dim v As Variant

    Dim href As String

    Dim childHtml As String

    Dim result As String

    Dim i As Long

    Set links = _
        ShotokuExtractLinks( _
            html, _
            currentPath)

    Debug.Print _
        "SEARCH LINKS = "; _
        links.Count

    For i = 1 To links.Count

        href = _
            CStr(links(i))

        If gShotokuVisited.Exists(href) Then

            GoTo ContinueLoop

        End If

        gShotokuVisited.Add _
            href, _
            True

        Debug.Print _
            "TRY [" & i & "/" & _
            links.Count & "] " & _
            href

        On Error Resume Next

        childHtml = _
            ShotokuHttpGetText(href)

        If Err.number <> 0 Then

            Err.Clear

            On Error GoTo 0

            GoTo ContinueLoop

        End If

        On Error GoTo 0

        '----------------------------------------------------
        ' ページ内検索
        '----------------------------------------------------

        result = _
            ShotokuExtractTsutatsuEntry( _
                childHtml, _
                number)

        If Len(result) > 0 Then

            Debug.Print
            Debug.Print "***** FOUND *****"
            Debug.Print href
            Debug.Print

            gShotokuFoundUrl = href

            ShotokuSearchLinks = result

            Exit Function

        End If

ContinueLoop:

        DoEvents

    Next i

End Function


'============================================================
' 通達本文抽出
'============================================================

Private Function ShotokuExtractTsutatsuEntry( _
    ByVal html As String, _
    ByVal targetNumber As String) As String

    Dim plain As String

    Dim p As Long
    Dim nextP As Long

    Dim target As String

    Dim nextNumber As String

    Debug.Print "TEXTIFY..."

    plain = _
        ShotokuHtmlToText(html)

    Debug.Print _
        "TEXT LENGTH = "; _
        Len(plain)

    target = _
        ShotokuNormalizeNumber( _
            targetNumber)

    Debug.Print _
        "TARGET=[" & target & "]"

    '--------------------------------------------------------
    ' 柔軟検索
    '--------------------------------------------------------

    p = _
        ShotokuFindNumberFlexible( _
            plain, _
            targetNumber)

    If p = 0 Then

        Debug.Print "NUMBER NOT FOUND"

        Exit Function

    End If

    Debug.Print _
        "NUMBER FOUND AT "; _
        p

    '--------------------------------------------------------
    ' 次の番号
    '--------------------------------------------------------

    nextNumber = _
        ShotokuGetNextNumber( _
            targetNumber)

    nextP = _
        ShotokuFindNumberFlexible( _
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

    ShotokuExtractTsutatsuEntry = _
        ShotokuCleanBody( _
            Mid$( _
                plain, _
                p, _
                nextP - p))

End Function


'============================================================
' 番号柔軟検索
'
' 2-1
' 2－1
' 2 ? 1
' 2-4の2
' 23-6の2
' などに対応
'============================================================

Private Function ShotokuFindNumberFlexible( _
    ByVal text As String, _
    ByVal number As String) As Long

    Dim re As Object

    Dim matches As Object

    Dim parts() As String

    Dim pattern As String

    number = _
        ShotokuNormalizeNumber( _
            number)

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

    If UBound(parts) >= 3 Then

        pattern = _
            pattern & _
            "\s*[-－??ー]\s*" & _
            parts(3)

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

        ShotokuFindNumberFlexible = _
            matches(0).FirstIndex + 1

    End If

End Function


'============================================================
' 次の番号
'
' 2-1 → 2-2
' 23-1 → 23-2
'============================================================

Private Function ShotokuGetNextNumber( _
    ByVal number As String) As String

    Dim parts() As String

    Dim lastIndex As Long

    Dim n As Long

    number = _
        ShotokuNormalizeNumber( _
            number)

    parts = _
        Split(number, "-")

    lastIndex = _
        UBound(parts)

    n = _
        Val(parts(lastIndex))

    parts(lastIndex) = _
        CStr(n + 1)

    ShotokuGetNextNumber = _
        Join(parts, "-")

End Function


'============================================================
' 番号の指定部分
'============================================================

Private Function ShotokuGetNumberPart( _
    ByVal number As String, _
    ByVal index As Long) As String

    Dim parts() As String

    number = _
        ShotokuNormalizeNumber( _
            number)

    parts = _
        Split(number, "-")

    If index <= _
        UBound(parts) + 1 Then

        ShotokuGetNumberPart = _
            parts(index - 1)

    End If

End Function


'============================================================
' HTML → TEXT
'============================================================

Private Function ShotokuHtmlToText( _
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
    ' SCRIPT
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

    '--------------------------------------------------------
    ' STYLE
    '--------------------------------------------------------

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
        ShotokuDecodeHtml(html)

    ShotokuHtmlToText = html

End Function


'============================================================
' HTML Entity
'============================================================

Private Function ShotokuDecodeHtml( _
    ByVal s As String) As String

    s = Replace(s, "&amp;", "&")
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quot;", """")
    s = Replace(s, "&#39;", "'")
    s = Replace(s, "&nbsp;", " ")
    s = Replace(s, "&emsp;", "　")
    s = Replace(s, "&ensp;", " ")

    ShotokuDecodeHtml = s

End Function


'============================================================
' HTTP
'============================================================

Private Function ShotokuHttpGetText( _
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
            vbObjectError + 4100, , _
            "HTTP ERROR " & _
            http.Status & _
            vbCrLf & _
            url

    End If

    bytes = _
        http.ResponseBody

    ShotokuHttpGetText = _
        ShotokuDecodeResponse(bytes)

End Function


'============================================================
' UTF-8 / Shift_JIS
'============================================================

Private Function ShotokuDecodeResponse( _
    ByVal bytes As Variant) As String

    Dim utf8Text As String

    Dim sjisText As String

    utf8Text = _
        ShotokuDecodeBytes( _
            bytes, _
            "utf-8")

    sjisText = _
        ShotokuDecodeBytes( _
            bytes, _
            "shift_jis")

    If ShotokuLooksLikeNta( _
        utf8Text) Then

        ShotokuDecodeResponse = _
            utf8Text

    ElseIf ShotokuLooksLikeNta( _
        sjisText) Then

        ShotokuDecodeResponse = _
            sjisText

    Else

        ShotokuDecodeResponse = _
            sjisText

    End If

End Function


'============================================================
' バイト列デコード
'============================================================

Private Function ShotokuDecodeBytes( _
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

    ShotokuDecodeBytes = _
        stm.ReadText

    stm.Close

    Set stm = Nothing

End Function


'============================================================
' NTAページ判定
'============================================================

Private Function ShotokuLooksLikeNta( _
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
        "所得税", _
        vbTextCompare) > 0 Then

        score = score + 3

    End If

    If InStr( _
        1, _
        text, _
        "所得税基本通達", _
        vbTextCompare) > 0 Then

        score = score + 3

    End If

    ShotokuLooksLikeNta = _
        (score >= 3)

End Function


'============================================================
' リンク抽出
'============================================================

Private Function ShotokuExtractLinks( _
    ByVal html As String, _
    ByVal basePath As String) As Collection

    Dim result As New Collection

    Dim info As Collection

    Dim v As Variant

    Set info = _
        ShotokuExtractLinkInfo( _
            html, _
            basePath)

    For Each v In info

        On Error Resume Next

        result.Add _
            CStr(v(0)), _
            CStr(v(0))

        On Error GoTo 0

    Next v

    Set ShotokuExtractLinks = _
        result

End Function


'============================================================
' href + text
'============================================================

Private Function ShotokuExtractLinkInfo( _
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
            ShotokuStripTags( _
                m.SubMatches(1))

        text = _
            ShotokuDecodeHtml( _
                text)

        text = _
            ShotokuCleanText( _
                text)

        If Len(href) = 0 Then
            GoTo ContinueLoop
        End If

        If Left$(href, 1) = "#" Then
            GoTo ContinueLoop
        End If

        href = _
            ShotokuResolveUrl( _
                href, _
                basePath)

        If Not ShotokuIsNtaUrl(href) Then
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

    Set ShotokuExtractLinkInfo = _
        result

End Function


'============================================================
' URL解決
'============================================================

Private Function ShotokuResolveUrl( _
    ByVal href As String, _
    ByVal basePath As String) As String

    Dim baseDir As String

    Dim p As Long

    href = _
        Trim$(href)

    If LCase$(Left$(href, 4)) = _
        "http" Then

        ShotokuResolveUrl = _
            href

        Exit Function

    End If

    If Left$(href, 1) = "/" Then

        ShotokuResolveUrl = _
            SHOTOKU_BASE_URL & _
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

    ShotokuResolveUrl = _
        SHOTOKU_BASE_URL & _
        ShotokuNormalizeRelative( _
            baseDir & href)

End Function


'============================================================
' 相対パス整理
'============================================================

Private Function ShotokuNormalizeRelative( _
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

    ShotokuNormalizeRelative = _
        result

End Function


'============================================================
' NTA URL判定
'============================================================

Private Function ShotokuIsNtaUrl( _
    ByVal url As String) As Boolean

    ShotokuIsNtaUrl = _
        InStr( _
            1, _
            url, _
            "https://www.nta.go.jp/", _
            vbTextCompare) = 1

End Function


'============================================================
' 番号正規化
'============================================================

Private Function ShotokuNormalizeNumber( _
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

    ShotokuNormalizeNumber = _
        Trim$(s)

End Function


'============================================================
' 本文整理
'============================================================

Private Function ShotokuCleanBody( _
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

    ShotokuCleanBody = _
        Trim$(s)

End Function


'============================================================
' テキスト整理
'============================================================

Private Function ShotokuCleanText( _
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

    ShotokuCleanText = _
        Trim$(s)

End Function


'============================================================
' HTMLタグ除去
'============================================================

Private Function ShotokuStripTags( _
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

    ShotokuStripTags = _
        re.Replace( _
            html, _
            "")

End Function


'============================================================
' 結果
'============================================================

Private Function ShotokuBuildResult( _
    ByVal number As String, _
    ByVal body As String, _
    ByVal url As String) As String

    ShotokuBuildResult = _
        "【所基通 " & _
        number & "】" & _
        vbCrLf & _
        body & _
        vbCrLf & _
        vbCrLf & _
        "出典:" & _
        vbCrLf & _
        url

End Function

