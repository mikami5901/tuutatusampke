Attribute VB_Name = "Module1"
Option Explicit

'============================================================
' NTA Tax Tsutatsu Retriever - VBA
' 実験作品
' 目的:
'   国税庁の法人税基本通達から指定番号を取得する
'
' 実行:
'   Test_GetTsutatsu
'
' 例:
'   GetTsutatsu("法基通", "1-1-1")
'
' 参照設定:
'   不要
'
' 使用:
'   WinHttp.WinHttpRequest.5.1
'   ADODB.Stream
'   VBScript.RegExp
'   Scripting.Dictionary
'
'============================================================

Private Const NTA_BASE As String = _
    "https://www.nta.go.jp"

Private Const HOJIN_TOC As String = _
    "/law/tsutatsu/kihon/hojin/01.htm"

Private Const MAX_DEPTH As Long = 4

Private gVisited As Object
Private gFoundUrl As String


'============================================================
' TEST
'============================================================

Public Sub Test_GetTsutatsu()

    Dim result As String

    Debug.Print String(70, "=")
    Debug.Print "NTA TSUTATSU TEST"
    Debug.Print String(70, "=")

    result = GetTsutatsu( _
                "法基通", _
                "1-1-1")

    Debug.Print
    Debug.Print "================ RESULT ================"
    Debug.Print result
    Debug.Print "========================================="
    Debug.Print

End Sub


'============================================================
' PUBLIC
'============================================================

Public Function GetTsutatsu( _
    ByVal tsutatsuName As String, _
    ByVal tsutatsuNumber As String) As String

    Dim tocHtml As String
    Dim body As String
    Dim url As String

    gFoundUrl = vbNullString
    Set gVisited = CreateObject("Scripting.Dictionary")

    Debug.Print "GET TOC..."

    tocHtml = HttpGetText( _
                    NTA_BASE & HOJIN_TOC)

    Debug.Print "TOC OK"
    Debug.Print "TOC LENGTH = "; Len(tocHtml)

    '--------------------------------------------------------
    ' まずTOCから目的ページを探す
    '--------------------------------------------------------

    Debug.Print
    Debug.Print "SEARCH TOC..."

    url = FindBestPageFromToc( _
                tocHtml, _
                HOJIN_TOC, _
                tsutatsuNumber)

    If Len(url) > 0 Then

        Debug.Print "CANDIDATE:"
        Debug.Print url

        body = GetEntryFromPage( _
                    url, _
                    tsutatsuNumber)

        If Len(body) > 0 Then

            GetTsutatsu = BuildResult( _
                                tsutatsuName, _
                                tsutatsuNumber, _
                                body, _
                                url)

            Exit Function

        End If

    End If

    '--------------------------------------------------------
    ' TOCの階層を探索
    '--------------------------------------------------------

    Debug.Print
    Debug.Print "SEARCH CHILD PAGES..."

    body = SearchChildPages( _
                tocHtml, _
                HOJIN_TOC, _
                tsutatsuNumber, _
                0)

    If Len(body) = 0 Then

        Err.Raise _
            vbObjectError + 1000, , _
            "Tsutatsu not found: " & _
            tsutatsuNumber

    End If

    GetTsutatsu = BuildResult( _
                        tsutatsuName, _
                        tsutatsuNumber, _
                        body, _
                        gFoundUrl)

End Function


'============================================================
' ページから通達を取得
'============================================================

Private Function GetEntryFromPage( _
    ByVal url As String, _
    ByVal number As String) As String

    Dim html As String
    Dim result As String

    On Error GoTo ErrorHandler

    Debug.Print "GET PAGE:"
    Debug.Print url

    html = HttpGetText(url)

    Debug.Print "PAGE LENGTH = "; Len(html)

    result = ExtractTsutatsuEntry( _
                html, _
                number)

    If Len(result) > 0 Then

        Debug.Print "ENTRY FOUND"

        GetEntryFromPage = result

    Else

        Debug.Print "ENTRY NOT FOUND"

    End If

    Exit Function

ErrorHandler:

    Debug.Print "PAGE ERROR:"
    Debug.Print Err.Description

End Function


'============================================================
' 階層探索
'============================================================

Private Function SearchChildPages( _
    ByVal html As String, _
    ByVal currentPath As String, _
    ByVal number As String, _
    ByVal depth As Long) As String

    Dim links As Collection
    Dim v As Variant

    Dim href As String
    Dim childHtml As String
    Dim result As String

    If depth > MAX_DEPTH Then Exit Function

    Set links = ExtractLinks( _
                    html, _
                    currentPath)

    Debug.Print _
        String(depth * 2, " ") & _
        "DEPTH=" & depth & _
        " LINKS=" & links.Count

    '--------------------------------------------------------
    ' 目的番号に近いリンクを先にする
    '--------------------------------------------------------

    Set links = SortLinksForTarget( _
                    links, _
                    number)

    For Each v In links

        href = CStr(v)

        If gVisited.Exists(href) Then
            GoTo ContinueLoop
        End If

        gVisited.Add href, True

        Debug.Print _
            String(depth * 2, " ") & _
            "TRY: " & href

        On Error Resume Next

        childHtml = HttpGetText(href)

        If Err.number <> 0 Then

            Debug.Print _
                String(depth * 2, " ") & _
                "HTTP ERROR"

            Err.Clear
            On Error GoTo 0

            GoTo ContinueLoop

        End If

        On Error GoTo 0

        '----------------------------------------------------
        ' このページ自身
        '----------------------------------------------------

        result = ExtractTsutatsuEntry( _
                    childHtml, _
                    number)

        If Len(result) > 0 Then

            Debug.Print
            Debug.Print "***** FOUND *****"
            Debug.Print href
            Debug.Print

            gFoundUrl = href

            SearchChildPages = result

            Exit Function

        End If

        '----------------------------------------------------
        ' さらに下
        '----------------------------------------------------

        result = SearchChildPages( _
                    childHtml, _
                    href, _
                    number, _
                    depth + 1)

        If Len(result) > 0 Then

            SearchChildPages = result

            Exit Function

        End If

ContinueLoop:

    Next v

End Function


'============================================================
' TOCから候補を探す
'============================================================

Private Function FindBestPageFromToc( _
    ByVal html As String, _
    ByVal basePath As String, _
    ByVal targetNumber As String) As String

    Dim links As Collection
    Dim info As Collection
    Dim v As Variant

    Dim targetPrefix As String
    Dim targetFirst As String
    Dim targetSecond As String

    Dim score As Long
    Dim bestScore As Long

    Dim bestUrl As String

    targetPrefix = GetNumberPrefix(targetNumber)

    targetFirst = GetNumberPart(targetNumber, 1)
    targetSecond = GetNumberPart(targetNumber, 2)

    Set info = ExtractLinkInfo( _
                    html, _
                    basePath)

    bestScore = 0

    For Each v In info

        score = ScoreLink( _
                    CStr(v(1)), _
                    CStr(v(0)), _
                    targetFirst, _
                    targetSecond, _
                    targetPrefix)

        If score > bestScore Then

            bestScore = score
            bestUrl = CStr(v(0))

        End If

    Next v

    If bestScore > 0 Then

        FindBestPageFromToc = bestUrl

    End If

End Function


'============================================================
' リンクスコア
'============================================================

Private Function ScoreLink( _
    ByVal text As String, _
    ByVal href As String, _
    ByVal firstPart As String, _
    ByVal secondPart As String, _
    ByVal prefix As String) As Long

    Dim score As Long

    text = NormalizeSearchText(text)

    ' 通達番号そのもの
    If InStr( _
        1, _
        text, _
        NormalizeSearchText(prefix), _
        vbTextCompare) > 0 Then

        score = score + 20

    End If

    ' 第X章
    If InStr( _
        1, _
        text, _
        "第" & firstPart & "章", _
        vbTextCompare) > 0 Then

        score = score + 30

    End If

    ' 第X節
    If InStr( _
        1, _
        text, _
        "第" & secondPart & "節", _
        vbTextCompare) > 0 Then

        score = score + 40

    End If

    ' URLにも番号が入っている場合
    If InStr( _
        1, _
        href, _
        "/" & firstPart & "/", _
        vbTextCompare) > 0 Then

        score = score + 10

    End If

    ScoreLink = score

End Function


'============================================================
' リンク抽出
'============================================================

Private Function ExtractLinks( _
    ByVal html As String, _
    ByVal basePath As String) As Collection

    Dim result As New Collection
    Dim info As Collection
    Dim v As Variant

    Set info = ExtractLinkInfo( _
                    html, _
                    basePath)

    For Each v In info

        On Error Resume Next

        result.Add _
            CStr(v(0)), _
            CStr(v(0))

        On Error GoTo 0

    Next v

    Set ExtractLinks = result

End Function


'============================================================
' href / text
'============================================================

Private Function ExtractLinkInfo( _
    ByVal html As String, _
    ByVal basePath As String) As Collection

    Dim result As New Collection

    Dim re As Object
    Dim matches As Object
    Dim m As Object

    Dim href As String
    Dim text As String

    Dim item(0 To 1) As Variant

    Set re = CreateObject( _
                "VBScript.RegExp")

    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True

    re.pattern = _
        "<a[^>]*href\s*=\s*[""]([^""]+)[""][^>]*>" & _
        "([\s\S]*?)</a>"

    Set matches = re.Execute(html)

    For Each m In matches

        href = Trim$( _
                    m.SubMatches(0))

        text = StripTags( _
                    m.SubMatches(1))

        text = DecodeHtml( _
                    text)

        text = CleanText( _
                    text)

        If Len(href) = 0 Then
            GoTo ContinueLoop
        End If

        If Left$(href, 1) = "#" Then
            GoTo ContinueLoop
        End If

        href = ResolveUrl( _
                    href, _
                    basePath)

        If Not IsNtaUrl(href) Then
            GoTo ContinueLoop
        End If

        item(0) = href
        item(1) = text

        On Error Resume Next

        result.Add item, href

        On Error GoTo 0

ContinueLoop:

    Next m

    Set ExtractLinkInfo = result

End Function


'============================================================
' ★ 通達本文抽出
'
' HTML構造に依存しない。
'
' HTML
'   ↓
' プレーンテキスト
'   ↓
' 1-1-1検索
'   ↓
' 次の1-1-2まで
'
'============================================================

Private Function ExtractTsutatsuEntry( _
    ByVal html As String, _
    ByVal targetNumber As String) As String

    Dim plain As String

    Dim p As Long
    Dim nextP As Long

    Dim target As String
    Dim nextNumber As String

    Debug.Print "TEXTIFY..."

    plain = HtmlToText(html)

    Debug.Print "TEXT LENGTH = "; Len(plain)

    target = NormalizeSearchText( _
                targetNumber)

    Debug.Print _
        "TARGET=[" & target & "]"

    '--------------------------------------------------------
    ' まず通常検索
    '--------------------------------------------------------

    p = InStr( _
            1, _
            plain, _
            target, _
            vbTextCompare)

    '--------------------------------------------------------
    ' 見つからない場合
    ' 数字とダッシュの間に空白等がある可能性を考慮
    '--------------------------------------------------------

    If p = 0 Then

        p = FindTsutatsuNumberFlexible( _
                plain, _
                targetNumber)

    End If

    If p = 0 Then

        Debug.Print "NUMBER NOT FOUND"

        ' デバッグ用
        DebugNumberPresence plain

        Exit Function

    End If

    Debug.Print _
        "NUMBER FOUND AT "; p

    '--------------------------------------------------------
    ' 次の通達番号
    '--------------------------------------------------------

    nextNumber = GetNextTsutatsuNumber( _
                    targetNumber)

    nextP = FindTsutatsuNumberFlexible( _
                Mid$(plain, p + Len(target) + 1), _
                nextNumber)

    If nextP > 0 Then

        nextP = _
            p + _
            Len(target) + _
            nextP

    Else

        nextP = Len(plain) + 1

    End If

    '--------------------------------------------------------
    ' 本文
    '--------------------------------------------------------

    ExtractTsutatsuEntry = _
        CleanTsutatsuBody( _
            Mid$( _
                plain, _
                p, _
                nextP - p))

End Function


'============================================================
' 柔軟な通達番号検索
'============================================================

Private Function FindTsutatsuNumberFlexible( _
    ByVal text As String, _
    ByVal number As String) As Long

    Dim re As Object
    Dim matches As Object

    Dim parts() As String
    Dim pattern As String

    number = NormalizeSearchText( _
                number)

    parts = Split(number, "-")

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

    Set re = CreateObject( _
                "VBScript.RegExp")

    re.Global = False
    re.IgnoreCase = True
    re.MultiLine = True

    re.pattern = pattern

    Set matches = _
        re.Execute(text)

    If matches.Count > 0 Then

        FindTsutatsuNumberFlexible = _
            matches(0).FirstIndex + 1

    End If

End Function


'============================================================
' デバッグ
'============================================================

Private Sub DebugNumberPresence( _
    ByVal text As String)

    Dim p As Long

    p = InStr( _
            1, _
            text, _
            "1", _
            vbBinaryCompare)

    If p > 0 Then

        Debug.Print _
            "FIRST '1' AT "; p

        Debug.Print _
            Left$( _
                Mid$(text, p, 500), _
                500)

    End If

End Sub


'============================================================
' HTML → TEXT
'============================================================

Private Function HtmlToText( _
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
    ' script/style
    '--------------------------------------------------------

    Set re = CreateObject( _
                "VBScript.RegExp")

    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True

    re.pattern = _
        "<script[^>]*>[\s\S]*?</script>"

    html = re.Replace( _
                html, _
                "")

    re.pattern = _
        "<style[^>]*>[\s\S]*?</style>"

    html = re.Replace( _
                html, _
                "")

    '--------------------------------------------------------
    ' HTMLタグ
    '--------------------------------------------------------

    re.pattern = "<[^>]*>"

    html = re.Replace( _
                html, _
                "")

    '--------------------------------------------------------
    ' HTML entity
    '--------------------------------------------------------

    html = DecodeHtml(html)

    HtmlToText = html

End Function


'============================================================
' HTML Entity
'============================================================

Private Function DecodeHtml( _
    ByVal s As String) As String

    s = Replace(s, "&amp;", "&")
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quot;", """")
    s = Replace(s, "&#39;", "'")
    s = Replace(s, "&nbsp;", " ")
    s = Replace(s, "&emsp;", "　")
    s = Replace(s, "&ensp;", " ")

    DecodeHtml = s

End Function


'============================================================
' ★ HTTP
'
' ResponseBodyを取得し、
' UTF-8 / Shift_JISを自動判定
'============================================================

Private Function HttpGetText( _
    ByVal url As String) As String

    Dim http As Object
    Dim bytes As Variant

    Set http = CreateObject( _
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
            vbObjectError + 1100, , _
            "HTTP ERROR " & _
            http.Status & _
            vbCrLf & url

    End If

    bytes = http.ResponseBody

    HttpGetText = _
        DecodeResponse(bytes)

End Function


'============================================================
' ResponseBodyデコード
'============================================================

Private Function DecodeResponse( _
    ByVal bytes As Variant) As String

    Dim utf8Text As String
    Dim sjisText As String

    utf8Text = DecodeBytes( _
                    bytes, _
                    "utf-8")

    sjisText = DecodeBytes( _
                    bytes, _
                    "shift_jis")

    '--------------------------------------------------------
    ' NTAページとして自然な方を判定
    '--------------------------------------------------------

    If IsGoodNtaHtml(utf8Text) Then

        DecodeResponse = utf8Text

    ElseIf IsGoodNtaHtml(sjisText) Then

        DecodeResponse = sjisText

    Else

        ' NTAの古いページはShift_JIS系が多いので
        ' 最後はShift_JISを採用
        DecodeResponse = sjisText

    End If

End Function


'============================================================
' バイト列 → 文字列
'============================================================

Private Function DecodeBytes( _
    ByVal bytes As Variant, _
    ByVal charset As String) As String

    Dim stm As Object

    Set stm = CreateObject( _
                "ADODB.Stream")

    stm.Type = 1
    stm.Open

    stm.Write bytes

    stm.Position = 0

    stm.Type = 2

    stm.charset = charset

    DecodeBytes = _
        stm.ReadText

    stm.Close

    Set stm = Nothing

End Function


'============================================================
' NTA HTML判定
'============================================================

Private Function IsGoodNtaHtml( _
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
        "法人税", _
        vbTextCompare) > 0 Then

        score = score + 3

    End If

    If InStr( _
        1, _
        text, _
        "第1節", _
        vbTextCompare) > 0 Then

        score = score + 3

    End If

    If InStr( _
        1, _
        text, _
        "<html", _
        vbTextCompare) > 0 Then

        score = score + 1

    End If

    IsGoodNtaHtml = _
        (score >= 3)

End Function


'============================================================
' URL
'============================================================

Private Function ResolveUrl( _
    ByVal href As String, _
    ByVal basePath As String) As String

    Dim baseDir As String
    Dim p As Long

    href = Trim$(href)

    If LCase$(Left$(href, 4)) = "http" Then

        ResolveUrl = href

        Exit Function

    End If

    If Left$(href, 1) = "/" Then

        ResolveUrl = _
            NTA_BASE & href

        Exit Function

    End If

    p = InStrRev( _
            basePath, _
            "/")

    If p > 0 Then

        baseDir = _
            Left$(basePath, p)

    End If

    ResolveUrl = _
        NTA_BASE & _
        NormalizeRelative( _
            baseDir & href)

End Function


'============================================================
' ../
'============================================================

Private Function NormalizeRelative( _
    ByVal path As String) As String

    Dim parts() As String
    Dim stack() As String

    Dim i As Long
    Dim n As Long

    Dim part As String
    Dim result As String

    parts = Split(path, "/")

    ReDim stack( _
        0 To UBound(parts))

    For i = LBound(parts) _
        To UBound(parts)

        part = parts(i)

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
            result & stack(i)

        If i < n - 1 Then
            result = result & "/"
        End If

    Next i

    NormalizeRelative = result

End Function


'============================================================
' NTA URL判定
'============================================================

Private Function IsNtaUrl( _
    ByVal url As String) As Boolean

    IsNtaUrl = _
        InStr( _
            1, _
            url, _
            "https://www.nta.go.jp/", _
            vbTextCompare) = 1

End Function


'============================================================
' 検索用文字列正規化
'============================================================

Private Function NormalizeSearchText( _
    ByVal s As String) As String

    On Error Resume Next

    s = StrConv( _
            s, _
            vbNarrow)

    On Error GoTo 0

    s = Replace(s, "－", "-")
    s = Replace(s, "?", "-")
    s = Replace(s, "?", "-")
    s = Replace(s, "ー", "-")

    s = Replace(s, "　", " ")

    NormalizeSearchText = s

End Function


'============================================================
' 通達番号の先頭
'
' 1-1-1 → 1
' 33-6 → 33
'============================================================

Private Function GetNumberPrefix( _
    ByVal number As String) As String

    Dim p As Long

    number = _
        NormalizeSearchText(number)

    p = InStr( _
            number, _
            "-")

    If p > 0 Then

        GetNumberPrefix = _
            Left$(number, p - 1)

    Else

        GetNumberPrefix = number

    End If

End Function


'============================================================
' 番号のn番目
'
' 1-1-1
'
' 1 → 1
' 2 → 1
' 3 → 1
'============================================================

Private Function GetNumberPart( _
    ByVal number As String, _
    ByVal index As Long) As String

    Dim arr() As String

    number = _
        NormalizeSearchText(number)

    arr = Split(number, "-")

    If index <= UBound(arr) + 1 Then

        GetNumberPart = _
            arr(index - 1)

    End If

End Function


'============================================================
' 次の通達番号
'
' 1-1-1 → 1-1-2
' 33-6 → 33-7
'============================================================

Private Function GetNextTsutatsuNumber( _
    ByVal number As String) As String

    Dim arr() As String
    Dim lastIndex As Long
    Dim n As Long

    number = _
        NormalizeSearchText(number)

    arr = Split(number, "-")

    lastIndex = UBound(arr)

    n = Val( _
            arr(lastIndex))

    arr(lastIndex) = _
        CStr(n + 1)

    GetNextTsutatsuNumber = _
        Join(arr, "-")

End Function


'============================================================
' 本文整理
'============================================================

Private Function CleanTsutatsuBody( _
    ByVal s As String) As String

    s = Replace( _
        s, _
        ChrW(160), _
        " ")

    Do While InStr( _
        s, _
        vbCrLf & vbCrLf & vbCrLf) > 0

        s = Replace( _
            s, _
            vbCrLf & vbCrLf & vbCrLf, _
            vbCrLf & vbCrLf)

    Loop

    CleanTsutatsuBody = _
        Trim$(s)

End Function


'============================================================
' HTMLタグ除去
'============================================================

Private Function StripTags( _
    ByVal html As String) As String

    Dim re As Object

    Set re = CreateObject( _
                "VBScript.RegExp")

    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True

    re.pattern = _
        "<script[^>]*>[\s\S]*?</script>"

    html = re.Replace( _
                html, _
                "")

    re.pattern = _
        "<style[^>]*>[\s\S]*?</style>"

    html = re.Replace( _
                html, _
                "")

    re.pattern = _
        "<[^>]*>"

    StripTags = _
        re.Replace( _
            html, _
            "")

End Function


'============================================================
' Clean
'============================================================

Private Function CleanText( _
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

    CleanText = _
        Trim$(s)

End Function


'============================================================
' リンク並び替え
'
' 今回はCollectionをそのまま返す。
' 後でスコア順に改良可能。
'============================================================

Private Function SortLinksForTarget( _
    ByVal links As Collection, _
    ByVal number As String) As Collection

    ' 現段階では元順序を維持
    Set SortLinksForTarget = links

End Function


'============================================================
' 結果
'============================================================

Private Function BuildResult( _
    ByVal name As String, _
    ByVal number As String, _
    ByVal body As String, _
    ByVal url As String) As String

    BuildResult = _
        "【" & name & " " & number & "】" & _
        vbCrLf & _
        body & _
        vbCrLf & vbCrLf & _
        "出典:" & _
        vbCrLf & _
        url

End Function
