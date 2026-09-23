Attribute VB_Name = "Module2"
'============================================================
' Excelから通達を取得
' 一旦、法人税基本通達専用
' B1 = 通達名
' B2 = 通達番号
'
' 例:
' B1 : 法基通
' B2 : 1-1-1
'
' 結果:
' B4以降に表示
'============================================================

Public Sub 通達取得()

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
    ' 入力チェック
    '--------------------------------------------------------

    If Len(tsutatsuName) = 0 Then

        MsgBox _
            "B1に通達名を入力してください。" & _
            vbCrLf & _
            "例：法基通", _
            vbExclamation

        Exit Sub

    End If

    If Len(tsutatsuNumber) = 0 Then

        MsgBox _
            "B2に通達番号を入力してください。" & _
            vbCrLf & _
            "例：1-1-1", _
            vbExclamation

        Exit Sub

    End If

    '--------------------------------------------------------
    ' 既存のGetTsutatsuを実行
    '--------------------------------------------------------

    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    ws.Range("B4:B10000").ClearContents

    Application.StatusBar = _
        "通達を取得しています..."

    result = GetTsutatsu( _
                tsutatsuName, _
                tsutatsuNumber)

    '--------------------------------------------------------
    ' 結果を表示
    '--------------------------------------------------------

    ws.Range("B4").Value = result

    ws.Range("B4").WrapText = True

    ws.Columns("B").ColumnWidth = 100

    ws.Rows("4:10000").RowHeight = 18

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox _
        "通達を取得しました。", _
        vbInformation

    Exit Sub

ErrorHandler:

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox _
        "通達の取得に失敗しました。" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbExclamation

End Sub
