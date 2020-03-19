
Public Sub ExportOthersCalendar()
    Const CSV_FILE_NAME = "C:\Users\20721\Documents\thismonth.csv" ' エクスポートするファイル名を指定
    Dim arrUsers() As Variant
    ' エクスポートするユーザーの名前かメールアドレスを指定
    arrUsers = Array("keita.moriyama@n-kokudo.co.jp", "tsuyoshi.furusawa@n-kokudo.co.jp", "masaki.tomita@n-kokudo.co.jp")
    Dim dtExport As Date
    Dim strStart As String
    Dim strEnd As String
    Dim objFSO As FileSystemObject
    Dim stmCSVFile As TextStream
    Dim strUserName As String
    Dim objRecip As Recipient
    Dim colAppts As Items
    Dim objAppt 'As AppointmentItem
    Dim strLine As String
    Dim i As Integer
    '
    dtExport = Now ' 来月の予定をエクスポートする場合は Now の代わりに DateAdd("m",1,Now) を使用します。
    ' 月単位ではなく任意の単位にする場合は以下の記述を変更します。
    strStart = Year(Now) & "/" & Month(Now) & "/1 00:00"
    strEnd = DateAdd("m", 1, CDate(strStart)) & " 00:00"
    Set objFSO = CreateObject("Scripting.FileSystemObject") 'インスタンス化
    Set stmCSVFile = objFSO.CreateTextFile(CSV_FILE_NAME, True) 'すでにテキストファイルがある場合は上書き
    ' CSV ファイルのヘッダ
    stmCSVFile.WriteLine """ユーザー"",""件名"",""場所"",""開始日時"",""終了日時"""
    For i = LBound(arrUsers) To UBound(arrUsers)
        strUserName = arrUsers(i)
        Set objRecip = Application.Session.CreateRecipient(strUserName)
        objRecip.Resolve
        'ユーザーが確認されなかった場合
        If Not objRecip.Resolved Then
            MsgBox "ユーザーが特定できませんでした。", vbCritical, "共有されている予定表のエクスポート"
            Exit Sub
        End If
        On Error Resume Next
        Set colAppts = Application.Session.GetSharedDefaultFolder(objRecip, olFolderCalendar).Items
        colAppts.Sort "[Start]"
        colAppts.IncludeRecurrences = True  '定期的な予定を含む
        Set objAppt = colAppts.Find("[Start] < """ & strEnd & """ AND [End] >= """ & strStart & """")
        While Not objAppt Is Nothing
            strLine = """" & objRecip.Name & _
                """,""" & objAppt.Subject & _
                """,""" & objAppt.Location & _
                """,""" & objAppt.Start & _
                """,""" & objAppt.End & _
                """"
    '
           stmCSVFile.WriteLine strLine
            Set objAppt = colAppts.FindNext
        Wend
    Next
    stmCSVFile.Close
    MsgBox "作業完了！", vbOKOnly, "～Fin.～"
End Sub







