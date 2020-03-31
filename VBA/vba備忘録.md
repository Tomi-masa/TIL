## Outlook連携VBA　使えそうなやつ
***
## 取込機能
```vb
' CIWEB結果転記(inf)
Private Sub cmbCIWEBinf_Click()
On Error GoTo ErrorHandle

    Dim strDate As String, sinTimer As Single
    Dim strCWOutFolder As String, strFilePath As String, strFileName As String, lngPos As Long
    Dim varDataType As Variant
    Dim varData As Variant
    Dim wb As Workbook
    Dim fso As Object

    ' 初期設定
    Const cProName As String = "CIWEB結果転記(inf)"
    strDate = Format(Now, "yyyymmdd")
    sinTimer = Timer
    Set fso = CreateObject("Scripting.FileSystemObject")
```
```vb
'**************************************************************************
'  モジュール名     ：  subFileInput(ファイル取込み)
'  引数             ：  入力ファイルパス、入力ファイル引用符、入力ファイル区切り、入力ファイル属性、入力ファイル文字コード、取込み先シート
'  戻り値           ：  なし
'  用途             ：  入力ファイルを取込み先シートにデータを取込む
'                       引用符     1(xlTextQualifierDoubleQuote):ダブル,2(xlTextQualifierSingleQuote):シングル,-4142(xlTextQualifierNone):なし
'                       区切り     1:カンマ区切り,2:タブ区切り,3:セミコロン区切り,4:スペース区切り,5(値指定):その他区切り
'                       属性       1:自動判別,2:文字列,5:日付(YMD),9:読込まない
'                       文字コード 932:Shift-JIS,65001:UTF-8
'**************************************************************************
Public Sub subFileInput(ByRef pstrFPath As String, ByRef plngQua As Long, ByRef pstrSep As Long, ByRef pvarType As Variant, ByRef plngChaCD As Long, ByRef pWs As Worksheet)

    Dim varDeli(1 To 5) As Variant
    Dim strFile As String
    Dim namName As Name

    blnPErrFlg = True
    For i = LBound(varDeli) To UBound(varDeli)
        If i = 5 Then
            If IsNumeric(pstrSep) = True Then
                varDeli(5) = ""
            Else
                varDeli(5) = pstrSep
            End If
        ElseIf i = CLng(pstrSep) Then
            varDeli(i) = True
        Else
            varDeli(i) = False
        End If
    Next i

    pWs.Cells.Delete ' クリア
    strFile = "text;" & pstrFPath
    With pWs.QueryTables.Add(Connection:=strFile, Destination:=pWs.Range("A1"))
       .Name = "link1"
       .AdjustColumnWidth = False               ' 列幅の自動調整 True:する,False:しない
       .TextFileStartRow = 1                    ' 読込み開始行
       .TextFileParseType = xlDelimited         ' ファイル形式 xlDelimited:区切り,xlFixedWidth:固定幅
       .TextFileCommaDelimiter = varDeli(1)     ' カンマ区切り     True:する,False:しない
       .TextFileTabDelimiter = varDeli(2)       ' タブ区切り       True:する,False:しない
       .TextFileSemicolonDelimiter = varDeli(3) ' セミコロン区切り True:する,False:しない
       .TextFileSpaceDelimiter = varDeli(4)     ' スペース区切り   True:する,False:しない
       .TextFileOtherDelimiter = varDeli(5)     ' その他区切り     値指定
       .TextFileColumnDataTypes = pvarType      ' 項目型 1:自動判別,2:文字列,5:日付(YMD),9:読込まない
       .TextFilePlatform = plngChaCD            ' 文字コード 932:Shift-JIS,65001:UTF-8
       .TextFileTextQualifier = plngQua         ' 引用符 xlTextQualifierNone:なし(-4142),xlTextQualifierDoubleQuote:ダブル(1),xlTextQualifierSingleQuote:シングル(2)
       .RefreshStyle = xlInsertDeleteCells      ' データ挿入 xlInsertDeleteCells:挿入又は削除,xlOverwriteCells:上書き,xlInsertEntireRows:挿入
       .Refresh BackgroundQuery:=False          ' QueryTableオブジェクト更新 BackgroundQuery False:Worksheet反映,True:使用しない
       .Delete                                  ' QueryTableオブジェクト削除QueryTablesQueryTables
    End With
    For Each namName In ThisWorkbook.Names
        If InStr(namName.Name, "link1") > 0 Then
            namName.Delete
        End If
    Next namName
    blnPErrFlg = False
End Sub
```
```vb
'**************************************************************************
' モジュール名 ：subImportBodyCsv
' 機能         ：CMS概要をCMS_body_workにインポートする
'                概要→詳細の順である必要がある
'**************************************************************************
Private Sub subImportBodyCsv()
    
    On Error GoTo subImportBodyCsv_Err
    
    'インポートまえに土木、建築のどちらのモードで動かすか判別する必要があり、PJコード内の文字列で判別する
    '土建モードチェック
    Yosan_detail.Cells(1, 4).value = Kouji_body_main.Cells(2, 7)
    '土建モードチェックを削除 redmine#598
    strDoKenMode = Mid(Yosan_detail.Cells(1, 4).value, 5, 1)
    'If strDoKenMode <> "D" And strDoKenMode <> "K" Then
    '    MsgBox "プロジェクトコードが不適切です。", vbCritical
    '    isSuccess = False
    '    Exit Sub
    'End If
    
    Dim strOpenFileName As String
    strOpenFileName = Application.GetOpenFilename("CSVファイル(*.csv),*.csv", , "概要CSVを指定してください。")
    
    Dim lngCountCsvColun As Long
    lngCountCsvColun = funCountCsvColumn(strOpenFileName)
    
    '列数によるファイルフォーマットチェック
    If strOpenFileName <> "" And lngCountCsvColun <> cBodyColumns Then
        MsgBox "ファイル未指定、または、概要CSVではありません。", vbCritical
        isSuccess = False
        Exit Sub
    Else
        'ファイル種類チェック
        If strOpenFileName <> "False" Then
            CMS_body_work.Cells.Clear
        
            With CMS_body_work.QueryTables.Add(Connection:= _
                "TEXT;" & strOpenFileName, Destination:=Range("CMS_body_work!$A$1"))
                .Name = "data"
                .FieldNames = True
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .RefreshStyle = xlOverwriteCells    'セルに上書き
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .TextFilePromptOnRefresh = False
                .TextFilePlatform = 932             'CMSはShift_JIS(932)。UTF-8の場合は65001
                .TextFileStartRow = 1               'CMSはヘッダ無し。1 行目から読み込み
                .TextFileParseType = xlDelimited
                .TextFileTextQualifier = xlTextQualifierDoubleQuote
                .TextFileConsecutiveDelimiter = False
                .TextFileTabDelimiter = False
                .TextFileSemicolonDelimiter = False
                .TextFileCommaDelimiter = True
                .TextFileSpaceDelimiter = False
                .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2) 'MAX列数にしておく
                .TextFileTrailingMinusNumbers = True
                .Refresh BackgroundQuery:=False
                .Delete                          ' CSV との接続を解除
            End With
        End If
    End If
    
    isSuccess = True
    Exit Sub
    
subImportBodyCsv_Err:
    isSuccess = False
    MsgBox "概要CSVのインポートに失敗しました。" & Err.Description, vbCritical
    
End Sub
```
## VBA高速化
```vb
'**************************************************************************
'  モジュール名     ：  subSetFocus(アプリケーション設定)
'  引数             ：  シート名、制御フラグ
'  戻り値           ：  なし
'  用途             ：  イベント抑制、画面描画、計算方法、ステータスバーを制御する
'                       pFlg = True:抑制、False:解除
'**************************************************************************
Public Sub subSetFocus(ByRef pFlg As Boolean)

    With Application
        ' イベント抑制
        .EnableEvents = Not pFlg
        ' 画面描画
        .ScreenUpdating = Not pFlg
        ' 自動計算
        ActiveSheet.Calculate
        .Calculation = IIf(pFlg, xlCalculationManual, xlCalculationAutomatic)
        ' ステータスバー
        If pFlg = False Then .StatusBar = pFlg
        ' A1参照形式
        .ReferenceStyle = xlA1
        ' 確認メッセージ
        .DisplayAlerts = Not pFlg
    End With
End Sub
```
## パス取得
```vb
'**************************************************************************
'  モジュール名     ：  funFilePath(ファイルパス取得)
'  引数             ：  ダイアログのタイトル、初期表示フォルダパス、拡張子
'  戻り値           ：  ファイルパス
'  用途             ：  ダイアログで選択されたファイルのパスを取得する
'**************************************************************************
Public Function funFilePath(ByRef pTitle As String, ByRef pInitPath As String, ByRef pExten As String) As String

    Dim lngPos As Long, strPath As String, strFile As String
    Dim wb As Workbook

    Const cProName As String = "ファイルパス取得"

    ' ファイルパス取得
    blnPErrFlg = False
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "CSV(タブ区切り)", "*" & pExten
        .Title = pTitle                     ' タイトル
        .InitialFileName = pInitPath & "\"  ' 初期表示フォルダ設定
        .AllowMultiSelect = False           ' True:複数ファイル,False:単一ファイル
        If .Show = True Then
            funFilePath = .SelectedItems(1)
            lngPos = InStrRev(funFilePath, "\")
            strPath = Left(funFilePath, lngPos - 1)
            strFile = Mid(funFilePath, lngPos + 1)
            ' ファイルOpenチェック
            For Each wb In Workbooks
                If wb.Name = strFile Then
                    strPMSG = Replace(W0006, "@FileName@", strFile)
                    lngPMSG = funClientMsg(cProName, strPMSG, Empty, Empty, 0, True)
                    blnPErrFlg = True
                    Exit For
                End If
            Next wb
        Else
            lngPMSG = funClientMsg(cProName, I0001, Empty, Empty, 0, True)
            blnPErrFlg = True
        End If
    End With
End Function
```
## マスタから値を取得
```vb
'**************************************************************************
'  モジュール名     ：  funYosan_body_GetVerName(予算概要シート名称を取得)
'  引数             ：  対象文字列
'  戻り値           ：  名称
'  用途             ：  予算概要シートに入力された前回/今回予算Ver.から名称を取得
'**************************************************************************
Public Function funYosan_body_GetVerName(strGetVer As String, Yosan_body_Sheet As Object) As String

    Dim strGetVerNmae As String
    Dim strCells As String

    ' 入力されたVerを元に名称を取得(redmine#393)
    If (Left(strGetVer, 1) = "M") Then
        'M=見積原価予算の名称を取得
        strCells = "DO8:DP" & CStr(Common_mst.Range("DO8").End(xlDown).Row)
    Else
        'J=実行予算の名称を取得
        strCells = "DR8:DS" & CStr(Common_mst.Range("DR8").End(xlDown).Row)
    End If
    
    strGetVerNmae = WorksheetFunction.VLookup(strGetVer, Common_mst.Range(strCells), 2, False)

    funYosan_body_GetVerName = strGetVerNmae

End Function
```
