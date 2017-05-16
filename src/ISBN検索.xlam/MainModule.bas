Attribute VB_Name = "MainModule"
Option Explicit

' ================================================================
' リボンはxlamファイル展開した中の_rels/.relsファイル中に指定された
' Relation項目から参照されるUI定義xmlで定義する.
' 本ファイルはcustomUI/customUI.xmlに定義している.
' ================================================================

' シート設定フォームを開く
' カスタムリボンから呼び出される
Public Sub ShowConfigFormRibbon(ribbon As IRibbonControl)
    ShowConfigForm
End Sub

' ISBN検索実行
' カスタムリボンから呼び出される
Public Sub UpdateSheetFromRibbon(ribbon As IRibbonControl)
    UpdateSheet
End Sub


' About表示
' カスタムリボンから呼び出される
Public Sub ShowAboutFromRibbon(ribbon As IRibbonControl)
    Dim frm As New AboutForm
    frm.Show
End Sub


Sub UpdateSheet()
' ------------------------------------------------------------
' 事前に定義されているシートの設定に従ってISBN検索を実施する.
'
' Remarks:
'   現在のアクティブシートが対象となる.
'   アクティブシートがない場合は何もしない.
' ------------------------------------------------------------
On Error GoTo Err
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    Dim Sheet As Worksheet
    Set Sheet = ActiveWorkbook.ActiveSheet

    If Sheet Is Nothing Then Exit Sub
    
    UpdateSheetWithConf Sheet
    
    MsgBox "完了しました"
    Exit Sub

Err:
    MsgBox "Error: " & Err.Description & vbCrLf & Err.Source
    Exit Sub
End Sub

Public Sub UpdateSheetWithConf(Sheet As Worksheet)
' ------------------------------------------------------------
' 事前に定義されているシートの設定に従ってISBN検索を実施する.
'
' Parameter:
'    [Sheet] IN 対象となるワークシート
'-------------------------------------------------------------

    ' 設定値の読み込み
    Dim conf As ConfigModule.SheetConfig
    Dim connectConf As ConfigModule.ConnectConfig
    conf = ConfigModule.GetSheetConfig()
    connectConf = ConfigModule.GetConnectConfig()
    
    ' ISBN検索オブジェクトの構築
    Dim SearcherImpl As New NDLSearchISBN
    ' ISBNの妥当性チェック
    Dim verifier As New ISBNVerifier
    
    ' XMLリーダーの構築
    Dim fac As New XMLReaderFactory
    Set SearcherImpl.XmlReader = fac.Create(connectConf)
    
    Dim searcher As ISearchISBN
    Set searcher = SearcherImpl
    
    ' シートで使用されているセル範囲
    Dim MaxRow As Integer, MaxCol As Integer
    With Sheet.UsedRange
        MaxRow = .Rows(.Rows.Count).row
        MaxCol = .Columns(.Columns.Count).Column
    End With
    
    Dim ISBN As String
    Dim TITLE As String
    
    Dim row As Integer
    If conf.StartCurPos Then row = ActiveCell.row
    If row < conf.StartRow Then row = conf.StartRow
    
    ' 開始行からシートで使用されている最大行までループする
    Do While row <= MaxRow
        DoEvents
        
        ISBN = Trim(Sheet.Cells(row, conf.ISBN).value)
        TITLE = Sheet.Cells(row, conf.TITLE).value
        
        If ISBN = "" Then
            ' ISBNが空欄なら何もしない
        
        ElseIf TITLE = "" Then
            ' まだタイトルが取得されていない行のみ処理する.
            
            ' ISBNからハイフンと空白を除去する
            ISBN = Replace(ISBN, "-", "")
            ISBN = Replace(ISBN, " ", "")
            
            ' ISBNの妥当性チェック
            If verifier.Verify(ISBN) Then
                ' 現在処理中のISBNをステータスバーに表示する.
                Application.StatusBar = "ISBN:" & ISBN
                
                ' 検索の実施
                Dim info As BookInfo
                If searcher.FindByISBN(ISBN, info) Then
                    TITLE = info.TITLE
                    
                    If info.VOLUME <> "" And conf.TITLE_WITH_VOLUME Then
                        ' Volumeの指定がある場合はタイトルに括弧付きでつなげる.
                        TITLE = TITLE & "(" & info.VOLUME & ")"
                    End If
                    
                    UpdateCell Sheet, row, conf.TITLE, TITLE
                    UpdateCell Sheet, row, conf.AUTHOR, info.AUTHOR
                    UpdateCell Sheet, row, conf.PUBLISHER, info.PUBLISHER
                    UpdateCell Sheet, row, conf.ISSUED, info.ISSUED
                    UpdateCell Sheet, row, conf.YOMI, info.YOMI
                    UpdateCell Sheet, row, conf.VOLUME, info.VOLUME
                Else
                    ' 書籍情報が取得できないISBNはセルの背景色を変えておく
                    Sheet.Cells(row, conf.ISBN).Interior.ColorIndex = 37
                End If
            Else
                ' ISBNが妥当でない場合はセルの背景色を変えておく
                Sheet.Cells(row, conf.ISBN).Interior.ColorIndex = 38
            End If
        End If
        row = row + 1
    Loop
    ' 完了したらステータスバーは戻す
    Application.StatusBar = False
End Sub

Private Sub UpdateCell(Sheet As Worksheet, row As Integer, col As Integer, value As String)
'------------------------------------------------------------------------
' セル列位置が有効であれば、そのセルに書き込む.
'
' Parameter:
'   [Sheet] IN 書き込み先シート
'   [Row]   IN 行番号
'   [Col]   IN 列番号
'   [Value] IN 書き込む文字列
'
' Remarks:
'   0以下の列番号が指定されている場合は書き込みせず省略する.
'------------------------------------------------------------------------
    If col > 0 Then
        Sheet.Cells(row, col).value = value
    End If
End Sub
