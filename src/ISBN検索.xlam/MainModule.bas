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


' ISBN検索実行
Sub UpdateSheet()
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

' シートの設定に従って検索を実施する.
Public Sub UpdateSheetWithConf(Sheet As Worksheet)
    Dim conf As ConfigModule.SheetConfig
    Dim connectConf As ConfigModule.ConnectConfig
    conf = ConfigModule.GetSheetConfig()
    connectConf = ConfigModule.GetConnectConfig()
    
    Dim sercher As New SearchISBN
    
    sercher.UseServerHTTP = connectConf.UseServerHTTP
    sercher.TimeoutMillis = connectConf.TimeoutMillis
    
    sercher.UseProxy = connectConf.UseProxy
    sercher.ProxyAddress = connectConf.ProxyAddress
    sercher.ProxyUser = connectConf.ProxyUser
    sercher.ProxyPassword = connectConf.ProxyPassword
    
    ' シートで使用されているセル範囲
    Dim MaxRow As Integer, MaxCol As Integer
    With Sheet.UsedRange
        MaxRow = .Rows(.Rows.Count).row
        MaxCol = .Columns(.Columns.Count).Column
    End With
    
    Dim row As Integer
    Dim ISBN As String
    Dim TITLE As String
    
    row = conf.StartRow
    
    ' 開始行からシートで使用されている最大行までループする
    Do While row <= MaxRow
        DoEvents
        
        ISBN = Sheet.Cells(row, conf.ISBN).value
        TITLE = Sheet.Cells(row, conf.TITLE).value
        
        If ISBN = "" Then
            ' ISBNが空欄なら何もしない
        
        ElseIf TITLE = "" Then
            ' まだタイトルが取得されていない行のみ処理する.
            ' 現在処理中のISBNをステータスバーに表示する.
            Application.StatusBar = "ISBN:" & ISBN
            
            Dim info As BookInfo
            If sercher.FindByISBN(ISBN, info) Then
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
        End If
        row = row + 1
    Loop
    ' 完了したらステータスバーは戻す
    Application.StatusBar = False
End Sub

' セル列位置が有効であれば、そのセルに書き込む.
' (0以下のセル列番号が指定されている場合は省略項目として書き込みしない)
Private Sub UpdateCell(Sheet As Worksheet, row As Integer, col As Integer, value As String)
    If col > 0 Then
        Sheet.Cells(row, col).value = value
    End If
End Sub
