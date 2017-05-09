Attribute VB_Name = "ConfigModule"
Option Explicit

Private Const APP_NAME As String = "SearchISBNAddin"

Private Const SHEET_SETTING As String = "SheetSetting"
Private Const CONNECT_SETTING As String = "ConnectSetting"

' シート設定の構造体
Public Type SheetConfig
    ' 開始行番号
    StartRow As Integer
    
    ' ISBNの列番号
    ISBN As Integer
    
    ' タイトルの列番号
    TITLE As Integer
    
    ' 著者の列番号(不要なら0以下)
    AUTHOR As Integer
    
    ' 出版社の列番号(不要なら0以下)
    PUBLISHER As Integer
    
    ' 発行年の列番号(不要なら0以下)
    ISSUED As Integer
    
    ' よみの列番号(不要なら0以下)
    YOMI As Integer
    
    ' 巻数の列番号(不要なら0以下)
    VOLUME As Integer
    
    ' タイトル末尾に関数を付与するか？
    TITLE_WITH_VOLUME As Boolean
End Type

' 接続情報の構造体
Public Type ConnectConfig
    ' ServerXMLHTTPを使用するか？
    UseServerHTTP As Boolean
    
    ' ServerXMLHTTP使用時のタイムアウト値(mSec)
    TimeoutMillis As Integer
    
    ' ServerXMLHTTP使用時のProxy有無
    UseProxy As Boolean
    ProxyAddress As String
    ProxyUser As String
    ProxyPassword As String
End Type

Public Function GetSheetConfig() As SheetConfig
    Dim conf As SheetConfig
            
    conf.StartRow = CInt(GetSetting(APP_NAME, SHEET_SETTING, "StartRow", "2"))
    conf.ISBN = CInt(GetSetting(APP_NAME, SHEET_SETTING, "ISBN", "2"))
    conf.TITLE = CInt(GetSetting(APP_NAME, SHEET_SETTING, "TITLE", "3"))
    conf.AUTHOR = CInt(GetSetting(APP_NAME, SHEET_SETTING, "AUTHOR", "4"))
    conf.PUBLISHER = CInt(GetSetting(APP_NAME, SHEET_SETTING, "PUBLISHER", "5"))
    conf.ISSUED = CInt(GetSetting(APP_NAME, SHEET_SETTING, "ISSUED", "6"))
    conf.YOMI = CInt(GetSetting(APP_NAME, SHEET_SETTING, "YOMI", "7"))
    conf.VOLUME = CInt(GetSetting(APP_NAME, SHEET_SETTING, "VOLUME", "-1"))
    conf.TITLE_WITH_VOLUME = CBool(GetSetting(APP_NAME, SHEET_SETTING, "TITLE_WITH_VOLUME", "True"))
    
    GetSheetConfig = conf
End Function

Public Sub SaveSheetConfig(conf As SheetConfig)
    SaveSetting APP_NAME, SHEET_SETTING, "StartRow", CStr(conf.StartRow)
    SaveSetting APP_NAME, SHEET_SETTING, "ISBN", CStr(conf.ISBN)
    SaveSetting APP_NAME, SHEET_SETTING, "TITLE", CStr(conf.TITLE)
    SaveSetting APP_NAME, SHEET_SETTING, "AUTHOR", CStr(conf.AUTHOR)
    SaveSetting APP_NAME, SHEET_SETTING, "PUBLISHER", CStr(conf.PUBLISHER)
    SaveSetting APP_NAME, SHEET_SETTING, "ISSUED", CStr(conf.ISSUED)
    SaveSetting APP_NAME, SHEET_SETTING, "YOMI", CStr(conf.YOMI)
    SaveSetting APP_NAME, SHEET_SETTING, "VOLUME", CStr(conf.VOLUME)
    SaveSetting APP_NAME, SHEET_SETTING, "TITLE_WITH_VOLUME", CStr(conf.TITLE_WITH_VOLUME)
End Sub

Public Function GetSheetConfigByWorkbook(wb As Workbook) As SheetConfig
On Error Resume Next ' カスタムプロパティが未設定の場合はエラーを無視して次ぎに進む
    Dim conf As SheetConfig
    conf = GetSheetConfig() ' デフォルト値の取得
    
    If Not wb Is Nothing Then
        With wb.CustomDocumentProperties
            conf.StartRow = .item("ISBN_ROW")
            conf.ISBN = .item("ISBN_COL")
            conf.TITLE = .item("TITLE_COL")
            conf.AUTHOR = .item("AUTHOR_COL")
            conf.PUBLISHER = .item("PUBLISHER_COL")
            conf.ISSUED = .item("ISSUED_COL")
            conf.YOMI = .item("YOMI_COL")
            conf.VOLUME = .item("VOLUME_COL")
            conf.TITLE_WITH_VOLUME = .item("SUFFIX_VOLUME")
        End With
    End If
    
    GetSheetConfigByWorkbook = conf
End Function

Public Sub SaveSheetConfigByWorkbook(wb As Workbook, conf As SheetConfig)
    SaveSheetConfig conf ' 次回のデフォルト値のためにレジストリにも保存する
    
    If wb Is Nothing Then Exit Sub
    
    ' 保存する値をコレクションに入れる
    Dim keyValues As New Collection
    AddValue keyValues, "ISBN_ROW", conf.StartRow, msoPropertyTypeNumber
    AddValue keyValues, "ISBN_COL", conf.ISBN, msoPropertyTypeNumber
    AddValue keyValues, "TITLE_COL", conf.TITLE, msoPropertyTypeNumber
    AddValue keyValues, "AUTHOR_COL", conf.AUTHOR, msoPropertyTypeNumber
    AddValue keyValues, "PUBLISHER_COL", conf.PUBLISHER, msoPropertyTypeNumber
    AddValue keyValues, "ISSUED_COL", conf.ISSUED, msoPropertyTypeNumber
    AddValue keyValues, "YOMI_COL", conf.YOMI, msoPropertyTypeNumber
    AddValue keyValues, "VOLUME_COL", conf.VOLUME, msoPropertyTypeNumber
    AddValue keyValues, "SUFFIX_VOLUME", conf.TITLE_WITH_VOLUME, msoPropertyTypeBoolean
    
    ' 現在のワークブックにすでにカスタムプロパティが登録済みの場合、それを更新する.
    Dim prop As DocumentProperty
    For Each prop In wb.CustomDocumentProperties
        If Contains(keyValues, prop.Name) Then
            prop.value = keyValues(prop.Name)(1)
            keyValues.Remove prop.Name
        End If
    Next
    
    ' 現在のワークブックに登録されていないカスタムプロパティを登録する
    Dim keyValue
    For Each keyValue In keyValues
        Dim key, value, typ
        key = keyValue(0)
        value = keyValue(1)
        typ = keyValue(2)
        
        wb.CustomDocumentProperties.Add _
            Name:=key, _
            LinkToContent:=False, _
            Type:=typ, _
            value:=value
            
    Next
End Sub

' コレクションに指定されたキーが存在するか判定する
Private Function Contains(col As Collection, key As Variant) As Boolean
On Error GoTo Err
    Dim dummy As Variant
    dummy = col(key)
    Contains = True
Err:
End Function

' コレクションにキー、値、型情報を保存する
' キーはキー値として保存される。値は、キー、値、型情報の３要素からなるArrayとして保存される.
Private Sub AddValue(col As Collection, key As String, item As Variant, typ As Integer)
    col.Add key:=key, item:=Array(key, item, typ)
End Sub


Public Function GetConnectConfig() As ConnectConfig
    Dim conf As ConnectConfig
    
    conf.UseServerHTTP = CBool(GetSetting(APP_NAME, CONNECT_SETTING, "UseServerHTTP", "True"))
    conf.TimeoutMillis = CInt(GetSetting(APP_NAME, CONNECT_SETTING, "TimeoutMillis", "10000"))
    
    conf.UseProxy = CBool(GetSetting(APP_NAME, CONNECT_SETTING, "UseProxy", "False"))
    conf.ProxyAddress = GetSetting(APP_NAME, CONNECT_SETTING, "ProxyAddress", "127.0.0.1:8080")
    conf.ProxyUser = GetSetting(APP_NAME, CONNECT_SETTING, "ProxyUser", "")
    conf.ProxyPassword = GetSetting(APP_NAME, CONNECT_SETTING, "ProxyPassword", "")
    
    GetConnectConfig = conf
End Function

Public Sub SaveConnectConfig(conf As ConnectConfig)
    SaveSetting APP_NAME, CONNECT_SETTING, "UseServerHTTP", CStr(conf.UseServerHTTP)
    SaveSetting APP_NAME, CONNECT_SETTING, "TimeoutMillis", CStr(conf.TimeoutMillis)

    SaveSetting APP_NAME, CONNECT_SETTING, "UseProxy", CStr(conf.UseProxy)
    SaveSetting APP_NAME, CONNECT_SETTING, "ProxyAddress", conf.ProxyAddress
    SaveSetting APP_NAME, CONNECT_SETTING, "ProxyUser", conf.ProxyUser
    SaveSetting APP_NAME, CONNECT_SETTING, "ProxyPassword", conf.ProxyPassword
End Sub

Public Sub ShowConfigForm()
    Dim frm As New ConfigForm
    frm.Show
End Sub


