VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLReaderFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Create(connInfo As ConnectConfig) As IXMLReader
' ----------------------------------------------------------------------
' 接続設定からWinINet, WinHTTPのいずれかのXMLReaderを作成して返す
'
' Parameter:
'     [connInfo] IN 接続情報
'
' Return value:
'     IXMLReaderの具象クラスのインスタンス
' ----------------------------------------------------------------------
    Dim XmlReader As IXMLReader
    If connInfo.UseServerHTTP Then
        Dim xmlReaderImpl As New WinHTTPXMLReader
        With xmlReaderImpl
            .TimeoutMillis = connInfo.TimeoutMillis
            .UseProxy = connInfo.UseProxy
            .ProxyAddress = connInfo.ProxyAddress
            .ProxyUser = connInfo.ProxyUser
            .ProxyPassword = connInfo.ProxyPassword
        End With
        Set XmlReader = xmlReaderImpl
    Else
        Set XmlReader = New WinINetXMLReader
    End If

    Set Create = XmlReader
End Function
