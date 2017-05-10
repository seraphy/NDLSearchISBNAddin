VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHTTPXMLReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IXMLReader

' =================================================================================
' ServerXMLHTTP(WinHTTP)を使用してURLからコンテンツを取得するIXMLReaderの実装クラス
' =================================================================================

' ServerXMLHTTP使用時のタイムアウト
Public TimeoutMillis As Integer

' ServerXMLHTTP使用時のProxyの設定
Public UseProxy As Boolean
Public ProxyAddress As String
Public ProxyUser As String
Public ProxyPassword As String

' クラスの初期化
Private Sub Class_Initialize()
    TimeoutMillis = 7000
End Sub

Private Function IXMLReader_GetContent(URL As String) As String
    ' バージョン非依存のProgIDはMSXML4以降廃止されているので、バージョンの明記が必要.
    Dim XmlHttp As IServerXMLHTTPRequest
    Set XmlHttp = New MSXML2.ServerXMLHTTP60
        
    If UseProxy And Len(ProxyAddress) > 0 Then
        XmlHttp.setProxy 2, ProxyAddress, "localhost" ' プロキシとバイパスリスト
        Debug.Print "ServerXMLHTTP: Proxy=" & ProxyAddress
    End If
    
    ' タイムアウトを設定する (Resolve, Connect, Send, Receive)
    If TimeoutMillis > 0 Then
        XmlHttp.setTimeouts TimeoutMillis, TimeoutMillis, TimeoutMillis, TimeoutMillis
        Debug.Print "ServerXMLHTTP: Timeout Millis=" & TimeoutMillis
    End If
    
    XmlHttp.Open "GET", URL, True
    
    If UseProxy And Len(ProxyUser) > 0 Then
        XmlHttp.setProxyCredentials ProxyUser, ProxyPassword
        Debug.Print "ServerXMLHTTP: Proxy user:" & ProxyUser
    End If
    
    XmlHttp.Send
    
    ' 応答がなかなか返らない場合があるのでブロックしないように
    XmlHttp.waitForResponse
    
    ' Proxy経由で応答が空になる場合があるようだ。
    If XmlHttp.Status = 0 Then
        Err.Clear
        Err.Raise Number:=999, Description:="接続できなかった可能性があります。"
    End If
    
    IXMLReader_GetContent = XmlHttp.responseText
End Function

