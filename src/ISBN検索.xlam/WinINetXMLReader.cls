VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinINetXMLReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IXMLReader

' ============================================================================
' XMLHTTP(WinINet)を使用してURLからコンテンツを取得するIXMLReaderの実装クラス
' ============================================================================

Private Function IXMLReader_GetContent(URL As String) As String
    ' バージョン非依存のProgIDはMSXML4以降廃止されているので、バージョンの明記が必要.
    Dim XmlHttp As IXMLHTTPRequest
    Set XmlHttp = New MSXML2.XMLHTTP60
    
    XmlHttp.Open "GET", URL, True
    XmlHttp.Send
    
    ' 応答がなかなか返らない場合があるのでブロックしないように
    Do While XmlHttp.readyState <> 4
        DoEvents
    Loop
    
    ' Proxy経由で応答が空になる場合があるようだ。
    If XmlHttp.Status = 0 Then
        Err.Clear
        Err.Raise Number:=999, Description:="接続できなかった可能性があります。"
    End If
    
    IXMLReader_GetContent = XmlHttp.responseText
End Function
