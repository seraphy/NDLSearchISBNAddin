VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SearchISBN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''
''' 国立国会図書館サーチのWebAPIを利用して，
''' ISBNをキーに国会図書館が公開する図書の書誌情報を抽出する.
'''
''' 参照設定として"Microsoft XML, v6"が必要である.
'''
''' LOD Challenge 2013 書誌検索シート 原田隆史 同志社大学 のマクロをベースにしている。
''' http://lod.sfc.keio.ac.jp/challenge2013/show_status.php?id=a072
'''

Option Explicit

' ServerXMLHTTPを使うか？
' Falseの場合はMSXML2.XMLHTTP, Trueの場合はMSXML2.ServerXMLHTTPが用いられる
Public UseServerHTTP As Boolean

' ServerXMLHTTP使用時のタイムアウト
Public TimeoutMillis As Integer

' ServerXMLHTTP使用時のProxyの設定
Public UseProxy As Boolean
Public ProxyAddress As String
Public ProxyUser As String
Public ProxyPassword As String

' クラスの初期化
Private Sub Class_Initialize()
    UseServerHTTP = True
    TimeoutMillis = 7000
End Sub

' ISBNから書籍タイプを設定してTrueを返す.
' 取得できない場合はFalseを返す.
Public Function FindByISBN(ISBN As String, ByRef info As BookInfo) As Boolean
    FindByISBN = False
    Dim LenISBN As Integer
    LenISBN = Len(ISBN)
    If LenISBN = 13 Or LenISBN = 10 Then
         Dim URL As String
         URL = "http://iss.ndl.go.jp/books.rss?gpid=iss-ndl-opac&search_mode=advanced&rft.isbn=" & ISBN
         
         Dim Xml As String
         Xml = GetContents(URL)
         
         Dim ItemNode As MSXML2.IXMLDOMNode
         Set ItemNode = GetItemNode("/rss/channel/item", Xml)
         
         If Not ItemNode Is Nothing Then
            info.TITLE = GetText(ItemNode, "title")
            info.AUTHOR = GetText(ItemNode, "author")
            info.PUBLISHER = GetText(ItemNode, "dc:publisher")
            info.ISSUED = GetText(ItemNode, "dcterms:issued")
            info.VOLUME = GetText(ItemNode, "dcndl:volume")
            info.YOMI = GetText(ItemNode, "dcndl:titleTranscription")
            FindByISBN = True
         End If
    End If
End Function

' URLを指定してXMLテキストを読み込む
' WinHTTPコンポーネントを使うServerXMLHTT60を使用した場合、
' Proxyの手動設定、タイムアウト検知等々が行うことができる.
Function GetContents(URL As String) As String
    
    ' バージョン非依存のProgIDはMSXML4以降廃止されているので、バージョンの明記が必要.
    Dim XmlHttp As IXMLHTTPRequest
    If UseServerHTTP Then
        Set XmlHttp = New MSXML2.ServerXMLHTTP60
        Debug.Print "Use ServerXMLHTTP"
    
    ' ServerXMLHTTP(WinHTTP使用)でプロキシの手動設定が必要な場合
        If UseProxy And Len(ProxyAddress) > 0 Then
            XmlHttp.setProxy 2, ProxyAddress, "localhost" ' プロキシとバイパスリスト
            Debug.Print "ServerXMLHTTP: Proxy=" & ProxyAddress
        End If
        
        ' タイムアウトを設定する (Resolve, Connect, Send, Receive)
        If TimeoutMillis > 0 Then
            XmlHttp.setTimeouts TimeoutMillis, TimeoutMillis, TimeoutMillis, TimeoutMillis
            Debug.Print "ServerXMLHTTP: Timeout Millis=" & TimeoutMillis
        End If
    Else
        Set XmlHttp = New MSXML2.XMLHTTP60
        Debug.Print "Use XMLHTTP"
    End If
    
    Debug.Print "Request: " & URL
    XmlHttp.Open "GET", URL, True
    
    ' ServerXMLHTTP(WinHTTP使用)でプロキシの手動設定が必要な場合
    If TypeOf XmlHttp Is IServerXMLHTTPRequest Then
        If UseProxy And Len(ProxyUser) > 0 Then
            XmlHttp.setProxyCredentials ProxyUser, ProxyPassword
            Debug.Print "ServerXMLHTTP: Proxy user:" & ProxyUser
        End If
    End If
    
    XmlHttp.Send
    
    ' 応答がなかなか返らない場合があるのでブロックしないように
    If TypeOf XmlHttp Is IServerXMLHTTPRequest Then
        XmlHttp.waitForResponse
    Else
        Do While XmlHttp.readyState <> 4
            DoEvents
        Loop
    End If
    
    ' Proxy経由で応答が空になる場合があるようだ。
    If XmlHttp.Status = 0 Then
        Err.Clear
        Err.Raise Number:=999, Description:="接続できなかった可能性があります。"
    End If
    
    GetContents = XmlHttp.responseText
End Function

' 国会図書館サーチのRSSから、書籍情報のあるノードを取得する.
' 該当が無い場合はNothingを返す.
Function GetItemNode(NodePath As String, Xml As String) As MSXML2.IXMLDOMNode
    Dim Dom As New MSXML2.DOMDocument60
    
    If Not Dom.LoadXML(Xml) Then Exit Function
    
    ' 名前空間つきのXMLをXPathで取得するためには事前定義が必要
    Dom.setProperty "SelectionLanguage", "XPath"
    Dom.setProperty "SelectionNamespaces", _
        "xmlns:atom='http://www.w3.org/2005/Atom' " & _
        "xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' " & _
        "xmlns:rdfs='http://www.w3.org/2000/01/rdf-schema#' " & _
        "xmlns:opensearch='http://a9.com/-/spec/opensearch/1.1/' " & _
        "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " & _
        "xmlns:dcndl='http://ndl.go.jp/dcndl/terms/' " & _
        "xmlns:dc='http://purl.org/dc/elements/1.1/' " & _
        "xmlns:dcterms='http://purl.org/dc/terms/'"
 
    Dim Nodes As MSXML2.IXMLDOMNodeList
    Set Nodes = Dom.DocumentElement.SelectNodes(NodePath)
    
    Dim Node As MSXML2.IXMLDOMNode
    For Each Node In Nodes
        Set GetItemNode = Node
        Exit For ' 最初の一件で打ち切り
    Next
End Function

' XPathでノードを指定して、そのノードのテキストを取得する.
Function GetText(ItemNode As MSXML2.IXMLDOMNode, XPath As String) As String
    Dim NodeList As MSXML2.IXMLDOMNodeList
    Set NodeList = ItemNode.SelectNodes(XPath)
    
    Dim Node As MSXML2.IXMLDOMNode
    For Each Node In NodeList
        GetText = Node.Text
        Exit Function ' 最初の一件で打ち切り
    Next
End Function

