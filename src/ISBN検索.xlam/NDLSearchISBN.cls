VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NDLSearchISBN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISearchISBN

' ----------------------------------
' XMLの読み込みに使うリーダー(必須)
' ----------------------------------
Public XmlReader As IXMLReader

Function ISearchISBN_FindByISBN(ISBN As String, ByRef info As BookInfo) As Boolean
' ------------------------------------------------------------------------------
' 国立国会図書館サーチWebAPIを用いてISBNを指定して書籍情報を取得してTrueを返す.
' 書籍情報が取得できない場合は書籍情報は未設定のままFalseを返す.
'
' Parameter:
'   [ISBN] IN ISBN文字列
'   [INFO] IN/OUT 書籍情報(書き込み先)
'
' Return value:
'   書籍情報が取得できない場合はFalseを返す.
'
' Remarks:
'   WebAPIとして OpenSearchを用いている.
'   http://iss.ndl.go.jp/api/opensearch?isbn=???? でアクセスしている.
' ------------------------------------------------------------------------------
    ISearchISBN_FindByISBN = False
    Dim LenISBN As Integer
    LenISBN = Len(ISBN)
    If LenISBN = 13 Or LenISBN = 10 Then
         Dim URL As String
         URL = "http://iss.ndl.go.jp/api/opensearch?isbn=" & ISBN
         
         Dim XML As String
         XML = XmlReader.GetContent(URL)
         
         Dim ItemNode As MSXML2.IXMLDOMNode
         Set ItemNode = GetItemNode("/rss/channel/item", XML)
         
         If Not ItemNode Is Nothing Then
            info.TITLE = GetText(ItemNode, "title")
            info.AUTHOR = GetText(ItemNode, "author")
            info.PUBLISHER = GetText(ItemNode, "dc:publisher")
            info.ISSUED = GetText(ItemNode, "dcterms:issued")
            info.VOLUME = GetText(ItemNode, "dcndl:volume")
            info.YOMI = GetText(ItemNode, "dcndl:titleTranscription")
            ISearchISBN_FindByISBN = True
         End If
    End If
End Function

Function GetItemNode(NodePath As String, XML As String) As MSXML2.IXMLDOMNode
' ------------------------------------------------------------
' 国会図書館サーチのRSSから、書籍情報のあるノードを取得する.
'
' Parameter:
'   [XML] IN XML文字列
'
' Return value:
'   書籍情報のある最初のIXMLNodeを返す.
'   該当が無い場合、もしくはXMLとして妥当でない場合はNothingを返す.
' ------------------------------------------------------------
    Dim Dom As New MSXML2.DOMDocument60
    
    If Not Dom.LoadXML(XML) Then Exit Function
    
    ' 名前空間つきのXMLをXPathで取得するためには事前定義が必要
    Dom.setProperty "SelectionLanguage", "XPath"
    Dom.setProperty "SelectionNamespaces", _
        "xmlns:dcmitype='http://purl.org/dc/dcmitype/' " & _
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

Function GetText(ItemNode As MSXML2.IXMLDOMNode, XPath As String) As String
'-----------------------------------------------------------
' XPathでノードを指定して、そのノードのテキストを取得する.
'
' Parameter:
'    [ItemNode] IN 親ノード
'    [XPath]    IN XPath
'
' Return value:
'    指定されたノードのテキスト、該当が無い場合は空文字を返す.
'-----------------------------------------------------------
    Dim NodeList As MSXML2.IXMLDOMNodeList
    Set NodeList = ItemNode.SelectNodes(XPath)
    
    Dim Node As MSXML2.IXMLDOMNode
    For Each Node In NodeList
        GetText = Node.Text
        Exit Function ' 最初の一件で打ち切り
    Next
End Function

