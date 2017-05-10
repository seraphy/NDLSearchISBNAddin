Attribute VB_Name = "SearchISBNModule"
Option Explicit

' ==================
'  図書情報の構造体
' ==================
Public Type BookInfo
    ' タイトル
    TITLE As String
    
    ' 著者
    AUTHOR As String
    
    ' 発行元
    PUBLISHER As String
    
    ' 発行年
    ISSUED As String
    
    ' 巻数
    VOLUME As String
    
    ' タイトルのよみがな
    YOMI As String
End Type

