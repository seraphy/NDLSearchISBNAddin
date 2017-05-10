VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISearchISBN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' VBAによるインターフェイス定義
''' 実装は空にしておくこと
''' 参考: http://www.cpearson.com/excel/Implements.aspx

Public Function FindByISBN(ISBN As String, ByRef info As BookInfo) As Boolean
' ------------------------------------------------
' ISBNを指定して書籍情報を取得してTrueを返す.
' 書籍情報が取得できない場合は書籍情報は未設定のままFalseを返す.
'
' Parameter:
'   [ISBN] IN ISBN文字列
'   [INFO] IN/OUT 書籍情報(書き込み先)
'
' Return value:
'   書籍情報が取得できない場合はFalseを返す.
' ------------------------------------------------
    ' declaration only. no code allowed.
End Function
