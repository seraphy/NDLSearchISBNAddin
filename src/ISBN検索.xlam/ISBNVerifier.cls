VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISBNVerifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private base10() As Variant
Private base13() As Variant

Private Sub Class_Initialize()
    ' チェックサム係数。最後の桁はチェックサム自身なので係数0で計算値から除外する
    base10 = Array(10, 9, 8, 7, 6, 5, 4, 3, 2, 0)
    base13 = Array(1, 3, 1, 3, 1, 3, 1, 3, 1, 3, 1, 3, 0)
End Sub

Public Function Verify(ISBN As String) As Boolean
' ----------------------------------------------------
' ISBN10/ISBN13が妥当であるかチェックする
' ISBNはハイフンで区切られていても良い
' 妥当であればTrue、そうでなければFalseを返す
'
' Parameter:
'  [ISBN] IN ISBN文字列
'
' Return value:
' 妥当であればTrue、そうでなければFalse
' ----------------------------------------------------
On Error GoTo Err
    ISBN = Replace(ISBN, "-", "")
    ISBN = Replace(ISBN, " ", "")
    
    Dim LenISBN As Integer
    LenISBN = Len(ISBN)
    If LenISBN = 10 Then
        Verify = ISBN10Verification(ISBN)
    ElseIf LenISBN = 13 Then
        Verify = ISBN13Verification(ISBN)
    End If
Err:
End Function

Private Function ISBN10Verification(ISBN As String) As Boolean
' ----------------------------------------------------
' ISBN10桁の妥当性検査を行う.
' 与えられた10桁の文字数がISBN10として妥当であるか
' チェックデジット計算値が一致するか検証する.
'
' Parameter:
'  [ISBN] IN ISBN文字列
'
' Return value:
' 妥当であればTrue、そうでなければFalse
' ----------------------------------------------------
    ' チェックサムの算定
    Dim i As Integer
    Dim ch As Integer
    Dim sum As Integer
    For i = 1 To 10
        ch = Asc(Mid(ISBN, i, 1)) - &H30
        If ch > 10 Then ' チェックデジットはXがありえる。その場合は10
            ch = 10
        End If
        sum = sum + ch * base10(i - 1)
    Next
    
    ' 最後の数値とチェックサム値が一致しているか？
    sum = 11 - (sum Mod 11)
    If sum = 11 Then sum = 0
    Debug.Print sum & ":" & ch
    ISBN10Verification = (sum = ch)
End Function

Private Function ISBN13Verification(ISBN As String) As Boolean
' ----------------------------------------------------
' ISBN13桁の妥当性検査を行う.
' 与えられた13桁の文字数がISBN13として妥当であるか
' 先頭978-979ではじまっており、
' 且つ、チェックデジット計算値が一致するか検証する.
'
' Parameter:
'  [ISBN] IN ISBN文字列
'
' Return value:
' 妥当であればTrue、そうでなければFalse
' ----------------------------------------------------
On Error GoTo Err
    ' ISBN13は978, 979で始まるはず
    Dim lead As Integer
    lead = CInt(Left(ISBN, 3))
    If lead < 978 Or lead > 979 Then Exit Function
    
    ' チェックサムの算定
    Dim i As Integer
    Dim ch As Integer
    Dim sum As Integer
    For i = 1 To 13
        ch = Asc(Mid(ISBN, i, 1)) - &H30
        sum = sum + ch * base13(i - 1)
    Next
    
    ' 最後の数値とチェックサム値が一致しているか？
    sum = 10 - (sum Mod 10)
    If sum = 10 Then sum = 0
    Debug.Print sum & ":" & ch
    ISBN13Verification = (sum = ch)
Err:
End Function

