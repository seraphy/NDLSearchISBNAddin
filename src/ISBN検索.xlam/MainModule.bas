Attribute VB_Name = "MainModule"
Option Explicit

' ================================================================
' ���{����xlam�t�@�C���W�J��������_rels/.rels�t�@�C�����Ɏw�肳�ꂽ
' Relation���ڂ���Q�Ƃ����UI��`xml�Œ�`����.
' �{�t�@�C����customUI/customUI.xml�ɒ�`���Ă���.
' ================================================================

' �V�[�g�ݒ�t�H�[�����J��
' �J�X�^�����{������Ăяo�����
Public Sub ShowConfigFormRibbon(ribbon As IRibbonControl)
    ShowConfigForm
End Sub

' ISBN�������s
' �J�X�^�����{������Ăяo�����
Public Sub UpdateSheetFromRibbon(ribbon As IRibbonControl)
    UpdateSheet
End Sub


' About�\��
' �J�X�^�����{������Ăяo�����
Public Sub ShowAboutFromRibbon(ribbon As IRibbonControl)
    Dim frm As New AboutForm
    frm.Show
End Sub


' ISBN�������s
Sub UpdateSheet()
On Error GoTo Err
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    Dim Sheet As Worksheet
    Set Sheet = ActiveWorkbook.ActiveSheet

    If Sheet Is Nothing Then Exit Sub
    
    UpdateSheetWithConf Sheet
    
    MsgBox "�������܂���"
    Exit Sub

Err:
    MsgBox "Error: " & Err.Description & vbCrLf & Err.Source
    Exit Sub
End Sub

' �V�[�g�̐ݒ�ɏ]���Č��������{����.
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
    
    ' �V�[�g�Ŏg�p����Ă���Z���͈�
    Dim MaxRow As Integer, MaxCol As Integer
    With Sheet.UsedRange
        MaxRow = .Rows(.Rows.Count).row
        MaxCol = .Columns(.Columns.Count).Column
    End With
    
    Dim row As Integer
    Dim ISBN As String
    Dim TITLE As String
    
    row = conf.StartRow
    
    ' �J�n�s����V�[�g�Ŏg�p����Ă���ő�s�܂Ń��[�v����
    Do While row <= MaxRow
        DoEvents
        
        ISBN = Sheet.Cells(row, conf.ISBN).value
        TITLE = Sheet.Cells(row, conf.TITLE).value
        
        If ISBN = "" Then
            ' ISBN���󗓂Ȃ牽�����Ȃ�
        
        ElseIf TITLE = "" Then
            ' �܂��^�C�g�����擾����Ă��Ȃ��s�̂ݏ�������.
            ' ���ݏ�������ISBN���X�e�[�^�X�o�[�ɕ\������.
            Application.StatusBar = "ISBN:" & ISBN
            
            Dim info As BookInfo
            If sercher.FindByISBN(ISBN, info) Then
                TITLE = info.TITLE
                
                If info.VOLUME <> "" And conf.TITLE_WITH_VOLUME Then
                    ' Volume�̎w�肪����ꍇ�̓^�C�g���Ɋ��ʕt���łȂ���.
                    TITLE = TITLE & "(" & info.VOLUME & ")"
                End If
                
                UpdateCell Sheet, row, conf.TITLE, TITLE
                UpdateCell Sheet, row, conf.AUTHOR, info.AUTHOR
                UpdateCell Sheet, row, conf.PUBLISHER, info.PUBLISHER
                UpdateCell Sheet, row, conf.ISSUED, info.ISSUED
                UpdateCell Sheet, row, conf.YOMI, info.YOMI
                UpdateCell Sheet, row, conf.VOLUME, info.VOLUME
            
            Else
                ' ���Џ�񂪎擾�ł��Ȃ�ISBN�̓Z���̔w�i�F��ς��Ă���
                Sheet.Cells(row, conf.ISBN).Interior.ColorIndex = 37
            End If
        End If
        row = row + 1
    Loop
    ' ����������X�e�[�^�X�o�[�͖߂�
    Application.StatusBar = False
End Sub

' �Z����ʒu���L���ł���΁A���̃Z���ɏ�������.
' (0�ȉ��̃Z����ԍ����w�肳��Ă���ꍇ�͏ȗ����ڂƂ��ď������݂��Ȃ�)
Private Sub UpdateCell(Sheet As Worksheet, row As Integer, col As Integer, value As String)
    If col > 0 Then
        Sheet.Cells(row, col).value = value
    End If
End Sub
