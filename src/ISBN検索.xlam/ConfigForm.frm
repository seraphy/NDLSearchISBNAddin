VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigForm 
   Caption         =   "ISBN�����̐ݒ�"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7065
   OleObjectBlob   =   "ConfigForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub OptClient_Change()
    OnConnectModeChange
End Sub

Private Sub OptServer_Change()
    OnConnectModeChange
End Sub

Private Sub OnConnectModeChange()
    Dim enable As Boolean
    enable = OptServer.value
    lblTimeout.Enabled = enable
    lblTimeoutSuffix.Enabled = enable
    txtTimeout.Enabled = enable
    
    UseProxy_Change ' Proxy�̏�Ԃ��ς�邽��
End Sub

Private Sub UseProxy_Change()
    Dim enable As Boolean
    enable = UseProxy.value And OptServer.value
    lblProxyUser.Enabled = enable
    lblProxyPassword.Enabled = enable
    ProxyAddress.Enabled = enable
    ProxyUser.Enabled = enable
    ProxyPassword.Enabled = enable
End Sub

Private Sub UserForm_Initialize()
    Dim sheetConf As ConfigModule.SheetConfig
    sheetConf = ConfigModule.GetSheetConfigByWorkbook(ActiveWorkbook)
    
    txtStartRow.Text = CStr(sheetConf.StartRow)
    txtISBN.Text = CStr(sheetConf.ISBN)
    txtTITLE.Text = CStr(sheetConf.TITLE)
    txtAUTHOR.Text = CStr(sheetConf.AUTHOR)
    txtPUBLISHER.Text = CStr(sheetConf.PUBLISHER)
    txtISSUED.Text = CStr(sheetConf.ISSUED)
    txtYOMI.Text = CStr(sheetConf.YOMI)
    txtVOLUME.Text = CStr(sheetConf.VOLUME)
    chkTitleWithVolume.value = sheetConf.TITLE_WITH_VOLUME
    
    Dim connectConf As ConfigModule.ConnectConfig
    connectConf = ConfigModule.GetConnectConfig()
    
    If connectConf.UseServerHTTP Then
        OptServer.value = True
    Else
        OptClient.value = True
    End If
    txtTimeout.value = CStr(connectConf.TimeoutMillis)
    
    UseProxy.value = connectConf.UseProxy
    ProxyAddress.value = connectConf.ProxyAddress
    ProxyUser.value = connectConf.ProxyUser
    ProxyPassword.value = connectConf.ProxyPassword
    
    OnConnectModeChange
End Sub

Private Sub btnSave_Click()
On Error GoTo Err
    Dim sheetConf As ConfigModule.SheetConfig
    sheetConf.StartRow = CInt(txtStartRow.Text)
    sheetConf.ISBN = CInt(txtISBN.Text)
    sheetConf.TITLE = CInt(txtTITLE.Text)
    sheetConf.AUTHOR = CInt(txtAUTHOR.Text)
    sheetConf.PUBLISHER = CInt(txtPUBLISHER.Text)
    sheetConf.ISSUED = CInt(txtISSUED.Text)
    sheetConf.YOMI = CInt(txtYOMI.Text)
    sheetConf.VOLUME = CInt(txtVOLUME.Text)
    sheetConf.TITLE_WITH_VOLUME = chkTitleWithVolume.value
    
    ' �K�{�J�����̃`�F�b�N
    If sheetConf.StartRow < 1 Or sheetConf.ISBN < 1 Or sheetConf.TITLE < 1 Then
        MsgBox "�J�n�s�AISBN��A�^�C�g�����1�ȏ�łȂ���΂Ȃ�܂���", vbExclamation
        Exit Sub
    End If
    
    Dim connectConf As ConfigModule.ConnectConfig
    connectConf.UseServerHTTP = OptServer.value
    connectConf.TimeoutMillis = CInt(txtTimeout.value)
    
    connectConf.UseProxy = UseProxy.value
    connectConf.ProxyAddress = ProxyAddress.value
    connectConf.ProxyUser = ProxyUser.value
    connectConf.ProxyPassword = ProxyPassword.value
    
    ConfigModule.SaveSheetConfigByWorkbook ActiveWorkbook, sheetConf
    ConfigModule.SaveConnectConfig connectConf
    
    Me.Hide
    Exit Sub
    
Err:
    MsgBox "ERROR: " & Err.Description
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

