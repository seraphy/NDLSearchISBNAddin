Attribute VB_Name = "ConfigModule"
Option Explicit

Private Const APP_NAME As String = "SearchISBNAddin"

Private Const SHEET_SETTING As String = "SheetSetting"
Private Const CONNECT_SETTING As String = "ConnectSetting"

' ====================
'  �V�[�g�ݒ�̍\����
' ====================
Public Type SheetConfig
    ' �J�n�s�ԍ�
    StartRow As Integer
    
    ' ISBN�̗�ԍ�
    ISBN As Integer
    
    ' �^�C�g���̗�ԍ�
    TITLE As Integer
    
    ' ���҂̗�ԍ�(�s�v�Ȃ�0�ȉ�)
    AUTHOR As Integer
    
    ' �o�ŎЂ̗�ԍ�(�s�v�Ȃ�0�ȉ�)
    PUBLISHER As Integer
    
    ' ���s�N�̗�ԍ�(�s�v�Ȃ�0�ȉ�)
    ISSUED As Integer
    
    ' ��݂̗�ԍ�(�s�v�Ȃ�0�ȉ�)
    YOMI As Integer
    
    ' �����̗�ԍ�(�s�v�Ȃ�0�ȉ�)
    VOLUME As Integer
    
    ' �^�C�g�������Ɋ֐���t�^���邩�H
    TITLE_WITH_VOLUME As Boolean
End Type

' ==================
'  �ڑ����̍\����
' ==================
Public Type ConnectConfig
    ' ServerXMLHTTP���g�p���邩�H
    UseServerHTTP As Boolean
    
    ' ServerXMLHTTP�g�p���̃^�C���A�E�g�l(mSec)
    TimeoutMillis As Integer
    
    ' ServerXMLHTTP�g�p����Proxy�L��
    UseProxy As Boolean
    ProxyAddress As String
    ProxyUser As String
    ProxyPassword As String
End Type

Public Function GetSheetConfig() As SheetConfig
' -----------------------------------------------------------
' �V�[�g�ݒ�����W�X�g�����ǂݍ���
'
' Return value:
'   �V�[�g�ݒ�
'
' Remarks:
'   �܂����W�X�g���ɏ������܂�Ă��Ȃ��ꍇ�͏����l���Ԃ����
' -----------------------------------------------------------
    Dim conf As SheetConfig
            
    conf.StartRow = CInt(GetSetting(APP_NAME, SHEET_SETTING, "StartRow", "2"))
    conf.ISBN = CInt(GetSetting(APP_NAME, SHEET_SETTING, "ISBN", "2"))
    conf.TITLE = CInt(GetSetting(APP_NAME, SHEET_SETTING, "TITLE", "3"))
    conf.AUTHOR = CInt(GetSetting(APP_NAME, SHEET_SETTING, "AUTHOR", "4"))
    conf.PUBLISHER = CInt(GetSetting(APP_NAME, SHEET_SETTING, "PUBLISHER", "5"))
    conf.ISSUED = CInt(GetSetting(APP_NAME, SHEET_SETTING, "ISSUED", "6"))
    conf.YOMI = CInt(GetSetting(APP_NAME, SHEET_SETTING, "YOMI", "7"))
    conf.VOLUME = CInt(GetSetting(APP_NAME, SHEET_SETTING, "VOLUME", "-1"))
    conf.TITLE_WITH_VOLUME = CBool(GetSetting(APP_NAME, SHEET_SETTING, "TITLE_WITH_VOLUME", "True"))
    
    GetSheetConfig = conf
End Function

Public Sub SaveSheetConfig(conf As SheetConfig)
' ----------------------------------------------
' �V�[�g�ݒ�����W�X�g���ɏ�������
'
' Parameter:
'   [conf]   IN �V�[�g�ݒ�
' ----------------------------------------------
    SaveSetting APP_NAME, SHEET_SETTING, "StartRow", CStr(conf.StartRow)
    SaveSetting APP_NAME, SHEET_SETTING, "ISBN", CStr(conf.ISBN)
    SaveSetting APP_NAME, SHEET_SETTING, "TITLE", CStr(conf.TITLE)
    SaveSetting APP_NAME, SHEET_SETTING, "AUTHOR", CStr(conf.AUTHOR)
    SaveSetting APP_NAME, SHEET_SETTING, "PUBLISHER", CStr(conf.PUBLISHER)
    SaveSetting APP_NAME, SHEET_SETTING, "ISSUED", CStr(conf.ISSUED)
    SaveSetting APP_NAME, SHEET_SETTING, "YOMI", CStr(conf.YOMI)
    SaveSetting APP_NAME, SHEET_SETTING, "VOLUME", CStr(conf.VOLUME)
    SaveSetting APP_NAME, SHEET_SETTING, "TITLE_WITH_VOLUME", CStr(conf.TITLE_WITH_VOLUME)
End Sub

Public Function GetSheetConfigByWorkbook(wb As Workbook) As SheetConfig
' --------------------------------------------------------------------------
' �V�[�g�ݒ���u�b�N����ǂݍ���
' �܂��u�b�N�ɃV�[�g�ݒ肪�������܂�Ă��Ȃ��ꍇ�̓��W�X�g���A�������͏����l���Ԃ����
'
' Parameter:
'    [wb]   IN ���[�N�u�b�N�ANothing�̏ꍇ�̓��[�N�u�b�N����͓ǂݍ��܂Ȃ�
'
' Return value:
'    �V�[�g�ݒ�
' --------------------------------------------------------------------------
On Error Resume Next ' �J�X�^���v���p�e�B�����ݒ�̏ꍇ�̓G���[�𖳎����Ď����ɐi��
    Dim conf As SheetConfig
    conf = GetSheetConfig() ' �f�t�H���g�l�̎擾
    
    If Not wb Is Nothing Then
        With wb.CustomDocumentProperties
            conf.StartRow = .item("ISBN_ROW")
            conf.ISBN = .item("ISBN_COL")
            conf.TITLE = .item("TITLE_COL")
            conf.AUTHOR = .item("AUTHOR_COL")
            conf.PUBLISHER = .item("PUBLISHER_COL")
            conf.ISSUED = .item("ISSUED_COL")
            conf.YOMI = .item("YOMI_COL")
            conf.VOLUME = .item("VOLUME_COL")
            conf.TITLE_WITH_VOLUME = .item("SUFFIX_VOLUME")
        End With
    End If
    
    GetSheetConfigByWorkbook = conf
End Function

Public Sub SaveSheetConfigByWorkbook(wb As Workbook, conf As SheetConfig)
' --------------------------------------------------------------------------
' �V�[�g�ݒ���u�b�N�ƃ��W�X�g���ɏ�������
'
' Parameter:
'    [wb]   IN ���[�N�u�b�N�ANothing�̏ꍇ�̓��[�N�u�b�N�ɂ͏������܂Ȃ�.
'    [conf] IN �V�[�g�̐ݒ�
' -------------------------------------------------------------------------
    SaveSheetConfig conf ' ����̃f�t�H���g�l�̂��߂Ƀ��W�X�g���ɂ��ۑ�����
    
    If wb Is Nothing Then Exit Sub
    
    ' �ۑ�����l���R���N�V�����ɓ����
    Dim keyValues As New Collection
    AddValue keyValues, "ISBN_ROW", conf.StartRow, msoPropertyTypeNumber
    AddValue keyValues, "ISBN_COL", conf.ISBN, msoPropertyTypeNumber
    AddValue keyValues, "TITLE_COL", conf.TITLE, msoPropertyTypeNumber
    AddValue keyValues, "AUTHOR_COL", conf.AUTHOR, msoPropertyTypeNumber
    AddValue keyValues, "PUBLISHER_COL", conf.PUBLISHER, msoPropertyTypeNumber
    AddValue keyValues, "ISSUED_COL", conf.ISSUED, msoPropertyTypeNumber
    AddValue keyValues, "YOMI_COL", conf.YOMI, msoPropertyTypeNumber
    AddValue keyValues, "VOLUME_COL", conf.VOLUME, msoPropertyTypeNumber
    AddValue keyValues, "SUFFIX_VOLUME", conf.TITLE_WITH_VOLUME, msoPropertyTypeBoolean
    
    ' ���݂̃��[�N�u�b�N�ɂ��łɃJ�X�^���v���p�e�B���o�^�ς݂̏ꍇ�A������X�V����.
    Dim prop As DocumentProperty
    For Each prop In wb.CustomDocumentProperties
        If Contains(keyValues, prop.Name) Then
            prop.value = keyValues(prop.Name)(1)
            keyValues.Remove prop.Name
        End If
    Next
    
    ' ���݂̃��[�N�u�b�N�ɓo�^����Ă��Ȃ��J�X�^���v���p�e�B��o�^����
    Dim keyValue
    For Each keyValue In keyValues
        Dim key, value, typ
        key = keyValue(0)
        value = keyValue(1)
        typ = keyValue(2)
        
        wb.CustomDocumentProperties.Add _
            Name:=key, _
            LinkToContent:=False, _
            Type:=typ, _
            value:=value
            
    Next
End Sub

Private Function Contains(col As Collection, key As Variant) As Boolean
'--------------------------------------------------------------
' �R���N�V�����Ɏw�肳�ꂽ�L�[�����݂��邩���肷��
'
' Parameter:
'    [col] IN �R���N�V����
'    [key] IN �L�[
'
' Return value:
'    �w�肳�ꂽ�L�[���R���N�V�����Ɋ܂܂�Ă���ꍇ��True�A
'    �����łȂ����False
'--------------------------------------------------------------
On Error GoTo Err
    Dim dummy As Variant
    dummy = col(key)
    Contains = True
Err:
End Function

Private Sub AddValue(col As Collection, key As String, item As Variant, typ As Integer)
'--------------------------------------------------------------
' �R���N�V�����ɃL�[�A�l�A�^����ۑ�����
' �L�[�̓L�[�l�Ƃ��ĕۑ������B
' �l�́A�L�[�A�l�A�^���̂R�v�f����Ȃ�Array�Ƃ��ĕۑ������.
'--------------------------------------------------------------
    col.Add key:=key, item:=Array(key, item, typ)
End Sub

Public Function GetConnectConfig() As ConnectConfig
' ---------------------------------------------------
' �ڑ��ݒ�����W�X�g������ǂݍ���
'
' Return value:
'    ���W�X�g������ǂݍ��܂ꂽ�ڑ����
'
' Remarks:
'    �܂����W�X�g���ɏ������܂�Ă��Ȃ��ꍇ�͏����l���Ԃ����.
' ---------------------------------------------------
    Dim conf As ConnectConfig
    
    conf.UseServerHTTP = CBool(GetSetting(APP_NAME, CONNECT_SETTING, "UseServerHTTP", "True"))
    conf.TimeoutMillis = CInt(GetSetting(APP_NAME, CONNECT_SETTING, "TimeoutMillis", "10000"))
    
    conf.UseProxy = CBool(GetSetting(APP_NAME, CONNECT_SETTING, "UseProxy", "False"))
    conf.ProxyAddress = GetSetting(APP_NAME, CONNECT_SETTING, "ProxyAddress", "127.0.0.1:8080")
    conf.ProxyUser = GetSetting(APP_NAME, CONNECT_SETTING, "ProxyUser", "")
    conf.ProxyPassword = GetSetting(APP_NAME, CONNECT_SETTING, "ProxyPassword", "")
    
    GetConnectConfig = conf
End Function

Public Sub SaveConnectConfig(conf As ConnectConfig)
'-------------------------------------------------
' �ڑ��ݒ�����W�X�g���ɏ�������
'
' Parameter:
'   conf [IN] �ڑ����
'-------------------------------------------------
    SaveSetting APP_NAME, CONNECT_SETTING, "UseServerHTTP", CStr(conf.UseServerHTTP)
    SaveSetting APP_NAME, CONNECT_SETTING, "TimeoutMillis", CStr(conf.TimeoutMillis)

    SaveSetting APP_NAME, CONNECT_SETTING, "UseProxy", CStr(conf.UseProxy)
    SaveSetting APP_NAME, CONNECT_SETTING, "ProxyAddress", conf.ProxyAddress
    SaveSetting APP_NAME, CONNECT_SETTING, "ProxyUser", conf.ProxyUser
    SaveSetting APP_NAME, CONNECT_SETTING, "ProxyPassword", conf.ProxyPassword
End Sub

Public Sub ShowConfigForm()
'--------------------------
' �ݒ�_�C�A���O��\������
'--------------------------
    Dim frm As New ConfigForm
    frm.Show
End Sub


