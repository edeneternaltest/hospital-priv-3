Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Friend Class frmNSK0000HG
    Inherits General.FormBase
	'=====================================================================================
	'   ��  ��  ��  ��
	'=====================================================================================
	
	'�E����񍀖� �萔
	Private Const M_StaffID As Short = 1 '�E���Ǘ��ԍ�
	Private Const M_SaiyoYMD As Short = 2 '�̗p�N����
	Private Const M_JobHyojiNo As Short = 3 '�E��\��No
	Private Const M_PostHyojiNo As Short = 4 '��E�\��No
	Private Const M_GiryoHyojiNo As Short = 5 '�Z�ʕ\��No
	Private Const M_StaffHyojiNo As Short = 6 '�E���\��No
	Private Const M_Team As Short = 7 '���
    Private Const M_SyokuinNo As Short = 8 '�E���ԍ�
	Private Const M_SortTeam As Short = 9 '�\�[�g�p
    Private Const M_HairetuNo As Short = 10 '�ް��i�[�z��ԍ�
	
	'-----------------------------------------------------------------
	'   �� �� �� ��
	'-----------------------------------------------------------------
    Private m_FormShowFlg As Boolean '��ʂ��\������Ă��邩
	Private m_SortKey1 As Short '1�Ԗڂ̃\�[�g�L�[
	Private m_SortKey2 As Short '2�Ԗڂ̃\�[�g�L�[
	Private m_SortKey3 As Short '3�Ԗڂ̃\�[�g�L�[
    Private m_SortOrder1 As Boolean '1�Ԗڂ̃\�[�g�L�[�̕��я�
    Private m_SortOrder2 As Boolean '2�Ԗڂ̃\�[�g�L�[�̕��я�
    Private m_SortOrder3 As Boolean '3�Ԗڂ̃\�[�g�L�[�̕��я�
    Private m_lstCmbSortKey As New List(Of Object)

	'------------------------------------------------------------------
	'  ����Đ錾
	'------------------------------------------------------------------
    Event MsgBox_Renamed(ByVal pMsgNo As String, ByRef p_strMsg As System.Array)
    Event Sort()

	Public Property pShowFlg() As Boolean
		Get
			pShowFlg = m_FormShowFlg
        End Get

		Set(ByVal Value As Boolean)
			m_FormShowFlg = Value
		End Set
	End Property
	
	Public ReadOnly Property pSortKey1() As Short
		Get
			pSortKey1 = m_SortKey1
		End Get
	End Property
	
    Public ReadOnly Property pSortOrder1() As Boolean
        Get
            pSortOrder1 = m_SortOrder1
        End Get
    End Property
	
    Public ReadOnly Property pSortOrder2() As Boolean
        Get
            pSortOrder2 = m_SortOrder2
        End Get
    End Property
	
    Public ReadOnly Property pSortOrder3() As Boolean
        Get
            pSortOrder3 = m_SortOrder3
        End Get
    End Property
	
    Public ReadOnly Property pSortKey2() As Short
        Get
            pSortKey2 = m_SortKey2
        End Get
    End Property
	
	Public ReadOnly Property pSortKey3() As Short
		Get
			pSortKey3 = m_SortKey3
		End Get
	End Property
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        '��\��
		Me.Hide()
		m_FormShowFlg = False
    End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		On Error GoTo cmdOK_Click
		Const W_SUBNAME As String = "NSK0000HG cmdOK_Click"
		
		Dim w_Index1 As Integer
		Dim w_Index2 As Integer
		Dim w_Index3 As Integer
        Dim w_strMsg() As String

		'�I���L�[�̲��ޯ���擾
        w_Index1 = m_lstCmbSortKey(0).SelectedIndex + 1
        w_Index2 = m_lstCmbSortKey(1).SelectedIndex
        w_Index3 = m_lstCmbSortKey(2).SelectedIndex
		
		'�L�[���I������Ă��邩
		If w_Index1 = 0 Then
			'�P�Ԗڂ̃L�[�����I���̂Ƃ�(�Ȃ��̂Ƃ�)
			ReDim w_strMsg(3)
			w_strMsg(1) = "�P�Ԗ�"
			w_strMsg(2) = "�\�[�g�L�["
			w_strMsg(3) = "�L�["
			RaiseEvent MsgBox_Renamed("NS0100", w_strMsg)
			
			Exit Sub
		End If
		
		If w_Index2 = 0 And w_Index3 <> 0 Then
			'�Q�Ԗڂ̃L�[�����I���łR�Ԗڂ̃L�[���I������Ă���Ƃ�
			ReDim w_strMsg(3)
			w_strMsg(1) = "�Q�Ԗ�"
			w_strMsg(2) = "�\�[�g�L�["
			w_strMsg(3) = "�L�["
			RaiseEvent MsgBox_Renamed("NS0100", w_strMsg)
			Exit Sub
		End If
		
		If w_Index1 = w_Index2 Then
			ReDim w_strMsg(0)
			RaiseEvent MsgBox_Renamed("NS0101", w_strMsg)
			Exit Sub
		ElseIf w_Index1 = w_Index3 Then 
			ReDim w_strMsg(0)
			RaiseEvent MsgBox_Renamed("NS0101", w_strMsg)
			Exit Sub
		ElseIf w_Index2 = w_Index3 Then 
			If w_Index2 <> 0 Then
				ReDim w_strMsg(0)
				RaiseEvent MsgBox_Renamed("NS0101", w_strMsg)
				Exit Sub
			End If
        End If
		
		'�I���L�[�̲��ޯ���ԍ����i�[
		m_SortKey1 = w_Index1
        If m_SortKey1 = M_StaffID Then
            '�E���ԍ��̏ꍇ�̓C���f�b�N�X��8�Ƃ���(�E���ԍ�)
            m_SortKey1 = M_SyokuinNo
        End If

		m_SortKey2 = w_Index2
        If m_SortKey2 = M_StaffID Then
            '�E���ԍ��̏ꍇ�̓C���f�b�N�X��8�Ƃ���(�E���ԍ�)
            m_SortKey2 = M_SyokuinNo
        End If

		m_SortKey3 = w_Index3
        If m_SortKey3 = M_StaffID Then
            '�E���ԍ��̏ꍇ�̓C���f�b�N�X��8�Ƃ���(�E���ԍ�)
            m_SortKey3 = M_SyokuinNo
        End If

        m_SortOrder1 = _optSort_0.Checked

        m_SortOrder2 = _optSort_2.Checked

        m_SortOrder3 = _optSort_4.Checked

		'�\�[�g���s
		RaiseEvent Sort()
		
		'��\��
		Me.Hide()
		m_FormShowFlg = False

		Exit Sub
cmdOK_Click: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
		End
	End Sub
	
    Public Sub frmNSK0000HG_Load()
        On Error GoTo Form_Load
        Const W_SUBNAME As String = "NSK0000HG Form_Load"

        Call subSetCtlList()

        '�ŏ�ʂɐݒ�
        Call General.paSetDialogPos(Me)

        '�����ޯ���ݒ�
        Call Set_ComboBox()

        '��ʂ̒����\��
        Me.StartPosition = FormStartPosition.CenterScreen

        Exit Sub
Form_Load:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub
	
    '��đΏۍ��ڂ̐ݒ�
	Private Sub Set_ComboBox()
		On Error GoTo Set_ComboBox
		Const W_SUBNAME As String = "NSK0000HG Set_ComboBox"
		
		Dim w_Int As Short
		Dim w_Int2 As Short
		Dim w_SortKey(7) As String

        w_SortKey(M_StaffID) = "�E���ԍ�"
		w_SortKey(M_SaiyoYMD) = "�̗p�N����"
		w_SortKey(M_JobHyojiNo) = "�E��"
		w_SortKey(M_PostHyojiNo) = "��E"
		w_SortKey(M_GiryoHyojiNo) = "�Z��"
		w_SortKey(M_StaffHyojiNo) = "�\��No."
		w_SortKey(M_Team) = "�`�[��"
		
		'--- �����ޯ���̐ݒ� ---
        m_lstCmbSortKey(0).Items.Clear()
		For w_Int2 = 1 To UBound(w_SortKey)
            m_lstCmbSortKey(0).Items.Add(w_SortKey(w_Int2))
		Next w_Int2
		
		For w_Int = 1 To 2
            m_lstCmbSortKey(w_Int).Items.Clear()
            m_lstCmbSortKey(w_Int).Items.Add("�Ȃ�")
			
			For w_Int2 = 1 To UBound(w_SortKey)
				
                m_lstCmbSortKey(w_Int).Items.Add(w_SortKey(w_Int2))
			Next w_Int2
        Next w_Int
		
		'��̫�Ēl�ݒ�
        m_lstCmbSortKey(0).SelectedIndex = M_StaffHyojiNo - 1
        m_lstCmbSortKey(1).SelectedIndex = 0
        m_lstCmbSortKey(2).SelectedIndex = 0
		
        m_SortKey1 = m_lstCmbSortKey(0).SelectedIndex
		
		Exit Sub
Set_ComboBox: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub frmNSK0000HG_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo Form_Unload
        Const W_SUBNAME As String = "NSK0000HG Form_Unload"

        '���ёւ���ʔ�\��
        Me.Hide()
        m_FormShowFlg = False

        Exit Sub
Form_Unload:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub subSetCtlList()
        m_lstCmbSortKey.Add(_cmbSortKey_0)
        m_lstCmbSortKey.Add(_cmbSortKey_1)
        m_lstCmbSortKey.Add(_cmbSortKey_2)
    End Sub
End Class