Option Strict Off
Option Explicit On
Friend Class frmNSK0000HP
    Inherits General.FormBase

    Private m_OKButtonFlg As Boolean
    Private m_Comment As String
    Private m_ComFlg As Boolean
    Private m_selRiyuKbn As String
	
    '�R�����g�����׸�
    Public ReadOnly Property pComFlg() As String
        Get
            pComFlg = CStr(m_ComFlg)
        End Get
    End Property
	
	'���݉����׸�
	Public ReadOnly Property pOKFlg() As Boolean
		Get
			pOKFlg = m_OKButtonFlg
		End Get
	End Property
	
    '���͂��ꂽ�R�����g
    Public ReadOnly Property pComment() As String
        Get
            pComment = m_Comment
        End Get
    End Property

    Public WriteOnly Property pRiyuKbn() As String
        Set(ByVal Value As String)
            '���R�敪
            m_selRiyuKbn = Value
        End Set
    End Property

    Public WriteOnly Property p_com() As String
        Set(ByVal Value As String)
            '���R�敪
            m_Comment = Value
        End Set
    End Property

    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        m_OKButtonFlg = False
        Me.Close()
    End Sub

    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        On Error GoTo cmdOK_Click
        Const W_SUBNAME As String = "cmdOK_Click"

        Dim w_strMsg() As String
        Dim w_MsgResult As MsgBoxResult

        '���͓��e���i�[
        m_Comment = Comtxt.Text
        'OK���݉���
        m_OKButtonFlg = True

        Me.Close()

        Exit Sub
cmdOK_Click:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub

    Public Sub frmNSK0000HP_Load()
        On Error GoTo Form_Load
        Const W_SUBNAME As String = "NSK0000HP Form_Load"

        'OK���ݔ��f�׸ޏ�����
        m_OKButtonFlg = False

        '���ɃR�����g�������Ă��邩�m�F
        If m_Comment <> "" Then
            '�R�����g�����݂���ꍇ�͕ύX���[�h
            Comtxt.Text = m_Comment
            Comtxt.HighlightText = True
        Else
            '�R�����g�����݂��Ȃ��ꍇ��
            '�e�L�X�g������
            Comtxt.Text = ""
        End If

        '��]�̏ꍇ
        lblComment.Text = "��]�R�����g����"

        Exit Sub
Form_Load:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub
End Class