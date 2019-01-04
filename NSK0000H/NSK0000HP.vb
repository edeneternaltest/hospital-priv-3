Option Strict Off
Option Explicit On
Friend Class frmNSK0000HP
    Inherits General.FormBase

    Private m_OKButtonFlg As Boolean
    Private m_Comment As String
    Private m_ComFlg As Boolean
    Private m_selRiyuKbn As String
	
    'コメント完了ﾌﾗｸﾞ
    Public ReadOnly Property pComFlg() As String
        Get
            pComFlg = CStr(m_ComFlg)
        End Get
    End Property
	
	'ﾎﾞﾀﾝ押下ﾌﾗｸﾞ
	Public ReadOnly Property pOKFlg() As Boolean
		Get
			pOKFlg = m_OKButtonFlg
		End Get
	End Property
	
    '入力されたコメント
    Public ReadOnly Property pComment() As String
        Get
            pComment = m_Comment
        End Get
    End Property

    Public WriteOnly Property pRiyuKbn() As String
        Set(ByVal Value As String)
            '理由区分
            m_selRiyuKbn = Value
        End Set
    End Property

    Public WriteOnly Property p_com() As String
        Set(ByVal Value As String)
            '理由区分
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

        '入力内容を格納
        m_Comment = Comtxt.Text
        'OKﾎﾞﾀﾝ押下
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

        'OKﾎﾞﾀﾝ判断ﾌﾗｸﾞ初期化
        m_OKButtonFlg = False

        '既にコメントが入っているか確認
        If m_Comment <> "" Then
            'コメントが存在する場合は変更モード
            Comtxt.Text = m_Comment
            Comtxt.HighlightText = True
        Else
            'コメントが存在しない場合は
            'テキスト初期化
            Comtxt.Text = ""
        End If

        '希望の場合
        lblComment.Text = "希望コメント入力"

        Exit Sub
Form_Load:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub
End Class