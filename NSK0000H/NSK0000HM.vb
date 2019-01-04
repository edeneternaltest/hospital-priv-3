Option Strict Off
Option Explicit On
Friend Class frmNSK0000HM
    Inherits General.FormBase
	
	Private m_CountMax As Single
    Private m_FormatCnt As Short

    '2016/04/05 Ishiga add start-------------------------
    Public WriteOnly Property pNumberDisp() As Boolean
        Set(ByVal value As Boolean)
            lblCount.Visible = value
        End Set
    End Property
    '2016/04/05 Ishiga add end---------------------------

	Public WriteOnly Property pCountValue() As Short
		Set(ByVal Value As Short)
			
			'Max値が設定されていない場合は抜け出し
			If m_CountMax <= 0 Then
				Exit Property
			End If
			
			'処理経過ﾗﾍﾞﾙ･ﾌﾟﾛｸﾞﾚｽﾊﾞｰ設定
            lblCount.Text = String.Format("{0, " & m_FormatCnt & "}", Value) & " / " & CStr(m_CountMax)
			If Value <= 0 Then
				pgbProcess.Value = 0
			ElseIf Value >= m_CountMax Then 
				pgbProcess.Value = m_CountMax
			Else
				pgbProcess.Value = Value
			End If
			lblCount.Refresh()
			
		End Set
	End Property
	
	Public WriteOnly Property pForeColor() As Integer
		Set(ByVal Value As Integer)
            lblSyori.ForeColor = ColorTranslator.FromOle(Value)
		End Set
	End Property
	
	Public WriteOnly Property pMaxValue() As Short
		Set(ByVal Value As Short)
			
			'ｶｳﾝﾄﾗﾍﾞﾙ、ｶｳﾝﾄﾌﾟﾛｸﾞﾚｽﾊﾞｰを初期化する
			m_CountMax = CSng(Value)
			If m_CountMax >= 10000 Then
                m_FormatCnt = 5
			ElseIf m_CountMax >= 1000 Then 
                m_FormatCnt = 4
			ElseIf m_CountMax >= 100 Then 
                m_FormatCnt = 3
			ElseIf m_CountMax >= 10 Then 
                m_FormatCnt = 2
			Else
                m_FormatCnt = 1
			End If
            lblCount.Text = String.Format("{0, " & m_FormatCnt & "}", 0) & " / " & CStr(m_CountMax)
			
			If m_CountMax > 0 Then
				pgbProcess.Maximum = m_CountMax
			End If
			pgbProcess.Value = 0
			
			lblCount.Refresh()
			
		End Set
	End Property
	
	
	Public WriteOnly Property pSyoriText() As String
		Set(ByVal Value As String)
			lblSyori.Text = Value
			lblSyori.Refresh()
		End Set
	End Property
	
	
	Private Sub frmNSK0000HM_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo Form_Load
		Const W_SUBNAME As String = "Nskk001h Form_Load"
		
		Dim w_Left As Single
		
		'処理表示ﾗﾍﾞﾙ
        lblSyori.SetBounds(0, General.paTwipsTopixels(240), Me.ClientRectangle.Width, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)

        '処理計画表示ﾗﾍﾞﾙ
        lblCount.SetBounds(0, General.paTwipsTopixels(840), Me.ClientRectangle.Width, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)

        '処理経過ﾌﾟﾛｸﾞﾚｽﾊﾞｰ
        w_Left = (Me.ClientRectangle.Width - pgbProcess.Width) / 2
        pgbProcess.SetBounds(w_Left, General.paTwipsTopixels(1200), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		
		'各ラベルの初期化
		lblSyori.Text = ""
		lblCount.Text = ""
		
        Me.StartPosition = FormStartPosition.CenterScreen
		
		Exit Sub
Form_Load: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
		End
	End Sub
End Class