Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Friend Class frmNSK0000HG
    Inherits General.FormBase
	'=====================================================================================
	'   定  数  宣  言
	'=====================================================================================
	
	'職員情報項目 定数
	Private Const M_StaffID As Short = 1 '職員管理番号
	Private Const M_SaiyoYMD As Short = 2 '採用年月日
	Private Const M_JobHyojiNo As Short = 3 '職種表示No
	Private Const M_PostHyojiNo As Short = 4 '役職表示No
	Private Const M_GiryoHyojiNo As Short = 5 '技量表示No
	Private Const M_StaffHyojiNo As Short = 6 '職員表示No
	Private Const M_Team As Short = 7 'ﾁｰﾑ
    Private Const M_SyokuinNo As Short = 8 '職員番号
	Private Const M_SortTeam As Short = 9 'ソート用
    Private Const M_HairetuNo As Short = 10 'ﾃﾞｰﾀ格納配列番号
	
	'-----------------------------------------------------------------
	'   変 数 宣 言
	'-----------------------------------------------------------------
    Private m_FormShowFlg As Boolean '画面が表示されているか
	Private m_SortKey1 As Short '1番目のソートキー
	Private m_SortKey2 As Short '2番目のソートキー
	Private m_SortKey3 As Short '3番目のソートキー
    Private m_SortOrder1 As Boolean '1番目のソートキーの並び順
    Private m_SortOrder2 As Boolean '2番目のソートキーの並び順
    Private m_SortOrder3 As Boolean '3番目のソートキーの並び順
    Private m_lstCmbSortKey As New List(Of Object)

	'------------------------------------------------------------------
	'  ｲﾍﾞﾝﾄ宣言
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
        '非表示
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

		'選択キーのｲﾝﾃﾞｯｸｽ取得
        w_Index1 = m_lstCmbSortKey(0).SelectedIndex + 1
        w_Index2 = m_lstCmbSortKey(1).SelectedIndex
        w_Index3 = m_lstCmbSortKey(2).SelectedIndex
		
		'キーが選択されているか
		If w_Index1 = 0 Then
			'１番目のキーが未選択のとき(なしのとき)
			ReDim w_strMsg(3)
			w_strMsg(1) = "１番目"
			w_strMsg(2) = "ソートキー"
			w_strMsg(3) = "キー"
			RaiseEvent MsgBox_Renamed("NS0100", w_strMsg)
			
			Exit Sub
		End If
		
		If w_Index2 = 0 And w_Index3 <> 0 Then
			'２番目のキーが未選択で３番目のキーが選択されているとき
			ReDim w_strMsg(3)
			w_strMsg(1) = "２番目"
			w_strMsg(2) = "ソートキー"
			w_strMsg(3) = "キー"
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
		
		'選択キーのｲﾝﾃﾞｯｸｽ番号を格納
		m_SortKey1 = w_Index1
        If m_SortKey1 = M_StaffID Then
            '職員番号の場合はインデックスを8とする(職員番号)
            m_SortKey1 = M_SyokuinNo
        End If

		m_SortKey2 = w_Index2
        If m_SortKey2 = M_StaffID Then
            '職員番号の場合はインデックスを8とする(職員番号)
            m_SortKey2 = M_SyokuinNo
        End If

		m_SortKey3 = w_Index3
        If m_SortKey3 = M_StaffID Then
            '職員番号の場合はインデックスを8とする(職員番号)
            m_SortKey3 = M_SyokuinNo
        End If

        m_SortOrder1 = _optSort_0.Checked

        m_SortOrder2 = _optSort_2.Checked

        m_SortOrder3 = _optSort_4.Checked

		'ソート実行
		RaiseEvent Sort()
		
		'非表示
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

        '最上位に設定
        Call General.paSetDialogPos(Me)

        'ｺﾝﾎﾞﾎﾞｯｸｽ設定
        Call Set_ComboBox()

        '画面の中央表示
        Me.StartPosition = FormStartPosition.CenterScreen

        Exit Sub
Form_Load:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub
	
    'ｿｰﾄ対象項目の設定
	Private Sub Set_ComboBox()
		On Error GoTo Set_ComboBox
		Const W_SUBNAME As String = "NSK0000HG Set_ComboBox"
		
		Dim w_Int As Short
		Dim w_Int2 As Short
		Dim w_SortKey(7) As String

        w_SortKey(M_StaffID) = "職員番号"
		w_SortKey(M_SaiyoYMD) = "採用年月日"
		w_SortKey(M_JobHyojiNo) = "職種"
		w_SortKey(M_PostHyojiNo) = "役職"
		w_SortKey(M_GiryoHyojiNo) = "技量"
		w_SortKey(M_StaffHyojiNo) = "表示No."
		w_SortKey(M_Team) = "チーム"
		
		'--- ｺﾝﾎﾞﾎﾞｯｸｽの設定 ---
        m_lstCmbSortKey(0).Items.Clear()
		For w_Int2 = 1 To UBound(w_SortKey)
            m_lstCmbSortKey(0).Items.Add(w_SortKey(w_Int2))
		Next w_Int2
		
		For w_Int = 1 To 2
            m_lstCmbSortKey(w_Int).Items.Clear()
            m_lstCmbSortKey(w_Int).Items.Add("なし")
			
			For w_Int2 = 1 To UBound(w_SortKey)
				
                m_lstCmbSortKey(w_Int).Items.Add(w_SortKey(w_Int2))
			Next w_Int2
        Next w_Int
		
		'ﾃﾞﾌｫﾙﾄ値設定
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

        '並び替え画面非表示
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