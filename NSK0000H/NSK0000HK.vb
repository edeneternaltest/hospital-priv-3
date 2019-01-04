Option Strict Off
Option Explicit On
Friend Class frmNSK0000HK
    Inherits General.FormBase
	
	Private m_KangoM() As KangoM_Type
	Private m_KangoCDFlg As Boolean
	Private m_SelKangoTCD As String '選択された勤務部署CD
	Private m_OKButtonFlg As Boolean
	
	Private Structure KangoM_Type
		Dim CD As String
        Dim Name_Renamed As String
    End Structure

	'勤務部署CD取得ﾌﾗｸﾞ
	Public ReadOnly Property pKangoCDFlg() As String
		Get
			pKangoCDFlg = CStr(m_KangoCDFlg)
		End Get
	End Property
	
	'ﾎﾞﾀﾝ押下ﾌﾗｸﾞ
	Public ReadOnly Property pOKFlg() As Boolean
		Get
			pOKFlg = m_OKButtonFlg
		End Get
	End Property
	
	'選択された勤務部署CD取得ﾌﾗｸﾞ
	Public ReadOnly Property pSelKangoTCD() As String
		Get
			pSelKangoTCD = m_SelKangoTCD
		End Get
	End Property
	
    Private Sub Get_KangoM()
        On Error GoTo Get_KangoM
        Const W_SUBNAME As String = "NSK0000HK Get_KangoM"

        Dim w_Int As Integer
        Dim w_DataCnt As Integer
        Dim w_Index As Integer
        Dim w_CD As String
        Dim w_YYYYMMDD As Integer

        w_YYYYMMDD = Integer.Parse(Format(Now, "yyyyMMdd"))

        '配列初期化
        ReDim m_KangoM(0)

        With General.g_objGetMaster
            .pHospitalCD = General.g_strHospitalCD '施設コード
            .pKD_KinmuDeptCD = "" '空白(全件)
            .pKD_KijunDate = w_YYYYMMDD '基準日

            If .mGetKinmuDept = False Then
                'ﾃﾞｰﾀがないとき
            Else
                '勤務部署件数取得
                w_DataCnt = .fKD_KinmuDeptCount

                For w_Int = 1 To w_DataCnt
                    'ｲﾝﾃﾞｯｸｽ引渡し
                    .mKD_KinmuDeptIdx = w_Int
                    '勤務部署CD取得
                    w_CD = .fKD_KinmuDeptCD

                    '自部署勤務部署CD以外のときは配列に格納
                    If w_CD <> General.g_strSelKinmuDeptCD Then
                        '配列拡張
                        w_Index = UBound(m_KangoM) + 1
                        ReDim Preserve m_KangoM(w_Index)

                        '●勤務部署CD
                        m_KangoM(w_Index).CD = w_CD
                        '●名称取得
                        m_KangoM(w_Index).Name_Renamed = .fKD_KinmuDeptName
                    End If
                Next w_Int
            End If
        End With

        Exit Sub
Get_KangoM:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub
	
	Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
		m_OKButtonFlg = False
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		On Error GoTo cmdOK_Click
		Const W_SUBNAME As String = "cmdOK_Click"
		
		Dim w_SelectIndex As Integer
		Dim w_StrReg As String
		
        '選択された勤務地のｲﾝﾃﾞｯｸｽ取得
		w_SelectIndex = cboKangoTani.SelectedIndex + 1
		
		m_SelKangoTCD = m_KangoM(w_SelectIndex).CD '選択された勤務部署CD

		'----- ﾚｼﾞｽﾄﾘ設定 ------------------
        w_StrReg = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY2
		'応援先勤務地CD
        Call General.paSaveSetting(w_StrReg, "NSK0000H", "OUENKINMUDEPTCD_" & General.g_strUserMngID, m_SelKangoTCD)
		
        'OKﾎﾞﾀﾝ押下
		m_OKButtonFlg = True
		
		Me.Close()
		
		Exit Sub
cmdOK_Click: 
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
		End
	End Sub
	
    Public Sub frmNSK0000HK_Load()
        On Error GoTo Form_Load
        Const W_SUBNAME As String = "NSK0000HK Form_Load"

        Dim w_Int As Integer
        Dim w_StrReg As String
        Dim w_strOuenKinmuDeptCd As String
        Dim w_lngOuenIndex As Integer

        '勤務部署ﾏｽﾀ取得ﾌﾗｸﾞ初期化
        m_KangoCDFlg = False
        'OKﾎﾞﾀﾝ判断ﾌﾗｸﾞ初期化
        m_OKButtonFlg = False

        '初期化
        m_SelKangoTCD = ""

        '勤務部署取得
        Call Get_KangoM()

        '勤務部署Ｍにﾃﾞｰﾀがなかった場合
        If UBound(m_KangoM) <= 0 Then
            Exit Sub
        End If

        'ｺﾝﾎﾞﾎﾞｯｸｽにｾｯﾄ
        cboKangoTani.Items.Clear()
        For w_Int = 1 To UBound(m_KangoM)
            cboKangoTani.Items.Add(m_KangoM(w_Int).Name_Renamed)
        Next w_Int

        '--- ﾚｼﾞｽﾄﾘ設定値 取得 ---
        '応援先勤務地CD
        w_StrReg = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY2

        w_strOuenKinmuDeptCd = General.paGetSetting(w_StrReg, "NSK0000H", "OUENKINMUDEPTCD_" & General.g_strUserMngID, "")

        w_lngOuenIndex = 0

        If w_strOuenKinmuDeptCd = "" Then
            cboKangoTani.SelectedIndex = 0
        Else
            For w_Int = 1 To UBound(m_KangoM)
                If w_strOuenKinmuDeptCd = m_KangoM(w_Int).CD Then
                    '応援先勤務地CDが一致した行を選択
                    w_lngOuenIndex = w_Int - 1
                End If
            Next w_Int

            cboKangoTani.SelectedIndex = w_lngOuenIndex
        End If

        'ﾃﾞｰﾀ取得OK
        m_KangoCDFlg = True

        Exit Sub
Form_Load:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub
End Class