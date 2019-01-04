Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Friend Class frmNSK0000HJ
    Inherits General.FormBase
    Public NSK0000H_sql As New clsNSK0000H_sql

	Private Structure Daikyu
		Dim lngYMD As Integer
		Dim strKinmuCD As String '休日出勤勤務CD
		Dim strKinmuNM As String '休日出勤勤務名称
		Dim strKinmuMark As String '休日出勤勤務記号
        Dim strDaikyuValueType As String '(0:1日、1:1.5日、2:0.5日)
    End Structure

	Private m_udtDaikyu() As Daikyu
    Private m_blnEndStatus As Boolean '終了状態

	'*** ﾌﾟﾛﾊﾟﾃｨ受け取り
	Private m_strSelDate As String '指定されている年月日
	Private m_strSelKinmuCD As String '選択された勤務CD
	Private m_strMngStaffID As String '職員管理番号
	Private m_KeikakuFlg As String '起動元判別ﾌﾗｸﾞ("0":計画画面 それ以外:初期画面)
	Private m_Index As Integer
	Private m_SelDate As Integer '選択された日付
    Private m_HalfDaikyuList_bk() As String
	Private m_SelDate2 As Integer '選択された日付
	Private m_HalfKinmuFlg As Boolean '半日代休対象勤務Flg
	Private m_ClearFlg As Boolean
    Private m_GetDaikyuType As Object
    Private m_lstCmbHalfDaikyuList As New List(Of Object)
    Private m_lstOptDaikyuType As New List(Of Object)

	Private Structure DaikyuDetailType
		Dim lngYMD As Integer '取得年月日
		Dim strKinmuCD As String '勤務CD
		Dim strGetDaikyuType As String '取得タイプ(0:1日、1:0.5日)
		Dim dblRegistFirstTimeDate As Double
	End Structure

	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		
		m_blnEndStatus = False
		Me.Close()
		
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		On Error GoTo Errhnadler
        Const W_SUBNAME As String = "NSK0000HJ cmdOK_Click"

		Dim w_Index As Integer
        Dim w_str As String
		Dim w_lngLoop As Integer
		Dim w_strMsg() As String

		m_blnEndStatus = True
		
        'データを取得
		m_SelDate = 0
		m_SelDate2 = 0
        If m_HalfKinmuFlg = True Or m_lstOptDaikyuType(0).Checked = True Then
            '半日または１日代休の場合
            w_str = Format(CDate(General.paLeft(cmbDaikyuList.Text, 11)), "yyyyMMdd")
            For w_lngLoop = 1 To UBound(m_udtDaikyu)
                If CDbl(w_str) = m_udtDaikyu(w_lngLoop).lngYMD Then
                    m_SelDate = m_udtDaikyu(w_lngLoop).lngYMD
                    Exit For
                End If
            Next w_lngLoop

            If m_KeikakuFlg <> "0" Then
                '計画画面以外から呼ばれた場合、代休取得年月日を更新する。
                Call fncSetDaikyuGetDate(m_SelDate)
            End If
        Else
            '半日＋半日代休の場合
            w_str = Format(CDate(General.paLeft(m_lstCmbHalfDaikyuList(0).Text, 11)), "yyyyMMdd")
            For w_lngLoop = 1 To UBound(m_udtDaikyu)
                If CDbl(w_str) = m_udtDaikyu(w_lngLoop).lngYMD Then
                    m_SelDate = m_udtDaikyu(w_lngLoop).lngYMD
                    Exit For
                End If
            Next w_lngLoop

            'ｴﾗｰﾁｪｯｸ
            w_str = Format(CDate(General.paLeft(m_lstCmbHalfDaikyuList(1).Text, 11)), "yyyyMMdd")

            '同じ日を選択している場合ＮＧ
            If w_str = "" Then
                '*******ﾒｯｾｰｼﾞ***********************************
                ReDim w_strMsg(1)
                w_strMsg(1) = "対象代休"
                Call General.paMsgDsp("NS0001", w_strMsg)
                '************************************************
                Call General.paSetFocus(m_lstCmbHalfDaikyuList(1))
                Exit Sub
            End If

            For w_lngLoop = 1 To UBound(m_udtDaikyu)
                If CDbl(w_str) = m_udtDaikyu(w_lngLoop).lngYMD Then
                    m_SelDate2 = m_udtDaikyu(w_lngLoop).lngYMD
                    Exit For
                End If
            Next w_lngLoop

            '同じ日を選択している場合ＮＧ
            If m_SelDate = m_SelDate2 Then
                '*******ﾒｯｾｰｼﾞ***********************************
                ReDim w_strMsg(1)
                w_strMsg(1) = "日付"
                Call General.paMsgDsp("NS0003", w_strMsg)
                '************************************************
                Call General.paSetFocus(m_lstCmbHalfDaikyuList(0))
                Exit Sub
            End If

            If m_KeikakuFlg <> "0" Then
                '計画画面以外から呼ばれた場合、代休取得年月日を更新する。
                Call fncSetDaikyuGetDate(m_SelDate)
                Call fncSetDaikyuGetDate(m_SelDate2)
            End If
        End If

		Me.Close()
		
		Exit Sub
Errhnadler: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
	End Sub
	
	'終了状態を引渡す
	Public ReadOnly Property pEndStatus() As Boolean
		Get
			pEndStatus = m_blnEndStatus
		End Get
	End Property
	
	'選択された日付を引渡す
    '選択年月日を受け取る
	Public Property pSelDate() As Integer
		Get
			pSelDate = m_SelDate
        End Get

		Set(ByVal Value As Integer)
			m_strSelDate = CStr(Value)
		End Set
	End Property
	
	'代休情報を受け取る
    Public WriteOnly Property pDaikyuData(ByVal p_HolDate As Integer, ByVal p_HolKinmuCD As String) As Double
        Set(ByVal Value As Double)

            Dim w_str As String

            '配列拡張
            m_Index = UBound(m_udtDaikyu) + 1
            ReDim Preserve m_udtDaikyu(m_Index)

            '休日出勤日
            m_udtDaikyu(m_Index).lngYMD = p_HolDate

            '休日出勤勤務CD
            m_udtDaikyu(m_Index).strKinmuCD = p_HolKinmuCD

            '代休タイプ
            w_str = ""

            Select Case Value
                Case 1
                    w_str = "0"
                Case 1.5
                    w_str = "1"
                Case 0.5
                    w_str = "2"
            End Select

            m_udtDaikyu(m_Index).strDaikyuValueType = w_str
        End Set
    End Property
	
	'起動元（計画画面からかどうか）を受け取る ("0": 計画画面から　それ以外:初期画面から)
	Public WriteOnly Property pKeikakuFlg() As String
		Set(ByVal Value As String)
			m_KeikakuFlg = Value
			
			'代休配列初期化
			ReDim m_udtDaikyu(0)
			m_Index = 0
        End Set
	End Property

	'選択勤務CDを受け取る
	Public WriteOnly Property pSelKinmuCD() As String
		Set(ByVal Value As String)
			m_strSelKinmuCD = Value
		End Set
	End Property
	
	'職員管理番号を受け取る
	Public WriteOnly Property pSelMngStaffID() As String
		Set(ByVal Value As String)
			m_strMngStaffID = Value
		End Set
	End Property
	
	'選択された日付を引渡す
	Public ReadOnly Property pSelDate2() As Integer
		Get
			pSelDate2 = m_SelDate2
		End Get
	End Property
	
	'半日代休取得勤務か引渡す
	Public ReadOnly Property pGetDaikyuType() As Boolean
		Get
			pGetDaikyuType = m_HalfKinmuFlg
		End Get
	End Property
	
	Private Sub frmNSK0000HJ_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo ErrHandler
        Const W_SUBNAME As String = "NSK0000HJ Form_Load"

		Dim w_lngLoop As Integer
		Dim w_strString As String

        Call subSetCtlList()

		'フォームを設定
		Call SetForm()

        Me.StartPosition = FormStartPosition.CenterScreen
		
		Exit Sub
ErrHandler: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
		End
	End Sub
	
	'初期画面用
	Public Function mfncDaikyuDate_Check() As Boolean
		On Error GoTo ErrHandler
        Const W_SUBNAME As String = "NSK0000HJ mfncDaikyuDate_Check"

		Dim w_strSql As String
		Dim w_Rs As ADODB.Recordset
		Dim w_lngDaikyuPastPeriod As Integer '過去の代休取得時の休日出勤日の有効範囲（何日前までの休日出勤は有効って感じ）
		Dim w_lngDaikyuDate As Integer
		Dim w_objDic As Object
		Dim w_lngKensu As Integer
		Dim w_lngLoop As Integer 'ﾙｰﾌﾟｶｳﾝﾄ
		Dim w_勤務CD_F As ADODB.Field
		Dim w_名称_F As ADODB.Field
		Dim w_記号_F As ADODB.Field
		Dim w_休日出勤年月日_F As ADODB.Field
		Dim w_休日出勤勤務CD_F As ADODB.Field
		Dim w_DaikyuAdvFlg As Integer
		Dim w_lngDaikyuDate_To As Integer
        Dim w_varWork As Object
		Dim w_strString As String
		Dim w_lngCount_Day As Integer
		Dim w_lngCount_HalfDay As Integer
		Dim w_DaikyuAdvThisMonthFlg As Integer
        Dim w_lngDaikyuDate_To2 As Integer
		
		'代休の有効期間を求める(ﾃﾞﾌｫﾙﾄは８週間)
        w_lngDaikyuPastPeriod = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "PASTDAIKYUPERIOD", "56", General.g_strHospitalCD))
        '代休先取りフラグ
        w_DaikyuAdvFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "DAIKYUADVANCEFLG", CStr(0), General.g_strHospitalCD))

        '代休先取り当月フラグ
        w_DaikyuAdvThisMonthFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "DAIKYUADVANCETHISMONTHFLG", CStr(0), General.g_strHospitalCD))

        '代休取得期間に制限をかける？-----
        If w_lngDaikyuPastPeriod = -1 Then '制限なし
            If w_DaikyuAdvFlg = 0 Then ''先取りなし

                '代休データ取得期間(計画期間の開始日から有効期間数過去の日付)
                w_lngDaikyuDate = 0
                '代休データ取得期間
                w_lngDaikyuDate_To = Integer.Parse(m_strSelDate)
            Else ''先取りあり
                '代休データ取得期間(計画期間の開始日から有効期間数過去の日付)
                w_lngDaikyuDate = 0

                If w_DaikyuAdvThisMonthFlg = 0 Then
                    w_lngDaikyuDate_To = 99999999
                Else
                    w_lngDaikyuDate_To = Integer.Parse(Format(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(Integer.Parse(General.paLeft(m_strSelDate, 6) & "01"), "0000/00/00")))), "yyyyMMdd"))
                End If
            End If
        Else '制限あり
            If w_DaikyuAdvFlg = 0 Then ''先取りなし

                '代休データ取得期間(計画期間の開始日から有効期間数過去の日付)
                w_lngDaikyuDate = Integer.Parse(Format(DateAdd(DateInterval.Day, w_lngDaikyuPastPeriod * -1, CDate(Format(Integer.Parse(m_strSelDate), "0000/00/00"))), "yyyyMMdd"))
                '代休データ取得期間
                w_lngDaikyuDate_To = Integer.Parse(m_strSelDate)
            Else ''先取りあり
                '代休データ取得期間(計画期間の開始日から有効期間数過去の日付)
                w_lngDaikyuDate = Integer.Parse(Format(DateAdd(DateInterval.Day, w_lngDaikyuPastPeriod * -1, CDate(Format(Integer.Parse(m_strSelDate), "0000/00/00"))), "yyyyMMdd"))

                '代休データ取得期間
                w_lngDaikyuDate_To = Integer.Parse(Format(DateAdd(DateInterval.Day, w_lngDaikyuPastPeriod * 1, CDate(Format(Integer.Parse(m_strSelDate), "0000/00/00"))), "yyyyMMdd"))
                w_lngDaikyuDate_To2 = Integer.Parse(Format(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(Integer.Parse(General.paLeft(m_strSelDate, 6) & "01"), "0000/00/00")))), "yyyyMMdd"))
                If w_DaikyuAdvThisMonthFlg <> 0 And w_lngDaikyuDate_To > w_lngDaikyuDate_To2 Then
                    w_lngDaikyuDate_To = w_lngDaikyuDate_To2
                End If
            End If
        End If

        w_objDic = CreateObject("scripting.dictionary")
        '2017/05/02 Christopher Upd Start
        'w_strSql = ""
        'w_strSql = "select KinmuCD, Name, MarkF from NS_KINMUNAME_M"
        'w_strSql = w_strSql & " where GetDaikyuFlg = '1'"
        'w_strSql = w_strSql & " and HospitalCD = '" & General.g_strHospitalCD & "'"

        'w_Rs = General.paDBRecordSetOpen(w_strSql)

        Call NSK0000H_sql.select_NS_KINMUNAME_M_04(w_Rs)
        'Upd End
        With w_Rs
            If .RecordCount > 0 Then
                .MoveLast()
                w_lngKensu = .RecordCount
                .MoveFirst()

                w_勤務CD_F = .Fields("KinmuCD")
                w_名称_F = .Fields("Name")
                w_記号_F = .Fields("MarkF")
                For w_lngLoop = 1 To w_lngKensu
                    w_objDic.Item(CStr(w_勤務CD_F.Value & "A")) = CStr(w_名称_F.Value & "")
                    w_objDic.Item(CStr(w_勤務CD_F.Value & "B")) = CStr(w_記号_F.Value & "")
                    .MoveNext()
                Next w_lngLoop
            End If
            .Close()
        End With

        w_Rs = Nothing

        '初期画面から呼ばれた場合
        If m_KeikakuFlg <> "0" Then

            '代休管理Ｆより取得可能な代休日付を取得する。
            w_strSql = ""
            w_strSql = "select WorkHolKinmuDate, WorkHolKinmuCD from NS_DAIKYUMNG_F"
            w_strSql = w_strSql & " where (GetDaikyuDate is null"
            w_strSql = w_strSql & " or GetDaikyuDate = 0)"
            w_strSql = w_strSql & " and WorkHolKinmuDate >= " & w_lngDaikyuDate
            w_strSql = w_strSql & " and WorkHolKinmuDate <= " & Integer.Parse(m_strSelDate)
            w_strSql = w_strSql & " and StaffMngID = '" & Trim(m_strMngStaffID) & "'"
            w_strSql = w_strSql & " and HospitalCD = '" & Trim(General.g_strHospitalCD) & "'"

            w_Rs = General.paDBRecordSetOpen(w_strSql)

            With w_Rs
                If .RecordCount <= 0 Then
                    mfncDaikyuDate_Check = False
                Else
                    .MoveLast()
                    w_lngKensu = .RecordCount
                    .MoveFirst()

                    w_休日出勤年月日_F = .Fields("WorkHolKinmuDate")
                    w_休日出勤勤務CD_F = .Fields("WorkHolKinmuCD")
                    ReDim m_udtDaikyu(w_lngKensu)
                    For w_lngLoop = 1 To w_lngKensu
                        m_udtDaikyu(w_lngLoop).lngYMD = Integer.Parse(w_休日出勤年月日_F.Value)
                        m_udtDaikyu(w_lngLoop).strKinmuCD = CStr(w_休日出勤勤務CD_F.Value & "")
                        m_udtDaikyu(w_lngLoop).strKinmuNM = w_objDic.Item(m_udtDaikyu(w_lngLoop).strKinmuCD & "A")
                        m_udtDaikyu(w_lngLoop).strKinmuMark = w_objDic.Item(m_udtDaikyu(w_lngLoop).strKinmuCD & "B")
                        .MoveNext()
                    Next w_lngLoop
                    mfncDaikyuDate_Check = True
                End If
                .Close()
            End With

            w_Rs = Nothing

        Else
            '計画画面から呼ばれた場合
            If UBound(m_udtDaikyu) > 0 Then
                mfncDaikyuDate_Check = True

                For w_lngLoop = 1 To UBound(m_udtDaikyu)
                    m_udtDaikyu(w_lngLoop).strKinmuNM = w_objDic.Item(m_udtDaikyu(w_lngLoop).strKinmuCD & "A")
                    m_udtDaikyu(w_lngLoop).strKinmuMark = w_objDic.Item(m_udtDaikyu(w_lngLoop).strKinmuCD & "B")
                Next w_lngLoop
            Else
                mfncDaikyuDate_Check = False
            End If
        End If

        '取得可能代休があるかチェック
        If mfncDaikyuDate_Check = True Then
            '半日代休取得可能勤務ＣＤ
            m_HalfKinmuFlg = False

            If g_KinmuM(CShort(m_strSelKinmuCD)).AMCD <> "" And g_KinmuM(CShort(m_strSelKinmuCD)).PMCD <> "" Then
                If g_KinmuM(CShort(g_KinmuM(CShort(m_strSelKinmuCD)).AMCD)).HolBunruiCD = General.G_STRDAIKYUBUNRUI Or g_KinmuM(CShort(g_KinmuM(CShort(m_strSelKinmuCD)).PMCD)).HolBunruiCD = General.G_STRDAIKYUBUNRUI Then
                    m_HalfKinmuFlg = True
                End If
            End If

            w_lngCount_Day = 0
            w_lngCount_HalfDay = 0

            '１日代休
            For w_lngLoop = 1 To UBound(m_udtDaikyu)
                Select Case m_udtDaikyu(w_lngLoop).strDaikyuValueType
                    Case "0"
                        '１日
                        w_lngCount_Day = w_lngCount_Day + 1
                        w_lngCount_HalfDay = w_lngCount_HalfDay + 2
                    Case "1"
                        '1.5日
                        w_lngCount_Day = w_lngCount_Day + 1
                        w_lngCount_HalfDay = w_lngCount_HalfDay + 3
                    Case "2"
                        w_lngCount_HalfDay = w_lngCount_HalfDay + 1
                End Select
            Next w_lngLoop

            If m_HalfKinmuFlg = True Then
                '半日用フォーム
                If w_lngCount_HalfDay < 1 Then
                    mfncDaikyuDate_Check = False
                End If
            Else
                '通常フォーム
                If w_lngCount_Day < 1 And w_lngCount_HalfDay < 2 Then
                    mfncDaikyuDate_Check = False
                End If
            End If
        End If

        w_objDic = Nothing
        Exit Function
ErrHandler:
        w_objDic = Nothing
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Function

    Private Sub frmNSK0000HJ_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As FormClosingEventArgs) Handles Me.FormClosing
        Erase m_udtDaikyu
    End Sub

    '代休管理Ｆに代休取得日を更新する
    Private Function fncSetDaikyuGetDate(ByRef p_SelDate As Integer) As Boolean
        On Error GoTo ErrHandler
        Const W_SUBNAME As String = "NSK0000HJ fncSetDaikyuGetDate"

        Dim w_strSql As String
        Dim w_Rs As ADODB.Recordset
        Dim w_GETFLG_F As ADODB.Field
        Dim w_GETDAIKYUDATE_F As ADODB.Field
        Dim w_GETDAIKYUKINMUCD_F As ADODB.Field
        Dim w_REGISTFIRSTTIMEDATE_F As ADODB.Field
        Dim w_DetaIdx As Short
        Dim w_DataCnt As Integer
        Dim w_DataLoop As Integer
        Dim w_SysDate As Double
        Dim w_DaikyuType As String '(0:１日,1:0.5日)
        Dim w_DetailInfo() As DaikyuDetailType

        fncSetDaikyuGetDate = False

        w_SysDate = CDbl(Format(Now, "yyyyMMddHHmmss"))

        'データが既にあるかチェック
        ReDim w_DetailInfo(0)
        '2017/05/02 Christopher Upd Start
        'w_strSql = "select * from NS_DAIKYUDETAILMNG_F"
        'w_strSql = w_strSql & " where  WorkHolKinmuDate = " & p_SelDate
        'w_strSql = w_strSql & " and GETDAIKYUDATE <> " & m_strSelDate
        'w_strSql = w_strSql & " and StaffMngID = '" & Trim(m_strMngStaffID) & "'"
        'w_strSql = w_strSql & " and HospitalCD = '" & Trim(General.g_strHospitalCD) & "'"
        'w_Rs = General.paDBRecordSetOpen(w_strSql)

        Call NSK0000H_sql.select_NS_DAIKYUDETAILMNG_F_01(w_Rs, p_SelDate, m_strSelDate, Trim(m_strMngStaffID), Trim(General.g_strHospitalCD))
        'Upd End
        With w_Rs
            If .RecordCount <= 0 Then
                'ﾃﾞｰﾀなし
            Else
                'ﾃﾞｰﾀあり
                .MoveLast()
                w_DataCnt = .RecordCount
                .MoveFirst()
                '配列確保
                ReDim Preserve w_DetailInfo(w_DataCnt)

                w_GETFLG_F = .Fields("GETFLG")
                w_GETDAIKYUDATE_F = .Fields("GETDAIKYUDATE")
                w_GETDAIKYUKINMUCD_F = .Fields("GETDAIKYUKINMUCD")
                w_REGISTFIRSTTIMEDATE_F = .Fields("REGISTFIRSTTIMEDATE")

                For w_DataLoop = 1 To w_DataCnt
                    'データ格納
                    w_DetailInfo(w_DataLoop).strGetDaikyuType = IIf(IsDBNull(w_GETFLG_F.Value), "0", w_GETFLG_F.Value)
                    w_DetailInfo(w_DataLoop).lngYMD = IIf(IsDBNull(w_GETDAIKYUDATE_F.Value), 0, w_GETDAIKYUDATE_F.Value)
                    w_DetailInfo(w_DataLoop).strKinmuCD = IIf(IsDBNull(w_GETDAIKYUKINMUCD_F.Value), "", w_GETDAIKYUKINMUCD_F.Value)
                    w_DetailInfo(w_DataLoop).dblRegistFirstTimeDate = IIf(IsDBNull(w_REGISTFIRSTTIMEDATE_F.Value), 0, w_REGISTFIRSTTIMEDATE_F.Value)

                    .MoveNext()
                Next w_DataLoop
            End If
        End With
        w_Rs.Close()

        '代休の取得タイプ取得
        w_DetaIdx = UBound(w_DetailInfo) + 1
        ReDim Preserve w_DetailInfo(w_DetaIdx)

        If m_HalfKinmuFlg = True Or m_lstOptDaikyuType(1).Checked = True Then
            '半日の場合
            w_DetailInfo(w_DetaIdx).strGetDaikyuType = "1"
        Else
            '１日の場合
            w_DetailInfo(w_DetaIdx).strGetDaikyuType = "0"
        End If

        w_DetailInfo(w_DetaIdx).lngYMD = Integer.Parse(m_strSelDate)
        w_DetailInfo(w_DetaIdx).strKinmuCD = m_strSelKinmuCD
        w_DetailInfo(w_DetaIdx).dblRegistFirstTimeDate = w_SysDate

        Call General.paBeginTrans()
        '2017/05/02 Christopher Upd Start
        'データ削除
        ''Delete文 編集
        'w_strSql = "Delete From NS_DAIKYUDETAILMNG_F "
        'w_strSql = w_strSql & " where WorkHolKinmuDate = " & p_SelDate
        'w_strSql = w_strSql & " and StaffMngID = '" & Trim(m_strMngStaffID) & "'"
        'w_strSql = w_strSql & " and HospitalCD = '" & Trim(General.g_strHospitalCD) & "'"

        'Call General.paDBExecute(w_strSql)

        Call NSK0000H_sql.delete_NS_DAIKYUDETAILMNG_F_02(p_SelDate, Trim(m_strMngStaffID))
        'Upd End
        For w_DataLoop = 1 To UBound(w_DetailInfo)
            '2017/05/22 Richard Upd Start
            ''Insert文 編集
            'w_strSql = "Insert Into NS_DAIKYUDETAILMNG_F ("
            'w_strSql = w_strSql & "HospitalCD,"
            'w_strSql = w_strSql & "StaffMngID,"
            'w_strSql = w_strSql & "WorkHolKinmuDate,"
            'w_strSql = w_strSql & "SEQ,"
            'w_strSql = w_strSql & "GETFLG,"
            'w_strSql = w_strSql & "GetDaikyuDate,"
            'w_strSql = w_strSql & "GetDaikyuKinmuCD,"
            'w_strSql = w_strSql & "RegistFirstTimeDate,"
            'w_strSql = w_strSql & "LastUpdTimeDate,"
            'w_strSql = w_strSql & "RegistrantID)"
            'w_strSql = w_strSql & "Values('"
            'w_strSql = w_strSql & Trim(General.g_strHospitalCD) & "'," '病院CD
            'w_strSql = w_strSql & "'" & Trim(m_strMngStaffID) & "'," '職員管理番号
            'w_strSql = w_strSql & p_SelDate & "," '発生日
            'w_strSql = w_strSql & w_DataLoop & "," 'SEQ
            'w_strSql = w_strSql & "'" & Trim(w_DetailInfo(w_DataLoop).strGetDaikyuType) & "'," '取得タイプ
            'w_strSql = w_strSql & w_DetailInfo(w_DataLoop).lngYMD & "," '取得日
            'w_strSql = w_strSql & "'" & Trim(w_DetailInfo(w_DataLoop).strKinmuCD) & "'," '取得勤務CD
            'w_strSql = w_strSql & w_DetailInfo(w_DetaIdx).dblRegistFirstTimeDate & ","
            'w_strSql = w_strSql & w_SysDate & ","
            'w_strSql = w_strSql & "'" & General.g_strUserID & "')"

            'Call General.paDBExecute(w_strSql)

            Call NSK0000H_sql.insert_NS_DAIKYUDETAILMNG_F_02(m_strMngStaffID,
                                                             p_SelDate,
                                                             w_DataLoop,
                                                             w_DetailInfo(w_DataLoop).strGetDaikyuType,
                                                             w_DetailInfo(w_DataLoop).lngYMD,
                                                             w_DetailInfo(w_DataLoop).strKinmuCD,
                                                             w_DetailInfo(w_DetaIdx).dblRegistFirstTimeDate,
                                                             w_SysDate)
            'Upd End
        Next w_DataLoop

        Call General.paCommit()

        fncSetDaikyuGetDate = True
        Exit Function
ErrHandler:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Function
	
    Private Sub m_lstOptDaikyuType_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optDaikyuType_0.CheckedChanged, _optDaikyuType_1.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = m_lstOptDaikyuType.IndexOf(eventSender)
            On Error GoTo Errhnadler
            Const W_SUBNAME As String = "NSK0000HJ m_lstOptDaikyuType_Click"

            Select Case Index
                Case 0
                    '１日代休
                    cmbDaikyuList.Enabled = True
                    m_lstCmbHalfDaikyuList(0).Enabled = False
                    m_lstCmbHalfDaikyuList(1).Enabled = False
                Case 1
                    '半日代休
                    cmbDaikyuList.Enabled = False
                    m_lstCmbHalfDaikyuList(0).Enabled = True
                    m_lstCmbHalfDaikyuList(1).Enabled = True
            End Select

            Exit Sub
Errhnadler:
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End If
    End Sub
	
    Private Sub m_lstCmbHalfDaikyuList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _cmbHalfDaikyuList_0.SelectedIndexChanged, _cmbHalfDaikyuList_1.SelectedIndexChanged
        Dim Index As Short = m_lstCmbHalfDaikyuList.IndexOf(eventSender)
        On Error GoTo Errhnadler
        Const W_SUBNAME As String = "NSK0000HJ m_lstCmbHalfDaikyuList_Click"

        Dim w_lngLoop As Integer
        Dim w_strText_bk As String

        Select Case Index
            Case 0
                '半日代休１入力時
                If m_lstOptDaikyuType(1).Checked = True And m_lstCmbHalfDaikyuList(0).Text <> "" Then
                    If m_ClearFlg = True Then
                        m_lstCmbHalfDaikyuList(1).Items.Clear()

                        '半日代休２に半日代休１以外のデータを格納
                        For w_lngLoop = 1 To UBound(m_HalfDaikyuList_bk)
                            If m_lstCmbHalfDaikyuList(0).Text <> m_HalfDaikyuList_bk(w_lngLoop) Then
                                m_lstCmbHalfDaikyuList(1).Items.Add(m_HalfDaikyuList_bk(w_lngLoop))
                                m_lstCmbHalfDaikyuList(1).SelectedIndex = 0
                            End If
                        Next w_lngLoop
                    End If
                End If
            Case 1
                '半日代休２入力時
                If m_lstOptDaikyuType(1).Checked = True And m_lstCmbHalfDaikyuList(1).Text <> "" Then
                    '半日代休１の内容を退避
                    w_strText_bk = m_lstCmbHalfDaikyuList(0).Text
                    m_lstCmbHalfDaikyuList(0).Items.Clear()

                    '半日代休１に半日代休２以外のデータを格納
                    For w_lngLoop = 1 To UBound(m_HalfDaikyuList_bk)
                        If m_lstCmbHalfDaikyuList(1).Text <> m_HalfDaikyuList_bk(w_lngLoop) Then
                            m_lstCmbHalfDaikyuList(0).Items.Add(m_HalfDaikyuList_bk(w_lngLoop))
                        End If
                    Next w_lngLoop

                    m_ClearFlg = False
                    '選択データを指定しなおす
                    For w_lngLoop = 0 To m_lstCmbHalfDaikyuList(0).Items.Count - 1
                        If m_lstCmbHalfDaikyuList(0).Items(w_lngLoop).ToString = w_strText_bk Then
                            m_lstCmbHalfDaikyuList(0).SelectedIndex = w_lngLoop
                        End If
                    Next w_lngLoop
                m_ClearFlg = True
                End If
        End Select

        Exit Sub
Errhnadler:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub
	
	Private Sub SetForm()
		On Error GoTo ErrHandler
		Const W_SUBNAME As String = "NSK0000HJ SetForm"
		
		Dim w_lngLoop As Integer
		Dim w_lngLoop2 As Integer
		Dim w_intDataLoop As Short
		Dim w_strString As String
		Dim w_strMsg() As String
		Dim w_varWork As Object
		Dim w_DataIdx As Integer
		'フォームプロパティ定数
		Const W_NOMALFORMHEIGHT As Integer = 4000
		Const W_NOMALFORMBOTTOMTOP As Integer = 3060
		Const W_HALFKINMUFORMHEIGHT As Integer = 2550
		Const W_HALFKINMUFORMBOTTOMTOP As Integer = 1620
		
		'初期設定
		m_ClearFlg = True
		w_DataIdx = 0
		ReDim m_HalfDaikyuList_bk(0)
		
		'半日代休対象勤務の場合
		If m_HalfKinmuFlg = True Then
			'フォームサイズとボタンの位置を設定
            Me.Height = General.paTwipsTopixels(W_HALFKINMUFORMHEIGHT)
            Me.cmdOK.Top = General.paTwipsTopixels(W_HALFKINMUFORMBOTTOMTOP)
            Me.cmdCancel.Top = General.paTwipsTopixels(W_HALFKINMUFORMBOTTOMTOP)

            '１日代休用オブジェクトを隠す
            m_lstOptDaikyuType(0).Visible = False
            m_lstOptDaikyuType(1).Visible = False
            m_lstCmbHalfDaikyuList(0).Visible = False
            m_lstCmbHalfDaikyuList(1).Visible = False

            For w_lngLoop = 1 To UBound(m_udtDaikyu)
                Select Case m_udtDaikyu(w_lngLoop).strDaikyuValueType
                    Case CStr(0) '1日
                        w_intDataLoop = 2
                    Case CStr(1) '1.5日
                        w_intDataLoop = 3
                    Case CStr(2) '0.5日
                        w_intDataLoop = 1
                End Select

                'コンボボックスにセット
                For w_lngLoop2 = 1 To w_intDataLoop
                    w_strString = Format(CDate(Format(m_udtDaikyu(w_lngLoop).lngYMD, "0000/00/00")), "yyyy年MM月dd日(ddd)")
                    w_strString = w_strString & Space(3) & m_udtDaikyu(w_lngLoop).strKinmuNM & "(" & m_udtDaikyu(w_lngLoop).strKinmuMark & ")"
                    cmbDaikyuList.Items.Add(w_strString)
                    w_strString = ""
                Next w_lngLoop2
            Next w_lngLoop

            If cmbDaikyuList.Items.Count > 0 Then
                cmbDaikyuList.SelectedIndex = 0
            End If
        Else
            'フォームサイズとボタンの位置を設定
            Me.Height = General.paTwipsTopixels(W_NOMALFORMHEIGHT)
            Me.cmdOK.Top = General.paTwipsTopixels(W_NOMALFORMBOTTOMTOP)
            Me.cmdCancel.Top = General.paTwipsTopixels(W_NOMALFORMBOTTOMTOP)

            '１日代休用オブジェクトを設定
            m_lstOptDaikyuType(0).Visible = True
            m_lstOptDaikyuType(1).Visible = True
            m_lstCmbHalfDaikyuList(0).Visible = True
            m_lstCmbHalfDaikyuList(1).Visible = True
            m_lstCmbHalfDaikyuList(1).Enabled = True
            m_lstOptDaikyuType(0).Checked = True
            For w_lngLoop = 1 To UBound(m_udtDaikyu)
                w_intDataLoop = 0
                Select Case m_udtDaikyu(w_lngLoop).strDaikyuValueType
                    Case CStr(0) '1日
                        w_intDataLoop = 2
                    Case CStr(1) '1.5日
                        w_intDataLoop = 3
                    Case CStr(2) '0.5日
                        w_intDataLoop = 1
                End Select

                '１日代休用コンボボックスにセット
                If m_udtDaikyu(w_lngLoop).strDaikyuValueType <> "2" And m_udtDaikyu(w_lngLoop).strDaikyuValueType <> "" Then
                    w_strString = Format(CDate(Format(m_udtDaikyu(w_lngLoop).lngYMD, "0000/00/00")), "yyyy年MM月dd日(ddd)")
                    w_strString = w_strString & Space(3) & m_udtDaikyu(w_lngLoop).strKinmuNM & "(" & m_udtDaikyu(w_lngLoop).strKinmuMark & ")"
                    cmbDaikyuList.Items.Add(w_strString)
                    w_strString = ""
                End If

                '半日＋半日代休用コンボボックスにセット
                '代休発生日が２件以上ある場合
                For w_lngLoop2 = 1 To w_intDataLoop
                    w_strString = Format(CDate(Format(m_udtDaikyu(w_lngLoop).lngYMD, "0000/00/00")), "yyyy年MM月dd日(ddd)")
                    w_strString = w_strString & Space(3) & m_udtDaikyu(w_lngLoop).strKinmuNM & "(" & m_udtDaikyu(w_lngLoop).strKinmuMark & ")"
                    w_strString = w_strString & "_" & w_lngLoop2

                    w_DataIdx = w_DataIdx + 1
                    '2件目以外をリスト1に追加
                    If w_DataIdx <> 2 Then
                        m_lstCmbHalfDaikyuList(0).Items.Add(w_strString)
                    End If

                    '2件目以降をリスト2に追加
                    If w_DataIdx > 1 Then
                        m_lstCmbHalfDaikyuList(1).Items.Add(w_strString)
                    End If

                    'データを退避
                    ReDim Preserve m_HalfDaikyuList_bk(UBound(m_HalfDaikyuList_bk) + 1)
                    m_HalfDaikyuList_bk(UBound(m_HalfDaikyuList_bk)) = w_strString

                    w_strString = ""
                Next w_lngLoop2
            Next w_lngLoop

            '初期選択
            '１日代休
            If cmbDaikyuList.Items.Count > 0 Then
                cmbDaikyuList.SelectedIndex = 0
            Else
                cmbDaikyuList.Enabled = False
                m_lstOptDaikyuType(0).Enabled = False
                m_lstOptDaikyuType(1).Checked = True
            End If

            '半日代休１
            If w_DataIdx > 0 Then
                m_lstCmbHalfDaikyuList(0).SelectedIndex = 0
            End If

            '半日代休２
            If w_DataIdx > 1 Then
                m_lstCmbHalfDaikyuList(1).SelectedIndex = 0
            Else
                m_lstOptDaikyuType(1).Enabled = False
            End If
		End If
		
		Exit Sub
ErrHandler: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
		End
    End Sub

    Private Sub subSetCtlList()
        m_lstCmbHalfDaikyuList.Add(_cmbHalfDaikyuList_0)
        m_lstCmbHalfDaikyuList.Add(_cmbHalfDaikyuList_1)

        m_lstOptDaikyuType.Add(_optDaikyuType_0)
        m_lstOptDaikyuType.Add(_optDaikyuType_1)
    End Sub
End Class