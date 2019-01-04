Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Friend Class frmNSK0000HJ
    Inherits General.FormBase
    Public NSK0000H_sql As New clsNSK0000H_sql

	Private Structure Daikyu
		Dim lngYMD As Integer
		Dim strKinmuCD As String '�x���o�΋Ζ�CD
		Dim strKinmuNM As String '�x���o�΋Ζ�����
		Dim strKinmuMark As String '�x���o�΋Ζ��L��
        Dim strDaikyuValueType As String '(0:1���A1:1.5���A2:0.5��)
    End Structure

	Private m_udtDaikyu() As Daikyu
    Private m_blnEndStatus As Boolean '�I�����

	'*** �����è�󂯎��
	Private m_strSelDate As String '�w�肳��Ă���N����
	Private m_strSelKinmuCD As String '�I�����ꂽ�Ζ�CD
	Private m_strMngStaffID As String '�E���Ǘ��ԍ�
	Private m_KeikakuFlg As String '�N���������׸�("0":�v���� ����ȊO:�������)
	Private m_Index As Integer
	Private m_SelDate As Integer '�I�����ꂽ���t
    Private m_HalfDaikyuList_bk() As String
	Private m_SelDate2 As Integer '�I�����ꂽ���t
	Private m_HalfKinmuFlg As Boolean '������x�ΏۋΖ�Flg
	Private m_ClearFlg As Boolean
    Private m_GetDaikyuType As Object
    Private m_lstCmbHalfDaikyuList As New List(Of Object)
    Private m_lstOptDaikyuType As New List(Of Object)

	Private Structure DaikyuDetailType
		Dim lngYMD As Integer '�擾�N����
		Dim strKinmuCD As String '�Ζ�CD
		Dim strGetDaikyuType As String '�擾�^�C�v(0:1���A1:0.5��)
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
		
        '�f�[�^���擾
		m_SelDate = 0
		m_SelDate2 = 0
        If m_HalfKinmuFlg = True Or m_lstOptDaikyuType(0).Checked = True Then
            '�����܂��͂P����x�̏ꍇ
            w_str = Format(CDate(General.paLeft(cmbDaikyuList.Text, 11)), "yyyyMMdd")
            For w_lngLoop = 1 To UBound(m_udtDaikyu)
                If CDbl(w_str) = m_udtDaikyu(w_lngLoop).lngYMD Then
                    m_SelDate = m_udtDaikyu(w_lngLoop).lngYMD
                    Exit For
                End If
            Next w_lngLoop

            If m_KeikakuFlg <> "0" Then
                '�v���ʈȊO����Ă΂ꂽ�ꍇ�A��x�擾�N�������X�V����B
                Call fncSetDaikyuGetDate(m_SelDate)
            End If
        Else
            '�����{������x�̏ꍇ
            w_str = Format(CDate(General.paLeft(m_lstCmbHalfDaikyuList(0).Text, 11)), "yyyyMMdd")
            For w_lngLoop = 1 To UBound(m_udtDaikyu)
                If CDbl(w_str) = m_udtDaikyu(w_lngLoop).lngYMD Then
                    m_SelDate = m_udtDaikyu(w_lngLoop).lngYMD
                    Exit For
                End If
            Next w_lngLoop

            '�װ����
            w_str = Format(CDate(General.paLeft(m_lstCmbHalfDaikyuList(1).Text, 11)), "yyyyMMdd")

            '��������I�����Ă���ꍇ�m�f
            If w_str = "" Then
                '*******ү����***********************************
                ReDim w_strMsg(1)
                w_strMsg(1) = "�Ώۑ�x"
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

            '��������I�����Ă���ꍇ�m�f
            If m_SelDate = m_SelDate2 Then
                '*******ү����***********************************
                ReDim w_strMsg(1)
                w_strMsg(1) = "���t"
                Call General.paMsgDsp("NS0003", w_strMsg)
                '************************************************
                Call General.paSetFocus(m_lstCmbHalfDaikyuList(0))
                Exit Sub
            End If

            If m_KeikakuFlg <> "0" Then
                '�v���ʈȊO����Ă΂ꂽ�ꍇ�A��x�擾�N�������X�V����B
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
	
	'�I����Ԃ����n��
	Public ReadOnly Property pEndStatus() As Boolean
		Get
			pEndStatus = m_blnEndStatus
		End Get
	End Property
	
	'�I�����ꂽ���t�����n��
    '�I��N�������󂯎��
	Public Property pSelDate() As Integer
		Get
			pSelDate = m_SelDate
        End Get

		Set(ByVal Value As Integer)
			m_strSelDate = CStr(Value)
		End Set
	End Property
	
	'��x�����󂯎��
    Public WriteOnly Property pDaikyuData(ByVal p_HolDate As Integer, ByVal p_HolKinmuCD As String) As Double
        Set(ByVal Value As Double)

            Dim w_str As String

            '�z��g��
            m_Index = UBound(m_udtDaikyu) + 1
            ReDim Preserve m_udtDaikyu(m_Index)

            '�x���o�Γ�
            m_udtDaikyu(m_Index).lngYMD = p_HolDate

            '�x���o�΋Ζ�CD
            m_udtDaikyu(m_Index).strKinmuCD = p_HolKinmuCD

            '��x�^�C�v
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
	
	'�N�����i�v���ʂ��炩�ǂ����j���󂯎�� ("0": �v���ʂ���@����ȊO:������ʂ���)
	Public WriteOnly Property pKeikakuFlg() As String
		Set(ByVal Value As String)
			m_KeikakuFlg = Value
			
			'��x�z�񏉊���
			ReDim m_udtDaikyu(0)
			m_Index = 0
        End Set
	End Property

	'�I���Ζ�CD���󂯎��
	Public WriteOnly Property pSelKinmuCD() As String
		Set(ByVal Value As String)
			m_strSelKinmuCD = Value
		End Set
	End Property
	
	'�E���Ǘ��ԍ����󂯎��
	Public WriteOnly Property pSelMngStaffID() As String
		Set(ByVal Value As String)
			m_strMngStaffID = Value
		End Set
	End Property
	
	'�I�����ꂽ���t�����n��
	Public ReadOnly Property pSelDate2() As Integer
		Get
			pSelDate2 = m_SelDate2
		End Get
	End Property
	
	'������x�擾�Ζ������n��
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

		'�t�H�[����ݒ�
		Call SetForm()

        Me.StartPosition = FormStartPosition.CenterScreen
		
		Exit Sub
ErrHandler: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
		End
	End Sub
	
	'������ʗp
	Public Function mfncDaikyuDate_Check() As Boolean
		On Error GoTo ErrHandler
        Const W_SUBNAME As String = "NSK0000HJ mfncDaikyuDate_Check"

		Dim w_strSql As String
		Dim w_Rs As ADODB.Recordset
		Dim w_lngDaikyuPastPeriod As Integer '�ߋ��̑�x�擾���̋x���o�Γ��̗L���͈́i�����O�܂ł̋x���o�΂͗L�����Ċ����j
		Dim w_lngDaikyuDate As Integer
		Dim w_objDic As Object
		Dim w_lngKensu As Integer
		Dim w_lngLoop As Integer 'ٰ�߶���
		Dim w_�Ζ�CD_F As ADODB.Field
		Dim w_����_F As ADODB.Field
		Dim w_�L��_F As ADODB.Field
		Dim w_�x���o�ΔN����_F As ADODB.Field
		Dim w_�x���o�΋Ζ�CD_F As ADODB.Field
		Dim w_DaikyuAdvFlg As Integer
		Dim w_lngDaikyuDate_To As Integer
        Dim w_varWork As Object
		Dim w_strString As String
		Dim w_lngCount_Day As Integer
		Dim w_lngCount_HalfDay As Integer
		Dim w_DaikyuAdvThisMonthFlg As Integer
        Dim w_lngDaikyuDate_To2 As Integer
		
		'��x�̗L�����Ԃ����߂�(��̫�Ă͂W�T��)
        w_lngDaikyuPastPeriod = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "PASTDAIKYUPERIOD", "56", General.g_strHospitalCD))
        '��x����t���O
        w_DaikyuAdvFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "DAIKYUADVANCEFLG", CStr(0), General.g_strHospitalCD))

        '��x���蓖���t���O
        w_DaikyuAdvThisMonthFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "DAIKYUADVANCETHISMONTHFLG", CStr(0), General.g_strHospitalCD))

        '��x�擾���Ԃɐ�����������H-----
        If w_lngDaikyuPastPeriod = -1 Then '�����Ȃ�
            If w_DaikyuAdvFlg = 0 Then ''����Ȃ�

                '��x�f�[�^�擾����(�v����Ԃ̊J�n������L�����Ԑ��ߋ��̓��t)
                w_lngDaikyuDate = 0
                '��x�f�[�^�擾����
                w_lngDaikyuDate_To = Integer.Parse(m_strSelDate)
            Else ''���肠��
                '��x�f�[�^�擾����(�v����Ԃ̊J�n������L�����Ԑ��ߋ��̓��t)
                w_lngDaikyuDate = 0

                If w_DaikyuAdvThisMonthFlg = 0 Then
                    w_lngDaikyuDate_To = 99999999
                Else
                    w_lngDaikyuDate_To = Integer.Parse(Format(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(Integer.Parse(General.paLeft(m_strSelDate, 6) & "01"), "0000/00/00")))), "yyyyMMdd"))
                End If
            End If
        Else '��������
            If w_DaikyuAdvFlg = 0 Then ''����Ȃ�

                '��x�f�[�^�擾����(�v����Ԃ̊J�n������L�����Ԑ��ߋ��̓��t)
                w_lngDaikyuDate = Integer.Parse(Format(DateAdd(DateInterval.Day, w_lngDaikyuPastPeriod * -1, CDate(Format(Integer.Parse(m_strSelDate), "0000/00/00"))), "yyyyMMdd"))
                '��x�f�[�^�擾����
                w_lngDaikyuDate_To = Integer.Parse(m_strSelDate)
            Else ''���肠��
                '��x�f�[�^�擾����(�v����Ԃ̊J�n������L�����Ԑ��ߋ��̓��t)
                w_lngDaikyuDate = Integer.Parse(Format(DateAdd(DateInterval.Day, w_lngDaikyuPastPeriod * -1, CDate(Format(Integer.Parse(m_strSelDate), "0000/00/00"))), "yyyyMMdd"))

                '��x�f�[�^�擾����
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

                w_�Ζ�CD_F = .Fields("KinmuCD")
                w_����_F = .Fields("Name")
                w_�L��_F = .Fields("MarkF")
                For w_lngLoop = 1 To w_lngKensu
                    w_objDic.Item(CStr(w_�Ζ�CD_F.Value & "A")) = CStr(w_����_F.Value & "")
                    w_objDic.Item(CStr(w_�Ζ�CD_F.Value & "B")) = CStr(w_�L��_F.Value & "")
                    .MoveNext()
                Next w_lngLoop
            End If
            .Close()
        End With

        w_Rs = Nothing

        '������ʂ���Ă΂ꂽ�ꍇ
        If m_KeikakuFlg <> "0" Then

            '��x�Ǘ��e���擾�\�ȑ�x���t���擾����B
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

                    w_�x���o�ΔN����_F = .Fields("WorkHolKinmuDate")
                    w_�x���o�΋Ζ�CD_F = .Fields("WorkHolKinmuCD")
                    ReDim m_udtDaikyu(w_lngKensu)
                    For w_lngLoop = 1 To w_lngKensu
                        m_udtDaikyu(w_lngLoop).lngYMD = Integer.Parse(w_�x���o�ΔN����_F.Value)
                        m_udtDaikyu(w_lngLoop).strKinmuCD = CStr(w_�x���o�΋Ζ�CD_F.Value & "")
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
            '�v���ʂ���Ă΂ꂽ�ꍇ
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

        '�擾�\��x�����邩�`�F�b�N
        If mfncDaikyuDate_Check = True Then
            '������x�擾�\�Ζ��b�c
            m_HalfKinmuFlg = False

            If g_KinmuM(CShort(m_strSelKinmuCD)).AMCD <> "" And g_KinmuM(CShort(m_strSelKinmuCD)).PMCD <> "" Then
                If g_KinmuM(CShort(g_KinmuM(CShort(m_strSelKinmuCD)).AMCD)).HolBunruiCD = General.G_STRDAIKYUBUNRUI Or g_KinmuM(CShort(g_KinmuM(CShort(m_strSelKinmuCD)).PMCD)).HolBunruiCD = General.G_STRDAIKYUBUNRUI Then
                    m_HalfKinmuFlg = True
                End If
            End If

            w_lngCount_Day = 0
            w_lngCount_HalfDay = 0

            '�P����x
            For w_lngLoop = 1 To UBound(m_udtDaikyu)
                Select Case m_udtDaikyu(w_lngLoop).strDaikyuValueType
                    Case "0"
                        '�P��
                        w_lngCount_Day = w_lngCount_Day + 1
                        w_lngCount_HalfDay = w_lngCount_HalfDay + 2
                    Case "1"
                        '1.5��
                        w_lngCount_Day = w_lngCount_Day + 1
                        w_lngCount_HalfDay = w_lngCount_HalfDay + 3
                    Case "2"
                        w_lngCount_HalfDay = w_lngCount_HalfDay + 1
                End Select
            Next w_lngLoop

            If m_HalfKinmuFlg = True Then
                '�����p�t�H�[��
                If w_lngCount_HalfDay < 1 Then
                    mfncDaikyuDate_Check = False
                End If
            Else
                '�ʏ�t�H�[��
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

    '��x�Ǘ��e�ɑ�x�擾�����X�V����
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
        Dim w_DaikyuType As String '(0:�P��,1:0.5��)
        Dim w_DetailInfo() As DaikyuDetailType

        fncSetDaikyuGetDate = False

        w_SysDate = CDbl(Format(Now, "yyyyMMddHHmmss"))

        '�f�[�^�����ɂ��邩�`�F�b�N
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
                '�ް��Ȃ�
            Else
                '�ް�����
                .MoveLast()
                w_DataCnt = .RecordCount
                .MoveFirst()
                '�z��m��
                ReDim Preserve w_DetailInfo(w_DataCnt)

                w_GETFLG_F = .Fields("GETFLG")
                w_GETDAIKYUDATE_F = .Fields("GETDAIKYUDATE")
                w_GETDAIKYUKINMUCD_F = .Fields("GETDAIKYUKINMUCD")
                w_REGISTFIRSTTIMEDATE_F = .Fields("REGISTFIRSTTIMEDATE")

                For w_DataLoop = 1 To w_DataCnt
                    '�f�[�^�i�[
                    w_DetailInfo(w_DataLoop).strGetDaikyuType = IIf(IsDBNull(w_GETFLG_F.Value), "0", w_GETFLG_F.Value)
                    w_DetailInfo(w_DataLoop).lngYMD = IIf(IsDBNull(w_GETDAIKYUDATE_F.Value), 0, w_GETDAIKYUDATE_F.Value)
                    w_DetailInfo(w_DataLoop).strKinmuCD = IIf(IsDBNull(w_GETDAIKYUKINMUCD_F.Value), "", w_GETDAIKYUKINMUCD_F.Value)
                    w_DetailInfo(w_DataLoop).dblRegistFirstTimeDate = IIf(IsDBNull(w_REGISTFIRSTTIMEDATE_F.Value), 0, w_REGISTFIRSTTIMEDATE_F.Value)

                    .MoveNext()
                Next w_DataLoop
            End If
        End With
        w_Rs.Close()

        '��x�̎擾�^�C�v�擾
        w_DetaIdx = UBound(w_DetailInfo) + 1
        ReDim Preserve w_DetailInfo(w_DetaIdx)

        If m_HalfKinmuFlg = True Or m_lstOptDaikyuType(1).Checked = True Then
            '�����̏ꍇ
            w_DetailInfo(w_DetaIdx).strGetDaikyuType = "1"
        Else
            '�P���̏ꍇ
            w_DetailInfo(w_DetaIdx).strGetDaikyuType = "0"
        End If

        w_DetailInfo(w_DetaIdx).lngYMD = Integer.Parse(m_strSelDate)
        w_DetailInfo(w_DetaIdx).strKinmuCD = m_strSelKinmuCD
        w_DetailInfo(w_DetaIdx).dblRegistFirstTimeDate = w_SysDate

        Call General.paBeginTrans()
        '2017/05/02 Christopher Upd Start
        '�f�[�^�폜
        ''Delete�� �ҏW
        'w_strSql = "Delete From NS_DAIKYUDETAILMNG_F "
        'w_strSql = w_strSql & " where WorkHolKinmuDate = " & p_SelDate
        'w_strSql = w_strSql & " and StaffMngID = '" & Trim(m_strMngStaffID) & "'"
        'w_strSql = w_strSql & " and HospitalCD = '" & Trim(General.g_strHospitalCD) & "'"

        'Call General.paDBExecute(w_strSql)

        Call NSK0000H_sql.delete_NS_DAIKYUDETAILMNG_F_02(p_SelDate, Trim(m_strMngStaffID))
        'Upd End
        For w_DataLoop = 1 To UBound(w_DetailInfo)
            '2017/05/22 Richard Upd Start
            ''Insert�� �ҏW
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
            'w_strSql = w_strSql & Trim(General.g_strHospitalCD) & "'," '�a�@CD
            'w_strSql = w_strSql & "'" & Trim(m_strMngStaffID) & "'," '�E���Ǘ��ԍ�
            'w_strSql = w_strSql & p_SelDate & "," '������
            'w_strSql = w_strSql & w_DataLoop & "," 'SEQ
            'w_strSql = w_strSql & "'" & Trim(w_DetailInfo(w_DataLoop).strGetDaikyuType) & "'," '�擾�^�C�v
            'w_strSql = w_strSql & w_DetailInfo(w_DataLoop).lngYMD & "," '�擾��
            'w_strSql = w_strSql & "'" & Trim(w_DetailInfo(w_DataLoop).strKinmuCD) & "'," '�擾�Ζ�CD
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
                    '�P����x
                    cmbDaikyuList.Enabled = True
                    m_lstCmbHalfDaikyuList(0).Enabled = False
                    m_lstCmbHalfDaikyuList(1).Enabled = False
                Case 1
                    '������x
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
                '������x�P���͎�
                If m_lstOptDaikyuType(1).Checked = True And m_lstCmbHalfDaikyuList(0).Text <> "" Then
                    If m_ClearFlg = True Then
                        m_lstCmbHalfDaikyuList(1).Items.Clear()

                        '������x�Q�ɔ�����x�P�ȊO�̃f�[�^���i�[
                        For w_lngLoop = 1 To UBound(m_HalfDaikyuList_bk)
                            If m_lstCmbHalfDaikyuList(0).Text <> m_HalfDaikyuList_bk(w_lngLoop) Then
                                m_lstCmbHalfDaikyuList(1).Items.Add(m_HalfDaikyuList_bk(w_lngLoop))
                                m_lstCmbHalfDaikyuList(1).SelectedIndex = 0
                            End If
                        Next w_lngLoop
                    End If
                End If
            Case 1
                '������x�Q���͎�
                If m_lstOptDaikyuType(1).Checked = True And m_lstCmbHalfDaikyuList(1).Text <> "" Then
                    '������x�P�̓��e��ޔ�
                    w_strText_bk = m_lstCmbHalfDaikyuList(0).Text
                    m_lstCmbHalfDaikyuList(0).Items.Clear()

                    '������x�P�ɔ�����x�Q�ȊO�̃f�[�^���i�[
                    For w_lngLoop = 1 To UBound(m_HalfDaikyuList_bk)
                        If m_lstCmbHalfDaikyuList(1).Text <> m_HalfDaikyuList_bk(w_lngLoop) Then
                            m_lstCmbHalfDaikyuList(0).Items.Add(m_HalfDaikyuList_bk(w_lngLoop))
                        End If
                    Next w_lngLoop

                    m_ClearFlg = False
                    '�I���f�[�^���w�肵�Ȃ���
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
		'�t�H�[���v���p�e�B�萔
		Const W_NOMALFORMHEIGHT As Integer = 4000
		Const W_NOMALFORMBOTTOMTOP As Integer = 3060
		Const W_HALFKINMUFORMHEIGHT As Integer = 2550
		Const W_HALFKINMUFORMBOTTOMTOP As Integer = 1620
		
		'�����ݒ�
		m_ClearFlg = True
		w_DataIdx = 0
		ReDim m_HalfDaikyuList_bk(0)
		
		'������x�ΏۋΖ��̏ꍇ
		If m_HalfKinmuFlg = True Then
			'�t�H�[���T�C�Y�ƃ{�^���̈ʒu��ݒ�
            Me.Height = General.paTwipsTopixels(W_HALFKINMUFORMHEIGHT)
            Me.cmdOK.Top = General.paTwipsTopixels(W_HALFKINMUFORMBOTTOMTOP)
            Me.cmdCancel.Top = General.paTwipsTopixels(W_HALFKINMUFORMBOTTOMTOP)

            '�P����x�p�I�u�W�F�N�g���B��
            m_lstOptDaikyuType(0).Visible = False
            m_lstOptDaikyuType(1).Visible = False
            m_lstCmbHalfDaikyuList(0).Visible = False
            m_lstCmbHalfDaikyuList(1).Visible = False

            For w_lngLoop = 1 To UBound(m_udtDaikyu)
                Select Case m_udtDaikyu(w_lngLoop).strDaikyuValueType
                    Case CStr(0) '1��
                        w_intDataLoop = 2
                    Case CStr(1) '1.5��
                        w_intDataLoop = 3
                    Case CStr(2) '0.5��
                        w_intDataLoop = 1
                End Select

                '�R���{�{�b�N�X�ɃZ�b�g
                For w_lngLoop2 = 1 To w_intDataLoop
                    w_strString = Format(CDate(Format(m_udtDaikyu(w_lngLoop).lngYMD, "0000/00/00")), "yyyy�NMM��dd��(ddd)")
                    w_strString = w_strString & Space(3) & m_udtDaikyu(w_lngLoop).strKinmuNM & "(" & m_udtDaikyu(w_lngLoop).strKinmuMark & ")"
                    cmbDaikyuList.Items.Add(w_strString)
                    w_strString = ""
                Next w_lngLoop2
            Next w_lngLoop

            If cmbDaikyuList.Items.Count > 0 Then
                cmbDaikyuList.SelectedIndex = 0
            End If
        Else
            '�t�H�[���T�C�Y�ƃ{�^���̈ʒu��ݒ�
            Me.Height = General.paTwipsTopixels(W_NOMALFORMHEIGHT)
            Me.cmdOK.Top = General.paTwipsTopixels(W_NOMALFORMBOTTOMTOP)
            Me.cmdCancel.Top = General.paTwipsTopixels(W_NOMALFORMBOTTOMTOP)

            '�P����x�p�I�u�W�F�N�g��ݒ�
            m_lstOptDaikyuType(0).Visible = True
            m_lstOptDaikyuType(1).Visible = True
            m_lstCmbHalfDaikyuList(0).Visible = True
            m_lstCmbHalfDaikyuList(1).Visible = True
            m_lstCmbHalfDaikyuList(1).Enabled = True
            m_lstOptDaikyuType(0).Checked = True
            For w_lngLoop = 1 To UBound(m_udtDaikyu)
                w_intDataLoop = 0
                Select Case m_udtDaikyu(w_lngLoop).strDaikyuValueType
                    Case CStr(0) '1��
                        w_intDataLoop = 2
                    Case CStr(1) '1.5��
                        w_intDataLoop = 3
                    Case CStr(2) '0.5��
                        w_intDataLoop = 1
                End Select

                '�P����x�p�R���{�{�b�N�X�ɃZ�b�g
                If m_udtDaikyu(w_lngLoop).strDaikyuValueType <> "2" And m_udtDaikyu(w_lngLoop).strDaikyuValueType <> "" Then
                    w_strString = Format(CDate(Format(m_udtDaikyu(w_lngLoop).lngYMD, "0000/00/00")), "yyyy�NMM��dd��(ddd)")
                    w_strString = w_strString & Space(3) & m_udtDaikyu(w_lngLoop).strKinmuNM & "(" & m_udtDaikyu(w_lngLoop).strKinmuMark & ")"
                    cmbDaikyuList.Items.Add(w_strString)
                    w_strString = ""
                End If

                '�����{������x�p�R���{�{�b�N�X�ɃZ�b�g
                '��x���������Q���ȏ゠��ꍇ
                For w_lngLoop2 = 1 To w_intDataLoop
                    w_strString = Format(CDate(Format(m_udtDaikyu(w_lngLoop).lngYMD, "0000/00/00")), "yyyy�NMM��dd��(ddd)")
                    w_strString = w_strString & Space(3) & m_udtDaikyu(w_lngLoop).strKinmuNM & "(" & m_udtDaikyu(w_lngLoop).strKinmuMark & ")"
                    w_strString = w_strString & "_" & w_lngLoop2

                    w_DataIdx = w_DataIdx + 1
                    '2���ڈȊO�����X�g1�ɒǉ�
                    If w_DataIdx <> 2 Then
                        m_lstCmbHalfDaikyuList(0).Items.Add(w_strString)
                    End If

                    '2���ڈȍ~�����X�g2�ɒǉ�
                    If w_DataIdx > 1 Then
                        m_lstCmbHalfDaikyuList(1).Items.Add(w_strString)
                    End If

                    '�f�[�^��ޔ�
                    ReDim Preserve m_HalfDaikyuList_bk(UBound(m_HalfDaikyuList_bk) + 1)
                    m_HalfDaikyuList_bk(UBound(m_HalfDaikyuList_bk)) = w_strString

                    w_strString = ""
                Next w_lngLoop2
            Next w_lngLoop

            '�����I��
            '�P����x
            If cmbDaikyuList.Items.Count > 0 Then
                cmbDaikyuList.SelectedIndex = 0
            Else
                cmbDaikyuList.Enabled = False
                m_lstOptDaikyuType(0).Enabled = False
                m_lstOptDaikyuType(1).Checked = True
            End If

            '������x�P
            If w_DataIdx > 0 Then
                m_lstCmbHalfDaikyuList(0).SelectedIndex = 0
            End If

            '������x�Q
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