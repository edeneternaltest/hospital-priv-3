Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Friend Class frmNSK0000HH
    Inherits General.FormBase
    Public NSK0000H_sql As New clsNSK0000H_sql
    '�����ΏۋΖ��I���߼�ݲ��ޯ��
    Private Const M_OPT_STANDARD As Short = 0
    Private Const M_OPT_SELECT As Short = 1

    '�Ζ���(�����������Ɏg�p)
    Private Const M_KINMU_NUM As Short = 6

    '�����������s�N���X�I�u�W�F�N�g�ϐ�
    Private m_AutoSched As Object

    '�����������s�����Ұ�
    Private m_CalStartDate As Integer '����ް�J�n��
    Private m_CalEndDate As Integer '����ް�I����
    Private m_ScheduleStartDate As Integer '�����J�n��
    Private m_ScheduleEndDate As Integer '�����I����
    Private m_DisplayStartDate As Integer '�\�����ԊJ�n��
    Private m_DisplayEndDate As Integer '�\�����ԏI����
    Private m_ScheduleStartCol As Integer '�����J�n��
    Private m_ScheduleEndCol As Integer '�����I����
    Private m_DisplayStartCol As Integer '�\�����ԊJ�n��
    Private m_DisplayEndCol As Integer '�\�����ԏI����
    Private m_UserStartCol As Integer 'հ�ް�w��J�n��
    Private m_UserEndCol As Integer 'հ�ް�w��I����
    Private m_SelectDate As Long     '���������J�n��
    '2012/11/16 Ishiga add start-------------
    Private m_PlanNO As Long     '�v��ԍ�
    '2012/11/16 Ishiga add end---------------

    '2014/06/12 TAKEBAYASHI P-07100 Add (�ϐ��̒ǉ�) START-->>
    Private m_SelTeamNo As Integer     '�I�����ꂽ��єԍ�
    Private m_OuenTeam As Integer     '��������єԍ�
    Private m_TeamCnt As Integer     '��ь���
    '2014/06/12 TAKEBAYASHI P-07100 Add (�ϐ��̒ǉ�) END<<--

    '�����������s�ς��׸�
    Private m_SchedExecute As Boolean
    '����������ݾ�Ӱ��
    Private m_CancelMode As Boolean 'True�̏ꍇ�͎����Эڰ��̷݂�ݾ�
    '�����������ʔ��f�׸ށi�I���㊄�����ʂ��v���ʂɔ��f�����邩�H�j
    Private m_SchedSaveFlg As Boolean

    '�����Ζ��i�[�\���́E�ϐ�
    Private Structure KinmuMaster_Type
        Dim KinmuCD As String '�Ζ�����
        Dim KinmuName As String 'KinmuName
        Dim BunruiCD As String '���޺���
        Dim OnOffFlg As Boolean '�I������׸�
    End Structure

    Private m_NsKinmuMData() As KinmuMaster_Type
    Private Const M_BackColor As String = "&H8080FF" '�I������ޯ��װ
    Private Const M_NomalColor As String = "&H8000000F" '�I������Ă��Ȃ��Ƃ��ޯ��װ
    Private m_LoadError As Boolean
    Private m_KikanStartDate As String '�����ޯ���I���J�n�N����
    Private m_KikanEndDate As String '�����ޯ���I���I���N����
    Private m_SelectKikanNo As Short '�����ޯ���I�����Բ��ޯ��
    Private m_lstCmdKinmu As New List(Of Object)

    Private Sub GetKinmuList()
        On Error GoTo GetKinmuList
        Const W_SUBNAME As String = "NSK0000HH GetKinmuList"

        Dim w_Sql As String
        Dim w_Rs As ADODB.Recordset
        Dim w_i As Short
        Dim w_Cnt As Short
        Dim w_Code_F As ADODB.Field
        Dim w_Name_F As ADODB.Field
        Dim w_BCode_F As ADODB.Field
        Dim w_strMsg() As String
        '2017/05/02 Christopher Upd Start
        ''--- �Ζ�����Ͻ�(�����Ώۂ̂���) -----
        'w_Sql = "Select KinmuCD, Name, AllocBunruiCD "
        'w_Sql = w_Sql & "From NS_KINMUNAME_M "
        'w_Sql = w_Sql & "Where AllocFlg = '2' "
        'w_Sql = w_Sql & "And HospitalCD = '" & General.g_strHospitalCD & "' "
        ''2014/06/13 TAKEBAYASHI P-07100 Add (�L�������I�����������ɒǉ�) START-->>
        ''Ͻ�����ݽ�ŗL�������I������ݒ�ł��Ȃ���(AllocFlg=2�̏ꍇ)�A���ĉ�(�ǉ��������)
        ''w_Sql = w_Sql & "And (EFFTODATE > " & CInt(Convert.ToString(m_DisplayEndDate))
        ''w_Sql = w_Sql & "OR EFFTODATE = 0) "
        ''2014/06/13 TAKEBAYASHI P-07100 Add (�L�������I�����������ɒǉ�) END<<--
        'w_Sql = w_Sql & "Order By DispNo "
        ''ں��޾�ĵ�޼ު�Đ���
        'w_Rs = General.paDBRecordSetOpen(w_Sql)

        Call NSK0000H_sql.select_NS_KINMUNAME_M_03(w_Rs)
        'Upd End
        With w_Rs

            'ں��ތ�����m�邽�ߍŏI�s�Ɉړ�����
            If .BOF = True And .EOF = True Then
                ReDim w_strMsg(1)
                w_strMsg(1) = "�����Ζ�"
                Call General.paMsgDsp("NS0010", w_strMsg)
                m_LoadError = True
                w_Cnt = 0
            Else
                'ں��ތ����i�[
                .MoveLast()
                w_Cnt = .RecordCount
                .MoveFirst()
                w_Code_F = .Fields("KinmuCD")
                w_Name_F = .Fields("Name")
                w_BCode_F = .Fields("AllocBunruiCD")
            End If

            '�z��Ɋi�[
            ReDim m_NsKinmuMData(w_Cnt)

            For w_i = 1 To w_Cnt
                m_NsKinmuMData(w_i - 1).KinmuCD = w_Code_F.Value
                m_NsKinmuMData(w_i - 1).KinmuName = w_Name_F.Value & ""
                m_NsKinmuMData(w_i - 1).BunruiCD = w_BCode_F.Value & ""
                m_NsKinmuMData(w_i - 1).OnOffFlg = False
                .MoveNext()
            Next w_i

        End With

        '��޼ު�Ẳ��
        w_Rs.Close()

        Exit Sub
GetKinmuList:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    '̫��۰�޲���Ă�����ɏI��������?
    Public ReadOnly Property pLoadState() As Boolean
        Get
            'True:����.False:�ُ�
            pLoadState = m_LoadError
        End Get
    End Property

    '����ް�I����
    Public WriteOnly Property pAutoCalEndDate() As Integer
        Set(ByVal Value As Integer)
            m_CalEndDate = Value
        End Set
    End Property

    '����ް�J�n��
    Public WriteOnly Property pAutoCalStartDate() As Integer
        Set(ByVal Value As Integer)
            m_CalStartDate = Value
        End Set
    End Property

    '�\�����ԏI����
    Public WriteOnly Property pAutoDisplayEndCol() As Integer
        Set(ByVal Value As Integer)
            m_DisplayEndCol = Value
        End Set
    End Property

    '�\�����ԏI����
    Public WriteOnly Property pAutoDisplayEndDate() As Integer
        Set(ByVal Value As Integer)
            m_DisplayEndDate = Value
        End Set
    End Property

    '�\�����ԊJ�n��
    Public WriteOnly Property pAutoDisplayStartCol() As Integer
        Set(ByVal Value As Integer)
            m_DisplayStartCol = Value
        End Set
    End Property

    '�\�����ԊJ�n��
    Public WriteOnly Property pAutoDisplayStartDate() As Integer
        Set(ByVal Value As Integer)
            m_DisplayStartDate = Value
        End Set
    End Property

    '�����I����
    Public WriteOnly Property pAutoScheduleEndCol() As Integer
        Set(ByVal Value As Integer)
            m_ScheduleEndCol = Value
        End Set
    End Property

    '�����I����
    Public WriteOnly Property pAutoScheduleEndDate() As Integer
        Set(ByVal Value As Integer)
            m_ScheduleEndDate = Value
        End Set
    End Property

    '�����J�n��
    Public WriteOnly Property pAutoScheduleStartCol() As Integer
        Set(ByVal Value As Integer)
            m_ScheduleStartCol = Value
        End Set
    End Property

    '�����J�n��
    Public WriteOnly Property pAutoScheduleStartDate() As Integer
        Set(ByVal Value As Integer)
            m_ScheduleStartDate = Value
        End Set
    End Property

    'հ�ް�w��I����
    Public WriteOnly Property pAutoUserEndCol() As Integer
        Set(ByVal Value As Integer)
            m_UserEndCol = Value
        End Set
    End Property

    'հ�ް�w��J�n��
    Public WriteOnly Property pAutoUserStartCol() As Integer
        Set(ByVal Value As Integer)
            m_UserStartCol = Value
        End Set
    End Property

    Public ReadOnly Property pSchedSaveFlg() As Boolean
        Get
            pSchedSaveFlg = m_SchedSaveFlg
        End Get
    End Property

    '�����ޯ���I���I���N����
    Public WriteOnly Property pKikanEnd() As String
        Set(ByVal Value As String)
            m_KikanEndDate = Value
        End Set
    End Property

    '�����ޯ���I���J�n�N����
    Public WriteOnly Property pKikanStart() As String
        Set(ByVal Value As String)
            m_KikanStartDate = Value
        End Set
    End Property

    '���������J�n��
    Public ReadOnly Property p_SelectDate() As Long
        Get
            Return m_SelectDate
        End Get
    End Property

    '2014/06/12 TAKEBAYASHI P-07100 Add (���Ұ��p�����èҿ��ޒǉ�) START-->>
    '�I����єԍ�
    Public WriteOnly Property p_SelTeamNo() As Integer
        Set(ByVal Value As Integer)
            m_SelTeamNo = Value
        End Set
    End Property
    '������єԍ�
    Public WriteOnly Property p_OuenTeam() As Integer
        Set(ByVal Value As Integer)
            m_OuenTeam = Value
        End Set
    End Property
    '��ь���
    Public WriteOnly Property p_TeamCnt() As Integer
        Set(ByVal Value As Integer)
            m_TeamCnt = Value
        End Set
    End Property
    '2014/06/12 TAKEBAYASHI P-07100 Add (���Ұ��p�����èҿ��ޒǉ�) END<<--

    '2012/11/16 Ishiga add start----------------------------------
    '�v��ԍ�
    Public WriteOnly Property p_PlanNO() As String
        Set(ByVal Value As String)
            m_PlanNO = Value
        End Set
    End Property
    '2012/11/16 Ishiga add start----------------------------------

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub m_lstCmdKinmu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _cmdKinmu_0.Click, _cmdKinmu_1.Click, _
                                                                                                                    _cmdKinmu_2.Click, _cmdKinmu_3.Click, _
                                                                                                                    _cmdKinmu_4.Click, _cmdKinmu_5.Click
        Dim Index As Short = m_lstCmdKinmu.IndexOf(eventSender)
        Dim w_Cap As String
        Dim w_i As Short
        Dim w_Font As Font

        w_Cap = m_lstCmdKinmu(Index).Text

        For w_i = 1 To UBound(m_NsKinmuMData)
            If w_Cap = m_NsKinmuMData(w_i - 1).KinmuName Then
                If m_NsKinmuMData(w_i - 1).OnOffFlg = True Then
                    m_NsKinmuMData(w_i - 1).OnOffFlg = False
                Else
                    m_NsKinmuMData(w_i - 1).OnOffFlg = True
                End If
            End If
        Next w_i

        If ColorTranslator.ToOle(m_lstCmdKinmu(Index).BackColor) = CDbl(M_BackColor) Then
            m_lstCmdKinmu(Index).BackColor = SystemColors.Control
            w_Font = m_lstCmdKinmu(Index).Font
            m_lstCmdKinmu(Index).Font = New Font(w_Font, FontStyle.Regular)
        Else
            m_lstCmdKinmu(Index).BackColor = ColorTranslator.FromOle(CDbl(M_BackColor))
            w_Font = m_lstCmdKinmu(Index).Font
            m_lstCmdKinmu(Index).Font = New Font(w_Font, FontStyle.Bold)
        End If

    End Sub

    Private Sub cmdSchedExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSchedExit.Click
        On Error GoTo cmdSchedExit_Click
        Const W_SUBNAME As String = "Nske001 cmdSchedExit_Click"

        Dim w_Res As Short
        Dim w_Rtn As Boolean
        Dim w_strMsg() As String

        '�J�n�{�^�����g�p�\�Ƃ��܂��B
        cmdSchedStart.Enabled = True

        '����{�^�����g�p�s�Ƃ��܂��B
        cmdSchedExit.Text = "�I��(&E)"

        If m_SchedExecute <> True Then
            Me.Close()
            Exit Sub
        End If

        '��ݾق̏ꍇ
        If m_CancelMode = True Then
            '�����Эڰ��̷݂�ݾ�

            ReDim w_strMsg(2)
            w_strMsg(1) = "�X�P�W���[�����O"
            w_strMsg(2) = "���~"
            w_Res = General.paMsgDsp("NS0097", w_strMsg)

            If w_Res = MsgBoxResult.Yes Then
                '��ݾ�SW��ON�ɂ���
                m_AutoSched.p_CancelSW = True
                '���~ү���ނ�\��
                lblSchedMessage.Text = "���ޭ��ݸނ𒆎~���܂����D�D�D"
                '�����o�߂��N���A����
                prbSchedProcess.Minimum = 0
                prbSchedProcess.Maximum = 10
                prbSchedProcess.Value = 0
            End If

            Exit Sub

        End If

        '���s�ð��(�����è�)���Q�Ƃ����s����Ă��������̪���ް����쐬
        If m_AutoSched.p_SchedStaus = "EXEC" Then
            '�����Эڰ��ݏI��ҿ��ނ̎��s
            w_Rtn = m_AutoSched.mSchedExit
            If w_Rtn = False Then
                '�����Эڰ��ݏI��ҿ��ނŎ��s���װ�����I�I
                Me.Close()
                Exit Sub
            End If

            '�������ʂ�ۑ����܂��B
            w_Rtn = m_AutoSched.mMakeOutputIfTbl
            If w_Rtn = False Then
                '�������ʕۑ�ҿ��ނŎ��s���װ�����I�I
                Me.Close()
                Exit Sub
            End If

            '���ʔ��f�׸ނ�ON�ɂ���
            m_SchedSaveFlg = True
        End If

        Me.Close()

        Exit Sub
cmdSchedExit_Click:
        Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
        End

    End Sub

    Private Sub cmdSchedStart_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSchedStart.Click
        On Error GoTo cmdSchedStart_Click
        Const W_SUBNAME As String = "NSK0000HH cmdSchedStart_Click"

        '�I���Ζ��i�[ڼ޽�؂�Key
        Const M_RegKey As String = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY2 & "\" & "NSK0020B"
        Const M_RegSubKey As String = "List_Check"
        Const M_RegValue As String = "Field"

        Dim w_RegKey As String
        Dim w_i As Short
        Dim w_Res As Short
        Dim w_RegPath As String
        Dim w_Select As String
        Dim w_Rtn As Boolean
        Dim w_Date As Integer
        Dim w_SelectDate As Integer '���͂��ꂽ���t
        Dim w_strMsg() As String

        'Ini�̾���ݖ���
        w_RegPath = "NSK0000H"

        '�����V�~�����[�V�������s�m�F���b�Z�[�W
        ReDim w_strMsg(2)
        w_strMsg(1) = "�����V�~�����[�V����"
        w_strMsg(2) = "�J�n"
        w_Res = General.paMsgDsp("NS0097", w_strMsg)

        If w_Res = MsgBoxResult.Yes Then
            '[�͂�] �{�^����I�������ꍇ

            Application.DoEvents()

            '�����J�n��
            If imdDate.Text = "    /  /" Then
                ReDim w_strMsg(1)
                w_strMsg(1) = "�����J�n��"
                Call General.paMsgDsp("NS0001", w_strMsg)
                imdDate.Focus()
                Exit Sub
            End If

            '���͓�����
            w_Date = Integer.Parse(Format(CDate(imdDate.Text), "yyyyMMdd"))
            Select Case m_SelectKikanNo

                Case 1, 2
                    If m_DisplayStartDate <= w_Date And w_Date <= m_DisplayEndDate Then
                    Else
                        ReDim w_strMsg(1)
                        w_strMsg(1) = "�����J�n��"
                        Call General.paMsgDsp("NS0003", w_strMsg)
                        imdDate.Focus()
                        Exit Sub
                    End If
                Case 3
                    If CDbl(m_KikanStartDate) <= w_Date And w_Date <= m_DisplayEndDate Then
                    Else
                        ReDim w_strMsg(1)
                        w_strMsg(1) = "�����J�n��"
                        Call General.paMsgDsp("NS0003", w_strMsg)
                        imdDate.Focus()
                        Exit Sub
                    End If
            End Select

            '�I���Ζ�������́H
            w_Select = Convert.ToString(False) '�I���Ζ������׸ޏ�����
            For w_i = 1 To UBound(m_NsKinmuMData)
                If m_NsKinmuMData(w_i - 1).OnOffFlg = True Then
                    '�I���Ζ�����
                    w_Select = Convert.ToString(True)
                    Exit For
                End If
            Next w_i

            If CBool(w_Select) = True Then

                '�I����Ԃ�ڼ޽�؂ɏ�������
                For w_i = 1 To UBound(m_NsKinmuMData)
                    '�I����ԁH
                    If m_NsKinmuMData(w_i - 1).OnOffFlg = True Then
                        '�I���Ζ�
                        Call General.paSaveSetting(M_RegKey, M_RegSubKey, M_RegValue & Str(w_i), "1")
                    Else
                        '���I���Ζ�
                        Call General.paSaveSetting(M_RegKey, M_RegSubKey, M_RegValue & Str(w_i), "0")
                    End If

                Next w_i

            End If

            If CBool(w_Select) = False Then
                ReDim w_strMsg(1)
                w_strMsg(1) = "�����ΏۋΖ�"
                Call General.paMsgDsp("NS0002", w_strMsg)
                Exit Sub
            End If

            '2014/06/18 TAKEBAYASHI P-07100 Add (�I����є���) START-->>
            If m_OuenTeam > 0 And m_OuenTeam = m_SelTeamNo Then
                ReDim w_strMsg(1)
                w_strMsg(1) = "�`�[��"
                Call General.paMsgDsp("NS0276", w_strMsg)
                Exit Sub
            End If
            '2014/06/18 TAKEBAYASHI P-07100 Add (�I����є���) END<<--

            'Text�ɓ��͂��ꂽ���t��ϐ��Ɋi�[
            w_SelectDate = General.paGetDateIntegerFromDate(imdDate.Value, General.G_DATE_ENUM.yyyyMMdd)
            m_SelectDate = w_SelectDate

            '�J�n�{�^�����g�p�s�Ƃ��܂��B
            cmdSchedStart.Enabled = False

            '�������玩���Эڰ��݂̊e�������è���ݒ肵�܂��B
            '�{�݃R�[�h
            m_AutoSched.p_HospitalCD = General.g_strHospitalCD
            '�Ζ������R�[�h
            m_AutoSched.p_KangoTCD = General.g_strSelKinmuDeptCD
            '�J�����_�[�J�n�N����
            m_AutoSched.p_calstart_ymd = Convert.ToString(m_CalStartDate)
            '�J�����_�[�I���N����
            m_AutoSched.p_calend_ymd = Convert.ToString(m_CalEndDate)
            '�����J�n�N����
            m_AutoSched.p_schedstart_ymd = Convert.ToString(m_ScheduleStartDate)
            '�����I���N����
            m_AutoSched.p_schedend_ymd = Convert.ToString(m_ScheduleEndDate)
            '�����J�n��
            m_AutoSched.p_schedstart_col = Convert.ToString(m_ScheduleStartCol)
            '�����I����
            m_AutoSched.p_schedend_col = Convert.ToString(m_ScheduleEndCol)
            '���͂��ꂽ���t
            m_AutoSched.p_SelectDate = Convert.ToString(w_SelectDate)
            '���[�U�[�w��J�n��
            m_AutoSched.p_usrstart_col = Convert.ToString(m_UserStartCol)
            '���[�U�[�w��I����
            m_AutoSched.p_usrend_col = Convert.ToString(m_UserEndCol)

            '2014/06/12 TAKEBAYASHI P-07100 Add (���Ұ��̐ݒ�) START-->>
            m_AutoSched.p_SelTeamNo = m_SelTeamNo
            m_AutoSched.p_TeamCnt = m_TeamCnt
            '2014/06/12 TAKEBAYASHI P-07100 Add (���Ұ��̐ݒ�) END<<--

            '2012/11/16 Ishiga add start--------
            '�v��ԍ�
            m_AutoSched.p_PlanNO = m_PlanNO
            '2012/11/16 Ishiga add end----------

            '�\���J�n�N����
            m_AutoSched.p_dispstart_ymd = Convert.ToString(m_DisplayStartDate)
            '�\���I���N����
            m_AutoSched.p_dispend_ymd = Convert.ToString(m_DisplayEndDate)
            '�\���J�n��
            m_AutoSched.p_dispstart_col = Convert.ToString(m_DisplayStartCol)
            '�\���I����
            m_AutoSched.p_dispend_col = Convert.ToString(m_DisplayEndCol)

            '���W�X�g���̎擾���z��ɑΏہ^��Ώۂ�ݒ�
            For w_i = 1 To UBound(m_NsKinmuMData)
                w_Select = General.paGetSetting(M_RegKey, M_RegSubKey, M_RegValue & Str(w_i), "0")
                If w_Select <> "0" Then
                    '2014/06/19 TAKEBAYASHI P-07100 Change (�Ζ����ނ�������A������悤�ɕύX)
                    m_AutoSched.p_scheditem = m_AutoSched.p_scheditem & m_NsKinmuMData(w_i - 1).KinmuCD & m_NsKinmuMData(w_i - 1).BunruiCD
                End If
            Next w_i

            '�Ďv�l��
            m_AutoSched.p_test_cnt = General.paGetItemValue(General.G_STRMAINKEY2, w_RegPath, "TESTCOUNT", Convert.ToString(5), General.g_strHospitalCD)
            '�I������
            m_AutoSched.p_end_jyoken = General.paGetItemValue(General.G_STRMAINKEY2, w_RegPath, "ENDJYOUKENPOINT", Convert.ToString(30000), General.g_strHospitalCD)
            'I/F �t�@�C���p�X�^�t�@�C����
            m_AutoSched.p_SchedDataPath = General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, "DataPath", "") & "Schedif.dat"
            '���s�󋵕\���p��۸�ڽ�ް��޼ު��
            m_AutoSched.p_ObjProgressBar = prbSchedProcess
            '���s�󋵕\���p���ٵ�޼ު��
            m_AutoSched.p_ObjLabel = lblSchedMessage
            '�\�[�g�p ؽ��ޯ����޼ު��
            m_AutoSched.p_ObjListBox = Lst_SortList

            '��ݾ����ݐ����׸�
            m_CancelMode = True

            '���s����
            m_SchedExecute = True

            '�����ł͎����Эڰ��ݸ׽�̎��s�J�nҿ��ނ����s���܂��B�����è����������ݒ肳���
            '���Ȃ��Ɠ��삵�܂���B���Aҿ��ނ��I������܂Ő���͖߂�܂���B
            w_Rtn = m_AutoSched.mSchedStart()

            '��ݾ����ݐ���OFF
            m_CancelMode = False

            '�����V�~�����[�V�����̖߂�l����
            If w_Rtn = False Then
                '�����V�~�����[�V�����Ŏ��s���G���[�����������ꍇ

                '���s�׸ނ�������
                m_SchedExecute = False
                '�������ʔ��f�׸ނ�������
                m_SchedSaveFlg = False

                Me.Close()

            Else
                '�J�n�{�^�����g�p�\�Ƃ��܂��B
                cmdSchedStart.Enabled = True
            End If

            Call cmdSchedExit_Click(cmdSchedExit, New System.EventArgs())
        End If

        Exit Sub
cmdSchedStart_Click:
        Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub frmNSK0000HH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo Form_Load
        Const W_SUBNAME As String = "NSK0000HH Form_Load"

        Dim w_i As Short
        Dim w_Select As String
        Dim w_KinmuCnt As Short
        Dim w_Font As Font

        '�I���Ζ��i�[ڼ޽�؂�Key
        Const M_RegKey As String = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY2 & "\" & "NSK0020B"
        Const M_RegSubKey As String = "List_Check"
        Const M_RegValue As String = "Field"

        '������
        m_LoadError = False

        Call subSetCtlList()

        '�����Ζ��̑I��ؽĂ��擾���܂��B
        Call GetKinmuList()

        '�����Ζ���
        w_KinmuCnt = UBound(m_NsKinmuMData)

        '�e���׸ނ̏�����
        m_SchedExecute = False '���ޭ��ݸގ��s�׸�
        m_CancelMode = False '�I������
        m_SchedSaveFlg = False '�������ʔ��f�׸�

        If m_LoadError = False Then
            '�����J�n��
            m_SelectKikanNo = 0

            '�����ޯ���̑I�����Ԃɂ������̫�Ă̓��t��ύX
            If m_KikanStartDate = "0" Then
                w_Select = Convert.ToString(m_DisplayStartDate)
                m_SelectKikanNo = 1
            ElseIf CDbl(m_KikanStartDate) < m_DisplayStartDate Then
                w_Select = Convert.ToString(m_DisplayStartDate)
                m_SelectKikanNo = 2
            Else
                w_Select = m_KikanStartDate
                m_SelectKikanNo = 3
            End If

            imdDate.Text = Mid(w_Select, 1, 4) & "/" & Mid(w_Select, 5, 2) & "/" & Mid(w_Select, 7, 2)

            '��������ݷ��߼��
            For w_i = 1 To M_KINMU_NUM
                If w_i <= w_KinmuCnt Then
                    m_lstCmdKinmu(w_i - 1).Visible = True
                    m_lstCmdKinmu(w_i - 1).Text = m_NsKinmuMData(w_i - 1).KinmuName
                Else
                    Exit For
                End If
            Next w_i

            For w_i = 1 To UBound(m_NsKinmuMData)
                'ڼ޽�؂��I���Ζ����ǂ������擾
                w_Select = General.paGetSetting(M_RegKey, M_RegSubKey, M_RegValue & Str(w_i), "0")
                If w_Select <> "0" Then
                    m_NsKinmuMData(w_i - 1).OnOffFlg = True

                    If w_i <= M_KINMU_NUM Then
                        '�I��
                        m_lstCmdKinmu(w_i - 1).BackColor = ColorTranslator.FromOle(CDbl(M_BackColor))
                        w_Font = m_lstCmdKinmu(w_i - 1).Font
                        m_lstCmdKinmu(w_i - 1).Font = New Font(w_Font, FontStyle.Bold)
                    End If
                End If
            Next w_i
        End If

        '�X�N���[���o�[�A�I�v�V�����{�^���̐ݒ�
        Select Case w_KinmuCnt
            Case 0 To M_KINMU_NUM
                For w_i = M_KINMU_NUM To (w_KinmuCnt + 1) Step -1
                    m_lstCmdKinmu(w_i - 1).Visible = False
                Next w_i
                HscKinmu.Maximum = (0 + HscKinmu.LargeChange - 1)
                HscKinmu.Visible = False
            Case Else
                For w_i = 1 To M_KINMU_NUM
                    m_lstCmdKinmu(w_i - 1).Visible = True
                    m_lstCmdKinmu(w_i - 1).Enabled = True
                Next w_i
                HscKinmu.Maximum = (w_KinmuCnt - M_KINMU_NUM + HscKinmu.LargeChange - 1)
                HscKinmu.Visible = True
                HscKinmu.Enabled = True
        End Select


        '�����Эڰ��ݸ׽�̵�޼ު�āi�ݽ�ݽ�j���쐬���܂��B��޼ު�ĕϐ���
        'Ӽޭ�����قŐ錾���Ă��܂��B�ݽ�ݽ�̊J����̫�ѱ�۰�ގ��ɍs���Ă��܂��B
        m_AutoSched = New NsAid_NSK0020B.ClsAutoSched

        '�ڑ���޼ު�ēn��

        'Inatalltype�n��
        m_AutoSched.pInstallType = General.g_InstallType
        '�}�X�^�擾���i
        m_AutoSched.pGetMasterObj = General.g_objGetMaster

        '����޳�𒆉��ɔz�u���܂��B
        Me.StartPosition = FormStartPosition.CenterScreen

        Exit Sub
Form_Load:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub frmNSK0000HH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo Form_Unload
        Const W_SUBNAME As String = "NSK0000HH Form_Unload"

        '�����Эڰ��ݸ׽�̲ݽ�ݽ��j�����܂��B
        m_AutoSched = Nothing

        Exit Sub
Form_Unload:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub HscKinmu_Change(ByVal newScrollValue As Integer)
        On Error GoTo HscKinmu_Change
        Const W_SUBNAME As String = "Nskk001d  HscKinmu_Change"

        '�X�N���[���o�[�̍X�V
        Dim w_i As Short
        Dim w_Hsc_Cnt As Short
        Dim w_KinmuCnt As Short
        Dim w_Font As Font

        '�R�}���h�{�^���̂b�`�o�s�h�n�m�ݒ�

        w_KinmuCnt = UBound(m_NsKinmuMData)

        '�Ζ�
        w_Hsc_Cnt = newScrollValue
        For w_i = 1 To M_KINMU_NUM
            If w_i + w_Hsc_Cnt <= w_KinmuCnt Then
                m_lstCmdKinmu(w_i - 1).Text = m_NsKinmuMData(w_i + w_Hsc_Cnt - 1).KinmuName
                If m_NsKinmuMData(w_i + w_Hsc_Cnt - 1).OnOffFlg = True Then
                    m_lstCmdKinmu(w_i - 1).BackColor = ColorTranslator.FromOle(CDbl(M_BackColor))
                    w_Font = m_lstCmdKinmu(w_i - 1).Font
                    m_lstCmdKinmu(w_i - 1).Font = New Font(w_Font, FontStyle.Bold)
                Else
                    m_lstCmdKinmu(w_i - 1).BackColor = SystemColors.Control
                    w_Font = m_lstCmdKinmu(w_i - 1).Font
                    m_lstCmdKinmu(w_i - 1).Font = New Font(w_Font, FontStyle.Regular)
                End If
            Else
                m_lstCmdKinmu(w_i - 1).Text = ""
            End If
        Next w_i

        Exit Sub
HscKinmu_Change:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub HscKinmu_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscKinmu.Scroll
        HscKinmu_Change(eventArgs.NewValue)
    End Sub

    Private Sub subSetCtlList()
        m_lstCmdKinmu.Add(_cmdKinmu_0)
        m_lstCmdKinmu.Add(_cmdKinmu_1)
        m_lstCmdKinmu.Add(_cmdKinmu_2)
        m_lstCmdKinmu.Add(_cmdKinmu_3)
        m_lstCmdKinmu.Add(_cmdKinmu_4)
        m_lstCmdKinmu.Add(_cmdKinmu_5)
    End Sub
End Class