Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Imports FarPoint.Win.Spread

Friend Class frmNSK0000HC
    Inherits General.FormBase
    Public NSK0000H_sql As New clsNSK0000H_sql
    '=======================================================
    '   ��  ��  ��  ��
    '=======================================================
    '��گċΖ��ő�\����
    Private Const M_PARET_NUM As Short = 50 '�Ζ��E�x�݁E����p
    Private Const M_PARET_NUM_SET As Short = 25 '�Z�b�g�Ζ��p
	'�Ζ��L���ް� �J�n�� �ʒu
    Private Const M_KinmuData_Row As Integer = 3
    Private Const M_KinmuData_Row_ChgJisseki As Integer = 4 '�Ζ��ύX��ʂ��۰�ނ��ꂽ�ꍇ�̏C���\�s
    Private Const M_KinmuData_Col As Integer = 2
    '�Ζ����͕s�\�i�O�����сE�����ٓ��j
	Private m_MonthBefore_Fore As Integer '�����F
	Private m_MonthBefore_Back As Integer '�w�i�F
	'�v����ԊO��4�T����
	Private m_Jisseki4W_Back As Integer '�w�i�F
	'�\��ˎ��ѕύX�Ζ�
    Private m_Comp_Fore As Integer '�����F
    Private m_WeekEnd_Back As Integer
    Private m_WeekEndColorFlg As String '�y���w�i�F�t���O
    Private m_HolidayColorFlg As String '�j�x���w�i�F�t���O
	'-------------------------------------------------------
	'  ��ײ�ްĕϐ�
	'-------------------------------------------------------
	'�v��/���щ��Ӱ�ނ̎擾Ӱ��("�V�K"�C"�v��ύX"�C"�Ζ��ύX")
	Private m_Mode As String
    '��i/��i�\��Ӱ�ނ̎擾Ӱ��("��i�\��"�C"��i�\��")
	Private m_blnOneTwo As Boolean
	'�E���ް��J�n�s
	Private m_StaffStartRow As Integer
	'�ő�\���i��
	Private m_MaxShowLine As Short
	'�������\���s
	Private m_DutyData As Short
	'�Ζ��\��\���s
	Private m_KinmuPlan As Short
	'�Ζ����ѕ\���s
	Private m_KinmuJisseki As Short
	'�͏o�\���s
	Private m_AppliData As Short
    '�v��P�ʥ�\�����Ԃ̎擾(1:�S�T�^�P����  2:�S�T�^�P�����ȊO)
	Private m_DispKikan As String
	'�X�V�׸�
	Private m_KosinFlg As Boolean
	'O.K.���݉����׸�
	Private m_OKFlg As Boolean
	'�v���ް� �J�n/�I�� ��ԍ����
	Private m_KeikakuD_StartCol_Param As Integer
	Private m_KeikakuD_EndCol_Param As Integer
	'�e�v���ް� �J�n/�I�� ��ԍ����
	Private m_KeikakuD_StartCol As Integer
	Private m_KeikakuD_EndCol As Integer
	'�e���čs �ʒu
	Private m_CUR_ROW_Param As Integer
	'�e���۰�
    Private m_Control_Param As FarPoint.Win.Spread.FpSpread
	'�Ζ��p
	Private m_Kinmu() As KinmuM_Type
	Private m_KinmuCnt As Short
	'�x�ݗp
	Private m_Yasumi() As KinmuM_Type
	Private m_YasumiCnt As Short
	'����Ζ��p
	Private m_Tokusyu() As KinmuM_Type
	Private m_TokusyuCnt As Short
	'��x�擾�֘A�ϐ�
	Private m_DaikyuMsgFlg As Integer '��x�擾���̊m�F���b�Z�[�W��\�����邩(0:�\��,1:��\��)
	Private m_SundayDaikyuFlg As Integer '��x�擾�\���ɓ��j�����܂߂邩(0:�܂߂Ȃ�,1:�܂߂�)
	Private m_DaikyuAdvFlg As Integer '��x�̐�����\�ɂ��邩(0:���Ȃ�,1:����)
    Private m_SaturdayDaikyuFlg As Integer '��x�擾�\���ɓy�j�����܂߂邩(0:�܂߂Ȃ�,1:�܂߂�)
    Private m_DaikyuAdvThisMonthFlg As Integer '��x���蓖�������t���O(0:OFF,1:ON)
    Private m_OuenDispFlg As Integer '�����Ζ��敪�̃��W�I�{�^�����p���b�g�ɕ\�����邩(1:���Ȃ�,0:����)
    '�e���ɂ������ް��ɂ��Ă��׸ށi�z��Ƃ��Ď󂯎��j
	Private m_DataFlg As Object '�ް��׸ށi"0":�v���ް��C"1":�����ް��C���̑�:�ް��Ȃ��j
	Private m_KakuteiFlg As Object '�m���׸ށi"0":�Y�������m���ް��C"1":�������m���ް��j
	Private m_DataHideFlg As Object '�ް��׸ށi"0":�v���ް��C"1":�����ް��C���̑�:�ް��Ȃ��j
	Private m_KakuteiHideFlg As Object '�m���׸ށi"0":�Y�������m���ް��C"1":�������m���ް��j
    Private Const M_MenuKibouChk As Short = 4 '��]�Ζ��ւ̓��͌x��

    '2014/04/23 Saijo add start P-06979-------------------------------------------------------------------
    Private m_strKinmuEmSecondFlg As String '�Ζ��L���S�p�Q�����Ή��t���O(0�F�Ή����Ȃ��A1:�Ή�����)
    '2014/04/23 Saijo add end P-06979---------------------------------------------------------------------

    '2015/04/14 Bando Add Start ========================
    Private m_DispKinmuCd As String '��]���[�h���̕\���ΏۋΖ�CD
    '2015/04/14 Bando Add End   ========================

    '��x�ڍחp�\����
    Private Structure DaikyuDetail_Type
        Dim DaikyuDate As Integer '��x�擾��
        Dim DaikyuKinmuCD As String '��x�擾�Ζ��b�c
        Dim GetFlg As String '��x�擾�^�C�v(0:1����x�A1:0.5����x)
    End Structure

    Private Structure Daikyu_Type
        Dim HolDate As Integer
        Dim HolKinmuCD As String
        Dim DaikyuDetail() As DaikyuDetail_Type
        Dim GetKbn As String '��x�����ʃ^�C�v(0:1����,1:1.5����)
        Dim RemainderHol As Double '��x���g�p��
        Dim OutPutList As String '��x�擾���Ƀ��X�g�ɂ����邩�ǂ������׸�("0":��x�擾����Ώ� "1":��x�擾���Ώ�)
    End Structure

    Private m_DaikyuData() As Daikyu_Type
    Private M_StaffID As String
    Private m_Index As Integer
    Private Const M_YYYYMMDDLabel_Row As Integer = 2 '���t�B���Z���s
    Private Const M_PASTE As String = "1"
    Private Const M_DELETE As String = "2"
    Private Const M_SET As String = "3"
    Private m_DaikyuBackColorFlg As Boolean '��x�Ζ��ޯ��װ�׸�

    Private m_toolTipTxt As String '�c�[���`�b�v�\��������
    Private m_empRowDispFlg As Boolean '�̗p��\���t���O
    Private Structure NightShortInfo
        Dim Date_St As Integer
        Dim Date_Ed As Integer
    End Structure
    Private m_nightWorkInfo() As NightShortInfo '��ΐ�]
    Private m_shortWorkInfo() As NightShortInfo '�Z����

    '�Z�b�g�Ζ��p
    Private Structure SetKinmu_Type
        Dim Mark As String
        <VBFixedArray(10)> Dim CD() As String
        Dim StrText As String
        Dim KinmuCnt As Integer
        Dim blnKinmu As Boolean

        Public Sub Initialize()
            ReDim CD(10)
        End Sub
    End Structure

    Private m_SetKinmu() As SetKinmu_Type
    Private m_SetCnt As Short
    Private m_lngDaikyuPastPeriod As Integer '�ߋ��̑�x�擾���̋x���o�Γ��̗L���͈́i�����O�܂ł̋x���o�΂͗L�����Ċ����j
    Private m_FontSize As Short
    '̫�Ļ��ޒ萔
    Private Const M_FontSize_Big As Short = 14 '�u��v̫�Ļ���=14
    Private Const M_FontSize_Middle As Short = 12 '�u���v̫�Ļ���=12
    Private Const M_FontSize_Small As Short = 9 '�u���v̫�Ļ���=9
    '2014/04/23 Saijo add start P-06979-----------------------------------------
    Private Const M_FontSize_Second_Big As Short = 10 '�u��v̫�Ļ���=10
    Private Const M_FontSize_Second_Middle As Short = 9 '�u���v̫�Ļ���=9
    Private Const M_FontSize_Second_Small As Short = 7 '�u���v̫�Ļ���=7
    '2014/04/23 Saijo add end P-06979-------------------------------------------
    Private m_SpreadSize As Double
    '�����p������
    Private m_HolDateStr As String '�j�� �����p������
    Private m_OffDayStr As String '�x�� �����p������
    Private m_Daikyu15KinmuCD() As String '��x��1.5����������Ζ��b�c
    Private m_PackageFLG As Short '�p�b�P�[�W�}�X�^(0:�͏o�~�������~,1:�͏o�~��������,2:�͏o���������~,3:�͏o����������)
    Private m_StartDate As Integer
    Private m_EndDate As Integer
    Private m_strUpdKojyoDate As String

    Private m_lstCmdKinmu As New List(Of Object)
    Private m_lstCmdYasumi As New List(Of Object)
    Private m_lstCmdTokusyu As New List(Of Object)
    Private m_lstCmdSet As New List(Of Object)

    '�J�n���擾
    Public WriteOnly Property pStartDate() As Integer
        Set(ByVal Value As Integer)
            m_StartDate = Value
        End Set
    End Property

    '�I�����擾
    Public WriteOnly Property pEndDate() As Integer
        Set(ByVal Value As Integer)
            m_EndDate = Value
        End Set
    End Property

    Public WriteOnly Property pDataHideFlg() As Object
        Set(ByVal Value As Object)
            m_DataHideFlg = Value
            If IsArray(m_DataHideFlg) = False Then
                ReDim m_DataHideFlg(0)
            End If
        End Set
    End Property

    Public WriteOnly Property pKakuteiHideFlg() As Object
        Set(ByVal Value As Object)
            m_KakuteiHideFlg = Value
            If IsArray(m_KakuteiHideFlg) = False Then
                ReDim m_KakuteiHideFlg(0)
            End If
        End Set
    End Property

    Public WriteOnly Property pDataFlg() As Object
        Set(ByVal Value As Object)
            m_DataFlg = Value
            If IsArray(m_DataFlg) = False Then
                ReDim m_DataFlg(0)
            End If
        End Set
    End Property

    Public WriteOnly Property pKakuteiFlg() As Object
        Set(ByVal Value As Object)
            m_KakuteiFlg = Value
            If IsArray(m_KakuteiFlg) = False Then
                ReDim m_KakuteiFlg(0)
            End If
        End Set
    End Property

    '��ʗ����グӰ�ނ̎擾(1:�S�T�^�P����  2:�S�T�^�P�����ȊO)
    Public WriteOnly Property pDispKikan() As String
        Set(ByVal Value As String)
            m_DispKikan = Value
        End Set
    End Property

    '̫�Ļ���(̫�т̕���ݒ肷��Ƃ��Ɏg�p)
    Public WriteOnly Property pFontSize() As Short
        Set(ByVal Value As Short)
            m_FontSize = Value
        End Set
    End Property

    '̫�т̕�
    Public WriteOnly Property pSpreadSize() As Double
        Set(ByVal Value As Double)
            m_SpreadSize = Value
        End Set
    End Property

    '�v��/���щ��Ӱ�ނ̎擾Ӱ��(1:�v��  2:����)
    Public WriteOnly Property pMode() As String
        Set(ByVal Value As String)
            m_Mode = Value
        End Set
    End Property

    '��i/��i�\��Ӱ�ނ̎擾Ӱ��(True:��i  False:��i)
    Public WriteOnly Property pPlanTwo() As Boolean
        Set(ByVal Value As Boolean)
            m_blnOneTwo = Value
        End Set
    End Property

    '�E���ް��J�n�s
    Public WriteOnly Property pStaffStartRow() As Integer
        Set(ByVal Value As Integer)
            m_StaffStartRow = Value
        End Set
    End Property

    '�ő�\���i��
    Public WriteOnly Property pMaxShowLine(ByVal p_DutyData As Short, ByVal p_KinmuPlan As Short, ByVal p_KinmuJisseki As Short, ByVal p_AppliData As Short) As Short
        Set(ByVal Value As Short)
            m_DutyData = p_DutyData
            m_KinmuPlan = p_KinmuPlan
            m_KinmuJisseki = p_KinmuJisseki
            m_AppliData = p_AppliData
            m_MaxShowLine = Value
        End Set
    End Property

    '�c�[���`�b�v�e�L�X�g
    Public WriteOnly Property pToolTxt() As String
        Set(ByVal value As String)
            m_toolTipTxt = value
        End Set
    End Property

    '�̗p��̕\���t���O
    Public WriteOnly Property pEmpRowVisible() As Boolean
        Set(ByVal value As Boolean)
            m_empRowDispFlg = value
        End Set
    End Property

    '��ΐ�]�E�玙�Z���ԏ�����
    Public WriteOnly Property pInitNightShortInfo(ByVal p_ngt As Integer, ByVal p_shr As Integer) As Boolean
        Set(ByVal value As Boolean)
            ReDim m_nightWorkInfo(p_ngt)
            ReDim m_shortWorkInfo(p_ngt)
        End Set
    End Property

    '��ΐ�]���
    Public WriteOnly Property pNightWork(ByVal p_st As Integer, ByVal p_ed As Integer) As Boolean
        Set(ByVal value As Boolean)
            Dim idx As Integer = UBound(m_nightWorkInfo)
            m_nightWorkInfo(idx).Date_St = p_st
            m_nightWorkInfo(idx).Date_Ed = p_ed

            If m_StartDate > p_st Then m_nightWorkInfo(idx).Date_St = m_StartDate
            If m_EndDate < p_ed Then m_nightWorkInfo(idx).Date_Ed = m_EndDate
        End Set
    End Property

    '�Z���Ԏҏ��
    Public WriteOnly Property pShortWork(ByVal p_st As Integer, ByVal p_ed As Integer) As Boolean
        Set(ByVal value As Boolean)
            Dim idx As Integer = UBound(m_shortWorkInfo)
            m_shortWorkInfo(idx).Date_St = p_st
            m_shortWorkInfo(idx).Date_Ed = p_ed

            If m_StartDate > p_st Then m_shortWorkInfo(idx).Date_St = m_StartDate
            If m_EndDate < p_ed Then m_shortWorkInfo(idx).Date_Ed = m_EndDate
        End Set
    End Property

    '��x�f�[�^�󂯎��
    Public WriteOnly Property pDaikyuData(ByVal p_HolDate As Integer, ByVal p_HolKinmuCD As String, ByVal p_GetKbn As String, ByVal p_RemainderHol As Double, ByVal p_OutPutList As String, ByVal p_DaikyuDate As Integer, ByVal p_DaikyuKinmuCD As String) As String
        Set(ByVal Value As String)
            Dim w_SeachLoop As Integer
            Dim w_TargetIdx As Integer
            Dim w_SubTargetIdx As Integer

            If m_Index = 0 Then
                '�z��m��
                ReDim m_DaikyuData(0)
            End If

            '���Ɏ擾�ς݂̓��t���`�F�b�N
            w_TargetIdx = 0
            For w_SeachLoop = 1 To UBound(m_DaikyuData)
                If m_DaikyuData(w_SeachLoop).HolDate = p_HolDate Then
                    '�擾�ς݂̏ꍇ���ޯ�����Z�b�g
                    w_TargetIdx = w_SeachLoop
                End If
            Next w_SeachLoop

            If w_TargetIdx = 0 Then
                '�V�K�̏ꍇ�A�z��g��
                '���ޯ�����ı���
                m_Index = m_Index + 1
                w_TargetIdx = m_Index
                ReDim Preserve m_DaikyuData(w_TargetIdx)
                m_DaikyuData(w_TargetIdx).HolDate = p_HolDate
                m_DaikyuData(w_TargetIdx).HolKinmuCD = p_HolKinmuCD
                m_DaikyuData(w_TargetIdx).GetKbn = p_GetKbn
                m_DaikyuData(w_TargetIdx).RemainderHol = p_RemainderHol
                m_DaikyuData(w_TargetIdx).OutPutList = p_OutPutList

                w_SubTargetIdx = 1
            Else
                w_SubTargetIdx = UBound(m_DaikyuData(w_TargetIdx).DaikyuDetail) + 1
            End If

            ReDim Preserve m_DaikyuData(w_TargetIdx).DaikyuDetail(w_SubTargetIdx)
            m_DaikyuData(w_TargetIdx).DaikyuDetail(w_SubTargetIdx).DaikyuDate = p_DaikyuDate
            m_DaikyuData(w_TargetIdx).DaikyuDetail(w_SubTargetIdx).DaikyuKinmuCD = p_DaikyuKinmuCD
            m_DaikyuData(w_TargetIdx).DaikyuDetail(w_SubTargetIdx).GetFlg = Value
        End Set
    End Property

    '�E���Ǘ��ԍ����󂯎��
    Public WriteOnly Property pStaffID() As String
        Set(ByVal Value As String)
            M_StaffID = Value

            '��x�p�z�񏉊���
            ReDim m_DaikyuData(0)

            '��x�z����ޯ��������
            m_Index = 0
        End Set
    End Property

    '�X�V�׸�(True:�X�V,False:���X�V)
    Public ReadOnly Property pKosinFlg() As Boolean
        Get
            pKosinFlg = m_KosinFlg
        End Get
    End Property

    '��x�f�[�^�e��ʈ��n��
    Public Function pDaikyuDataGet(ByRef p_HolKinmuCD As String, ByRef p_OutPutList As String, ByRef p_GetKbn As String, ByRef p_RemainderHol As Double, ByRef p_DetailDataCnt As Integer, ByVal p_Int As Integer) As Integer

        pDaikyuDataGet = m_DaikyuData(p_Int).HolDate
        p_HolKinmuCD = m_DaikyuData(p_Int).HolKinmuCD


        p_OutPutList = m_DaikyuData(p_Int).OutPutList
        p_GetKbn = m_DaikyuData(p_Int).GetKbn
        p_RemainderHol = m_DaikyuData(p_Int).RemainderHol
        p_DetailDataCnt = UBound(m_DaikyuData(p_Int).DaikyuDetail)
    End Function

    Public Function pDaikyuDetailDataGet(ByRef p_DaikyuKinmuCD As String, ByRef p_GetFlg As String, ByVal p_Int As Integer, ByVal p_SubInt As Integer) As Integer
        pDaikyuDetailDataGet = m_DaikyuData(p_Int).DaikyuDetail(p_SubInt).DaikyuDate
        p_DaikyuKinmuCD = m_DaikyuData(p_Int).DaikyuDetail(p_SubInt).DaikyuKinmuCD
        p_GetFlg = m_DaikyuData(p_Int).DaikyuDetail(p_SubInt).GetFlg
    End Function

    '��x�f�[�^����
    Public ReadOnly Property pDaikyuCnt() As Integer
        Get
            pDaikyuCnt = UBound(m_DaikyuData)
        End Get
    End Property

    '�x�������擾
    Public WriteOnly Property pHolData(ByVal p_HolDateStr As String) As String
        Set(ByVal Value As String)
            m_HolDateStr = p_HolDateStr
            m_OffDayStr = Value
        End Set
    End Property

    '�v���ް� �J�n/�I�� ��ԍ������󂯎��
    Public WriteOnly Property pKeikakuD_EndCol() As Integer
        Set(ByVal Value As Integer)
            m_KeikakuD_EndCol = Value
        End Set
    End Property

    '�v���ް� �J�n/�I�� ��ԍ������󂯎��
    Public WriteOnly Property pKeikakuD_StartCol() As Integer
        Set(ByVal Value As Integer)
            m_KeikakuD_StartCol = Value
        End Set
    End Property

    '�e�v���ް� �J�n/�I�� ��ԍ������󂯎��
    Public WriteOnly Property pKeikakuD_KinmuDataEndCol() As Integer
        Set(ByVal Value As Integer)
            m_KeikakuD_EndCol_Param = Value
        End Set
    End Property

    '�e�v���ް� �J�n/�I�� ��ԍ������󂯎��
    Public WriteOnly Property pKeikakuD_KinmuDataStartCol() As Integer
        Set(ByVal Value As Integer)
            m_KeikakuD_StartCol_Param = Value
        End Set
    End Property

    '�e���čs �ʒu�����󂯎��
    Public WriteOnly Property pCUR_ROW() As Integer
        Set(ByVal Value As Integer)
            m_CUR_ROW_Param = Value
        End Set
    End Property

    '�e��ʂŕ\������Ă�����گ�޼�Ă��󂯎��
    Public WriteOnly Property pControl() As FarPoint.Win.Spread.FpSpread
        Set(ByVal Value As FarPoint.Win.Spread.FpSpread)
            m_Control_Param = Value
        End Set
    End Property

    Public ReadOnly Property pUpdKojyoDate() As String
        Get
            pUpdKojyoDate = m_strUpdKojyoDate
        End Get
    End Property

    '��x�`�F�b�N
    Private Function Check_Daikyu(ByVal p_Ivent As String, ByVal p_Date As Object, ByVal p_KinmuCD As String) As Boolean

        Const W_SUBNAME As String = "NSK0000HC Check_Daikyu"

        Dim w_frmDaikyu As Object
        Dim w_Int As Integer
        Dim w_HolDate As Integer
        Dim w_HolKinmuCD As String
        Dim w_SelDate As Integer
        Dim w_STS As Integer
        Dim w_strMsg() As String
        Dim w_indIdx As Short
        Dim w_lngLoop As Integer
        Dim w_lngLoop2 As Integer
        Dim w_SelDate2 As Integer
        Dim w_lngIdx As Integer
        Dim w_lngDataCnt As Integer
        Dim w_HalfDaikyuFlg As Boolean

        '�����l
        Check_Daikyu = False
        Try
            m_DaikyuBackColorFlg = False

            '�y�[�X�g�̏ꍇ
            If p_Ivent = "1" Then
                '�j���ł��邩�ǂ���
                If General.pafncDaikyuCheck(General.g_strHospitalCD, p_Date, General.g_strSelKinmuDeptCD) = True Or (m_SundayDaikyuFlg = 1 And Weekday(CDate(Format(Integer.Parse(p_Date), "0000/00/00"))) = 1) Or (m_SaturdayDaikyuFlg = 1 And Weekday(CDate(Format(Integer.Parse(p_Date), "0000/00/00"))) = 7) Then
                    If p_KinmuCD <> "" Then

                        '���łɂ���z��̋x���Ζ����ƑI�����ꂽ���t�œ����������邩
                        For w_Int = 1 To UBound(m_DaikyuData)
                            '�������łɂ���z��̒��ɂ�������A��x�擾�ς݂ł��邩�`�F�b�N
                            If m_DaikyuData(w_Int).HolDate = p_Date Then
                                For w_lngLoop = 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                    If m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate <> 0 Then
                                        ReDim w_strMsg(1)
                                        w_strMsg(1) = "��x�擾�ς݂̋Ζ��ł��B~n"
                                        Call General.paMsgDsp("NS0110", w_strMsg)
                                        Exit Function
                                    End If
                                Next w_lngLoop
                            End If
                        Next w_Int

                        Select Case g_KinmuM(CShort(p_KinmuCD)).DaikyuFlg
                            Case "1"
                                '�j���ɋΖ�����=�Ζ��̋Ζ�CD��\��t�������͑�x�f�[�^���쐬����
                                If m_DaikyuMsgFlg = 0 Then
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = ""
                                    w_STS = General.paMsgDsp("NS0111", w_strMsg)
                                Else
                                    '��\���̏ꍇ�A�K���擾
                                    w_STS = MsgBoxResult.Yes
                                End If

                                If w_STS = MsgBoxResult.Yes Then
                                    '��x���X�V
                                    Call UpDate_DaikyuData(p_Date, p_KinmuCD, 0)

                                    m_DaikyuBackColorFlg = True
                                Else
                                    '��x���X�V
                                    Call UpDate_DaikyuData(p_Date, p_KinmuCD, 1)
                                End If
                            Case "2"
                                If g_KinmuM(CShort(p_KinmuCD)).HolBunruiCD = General.G_STRDAIKYUBUNRUI Then
                                    '�j���ɑ�x��\��t�����ꍇ�A��x�͎擾�ł��Ȃ�
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "���̓��ɑ�x��"
                                    Call General.paMsgDsp("NS0112", w_strMsg)
                                    Exit Function
                                Else
                                    '��x���X�V
                                    Call UpDate_DaikyuData(p_Date, p_KinmuCD, 1)
                                End If
                        End Select
                    End If
                Else
                    '�j���łȂ��ꍇ
                    If g_KinmuM(CShort(p_KinmuCD)).HolBunruiCD = General.G_STRDAIKYUBUNRUI Then

                        w_frmDaikyu = New frmNSK0000HJ

                        w_frmDaikyu.pSelDate = Integer.Parse(p_Date)
                        w_frmDaikyu.pSelKinmuCD = p_KinmuCD
                        '�N�����i"0":�v�� ����ȊO:������ʁj
                        w_frmDaikyu.pKeikakuFlg = "0"
                        '��x�f�[�^��n��
                        For w_Int = 1 To UBound(m_DaikyuData)
                            '���X�g�Ώۃf�[�^�ł��邩
                            If m_DaikyuData(w_Int).OutPutList = "1" Then
                                If m_DaikyuAdvFlg = 0 Then
                                    '��x�L�����������邩�H
                                    If m_lngDaikyuPastPeriod = -1 Then
                                        '��x�L�������Ȃ�
                                        '�Ώۓ��t���ߋ��̃f�[�^���H
                                        If m_DaikyuData(w_Int).HolDate < Integer.Parse(p_Date) Then
                                            w_HolDate = m_DaikyuData(w_Int).HolDate
                                            w_HolKinmuCD = m_DaikyuData(w_Int).HolKinmuCD
                                            '�����n��
                                            w_frmDaikyu.pDaikyuData(w_HolDate, w_HolKinmuCD) = m_DaikyuData(w_Int).RemainderHol
                                        End If
                                    Else
                                        '��x�L����������
                                        '�Ώۓ��t���ߋ��̃f�[�^���w���(��̫��56��)�O�܂łł��邩
                                        If (m_DaikyuData(w_Int).HolDate < Integer.Parse(p_Date)) And (DateDiff(DateInterval.Day, CDate(Format(m_DaikyuData(w_Int).HolDate, "0000/00/00")), CDate(Format(Integer.Parse(p_Date), "0000/00/00"))) <= m_lngDaikyuPastPeriod) Then
                                            '�Ώۓ��t���ߋ��̃f�[�^�ł��邩
                                            w_HolDate = m_DaikyuData(w_Int).HolDate
                                            w_HolKinmuCD = m_DaikyuData(w_Int).HolKinmuCD
                                            '�����n��
                                            w_frmDaikyu.pDaikyuData(w_HolDate, w_HolKinmuCD) = m_DaikyuData(w_Int).RemainderHol
                                        End If
                                    End If
                                Else
                                    '�w����O��ł��邩�ǂ���
                                    '��x�L�����������邩�H
                                    If m_lngDaikyuPastPeriod = -1 Then
                                        If m_DaikyuAdvThisMonthFlg = 0 Then
                                            '��x�L�������Ȃ�
                                            w_HolDate = m_DaikyuData(w_Int).HolDate
                                            w_HolKinmuCD = m_DaikyuData(w_Int).HolKinmuCD
                                            '�����n��
                                            w_frmDaikyu.pDaikyuData(w_HolDate, w_HolKinmuCD) = m_DaikyuData(w_Int).RemainderHol
                                        Else
                                            If m_DaikyuData(w_Int).HolDate <= CDbl(Format(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(Integer.Parse(General.paLeft(p_Date, 6) & "01"), "0000/00/00")))), "yyyyMMdd")) Then
                                                w_HolDate = m_DaikyuData(w_Int).HolDate
                                                w_HolKinmuCD = m_DaikyuData(w_Int).HolKinmuCD
                                                '�����n��
                                                '��x�̖��g�p�����n��
                                                w_frmDaikyu.pDaikyuData(w_HolDate, w_HolKinmuCD) = m_DaikyuData(w_Int).RemainderHol
                                            End If
                                        End If
                                    Else
                                        If m_DaikyuAdvThisMonthFlg = 0 Then
                                            '��x�L����������
                                            '��x�L�������͈͓����H
                                            If (DateDiff(DateInterval.Day, CDate(Format(m_DaikyuData(w_Int).HolDate, "0000/00/00")), CDate(Format(Integer.Parse(p_Date), "0000/00/00"))) <= m_lngDaikyuPastPeriod) And (DateDiff(DateInterval.Day, CDate(Format(Integer.Parse(p_Date), "0000/00/00")), CDate(Format(m_DaikyuData(w_Int).HolDate, "0000/00/00"))) <= m_lngDaikyuPastPeriod) Then
                                                w_HolDate = m_DaikyuData(w_Int).HolDate
                                                w_HolKinmuCD = m_DaikyuData(w_Int).HolKinmuCD
                                                '�����n��
                                                w_frmDaikyu.pDaikyuData(w_HolDate, w_HolKinmuCD) = m_DaikyuData(w_Int).RemainderHol
                                            End If
                                        Else
                                            If DateDiff(DateInterval.Day, CDate(Format(m_DaikyuData(w_Int).HolDate, "0000/00/00")), CDate(Format(Integer.Parse(p_Date), "0000/00/00"))) <= m_lngDaikyuPastPeriod And m_DaikyuData(w_Int).HolDate <= CDbl(Format(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(Integer.Parse(General.paLeft(p_Date, 6) & "01"), "0000/00/00")))), "yyyyMMdd")) And DateDiff(DateInterval.Day, CDate(Format(Integer.Parse(p_Date), "0000/00/00")), CDate(Format(m_DaikyuData(w_Int).HolDate, "0000/00/00"))) <= m_lngDaikyuPastPeriod Then
                                                w_HolDate = m_DaikyuData(w_Int).HolDate
                                                w_HolKinmuCD = m_DaikyuData(w_Int).HolKinmuCD
                                                '�����n��
                                                '��x�̖��g�p�����n��
                                                w_frmDaikyu.pDaikyuData(w_HolDate, w_HolKinmuCD) = m_DaikyuData(w_Int).RemainderHol
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next w_Int

                        If w_frmDaikyu.mfncDaikyuDate_Check = False Then
                            ReDim w_strMsg(1)
                            w_strMsg(1) = "���̓��ɑ�x��"
                            Call General.paMsgDsp("NS0112", w_strMsg)
                            Exit Function
                        Else
                            '�j���ȊO�ɑ�x��\��t�����ꍇ�A��x�Ǘ��e���擾���A�擾�ł����x�̈ꗗ��ʂ��o�͂���B
                            w_frmDaikyu.ShowDialog(Me)
                            If w_frmDaikyu.pEndStatus = False Then
                                Exit Function
                            Else
                                'OK���݉�����
                                w_SelDate = w_frmDaikyu.pSelDate

                                '���� ������x�擾����2�ڂ̎擾�N�������擾
                                w_SelDate2 = w_frmDaikyu.pSelDate2
                                '������xDlg
                                w_HalfDaikyuFlg = w_frmDaikyu.pGetDaikyuType

                                '��x�̓��ɂ���ɑ�x��\��t�����ꍇ���̑�x�f�[�^���N���A
                                For w_Int = 1 To UBound(m_DaikyuData)
                                    w_lngDataCnt = UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                    For w_lngLoop = 1 To w_lngDataCnt
                                        If w_lngLoop > UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                            Exit For
                                        End If

                                        If m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = Integer.Parse(p_Date) Then
                                            '�N���A����̂ő�x���g�p���ɉ��Z
                                            '������x���`�F�b�N
                                            Select Case m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg
                                                Case "0"
                                                    '�P��
                                                    m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 1
                                                Case "1"
                                                    '����
                                                    m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 0.5
                                            End Select

                                            m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = 0
                                            m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuKinmuCD = ""
                                            m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg = ""
                                            m_DaikyuData(w_Int).OutPutList = "1"

                                            If w_lngLoop < UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                                w_indIdx = w_lngLoop
                                                For w_lngLoop2 = w_lngLoop + 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                                    m_DaikyuData(w_Int).DaikyuDetail(w_indIdx) = m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop2)
                                                    w_indIdx = w_indIdx + 1
                                                Next w_lngLoop2

                                                '�z�񒲐�
                                                ReDim Preserve m_DaikyuData(w_Int).DaikyuDetail(UBound(m_DaikyuData(w_Int).DaikyuDetail) - 1)
                                            End If
                                        End If
                                    Next w_lngLoop
                                Next w_Int

                                '�I�����ꂽ���t�̃f�[�^�ɑ�xCD�Ɠ��t���i�[
                                For w_Int = 1 To UBound(m_DaikyuData)
                                    If m_DaikyuData(w_Int).OutPutList = "1" Then
                                        If m_DaikyuData(w_Int).HolDate = w_SelDate Or m_DaikyuData(w_Int).HolDate = w_SelDate2 Then
                                            w_lngIdx = 1
                                            '�P���ڂɃf�[�^���͂����Ă��邩�`�F�b�N
                                            If m_DaikyuData(w_Int).DaikyuDetail(w_lngIdx).DaikyuDate <> 0 Then
                                                '�͂����Ă���ꍇ�A�z��g��
                                                w_lngIdx = UBound(m_DaikyuData(w_Int).DaikyuDetail) + 1
                                                ReDim Preserve m_DaikyuData(w_Int).DaikyuDetail(w_lngIdx)
                                            End If

                                            m_DaikyuData(w_Int).DaikyuDetail(w_lngIdx).DaikyuDate = Integer.Parse(p_Date)
                                            m_DaikyuData(w_Int).DaikyuDetail(w_lngIdx).DaikyuKinmuCD = p_KinmuCD

                                            '��x���g�p�����獷������
                                            If w_HalfDaikyuFlg = True Or w_SelDate2 <> 0 Then
                                                '������x�̏ꍇ
                                                m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol - 0.5
                                                m_DaikyuData(w_Int).DaikyuDetail(w_lngIdx).GetFlg = "1"
                                            Else
                                                '�P����x�̏ꍇ
                                                m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol - 1
                                                m_DaikyuData(w_Int).DaikyuDetail(w_lngIdx).GetFlg = "0"
                                            End If

                                            '��x�̖��g�p�����Ȃ��Ȃ����ꍇ
                                            If m_DaikyuData(w_Int).RemainderHol <= 0 Then
                                                m_DaikyuData(w_Int).OutPutList = "0"
                                            End If
                                        End If
                                    End If
                                Next w_Int
                            End If
                        End If
                    Else
                        '�j���ȊO�ɑ�x�ȊO�̋Ζ���\��t�����ꍇ�A���̓��̌��Ζ�����x�̏ꍇ,��x�Ǘ��e�̑�x�擾�N������NULL�ōX�V����
                        '��x���擾���Ă��邩�`�F�b�N���A�擾���Ă�����z�񂩂��x�f�[�^���폜
                        For w_Int = 1 To UBound(m_DaikyuData)
                            For w_lngLoop = 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                If w_lngLoop > UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                    Exit For
                                End If

                                If m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = Integer.Parse(p_Date) Then
                                    '�N���A����̂ő�x���g�p���ɉ��Z
                                    '������x���`�F�b�N
                                    Select Case m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg
                                        Case "0"
                                            '�P��
                                            m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 1
                                        Case "1"
                                            '����
                                            m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 0.5
                                    End Select

                                    m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = 0
                                    m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuKinmuCD = ""
                                    m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg = ""
                                    m_DaikyuData(w_Int).OutPutList = "1"

                                    If w_lngLoop < UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                        w_indIdx = w_lngLoop
                                        For w_lngLoop2 = w_lngLoop + 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                            m_DaikyuData(w_Int).DaikyuDetail(w_indIdx) = m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop2)
                                            w_indIdx = w_indIdx + 1
                                        Next w_lngLoop2
                                        '�z�񒲐�
                                        ReDim Preserve m_DaikyuData(w_Int).DaikyuDetail(UBound(m_DaikyuData(w_Int).DaikyuDetail) - 1)
                                    End If
                                End If
                            Next w_lngLoop
                        Next w_Int
                    End If
                End If
            ElseIf p_Ivent = "2" Then
                '�폜�̏ꍇ
                '�j���ł��邩�ǂ���
                If General.pafncDaikyuCheck(General.g_strHospitalCD, p_Date, General.g_strSelKinmuDeptCD) = True Then
                    '��x���擾���Ă���Ζ��łȂ����`�F�b�N
                    For w_Int = 1 To UBound(m_DaikyuData)
                        If m_DaikyuData(w_Int).HolDate = Integer.Parse(p_Date) Then
                            If m_DaikyuData(w_Int).OutPutList = "0" Then
                                ReDim w_strMsg(1)
                                w_strMsg(1) = "��x���擾���Ă���̂�"
                                Call General.paMsgDsp("NS0098", w_strMsg)

                                Exit Function
                            End If
                        End If
                    Next w_Int

                    '��x���X�V
                    Call UpDate_DaikyuData(p_Date, p_KinmuCD, 1)
                Else
                    '�j���łȂ��ꍇ
                    '��x���폜���Ă��邩�`�F�b�N���A�폜����ꍇ�͔z�񂩂��x�f�[�^���폜
                    For w_Int = 1 To UBound(m_DaikyuData)
                        For w_lngLoop = 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                            If w_lngLoop > UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                Exit For
                            End If

                            If m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = Integer.Parse(p_Date) Then
                                '�N���A����̂ő�x���g�p���ɉ��Z
                                '������x���`�F�b�N
                                Select Case m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg
                                    Case "0"
                                        '�P��
                                        m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 1
                                    Case "1"
                                        '����
                                        m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 0.5
                                End Select

                                m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = 0
                                m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuKinmuCD = ""
                                m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg = ""
                                m_DaikyuData(w_Int).OutPutList = "1"

                                If w_lngLoop < UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                    w_indIdx = w_lngLoop
                                    For w_lngLoop2 = w_lngLoop + 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                        m_DaikyuData(w_Int).DaikyuDetail(w_indIdx) = m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop2)
                                        w_indIdx = w_indIdx + 1
                                    Next w_lngLoop2

                                    '�z�񒲐�
                                    ReDim Preserve m_DaikyuData(w_Int).DaikyuDetail(UBound(m_DaikyuData(w_Int).DaikyuDetail) - 1)
                                End If
                            End If
                        Next w_lngLoop
                    Next w_Int
                End If
            ElseIf p_Ivent = "3" Then
                '�Z�b�g�Ζ��\��t����
                '��x���X�V
                If General.pafncDaikyuCheck(General.g_strHospitalCD, p_Date, General.g_strSelKinmuDeptCD) = True Then
                    Call UpDate_DaikyuData(p_Date, p_KinmuCD, 1)
                Else
                    For w_Int = 1 To UBound(m_DaikyuData)
                        For w_lngLoop = 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                            If w_lngLoop > UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                Exit For
                            End If

                            If m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = Integer.Parse(p_Date) Then
                                '�N���A����̂ő�x���g�p���ɉ��Z
                                '������x���`�F�b�N
                                Select Case m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg
                                    Case "0"
                                        '�P��
                                        m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 1
                                    Case "1"
                                        '����
                                        m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 0.5
                                End Select

                                m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = 0
                                m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuKinmuCD = ""
                                m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg = ""
                                m_DaikyuData(w_Int).OutPutList = "1"

                                If w_lngLoop < UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                    w_indIdx = w_lngLoop
                                    For w_lngLoop2 = w_lngLoop + 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                        m_DaikyuData(w_Int).DaikyuDetail(w_indIdx) = m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop2)
                                        w_indIdx = w_indIdx + 1
                                    Next w_lngLoop2

                                    '�z�񒲐�
                                    ReDim Preserve m_DaikyuData(w_Int).DaikyuDetail(UBound(m_DaikyuData(w_Int).DaikyuDetail) - 1)
                                End If
                            End If
                        Next w_lngLoop
                    Next w_Int
                End If
            End If

            Check_Daikyu = True

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    '��x���X�V
    Private Sub UpDate_DaikyuData(ByVal p_Date As Integer, ByVal p_KinmuCD As String, ByVal p_Mode As Integer)

        Const W_SUBNAME As String = "NSK0000HC UpDate_DaikyuData"

        Dim w_Int As Integer
        Dim w_Int2 As Integer
        Dim w_DataFlg As Boolean
        Dim w_Index As Integer
        Dim w_WorkTbl As Daikyu_Type
        Dim w_WorkTbl2() As Daikyu_Type
        Try
            If p_Mode = CDbl("0") Then

                w_DataFlg = False

                '��x�f�[�^�������[�v
                For w_Int = 1 To UBound(m_DaikyuData)
                    '�x���o�ΔN����������\��t�������t�ƈ�v���邩
                    If m_DaikyuData(w_Int).HolDate = p_Date Then

                        w_DataFlg = True

                        '�Ζ����ύX�ɂȂ��Ă��Ȃ���
                        If m_DaikyuData(w_Int).HolKinmuCD = p_KinmuCD Then
                            '��x�擾���ɂ��Ă��邩
                            If m_DaikyuData(w_Int).OutPutList = "0" Then
                                m_DaikyuData(w_Int).OutPutList = "1"
                                Exit For
                            End If
                        Else
                            m_DaikyuData(w_Int).HolKinmuCD = p_KinmuCD
                            m_DaikyuData(w_Int).OutPutList = "1"

                            '�擾�敪
                            '�Ƃ肠�����P���Ƃ��ăZ�b�g
                            m_DaikyuData(w_Int).GetKbn = "0"
                            m_DaikyuData(w_Int).RemainderHol = 1

                            '��x1.5���������ΏۋΖ����`�F�b�N
                            For w_Int2 = 1 To UBound(m_Daikyu15KinmuCD)
                                If p_KinmuCD = m_Daikyu15KinmuCD(w_Int2) Then
                                    m_DaikyuData(w_Int).GetKbn = "1"
                                    m_DaikyuData(w_Int).RemainderHol = 1.5
                                    Exit For
                                End If
                            Next w_Int2

                            Exit For
                        End If
                    End If
                Next w_Int

                '��v����f�[�^���Ȃ��ꍇ�͔z��ɒǉ�
                If w_DataFlg = False Then
                    '�z��g��
                    w_Index = UBound(m_DaikyuData) + 1
                    ReDim Preserve m_DaikyuData(w_Index)
                    ReDim m_DaikyuData(w_Index).DaikyuDetail(1)

                    m_DaikyuData(w_Index).HolDate = p_Date
                    m_DaikyuData(w_Index).HolKinmuCD = p_KinmuCD
                    m_DaikyuData(w_Index).OutPutList = "1"

                    '�擾�敪
                    '�Ƃ肠�����P���Ƃ��ăZ�b�g
                    m_DaikyuData(w_Index).GetKbn = "0"
                    m_DaikyuData(w_Index).RemainderHol = 1
                    '1.5���ΏۋΖ����`�F�b�N
                    For w_Int = 1 To UBound(m_Daikyu15KinmuCD)
                        If p_KinmuCD = m_Daikyu15KinmuCD(w_Int) Then
                            m_DaikyuData(w_Index).GetKbn = "1"
                            m_DaikyuData(w_Index).RemainderHol = 1.5
                            Exit For
                        End If
                    Next w_Int

                    '�x���o�Γ��Ń\�[�g
                    For w_Int = 1 To UBound(m_DaikyuData)
                        For w_Int2 = 1 To UBound(m_DaikyuData) - w_Int
                            If (m_DaikyuData(w_Int2).HolDate > m_DaikyuData(w_Int2 + 1).HolDate) Then
                                w_WorkTbl = m_DaikyuData(w_Int2)
                                m_DaikyuData(w_Int2) = m_DaikyuData(w_Int2 + 1)
                                m_DaikyuData(w_Int2 + 1) = w_WorkTbl
                            End If
                        Next w_Int2
                    Next w_Int
                End If
            ElseIf p_Mode = CDbl("1") Then
                '��x�f�[�^�������[�v
                For w_Int = 1 To UBound(m_DaikyuData)
                    '�x���o�ΔN����������\��t�������t�ƈ�v���邩
                    If m_DaikyuData(w_Int).HolDate = p_Date Then

                        '���[�N�e�[�u���Ƀf�[�^��ޔ�
                        ReDim w_WorkTbl2(UBound(m_DaikyuData))
                        For w_Int2 = 1 To UBound(m_DaikyuData)
                            w_WorkTbl2(w_Int2) = m_DaikyuData(w_Int2)
                        Next w_Int2

                        '��x�p�z�񏉊���
                        ReDim m_DaikyuData(0)

                        '��������I�����Ă���̂Ŕz�񂩂�폜
                        For w_Int2 = 1 To UBound(w_WorkTbl2)
                            If w_Int <> w_Int2 Then
                                w_Index = UBound(m_DaikyuData) + 1
                                ReDim Preserve m_DaikyuData(w_Index)
                                m_DaikyuData(w_Index) = w_WorkTbl2(w_Int2)
                            End If
                        Next w_Int2

                        Exit For
                    End If
                Next w_Int
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Public Sub ShowWindow()

        Const W_SUBNAME As String = "NSK0000HC  ShowWindow"

        Dim w_Size As Double
        '2018/09/21 K.I Add Start---------
        Dim w_Left As String
        Dim w_Top As String
        '2018/09/21 K.I Add Start---------

        Try
            '̫�т̻��ނ𒲐�
            '2014/04/23 Saijo upd start P-06979----------------------------
            'w_Size = m_SpreadSize + General.paTwipsTopixels(500)
            'If w_Size < General.paTwipsTopixels(9750) Then
            '    w_Size = General.paTwipsTopixels(9750)
            'Else
            '    '�t�H���g���̂Ƃ��̓T�C�Y�Œ�
            '    If m_FontSize = M_FontSize_Small Then
            '       w_Size = General.paTwipsTopixels(12000)
            '    End If
            'End If
            If m_strKinmuEmSecondFlg = "0" Then
                w_Size = m_SpreadSize + General.paTwipsTopixels(500)
                If w_Size < General.paTwipsTopixels(9750) Then
                    w_Size = General.paTwipsTopixels(9750)
                Else
                    '�t�H���g���̂Ƃ��̓T�C�Y�Œ�
                    If m_FontSize = M_FontSize_Small Then
                        w_Size = General.paTwipsTopixels(12000)
                    End If
                End If
            Else
                If m_FontSize = M_FontSize_Second_Big Then
                    w_Size = m_SpreadSize + General.paTwipsTopixels(700)
                Else
                    w_Size = 960.0 + General.paTwipsTopixels(2700)
                End If
            End If

            '2014/04/23 Saijo upd end P-06979------------------------------

            Me.Width = w_Size

            Call subSetCtlList()

            '2014/04/23 Saijo upd start P-06979----------------------------
            '�Ζ��L���S�p�Q�����Ή��̃��C�A�E�g�ύX
            Call SetKinmuSecondView()
            '2014/04/23 Saijo upd end P-06979------------------------------

            '-----------��گ� �ݒ�-----------
            'NSKINMUNAMEM �擾
            Call GetKinmuName()

            '�p�l���E�B���h�E�ɋL�����Z�b�g
            Call SetKinmuData()
            '--------------------------------

            '�����S��ICON�ݒ�
            cmdErase.Image = Image.FromFile(g_ImagePath & G_ERASER_ICO)

            '��ʾ���ݸ�
            Me.StartPosition = FormStartPosition.CenterScreen

            '����޳�̕\���߼޼�݂�ݒ肷��
            '2018/09/21 K.I Upd Start------------------------------------------------------------------------------------------
            '���W�X�g���擾���폜
            'Call General.paGetWindowPositon(Me, General.G_STRMAINKEY2 & "\" & g_AppName)
            '��ʒ���
            w_Left = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - CSng(Me.Width)) / 2)
            w_Top = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - CSng(Me.Height)) / 2)
            Me.SetBounds((CSng(w_Left)), (CSng(w_Top)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
            '2018/09/21 K.I Upd End--------------------------------------------------------------------------------------------
            '����޳�̕\���߼޼�݂�ݒ肷��

            '���پوړ�
            Call SetStartCursol()

            Me.ShowDialog(pProcessObj)

        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub GetKinmuName()
        Const W_SUBNAME As String = "NSK0000HC  GetKinmuName"

        Dim w_Int As Short
        Dim w_Int2 As Short
        Dim w_DataCnt As Short
        Dim w_Sql As String
        'DAO��޼ު��
        Dim w_Rs As ADODB.Recordset
        '̨����
        Dim w_�L��_F As ADODB.Field
        Dim w_�Ζ�CD1_F As ADODB.Field
        Dim w_�Ζ�CD2_F As ADODB.Field
        Dim w_�Ζ�CD3_F As ADODB.Field
        Dim w_�Ζ�CD4_F As ADODB.Field
        Dim w_�Ζ�CD5_F As ADODB.Field
        Dim w_�Ζ�CD6_F As ADODB.Field
        Dim w_�Ζ�CD7_F As ADODB.Field
        Dim w_�Ζ�CD8_F As ADODB.Field
        Dim w_�Ζ�CD9_F As ADODB.Field
        Dim w_�Ζ�CD10_F As ADODB.Field
        Dim w_KinmuCnt As Integer
        Dim w_Int3 As Integer
        Dim w_lngEffEndDate As Integer
        Dim w_strTaihiCD() As String
        Dim w_blnEndDate As Boolean
        Dim w_strKinmuCD1 As String
        Dim w_strKinmuCD2 As String
        Dim w_strKinmuCD3 As String
        Dim w_strKinmuCD4 As String
        Dim w_strKinmuCD5 As String
        Dim w_strKinmuCD6 As String
        Dim w_strKinmuCD7 As String
        Dim w_strKinmuCD8 As String
        Dim w_strKinmuCD9 As String
        Dim w_strKinmuCD10 As String
        Dim w_strKinmuBunruiCD As String
        Try
            '������
            ReDim w_strTaihiCD(0)
            w_Int2 = 1

            '�Ζ����̃}�X�^�擾(�Ζ�����)
            With General.g_objGetMaster
                .pHospitalCD = General.g_strHospitalCD '�{�݃R�[�h
                .pKN_GetKbn = 1 '0:�S�� 1:�w��Ζ�����
                .pKN_KinmuDeptCD = General.g_strSelKinmuDeptCD '�I���Ζ�����

                If .mGetKinmuNameM = False Then
                Else
                    '�}�X�^����
                    w_DataCnt = .fKN_KinmuCount

                    For w_Int = 1 To w_DataCnt

                        '����
                        .mKN_KinmuIdx = w_Int

                        .pKN_GetKbn = 0
                        w_lngEffEndDate = .fKN_EffToDate
                        .pKN_GetKbn = 1

                        If w_lngEffEndDate >= m_StartDate Or w_lngEffEndDate = 0 Or w_lngEffEndDate = 99999999 Then

                            '���Ζ����ރR�[�h
                            w_strKinmuBunruiCD = .fKN_KinmuBunruiCD

                            If w_strKinmuBunruiCD = "1" Then
                                '-- �Ζ� --
                                '2015/04/14 Bando Upd Start ============================
                                '��]���[�h�̏ꍇ�A�\���ΏۋΖ��̂݃p���b�g�ɕ\��
                                'If g_HopeMode = 1 Then
                                If g_HopeMode = 1 AndAlso m_DispKinmuCd <> "" Then '2015/06/02 Bando Chg
                                    If InStr(m_DispKinmuCd, .fKN_KinmuCD) > 0 Then
                                        m_KinmuCnt = m_KinmuCnt + 1
                                        'NS_KINMUNAME_M�i�[�p�ϐ��̍Ē�`
                                        ReDim Preserve m_Kinmu(m_KinmuCnt)

                                        m_Kinmu(m_KinmuCnt - 1).CD = .fKN_KinmuCD
                                        m_Kinmu(m_KinmuCnt - 1).KinmuName = .fKN_Name
                                        m_Kinmu(m_KinmuCnt - 1).Mark = .fKN_MarkF
                                        m_Kinmu(m_KinmuCnt - 1).KBunruiCD = w_strKinmuBunruiCD
                                        m_Kinmu(m_KinmuCnt - 1).Setumei = .fKN_KinmuExplan
                                    End If
                                Else
                                    m_KinmuCnt = m_KinmuCnt + 1
                                    'NS_KINMUNAME_M�i�[�p�ϐ��̍Ē�`
                                    ReDim Preserve m_Kinmu(m_KinmuCnt)

                                    m_Kinmu(m_KinmuCnt - 1).CD = .fKN_KinmuCD
                                    m_Kinmu(m_KinmuCnt - 1).KinmuName = .fKN_Name
                                    m_Kinmu(m_KinmuCnt - 1).Mark = .fKN_MarkF
                                    m_Kinmu(m_KinmuCnt - 1).KBunruiCD = w_strKinmuBunruiCD
                                    m_Kinmu(m_KinmuCnt - 1).Setumei = .fKN_KinmuExplan
                                End If
                                '2015/04/14 Bando Upd End   ============================

                            ElseIf w_strKinmuBunruiCD = "2" Then
                                '-- �x�� --
                                '2015/04/14 Bando Upd Start ============================
                                '��]���[�h�̏ꍇ�A�\���ΏۋΖ��̂݃p���b�g�ɕ\��
                                'If g_HopeMode = 1 Then
                                If g_HopeMode = 1 AndAlso m_DispKinmuCd <> "" Then '2015/06/02 Bando Chg
                                    If InStr(m_DispKinmuCd, .fKN_KinmuCD) > 0 Then
                                        m_YasumiCnt = m_YasumiCnt + 1
                                        ReDim Preserve m_Yasumi(m_YasumiCnt)

                                        m_Yasumi(m_YasumiCnt - 1).CD = .fKN_KinmuCD
                                        m_Yasumi(m_YasumiCnt - 1).KinmuName = .fKN_Name
                                        m_Yasumi(m_YasumiCnt - 1).Mark = .fKN_MarkF
                                        m_Yasumi(m_YasumiCnt - 1).KBunruiCD = w_strKinmuBunruiCD
                                        m_Yasumi(m_YasumiCnt - 1).Setumei = .fKN_KinmuExplan
                                    End If
                                Else
                                    m_YasumiCnt = m_YasumiCnt + 1
                                    ReDim Preserve m_Yasumi(m_YasumiCnt)

                                    m_Yasumi(m_YasumiCnt - 1).CD = .fKN_KinmuCD
                                    m_Yasumi(m_YasumiCnt - 1).KinmuName = .fKN_Name
                                    m_Yasumi(m_YasumiCnt - 1).Mark = .fKN_MarkF
                                    m_Yasumi(m_YasumiCnt - 1).KBunruiCD = w_strKinmuBunruiCD
                                    m_Yasumi(m_YasumiCnt - 1).Setumei = .fKN_KinmuExplan
                                End If
                                '2015/04/14 Bando Upd End   ============================

                            ElseIf w_strKinmuBunruiCD = "3" Then
                                '-- ���� --
                                '2015/04/14 Bando Upd Start ============================
                                '��]���[�h�̏ꍇ�A�\���ΏۋΖ��̂݃p���b�g�ɕ\��
                                'If g_HopeMode = 1 Then
                                If g_HopeMode = 1 AndAlso m_DispKinmuCd <> "" Then '2015/06/02 Bando Chg
                                    If InStr(m_DispKinmuCd, .fKN_KinmuCD) > 0 Then
                                        m_TokusyuCnt = m_TokusyuCnt + 1
                                        ReDim Preserve m_Tokusyu(m_TokusyuCnt)

                                        m_Tokusyu(m_TokusyuCnt - 1).CD = .fKN_KinmuCD
                                        m_Tokusyu(m_TokusyuCnt - 1).KinmuName = .fKN_Name
                                        m_Tokusyu(m_TokusyuCnt - 1).Mark = .fKN_MarkF
                                        m_Tokusyu(m_TokusyuCnt - 1).KBunruiCD = w_strKinmuBunruiCD
                                        m_Tokusyu(m_TokusyuCnt - 1).Setumei = .fKN_KinmuExplan
                                    End If
                                Else
                                    m_TokusyuCnt = m_TokusyuCnt + 1
                                    ReDim Preserve m_Tokusyu(m_TokusyuCnt)

                                    m_Tokusyu(m_TokusyuCnt - 1).CD = .fKN_KinmuCD
                                    m_Tokusyu(m_TokusyuCnt - 1).KinmuName = .fKN_Name
                                    m_Tokusyu(m_TokusyuCnt - 1).Mark = .fKN_MarkF
                                    m_Tokusyu(m_TokusyuCnt - 1).KBunruiCD = w_strKinmuBunruiCD
                                    m_Tokusyu(m_TokusyuCnt - 1).Setumei = .fKN_KinmuExplan
                                End If
                                '2015/04/14 Bando Upd End   ============================

                            End If
                        Else
                            '2015/04/14 Bando Upd Start ============================
                            '��]���[�h�̏ꍇ�A�\���ΏۋΖ��̂݃p���b�g�ɕ\��
                            'If g_HopeMode = 1 Then
                            If g_HopeMode = 1 AndAlso m_DispKinmuCd <> "" Then '2015/06/02 Bando Chg
                                If InStr(m_DispKinmuCd, .fKN_KinmuCD) > 0 Then
                                    ReDim Preserve w_strTaihiCD(w_Int2)
                                    w_strTaihiCD(w_Int2) = .fKN_KinmuCD
                                    w_Int2 = w_Int2 + 1
                                End If
                            Else
                                ReDim Preserve w_strTaihiCD(w_Int2)
                                w_strTaihiCD(w_Int2) = .fKN_KinmuCD
                                w_Int2 = w_Int2 + 1
                            End If
                            '2015/04/14 Bando Upd End   ============================
                        End If
                    Next w_Int
                End If
            End With

            '2015/06/02 Bando Upd Start =====================================
            If g_HopeMode <> 1 Then
                '�Z�b�g�Ζ�
                '2017/05/02 Christopher Upd Start
                ''SQL���ҏW
                'w_Sql = "SELECT * FROM NS_SETKINMU_M "
                'w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                'w_Sql = w_Sql & "AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                'w_Sql = w_Sql & "ORDER BY DISPNO "

                'w_Rs = General.paDBRecordSetOpen(w_Sql)
                '<1>
                Call NSK0000H_sql.select_NS_SETKINMU_M_01(w_Rs)
                'Upd End
                '�Z�b�g�Ζ��z�񏉊���
                ReDim m_SetKinmu(0)

                If w_Rs.RecordCount <= 0 Then
                    m_SetCnt = 0
                Else
                    w_Int3 = 0

                    With w_Rs
                        .MoveLast()
                        w_DataCnt = .RecordCount
                        .MoveFirst()

                        ReDim m_SetKinmu(w_DataCnt)
                        m_SetCnt = w_DataCnt

                        w_�L��_F = .Fields("SetMark")
                        w_�Ζ�CD1_F = .Fields("KinmuCD1")
                        w_�Ζ�CD2_F = .Fields("KinmuCD2")
                        w_�Ζ�CD3_F = .Fields("KinmuCD3")
                        w_�Ζ�CD4_F = .Fields("KinmuCD4")
                        w_�Ζ�CD5_F = .Fields("KinmuCD5")
                        w_�Ζ�CD6_F = .Fields("KinmuCD6")
                        w_�Ζ�CD7_F = .Fields("KinmuCD7")
                        w_�Ζ�CD8_F = .Fields("KinmuCD8")
                        w_�Ζ�CD9_F = .Fields("KinmuCD9")
                        w_�Ζ�CD10_F = .Fields("KinmuCD10")

                        For w_Int = 1 To w_DataCnt

                            w_blnEndDate = True
                            w_strKinmuCD1 = w_�Ζ�CD1_F.Value & ""
                            w_strKinmuCD2 = w_�Ζ�CD2_F.Value & ""
                            w_strKinmuCD3 = w_�Ζ�CD3_F.Value & ""
                            w_strKinmuCD4 = w_�Ζ�CD4_F.Value & ""
                            w_strKinmuCD5 = w_�Ζ�CD5_F.Value & ""
                            w_strKinmuCD6 = w_�Ζ�CD6_F.Value & ""
                            w_strKinmuCD7 = w_�Ζ�CD7_F.Value & ""
                            w_strKinmuCD8 = w_�Ζ�CD8_F.Value & ""
                            w_strKinmuCD9 = w_�Ζ�CD9_F.Value & ""
                            w_strKinmuCD10 = w_�Ζ�CD10_F.Value & ""

                            '�Ζ�CD1�`10�܂łőޔ����Ă����Ζ�CD�ƈ�v���Ă�����̂͊����؂�
                            For w_Int2 = 1 To UBound(w_strTaihiCD)
                                Select Case w_strTaihiCD(w_Int2)
                                    Case w_strKinmuCD1
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD2
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD3
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD4
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD5
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD6
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD7
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD8
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD9
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD10
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                End Select
                            Next w_Int2

                            If w_blnEndDate = True Then
                                m_SetKinmu(w_Int3).Initialize()
                                m_SetKinmu(w_Int3).Mark = w_�L��_F.Value & ""
                                m_SetKinmu(w_Int3).CD(1) = w_�Ζ�CD1_F.Value & ""
                                m_SetKinmu(w_Int3).CD(2) = w_�Ζ�CD2_F.Value & ""
                                m_SetKinmu(w_Int3).CD(3) = w_�Ζ�CD3_F.Value & ""
                                m_SetKinmu(w_Int3).CD(4) = w_�Ζ�CD4_F.Value & ""
                                m_SetKinmu(w_Int3).CD(5) = w_�Ζ�CD5_F.Value & ""
                                m_SetKinmu(w_Int3).CD(6) = w_�Ζ�CD6_F.Value & ""
                                m_SetKinmu(w_Int3).CD(7) = w_�Ζ�CD7_F.Value & ""
                                m_SetKinmu(w_Int3).CD(8) = w_�Ζ�CD8_F.Value & ""
                                m_SetKinmu(w_Int3).CD(9) = w_�Ζ�CD9_F.Value & ""
                                m_SetKinmu(w_Int3).CD(10) = w_�Ζ�CD10_F.Value & ""
                                m_SetKinmu(w_Int3).blnKinmu = True

                                '�Ζ����������邩(�Ԃɋ󔒂͂Ȃ����̂Ƃ���)
                                w_KinmuCnt = 0

                                For w_Int2 = 1 To 10
                                    If m_SetKinmu(w_Int3).CD(w_Int2) <> "" Then
                                        w_KinmuCnt = w_KinmuCnt + 1
                                    Else
                                        Exit For
                                    End If
                                Next w_Int2

                                m_SetKinmu(w_Int3).KinmuCnt = w_KinmuCnt
                                w_Int3 = w_Int3 + 1
                            End If
                            .MoveNext()
                        Next w_Int
                    End With
                End If
                '2015/06/02 Bando Upd End   =====================================

                w_Rs.Close()
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub SetKinmuData()

        Const W_SUBNAME As String = "NSK0000HC  SetKinmuData"

        Dim w_i As Short
        Try
            '�R�}���h�{�^���̂b�`�o�s�h�n�m�ݒ�
            '�Ζ�
            For w_i = 1 To M_PARET_NUM
                If w_i <= m_KinmuCnt Then
                    m_lstCmdKinmu(w_i - 1).Text = m_Kinmu(w_i - 1).Mark
                    If m_Kinmu(w_i - 1).Setumei = "" Then
                        ToolTip1.SetToolTip(m_lstCmdKinmu(w_i - 1), Get_KinmuTipText(m_Kinmu(w_i - 1).CD))
                    Else
                        ToolTip1.SetToolTip(m_lstCmdKinmu(w_i - 1), Get_KinmuTipText(m_Kinmu(w_i - 1).CD) & "�F" & m_Kinmu(w_i - 1).Setumei)
                    End If
                Else
                    Exit For
                End If
            Next w_i

            '�x��
            For w_i = 1 To M_PARET_NUM
                If w_i <= m_YasumiCnt Then
                    m_lstCmdYasumi(w_i - 1).Text = m_Yasumi(w_i - 1).Mark
                    If m_Yasumi(w_i - 1).Setumei = "" Then
                        ToolTip1.SetToolTip(m_lstCmdYasumi(w_i - 1), Get_KinmuTipText(m_Yasumi(w_i - 1).CD))
                    Else
                        ToolTip1.SetToolTip(m_lstCmdYasumi(w_i - 1), Get_KinmuTipText(m_Yasumi(w_i - 1).CD) & "�F" & m_Yasumi(w_i - 1).Setumei)
                    End If
                Else
                    Exit For
                End If
            Next w_i

            '����Ζ�
            For w_i = 1 To M_PARET_NUM
                If w_i <= m_TokusyuCnt Then
                    m_lstCmdTokusyu(w_i - 1).Text = m_Tokusyu(w_i - 1).Mark
                    If m_Tokusyu(w_i - 1).Setumei = "" Then
                        ToolTip1.SetToolTip(m_lstCmdTokusyu(w_i - 1), Get_KinmuTipText(m_Tokusyu(w_i - 1).CD))
                    Else
                        ToolTip1.SetToolTip(m_lstCmdTokusyu(w_i - 1), Get_KinmuTipText(m_Tokusyu(w_i - 1).CD) & "�F" & m_Tokusyu(w_i - 1).Setumei)
                    End If
                Else
                    Exit For
                End If
            Next w_i

            '�Z�b�g�Ζ�
            For w_i = 1 To M_PARET_NUM_SET
                If w_i <= m_SetCnt Then
                    If m_SetKinmu(w_i - 1).blnKinmu = True Then
                        m_lstCmdSet(w_i - 1).Text = m_SetKinmu(w_i - 1).Mark
                        ToolTip1.SetToolTip(m_lstCmdSet(w_i - 1), Get_SetKinmuTipText(w_i - 1))
                    End If
                Else
                    Exit For
                End If
            Next w_i

            '�X�N���[���o�[�A�I�v�V�����{�^���̐ݒ�
            '�Ζ�
            Select Case m_KinmuCnt
                Case 0 To M_PARET_NUM
                    For w_i = M_PARET_NUM To (m_KinmuCnt + 1) Step -1
                        m_lstCmdKinmu(w_i - 1).Visible = False
                    Next w_i

                    HscKinmu.Maximum = (0 + HscKinmu.LargeChange - 1)
                    HscKinmu.Visible = False
                Case Else
                    For w_i = 1 To M_PARET_NUM
                        m_lstCmdKinmu(w_i - 1).Visible = True
                        m_lstCmdKinmu(w_i - 1).Enabled = True
                    Next w_i

                    HscKinmu.Maximum = (General.paRoundUp((m_KinmuCnt - M_PARET_NUM) / 2, 0) + HscKinmu.LargeChange - 1)
                    HscKinmu.Visible = True
                    HscKinmu.Enabled = True
            End Select

            '�x��
            Select Case m_YasumiCnt
                Case 0 To M_PARET_NUM
                    For w_i = M_PARET_NUM To (m_YasumiCnt + 1) Step -1
                        m_lstCmdYasumi(w_i - 1).Visible = False
                    Next w_i

                    HscYasumi.Maximum = (0 + HscYasumi.LargeChange - 1)
                    HscYasumi.Visible = False
                Case Else
                    For w_i = 1 To M_PARET_NUM
                        m_lstCmdYasumi(w_i - 1).Visible = True
                        m_lstCmdYasumi(w_i - 1).Enabled = True
                    Next w_i

                    HscYasumi.Maximum = (General.paRoundUp((m_YasumiCnt - M_PARET_NUM) / 2, 0) + HscYasumi.LargeChange - 1)
                    HscYasumi.Visible = True
                    HscYasumi.Enabled = True
            End Select

            '����Ζ�
            Select Case m_TokusyuCnt
                Case 0 To M_PARET_NUM
                    For w_i = M_PARET_NUM To (m_TokusyuCnt + 1) Step -1
                        m_lstCmdTokusyu(w_i - 1).Visible = False
                    Next w_i

                    HscTokusyu.Maximum = (0 + HscTokusyu.LargeChange - 1)
                    HscTokusyu.Visible = False
                Case Else
                    For w_i = 1 To M_PARET_NUM
                        m_lstCmdTokusyu(w_i - 1).Visible = True
                        m_lstCmdTokusyu(w_i - 1).Enabled = True
                    Next w_i

                    HscTokusyu.Maximum = (General.paRoundUp((m_TokusyuCnt - M_PARET_NUM) / 2, 0) + HscTokusyu.LargeChange - 1)
                    HscTokusyu.Visible = True
                    HscTokusyu.Enabled = True
            End Select

            '�Z�b�g�Ζ�
            Select Case m_SetCnt
                Case 0 To M_PARET_NUM_SET
                    For w_i = M_PARET_NUM_SET To (m_SetCnt + 1) Step -1
                        m_lstCmdSet(w_i - 1).Visible = False
                    Next w_i

                    HscSet.Maximum = (0 + HscSet.LargeChange - 1)
                    HscSet.Visible = False
                Case Else
                    For w_i = 1 To M_PARET_NUM_SET
                        m_lstCmdSet(w_i - 1).Visible = True
                        m_lstCmdSet(w_i - 1).Enabled = True
                    Next w_i

                    HscSet.Maximum = (m_SetCnt - M_PARET_NUM_SET + HscSet.LargeChange - 1)
                    HscSet.Visible = True
                    HscSet.Enabled = True
            End Select

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '�Z�b�g�Ζ��c�[���`�b�v�p������擾
    Public Function Get_SetKinmuTipText(ByVal p_Int As Integer) As String

        Const W_SUBNAME As String = "NSK0000HC Get_SetKinmuTipText"

        Dim w_str As String = String.Empty
        Dim w_strTEXT As String = String.Empty
        Dim w_Cnt As Integer
        Dim w_CD As String
        Try
            For w_Cnt = 1 To 10
                '�Ζ�CD���擾
                w_CD = m_SetKinmu(p_Int).CD(w_Cnt)

                '�󔒂łȂ����l�ł���ꍇ
                If w_CD <> "" And IsNumeric(w_CD) = True Then
                    If w_str <> "" Then
                        w_str = w_str & "-" & g_KinmuM(CShort(w_CD)).KinmuName
                    Else
                        w_str = g_KinmuM(CShort(w_CD)).KinmuName
                    End If

                    w_strTEXT = w_strTEXT & g_KinmuM(CShort(w_CD)).Mark
                Else
                    w_strTEXT = w_strTEXT & Space(2)
                End If
            Next w_Cnt

            m_SetKinmu(p_Int).StrText = w_strTEXT

            Get_SetKinmuTipText = w_str

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    Private Function Disp_Ouen(ByRef p_CD As String) As Boolean

        Const W_SUBNAME As String = "NSK0000HC  Disp_Ouen"

        Dim w_Form As frmNSK0000HK
        Dim w_strMsg() As String

        '������
        Disp_Ouen = False
        Try
            w_Form = New frmNSK0000HK
            w_Form.frmNSK0000HK_Load()
            If w_Form.pKangoCDFlg = True Then
                '�\��
                w_Form.ShowDialog(Me)

                'OK�������ݎ��̂ݏ������s
                If w_Form.pOKFlg = True Then
                    p_CD = w_Form.pSelKangoTCD
                    Disp_Ouen = True
                End If
            Else
                '�Ζ������}�X�^�擾���s
                ReDim w_strMsg(1)
                w_strMsg(1) = "�Ζ��������"
                Call General.paMsgDsp("NS0031", w_strMsg)
            End If

            '���
            w_Form.Dispose()

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    '2015/04/13 Bando Add Start =======================================================
    Private Function Disp_Comment(ByRef p_com As String, ByRef p_riyu As String) As Boolean

        Const W_SUBNAME As String = "NSK0000HC  Disp_Comment"

        Dim w_Form As Object

        '������
        Disp_Comment = False
        Try
            w_Form = New frmNSK0000HP
            w_Form.pRiyuKbn = p_riyu
            '�R�����g�����ɑ��݂���ꍇ�̓v���p�e�B�Ŏ󂯓n��
            If p_com <> "" Then
                w_Form.p_com = p_com
            End If

            w_Form.frmNSK0000HP_Load()

            '�\��
            w_Form.ShowDialog(Me)

            'OK�������ݎ��̂ݏ������s
            If w_Form.pOKFlg = True Then
                p_com = w_Form.pComment
                Disp_Comment = True
            End If

            '���
            w_Form = Nothing

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function
    '2015/04/13 Bando Add End   =======================================================

    Private Sub chkSet_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSet.CheckStateChanged

        Const W_SUBNAME As String = "NSK0000HC  chkSet_Click"
        Try
            If chkSet.CheckState = 1 Then
                '���ѕύX�ł͂Ȃ�
                If m_Mode <> General.G_PGMSTARTFLG_CHANGEJISSEKI Then
                    If g_LimitedFlg = False Then
                        If g_SaikeiFlg = True Then
                            _OptRiyu_0.Enabled = False
                            _OptRiyu_1.Enabled = False
                            _OptRiyu_2.Enabled = False
                            _OptRiyu_3.Enabled = True
                            _OptRiyu_3.Checked = True
                            _OptRiyu_4.Enabled = False
                            _Frame_3.Enabled = True
                            _Frame_0.Enabled = False
                            _Frame_1.Enabled = False
                            _Frame_2.Enabled = False
                        Else
                            picFrame.Enabled = False
                            _OptRiyu_0.Checked = True
                            _OptRiyu_0.Enabled = True
                            _OptRiyu_1.Enabled = False
                            '2014/05/14 Shimpo upd start P-06991-----------------------------------------------------------------------
                            '_OptRiyu_2.Enabled = False
                            If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                                _OptRiyu_2.Enabled = False
                            Else
                                _OptRiyu_2.Enabled = True
                            End If
                            '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                            _OptRiyu_4.Enabled = False
                            _Frame_3.Enabled = True
                            _Frame_0.Enabled = False
                            _Frame_1.Enabled = False
                            _Frame_2.Enabled = False
                        End If
                    Else
                        _Frame_3.Enabled = True
                        _Frame_0.Enabled = False
                        _Frame_1.Enabled = False
                        _Frame_2.Enabled = False
                    End If
                End If
            Else
                '���ѕύX�ł͂Ȃ�
                If m_Mode <> General.G_PGMSTARTFLG_CHANGEJISSEKI Then
                    If g_LimitedFlg = False Then
                        If g_SaikeiFlg = True Then
                            picFrame.Enabled = True
                            _OptRiyu_0.Enabled = False
                            _OptRiyu_1.Enabled = False
                            _OptRiyu_2.Enabled = False
                            _OptRiyu_3.Enabled = True
                            _OptRiyu_3.Checked = True
                            _OptRiyu_4.Enabled = False
                            _Frame_3.Enabled = False
                            _Frame_0.Enabled = True
                            _Frame_1.Enabled = True
                            _Frame_2.Enabled = True
                        Else
                            picFrame.Enabled = True
                            _OptRiyu_0.Enabled = True
                            _OptRiyu_1.Enabled = True
                            _OptRiyu_2.Enabled = True
                            _OptRiyu_4.Enabled = True
                            _Frame_3.Enabled = False
                            _Frame_0.Enabled = True
                            _Frame_1.Enabled = True
                            _Frame_2.Enabled = True
                        End If
                    Else
                        _Frame_3.Enabled = False
                        _Frame_0.Enabled = True
                        _Frame_1.Enabled = True
                        _Frame_2.Enabled = True
                    End If
                End If
            End If
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '�������݉���
    Private Sub cmdEnd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEnd.Click

        Const W_SUBNAME As String = "NSK0000HC  cmdEnd_Click"
        Try
            '��ʏ���
            Me.Close()

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '�����S�����݉���
    Private Sub cmdErase_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdErase.Click

        Const W_SUBNAME As String = "NSK0000HC  cmdErase_Click"

        Dim w_Lng As Integer
        Dim w_i As Integer
        Dim w_Cnt As Short
        Dim w_Row As Integer
        Dim w_YYYYMMDD As Integer
        Dim w_Var As Object
        Dim w_strMsg() As String
        Dim w_strUpdKojyoDate As String = String.Empty
        Dim w_CellRange() As Model.CellRange
        Dim w_KinmuCD As String   'KinmuCD
        Dim w_RiyuKBN As String   '���R�敪
        Dim w_Time As String  '���ԔN�x
        Dim w_Flg As String   '�m���׸�
        Dim w_KinmuPlanCD As String
        Dim w_KangoCD As String
        Dim w_STS As Integer
        Dim w_ActiveCol As Long
        Dim w_ActiveRow As Long
        Dim w_MsgFlg1 As Boolean '��]�Ζ�
        Dim w_MsgFlg2 As Boolean '�Čf�Ζ�
        Dim w_MsgFlg3 As Boolean '�ψ���Ζ�
        Dim w_MsgFlg4 As Boolean '�����Ζ�
        Dim w_MsgFlg5 As Boolean '�v���Ζ�

        Try
            With sprSheet.Sheets(0)

                w_MsgFlg1 = False
                w_MsgFlg2 = False
                w_MsgFlg3 = False
                w_MsgFlg4 = False
                w_MsgFlg5 = False

                '�����ʒu����(��)
                If .ActiveColumn.Index < m_KeikakuD_StartCol Or .ActiveColumn.Index2 > m_KeikakuD_EndCol Then
                    Exit Sub
                End If

                '�����ʒu����(�s)
                If .ActiveRow.Index < M_KinmuData_Row Or .ActiveRow.Index2 < M_KinmuData_Row Then
                    Exit Sub
                End If

                '�A����������ۯ����H
                w_CellRange = .GetSelections

                For w_i = 0 To w_CellRange.Length - 1
                    '�����ʒu����(��)
                    If w_CellRange(w_i).Column < m_KeikakuD_StartCol Or (w_CellRange(w_i).Column + w_CellRange(w_i).ColumnCount - 1) > m_KeikakuD_EndCol Then
                        Exit Sub
                    End If

                    '�����ʒu����(�s)
                    If w_CellRange(w_i).Row <> M_KinmuData_Row Or w_CellRange(w_i).RowCount <> 1 Then
                        Exit Sub
                    End If
                Next w_i

                If w_CellRange.Length = 0 Then
                    ReDim w_CellRange(0)
                    w_CellRange(0) = New Model.CellRange(.ActiveRow.Index, .ActiveColumn.Index, 1, 1)
                End If

                For w_i = 0 To w_CellRange.Length - 1
                    w_Row = w_CellRange(w_i).Row

                    For w_Lng = w_CellRange(w_i).Column To w_CellRange(w_i).Column + w_CellRange(w_i).ColumnCount - 1
                        If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_Row, w_Lng).BackColor) = m_MonthBefore_Back Then
                            '�z�����ԊO(�w�i�F���O���[)�̏ꍇ���͕s��
                            Exit Sub
                        End If

                        w_Cnt = CShort(w_Lng - M_KinmuData_Col + 1)

                        If UBound(m_DataFlg) >= w_Cnt Then
                            If g_SaikeiFlg = True Then
                                '�Čf�����̏ꍇ
                                If m_DataFlg(w_Cnt) = "1" Then
                                    '�����ް��̏ꍇ�͏����s��
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "�m��ς݋Ζ���"
                                    Call General.paMsgDsp("NS0098", w_strMsg)
                                    Exit Sub
                                End If
                            Else
                                '�Čf�����ȊO�̏ꍇ
                                If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN And m_DataFlg(w_Cnt) = "1" And m_KakuteiFlg(w_Cnt) = "0" Then
                                    '�v��ύX�ŊY�������m���ް��̏ꍇ�͏����s��
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "�m��ς݋Ζ���"
                                    Call General.paMsgDsp("NS0098", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If

                        w_Var = .GetText(w_Row, w_CellRange(w_i).Column)
                        '2015/04/10 Bando Upd Start =======================
                        'Call Get_KinmuMark(w_Var, w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        Call Get_KinmuMark(w_Var, w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, "")
                        '2015/04/10 Bando Upd End   =======================
                        '�x���\��
                        If g_SaikeiFlg = False Then
                            Select Case w_RiyuKBN
                                Case "2"
                                    If w_MsgFlg5 = False Then
                                        If frmNSK0000HA._mnuTool_5.Checked = True Then
                                            ReDim w_strMsg(2)
                                            w_strMsg(1) = "�v���Ζ�"
                                            w_strMsg(2) = "�폜"
                                            w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                            If w_STS = MsgBoxResult.No Then
                                                Exit Sub
                                            End If
                                            w_MsgFlg5 = True
                                        End If
                                    End If
                                Case "3"
                                    If w_MsgFlg1 = False Then
                                        If frmNSK0000HA._mnuTool_4.Checked = True Then
                                            ReDim w_strMsg(2)
                                            w_strMsg(1) = "��]�Ζ�"
                                            w_strMsg(2) = "�폜"
                                            w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                            If w_STS = MsgBoxResult.No Then
                                                Exit Sub
                                            End If
                                            w_MsgFlg1 = True
                                        End If
                                    End If
                                Case "4"
                                    If w_MsgFlg2 = False Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "�Čf�Ζ�"
                                        w_strMsg(2) = "�폜"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                        w_MsgFlg2 = True
                                    End If
                                Case "5"
                                    If w_MsgFlg3 = False Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "�ψ���Ζ�"
                                        w_strMsg(2) = "�폜"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                        w_MsgFlg3 = True
                                    End If
                                Case "6"
                                    If w_MsgFlg4 = False Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "�����Ζ�"
                                        w_strMsg(2) = "�폜"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                        w_MsgFlg4 = True
                                    End If
                            End Select
                        End If

                        If General.g_lngDaikyuMng = 0 Then
                            '��x����
                            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_Lng)
                            w_YYYYMMDD = w_Var
                            If Check_Daikyu(M_DELETE, w_YYYYMMDD, "") = False Then
                                Exit Sub
                            End If
                        End If

                        w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_Lng)
                        If InStr(m_strUpdKojyoDate, w_Var) = 0 Then
                            w_strUpdKojyoDate = w_strUpdKojyoDate & w_Var & ","
                        End If
                    Next w_Lng

                    sprSheet.Sheets(0).ClearRange(w_CellRange(w_i).Row, w_CellRange(w_i).Column, w_CellRange(w_i).RowCount, w_CellRange(w_i).ColumnCount, True)

                    .Cells(w_CellRange(w_i).Row, w_CellRange(w_i).Column, w_CellRange(w_i).Row + w_CellRange(w_i).RowCount - 1, w_CellRange(w_i).Column + w_CellRange(w_i).ColumnCount - 1).ForeColor = Color.Black
                    .Cells(w_CellRange(w_i).Row, w_CellRange(w_i).Column, w_CellRange(w_i).Row + w_CellRange(w_i).RowCount - 1, w_CellRange(w_i).Column + w_CellRange(w_i).ColumnCount - 1).BackColor = Color.White
                Next w_i
            End With

            m_strUpdKojyoDate = m_strUpdKojyoDate & w_strUpdKojyoDate

            '�X�V�׸ސݒ�
            m_KosinFlg = True

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    'Private Sub m_lstCmdKinmu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _cmdKinmu_0.Click, _cmdKinmu_2.Click, _cmdKinmu_4.Click, _
    '                                                                                                            _cmdKinmu_6.Click, _cmdKinmu_8.Click, _cmdKinmu_10.Click, _cmdKinmu_12.Click, _
    '                                                                                                            _cmdKinmu_14.Click, _cmdKinmu_16.Click, _cmdKinmu_18.Click, _cmdKinmu_20.Click, _
    '                                                                                                            _cmdKinmu_22.Click, _cmdKinmu_24.Click, _cmdKinmu_26.Click, _cmdKinmu_28.Click, _
    '                                                                                                            _cmdKinmu_30.Click, _cmdKinmu_34.Click, _cmdKinmu_36.Click, _cmdKinmu_38.Click
    Private Sub m_lstCmdKinmu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim Index As Short = m_lstCmdKinmu.IndexOf(eventSender)
        Const W_SUBNAME As String = "NSK0000HC  m_lstCmdKinmu_Click"

        Dim w_RegStr As String
        Dim w_Var As Object
        Dim w_ActiveCol As Integer
        Dim w_ActiveRow As Integer
        Dim w_KinmuCD As String 'KinmuCD
        Dim w_RiyuKBN As String '���R�敪
        Dim w_Time As String '���ԔN�x
        Dim w_Flg As String '�m���׸�
        Dim w_ForeColor As Integer '�����F
        Dim w_BackColor As Integer '�w�i�F
        Dim w_InputFlg As Boolean '�����׸�
        Dim w_Cnt As Short
        Dim w_KinmuPlanCD As String = String.Empty 'KinmuCD(�\���ް�)
        Dim w_STS As Short
        Dim w_KangoCD As String
        Dim w_RiyuPlanKbn As String
        Dim w_KangoPlanCD As String
        Dim w_YYYYMMDD As Integer
        Dim w_DaikyuInputFlg As Boolean
        Dim w_strMsg() As String
        Dim w_lngBackColor As Integer
        Dim w_IntCol As Short
        Dim w_KibouCnt As Short
        Dim w_KibouCol() As Integer
        Dim w_blnColChk As Boolean
        Dim w_Style As New StyleInfo
        '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
        Dim w_KibouCntDate As Integer
        '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------
        Dim w_Comment As String = String.Empty  '��]�Ζ����̃R�����g 2015/04/13 Band Add
        Try
            '̫����ړ�
            sprSheet.Focus()

            'ڼ޽�؊i�[��
            w_RegStr = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            '������
            w_KangoCD = ""

            '�P���ł����݂���΁E�E�E
            If m_KinmuCnt <> 0 Then

                w_ActiveCol = sprSheet.Sheets(0).ActiveColumn.Index
                w_ActiveRow = sprSheet.Sheets(0).ActiveRow.Index

                '���͏ꏊ����
                If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row_ChgJisseki Then
                        Exit Sub
                    End If
                Else
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row Then
                        Exit Sub
                    End If
                End If

                '�����ٓ������i�z���͈́j
                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_ActiveCol).BackColor) = m_MonthBefore_Back Then
                    '�z�����ԊO(�w�i�F���O���[)�̏ꍇ���͕s��
                    Exit Sub
                End If

                '�����׸�
                w_InputFlg = True
                w_Cnt = CShort(w_ActiveCol - M_KinmuData_Col + 1)
                If g_SaikeiFlg = True Then
                    '�Čf�����̏ꍇ
                    If m_DataFlg(w_Cnt) = "1" Then
                        '�����ް��̏ꍇ
                        ReDim w_strMsg(2)
                        w_strMsg(1) = "�m��ς݋Ζ�"
                        w_strMsg(2) = "�Čf�Ζ�"
                        Call General.paMsgDsp("NS0011", w_strMsg)

                        w_InputFlg = False
                    End If
                End If


                '�Ζ��̗j�������`�F�b�N
                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                '2013/10/02 Bando Chg Start ====================================================
                'If Check_YoubiLimit(w_YYYYMMDD, m_Kinmu(Index + HscKinmu.Value).CD) = False Then
                If Check_YoubiLimit(w_YYYYMMDD, m_Kinmu(Index + HscKinmu.Value * 2).CD) = False Then
                    '2013/10/02 Bando Chg End ====================================================
                    Exit Sub
                End If

                '2013/01/07 Ishiga add start------------------------------------------
                '���΃f�[�^�̗L���`�F�b�N
                If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If Check_OverKinmuData(w_YYYYMMDD) = False Then
                        Exit Sub
                    End If
                End If
                '2013/01/07 Ishiga add end--------------------------------------------

                '�͏o���݃`�F�b�N
                If fncChkAppliData(w_YYYYMMDD) = False Then
                    Exit Sub
                End If

                If m_PackageFLG = 1 Or m_PackageFLG = 3 Then
                    '�������̃R���|�[�l���g����̎�
                    With General.g_objGetData
                        .p�E���敪 = 0
                        .p�E���ԍ� = M_StaffID
                        .p�`�F�b�N��� = w_YYYYMMDD
                        .p�����敪 = 0
                        '2013/10/02 Bando Chg Start ====================================================
                        '.p�`�F�b�N�Ζ�CD = m_Kinmu(Index + HscKinmu.Value).CD
                        .p�`�F�b�N�Ζ�CD = m_Kinmu(Index + HscKinmu.Value * 2).CD
                        '2013/10/02 Bando Chg Start ====================================================

                        If .mChkKinmuDuty = False Then
                            '�Ζ��ύX�s��
                            '*******ү����***********************************
                            ReDim w_strMsg(1)
                            w_strMsg(1) = ""
                            Call General.paMsgDsp("NS0110", w_strMsg)
                            '************************************************
                            Exit Sub
                        End If
                    End With
                End If


                w_DaikyuInputFlg = True

                If General.g_lngDaikyuMng = 0 Then
                    '��x����
                    w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                    w_YYYYMMDD = w_Var
                    '2013/10/02 Bando Chg Start ======================================================================
                    'If Check_Daikyu(M_PASTE, w_YYYYMMDD, m_Kinmu(Index + HscKinmu.Value).CD) = False Then
                    If Check_Daikyu(M_PASTE, w_YYYYMMDD, m_Kinmu(Index + HscKinmu.Value * 2).CD) = False Then
                        '2013/10/02 Bando Chg End ======================================================================
                        w_DaikyuInputFlg = False
                        w_InputFlg = False
                    End If
                End If

                If w_InputFlg = True Then
                    '���͉\�ȏꍇ
                    With sprSheet.Sheets(0)

                        '�x���\��+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        '�ύX����Ζ��̒l����Ōv��ύX�̏ꍇ�A���̋Ζ����擾����B
                        w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                        If w_Var = "" And m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            w_Var = .GetText(w_ActiveRow - 1, w_ActiveCol)
                        End If

                        '2015/04/13 Bando Upd Start ========================
                        'Call Get_KinmuMark(w_Var, w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        Call Get_KinmuMark(w_Var, w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, "")
                        '2015/04/13 Bando Upd End   ========================

                        If g_SaikeiFlg = False Then
                            Select Case w_RiyuKBN
                                Case "2"
                                    If frmNSK0000HA._mnuTool_5.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "�v���Ζ�"
                                        w_strMsg(2) = "�ҏW"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "3"
                                    If frmNSK0000HA._mnuTool_4.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "��]�Ζ�"
                                        w_strMsg(2) = "�ҏW"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "4"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�Čf�Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "5"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�ψ���Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "6"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�����Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                            End Select
                        Else
                            If w_RiyuKBN = "2" Then
                                If frmNSK0000HA._mnuTool_5.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�v���Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If

                            If w_RiyuKBN = "3" Then
                                If frmNSK0000HA._mnuTool_4.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "��]�Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If

                        w_KinmuCD = m_Kinmu(Index + HscKinmu.Value * 2).CD
                        w_Time = ""
                        w_Flg = "0"
                        w_KangoCD = ""

                        '��]�񐔏W�v�`�F�b�N
                        w_lngBackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                        If CDbl(g_HopeNumFlg) = 1 And _OptRiyu_2.Checked = True Then
                            w_KibouCnt = 0
                            ReDim w_KibouCol(0)
                            For w_IntCol = m_KeikakuD_StartCol To m_KeikakuD_EndCol

                                If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_IntCol).BackColor) = w_lngBackColor Then
                                    w_KibouCnt = w_KibouCnt + 1
                                    ReDim Preserve w_KibouCol(w_KibouCnt)
                                    w_KibouCol(w_KibouCnt) = w_IntCol
                                End If
                            Next w_IntCol

                            '�����ꏊ�ɓ����Ζ���\��t�����ꍇ�A�X���[
                            w_blnColChk = False
                            For w_IntCol = 1 To UBound(w_KibouCol)
                                If w_ActiveCol = w_KibouCol(w_IntCol) And w_KinmuCD = w_KinmuPlanCD Then
                                    w_blnColChk = True
                                End If
                            Next w_IntCol

                            If g_HopeNum <= w_KibouCnt And w_blnColChk = False Then
                                '��]�񐔐����I�[�o�[�_�C�A���O�\��
                                If g_KibouNumDiaLogFlg = 1 Then
                                    '���[�j���O
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "��]�Ζ�"
                                    w_strMsg(2) = "�ݒ肳�ꂽ��]�Ζ���"
                                    '�u&1��&2�𒴂��Ă��܂��B~n���̂܂ܓo�^���Ă���낵���ł����B�v
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    '�G���[
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "�ݒ肳�ꂽ��]�Ζ��񐔂𒴂��Ă��邽��"
                                    '�u&1���͂ł��܂���B�v
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If

                        '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
                        If CDbl(g_HopeNumDateFlg) = 1 And _OptRiyu_2.Checked = True Then
                            '���t�ʂ̊�]�Ζ����`�F�b�N
                            w_KibouCntDate = frmNSK0000HA.Get_HopeNum_Of_Date(w_YYYYMMDD)
                            If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor) = w_lngBackColor Then
                                w_KibouCntDate = w_KibouCntDate - 1
                            End If

                            If g_HopeNumDate <= w_KibouCntDate Then
                                '��]�񐔐����I�[�o�[�_�C�A���O�\��
                                If g_KibouNumDiaLogFlg = 1 Then
                                    '���[�j���O
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "��]�Ζ�"
                                    w_strMsg(2) = "�ݒ肳�ꂽ��]�Ζ�(���t��)��"
                                    '�u&1��&2�𒴂��Ă��܂��B~n���̂܂ܓo�^���Ă���낵���ł����B�v
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    '�G���[
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "�ݒ肳�ꂽ��]�Ζ�(���t��)�񐔂𒴂��Ă��邽��"
                                    '�u&1���͂ł��܂���B�v
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If
                        '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------

                        Select Case True
                            Case _OptRiyu_0.Checked '�ʏ�
                                '���R�敪 �ʏ�
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                            Case _OptRiyu_1.Checked '�v��
                                '���R�敪 �v��
                                w_RiyuKBN = "2"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", General.G_PALEGREEN))
                            Case _OptRiyu_2.Checked '��]
                                '���R�敪 ��]
                                w_RiyuKBN = "3"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                            Case _OptRiyu_3.Checked
                                '���R�敪 �Čf
                                w_RiyuKBN = "4"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", General.G_LIGHTCYAN))
                            Case _OptRiyu_4.Checked
                                w_RiyuKBN = "6"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", General.G_ORANGE))
                            Case Else
                                '���R�敪 �ʏ�
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                        End Select

                        '�敪�������̏ꍇ�̂݁A������Ζ��n�I����ʂ�\��
                        If w_RiyuKBN = "6" Then
                            If Disp_Ouen(w_KangoCD) = False Then
                                Exit Sub
                            End If
                        End If

                        '2015/04/13 Bando Add Start ========================
                        '��]�̏ꍇ�R�����g���͉�ʕ\��
                        If w_RiyuKBN = "3" And g_InputHopeCommentFlg = "1" Then
                            If Disp_Comment(w_Comment, w_RiyuKBN) = False Then
                                Exit Sub
                            End If
                        End If
                        '2015/04/13 Bando Add End   ========================

                        '2015/07/22 Bando Add Start ========================
                        If w_Comment <> "" Then
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Fore", Convert.ToString(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Back", ColorTranslator.ToOle(Color.IndianRed)))
                        End If
                        '2015/07/22 Bando Add End   ========================

                        '��x�����Ζ��ޯ��װ
                        If General.g_lngDaikyuMng = 0 Then
                            If m_DaikyuBackColorFlg = True Then
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Back", General.G_LIGHTYELLOW))
                            End If
                        End If

                        '�قɋL����ݒ�
                        '2015/04/13 Bando Upd Start =======================
                        'w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, w_Comment)
                        '2015/04/13 Bando Upd End   =======================
                        .SetText(w_ActiveRow, w_ActiveCol, w_Var)
                        '���R�ʐF�ݒ�
                        .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                        .Cells(w_ActiveRow, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)

                        '�Ζ��ύX�̎�
                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then

                            '2015/04/13 Bando Upd Start ====================
                            'Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time)
                            Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time, "")
                            '2015/04/13 Bando Upd End   ====================

                            '�\��ƈقȂ�ꍇ�F�ύX
                            If Trim(w_KinmuCD) <> Trim(w_KinmuPlanCD) Or w_RiyuKBN <> w_RiyuPlanKbn Or w_KangoCD <> w_KangoPlanCD Then
                                .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(m_Comp_Fore)
                            End If

                            '�ύX�\�s�̓��e���i�[
                            w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                            w_ForeColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).ForeColor)
                            w_BackColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor)

                            '�ύX�\�s�̓��e���Ζ��ύX��ʂɓ\��t���s�փR�s�[
                            .SetText(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol, w_Var)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        End If
                    End With
                End If
            End If

            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
            If InStr(m_strUpdKojyoDate, w_Var) = 0 Then
                m_strUpdKojyoDate = m_strUpdKojyoDate & w_Var & ","
            End If

            '���پوړ�
            If w_DaikyuInputFlg = True Then
                Call SetCursol()
            End If

            '�X�V�׸޾��
            If w_InputFlg = True Then
                m_KosinFlg = True
            End If
            '�v�C�� ed ********************

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '���΃f�[�^�̗L���`�F�b�N
    Private Function Check_OverKinmuData(ByVal p_Date As Object) As Boolean
        Const W_SUBNAME As String = "NSK0000HA Check_OverKinmuData"


        Dim w_strMsg() As String

        Check_OverKinmuData = False
        Try
            '���΃f�[�^�擾
            With General.g_objGetData
                .p�a�@CD = General.g_strHospitalCD
                .p���F�敪 = 2 '0:����F�A1:����F�A2:����
                .p�E���敪 = 0 '0:�E���Ǘ��ԍ�
                .p�E���ԍ� = M_StaffID '�I��E���Ǘ��ԍ�
                .p���t�敪 = 0 '0:�P���
                .p�J�n�N���� = p_Date '�J�n�N����
                .p�I���N���� = 0 '�I���N����

                If .mGetOverKinmu = True Then
                    ReDim w_strMsg(1)
                    w_strMsg(1) = "���ԊO�����ɓo�^����Ă��邽��~n"
                    Call General.paMsgDsp("NS0110", w_strMsg)
                    Exit Function
                End If
            End With

            Check_OverKinmuData = True

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    '��x�f�[�^�`�F�b�N�i�w�i�F�j
    Private Function Check_DaikyuBackColor(ByVal p_Date As Integer) As Boolean

        Const W_SUBNAME As String = "NSK0000HC Check_DaikyuBackColor"

        Dim w_Int As Integer

        '�����l
        Check_DaikyuBackColor = False
        Try
            '��x�f�[�^���[�v
            For w_Int = 1 To UBound(m_DaikyuData)
                If m_DaikyuData(w_Int).HolDate = p_Date Or (m_SundayDaikyuFlg = 1 And Weekday(CDate(Format(p_Date, "0000/00/00"))) = 1) Or (m_SaturdayDaikyuFlg = 1 And Weekday(CDate(Format(p_Date, "0000/00/00"))) = 7) Then
                    Check_DaikyuBackColor = True
                    Exit Function
                End If
            Next w_Int

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    '�Ζ��\�t��
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click

        Const W_SUBNAME As String = "NSK0000HC  cmdOK_Click"

        Dim w_Cnt As Short
        Dim w_Kinmu As Object
        Dim w_Color As Integer
        Dim w_Row As Integer '�Ζ��L���ް� �J�n�� �ʒu
        Dim w_Kinmu_Param As Object
        Dim w_Row_Param As Integer
        Dim w_blnUpdFLG As Boolean
        Dim w_Style As New StyleInfo()
        Try
            If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                w_Row = M_KinmuData_Row_ChgJisseki + 1
            Else
                w_Row = M_KinmuData_Row
            End If

            '�I���s���Ζ��\��̍s�ɓ��ꂷ��
            m_CUR_ROW_Param = m_CUR_ROW_Param - ((m_CUR_ROW_Param - m_StaffStartRow) Mod m_MaxShowLine) + m_KinmuPlan

            For w_Cnt = 1 To (m_KeikakuD_EndCol - m_KeikakuD_StartCol + 1)

                '�Ζ��擾
                w_Kinmu = sprSheet.Sheets(0).GetText(w_Row, m_KeikakuD_StartCol + w_Cnt - 1) '�Ζ��ύX�̎�

                '�X�V����H
                w_blnUpdFLG = False
                If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If m_blnOneTwo Then
                        '��i�\���̎��@���i���X�V
                        w_Row_Param = m_CUR_ROW_Param + 1
                        '��i�̔w�i�F���m�F
                        w_Style.Reset()
                        w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(m_CUR_ROW_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                        If w_Style.BackColor.A = 0 Then
                            w_Color = ColorTranslator.ToOle(Color.White)
                        Else
                            w_Color = ColorTranslator.ToOle(w_Style.BackColor)
                        End If
                    Else
                        '��i�\���̎��@��i���X�V
                        w_Row_Param = m_CUR_ROW_Param
                        '���i�̔w�i�F���m�F
                        w_Style.Reset()
                        w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                        If w_Style.BackColor.A = 0 Then
                            w_Color = ColorTranslator.ToOle(Color.White)
                        Else
                            w_Color = ColorTranslator.ToOle(w_Style.BackColor)
                        End If
                    End If

                    w_Kinmu_Param = m_Control_Param.Sheets(0).GetText(m_CUR_ROW_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1)
                    If m_Control_Param.Sheets(0).GetText(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1) <> "" Or w_Color = ColorTranslator.ToOle(Color.Black) Then
                        '���i�Ƀf�[�^�����݂���Ƃ� �܂��� �w�i�F�����̂Ƃ��@�X�V
                        w_blnUpdFLG = True
                    ElseIf w_Kinmu_Param <> "" Or (w_Kinmu_Param = "" And m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN) Then
                        '���i�Ƀf�[�^�����݂��Ȃ��Ƃ�
                        If w_Kinmu_Param <> w_Kinmu Then
                            '��i�̃f�[�^�ƈقȂ�Ζ��̂Ƃ��@�X�V
                            w_blnUpdFLG = True
                            If m_blnOneTwo = False Then
                                '��i�\���̎��@��i�̃f�[�^�����i�ɃR�s�[

                                '�e��ʋΖ��\�t��
                                If m_Control_Param.Sheets(0).GetText(m_CUR_ROW_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1) = "" And m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                                    m_Control_Param.Sheets(0).SetText(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Kinmu)
                                Else
                                    m_Control_Param.Sheets(0).SetText(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Kinmu_Param)
                                End If

                                '�F�擾
                                w_Style.Reset()
                                w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(m_CUR_ROW_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                                w_Color = ColorTranslator.ToOle(w_Style.ForeColor)

                                '�F�\�t��
                                w_Style.Reset()
                                w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                                w_Style.ForeColor = ColorTranslator.FromOle(w_Color)
                                m_Control_Param.Sheets(0).Models.Style.SetDirectInfo(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)

                                '�F�擾
                                w_Style.Reset()
                                w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(m_CUR_ROW_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                                If w_Style.BackColor.A = 0 Then
                                    w_Color = ColorTranslator.ToOle(Color.White)
                                Else
                                    w_Color = ColorTranslator.ToOle(w_Style.BackColor)
                                End If

                                '�F�\�t��
                                w_Style.Reset()
                                w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                                w_Style.BackColor = ColorTranslator.FromOle(w_Color)
                                m_Control_Param.Sheets(0).Models.Style.SetDirectInfo(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                            End If
                        End If
                    Else
                        '��i�Ƀf�[�^�����݂��Ȃ��Ƃ��@��i���X�V
                        If m_Mode <> General.G_PGMSTARTFLG_CHANGEPLAN Then
                            w_Row_Param = m_CUR_ROW_Param
                        End If

                        w_blnUpdFLG = True
                    End If
                Else
                    w_Row_Param = m_CUR_ROW_Param
                    w_blnUpdFLG = True
                End If

                If w_blnUpdFLG Then
                    '�e��ʋΖ��\�t��
                    m_Control_Param.Sheets(0).SetText(w_Row_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Kinmu)

                    '�F�擾
                    w_Color = ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_Row, m_KeikakuD_StartCol + w_Cnt - 1).ForeColor)

                    '�F�\�t��
                    w_Style.Reset()
                    w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(w_Row_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                    w_Style.ForeColor = ColorTranslator.FromOle(w_Color)
                    m_Control_Param.Sheets(0).Models.Style.SetDirectInfo(w_Row_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)

                    '�F�擾
                    If sprSheet.Sheets(0).Cells(w_Row, m_KeikakuD_StartCol + w_Cnt - 1).BackColor.A = 0 Then
                        w_Color = ColorTranslator.ToOle(Color.White)
                    Else
                        w_Color = ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_Row, m_KeikakuD_StartCol + w_Cnt - 1).BackColor)
                    End If

                    '�F�\�t��
                    w_Style.Reset()
                    w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(w_Row_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                    w_Style.BackColor = ColorTranslator.FromOle(w_Color)
                    m_Control_Param.Sheets(0).Models.Style.SetDirectInfo(w_Row_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                End If
            Next w_Cnt

            m_OKFlg = True

            '��ʏ���
            Me.Close()
        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    'Private Sub m_lstCmdSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _CmdSet_0.Click, _CmdSet_1.Click, _CmdSet_2.Click, _CmdSet_3.Click, _
    '                                                                                                                _CmdSet_4.Click, _CmdSet_5.Click, _CmdSet_6.Click, _CmdSet_7.Click, _
    '                                                                                                                _CmdSet_8.Click, _CmdSet_9.Click, _CmdSet_10.Click, _CmdSet_11.Click, _
    '                                                                                                                _CmdSet_12.Click, _CmdSet_13.Click, _CmdSet_14.Click, _CmdSet_15.Click, _
    '                                                                                                                _CmdSet_16.Click, _CmdSet_17.Click, _CmdSet_18.Click, _CmdSet_19.Click
    Private Sub m_lstCmdSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim Index As Short = m_lstCmdSet.IndexOf(eventSender)
        Const W_SUBNAME As String = "NSK0000HC  m_lstCmdSet_Click"

        Dim w_RegStr As String
        Dim w_Var As Object
        Dim w_ActiveCol As Integer
        Dim w_ActiveRow As Integer
        Dim w_KinmuCD As String 'KinmuCD
        Dim w_RiyuKBN As String '���R�敪
        Dim w_Time As String '���ԔN�x
        Dim w_Flg As String '�m���׸�
        Dim w_ForeColor As Integer '�����F
        Dim w_BackColor As Integer '�w�i�F
        Dim w_KinmuPlanCD As String 'KinmuCD(�\���ް�)
        Dim w_STS As Short
        Dim w_KangoCD As String
        Dim w_KangoPlanCD As String
        Dim w_RiyuPlanKbn As String
        Dim w_YYYYMMDD As Integer
        Dim w_Col As Integer
        Dim w_StopFlg As Boolean
        Dim w_DaikyuCnt As Integer
        Dim w_MsgFlg As Boolean
        Dim w_strMsg() As String
        Dim w_lngBackColor As Integer
        Dim w_IntCol As Short
        Dim w_KibouCnt As Short
        Dim w_objSyouninData As Object
        Dim w_ColCnt As Short
        Dim w_lngLoop As Integer
        Try
            '̫����ړ�
            sprSheet.Focus()

            ''ڼ޽�؊i�[��
            w_RegStr = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            '�P���ł����݂���΁E�E�E
            If m_SetCnt <> 0 Then

                w_ActiveCol = sprSheet.Sheets(0).ActiveColumn.Index
                w_ActiveRow = sprSheet.Sheets(0).ActiveRow.Index

                w_ColCnt = w_ActiveCol

                With sprSheet.Sheets(0)
                    '�x���\��+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    If g_SaikeiFlg = False Then

                        w_MsgFlg = False
                        '2013/10/02 Bando Chg Start =========================================================
                        'For w_Col = 1 To m_SetKinmu(Index + HscSet.Value).KinmuCnt
                        For w_Col = 1 To m_SetKinmu(Index + HscSet.Value * 2).KinmuCnt
                            '2013/10/02 Bando Chg End =========================================================

                            '�v���ް��� �͈͓��� ?
                            If Not ((m_KeikakuD_StartCol <= w_ColCnt) And (w_ColCnt <= m_KeikakuD_EndCol)) Then
                                'ERROR
                                Exit For
                            End If

                            '2015/04/13 Bando Upd Start ==================
                            'Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                            Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, "")
                            '2015/04/13 Bando Upd End   ==================
                            Select Case w_RiyuKBN
                                Case "3"
                                    If frmNSK0000HA._mnuTool_4.Checked = True Then
                                        w_MsgFlg = True
                                        Exit For
                                    End If
                                Case "4", "5", "6"
                                    w_MsgFlg = True
                                    Exit For
                            End Select

                            '�Ζ��̗j�������`�F�b�N
                            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                            w_YYYYMMDD = w_Var
                            '2013/10/02 Bando Chg Start =======================================================
                            'If Check_YoubiLimit(w_YYYYMMDD, m_SetKinmu(Index + HscSet.Value).CD(w_Col)) = False Then
                            If Check_YoubiLimit(w_YYYYMMDD, m_SetKinmu(Index + HscSet.Value * 2).CD(w_Col)) = False Then
                                '2013/10/02 Bando Chg End =======================================================
                                Exit Sub
                            End If

                            '2013/01/07 Ishiga add start------------------------------------------
                            '���΃f�[�^�̗L���`�F�b�N
                            If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                                If Check_OverKinmuData(w_YYYYMMDD) = False Then
                                    Exit Sub
                                End If
                            End If
                            '2013/01/07 Ishiga add end--------------------------------------------

                            '�͏o���݃`�F�b�N
                            If fncChkAppliData(w_YYYYMMDD) = False Then
                                Exit Sub
                            End If


                            If m_PackageFLG = 1 Or m_PackageFLG = 3 Then
                                '�������̃R���|�[�l���g����̎�
                                With General.g_objGetData
                                    .p�E���敪 = 0
                                    .p�E���ԍ� = M_StaffID
                                    .p�`�F�b�N��� = w_YYYYMMDD
                                    .p�����敪 = 0
                                    '2013/10/02 Bando Chg Start =======================================================
                                    '.p�`�F�b�N�Ζ�CD = m_SetKinmu(Index + HscSet.Value).CD(w_Col)
                                    .p�`�F�b�N�Ζ�CD = m_SetKinmu(Index + HscSet.Value * 2).CD(w_Col)
                                    '2013/10/02 Bando Chg End   =======================================================

                                    If .mChkKinmuDuty = False Then
                                        '�Ζ��ύX�s��
                                        '*******ү����***********************************
                                        ReDim w_strMsg(1)
                                        w_strMsg(1) = ""
                                        Call General.paMsgDsp("NS0110", w_strMsg)
                                        '************************************************
                                        Exit Sub
                                    End If
                                End With
                            End If

                            w_ActiveCol = w_ActiveCol + 1

                            w_ColCnt = w_ColCnt + 1
                        Next w_Col

                        If w_MsgFlg = True Then
                            ReDim w_strMsg(0)
                            w_STS = General.paMsgDsp("NS0120", w_strMsg)
                            If w_STS = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                    '��]�񐔏W�v�`�F�b�N
                    w_lngBackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                    If CDbl(g_HopeNumFlg) = 1 And _OptRiyu_2.Checked = True Then
                        w_KibouCnt = 0
                        For w_IntCol = m_KeikakuD_StartCol To m_KeikakuD_EndCol
                            If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_IntCol).BackColor) = w_lngBackColor Then
                                w_KibouCnt = w_KibouCnt + 1
                            End If
                        Next w_IntCol
                        '2013/10/02 Bando Chg Start ==============================================================
                        'If g_HopeNum < w_KibouCnt + m_SetKinmu(Index + HscSet.Value).KinmuCnt Then
                        If g_HopeNum < w_KibouCnt + m_SetKinmu(Index + HscSet.Value * 2).KinmuCnt Then
                            '2013/10/02 Bando Chg End ==============================================================
                            '��]�񐔐����I�[�o�[�_�C�A���O�\��
                            If g_KibouNumDiaLogFlg = 1 Then
                                '���[�j���O
                                ReDim w_strMsg(2)
                                w_strMsg(1) = "��]�Ζ�"
                                w_strMsg(2) = "�ݒ肳�ꂽ��]�Ζ���"
                                '�u&1��&2�𒴂��Ă��܂��B~n���̂܂ܓo�^���Ă���낵���ł����B�v
                                w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                If w_STS = MsgBoxResult.No Then
                                    Exit Sub
                                End If
                            Else
                                '�G���[
                                ReDim w_strMsg(1)
                                w_strMsg(1) = "�ݒ肳�ꂽ��]�Ζ��񐔂𒴂��Ă��邽��"
                                '�u&1���͂ł��܂���B�v
                                w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                Exit Sub
                            End If
                        End If
                    End If

                    w_ActiveCol = sprSheet.Sheets(0).ActiveColumn.Index
                    w_ActiveRow = sprSheet.Sheets(0).ActiveRow.Index

                    '2013/10/02 Bando Chg Start ==========================================
                    'For w_Col = 1 To m_SetKinmu(Index + HscSet.Value).KinmuCnt
                    For w_Col = 1 To m_SetKinmu(Index + HscSet.Value * 2).KinmuCnt
                        '2013/10/02 Bando Chg End ==========================================

                        w_StopFlg = False

                        '�\��t�����߰����Ȃ��ꍇ�͏����I��
                        If m_KeikakuD_EndCol < w_ActiveCol Then
                            w_ActiveCol = w_ActiveCol - 1
                            Exit For
                        End If

                        '���͏ꏊ����
                        If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row Then
                            w_ActiveCol = w_ActiveCol - 1
                            Exit For
                        End If

                        '�����ٓ������i�z���͈́j
                        If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_ActiveCol).BackColor) = m_MonthBefore_Back Then
                            '�z�����ԊO(�w�i�F���O���[)�̏ꍇ���͕s��
                            Exit Sub
                        End If

                        If General.g_lngDaikyuMng = 0 Then
                            '��x����
                            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                            w_YYYYMMDD = w_Var
                            '���łɑ�x�擾�ς݂̃f�[�^�����邩
                            For w_DaikyuCnt = 1 To UBound(m_DaikyuData)
                                If m_DaikyuData(w_DaikyuCnt).HolDate = w_YYYYMMDD Then
                                    For w_lngLoop = 1 To UBound(m_DaikyuData(w_DaikyuCnt).DaikyuDetail)
                                        If m_DaikyuData(w_DaikyuCnt).DaikyuDetail(w_lngLoop).DaikyuDate <> 0 Then
                                            w_StopFlg = True
                                            Exit For
                                        End If
                                    Next w_lngLoop

                                    If w_StopFlg = True Then
                                        Exit For
                                    End If
                                End If
                            Next w_DaikyuCnt
                        End If

                        '������x�擾�ς݂̃f�[�^������Ώ����I��
                        If General.g_lngDaikyuMng = 0 Then
                            If w_StopFlg = True Then
                                w_ActiveCol = w_ActiveCol - 1
                                Exit For
                            Else
                                '��x�擾�ς݂̃f�[�^���Ȃ��ꍇ�́A��x�z��X�V
                                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                                w_YYYYMMDD = w_Var
                                '�Z�b�g�Ζ��ł���x��\��
                                '2013/10/02 Bando Chg Start ==========================================================
                                'Call Check_Daikyu(M_PASTE, w_YYYYMMDD, m_SetKinmu(Index + HscSet.Value).CD(w_Col))
                                Call Check_Daikyu(M_PASTE, w_YYYYMMDD, m_SetKinmu(Index + HscSet.Value * 2).CD(w_Col))
                                '2013/10/02 Bando Chg End   ==========================================================
                            End If
                        End If

                        '2013/10/02 Bando Chg Start ==========================================================
                        'w_KinmuCD = m_SetKinmu(Index + HscSet.Value).CD(w_Col)
                        w_KinmuCD = m_SetKinmu(Index + HscSet.Value * 2).CD(w_Col)
                        '2013/10/02 Bando Chg End   ==========================================================
                        w_Time = ""
                        w_Flg = "0"
                        w_KangoCD = ""

                        Select Case True
                            Case _OptRiyu_0.Checked '�ʏ�
                                '���R�敪 �ʏ�
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                            Case _OptRiyu_1.Checked '�v��
                                '���R�敪 �v��
                                w_RiyuKBN = "2"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", General.G_PALEGREEN))
                            Case _OptRiyu_2.Checked '��]
                                '���R�敪 ��]
                                w_RiyuKBN = "3"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                            Case _OptRiyu_3.Checked
                                '���R�敪 �Čf
                                w_RiyuKBN = "4"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", General.G_LIGHTCYAN))
                            Case Else
                                '���R�敪 �ʏ�
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                        End Select

                        '�Z�b�g�Ζ��ł���x��\��
                        If General.g_lngDaikyuMng = 0 Then
                            If m_DaikyuBackColorFlg = True Then
                                '��x�����Ζ��@����/�w�i�F
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Back", General.G_LIGHTYELLOW))
                            End If
                        End If

                        '�قɋL����ݒ肷��
                        '2015/04/13 Bando Upd Start =====================
                        'w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, "")
                        '2015/04/13 Bando Upd End   =====================
                        .SetText(w_ActiveRow, w_ActiveCol, w_Var)
                        '���R�ʂ̐F�ݒ�
                        .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                        .Cells(w_ActiveRow, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)

                        '�Ζ��ύX�̎�
                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            '2015/04/13 Bando Upd Start =====================
                            'Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time)
                            Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time, "")
                            '2015/04/13 Bando Upd End   =====================
                            '�\��ƈقȂ�ꍇ�F�ύX
                            If Trim(w_KinmuCD) <> Trim(w_KinmuPlanCD) Or w_RiyuKBN <> w_RiyuPlanKbn Or w_KangoCD <> w_KangoPlanCD Then
                                .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(m_Comp_Fore)
                            End If

                            '�ύX�\�s�̓��e���i�[
                            w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                            w_ForeColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).ForeColor)
                            w_BackColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor)
                            '�ύX�\�s�̓��e���Ζ��ύX��ʂɓ\��t���s�փR�s�[
                            .SetText(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol, w_Var)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        End If

                        '�X�V�׸޾��
                        m_KosinFlg = True

                        w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                        If InStr(m_strUpdKojyoDate, w_Var) = 0 Then
                            m_strUpdKojyoDate = m_strUpdKojyoDate & w_Var & ","
                        End If

                        '2013/10/02 Bando Chg Start ===================================
                        'If w_Col <> m_SetKinmu(Index + HscSet.Value).KinmuCnt Then
                        If w_Col <> m_SetKinmu(Index + HscSet.Value * 2).KinmuCnt Then
                            '2013/10/02 Bando Chg End ===================================
                            w_ActiveCol = w_ActiveCol + 1
                        End If
                    Next w_Col
                End With
            End If

            '���پوړ�
            '�وʒu�ݒ�(�Z�b�g�Ζ��̍Ō�̃Z���ɂ��킷)
            sprSheet.Sheets(0).SetActiveCell(w_ActiveRow, w_ActiveCol)

            Call SetCursol()

            If w_StopFlg = True Then
                ReDim w_strMsg(1)
                w_strMsg(1) = "��x�擾�ς݋Ζ������݂��邽��"
                Call General.paMsgDsp("NS0107", w_strMsg)
            End If
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    'Private Sub m_lstCmdTokusyu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _CmdTokusyu_0.Click, _CmdTokusyu_2.Click, _CmdTokusyu_4.Click, _CmdTokusyu_6.Click, _
    '                                                                                                                    _CmdTokusyu_8.Click, _CmdTokusyu_10.Click, _CmdTokusyu_12.Click, _CmdTokusyu_14.Click, _
    '                                                                                                                    _CmdTokusyu_16.Click, _CmdTokusyu_18.Click, _CmdTokusyu_20.Click, _CmdTokusyu_22.Click, _
    '                                                                                                                    _CmdTokusyu_24.Click, _CmdTokusyu_26.Click, _CmdTokusyu_28.Click, _CmdTokusyu_30.Click, _
    '                                                                                                                    _CmdTokusyu_32.Click, _CmdTokusyu_34.Click, _CmdTokusyu_36.Click, _CmdTokusyu_38.Click
    Private Sub m_lstCmdTokusyu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim Index As Short = m_lstCmdTokusyu.IndexOf(eventSender)
        Const W_SUBNAME As String = "NSK0000HC  m_lstCmdTokusyu_Click"

        Dim w_RegStr As String
        Dim w_Var As Object
        Dim w_ActiveCol As Integer
        Dim w_ActiveRow As Integer
        Dim w_KinmuCD As String 'KinmuCD
        Dim w_RiyuKBN As String '���R�敪
        Dim w_Time As String '���ԔN�x
        Dim w_Flg As String '�m���׸�
        Dim w_ForeColor As Integer '�����F
        Dim w_BackColor As Integer '�w�i�F
        Dim w_InputFlg As Boolean '�����׸�
        Dim w_Cnt As Short
        Dim w_KinmuPlanCD As String 'KinmuCD(�\���ް�)
        Dim w_STS As Short
        Dim w_KangoCD As String
        Dim w_KangoPlanCD As String
        Dim w_RiyuPlanKbn As String
        Dim w_YYYYMMDD As Integer
        Dim w_DaikyuInputFlg As Boolean
        Dim w_strMsg() As String
        Dim w_lngBackColor As Integer
        Dim w_IntCol As Short
        Dim w_KibouCnt As Short
        Dim w_KibouCol() As Integer
        Dim w_blnColChk As Boolean
        '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
        Dim w_KibouCntDate As Integer
        '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------
        Dim w_Comment As String = String.Empty  '��]�Ζ����̃R�����g 2015/04/13 Bando Add

        Try
            '̫����ړ�
            sprSheet.Focus()

            'ڼ޽�؊i�[��
            w_RegStr = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            '�P���ł����݂���΁E�E�E
            If m_TokusyuCnt <> 0 Then

                w_ActiveCol = sprSheet.Sheets(0).ActiveColumn.Index
                w_ActiveRow = sprSheet.Sheets(0).ActiveRow.Index

                '���͏ꏊ����
                If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row_ChgJisseki Then
                        Exit Sub
                    End If
                Else
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row Then
                        Exit Sub
                    End If
                End If

                '�����ٓ������i�z���͈́j
                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_ActiveCol).BackColor) = m_MonthBefore_Back Then
                    '�z�����ԊO(�w�i�F���O���[)�̏ꍇ���͕s��
                    Exit Sub
                End If

                '�����׸�
                w_InputFlg = True
                w_Cnt = CShort(w_ActiveCol - M_KinmuData_Col + 1)
                If g_SaikeiFlg = True Then
                    '�Čf�����̏ꍇ
                    If m_DataFlg(w_Cnt) = "1" Then
                        '�����ް��̏ꍇ
                        ReDim w_strMsg(2)
                        w_strMsg(1) = "�m��ς݋Ζ�"
                        w_strMsg(2) = "�Čf�Ζ�"
                        Call General.paMsgDsp("NS0011", w_strMsg)
                        w_InputFlg = False
                    End If
                End If


                '�Ζ��̗j�������`�F�b�N
                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                '2013/10/02 Bando Chg Start ===========================================================
                'If Check_YoubiLimit(w_YYYYMMDD, m_Tokusyu(Index + HscTokusyu.Value).CD) = False Then
                If Check_YoubiLimit(w_YYYYMMDD, m_Tokusyu(Index + HscTokusyu.Value * 2).CD) = False Then
                    '2013/10/02 Bando Chg End   ===========================================================
                    Exit Sub
                End If

                '2013/01/07 Ishiga add start------------------------------------------
                '���΃f�[�^�̗L���`�F�b�N
                If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If Check_OverKinmuData(w_YYYYMMDD) = False Then
                        Exit Sub
                    End If
                End If
                '2013/01/07 Ishiga add end--------------------------------------------

                '�͏o���݃`�F�b�N
                If fncChkAppliData(w_YYYYMMDD) = False Then
                    Exit Sub
                End If

                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                If m_PackageFLG = 1 Or m_PackageFLG = 3 Then
                    '�������̃R���|�[�l���g����̎�
                    With General.g_objGetData
                        .p�E���敪 = 0
                        .p�E���ԍ� = M_StaffID
                        .p�`�F�b�N��� = w_YYYYMMDD
                        .p�����敪 = 0
                        '2013/10/02 Bando Chg Start ======================================
                        '.p�`�F�b�N�Ζ�CD = m_Tokusyu(Index + HscTokusyu.Value).CD
                        .p�`�F�b�N�Ζ�CD = m_Tokusyu(Index + HscTokusyu.Value * 2).CD
                        '2013/10/02 Bando Chg End   ======================================

                        If .mChkKinmuDuty = False Then
                            '�Ζ��ύX�s��
                            '*******ү����***********************************
                            ReDim w_strMsg(1)
                            w_strMsg(1) = ""
                            Call General.paMsgDsp("NS0110", w_strMsg)
                            '************************************************
                            Exit Sub
                        End If
                    End With
                End If

                w_DaikyuInputFlg = True

                If General.g_lngDaikyuMng = 0 Then
                    '��x����
                    w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                    w_YYYYMMDD = w_Var
                    '2013/10/02 Bando Chg Start ======================================================
                    'If Check_Daikyu(M_PASTE, w_YYYYMMDD, m_Tokusyu(Index + HscTokusyu.Value).CD) = False Then
                    If Check_Daikyu(M_PASTE, w_YYYYMMDD, m_Tokusyu(Index + HscTokusyu.Value * 2).CD) = False Then
                        '2013/10/02 Bando Chg End   ======================================================
                        w_DaikyuInputFlg = False
                        w_InputFlg = False
                    End If
                End If

                If w_InputFlg = True Then
                    '���͉\�ȏꍇ

                    With sprSheet.Sheets(0)

                        '�x���\��+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        '�ύX����Ζ��̒l����Ōv��ύX�̏ꍇ�A���̋Ζ����擾����B
                        w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                        If w_Var = "" And m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            w_Var = .GetText(w_ActiveRow - 1, w_ActiveCol)
                        End If

                        '2015/04/13 Bando Upd Start =======================
                        'Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, w_Comment)
                        '2015/04/13 Bando Upd End   =======================
                        If g_SaikeiFlg = False Then
                            Select Case w_RiyuKBN
                                Case "2"
                                    If frmNSK0000HA._mnuTool_5.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "�v���Ζ�"
                                        w_strMsg(2) = "�ҏW"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "3"
                                    If frmNSK0000HA._mnuTool_4.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "��]�Ζ�"
                                        w_strMsg(2) = "�ҏW"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "4"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�Čf�Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "5"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�ψ���Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "6"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�����Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                            End Select
                        Else
                            If w_RiyuKBN = "2" Then
                                If frmNSK0000HA._mnuTool_5.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�v���Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If

                            If w_RiyuKBN = "3" Then
                                If frmNSK0000HA._mnuTool_4.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "��]�Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                        w_KinmuCD = m_Tokusyu(Index + HscTokusyu.Value * 2).CD
                        w_Time = ""
                        w_Flg = "0"
                        w_KangoCD = ""

                        '��]�񐔏W�v�`�F�b�N
                        w_lngBackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                        If CDbl(g_HopeNumFlg) = 1 And _OptRiyu_2.Checked = True Then
                            w_KibouCnt = 0
                            ReDim w_KibouCol(0)
                            For w_IntCol = m_KeikakuD_StartCol To m_KeikakuD_EndCol
                                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_IntCol).BackColor) = w_lngBackColor Then
                                    w_KibouCnt = w_KibouCnt + 1
                                    ReDim Preserve w_KibouCol(w_KibouCnt)
                                    w_KibouCol(w_KibouCnt) = w_IntCol
                                End If
                            Next w_IntCol

                            '�����ꏊ�ɓ����Ζ���\��t�����ꍇ�A�X���[
                            w_blnColChk = False
                            For w_IntCol = 1 To UBound(w_KibouCol)
                                If w_ActiveCol = w_KibouCol(w_IntCol) And w_KinmuCD = w_KinmuPlanCD Then
                                    w_blnColChk = True
                                End If
                            Next w_IntCol

                            If g_HopeNum <= w_KibouCnt And w_blnColChk = False Then
                                '��]�񐔐����I�[�o�[�_�C�A���O�\��
                                If g_KibouNumDiaLogFlg = 1 Then
                                    '���[�j���O
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "��]�Ζ�"
                                    w_strMsg(2) = "�ݒ肳�ꂽ��]�Ζ���"
                                    '�u&1��&2�𒴂��Ă��܂��B~n���̂܂ܓo�^���Ă���낵���ł����B�v
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    '�G���[
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "�ݒ肳�ꂽ��]�Ζ��񐔂𒴂��Ă��邽��"
                                    '�u&1���͂ł��܂���B�v
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If

                        '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
                        If CDbl(g_HopeNumDateFlg) = 1 And _OptRiyu_2.Checked = True Then
                            '���t�ʂ̊�]�Ζ����`�F�b�N
                            w_KibouCntDate = frmNSK0000HA.Get_HopeNum_Of_Date(w_YYYYMMDD)
                            If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor) = w_lngBackColor Then
                                w_KibouCntDate = w_KibouCntDate - 1
                            End If

                            If g_HopeNumDate <= w_KibouCntDate Then
                                '��]�񐔐����I�[�o�[�_�C�A���O�\��
                                If g_KibouNumDiaLogFlg = 1 Then
                                    '���[�j���O
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "��]�Ζ�"
                                    w_strMsg(2) = "�ݒ肳�ꂽ��]�Ζ�(���t��)��"
                                    '�u&1��&2�𒴂��Ă��܂��B~n���̂܂ܓo�^���Ă���낵���ł����B�v
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    '�G���[
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "�ݒ肳�ꂽ��]�Ζ�(���t��)�񐔂𒴂��Ă��邽��"
                                    '�u&1���͂ł��܂���B�v
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If
                        '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------

                        Select Case True
                            Case _OptRiyu_0.Checked '�ʏ�
                                '���R�敪 �ʏ�
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                            Case _OptRiyu_1.Checked '�v��
                                '���R�敪 �v��
                                w_RiyuKBN = "2"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", General.G_PALEGREEN))
                            Case _OptRiyu_2.Checked '��]
                                '���R�敪 ��]
                                w_RiyuKBN = "3"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                            Case _OptRiyu_3.Checked
                                '���R�敪 �Čf
                                w_RiyuKBN = "4"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", General.G_LIGHTCYAN))
                            Case _OptRiyu_4.Checked
                                w_RiyuKBN = "6"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", General.G_ORANGE))
                            Case Else
                                '���R�敪 ���̑��i�ʏ툵���Ƃ���j
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                        End Select

                        '�敪�������̏ꍇ�̂݁A������Ζ��n�I����ʂ�\��
                        If w_RiyuKBN = "6" Then
                            If Disp_Ouen(w_KangoCD) = False Then
                                Exit Sub
                            End If
                        End If

                        '2015/04/13 Bando Upd Start ======================
                        '��]o�̏ꍇ�R�����g���͉�ʕ\��
                        If w_RiyuKBN = "3" And g_InputHopeCommentFlg = "1" Then
                            If Disp_Comment(w_Comment, w_RiyuKBN) = False Then
                                Exit Sub
                            End If
                        End If
                        '2015/04/13 Bando Upd End   ======================

                        '2015/07/22 Bando Add Start ========================
                        If w_Comment <> "" Then
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Fore", Convert.ToString(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Back", ColorTranslator.ToOle(Color.IndianRed)))
                        End If
                        '2015/07/22 Bando Add End   ========================

                        '��x�����Ζ��ޯ��װ
                        If General.g_lngDaikyuMng = 0 Then
                            If m_DaikyuBackColorFlg = True Then
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Back", General.G_LIGHTYELLOW))
                            End If
                        End If

                        '�قɋL����ݒ肷��
                        '2015/04/13 Bando Upd Start ====================
                        'w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, w_Comment)
                        '2015/04/13 Bando Upd End   ====================
                        .SetText(w_ActiveRow, w_ActiveCol, w_Var)
                        '���R�ʂ̐F�ݒ�
                        .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                        .Cells(w_ActiveRow, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        '�Ζ��ύX�̎�
                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            '2015/04/13 Bando Upd Start ====================
                            'Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time)
                            Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time, w_Comment)
                            '2015/04/13 Bando Upd End   ====================
                            '�\��ƈقȂ�ꍇ�F�ύX
                            If Trim(w_KinmuCD) <> Trim(w_KinmuPlanCD) Or w_RiyuKBN <> w_RiyuPlanKbn Or w_KangoCD <> w_KangoPlanCD Then
                                .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(m_Comp_Fore)
                            End If

                            '�ύX�\�s�̓��e���i�[
                            w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                            w_ForeColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).ForeColor)
                            w_BackColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor)
                            '�ύX�\�s�̓��e���Ζ��ύX��ʂɓ\��t���s�փR�s�[
                            .SetText(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol, w_Var)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        End If
                    End With
                End If
            End If

            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
            If InStr(m_strUpdKojyoDate, w_Var) = 0 Then
                m_strUpdKojyoDate = m_strUpdKojyoDate & w_Var & ","
            End If

            '���پوړ�
            If w_DaikyuInputFlg = True Then
                Call SetCursol()
            End If

            '�X�V�׸޾��
            If w_InputFlg = True Then
                m_KosinFlg = True
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    'Private Sub m_lstCmdYasumi_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _cmdYasumi_0.Click, _cmdYasumi_2.Click, _cmdYasumi_4.Click, _cmdYasumi_6.Click, _
    '                                                                                                                    _cmdYasumi_8.Click, _cmdYasumi_10.Click, _cmdYasumi_12.Click, _cmdYasumi_14.Click, _
    '                                                                                                                    _cmdYasumi_16.Click, _cmdYasumi_18.Click, _cmdYasumi_20.Click, _cmdYasumi_22.Click, _
    '                                                                                                                    _cmdYasumi_24.Click, _cmdYasumi_26.Click, _cmdYasumi_28.Click, _cmdYasumi_30.Click, _
    '                                                                                                                    _cmdYasumi_32.Click, _cmdYasumi_34.Click, _cmdYasumi_36.Click, _cmdYasumi_38.Click
    Private Sub m_lstCmdYasumi_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim Index As Short = m_lstCmdYasumi.IndexOf(eventSender)
        Const W_SUBNAME As String = "NSK0000HC  m_lstCmdYasumi_Click"

        Dim w_RegStr As String
        Dim w_Var As Object
        Dim w_ActiveCol As Integer
        Dim w_ActiveRow As Integer
        Dim w_KinmuCD As String 'KinmuCD
        Dim w_RiyuKBN As String '���R�敪
        Dim w_Time As String '���ԔN�x
        Dim w_Flg As String '�m���׸�
        Dim w_ForeColor As Integer '�����F
        Dim w_BackColor As Integer '�w�i�F
        Dim w_InputFlg As Boolean '�����׸�
        Dim w_Cnt As Short
        Dim w_KinmuPlanCD As String 'KinmuCD(�\���ް�)
        Dim w_STS As Short
        Dim w_KangoCD As String
        Dim w_KangoPlanCD As String
        Dim w_RiyuPlanKbn As String
        Dim w_YYYYMMDD As Integer
        Dim w_DaikyuInputFlg As Boolean
        Dim w_strMsg() As String
        Dim w_lngBackColor As Integer
        Dim w_IntCol As Short
        Dim w_KibouCnt As Short
        Dim w_KibouCol() As Integer
        Dim w_blnColChk As Boolean
        Dim w_objSyouninData As Object
        '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
        Dim w_KibouCntDate As Integer
        '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------
        Dim w_Comment As String = String.Empty  '��]�Ζ����̃R�����g 2015/04/13 Bando Add
        Try
            '̫����ړ�
            sprSheet.Focus()

            'ڼ޽�؊i�[��
            w_RegStr = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            '�P���ł����݂���΁E�E�E
            If m_YasumiCnt <> 0 Then

                w_ActiveCol = sprSheet.Sheets(0).ActiveColumn.Index
                w_ActiveRow = sprSheet.Sheets(0).ActiveRow.Index

                '���͏ꏊ����
                If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row_ChgJisseki Then
                        Exit Sub
                    End If
                Else
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row Then
                        Exit Sub
                    End If
                End If

                '�����ٓ������i�z���͈́j
                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_ActiveCol).BackColor) = m_MonthBefore_Back Then
                    '�z�����ԊO(�w�i�F���O���[)�̏ꍇ���͕s��
                    Exit Sub
                End If

                '�����׸�
                w_InputFlg = True
                w_Cnt = CShort(w_ActiveCol - M_KinmuData_Col + 1)
                If g_SaikeiFlg = True Then
                    '�Čf�����̏ꍇ
                    If m_DataFlg(w_Cnt) = "1" Then
                        '�����ް��̏ꍇ
                        ReDim w_strMsg(2)
                        w_strMsg(1) = "�m��ς݋Ζ�"
                        w_strMsg(2) = "�Čf�Ζ�"
                        Call General.paMsgDsp("NS0011", w_strMsg)

                        w_InputFlg = False
                    End If
                End If

                '�Ζ��̗j�������`�F�b�N
                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                '2013/10/02 Bando Chg Start ===================================================
                'If Check_YoubiLimit(w_YYYYMMDD, m_Yasumi(Index + HscYasumi.Value).CD) = False Then
                If Check_YoubiLimit(w_YYYYMMDD, m_Yasumi(Index + HscYasumi.Value * 2).CD) = False Then
                    '2013/10/02 Bando Chg End  ===================================================
                    Exit Sub
                End If

                '2013/01/07 Ishiga add start------------------------------------------
                '���΃f�[�^�̗L���`�F�b�N
                If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If Check_OverKinmuData(w_YYYYMMDD) = False Then
                        Exit Sub
                    End If
                End If
                '2013/01/07 Ishiga add end--------------------------------------------

                '�͏o���݃`�F�b�N
                If fncChkAppliData(w_YYYYMMDD) = False Then
                    Exit Sub
                End If

                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                If m_PackageFLG = 1 Or m_PackageFLG = 3 Then
                    '�������̃R���|�[�l���g����̎�
                    With General.g_objGetData
                        .p�E���敪 = 0
                        .p�E���ԍ� = M_StaffID
                        .p�`�F�b�N��� = w_YYYYMMDD
                        .p�����敪 = 0
                        '2013/10/02 Bando Chg Start ==============================
                        '.p�`�F�b�N�Ζ�CD = m_Yasumi(Index + HscYasumi.Value).CD
                        .p�`�F�b�N�Ζ�CD = m_Yasumi(Index + HscYasumi.Value * 2).CD
                        '2013/10/02 Bando Chg End   ==============================

                        If .mChkKinmuDuty = False Then
                            '�Ζ��ύX�s��
                            '*******ү����***********************************
                            ReDim w_strMsg(1)
                            w_strMsg(1) = ""
                            Call General.paMsgDsp("NS0110", w_strMsg)
                            '************************************************
                            Exit Sub
                        End If
                    End With
                End If

                w_DaikyuInputFlg = True

                If General.g_lngDaikyuMng = 0 Then
                    '��x����
                    w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                    w_YYYYMMDD = w_Var
                    '2013/10/02 Bando Chg Start ======================================================
                    'If Check_Daikyu(M_PASTE, w_YYYYMMDD, m_Yasumi(Index + HscYasumi.Value).CD) = False Then
                    If Check_Daikyu(M_PASTE, w_YYYYMMDD, m_Yasumi(Index + HscYasumi.Value * 2).CD) = False Then
                        '2013/10/02 Bando Chg End   ======================================================
                        w_DaikyuInputFlg = False
                        w_InputFlg = False
                    End If
                End If

                If w_InputFlg = True Then

                    With sprSheet.Sheets(0)
                        '�x���\��+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        '�ύX����Ζ��̒l����Ōv��ύX�̏ꍇ�A���̋Ζ����擾����B
                        w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                        If w_Var = "" And m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            w_Var = .GetText(w_ActiveRow - 1, w_ActiveCol)
                        End If

                        '2015/04/13 Bando Upd Start ==========================
                        'Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, "")
                        '2015/04/13 Bando Upd End   ==========================

                        If g_SaikeiFlg = False Then
                            Select Case w_RiyuKBN
                                Case "2"
                                    If frmNSK0000HA._mnuTool_5.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "�v���Ζ�"
                                        w_strMsg(2) = "�ҏW"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "3"
                                    If frmNSK0000HA._mnuTool_4.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "��]�Ζ�"
                                        w_strMsg(2) = "�ҏW"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "4"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�Čf�Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "5"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�ψ���Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "6"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�����Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                            End Select
                        Else
                            If w_RiyuKBN = "2" Then
                                If frmNSK0000HA._mnuTool_5.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�v���Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If

                            If w_RiyuKBN = "3" Then
                                If frmNSK0000HA._mnuTool_4.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "��]�Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If

                        w_KinmuCD = m_Yasumi(Index + HscYasumi.Value * 2).CD
                        w_Time = ""
                        w_Flg = "0"
                        w_KangoCD = ""

                        '��]�񐔏W�v�`�F�b�N
                        w_lngBackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                        If CDbl(g_HopeNumFlg) = 1 And _OptRiyu_2.Checked = True Then
                            w_KibouCnt = 0
                            ReDim w_KibouCol(0)
                            For w_IntCol = m_KeikakuD_StartCol To m_KeikakuD_EndCol
                                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_IntCol).BackColor) = w_lngBackColor Then
                                    w_KibouCnt = w_KibouCnt + 1
                                    ReDim Preserve w_KibouCol(w_KibouCnt)
                                    w_KibouCol(w_KibouCnt) = w_IntCol
                                End If
                            Next w_IntCol

                            '�����ꏊ�ɓ����Ζ���\��t�����ꍇ�A�X���[
                            w_blnColChk = False
                            For w_IntCol = 1 To UBound(w_KibouCol)
                                If w_ActiveCol = w_KibouCol(w_IntCol) And w_KinmuCD = w_KinmuPlanCD Then
                                    w_blnColChk = True
                                End If
                            Next w_IntCol

                            If g_HopeNum <= w_KibouCnt And w_blnColChk = False Then
                                '��]�񐔐����I�[�o�[�_�C�A���O�\��
                                If g_KibouNumDiaLogFlg = 1 Then
                                    '���[�j���O
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "��]�Ζ�"
                                    w_strMsg(2) = "�ݒ肳�ꂽ��]�Ζ���"
                                    '�u&1��&2�𒴂��Ă��܂��B~n���̂܂ܓo�^���Ă���낵���ł����B�v
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    '�G���[
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "�ݒ肳�ꂽ��]�Ζ��񐔂𒴂��Ă��邽��"
                                    '�u&1���͂ł��܂���B�v
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If

                        '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
                        If CDbl(g_HopeNumDateFlg) = 1 And _OptRiyu_2.Checked = True Then
                            '���t�ʂ̊�]�Ζ����`�F�b�N
                            w_KibouCntDate = frmNSK0000HA.Get_HopeNum_Of_Date(w_YYYYMMDD)
                            If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor) = w_lngBackColor Then
                                w_KibouCntDate = w_KibouCntDate - 1
                            End If

                            If g_HopeNumDate <= w_KibouCntDate Then
                                '��]�񐔐����I�[�o�[�_�C�A���O�\��
                                If g_KibouNumDiaLogFlg = 1 Then
                                    '���[�j���O
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "��]�Ζ�"
                                    w_strMsg(2) = "�ݒ肳�ꂽ��]�Ζ�(���t��)��"
                                    '�u&1��&2�𒴂��Ă��܂��B~n���̂܂ܓo�^���Ă���낵���ł����B�v
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    '�G���[
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "�ݒ肳�ꂽ��]�Ζ�(���t��)�񐔂𒴂��Ă��邽��"
                                    '�u&1���͂ł��܂���B�v
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If
                        '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------

                        Select Case True
                            Case _OptRiyu_0.Checked
                                '���R�敪 �ʏ�
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                            Case _OptRiyu_1.Checked
                                '���R�敪 �v��
                                w_RiyuKBN = "2"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", General.G_PALEGREEN))
                            Case _OptRiyu_2.Checked
                                '���R�敪 ��]
                                w_RiyuKBN = "3"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                            Case _OptRiyu_3.Checked
                                '���R�敪 �Čf
                                w_RiyuKBN = "4"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", General.G_LIGHTCYAN))
                            Case _OptRiyu_4.Checked
                                w_RiyuKBN = "6"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", General.G_ORANGE))
                            Case Else
                                '���R�敪 ���̑��i�ʏ툵���Ƃ���j
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                        End Select

                        '�敪�������̏ꍇ�̂݁A������Ζ��n�I����ʂ�\��
                        If w_RiyuKBN = "6" Then
                            If Disp_Ouen(w_KangoCD) = False Then
                                Exit Sub
                            End If
                        End If

                        '2015/04/13 Bando Add Start =====================
                        '��]�̏ꍇ�R�����g���͉�ʕ\��
                        If w_RiyuKBN = "3" And g_InputHopeCommentFlg = "1" Then
                            If Disp_Comment(w_Comment, w_RiyuKBN) = False Then
                                Exit Sub
                            End If
                        End If
                        '2015/04/13 Bando Add End   =====================

                        '2015/07/22 Bando Add Start ========================
                        If w_Comment <> "" Then
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Fore", Convert.ToString(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Back", ColorTranslator.ToOle(Color.IndianRed)))
                        End If
                        '2015/07/22 Bando Add End   ========================

                        '��x�����Ζ��ޯ��װ
                        If General.g_lngDaikyuMng = 0 Then
                            If m_DaikyuBackColorFlg = True Then
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Back", General.G_LIGHTYELLOW))
                            End If
                        End If

                        '�قɋL����ݒ肷��
                        '2015/04/13 Bando Upd Start ===========================
                        'w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, w_Comment)
                        '2015/04/13 Bando Upd End   ===========================
                        .SetText(w_ActiveRow, w_ActiveCol, w_Var)

                        '���R�ʂ̐F�ݒ�
                        .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                        .Cells(w_ActiveRow, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)

                        '�Ζ��ύX�̎�
                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            '2015/04/13 Bando Upd Start =========================================
                            'Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time)
                            Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time, "")
                            '2015/04/13 Bando Upd End   =========================================

                            '�\��ƈقȂ�ꍇ�F�ύX
                            If Trim(w_KinmuCD) <> Trim(w_KinmuPlanCD) Or w_RiyuKBN <> w_RiyuPlanKbn Or w_KangoCD <> w_KangoPlanCD Then
                                .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(m_Comp_Fore)
                            End If

                            '�ύX�\�s�̓��e���i�[
                            w_Var = .GetText(w_ActiveRow, w_ActiveCol)

                            w_ForeColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).ForeColor)
                            w_BackColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor)

                            '�ύX�\�s�̓��e���Ζ��ύX��ʂɓ\��t���s�փR�s�[
                            .SetText(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol, w_Var)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        End If
                    End With
                End If
            End If

            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
            If InStr(m_strUpdKojyoDate, w_Var) = 0 Then
                m_strUpdKojyoDate = m_strUpdKojyoDate & w_Var & ","
            End If

            '���پوړ�
            If w_DaikyuInputFlg = True Then
                Call SetCursol()
            End If

            If w_InputFlg = True Then
                '�X�V�׸޾��
                m_KosinFlg = True
            End If
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Public Sub frmNSK0000HC_Load()

        Const W_SUBNAME As String = "NSK0000HC  Form_Load"

        Dim w_Int As Short
        Dim w_str As String
        Dim w_varWork As Object
        Try
            Me.Hide()

            m_strUpdKojyoDate = ""

            'ڼ޽�؊i�[��
            Const w_RegStr As String = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            '�X�v���b�h���蓖�ăL�[������
            subChgSpreadKeyMap()

            '�O�� ����/�w�i�F
            m_MonthBefore_Fore = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "MonthBefore_Fore", General.G_BLACK))
            m_MonthBefore_Back = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "MonthBefore_Back", General.G_LIGHTGRAY))

            '�v����ԊO��4�T���� �w�i�F
            m_Jisseki4W_Back = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Jisseki4W_Back", ColorTranslator.ToOle(Color.Cyan).ToString))

            '���т��\��ƈقȂ�ꍇ ����/�w�i�F
            m_Comp_Fore = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Comp_Fore", CStr(ColorTranslator.ToOle(Color.Red))))

            '�y���j �w�i�F
            m_WeekEnd_Back = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "WeekEnd_Back", ColorTranslator.ToOle(Color.LavenderBlush)))

            '--- �y���w�i�F�t���O
            m_WeekEndColorFlg = General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY7, "WEEKENDCOLORFLG", "0", General.g_strHospitalCD)
            '--- �j�x���w�i�F�t���O
            m_HolidayColorFlg = General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY7, "HOLIDAYCOLORFLG", "0", General.g_strHospitalCD)

            '��x�̗L�����Ԃ����߂�(��̫�Ă͂W�T��)
            m_lngDaikyuPastPeriod = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "PASTDAIKYUPERIOD", "56", General.g_strHospitalCD))
            m_DaikyuMsgFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "DAIKYUMSGFLG", CStr(0), General.g_strHospitalCD))
            m_SundayDaikyuFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "SUNDAYDAIKYUFLG", CStr(0), General.g_strHospitalCD))
            m_DaikyuAdvFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "DAIKYUADVANCEFLG", CStr(0), General.g_strHospitalCD))
            m_SaturdayDaikyuFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "SATURDAYDAIKYUFLG", CStr(0), General.g_strHospitalCD))

            '�����Ζ��敪�̕\��FLG
            m_OuenDispFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY7, "OUENDISPFLG", "1", General.g_strHospitalCD))

            '��x���蓖�������t���O
            m_DaikyuAdvThisMonthFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "DAIKYUADVANCETHISMONTHFLG", CStr(0), General.g_strHospitalCD))

            '1.5�����̑�x����������Ζ��b�c
            w_str = General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "DAIKYU15KINNMUCD", "", General.g_strHospitalCD)
            w_varWork = General.paSplit(w_str, ",")
            ReDim m_Daikyu15KinmuCD(UBound(w_varWork) + 1)
            For w_Int = 0 To UBound(w_varWork)
                m_Daikyu15KinmuCD(w_Int + 1) = w_varWork(w_Int)
            Next w_Int

            '2014/04/23 Saijo add start P-06979-----------------------------------
            '�Ζ��L���S�p�Q�����Ή��t���O(0�F�Ή����Ȃ��A1:�Ή�����)
            m_strKinmuEmSecondFlg = Get_ItemValue(General.g_strHospitalCD)
            '2014/04/23 Saijo add end P-06979-------------------------------------

            '2015/04/14 Bando Add Start ========================================
            '��]���[�h���̕\���ΏۋΖ�CD
            m_DispKinmuCd = General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY15, "DISPKINMUCD", "", General.g_strHospitalCD)
            '2015/04/14 Bando Add End   ========================================

            Call Get_PackageUseFLG()

            '���тŎg�p����ꍇ�́A���R�敪�̓��͍͂s��Ȃ��B1:�v��A2:����
            If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                '�Ζ��ύX�i���яC���j�̏ꍇ
                _OptRiyu_0.Enabled = True '�ʏ�
                _OptRiyu_0.Checked = True
                _OptRiyu_1.Enabled = False '�v��
                _OptRiyu_2.Enabled = False '��]
                _OptRiyu_3.Enabled = False '�Čf
                _OptRiyu_3.Visible = False

                If m_OuenDispFlg = 0 Then
                    _OptRiyu_4.Enabled = True '����
                Else
                    _OptRiyu_4.Enabled = False
                    _OptRiyu_4.Visible = False
                End If

                chkSet.CheckState = CheckState.Unchecked
                chkSet.Enabled = False
                _Frame_3.Enabled = False
                chkSet.Visible = False
                _Frame_3.Visible = False

                Me.Text = "�ʋΖ��ύX"

                '�����S���g�p�s��
                cmdErase.Enabled = False
                cmdErase.Visible = False
            Else
                '�v�� �̏ꍇ
                '�Čf�����̏ꍇ�́A�Čf�݂̂��g�p�ɂ���
                If g_SaikeiFlg = True Then
                    _OptRiyu_0.Enabled = False '�ʏ�
                    _OptRiyu_0.Visible = False
                    _OptRiyu_1.Enabled = False '�v��
                    _OptRiyu_2.Enabled = False '��]
                    _OptRiyu_3.Enabled = True '�Čf
                    _OptRiyu_3.Visible = True
                    _OptRiyu_3.Checked = True
                    _OptRiyu_4.Enabled = False '����

                    If m_OuenDispFlg = 1 Then
                        _OptRiyu_4.Visible = False
                    End If

                    chkSet.Visible = False
                    _Frame_3.Enabled = True
                Else
                    '�v�� �̏ꍇ
                    _OptRiyu_0.Enabled = True '�ʏ�
                    _OptRiyu_0.Checked = True
                    _OptRiyu_1.Enabled = True '�v��


                    '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                    'If g_HopeNumFlg = "1" And g_HopeNum = 0 Then
                    If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                        '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                        '��]�񐔐�������@���@��]��0��@�̏ꍇ
                        _OptRiyu_2.Enabled = False '��]
                    Else
                        '�ȊO
                        _OptRiyu_2.Enabled = True '��]
                    End If

                    _OptRiyu_3.Enabled = True '�Čf
                    _OptRiyu_3.Visible = False

                    If m_OuenDispFlg = 0 Then
                        _OptRiyu_4.Enabled = True '����
                    Else
                        _OptRiyu_4.Enabled = False
                        _OptRiyu_4.Visible = False
                    End If

                    chkSet.CheckState = CheckState.Unchecked
                    _Frame_3.Enabled = False
                End If

                Me.Text = "�ʋΖ��v��쐬"
                '�����S���g�p��
                cmdErase.Enabled = True
            End If

            '��]���̓��[�h�̏ꍇ,���R�敪��]�̂ݎg�p��
            If g_SaikeiFlg = False Then
                If g_LimitedFlg = True Then
                    _OptRiyu_0.Enabled = False '�ʏ�
                    _OptRiyu_1.Enabled = False '�v��

                    '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                    'If g_HopeNumFlg = "1" And g_HopeNum = 0 Then
                    If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                        '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                        '��]�񐔐�������@���@��]��0��@�̏ꍇ
                        _OptRiyu_2.Enabled = False '��]
                        _OptRiyu_2.Checked = False
                    Else
                        '�ȊO
                        _OptRiyu_2.Enabled = True '��]
                        _OptRiyu_2.Checked = True
                    End If

                    _OptRiyu_2.Enabled = True '��]
                    _OptRiyu_2.Checked = True
                    _OptRiyu_3.Enabled = False '�Čf
                    _OptRiyu_3.Visible = False
                    _OptRiyu_4.Enabled = False '����

                    If m_OuenDispFlg = 1 Then
                        _OptRiyu_4.Visible = False
                    End If

                    chkSet.Visible = False
                    _Frame_3.Enabled = True
                End If
            End If


            '2014/04/23 Saijo upd start P-06979---------------------------
            ''̫�Ļ��ނɂ����̫�т̕���ݒ�
            'Select Case m_FontSize
            'Case M_FontSize_Big
            '    Me.Width = General.paTwipsTopixels(14300)
            'Case M_FontSize_Middle
            '    Me.Width = General.paTwipsTopixels(12500)
            'Case M_FontSize_Small
            '    Me.Width = General.paTwipsTopixels(10750)
            'Case Else
            'End Select
            If m_strKinmuEmSecondFlg = "0" Then
                '̫�Ļ��ނɂ����̫�т̕���ݒ�
                Select Case m_FontSize
                    Case M_FontSize_Big
                        Me.Width = General.paTwipsTopixels(14300)
                    Case M_FontSize_Middle
                        Me.Width = General.paTwipsTopixels(12500)
                    Case M_FontSize_Small
                        Me.Width = General.paTwipsTopixels(10750)
                    Case Else
                End Select
            Else
                '̫�Ļ��ނɂ����̫�т̕���ݒ�
                Select Case m_FontSize
                    Case M_FontSize_Second_Big
                        Me.Width = General.paTwipsTopixels(14300)
                    Case M_FontSize_Second_Middle
                        Me.Width = General.paTwipsTopixels(12500)
                    Case M_FontSize_Second_Small
                        Me.Width = General.paTwipsTopixels(10750)
                    Case Else
                End Select
            End If
            '2014/04/23 Saijo upd end P-06979-----------------------------

            If Not m_empRowDispFlg Then
                '�̗p�������\��
                sprSheet.Sheets(0).SetColumnWidth(1, 0)
            End If
            '��ΐ�]�E�Z���ԏ��\��
            Call DispNightShortInfo()

            '�X�V�׸ޏ�����
            m_KosinFlg = False
            m_OKFlg = False

        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub frmNSK0000HC_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As FormClosingEventArgs) Handles Me.FormClosing
        Dim UnloadMode As CloseReason = eventArgs.CloseReason
        Const W_SUBNAME As String = "NSK0000HC  Form_QueryUnload"

        Dim w_strMsg() As String
        Dim w_MsgRc As Short
        Try
            'O.K.���݉�������
            If m_OKFlg = False Then
                If m_KosinFlg Then
                    'ү���ނ�\��
                    ReDim w_strMsg(0)
                    w_MsgRc = General.paMsgDsp("NS0041", w_strMsg)

                    Select Case w_MsgRc
                        Case MsgBoxResult.Yes
                            '�Ζ��\�t��
                            Call cmdOK_Click(cmdOK, New System.EventArgs())

                        Case MsgBoxResult.No
                            '�ύX�͔j������
                            m_KosinFlg = False
                            m_strUpdKojyoDate = ""

                        Case MsgBoxResult.Cancel
                            '�I�����~
                            eventArgs.Cancel = True
                            Exit Sub
                    End Select
                End If
            End If

        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '���̾قֶ��وړ�
    Private Sub SetCursol()

        Const W_SUBNAME As String = "NSK0000HC  SetCursol"

        Dim w_ActiveRow As Short
        Dim w_ActiveCol As Short
        Dim w_Col As Short
        Dim w_tmpCol As Short
        Dim w_flg As Boolean = False
        Try
            With sprSheet.Sheets(0)
                w_ActiveRow = .ActiveRow.Index
                w_ActiveCol = .ActiveColumn.Index

                For w_Col = w_ActiveCol + 1 To m_KeikakuD_EndCol
                    '�وʒu�ݒ�
                    If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_MonthBefore_Back And ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_Jisseki4W_Back Then
                        If Not IsExistBackColor(sprSheet, w_ActiveRow, w_Col) Then
                            '�Ζ��\�t���ς�����
                            If Trim(.GetText(w_ActiveRow, w_Col)) = "" Then
                                .SetActiveCell(w_ActiveRow, w_Col)
                                Exit Sub
                            End If
                        End If
                    End If
                Next w_Col

                '�ړ��ق����݂��Ȃ��ꍇ

                '�وʒu�ݒ�
                If w_ActiveCol + 1 <= m_KeikakuD_EndCol Then
                    For w_Col = w_ActiveCol + 1 To m_KeikakuD_EndCol
                        '�وʒu�ݒ�
                        If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_MonthBefore_Back And ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_Jisseki4W_Back Then
                            If Not IsExistBackColor(sprSheet, w_ActiveRow, w_Col) Then
                                .SetActiveCell(w_ActiveRow, w_Col)
                                Exit Sub
                            End If
                        End If
                    Next w_Col

                    For w_Col = m_KeikakuD_StartCol To m_KeikakuD_EndCol
                        '�وʒu�ݒ�
                        If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_MonthBefore_Back And ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_Jisseki4W_Back Then
                            w_tmpCol = w_Col
                            If Not IsExistBackColor(sprSheet, w_ActiveRow, w_Col) Then
                                w_flg = True
                                .SetActiveCell(w_ActiveRow, w_Col)
                                Exit For
                            End If
                        End If
                    Next w_Col

                    If Not w_flg Then
                        .SetActiveCell(w_ActiveRow, w_tmpCol)
                    End If

                    w_ActiveRow = .ActiveRow.Index
                    w_ActiveCol = .ActiveColumn.Index

                    If Trim(.GetText(w_ActiveRow, w_ActiveCol)) <> "" Then
                        For w_Col = w_ActiveCol + 1 To m_KeikakuD_EndCol
                            '�وʒu�ݒ�
                            If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_MonthBefore_Back And ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_Jisseki4W_Back Then
                                If Not IsExistBackColor(sprSheet, w_ActiveRow, w_Col) Then
                                    '�Ζ��\�t���ς�����
                                    If Trim(.GetText(w_ActiveRow, w_Col)) = "" Then
                                        .SetActiveCell(w_ActiveRow, w_Col)
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next w_Col
                    End If
                Else
                    For w_Col = m_KeikakuD_StartCol To m_KeikakuD_EndCol
                        '�وʒu�ݒ�
                        If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_MonthBefore_Back And ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_Jisseki4W_Back Then
                            w_tmpCol = w_Col
                            If Not IsExistBackColor(sprSheet, w_ActiveRow, w_Col) Then
                                w_flg = True
                                .SetActiveCell(w_ActiveRow, w_Col)
                                Exit For
                            End If
                        End If
                    Next w_Col

                    If Not w_flg Then
                        .SetActiveCell(w_ActiveRow, w_tmpCol)
                    End If

                    w_ActiveRow = .ActiveRow.Index
                    w_ActiveCol = .ActiveColumn.Index

                    If Trim(.GetText(w_ActiveRow, w_ActiveCol)) <> "" Then
                        For w_Col = w_ActiveCol + 1 To m_KeikakuD_EndCol
                            '�وʒu�ݒ�
                            If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_MonthBefore_Back And ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_Jisseki4W_Back Then
                                If Not IsExistBackColor(sprSheet, w_ActiveRow, w_Col) Then
                                    '�Ζ��\�t���ς�����
                                    If Trim(.GetText(w_ActiveRow, w_Col)) = "" Then
                                        .SetActiveCell(w_ActiveRow, w_Col)
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next w_Col
                    End If
                End If

            End With

        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '�������وʒu�ݒ�
    Private Sub SetStartCursol()

        Const W_SUBNAME As String = "NSK0000HC  SetStartCursol"

        Dim w_Col As Short
        Dim w_Row As Short
        Dim w_Color As Integer
        Dim w_ActiveRow As Short

        Try
            With sprSheet.Sheets(0)
                w_ActiveRow = .ActiveRow.Index

                '�وʒu�ݒ�i���ѕύX���͗\����Q�Ƃ���j
                If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    w_Row = M_KinmuData_Row_ChgJisseki
                Else
                    w_Row = w_ActiveRow
                End If

                For w_Col = M_KinmuData_Col To m_KeikakuD_EndCol
                    '�Z���J���[�擾
                    w_Color = ColorTranslator.ToOle(.Cells(w_Row, w_Col).BackColor)
                    If w_Color <> m_MonthBefore_Back And w_Color <> m_Jisseki4W_Back Then
                        If Not IsExistBackColor(sprSheet, w_Row, w_Col) Then
                            '�Ζ��\�t���ς�����
                            If Trim(.GetText(w_Row, w_Col)) = "" Then
                                .SetActiveCell(w_Row, w_Col)
                                Exit Sub
                            End If
                        End If
                    End If
                Next w_Col

                For w_Col = M_KinmuData_Col To m_KeikakuD_EndCol
                    '�Z���J���[�擾
                    w_Color = ColorTranslator.ToOle(.Cells(w_Row, w_Col).BackColor)
                    If w_Color <> m_MonthBefore_Back And w_Color <> m_Jisseki4W_Back Then
                        If Not IsExistBackColor(sprSheet, w_Row, w_Col) Then
                            .SetActiveCell(w_Row, w_Col)
                            Exit Sub
                        End If
                    End If
                Next w_Col

                .SetActiveCell(w_Row, M_KinmuData_Col)
            End With

        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub frmNSK0000HC_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As FormClosedEventArgs) Handles Me.FormClosed

        Const W_SUBNAME As String = "NSK0000HC  Form_Unload"
        Try
            '����޳�̕\���߼޼�݂��i�[����
            Call General.paPutWindowPositon(Me, General.G_STRMAINKEY2 & "\" & g_AppName)

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End Try
    End Sub

    Private Sub HscKinmu_Change(ByVal newScrollValue As Integer)

        Const W_SUBNAME As String = "NSK0000HC  HscKinmu_Change"

        '�X�N���[���o�[�̍X�V
        Dim w_i As Short
        Dim w_Hsc_Cnt As Short
        Dim w_int As Integer
        Try
            '�R�}���h�{�^���̂b�`�o�s�h�n�m�ݒ�
            '�Ζ�
            w_Hsc_Cnt = newScrollValue
            For w_i = 1 To M_PARET_NUM
                w_int = w_i + w_Hsc_Cnt * 2
                If w_int <= m_KinmuCnt Then
                    m_lstCmdKinmu(w_i - 1).Text = m_Kinmu(w_int - 1).Mark
                    If m_Kinmu(w_int - 1).Setumei = "" Then
                        ToolTip1.SetToolTip(m_lstCmdKinmu(w_i - 1), Get_KinmuTipText(m_Kinmu(w_int - 1).CD))
                    Else
                        ToolTip1.SetToolTip(m_lstCmdKinmu(w_i - 1), Get_KinmuTipText(m_Kinmu(w_int - 1).CD) & "�F" & m_Kinmu(w_int - 1).Setumei)
                    End If
                    m_lstCmdKinmu(w_i - 1).Enabled = True
                    m_lstCmdKinmu(w_i - 1).visible = True
                Else
                    m_lstCmdKinmu(w_i - 1).Text = ""
                    m_lstCmdKinmu(w_i - 1).Enabled = False
                    m_lstCmdKinmu(w_i - 1).visible = False
                End If
            Next w_i

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub HscSet_Change(ByVal newScrollValue As Integer)

        Const W_SUBNAME As String = "NSK0000HC  HscSet_Change"

        '�X�N���[���o�[�̍X�V
        Dim w_i As Short
        Dim w_Hsc_Cnt As Short
        Dim w_int As Integer
        Try
            '�R�}���h�{�^���̂b�`�o�s�h�n�m�ݒ�
            '����Ζ�
            w_Hsc_Cnt = newScrollValue
            For w_i = 1 To M_PARET_NUM_SET
                w_int = w_i + w_Hsc_Cnt * 2
                If w_int <= m_SetCnt Then
                    m_lstCmdSet(w_i - 1).Text = m_SetKinmu(w_int - 1).Mark
                    ToolTip1.SetToolTip(m_lstCmdSet(w_i - 1), Get_SetKinmuTipText(w_int - 1))
                    m_lstCmdSet(w_i - 1).Enabled = True
                    m_lstCmdSet(w_i - 1).visible = True
                Else
                    m_lstCmdSet(w_i - 1).Text = ""
                    m_lstCmdSet(w_i - 1).Enabled = False
                    m_lstCmdSet(w_i - 1).visible = False
                End If
            Next w_i

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub HscTokusyu_Change(ByVal newScrollValue As Integer)

        Const W_SUBNAME As String = "NSK0000HC  HscTokusyu_Change"

        '�X�N���[���o�[�̍X�V
        Dim w_i As Short
        Dim w_Hsc_Cnt As Short
        Dim w_int As Integer
        Try
            '�R�}���h�{�^���̂b�`�o�s�h�n�m�ݒ�
            '����Ζ�
            w_Hsc_Cnt = newScrollValue
            For w_i = 1 To M_PARET_NUM
                w_int = w_i + w_Hsc_Cnt * 2
                If w_int <= m_TokusyuCnt Then
                    m_lstCmdTokusyu(w_i - 1).Text = m_Tokusyu(w_int - 1).Mark
                    If m_Tokusyu(w_int - 1).Setumei = "" Then
                        ToolTip1.SetToolTip(m_lstCmdTokusyu(w_i - 1), Get_KinmuTipText(m_Tokusyu(w_int - 1).CD))
                    Else
                        ToolTip1.SetToolTip(m_lstCmdTokusyu(w_i - 1), Get_KinmuTipText(m_Tokusyu(w_int - 1).CD) & "�F" & m_Tokusyu(w_int - 1).Setumei)
                    End If
                    m_lstCmdTokusyu(w_i - 1).Enabled = True
                    m_lstCmdTokusyu(w_i - 1).visible = True
                Else
                    m_lstCmdTokusyu(w_i - 1).Text = ""
                    m_lstCmdTokusyu(w_i - 1).Enabled = False
                    m_lstCmdTokusyu(w_i - 1).visible = False
                End If
            Next w_i

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub HscYasumi_Change(ByVal newScrollValue As Integer)
        Const W_SUBNAME As String = "NSK0000HC  HscYasumi_Change"

        '�X�N���[���o�[�̍X�V
        Dim w_i As Short
        Dim w_Hsc_Cnt As Short
        Dim w_int As Integer
        Try
            '�R�}���h�{�^���̂b�`�o�s�h�n�m�ݒ�
            '�x��
            w_Hsc_Cnt = newScrollValue
            For w_i = 1 To M_PARET_NUM
                w_int = w_i + w_Hsc_Cnt * 2
                If w_int <= m_YasumiCnt Then
                    m_lstCmdYasumi(w_i - 1).Text = m_Yasumi(w_int - 1).Mark
                    If m_Yasumi(w_int - 1).Setumei = "" Then
                        ToolTip1.SetToolTip(m_lstCmdYasumi(w_i - 1), Get_KinmuTipText(m_Yasumi(w_int - 1).CD))
                    Else
                        ToolTip1.SetToolTip(m_lstCmdYasumi(w_i - 1), Get_KinmuTipText(m_Yasumi(w_int - 1).CD) & "�F" & m_Yasumi(w_int - 1).Setumei)
                    End If
                    m_lstCmdYasumi(w_i - 1).Enabled = True
                    m_lstCmdYasumi(w_i - 1).visible = True
                Else
                    m_lstCmdYasumi(w_i - 1).Text = ""
                    m_lstCmdYasumi(w_i - 1).Enabled = False
                    m_lstCmdYasumi(w_i - 1).visible = False
                End If
            Next w_i

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '�̗pCD���擾 (���ԔN�x�擾���g�p)
    Private Function Get_SaiyoCD(ByVal p_Date As Integer, ByVal p_StaffID As String) As String
        Const W_SUBNAME As String = "NSK0000HC Get_SaiyoCD"


        Dim w_RecCnt As Integer
        Try
            '�̗pCD���擾
            General.g_objGetData.p�a�@CD = General.g_strHospitalCD
            General.g_objGetData.p�E���ԍ� = p_StaffID '�E���Ǘ��ԍ�
            General.g_objGetData.p���t�敪 = 0 '���t�͒P������w��
            General.g_objGetData.p�J�n�N���� = p_Date '�J�n�N����
            General.g_objGetData.p�����\�[�gFLG = 1 '�~��

            If General.g_objGetData.mStaffInit = False Then
                Get_SaiyoCD = ""
            Else
                w_RecCnt = General.g_objGetData.f�E���Ǘ�����

                '�P����w��Ȃ̂ŕK���P���ɂȂ�
                General.g_objGetData.p�E���Ǘ����� = 1

                Get_SaiyoCD = General.g_objGetData.f�̗p����CD
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    Private Sub sprSheet_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles sprSheet.CellClick
        Const W_SUBNAME As String = "NSK0000HC sprSheet_CellClick"

        Dim w_ToolTipText As String = ""

        Try
            If e.Column = 1 AndAlso e.Row = M_KinmuData_Row Then
                '�̗p�����񉟉����A�c�[���`�b�v��\������
                If TypeOf (sprSheet.Sheets(0).GetCellType(e.Row, e.Column)) Is CellType.ButtonCellType Then
                    ToolTip1.Show(m_toolTipTxt, sprSheet)
                End If
            End If
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub sprSheet_MouseMoveEvent(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles sprSheet.MouseMove
        Const W_SUBNAME As String = "NSK0000HC sprSheet_MouseMove"

        Dim w_Col As Integer
        Dim w_Row As Integer
        Dim w_Kinmu As Object
        Dim w_str As String
        Dim w_StrData As Object
        Dim w_KBN As String
        Dim w_KangoCD As String
        Dim w_CellRange As Model.CellRange

        Static s_Col As Integer
        Static s_Row As Integer
        Try
            'ϳ��ʒu��^�s�擾
            w_CellRange = sprSheet.GetCellFromPixel(0, 0, eventArgs.X, eventArgs.Y)
            w_Row = w_CellRange.Row
            w_Col = w_CellRange.Column

            If (w_Col = s_Col) And (w_Row = s_Row) Then
                Exit Sub
            End If

            s_Col = w_Col
            s_Row = w_Row

            '�v����͔͈͓��i��j
            If (w_Col < m_KeikakuD_StartCol) Or (w_Col > m_KeikakuD_EndCol) Then
                ToolTip1.SetToolTip(sprSheet, "")
                Exit Sub
            End If

            '�v����͔͈͓��i�s�j
            If (w_Row < M_KinmuData_Row) Then
                ToolTip1.SetToolTip(sprSheet, "")
                Exit Sub
            End If

            '�Ζ��擾
            w_StrData = sprSheet.Sheets(0).GetText(w_Row, w_Col)

            w_str = General.paRight(w_StrData, 11)

            '�Ζ���������
            w_Kinmu = Trim(General.paLeft(w_str, 3))
            If w_Kinmu = "" Then
                ToolTip1.SetToolTip(sprSheet, "")
                Exit Sub
            End If

            '������
            ToolTip1.SetToolTip(sprSheet, "")

            '�敪��"6"�̏ꍇ�̂݉����Ō�P�ʖ��̂��擾��°����߂ɐݒ�
            w_KBN = Mid(w_str, 4, 1)
            If w_KBN = "6" Then
                w_KangoCD = Trim(Mid(w_str, 8))
                ToolTip1.SetToolTip(sprSheet, "������ : " & w_KangoCD)
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '�j�������`�F�b�N
    Private Function Check_YoubiLimit(ByVal p_Date As Object, ByVal p_KinmuCD As String) As Boolean
        Const W_SUBNAME As String = "NSK0000HA Check_YoubiLimit"

        Dim w_intLoop1 As Short
        Dim w_intLoop2 As Short
        Dim w_chkValue As String
        Dim w_strYoubi As String
        Dim w_strMsg() As String

        Check_YoubiLimit = False
        Try
            '�j�����擾
            If InStr(m_HolDateStr, p_Date) > 0 Then
                '�j���E�x���̏ꍇ�͗j���G���[�`�F�b�N���s��Ȃ�
                If InStr(m_OffDayStr, p_Date) > 0 Then
                    w_chkValue = "8"
                    w_strYoubi = "�x��"
                Else
                    w_chkValue = "7"
                    w_strYoubi = "�j��"
                End If

                '�j�������`�F�b�N
                For w_intLoop1 = 1 To UBound(g_KinmuM)
                    If g_KinmuM(w_intLoop1).CD = p_KinmuCD Then
                        For w_intLoop2 = 1 To UBound(g_KinmuM(w_intLoop1).YoubiLimit)
                            If w_chkValue = g_KinmuM(w_intLoop1).YoubiLimit(w_intLoop2) Then
                                '�����Ώۂ̏ꍇ�A��������
                                ReDim w_strMsg(2)
                                w_strMsg(1) = w_strYoubi
                                w_strMsg(2) = g_KinmuM(w_intLoop1).KinmuName
                                Call General.paMsgDsp("NS0011", w_strMsg)
                                Exit Function
                            End If
                        Next w_intLoop2
                        Exit For
                    End If
                Next w_intLoop1
            Else

                Select Case Weekday(CDate(Format(Integer.Parse(p_Date), "0000/00/00")))
                    Case FirstDayOfWeek.Monday
                        w_chkValue = "0"
                        w_strYoubi = "���j��"
                    Case FirstDayOfWeek.Tuesday
                        w_chkValue = "1"
                        w_strYoubi = "�Ηj��"
                    Case FirstDayOfWeek.Wednesday
                        w_chkValue = "2"
                        w_strYoubi = "���j��"
                    Case FirstDayOfWeek.Thursday
                        w_chkValue = "3"
                        w_strYoubi = "�ؗj��"
                    Case FirstDayOfWeek.Friday
                        w_chkValue = "4"
                        w_strYoubi = "���j��"
                    Case FirstDayOfWeek.Saturday
                        w_chkValue = "5"
                        w_strYoubi = "�y�j��"
                    Case FirstDayOfWeek.Sunday
                        w_chkValue = "6"
                        w_strYoubi = "���j��"
                End Select

                '�j�������`�F�b�N
                For w_intLoop1 = 1 To UBound(g_KinmuM)
                    If g_KinmuM(w_intLoop1).CD = p_KinmuCD Then
                        For w_intLoop2 = 1 To UBound(g_KinmuM(w_intLoop1).YoubiLimit)
                            If w_chkValue = g_KinmuM(w_intLoop1).YoubiLimit(w_intLoop2) Then
                                '�����Ώۂ̏ꍇ�A��������
                                ReDim w_strMsg(2)
                                w_strMsg(1) = w_strYoubi
                                w_strMsg(2) = g_KinmuM(w_intLoop1).KinmuName
                                Call General.paMsgDsp("NS0011", w_strMsg)
                                Exit Function
                            End If
                        Next w_intLoop2
                        Exit For
                    End If
                Next w_intLoop1
            End If
            Check_YoubiLimit = True

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    '�p�b�P�[�W��� �ް��擾
    Private Function Get_PackageUseFLG() As Boolean

        Const W_SUBNAME As String = "NSK0000HA Get_PackageUseFLG"

        Dim w_lngRecCnt As Integer
        Dim w_lngLoop As Integer
        Dim w_strSql As String
        Dim w_Rs As ADODB.Recordset
        Dim w_CD_F As ADODB.Field
        Dim w_FLG_F As ADODB.Field
        Dim w_strAppliUseFlg As String
        Dim w_strDutyUseFlg As String
        Dim w_strMsg() As String
        Try
            w_strAppliUseFlg = "0"
            w_strDutyUseFlg = "0"

            ''�p�b�P�[�WM���͏o�Ɠ�������USEFLG���擾
            'w_strSql = ""
            'w_strSql = "Select PACKAGECD, USEFLG"
            'w_strSql = w_strSql & " From NS_PACKAGE_M"
            'w_strSql = w_strSql & " Where HospitalCD = '" & General.g_strHospitalCD & "'"

            ''ں��޾�ĵ�޼ު�� ����
            'w_Rs = General.paDBRecordSetOpen(w_strSql)

            Call NSK0000H_sql.select_NS_PACKAGE_M_01(w_Rs)

            With w_Rs
                If .RecordCount <= 0 Then
                    '�ް����Ȃ��Ƃ�
                    ReDim w_strMsg(1)
                    w_strMsg(1) = "�p�b�P�[�W�}�X�^"
                    Call General.paMsgDsp("NS0032", w_strMsg)
                    .Close()
                    Exit Function
                Else
                    .MoveLast()
                    w_lngRecCnt = .RecordCount
                    .MoveFirst()

                    w_CD_F = .Fields("PACKAGECD")
                    w_FLG_F = .Fields("USEFLG")

                    For w_lngLoop = 1 To w_lngRecCnt
                        Select Case w_CD_F.Value
                            Case "A"
                                w_strAppliUseFlg = w_FLG_F.Value & ""
                            Case "D"
                                w_strDutyUseFlg = w_FLG_F.Value & ""
                        End Select

                        .MoveNext()
                    Next w_lngLoop
                End If

                .Close()
            End With

            w_Rs = Nothing

            '����
            If w_strAppliUseFlg = "0" Then
            Else
                '���ڐݒ�
                w_strAppliUseFlg = General.paGetItemValue(General.G_STRMAINKEY1, General.G_STRSUBKEY1, "USEAPPLIFLG", "0", General.g_strHospitalCD)
            End If

            '�p�b�P�[�W�}�X�^(0:�͏o�~�������~,1:�͏o�~��������,2:�͏o���������~,3:�͏o����������)
            m_PackageFLG = 0

            If w_strAppliUseFlg = "1" Then
                '�͏o����
                m_PackageFLG = m_PackageFLG + 2
            End If

            If w_strDutyUseFlg = "1" Then
                '����������
                m_PackageFLG = m_PackageFLG + 1
            End If

            '����I��
            Get_PackageUseFLG = True

        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    Private Sub HscKinmu_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscKinmu.Scroll
        HscKinmu_Change(eventArgs.NewValue)
    End Sub

    Private Sub HscSet_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscSet.Scroll
        HscSet_Change(eventArgs.NewValue)
    End Sub

    Private Sub HscTokusyu_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscTokusyu.Scroll
        HscTokusyu_Change(eventArgs.NewValue)
    End Sub

    Private Sub HscYasumi_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscYasumi.Scroll
        HscYasumi_Change(eventArgs.NewValue)
    End Sub

    '�R���g���[���z��̑���Ƀ��X�g�Ɋi�[����
    Private Sub subSetCtlList()

        General.paSetControlList(_Frame_0, "_cmdKinmu_", m_lstCmdKinmu)
        General.paSetControlList(_Frame_1, "_cmdYasumi_", m_lstCmdYasumi)
        General.paSetControlList(_Frame_2, "_CmdTokusyu_", m_lstCmdTokusyu)
        General.paSetControlList(_Frame_3, "_CmdSet_", m_lstCmdSet)

        '�Ζ�
        For Each w_control As Button In m_lstCmdKinmu
            AddHandler w_control.Click, AddressOf m_lstCmdKinmu_Click
        Next
        '�x��
        For Each w_control As Button In m_lstCmdYasumi
            AddHandler w_control.Click, AddressOf m_lstCmdYasumi_Click
        Next
        '����
        For Each w_control As Button In m_lstCmdTokusyu
            AddHandler w_control.Click, AddressOf m_lstCmdTokusyu_Click
        Next
        '�Z�b�g
        For Each w_control As Button In m_lstCmdSet
            AddHandler w_control.Click, AddressOf m_lstCmdSet_Click
        Next
    End Sub


    ''' <summary>
    ''' �͏o�f�[�^���݃`�F�b�N
    ''' </summary>
    ''' <param name="p_YYYYMMDD"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncChkAppliData(ByVal p_YYYYMMDD As Integer) As Boolean

        Dim w_strMsg As String()
        Try
            If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then

                With General.g_objSyouninData
                    .pHospitalCD = General.g_strHospitalCD
                    .mKC_UNIQUESEQNO = ""
                    .mKC_StaffMngID = M_StaffID
                    .mKC_TargetDate = p_YYYYMMDD
                    .pSecurityObj = General.g_objSecurity
                    If .mChkKinmuChange = False Then
                        ReDim w_strMsg(1)
                        w_strMsg(1) = "���ԊO�����ɓo�^����Ă��邽��~n"
                        Call General.paMsgDsp("NS0110", w_strMsg)
                        Return False
                    End If
                End With
            End If
            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' �X�v���b�h�f�t�H���g�L�[������
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subChgSpreadKeyMap()
        Dim w_PreErrorProc As String = General.g_ErrorProc
        General.g_ErrorProc = "NSC0000HA subChgSpreadKeyMap"

        Dim im_m As New FarPoint.Win.Spread.InputMap

        Try
            '�f�t�H���g�Őݒ肳��Ă���F2�AF3�AF4���Ƃ肠����������
            '�uF2�v�F�ҏW���[�h���L���ɂȂ��Ă���ꍇ�́A�A�N�e�B�u�Z�����̒l������
            im_m = sprSheet.GetInputMap(FarPoint.Win.Spread.InputMapMode.WhenFocused)
            im_m.Put(New FarPoint.Win.Spread.Keystroke(Keys.F2, Keys.None), FarPoint.Win.Spread.SpreadActions.None)

            im_m = sprSheet.GetInputMap(FarPoint.Win.Spread.InputMapMode.WhenAncestorOfFocused)
            im_m.Put(New FarPoint.Win.Spread.Keystroke(Keys.F2, Keys.None), FarPoint.Win.Spread.SpreadActions.None)

            '�uF3�v�F�ҏW���[�h���L���ɂȂ��Ă���ꍇ�́A���t�����^�Z���Ɍ��݂̓��������
            im_m = sprSheet.GetInputMap(FarPoint.Win.Spread.InputMapMode.WhenFocused)
            im_m.Put(New FarPoint.Win.Spread.Keystroke(Keys.F3, Keys.None), FarPoint.Win.Spread.SpreadActions.None)

            im_m = sprSheet.GetInputMap(FarPoint.Win.Spread.InputMapMode.WhenAncestorOfFocused)
            im_m.Put(New FarPoint.Win.Spread.Keystroke(Keys.F3, Keys.None), FarPoint.Win.Spread.SpreadActions.None)

            '�uF4�v�F���t�����^�Z���ŕҏW���[�h���L���ɂȂ��Ă���ꍇ�́A���t��I�����邽�߂̃|�b�v�A�b�v�J�����_�[���X�v���b�h�V�[�g�ɕ\��
            im_m = sprSheet.GetInputMap(FarPoint.Win.Spread.InputMapMode.WhenFocused)
            im_m.Put(New FarPoint.Win.Spread.Keystroke(Keys.F4, Keys.None), FarPoint.Win.Spread.SpreadActions.None)

            im_m = sprSheet.GetInputMap(FarPoint.Win.Spread.InputMapMode.WhenAncestorOfFocused)
            im_m.Put(New FarPoint.Win.Spread.Keystroke(Keys.F4, Keys.None), FarPoint.Win.Spread.SpreadActions.None)

            General.g_ErrorProc = w_PreErrorProc
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' �L�[�_�E������
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub sprSheet_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles sprSheet.KeyDown
        Const W_SUBNAME As String = "NSK0000HC  sprSheet_KeyDown"

        Try
            If Not e.Control Then
                '�R���g���[���L�[���������̓X���[
                If IsNumOrFuncKey(e.KeyCode) Then
                    '�ݒ肪�Ȃ���΃X���[
                    If Not g_objKeyBoard.ContainsKey(e.KeyCode) Then Exit Sub
                    '�Ζ��\��t��
                    pasteKeyBoardKinmu(g_objKeyBoard(e.KeyCode))
                End If

                If e.KeyCode = Keys.Delete Then
                    cmdErase.PerformClick()
                End If
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    ''' <summary>
    ''' �Ζ��\��t���i�w��Ζ��R�[�h�j
    ''' </summary>
    ''' <param name="p_kinmuCd"></param>
    ''' <remarks>�L�[�{�[�h�Ή��Ŏg�p</remarks>
    Private Sub pasteKeyBoardKinmu(ByVal p_kinmuCd As String)
        Const W_SUBNAME As String = "NSK0000HC  sprSheet_KeyDown"

        Dim w_RegStr As String
        Dim w_Var As Object
        Dim w_ActiveCol As Integer
        Dim w_ActiveRow As Integer
        Dim w_RiyuKBN As String '���R�敪
        Dim w_Time As String '���ԔN�x
        Dim w_Flg As String '�m���׸�
        Dim w_ForeColor As Integer '�����F
        Dim w_BackColor As Integer '�w�i�F
        Dim w_InputFlg As Boolean '�����׸�
        Dim w_Cnt As Short
        Dim w_KinmuPlanCD As String 'KinmuCD(�\���ް�)
        Dim w_STS As Short
        Dim w_KangoCD As String
        Dim w_KangoPlanCD As String
        Dim w_RiyuPlanKbn As String
        Dim w_YYYYMMDD As Integer
        Dim w_DaikyuInputFlg As Boolean
        Dim w_strMsg() As String
        Dim w_lngBackColor As Integer
        Dim w_IntCol As Short
        Dim w_KibouCnt As Short
        Dim w_KibouCol() As Integer
        Dim w_blnColChk As Boolean
        Dim w_Comment As String = String.Empty  '��]�Ζ����̃R�����g 2015/04/13 Bando Add

        Try
            '̫����ړ�
            sprSheet.Focus()

            'ڼ޽�؊i�[��
            w_RegStr = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            '�P���ł����݂���΁E�E�E
            If m_TokusyuCnt <> 0 Then

                w_ActiveCol = sprSheet.Sheets(0).ActiveColumn.Index
                w_ActiveRow = sprSheet.Sheets(0).ActiveRow.Index

                '���͏ꏊ����
                If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row_ChgJisseki Then
                        Exit Sub
                    End If
                Else
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row Then
                        Exit Sub
                    End If
                End If

                '�����ٓ������i�z���͈́j
                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_ActiveCol).BackColor) = m_MonthBefore_Back Then
                    '�z�����ԊO(�w�i�F���O���[)�̏ꍇ���͕s��
                    Exit Sub
                End If

                '�����׸�
                w_InputFlg = True
                w_Cnt = CShort(w_ActiveCol - M_KinmuData_Col + 1)
                If g_SaikeiFlg = True Then
                    '�Čf�����̏ꍇ
                    If m_DataFlg(w_Cnt) = "1" Then
                        '�����ް��̏ꍇ
                        ReDim w_strMsg(2)
                        w_strMsg(1) = "�m��ς݋Ζ�"
                        w_strMsg(2) = "�Čf�Ζ�"
                        Call General.paMsgDsp("NS0011", w_strMsg)
                        w_InputFlg = False
                    End If
                End If


                '�Ζ��̗j�������`�F�b�N
                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                If Check_YoubiLimit(w_YYYYMMDD, p_kinmuCd) = False Then
                    Exit Sub
                End If

                '2013/01/07 Ishiga add start------------------------------------------
                '���΃f�[�^�̗L���`�F�b�N
                If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If Check_OverKinmuData(w_YYYYMMDD) = False Then
                        Exit Sub
                    End If
                End If
                '2013/01/07 Ishiga add end--------------------------------------------

                '�͏o���݃`�F�b�N
                If fncChkAppliData(w_YYYYMMDD) = False Then
                    Exit Sub
                End If

                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                If m_PackageFLG = 1 Or m_PackageFLG = 3 Then
                    '�������̃R���|�[�l���g����̎�
                    With General.g_objGetData
                        .p�E���敪 = 0
                        .p�E���ԍ� = M_StaffID
                        .p�`�F�b�N��� = w_YYYYMMDD
                        .p�����敪 = 0
                        .p�`�F�b�N�Ζ�CD = p_kinmuCd

                        If .mChkKinmuDuty = False Then
                            '�Ζ��ύX�s��
                            '*******ү����***********************************
                            ReDim w_strMsg(1)
                            w_strMsg(1) = ""
                            Call General.paMsgDsp("NS0110", w_strMsg)
                            '************************************************
                            Exit Sub
                        End If
                    End With
                End If

                w_DaikyuInputFlg = True

                If General.g_lngDaikyuMng = 0 Then
                    '��x����
                    w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                    w_YYYYMMDD = w_Var
                    If Check_Daikyu(M_PASTE, w_YYYYMMDD, p_kinmuCd) = False Then
                        w_DaikyuInputFlg = False
                        w_InputFlg = False
                    End If
                End If

                If w_InputFlg = True Then
                    '���͉\�ȏꍇ

                    With sprSheet.Sheets(0)

                        '�x���\��+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        '�ύX����Ζ��̒l����Ōv��ύX�̏ꍇ�A���̋Ζ����擾����B
                        w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                        If w_Var = "" And m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            w_Var = .GetText(w_ActiveRow - 1, w_ActiveCol)
                        End If

                        '2015/04/13 Bando Upd Start =========================================
                        'Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, "")
                        '2015/04/13 Bando Upd End   =========================================
                        If g_SaikeiFlg = False Then
                            Select Case w_RiyuKBN
                                Case "2"
                                    If frmNSK0000HA._mnuTool_5.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "�v���Ζ�"
                                        w_strMsg(2) = "�ҏW"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "3"
                                    If frmNSK0000HA._mnuTool_4.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "��]�Ζ�"
                                        w_strMsg(2) = "�ҏW"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "4"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�Čf�Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "5"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�ψ���Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "6"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�����Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                            End Select
                        Else
                            If w_RiyuKBN = "2" Then
                                If frmNSK0000HA._mnuTool_5.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "�v���Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If

                            If w_RiyuKBN = "3" Then
                                If frmNSK0000HA._mnuTool_4.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "��]�Ζ�"
                                    w_strMsg(2) = "�ҏW"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                        w_Time = ""
                        w_Flg = "0"
                        w_KangoCD = ""

                        '��]�񐔏W�v�`�F�b�N
                        w_lngBackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                        If CDbl(g_HopeNumFlg) = 1 And _OptRiyu_2.Checked = True Then
                            w_KibouCnt = 0
                            ReDim w_KibouCol(0)
                            For w_IntCol = m_KeikakuD_StartCol To m_KeikakuD_EndCol
                                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_IntCol).BackColor) = w_lngBackColor Then
                                    w_KibouCnt = w_KibouCnt + 1
                                    ReDim Preserve w_KibouCol(w_KibouCnt)
                                    w_KibouCol(w_KibouCnt) = w_IntCol
                                End If
                            Next w_IntCol

                            '�����ꏊ�ɓ����Ζ���\��t�����ꍇ�A�X���[
                            w_blnColChk = False
                            For w_IntCol = 1 To UBound(w_KibouCol)
                                If w_ActiveCol = w_KibouCol(w_IntCol) And p_kinmuCd = w_KinmuPlanCD Then
                                    w_blnColChk = True
                                End If
                            Next w_IntCol

                            If g_HopeNum <= w_KibouCnt And w_blnColChk = False Then
                                '��]�񐔐����I�[�o�[�_�C�A���O�\��
                                If g_KibouNumDiaLogFlg = 1 Then
                                    '���[�j���O
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "��]�Ζ�"
                                    w_strMsg(2) = "�ݒ肳�ꂽ��]�Ζ���"
                                    '�u&1��&2�𒴂��Ă��܂��B~n���̂܂ܓo�^���Ă���낵���ł����B�v
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    '�G���[
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "�ݒ肳�ꂽ��]�Ζ��񐔂𒴂��Ă��邽��"
                                    '�u&1���͂ł��܂���B�v
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If

                        Select Case True
                            Case _OptRiyu_0.Checked '�ʏ�
                                '���R�敪 �ʏ�
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                            Case _OptRiyu_1.Checked '�v��
                                '���R�敪 �v��
                                w_RiyuKBN = "2"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", General.G_PALEGREEN))
                            Case _OptRiyu_2.Checked '��]
                                '���R�敪 ��]
                                w_RiyuKBN = "3"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                            Case _OptRiyu_3.Checked
                                '���R�敪 �Čf
                                w_RiyuKBN = "4"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", General.G_LIGHTCYAN))
                            Case _OptRiyu_4.Checked
                                w_RiyuKBN = "6"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", General.G_ORANGE))
                            Case Else
                                '���R�敪 ���̑��i�ʏ툵���Ƃ���j
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                        End Select

                        '�敪�������̏ꍇ�̂݁A������Ζ��n�I����ʂ�\��
                        If w_RiyuKBN = "6" Then
                            If Disp_Ouen(w_KangoCD) = False Then
                                Exit Sub
                            End If
                        End If

                        '2015/04/13 Bando Add Start =============================
                        '��]or�v���̏ꍇ�R�����g���͉�ʕ\��
                        If w_RiyuKBN = "3" And g_InputHopeCommentFlg = "1" Then
                            If Disp_Comment(w_Comment, w_RiyuKBN) = False Then
                                Exit Sub
                            End If
                        End If
                        '2015/04/13 Bando Add End   =============================

                        '2015/07/22 Bando Add Start ========================
                        If w_Comment <> "" Then
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Fore", Convert.ToString(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Back", ColorTranslator.ToOle(Color.IndianRed)))
                        End If
                        '2015/07/22 Bando Add End   ========================

                        '��x�����Ζ��ޯ��װ
                        If General.g_lngDaikyuMng = 0 Then
                            If m_DaikyuBackColorFlg = True Then
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Back", General.G_LIGHTYELLOW))
                            End If
                        End If

                        '�قɋL����ݒ肷��
                        '2015/04/13 Bando Upd Start =========================
                        'w_Var = Set_KinmuMark(p_kinmuCd, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        w_Var = Set_KinmuMark(p_kinmuCd, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, w_Comment)
                        '2015/04/13 Bando Upd End =========================
                        .SetText(w_ActiveRow, w_ActiveCol, w_Var)
                        '���R�ʂ̐F�ݒ�
                        .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                        .Cells(w_ActiveRow, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        '�Ζ��ύX�̎�
                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            '2015/04/13 Bando Upd Start =========================================
                            'Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time)
                            Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time, "")
                            '2015/04/13 Bando Upd End   =========================================
                            '�\��ƈقȂ�ꍇ�F�ύX
                            If Trim(p_kinmuCd) <> Trim(w_KinmuPlanCD) Or w_RiyuKBN <> w_RiyuPlanKbn Or w_KangoCD <> w_KangoPlanCD Then
                                .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(m_Comp_Fore)
                            End If

                            '�ύX�\�s�̓��e���i�[
                            w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                            w_ForeColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).ForeColor)
                            w_BackColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor)
                            '�ύX�\�s�̓��e���Ζ��ύX��ʂɓ\��t���s�փR�s�[
                            .SetText(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol, w_Var)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        End If
                    End With
                End If
            End If

            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
            If InStr(m_strUpdKojyoDate, w_Var) = 0 Then
                m_strUpdKojyoDate = m_strUpdKojyoDate & w_Var & ","
            End If

            '���پوړ�
            If w_DaikyuInputFlg = True Then
                Call SetCursol()
            End If

            '�X�V�׸޾��
            If w_InputFlg = True Then
                m_KosinFlg = True
            End If
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    ''' <summary>
    ''' �w�i�F�`�F�b�N
    ''' </summary>
    ''' <param name="p_spr"></param>
    ''' <param name="p_row"></param>
    ''' <param name="p_col"></param>
    ''' <returns></returns>
    ''' <remarks>�ʏ�J���[���ǂ�������</remarks>
    Private Function IsExistBackColor(ByVal p_spr As FarPoint.Win.Spread.FpSpread, ByVal p_row As Integer, ByVal p_col As Integer) As Boolean
        Const W_SUBNAME As String = "NSK0000HC  IsExistBackColor"

        Dim rtnFlg As Boolean = True
        Dim frCl_Normal As Integer
        Dim bkCl_Normal As Integer

        Try
            '�ʏ�J���[
            frCl_Normal = ColorTranslator.ToOle(Color.Black)
            bkCl_Normal = ColorTranslator.ToOle(Color.White)

            With p_spr.Sheets(0)
                If ColorTranslator.ToOle(.Cells(p_row, p_col).BackColor) = bkCl_Normal _
                        OrElse .Cells(p_row, p_col).BackColor.A = 0 _
                        OrElse ColorTranslator.ToOle(.Cells(p_row, p_col).BackColor) = m_WeekEnd_Back Then
                    rtnFlg = False
                End If
            End With

            Return rtnFlg
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    ''' <summary>
    ''' ��ΐ�]�E�Z���ԏ��\��
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DispNightShortInfo()
        Const W_SUBNAME As String = "NSK0000HC  DispNightShortInfo"

        Dim cnt As Integer
        Dim bigDate As Integer
        Dim smallDate As Integer

        Try
            '�������i��\���j
            lblText1.Visible = False
            lblTermS1.Visible = False
            lblTermS2.Visible = False
            lblText2.Visible = False
            lblTermN1.Visible = False
            lblTermN2.Visible = False
            Panel1.Visible = False
            Panel2.Visible = False

            '�Z���ԃ`�F�b�N
            cnt = 0
            For i As Integer = 1 To UBound(m_shortWorkInfo)
                '���ԓ�������
                If m_shortWorkInfo(i).Date_St <= m_EndDate AndAlso m_StartDate <= m_shortWorkInfo(i).Date_Ed Then
                    cnt += 1
                    '�J�n��
                    smallDate = Integer.Parse(General.paGetDateStringFromInteger(m_shortWorkInfo(i).Date_St, General.G_DATE_ENUM.dd))

                    '�I����
                    bigDate = Integer.Parse(General.paGetDateStringFromInteger(m_shortWorkInfo(i).Date_Ed, General.G_DATE_ENUM.dd))

                    If cnt = 1 Then
                        lblTermS1.Text = General.paFormatSpace(smallDate, 2) & "���`" & General.paFormatSpace(bigDate, 2) & "��"
                        lblTermS1.Visible = True
                    Else
                        lblTermS2.Text = General.paFormatSpace(smallDate, 2) & "���`" & General.paFormatSpace(bigDate, 2) & "��"
                        lblTermS2.Visible = True
                    End If

                    '2���ȏ�͏�������
                    If cnt >= 2 Then Exit For
                End If
            Next
            '1���ȏ゠���
            If cnt >= 1 Then Panel1.Visible = True

            '��ΐ�]
            cnt = 0
            For i As Integer = 1 To UBound(m_nightWorkInfo)
                '���ԓ�������
                If m_nightWorkInfo(i).Date_St <= m_EndDate AndAlso m_StartDate <= m_nightWorkInfo(i).Date_Ed Then
                    cnt += 1
                    '�J�n��
                    smallDate = Integer.Parse(General.paGetDateStringFromInteger(m_nightWorkInfo(i).Date_St, General.G_DATE_ENUM.dd))
                    '�I����
                    bigDate = Integer.Parse(General.paGetDateStringFromInteger(m_nightWorkInfo(i).Date_Ed, General.G_DATE_ENUM.dd))

                    If cnt = 1 Then
                        lblTermN1.Text = General.paFormatSpace(smallDate, 2) & "���`" & General.paFormatSpace(bigDate, 2) & "��"
                        lblTermN1.Visible = True
                    Else
                        lblTermN2.Text = General.paFormatSpace(smallDate, 2) & "���`" & General.paFormatSpace(bigDate, 2) & "��"
                        lblTermN2.Visible = True
                    End If

                    '2���ȏ�͏�������
                    If cnt >= 2 Then Exit For
                End If
            Next
            '1���ȏ゠���
            If cnt >= 1 Then lblText2.Visible = True

            If Not Panel1.Visible AndAlso Panel2.Visible Then
                '�Z���Ԕ�\���Ŗ�ΐ�]������΋l�߂�
                Panel2.Top = Panel2.Top - Panel1.Height
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub sprSheet_Paint(ByVal sender As Object, ByVal e As PaintEventArgs) Handles sprSheet.Paint
        Dim colLoop As Integer
        Dim weekDayStr As String

        Try
            With sprSheet.Sheets(0)
                For colLoop = M_KinmuData_Col To m_KeikakuD_EndCol
                    weekDayStr = .GetText(1, colLoop)

                    If (m_WeekEndColorFlg = "1" AndAlso (weekDayStr = "�y" OrElse weekDayStr = "��")) OrElse _
                            (m_HolidayColorFlg = "1" AndAlso (weekDayStr = "�j" OrElse weekDayStr = "�x")) Then

                        If .Cells(M_KinmuData_Row, colLoop).BackColor.A = 0 Then
                            .Cells(M_KinmuData_Row, colLoop).BackColor = ColorTranslator.FromOle(m_WeekEnd_Back)
                        ElseIf .Cells(M_KinmuData_Row, colLoop).BackColor = Color.White Then
                            .Cells(M_KinmuData_Row, colLoop).BackColor = ColorTranslator.FromOle(m_WeekEnd_Back)
                        End If

                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            If .Cells(M_KinmuData_Row_ChgJisseki, colLoop).BackColor.A = 0 Then
                                .Cells(M_KinmuData_Row_ChgJisseki, colLoop).BackColor = ColorTranslator.FromOle(m_WeekEnd_Back)
                            ElseIf .Cells(M_KinmuData_Row_ChgJisseki, colLoop).BackColor = Color.White Then
                                .Cells(M_KinmuData_Row_ChgJisseki, colLoop).BackColor = ColorTranslator.FromOle(m_WeekEnd_Back)
                            End If
                        End If

                    Else

                        If .Cells(M_KinmuData_Row, colLoop).BackColor = ColorTranslator.FromOle(m_WeekEnd_Back) Then
                            .Cells(M_KinmuData_Row, colLoop).BackColor = Color.White
                        End If

                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            If .Cells(M_KinmuData_Row_ChgJisseki, colLoop).BackColor = ColorTranslator.FromOle(m_WeekEnd_Back) Then
                                .Cells(M_KinmuData_Row_ChgJisseki, colLoop).BackColor = Color.White
                            End If
                        End If

                    End If
                Next
            End With
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), "")
            End
        End Try
    End Sub

    '2014/04/23 Saijo add start P-06979--------------------------------------------------------------------------------------------------
    '/----------------------------------------------------------------------/
    '/  �T�v�@�@�@�@  : �Ζ��L���S�p�Q�����Ή��̃��C�A�E�g�ύX
    '/  �p�����[�^    : �Ȃ�
    '/  �߂�l        : �Ȃ�
    '/----------------------------------------------------------------------/
    Private Sub SetKinmuSecondView()
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc '��ۼ��ެ���̈ꎞ�Ҕ�
        General.g_ErrorProc = "NSK0000HC SetKinmuSecondView"

        Const W_FRAME_FIRST_HEIGHT As Integer = 20 '1�s�ڂ̏c�ʒu
        Const W_FRAME_ADD_HEIGHT As Integer = 27 '�s�̏c�ʒu������

        Const W_FRAME_FIRST_WIDTH As Integer = 8 '1��ڂ̉��ʒu
        Const W_FRAME_ADD_WIDTH As Integer = 39 '��̉��ʒu������

        Const W_FRAME_HEIGHT As Integer = 95 '�t���[���̏c��
        Const W_FRAME_WIDTH As Integer = 990 '�t���[���̉���
        Const W_SCL_WIDTH As Integer = 976 '�X�N���[���̉���
        Const W_SCL_HEIGHT As Integer = 16 '�X�N���[���̏c��
        Const W_KINMU_WIDTH As Integer = 40 '�Ζ��̉���
        Const W_KINMU_HEIGHT As Integer = 25 '�Ζ��̏c��

        Try
            '�Ζ��L���S�p�Q�����Ή��t���O����
            If m_strKinmuEmSecondFlg = "0" Then
                '0�F�Ή����Ȃ�(�]���̋Ζ��L�����̓T�C�Y�ƍő�2�o�C�g)
            Else
                '1�F�Ή�����(�S�p�Q�������\���ł���Ζ��L�����̓T�C�Y�ƍő�4�o�C�g)
                '�p���b�g���ڂ��Ă���t���[��
                FramAll.Size = New System.Drawing.Size(1200, 400)

                '�t���[��
                _Frame_0.Size = New System.Drawing.Size(W_FRAME_WIDTH, W_FRAME_HEIGHT)
                _Frame_1.Size = New System.Drawing.Size(W_FRAME_WIDTH, W_FRAME_HEIGHT)
                _Frame_2.Size = New System.Drawing.Size(W_FRAME_WIDTH, W_FRAME_HEIGHT)
                _Frame_3.Size = New System.Drawing.Size(W_FRAME_WIDTH, W_FRAME_HEIGHT)

                '�X�N���[��
                HscKinmu.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)
                HscYasumi.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)
                HscTokusyu.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)
                HscSet.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)

                '�Ζ�
                General.setSizeAndLocal(m_lstCmdKinmu, 2, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

                '�Ζ�(�x��)
                General.setSizeAndLocal(m_lstCmdYasumi, 2, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

                '�Ζ�(����Ζ�)
                General.setSizeAndLocal(m_lstCmdTokusyu, 2, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

                '�Ζ�(�Z�b�g)
                General.setSizeAndLocal(m_lstCmdSet, 1, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

            End If

            General.g_ErrorProc = w_PreErrorProc '�Ҕ���ۼ��ެ�������ɖ߂�

        Catch ex As Exception
            Err.Raise(Err.Number)
        End Try
    End Sub
    '2014/04/23 Saijo add end P-06979----------------------------------------------------------------------------------------------------
End Class