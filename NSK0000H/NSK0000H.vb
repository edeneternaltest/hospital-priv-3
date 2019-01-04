Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Module BasNSK0000H
    '/----------------------------------------------------------------------/
    '/
    '/    ���і��́F�Ō�x���V�X�e��(�Ζ��Ǘ�)
    '/ ��۸��і��́F�v��쐬���j���[
    '/        �h�c�FNSK0000H
    '/        �T�v�F�Ώی��̌v��𗧈Ă���
    '/
    '/
    '/      �쐬�ҁF S.Y    CREATE 2000/07/24           REV 01.00
    '/      �X�V�ҁF M.N           2008/11/25           �yP-00859�z
    '/      �X�V�ҁF M.I           2008/12/08           �yP-00931�z
    '/      �X�V�ҁF M.I           2008/12/09           �yP-00947�z
    '/      �X�V�ҁF M.I           2008/12/09           �yP-00958�z
    '/      �X�V�ҁF M.I           2008/12/09           �yPRE-0314�z
    '/      �X�V�ҁF M.I           2008/12/16           �yP-01000�z
    '/      �X�V�ҁF M.I           2008/12/17           �yP-01006�z
    '/      �X�V�ҁF M.I           2008/12/24           �yP-01053�z
    '/      �X�V�ҁF M.I           2008/12/25           �yP-01082�z
    '/      �X�V�ҁF M.I           2009/01/08           �yP-01127�z
    '/      �X�V�ҁF M.I           2009/01/08           �yP-01132�z
    '/      �X�V�ҁF M.I           2009/01/15           �yP-01172�z
    '/      �X�V�ҁF M.I           2009/02/09           �yP-01424�z
    '/      �X�V�ҁF M.I           2009/06/10           �yPKG-0215�z
    '/      �X�V�ҁF M.I           2009/06/10           �yPKG-0129�z
    '/      �X�V�ҁF M.I           2009/06/15           �yPKG-0089�z
    '/      �X�V�ҁF M.I           2009/06/16           �yPRE-0683�z
    '/      �X�V�ҁF T.I           2009/06/18           �yPRE-0706�z
    '/      �X�V�ҁF M.I           2009/06/19           �yPRE-0709�z
    '/      �X�V�ҁF M.I           2009/06/19           �yPRE-0713�z
    '/      �X�V�ҁF M.I           2009/06/19           �yPRE-0721�z
    '/      �X�V�ҁF M.I           2009/06/19           �yPRE-0726�z
    '/      �X�V�ҁF M.I           2009/06/25           �yPRE-0772�z
    '/      �X�V�ҁF M.I           2009/07/07           �yP-01967�z
    '/      �X�V�ҁF M.I           2009/07/07           �yPRE-0906�z
    '/      �X�V�ҁF M.I           2009/07/13           �yPRE-0914�z
    '/      �X�V�ҁF M.I           2009/07/15           �yP-02030�z
    '/      �X�V�ҁF okamura       2009/07/17           �yP-01981�z
    '/      �X�V�ҁF okamoto       2009/07/23           �yP-02050�z
    '/      �X�V�ҁF okamura       2009/08/07           �yPRE-1013�z
    '/      �X�V�ҁF okamura       2009/09/02           �yP-02215�z
    '/      �X�V�ҁF M.I           2009/11/12           �yP-02390�z
    '/      �X�V�ҁF Y.I           2012/10/24           �yP-*****�z�iPKG�o�[�W����UP_7.0�j
    '/      �X�V�ҁF T.Ishiga      2013/01/07           �yP-05697�z
    '/      �X�V�ҁF Y.Bando       2015/04/10           �yP-07830�z (PKG7.5)��]�Ζ����R�����g���͋@�\�ǉ�
    '/      �X�V�ҁF Angelo        2017/08/24           �yPKG�o�[�W�����A�b�v�z
    '/     Copyright (C) Inter co.,ltd 2000
    '/----------------------------------------------------------------------/

    '--------------------------------------------------------------------------------
    '       NSK0000H �萔 �錾
    '--------------------------------------------------------------------------------
    '�A�C�R��
    Public Const G_FORM_ICO As String = "kinmu.ico" '�t�H�[���A�C�R��
	Public Const G_ERASER_ICO As String = "Eraser.ico" '�����S��(NSK0000HB)
	Public Const G_CLOSE_ICO As String = "Close.ico" '����(NSK0000HB)
	Public Const G_SEARCH_ICO As String = "Search.ico" '����(NSK0000HD)
	Public Const G_PERMUTATION_ICO As String = "Permutation.ico" '�u��(NSK0000HD)
	
	'������
	Public Const G_LOAD_STR As String = "�ꎞ�t�@�C���ǂݍ��ݒ�..." '(1002)
	Public Const G_SAVE_STR As String = "�ꎞ�t�@�C���ۑ���..." '(1005)
	Public Const G_KEIKAKUSAVE_STR As String = "�v��f�[�^ �ۑ���..." '(1010)
	Public Const G_PICKUP_STR As String = "�f�[�^���o���c" '(1011)
	Public Const G_SORT_STR As String = "�f�[�^���ёւ���..." '(1012)
	Public Const G_DELTE_STR As String = "������..." '(1014)
	Public Const G_CUT_STR As String = "�؂��蒆..." '(1015)
    Public Const G_ROLLBACK_STR As String = "���ɖ߂���..." '(1016)
    Public Const G_REDO_STR As String = "��蒼����..." '2016/04/05 Ishiga add
	Public Const G_PASTE_STR As String = "�Ζ��L���\��t����..." '(1017)
	Public Const G_COPY_STR As String = "�R�s�[��..." '(1018)
	Public Const G_JISSEKISAVE_STR As String = "���уf�[�^ �X�V��..." '(1020)
	Public Const G_SEARCH_STR As String = "������..." '(1023)
	Public Const G_TEAM_STR As String = "�`�[��" '(5008)
    Public Const G_KINMUDEPT_STR As String = "�Ζ�����" '(5009)

    '2017/08/24 Angelo add st---------------------------------
    '�\�����e�̕ҏW���[�h
    Public Const G_EDITMODE_NO As Short = 0
    Public Const G_EDITMODE_DATETIME As Short = 1
    '2017/08/24 Angelo add en---------------------------------

    '2017/09/08 Angelo add st----------------------------------------------
    '�ް�����
    Public Const M_PLANDATA As String = "0" '�v���ް�
    Public Const M_JISSEKIDATA As String = "1" '�����ް�
    '2017/09/08 Angelo add en----------------------------------------------

    '--------------------------------------------------------------------------------
    '       NSK0000H �ϐ� �錾
    '--------------------------------------------------------------------------------
    Public g_AppName As String '��۸���ID �i�[
    Public g_KinmuM() As KinmuM_Type '�Ζ����z��
    Public g_HolidayBunruiM() As HolidayM_Type '�x�ݕ��ޏ��z��
	Public w_KinmuCDCount As Short
	Public g_SaikeiFlg As Boolean 'True:�Čf����,False:�Čf�����ȊO
	Public g_LimitedFlg As Boolean '(True�F�Ζ���  False�F�Ǘ���)
	Public g_ImagePath As String '�C���[�W�p�X
	
    Public g_HopeNum As Short '��]��
    '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
    Public g_HopeNumDate As Short '��]��(���t��)
    '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------
    Public g_HopeNumFlg As String '��]�񐔐����t���O(1:��������@2:�������Ȃ�)
    '2014/05/22 Shimpo add start P-06991-----------------------------------------------------------------------
    Public g_HopeNumDateFlg As String '��]��(���t��)�����t���O(1:��������@2:�������Ȃ�)
    '2014/05/22 Shimpo add end P-06991-------------------------------------------------------------------------
    Public g_KibouNumDiaLogFlg As Integer '��]�񐔐����_�C�A���O�i1:���[�j���O�@�ȊO:�G���[�j
    Public g_HopeMode As String = "0"        '�i1:��]���[�h�A0:����ȊO�j
    Public g_objKeyBoard As Dictionary(Of Integer, String)
    '2015/04/13 Bando Add Start ===================
    Public g_InputHopeCommentFlg As String '��]�R�����g�t���O(1�F�R�����g�� 2:�R�����g�s��)
    '2015/04/13 Bando Add Start ===================

    '2017/09/04 Angelo Add st-----------------------------------------------------------------------------------------------------------------------
    '�Ζ��n�ٓ����
    Public Structure IdoData_Type
        Dim CD As String '�ǉ��d�l�F�����ҁE�����ٓ��҂̑���΂�S���o��
        Dim IdoYMD As Integer '�ٓ��N����
        Dim IdoYMD2 As Integer '�ٓ��N�����Q
        Dim SyuryoYMD As Integer '�I���N����
    End Structure

    '�Ζ����
    Public Structure KinmuData_Type
        Dim KinmuCD As String 'KinmuCD
        Dim Date_Renamed As Integer '���t
        Dim RiyuKBN As String '���R�敪
        Dim Time As String '���ԔN�x�i�ő�S���܂Łj
        Dim DataFlg As String '�v���ް�,�����ް����ʗp�׸�(0:�v���ް�,1:�����ް�)
        Dim KakuteiFlg As String '�m�蔻�ʗp�׸�(0:�Y������,1:������)
        Dim DataChk As String 'DB�̑��݂̗L��("1":DB�ް�����C"":DB�ް��Ȃ��j
        Dim OuenKangoCD As String '������Ō�P��CD
        Dim FirstRegistTimeDate As Double
        Dim LastUpdTimeDate As Double
        Dim RegistID As String
        Dim Comment As String '��]�R�����g�@2015/04/10 Bando Add
    End Structure

    '�N�x�ڍ׏��
    Public Structure NenkyuDetail_Type
        Dim GetContentsKbn As String '�擾���e�敪(1:�S��,2:�O��,3:�㔼,4:���ԔN�x)
        Dim HolidayBunruiCD As String '�x�ݕ���CD
        Dim FromTime As Integer '�J�n����
        Dim ToTime As Integer '�I������
        Dim DateKbn As String '�N�����敪(0:����,1:����)
        Dim NenkyuTime As Integer '���ԔN�x
        Dim HolSubFlg As String '�x�e���Z�t���O
        Dim DayTime As Integer '���Ύ���
        Dim NightTime As Integer '��Ύ���
        Dim NextNightTime As Integer '������Ύ���
    End Structure

    '�N�x���
    Public Structure NenkyuData_Type
        Dim Date_Renamed As Integer '���t
        '*=*=*=*=*=*=*=*=*=*=*=*=�r�������Ή��ׂ̈̈ꎞ�ۑ��X�y�[�X*=*=*=*=*=*=*=*=*=*=*=*=*
        '���̍\���͍̂X�V����(�uSave_PlanData�v�uSave_PlanData_4W�v�uSave_JissekiData�v)�̎�
        '�ɒl�����uUpDate_JikanNenkyu�v�łc�a��(NS_NENKYU_F)�ɕۑ����ɂ����ׂɎg�p���Ă܂�
        Dim Detail() As NenkyuDetail_Type
        '*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
        Dim DataFlg As Boolean '�f�[�^�X�VFLG(True:�X�V����CFalse:�X�V���Ȃ�)
        Dim DataChk As Boolean 'DB�̑��݂̗L��(True:DB�ް�����CFalse:DB�ް��Ȃ��j
        Dim FirstRegistTimeDate As Double
        Dim LastUpdTimeDate As Double
        Dim RegistID As String
    End Structure

    '�l�ʋΖ�����
    Public Structure PersonalCondition_Type
        Dim NotKinmu As Boolean 'True:�����s�CFalse:������
        Dim CountMax As Integer
        Dim CountMin As Integer
        Dim IntervalMax As Integer
        Dim IntervalMin As Integer
        Dim MondayNot As Boolean 'True:�����s�CFalse:������
        Dim TuesdayNot As Boolean 'True:�����s�CFalse:������
        Dim WednesdayNot As Boolean 'True:�����s�CFalse:������
        Dim ThursdayNot As Boolean 'True:�����s�CFalse:������
        Dim FridayNot As Boolean 'True:�����s�CFalse:������
        Dim SaturdayNot As Boolean 'True:�����s�CFalse:������
        Dim SundayNot As Boolean 'True:�����s�CFalse:������
        '2015/05/13 Ishiga add start---------------------
        Dim RenzokuCountMax As Integer '�A���Ζ������
        Dim RenkyuCountMax As Integer '�A�x�����
        '2015/05/13 Ishiga add end-----------------------
    End Structure

    '�Ō���Z�v�Z�p�f�[�^�\����
    Public Structure Kangokasan_Type
        Dim PostCD As String
        Dim JobCD As String
        Dim SaiyoCD As String
        Dim YakinKBN As String
        Dim ChildShort As Boolean '�Z����
    End Structure

    '��x�ڍחp�\����
    'Private Structure DaikyuDetail_Type
    <Serializable()> Public Structure DaikyuDetail_Type '2016/04/06 Yamanishi Upd
        Dim DaikyuDate As Integer '��x������
        Dim DaikyuKinmuCD As String '��x�����Ζ��b�c
        Dim GetFlg As String '�擾�^�C�v(0:1���A1:0.5��)
    End Structure

    '��x�p�\����
    'Private Structure Daikyu_Type
    <Serializable()> Public Structure Daikyu_Type '2016/04/06 Yamanishi Upd
        Dim HolDate As Integer
        Dim HolKinmuCD As String
        Dim DaikyuDetail() As DaikyuDetail_Type '��x�ڍ�
        Dim GetKbn As String '��x�����ʃ^�C�v(0:1����,1:1.5����)
        Dim RemainderHol As Double '�c���x
        Dim OutPutList As String '��x�擾���Ƀ��X�g�ɂ����邩�ǂ������׸�("0":��x�擾����Ώ� "1":��x�擾���Ώ�)
        Dim FirstRegistTimeDate As Double
        Dim LastUpdTimeDate As Double
        Dim RegistID As String
    End Structure

    '���ԔN�x�p�\����
    Public Structure JikanNenkyu_Type
        Dim Date_Renamed As Integer
        Dim StrData As String
    End Structure

    '�̗p�������
    Public Structure SaiyoData_Type
        Dim SaiyoDate As Integer
        Dim TentaiDate As Integer
        Dim SaiyoCD As String
        Dim strStaffNo As String
        Dim lngHaizoku As Integer
        Dim blnHaizokuFLG As Boolean
        Dim strSecName As String
    End Structure

    'Private Structure KojyoData_Type
    <Serializable()> Public Structure KojyoData_Type '2016/04/06 Yamanishi Upd
        Dim lngDate As Integer '�N����
        Dim lngKinmuDetailDate() As Integer '�Ζ��ڍהN����
        Dim strKinmuDetailCD() As String '�Ζ��ڍ�CD
        Dim lngTimeFrom() As Integer '���ԑ�From
        Dim lngTimeTo() As Integer '���ԑ�To
        Dim lngNikkinTime() As Integer '���Ύ���
        Dim lngYakinTime() As Integer '��Ύ���
        Dim lngYokuYakinTime() As Integer '������Ύ���
        Dim strNextFlg() As String '���TFLG
        Dim lngKinmuDetailTime() As Integer '�T������
        Dim strHolSubFlg() As String '�x�e���Z�t���O
        Dim strShinryoKbn() As String '�f�Õ�V�v�Z�敪
        Dim UniqueseqNo() As String '���j�[�NNO
        Dim Seq() As Integer
        '2018/02/23 Yamanishi Upd Start ------------------------------------------
        'Sub init()
        '    ReDim lngKinmuDetailDate(0)
        '    ReDim strKinmuDetailCD(0)
        '    ReDim lngTimeFrom(0)
        '    ReDim lngTimeTo(0)
        '    ReDim lngNikkinTime(0)
        '    ReDim lngYakinTime(0)
        '    ReDim lngYokuYakinTime(0)
        '    ReDim strNextFlg(0)
        '    ReDim lngKinmuDetailTime(0)
        '    ReDim strHolSubFlg(0)
        '    ReDim strShinryoKbn(0)
        '    ReDim UniqueseqNo(0)
        '    ReDim Seq(0)
        'End Sub

        Dim OuenKinmuDeptCD() As String

        Sub init(Optional ByVal p_Cnt As Integer = 0)
            ReDim lngKinmuDetailDate(p_Cnt)
            ReDim strKinmuDetailCD(p_Cnt)
            ReDim lngTimeFrom(p_Cnt)
            ReDim lngTimeTo(p_Cnt)
            ReDim lngNikkinTime(p_Cnt)
            ReDim lngYakinTime(p_Cnt)
            ReDim lngYokuYakinTime(p_Cnt)
            ReDim strNextFlg(p_Cnt)
            ReDim lngKinmuDetailTime(p_Cnt)
            ReDim strHolSubFlg(p_Cnt)
            ReDim strShinryoKbn(p_Cnt)
            ReDim UniqueseqNo(p_Cnt)
            ReDim Seq(p_Cnt)
            ReDim OuenKinmuDeptCD(p_Cnt)
        End Sub
        '2018/02/23 Yamanishi Upd End --------------------------------------------
    End Structure

    Public Structure SumCntDeteil_Type
        Dim GetFlg As Boolean
        Dim Cnt As Double
    End Structure

    Public Structure SumCntData_Type
        Dim SumSeq() As SumCntDeteil_Type
    End Structure

    '�ΏېE�����
    Public Structure StaffData_Type
        Dim ID As String '�E���Ǘ��ԍ�
        Dim PreID As String '�E���ԍ�
        Dim StaffName As String '����
        Dim SaiyoYMD As Integer '�̗p�N����
        Dim TentaiYMD As Integer '�]�ޔN����
        Dim IdoData() As IdoData_Type '�ٓ����i�Y���������̂݁j
        Dim InitialHyojiNo As Integer '�\��No(�����l)
        Dim HyojiNo As Integer '�\��No
        Dim HyojiNo1 As Integer '�\��No1
        Dim HyojiNo2 As Integer '�\��No2
        Dim HyojiNo3 As Integer '�\��No3
        Dim HyojiNo4 As Integer '�\��No4
        Dim HyojiNo5 As Integer '�\��No5
        Dim Team As Integer '���
        Dim AutoKBN As String '���������敪
        Dim JobHyojiNo As Integer '�E��\��No
        Dim PostHyojiNo As Integer '��E�\��No
        Dim GiryoHyojiNo As Integer '�Z�ʕ\��No
        Dim GiryoLvCD As String 'SkillLvlCD
        Dim GiryoBunruiCD As String 'SkillBunruiCD
        Dim KinmuData() As KinmuData_Type '�Ζ��ް�
        Dim SaikeiData() As KinmuData_Type '�Čf�ް��i�Čf�����̂݁j
        Dim CompKinmuData() As KinmuData_Type '�Ζ��ް�(�\���ް�[�����ް��Ƃ̔�r�p])
        Dim NenkyuData() As NenkyuData_Type '�N�x�ް�
        Dim KinmuCondition() As PersonalCondition_Type '�l�ʋΖ������f�[�^
        Dim PersonalError As Boolean 'True:�G���[�CFalse:�G���[�Ȃ�
        Dim JobCode As String '�E��CD     (I/F)
        Dim PostCode As String '��ECD     (I/F)
        Dim PostName As String '��E����     (I/F)�@
        Dim HaizokuIF As Integer '�z����     (I/F)
        Dim TensyutuIF As Integer '�]�o��     (I/F)
        Dim Syokai As Double 'RegistFirstTimeDate
        Dim Saisyu As Double 'LastUpdTimeDate
        Dim UpdateFlg As String '�E�����e�i��̫�Ēl:"0" Or �Y���v��ԍ�:"1" Or �O�v��ԍ�:"2"�j
        Dim YakinKBN As String '��ΐ�]�ҋ敪
        Dim PatternCD As String '�p�^�[���R�[�h
        Dim OuenStaffFlg As Integer '�����Ζ��҂��ǂ����̔��f�i0�F�Ώە��������@1�F�����Ζ��ҁj
        Dim TargetKikanFlg As Boolean '�Ώۊ��ԓ��i�P�����j�̊ԂɍݐE���Ă��邩�i�E���擾�͂P�����ōs���Ă��Ȃ����߁A�E�����ݒ��ʂɃf�[�^��n���ۂɎg�p�j (True:�ݐE, False:�ݐE���ĂȂ�)
        Dim KangoKasanData() As Kangokasan_Type '�Ō���Z�v�Z�p�f�[�^
        Dim Daikyu() As Daikyu_Type '��x�f�[�^
        Dim LoadDaikyu() As Daikyu_Type '���۰�ގ��̑�x�f�[�^�i�Ζ��ύX�l�ʉ�ʎ��Ɏg�p�j
        'Dim BackDaikyu() As Daikyu_Type '��x�f�[�^�ޔ�p '2016/04/06 Yamanishi Del
        Dim WariateJikanNenkyu() As JikanNenkyu_Type '�������蓖�Ď��̎��ԔN�x�ޔ�p�z��
        Dim WariateOuenKinmu() As JikanNenkyu_Type '�������蓖�Ď��̉����Ζ��ޔ�p�z��
        Dim WariateComment() As JikanNenkyu_Type '�������蓖�Ď��̊�]�R�����g�ޔ�p�z�� 2015/04/10 Bando Add
        Dim SaiyoData() As SaiyoData_Type '�̗p�������
        Dim RuikeiTime As Integer '�݌v����(�Y���v��ԍ��̂Q�O�܂�)
        Dim RuikeiTime_Jisseki As Integer '���ъ���(�Y���v��ԍ��̂P�O�̎���)
        Dim Kojyo() As KojyoData_Type
        Dim blnEndDayChangeFlg As Boolean '�����I�Ɉٓ����ENDDATE�����Ԃ̍ŏI���ɕϊ�������(99999999��0��20080331��)
        Dim NightWork() As NghtShrtData_Type '��ΐ�]���
        Dim ShortWork() As NghtShrtData_Type '�Z���Ԑ��x���
        Dim SumCntData As SumCntData_Type
        Dim SumCntData_4W() As SumCntData_Type
        Dim ResultCnt() As Integer
        '�ǉ��d�l�F�����ҁE�����ٓ��҂̑���΂�S���o��---------------------------------------------------------------------
        Dim BeforeKinmuData() As KinmuData_Type
        Dim IdoHistory() As IdoData_Type
        'Dim BeforeJikanNenkyu() As JikanNenkyu_Type '2016/06/15 Yamanishi Del
        '�ǉ��d�l�F�����ҁE�����ٓ��҂̑���΂�S���o��---------------------------------------------------------------------
    End Structure
    '2017/09/04 Angelo Add en-----------------------------------------------------------------------------------------------------------------------

    '-- �Ζ��L�� �ޔ�z�� ----------------------
    Public Structure KinmuM_Type
		Dim CD As String 'KinmuCD
		Dim KinmuName As String '����
		Dim Mark As String '�L��
		Dim KBunruiCD As String '�Ζ�����CD
		Dim WBunruiCD As String 'AllocBunruiCD
        Dim WFlg As String '�����׸�
		Dim HFlg As String '�����Ζ��׸�
		Dim AMCD As String '�`�l��CD
		Dim PMCD As String '�o�l��CD
		Dim From As Short '�Ζ����ԑ�FROM
        Dim To_Renamed As Short '�Ζ����ԑ�TO
		Dim Time As Short '�Ζ�����
		Dim TimeFlg As String '���ԋx�׸�
		Dim Setumei As String '����
        Dim DaikyuFlg As String '��x�擾�׸�(1:�\ 2:�s��)
        Dim HolBunruiCD As String '�x�ݕ���CD
        Dim KinmuKBN As String '�Ζ��敪(0:�Ζ� 1:�Ζ��ȊO)
        Dim YoubiLimit() As String '�j������
        Dim EfftoDate As String '�L���I���� 2015/05/22 Bando Add
	End Structure
	
    '-- �x�ݕ��ދL�� �ޔ�z�� --
	Public Structure HolidayM_Type
		Dim CD As String 'HolidayBunruiCD
		Dim HolidayName As String '����
		Dim SecName As String '����
		Dim Mark As String '�L��
		Dim Setumei As String '����
		Dim DivFlg As String '�����x�Ɏ擾�׸�(0:�s�� 1:�����\)
		Dim GetPosFrom As Short '�擾�\��FROM
		Dim GetPosTo As Short '�擾�\��TO
		Dim AppliCD As String '�͏oCD
	End Structure
	
	'�װ�����p
	'*** �񐔃`�F�b�N�p ***
    Public Structure SpanCount_Type
        Dim KinmuCount As Single
        Dim ErrorFlg As Boolean 'True:�װ����CFalse:�װ�Ȃ�
    End Structure

	'*** �Ԋu�`�F�b�N�p ***
    Public Structure IntervalErr_Type
        Dim ErrorDate As Integer
        Dim ErrorFlg As Boolean 'True:�װ����CFalse:�װ�Ȃ�
    End Structure

	'�񐔁^�Ԋu�G���[�`�F�b�N�p
    Public Structure CountErr_Type
        Dim ErrName As String
        Dim ErrBunrui As String
        Dim CheckSpan() As SpanCount_Type '�񐔃G���[�i�z��͌v����� 0:�\�����ԁCn:�e�v����Ԃ��Ɓj
        Dim InterValErr As IntervalErr_Type '�Ԋu�G���[
    End Structure

    Public g_KikanError() As CountErr_Type '�z��͏W�v�Ζ���
    Public g_RenzokuError() As CountErr_Type
	
	'*** �֎~�p�^�[���`�F�b�N�p ***
    Public Structure NotPatternErr_Type
        Dim ErrorPattern As String
        Dim ErrorDate As Integer
        Dim ErrorFlg As Boolean 'True:�װ����CFalse:�װ�Ȃ�
    End Structure

    Public g_NotPatternError As NotPatternErr_Type

	'*** �ے�Ζ��`�F�b�N�p ***
    Public Structure NotKinmuErr_Type
        Dim KinmuName As String
        Dim ErrorDate As Integer
        Dim ErrorFlg As Boolean 'True:�װ����CFalse:�װ�Ȃ�
    End Structure

    Public g_NotKinmuError As NotKinmuErr_Type

    '�v��P�ʗp�񋓌^�i�S�T�^�P�����j
	Public Enum gePlanType
		PlanType_Month '�P�����i�ް��F"1"�j
		PlanType_Week '�S�T�i�ް��F"2"�j
    End Enum

	'�\�����ԗp�񋓌^�i�S�T�^�P�����j
	Public Enum geViewType
		ViewType_Month '�P�����i�ް��F"1"�j
		ViewType_Week '�S�T�i�ް��F"2"�j
    End Enum

	'Window�ʒu��傫���ݒ�p�񋓌^
	Public Enum geWindowPosition
		GetSettingValue '����޳�̈ʒu�A�傫���ݒ�
		SaveSettingValue '����޳�̈ʒu�A�傫���ۑ�
    End Enum

    '����/�u�����
    Public Enum geKenChiFlg
        FormType_Kensaku '�������
        FormType_Chikan '�u�����
    End Enum

    '�װ�����p
    '*** �񐔃`�F�b�N�p ***
    Public Structure SpanCount2_Type
        Dim KinmuCount As Single
        Dim ErrorDate As Integer '�G���[�J�n��
        Dim ErrorFlg As Boolean 'True:�װ����CFalse:�װ�Ȃ�
        Dim ColIdx As Short '���t(��)�C���f�b�N�X
        Dim KinmuName As String '�G���[�Ζ�
    End Structure

	'*** �Ԋu�`�F�b�N�p ***
    Public Structure IntervalErr2_Type
        Dim ErrorDate As Integer
        Dim ErrorFlg As Boolean 'True:�װ����CFalse:�װ�Ȃ�
        Dim ColIdx As Short '���t(��)�C���f�b�N�X
        Dim ErrorName As String
    End Structure

	'�񐔁^�Ԋu�G���[�`�F�b�N�p
    Public Structure CountErr2_Type
        Dim ErrName As String
        Dim ErrBunrui As String
        Dim StaffName As String '�E������
        Dim StaffIdx As Short '�E��(�s)�C���f�b�N�X
        Dim CheckSpan() As SpanCount2_Type '�񐔃G���[�i�z��͌v����� 0:�\�����ԁCn:�e�v����Ԃ��Ɓj
        Dim InterValErr() As IntervalErr2_Type '�Ԋu�G���[
    End Structure

    Public g_KikanError2() As CountErr2_Type '�z��͏W�v�Ζ���
    Public g_RenzokuError2() As CountErr2_Type

    '��ΐ�]�E�Z���Ԏҏ��
    <Serializable()> Public Structure NghtShrtData_Type
        Dim Date_from As Integer
        Dim Date_to As Integer
        Dim ReasonCd As String
        Dim ReasonRNm As String
    End Structure

    '�s���Q����
    <Serializable()> Public Structure EventStaff_Type
        Dim staffMngId As String
        Dim staffNm As String
        Dim postNm As String
    End Structure

    '�s���\��
    <Serializable()> Public Structure EventList_Type
        Dim DateF As Integer
        Dim Time_st As Short
        Dim Time_ed As Short
        Dim EventName As String
        Dim allFlg As Boolean
        Dim uniqNo As String
        Dim EventStaff() As EventStaff_Type
        Sub init()
            DateF = 0
            Time_st = 0
            Time_ed = 0
            EventName = ""
            allFlg = False
            uniqNo = ""
            ReDim EventStaff(0)
        End Sub
    End Structure

    '*** �֎~�p�^�[���`�F�b�N�p ***
    Public Structure NotPatternDetail_Type
        Dim ErrorPattern As String
        Dim ErrorDate As Integer
        Dim ErrorFlg As Boolean 'True:�װ����CFalse:�װ�Ȃ�
        Dim ColIdx As Short '���t(��)�C���f�b�N�X
        Dim EndDate As Integer
    End Structure

    Public Structure NotPatternErr2_Type
        Dim StaffName As String
        Dim StaffIdx As Short '�E��(�s)�C���f�b�N�X
        Dim Data() As NotPatternDetail_Type
    End Structure

    Public g_NotPatternError2() As NotPatternErr2_Type

    '*** �ے�Ζ��`�F�b�N�p ***
    Public Structure NotKinmuDetail_Type
        Dim KinmuName As String
        Dim ErrorDate As Integer
        Dim ErrorFlg As Boolean 'True:�װ����CFalse:�װ�Ȃ�
        Dim ColIdx As Short '���t(��)�C���f�b�N�X
    End Structure

    Public Structure NotKinmuErr2_Type
        Dim StaffName As String
        Dim StaffIdx As Short '�E��(�s)�C���f�b�N�X
        Dim Data() As NotKinmuDetail_Type
    End Structure
    Public g_NotKinmuError2() As NotKinmuErr2_Type

    '*** �֎~�E���p�^�[���`�F�b�N�p ***
    Public g_NotStaffPatternError2() As NotPatternErr2_Type

    '�o���敪�̑g�ݍ��킹�`�F�b�N�p
    Public g_NotGiryoCheckError() As NotPatternErr2_Type

    Public g_NotAbsKinmuCheckError() As NotPatternErr2_Type '�K�{�Ζ�
    '2017/09/29 Yamanishi Add ---------------------------------------------------------------------------------------------------------
    Private m_dicRandomKey2TimeNenkyu As New Dictionary(Of String, String)
    ''' <summary>
    ''' �����_���ȕ�����𐶐����A�����Key�Ƃ���Dictionary��Value�ƂȂ�f�[�^���i�[
    ''' </summary>
    ''' <param name="p_Len">�������镶����̒���</param>
    ''' <param name="p_Val">Value�ƂȂ�f�[�^</param>
    ''' <param name="p_Dic">Dictionary</param>
    ''' <returns>�������ꂽ������</returns>
    Public Function GenerateRandomKeyAndSetDictionary(ByVal p_Len As Integer, ByVal p_Val As String, ByRef p_Dic As Dictionary(Of String, String)) As String
        '�g�p���镶��
        Const W_Chars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ@[]{};:*;-/.,?!#$%&()=^~|"
        Dim w_RetVal As New System.Text.StringBuilder(p_Len)
        Dim w_Random As New Random
        For i As Integer = 0 To p_Len - 1
            '�I�����ꂽ�ʒu�̕������擾�E�ǉ�
            w_RetVal.Append(W_Chars(w_Random.Next(W_Chars.Length)))
        Next i
        GenerateRandomKeyAndSetDictionary = w_RetVal.ToString
        If Not p_Dic.ContainsKey(GenerateRandomKeyAndSetDictionary) Then
            p_Dic.Add(GenerateRandomKeyAndSetDictionary, p_Val)
        Else
            '����Ă����蒼��
            Return GenerateRandomKeyAndSetDictionary(p_Len, p_Val, p_Dic)
        End If
    End Function
    '----------------------------------------------------------------------------------------------------------------------------------
    '*****************************************************************************************************
    '   ���گ�ނ̋Ζ��L����������e�f�[�^���ɕ�������
    '   ���Ұ��Fp_Val(I)�i�Ζ����ށj
    '           p_KinmuCD(O)�i�Ζ����ށj
    '           p_RiyuKBN(O)�i���R�敪�j
    '           p_Time(O)�i���ԔN�x�j
    '           p_Flg(0)�i�m�蕔���t���O�j
    '           p_KangoCD�i������Ō�P��CD�j
    '   �ҏW�d�l
    '       �Ζ��L��(2)�{Space(5)�{�Ζ�����(3)�{���R�敪(1)�{�m�蕔���t���O(1)+������Ō�P��CD(4) + ��]�R�����g(20) +���ԔN�x(44)
    '       �S�U�O�o�C�g��
    '           1) �Ζ��L��---�P�޲Ėڂ���V�޲ĕ��iTrim�ŗ]���Ƚ�߰����Ȃ��j
    '           2) KinmuCD---�W�޲Ėڂ���R�޲ĕ��iTrim�ŗ]���Ƚ�߰����Ȃ��j
    '           3) ���R�敪---11�޲Ėڂ���P�޲ĕ�
    '           4) �m�蕔��FLG---12�޲Ėڂ���P�o�C�g��
    '             �i"1"�F�������m���ް��C"0"("1"�ȊO):�Y�������m���ް��A���́A�\���ް��j
    '           5) ������Ō�P��CD---13�޲Ėڂ���S�o�C�g��     'Add Tanaka 2003/06/23
    '�@�@�@�@�@��2015/04/10 Bando Mod
    '           6) ��]�R�����g 16�޲Ėڂ���Q�O�o�C�g��
    '           7) ���ԔN�x---36�޲Ėڂ���S�S�޲ĕ�
    '*****************************************************************************************************
    '2015/04/10 Bando Upd Start ========================================
    'Public Sub Get_KinmuMark(ByVal p_Val As Object, ByRef p_KinmuCD As String, ByRef p_RiyuKBN As String, ByRef p_Flg As String, ByRef p_KangoCD As String, ByRef p_Time As String)
    Public Sub Get_KinmuMark(ByVal p_Val As Object, ByRef p_KinmuCD As String, ByRef p_RiyuKBN As String, ByRef p_Flg As String, ByRef p_KangoCD As String, ByRef p_Time As String, ByRef p_Comment As String)
        'On Error GoTo Get_KinmuMark
        Const W_SUBNAME As String = "BasNSK0000H Get_KinmuMark"

        Dim w_str As String

        Try
            '�Z���̒l���ݒ肳��Ă��Ȃ��ꍇ
            If Trim(p_Val) = "" Then
                'If IsDBNull(p_Val) = True Then
                p_KinmuCD = ""
                p_RiyuKBN = ""
                p_Flg = "0"
                p_KangoCD = ""
                p_Time = ""
                p_Comment = ""
                Exit Sub
            End If

            '�Ζ��L���ȊO�͔��p�Ƃ���
            w_str = CStr(p_Val)

            'w_str = Right(w_str, 127)
            w_str = General.paRightB(w_str, 147)

            '���گ�ޓ\��t��������𕪊�����
            p_KinmuCD = Trim(Left(w_str, 3)) '�Ζ��L���͏�����������3�o�C�g���m�ۂ���
            p_RiyuKBN = Trim(Mid(w_str, 4, 1))
            p_Flg = Trim(Mid(w_str, 5, 1))
            p_KangoCD = Trim(Mid(w_str, 6, 10))
            p_Comment = Trim(General.paMidB(w_str, 16, 20))
            'p_Time = Trim(Mid(w_str, 16, 112))
            p_Time = Trim(General.paMidB(w_str, 36, 112))
            '2017/09/29 Yamanishi Add ---------------------------------------------------------------------------------------------------------
            If m_dicRandomKey2TimeNenkyu.ContainsKey(p_Time) Then
                p_Time = Trim(m_dicRandomKey2TimeNenkyu(p_Time))
            End If
            '----------------------------------------------------------------------------------------------------------------------------------
            'Get_KinmuMark:
            '        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            '        End
        Catch ex As Exception
            End
        End Try
    End Sub
    '2015/04/10 Bando Upd End   ========================================

    '******************************************************************************************
    '   ���گ�ނɓ\��t����Ζ��L����ҏW����
    '   ���Ұ��Fp_KinmuCD�i�Ζ����ށj
    '           p_RiyuKBN�i���R�敪�j
    '           p_Time�i���ԔN�x�j
    '           p_Flg�i�m�蕔���t���O�j
    '           p_KangoCD (������Ō�P��CD)
    '   �ҏW�d�l
    '       �Ζ��L��(2)�{Space(5)�{�Ζ�����(3)�{���R�敪(1)�{�m�蕔���t���O(1)+������Ō�P��CD(4) + ��]�R�����g(20) +���ԔN�x(44)
    '       �S�U�O�o�C�g��
    '           1) �Ζ��L��---�P�޲Ėڂ���V�޲ĕ��iTrim�ŗ]���Ƚ�߰����Ȃ��j
    '           2) KinmuCD---�W�޲Ėڂ���R�޲ĕ��iTrim�ŗ]���Ƚ�߰����Ȃ��j
    '           3) ���R�敪---11�޲Ėڂ���P�޲ĕ�
    '           4) �m�蕔��FLG---12�޲Ėڂ���P�o�C�g��
    '               �i"1"�F�������m���ް��C"0"("1"�ȊO):�Y�������m���ް��A���́A�\���ް��j
    '           5) ������Ō�P��CD---13�޲Ėڂ���4�޲�     'Add Tanaka 2003/06/23
    '�@�@�@�@�@��2014/04/10 Bando Mod
    '           6) ��]�R�����g 16�޲Ėڂ���Q�O�o�C�g��
    '           7) ���ԔN�x---36�޲Ėڂ���S�S�޲ĕ�
    '******************************************************************************************
    '2015/04/10 Bando Upd Start =======================================
    'Public Function Set_KinmuMark(ByVal p_KinmuCD As String, ByVal p_RiyuKBN As String, ByVal p_Flg As String, ByVal p_KangoCD As String, ByVal p_Time As String) As Object
    Public Function Set_KinmuMark(ByVal p_KinmuCD As String, ByVal p_RiyuKBN As String, ByVal p_Flg As String, ByVal p_KangoCD As String, ByVal p_Time As String, ByVal p_Comment As String) As Object
        On Error GoTo Set_KinmuMark
        Const W_SUBNAME As String = "BasNSK0000H Set_KinmuMark"

        Dim w_SprText As String

        If IsNumeric(p_KinmuCD) = False Then
            Set_KinmuMark = CObj(Space(69))
            Exit Function
            'ElseIf CShort(p_KinmuCD) < 0 Or CShort(p_KinmuCD) > UBound(g_KinmuM) Then
        ElseIf CShort(p_KinmuCD) <= 0 Or CShort(p_KinmuCD) > UBound(g_KinmuM) Then
            Set_KinmuMark = CObj(Space(69))
            Exit Function
        Else
            '�L��
            w_SprText = g_KinmuM(CShort(p_KinmuCD)).Mark
            If 7 - General.paLenB(w_SprText) >= 0 Then
                w_SprText = w_SprText & Space(7 - General.paLenB(w_SprText))
            End If
            '���ށi�Ζ����ނ͏�����������3�޲ĕ��m�ۂ���j
            w_SprText = w_SprText & p_KinmuCD
            If 3 - General.paLenB(p_KinmuCD) >= 0 Then
                w_SprText = w_SprText & Space(3 - General.paLenB(p_KinmuCD))
            End If
            '���R�敪
            w_SprText = w_SprText & Left(p_RiyuKBN & Space(1), 1)
            '�m�蕔���׸�
            If p_Flg = "" Then
                w_SprText = w_SprText & "0"
            Else
                w_SprText = w_SprText & Left(p_Flg, 1)
            End If
            '�����Ζ��̊Ō�P��CD
            w_SprText = w_SprText & Left(p_KangoCD & Space(10), 10)

            '��]
            w_SprText = w_SprText & General.paLeftB(p_Comment & Space(20), 20)

            '���Ԑ�
            '2017/09/29 Yamanishi Upd ---------------------------------------------------------------------------------------------------------
            'w_SprText = w_SprText & Left(p_Time & Space(112), 112)
            p_Time = Trim(p_Time)
            If p_Time.Length <= 112 Then
                w_SprText = w_SprText & Left(p_Time & Space(112), 112)
            Else
                w_SprText = w_SprText & GenerateRandomKeyAndSetDictionary(112, p_Time, m_dicRandomKey2TimeNenkyu)
            End If
            '----------------------------------------------------------------------------------------------------------------------------------
            '�ҏW�������ԋp����
            Set_KinmuMark = CObj(w_SprText)
        End If

        Exit Function
Set_KinmuMark:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Function
    '2015/04/10 Bando Upd End   =======================================

    '2018/03/08 Yamanishi Add Start ------------------------------------------------------------------------------------------
    ''' <summary>
    ''' ���ԋx�̊e���ڂ��玞�ԋx������쐬
    ''' </summary>
    ''' <param name="p_BunruiCD">�x�ɕ��ރR�[�h</param>
    ''' <param name="p_FromTime">�J�n����</param>
    ''' <param name="p_ToTime">�I������</param>
    ''' <param name="p_DateKbn">�����t���O</param>
    ''' <param name="p_NenkyuTime">�N�x�擾����</param>
    ''' <param name="p_HolSubFlg">�x�e���Z�t���O</param>
    ''' <param name="p_DayTime">���Ύ���</param>
    ''' <param name="p_NightTime">��Ύ���</param>
    ''' <param name="p_NextNightTime">������Ύ���</param>
    ''' <returns>���ԋx������</returns>
    Public Function Set_NenkyuTime(ByVal p_BunruiCD As Object, ByVal p_FromTime As Object, ByVal p_ToTime As Object,
                                   ByVal p_DateKbn As Object, ByVal p_NenkyuTime As Object, ByVal p_HolSubFlg As Object,
                                   ByVal p_DayTime As Object, ByVal p_NightTime As Object, ByVal p_NextNightTime As Object) As String
        Const W_SUBNAME As String = "BasNSK0000H Set_NenkyuTime"

        Const WC_BunuruiCDLength As Integer = 2
        Const WC_TimeLength As Integer = 4
        Const WC_KbnLength As Integer = 1

        Dim w_Time As String
        Try
            '�x�ɕ��ރR�[�h
            w_Time = General.paFormatSpace(Convert.ToString(p_BunruiCD), WC_BunuruiCDLength)

            '�J�n����
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_FromTime), WC_TimeLength)

            '�I������
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_ToTime), WC_TimeLength)

            '�����t���O
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_DateKbn), WC_KbnLength)

            '�N�x�擾����
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_NenkyuTime), WC_TimeLength)

            '�x�e���Z�t���O
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_HolSubFlg), WC_KbnLength)

            '���Ύ���
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_DayTime), WC_TimeLength)

            '��Ύ���
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_NightTime), WC_TimeLength)

            '������Ύ���
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_NextNightTime), WC_TimeLength)

            Return w_Time

        Catch ex As Exception
            Call General.paTrpMsg(Err.Number, W_SUBNAME)
            End
        End Try
    End Function

    ''' <summary>
    ''' ���ԋx�����񂩂�x�ɕ��ނƎ擾���Ԃ��擾����
    ''' </summary>
    ''' <param name="p_Time">���ԋx������</param>
    ''' <param name="p_BunruiCD">�x�ɕ��ރR�[�h</param>
    ''' <param name="p_NenkyuTime">�N�x�擾����</param>
    ''' <returns>�����ϕ�����菜�������ԋx������</returns>
    Public Function Get_NenkyuTime(ByVal p_Time As String,
                                   ByRef p_BunruiCD As String, ByRef p_NenkyuTime As Integer) As String
        Return Get_NenkyuTime(p_Time, p_BunruiCD, 0, 0, "", p_NenkyuTime, "", 0, 0, 0)
    End Function

    ''' <summary>
    ''' ���ԋx�����񂩂���΁E��΁E����Ύ��Ԃ��擾����
    ''' </summary>
    ''' <param name="p_Time">���ԋx������</param>
    ''' <param name="p_DayTime">���Ύ���</param>
    ''' <param name="p_NightTime">��Ύ���</param>
    ''' <param name="p_NextNightTime">������Ύ���</param>
    ''' <returns>�����ϕ�����菜�������ԋx������</returns>
    Public Function Get_NenkyuTime(ByVal p_Time As String,
                                   ByRef p_DayTime As Integer, ByRef p_NightTime As Integer, ByRef p_NextNightTime As Integer) As String
        Return Get_NenkyuTime(p_Time, "", 0, 0, "", 0, "", p_DayTime, p_NightTime, p_NextNightTime)
    End Function

    ''' <summary>
    ''' ���ԋx�����񂩂�e��p�����[�^�𕜌�����
    ''' </summary>
    ''' <param name="p_Time">���ԋx������</param>
    ''' <param name="p_BunruiCD">�x�ɕ��ރR�[�h</param>
    ''' <param name="p_FromTime">�J�n����</param>
    ''' <param name="p_ToTime">�I������</param>
    ''' <param name="p_DateKbn">�����t���O</param>
    ''' <param name="p_NenkyuTime">�N�x�擾����</param>
    ''' <param name="p_HolSubFlg">�x�e���Z�t���O</param>
    ''' <param name="p_DayTime">���Ύ���</param>
    ''' <param name="p_NightTime">��Ύ���</param>
    ''' <param name="p_NextNightTime">������Ύ���</param>
    ''' <returns>�����ϕ�����菜�������ԋx������</returns>
    Public Function Get_NenkyuTime(ByVal p_Time As String,
                                   ByRef p_BunruiCD As String, ByRef p_FromTime As Integer, ByRef p_ToTime As Integer,
                                   ByRef p_DateKbn As String, ByRef p_NenkyuTime As Integer, ByRef p_HolSubFlg As String,
                                   ByRef p_DayTime As Integer, ByRef p_NightTime As Integer, ByRef p_NextNightTime As Integer) As String
        Const W_SUBNAME As String = "BasNSK0000H Get_NenkyuTime"

        Const WC_BunuruiCDLength As Integer = 2
        Const WC_TimeLength As Integer = 4
        Const WC_KbnLength As Integer = 1

        Dim w_Length As Integer
        Dim w_Position As Integer
        Try
            w_Position = 1

            '�x�ɕ��ރR�[�h
            w_Length = WC_BunuruiCDLength
            p_BunruiCD = Mid(p_Time, w_Position, w_Length)
            w_Position = w_Position + w_Length

            '�J�n����
            w_Length = WC_TimeLength
            If Not Integer.TryParse(Mid(p_Time, w_Position, w_Length), p_FromTime) Then
                p_FromTime = 0
            End If
            w_Position = w_Position + w_Length

            '�I������
            w_Length = WC_TimeLength
            If Not Integer.TryParse(Mid(p_Time, w_Position, w_Length), p_ToTime) Then
                p_ToTime = 0
            End If
            w_Position = w_Position + w_Length

            '�����t���O
            w_Length = WC_KbnLength
            p_DateKbn = Mid(p_Time, w_Position, w_Length)
            w_Position = w_Position + w_Length

            '�N�x�擾����
            w_Length = WC_TimeLength
            If Not Integer.TryParse(Mid(p_Time, w_Position, w_Length), p_NenkyuTime) Then
                p_NenkyuTime = 0
            End If
            w_Position = w_Position + w_Length

            '�x�e���Z�t���O
            w_Length = WC_KbnLength
            p_HolSubFlg = Mid(p_Time, w_Position, w_Length)
            w_Position = w_Position + w_Length

            '���Ύ���
            w_Length = WC_TimeLength
            If Not Integer.TryParse(Mid(p_Time, w_Position, w_Length), p_DayTime) Then
                p_DayTime = 0
            End If
            w_Position = w_Position + w_Length

            '��Ύ���
            w_Length = WC_TimeLength
            If Not Integer.TryParse(Mid(p_Time, w_Position, w_Length), p_NightTime) Then
                p_NightTime = 0
            End If
            w_Position = w_Position + w_Length

            '������Ύ���
            w_Length = WC_TimeLength
            If Not Integer.TryParse(Mid(p_Time, w_Position, w_Length), p_NextNightTime) Then
                p_NextNightTime = 0
            End If
            w_Position = w_Position + w_Length

            Return Mid(p_Time, w_Position)

        Catch ex As Exception
            Call General.paTrpMsg(Err.Number, W_SUBNAME)
            End
        End Try
    End Function

    ''' <summary>
    ''' HHmm�`���𕪐��ɕϊ�
    ''' </summary>
    ''' <param name="p_HHmm">HHmm</param>
    ''' <returns>����</returns>
    Public Function HHmmToMin(ByVal p_HHmm As Integer) As Integer
        Return (p_HHmm \ 100) * 60 + (p_HHmm Mod 100)
    End Function
    '2018/03/08 Yamanishi Add End --------------------------------------------------------------------------------------------

    Public Function Get_KinmuTipText(ByVal p_KinmuCD As String) As String
        On Error GoTo Get_KinmuTipText
        Const W_SUBNAME As String = "BasNSK0000H Get_KinmuTipText"

        Dim w_str As String

        If IsNumeric(p_KinmuCD) = False Then
            Get_KinmuTipText = ""
            Exit Function
        ElseIf CShort(p_KinmuCD) <= 0 Or UBound(g_KinmuM) < CShort(p_KinmuCD) Then
            Get_KinmuTipText = ""
            Exit Function
        End If

        '�����Ζ����H
        Select Case g_KinmuM(CShort(p_KinmuCD)).HFlg
            Case "1" '�S��
                w_str = g_KinmuM(CShort(p_KinmuCD)).KinmuName
            Case "2" '����
                w_str = g_KinmuM(CShort(g_KinmuM(CShort(p_KinmuCD)).AMCD)).KinmuName & "�^" & g_KinmuM(CShort(g_KinmuM(CShort(p_KinmuCD)).PMCD)).KinmuName
            Case Else
                w_str = ""
        End Select

        Get_KinmuTipText = w_str

        Exit Function
Get_KinmuTipText:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Function

    '*****************************************************************************************************
    '   ���گ�ނ̋x�ݕ��ދL����������e�f�[�^���ɕ�������
    '   ���Ұ��Fp_Val(I)�i�Ζ����ށj
    '           p_UniqueSeqNO(O)�iUNIQUESEQNO�j
    '           p_AppliCD(O)�i�͏o���ށj
    '           p_HolBunruiCD(O)�i�x�ݕ��޺��ށj
    '           p_GetContentsKBN(O)�i�擾���e�敪�j
    '           p_intTimeFrom(O)�i����FROM�j
    '           p_intTimeTo(O)�i����TO�j
    '           p_strNextDayFlg(O)�i����FLG�j
    '   �ҏW�d�l
    '       �x�ݕ��ދL��(2)�{Space(5)�{�x�ݕ��޺���(2)�{UNIQUESEQNO(18)�{�͏o����(6)+�擾���e�敪(1)�{����FROM(4)�{����TO(4)�{����FLG(1)
    '       �S�S�R�o�C�g��
    '           1) �x�ݕ��ދL��---�P�޲Ėڂ���V�޲ĕ��iTrim�ŗ]���Ƚ�߰����Ȃ��j
    '           2) �x�ݕ���CD---�W�޲Ėڂ���Q�޲ĕ��iTrim�ŗ]���Ƚ�߰����Ȃ��j
    '           3) �͏oUniqueSeqNO---10�޲Ėڂ���P�W�޲ĕ�
    '           4) �͏oCD---28�޲Ėڂ���U�޲ĕ�
    '           5) �擾���e�敪---34�޲Ėڂ���P�޲ĕ�
    '               �i"1":�S���A"2":�O���A"3":�㔼�A"4":���ԔN�x�j
    '           6) ����FROM---35�޲Ėڂ���S�޲ĕ�
    '           7) ����TO---39�޲Ėڂ���S�޲ĕ�
    '           8) ����FLG---43�޲Ėڂ���P�޲ĕ�
    '*****************************************************************************************************
    Public Sub Get_AppliMark(ByVal p_Val As Object, ByRef p_UniqueSeqNO As String, ByRef p_AppliCD As String, ByRef p_HolBunruiCD As String, ByRef p_GetContentsKBN As String, ByRef p_intTimeFrom As Short, ByRef p_intTimeTo As Short, ByRef p_strNextDayFlg As String)
        On Error GoTo Get_AppliMark
        Const W_SUBNAME As String = "BasNSK0000H Get_AppliMark"

        Dim w_str As String

        '�Z���̒l���ݒ肳��Ă��Ȃ��ꍇ
        If IsDBNull(p_Val) = True Then
            p_HolBunruiCD = ""
            p_UniqueSeqNO = ""
            p_AppliCD = ""
            p_GetContentsKBN = ""
            p_intTimeFrom = 0
            p_intTimeTo = 0
            p_strNextDayFlg = ""
            Exit Sub
        End If

        '�Ζ��L���ȊO�͔��p�Ƃ���
        w_str = CStr(p_Val)
        w_str = Right(w_str, 36)

        '���گ�ޓ\��t��������𕪊�����
        p_HolBunruiCD = Trim(Left(w_str, 2))
        p_UniqueSeqNO = Trim(Mid(w_str, 3, 18))
        p_AppliCD = Trim(Mid(w_str, 21, 6))
        p_GetContentsKBN = Trim(Mid(w_str, 27, 1))
        If IsNumeric(Mid(w_str, 28, 4)) Then
            p_intTimeFrom = CShort(Mid(w_str, 28, 4))
        Else
            p_intTimeFrom = 0
        End If
        If IsNumeric(Mid(w_str, 32, 4)) Then
            p_intTimeTo = CShort(Mid(w_str, 32, 4))
        Else
            p_intTimeTo = 0
        End If
        p_strNextDayFlg = Trim(Mid(w_str, 36, 1))

        Exit Sub
Get_AppliMark:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub

    '******************************************************************************************
    '   ���گ�ނɓ\��t����x�ݕ��ދL����ҏW����
    '   ���Ұ��Fp_UniqueSeqNO�iUNIQUESEQNO�j
    '           p_AppliCD�i�͏o���ށj
    '           p_HolBunruiCD�i�x�ݕ��޺��ށj
    '           p_GetContentsKBN�i�擾���e�敪�j
    '           p_intTimeFrom�i����FROM�j
    '           p_intTimeTo�i����TO�j
    '           p_strNextDayFlg�i����FLG�j
    '   �ҏW�d�l
    '       �x�ݕ��ދL��(2)�{Space(5)�{�x�ݕ��޺���(2)�{UNIQUESEQNO(18)�{�͏o����(6)+�擾���e�敪(1)�{����FROM(4)�{����TO(4)�{����FLG(1)
    '       �S�S�R�o�C�g��
    '           1) �x�ݕ��ދL��---�P�޲Ėڂ���V�޲ĕ��iTrim�ŗ]���Ƚ�߰����Ȃ��j
    '           2) �x�ݕ���CD---�W�޲Ėڂ���Q�޲ĕ��iTrim�ŗ]���Ƚ�߰����Ȃ��j
    '           3) �͏oUniqueSeqNO---10�޲Ėڂ���P�W�޲ĕ�
    '           4) �͏oCD---28�޲Ėڂ���U�޲ĕ�
    '           5) �擾���e�敪---34�޲Ėڂ���P�޲ĕ�
    '               �i"1":�S���A"2":�O���A"3":�㔼�A"4":���ԔN�x�j
    '           6) ����FROM---35�޲Ėڂ���S�޲ĕ�
    '           7) ����TO---39�޲Ėڂ���S�޲ĕ�
    '           8) ����FLG---43�޲Ėڂ���P�޲ĕ�
    '******************************************************************************************
    Public Function Set_AppliMark(ByVal p_UniqueSeqNO As String, ByVal p_AppliCD As String, ByVal p_HolBunruiCD As String, ByVal p_GetContentsKBN As String, ByVal p_intTimeFrom As Short, ByVal p_intTimeTo As Short, ByVal p_strNextDayFlg As String) As Object
        On Error GoTo Set_AppliMark
        Const W_SUBNAME As String = "BasNSK0000H Set_AppliMark"

        Dim w_SprText As String
        Dim w_intLoop As Short
        Dim w_intIndex As Short
        Dim w_blnFLG As Boolean

        w_blnFLG = False
        For w_intLoop = 1 To UBound(g_HolidayBunruiM)
            If g_HolidayBunruiM(w_intLoop).CD = p_HolBunruiCD Then
                w_blnFLG = True
                w_intIndex = w_intLoop
                Exit For
            End If
        Next w_intLoop

        If w_blnFLG = False Then
            w_SprText = Space(34)
            '����FROM
            w_SprText = w_SprText & "0000"
            '����TO
            w_SprText = w_SprText & "0000"
            w_SprText = w_SprText & Space(1)
            Set_AppliMark = CObj(w_SprText)
            Exit Function
        Else
            '�L��
            w_SprText = g_HolidayBunruiM(w_intIndex).Mark
            If 7 - General.paLenB(w_SprText) >= 0 Then
                w_SprText = w_SprText & Space(7 - General.paLenB(w_SprText))
            End If
            '����
            w_SprText = w_SprText & p_HolBunruiCD
            If 2 - General.paLenB(p_HolBunruiCD) >= 0 Then
                w_SprText = w_SprText & Space(2 - General.paLenB(p_HolBunruiCD))
            End If
            'UNIQUESEQNO
            w_SprText = w_SprText & p_UniqueSeqNO
            If 18 - General.paLenB(p_UniqueSeqNO) >= 0 Then
                w_SprText = w_SprText & Space(18 - General.paLenB(p_UniqueSeqNO))
            End If
            '�͏o����
            w_SprText = w_SprText & p_AppliCD
            If 6 - General.paLenB(p_AppliCD) >= 0 Then
                w_SprText = w_SprText & Space(6 - General.paLenB(p_AppliCD))
            End If
            '�擾���e�敪
            w_SprText = w_SprText & Left(p_GetContentsKBN & Space(1), 1)
            '����FROM
            w_SprText = w_SprText & Left(Format(p_intTimeFrom, "0000"), 4)
            '����TO
            w_SprText = w_SprText & Left(Format(p_intTimeTo, "0000"), 4)
            '����FLG
            w_SprText = w_SprText & Left(p_strNextDayFlg & Space(1), 1)

            '�ҏW�������ԋp����
            Set_AppliMark = CObj(w_SprText)

        End If

        Exit Function
Set_AppliMark:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Function

    Public Function Get_AppliTipText(ByVal p_HolBunruiCD As String) As String
        On Error GoTo Get_AppliTipText
        Const W_SUBNAME As String = "BasNSK0000H Get_AppliTipText"

        Dim w_intLoop As Short
        Dim w_intIndex As Short

        w_intIndex = -1
        For w_intLoop = 1 To UBound(g_HolidayBunruiM)
            If g_HolidayBunruiM(w_intLoop).CD = p_HolBunruiCD Then
                w_intIndex = w_intLoop
                Exit For
            End If
        Next w_intLoop

        If w_intIndex = -1 Then
            Get_AppliTipText = ""
            Exit Function
        End If

        Get_AppliTipText = g_HolidayBunruiM(w_intIndex).HolidayName

        Exit Function
Get_AppliTipText:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Function

    '*****************************************************************************************************
    '   ���گ�ނ̋Ζ��L����������e�f�[�^���ɕ�������
    '   ���Ұ��Fp_Val(I)�i�Ζ����ށj
    '           p_KinmuCD(O)�i�������Ζ����ށj
    '           p_GroupCD(O)�i�������O���[�v���ށj
    '   �ҏW�d�l
    '       �������Ζ��L��(2)�{Space(5)�{�������Ζ�����(3)�{�������O���[�v����(2)�{�Q����(�������Ζ��L��(2)�{Space(5)�{�������Ζ�����(3)�{�������O���[�v����(2))�{�R���ڈȍ~�E�E�E
    '       �S�P�Q�o�C�g�~������
    '           1) �������Ζ��L��---�P�޲Ėڂ���V�޲ĕ��iTrim�ŗ]���Ƚ�߰����Ȃ��j
    '           2) �������Ζ�CD---�W�޲Ėڂ���R�޲ĕ��iTrim�ŗ]���Ƚ�߰����Ȃ��j
    '           3) �������O���[�vCD---11�޲Ėڂ���Q�޲ĕ�
    '           4) �Q���ڈȍ~---13�޲Ėڂ���12�o�C�g��
    '*****************************************************************************************************
    Public Sub Get_DutyMark(ByVal p_Val As Object, ByRef p_KinmuCD As Object, ByRef p_GroupCD As Object)
        On Error GoTo Get_DutyMark
        Const W_SUBNAME As String = "BasNSK0000H Get_DutyMark"

        Dim w_str As String
        Dim w_Str2 As String
        Dim w_intCount As Short
        Dim w_intLoop As Short

        '�Z���̒l���ݒ肳��Ă��Ȃ��ꍇ
        If IsDBNull(p_Val) = True Then
            ReDim p_KinmuCD(0)
            ReDim p_GroupCD(0)
            Exit Sub
        End If

        '�Ζ��L���ȊO�͔��p�Ƃ���
        w_str = CStr(p_Val)
        w_intCount = General.paLenB(w_str) / 12
        ReDim p_KinmuCD(w_intCount)
        ReDim p_GroupCD(w_intCount)
        For w_intLoop = 1 To w_intCount
            w_Str2 = General.paLeftB(w_str, w_intLoop * 12)
            w_Str2 = Right(w_Str2, 5)

            '���گ�ޓ\��t��������𕪊�����
            p_KinmuCD(w_intLoop) = Trim(Left(w_Str2, 3)) '�Ζ��L���͏�����������3�o�C�g���m�ۂ���
            p_GroupCD(w_intLoop) = Trim(Mid(w_Str2, 4, 2))
        Next w_intLoop

        Exit Sub
Get_DutyMark:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub

    '******************************************************************************************
    '   ���گ�ނɓ\��t����Ζ��L����ҏW����
    '   ���Ұ��Fp_KinmuCD�i�������Ζ����ށj
    '           p_GroupCD�i�������O���[�v���ށj
    '   �ҏW�d�l
    '       �������Ζ��L��(2)�{Space(5)�{�������Ζ�����(3)�{�������O���[�v����(2)�{�Q����(�������Ζ��L��(2)�{Space(5)�{�������Ζ�����(3)�{�������O���[�v����(2))�{�R���ڈȍ~�E�E�E
    '       �S�P�Q�o�C�g�~������
    '           1) �������Ζ��L��---�P�޲Ėڂ���V�޲ĕ��iTrim�ŗ]���Ƚ�߰����Ȃ��j
    '           2) �������Ζ�CD---�W�޲Ėڂ���R�޲ĕ��iTrim�ŗ]���Ƚ�߰����Ȃ��j
    '           3) �������O���[�vCD---11�޲Ėڂ���Q�޲ĕ�
    '           4) �Q���ڈȍ~---13�޲Ėڂ���12�o�C�g��
    '******************************************************************************************
    Public Function Set_DutyMark(ByVal p_KinmuCD As Object, ByVal p_GroupCD As Object) As Object
        On Error GoTo Set_DutyMark
        Const W_SUBNAME As String = "BasNSK0000H Set_DutyMark"

        Dim w_SprText As String
        Dim w_intLoop As Short

        Set_DutyMark = ""

        For w_intLoop = 1 To UBound(p_KinmuCD)
            If IsNumeric(p_KinmuCD(w_intLoop)) = False Then
            ElseIf CShort(p_KinmuCD(w_intLoop)) < 0 Or CShort(p_KinmuCD(w_intLoop)) > UBound(g_KinmuM) Then
            Else
                '�L��
                w_SprText = g_KinmuM(CShort(p_KinmuCD(w_intLoop))).Mark
                If 7 - General.paLenB(w_SprText) >= 0 Then
                    w_SprText = w_SprText & Space(7 - General.paLenB(w_SprText))
                End If
                '���ށi�Ζ����ނ͏�����������3�޲ĕ��m�ۂ���j
                w_SprText = w_SprText & p_KinmuCD(w_intLoop)
                If 3 - General.paLenB(p_KinmuCD(w_intLoop)) >= 0 Then
                    w_SprText = w_SprText & Space(3 - General.paLenB(p_KinmuCD(w_intLoop)))
                End If
                '�O���[�v����
                w_SprText = w_SprText & Left(p_GroupCD(w_intLoop) & Space(2), 2)

                '�ҏW�������ԋp����
                Set_DutyMark = Set_DutyMark & CObj(w_SprText)
            End If
        Next w_intLoop

        If Set_DutyMark = "" Then
            Set_DutyMark = CObj(Space(12))
        End If

        Exit Function
Set_DutyMark:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Function

    '�\�������i�[�p
    Public m_HyoujijunMeiDate() As HyoujijunMeiDate_Type

    Public Structure HyoujijunMeiDate_Type
        Dim HName As String '�\��������
        Dim HMasterCD As String '�\����CD
    End Structure

    ''' <summary>
    ''' �����L�[����
    ''' </summary>
    ''' <param name="p_key"></param>
    ''' <returns></returns>
    ''' <remarks>�L�[�{�[�h�Ή��L�[���ǂ������肷��</remarks>
    Public Function IsNumOrFuncKey(ByVal p_key As System.Windows.Forms.Keys) As Boolean
        Dim w_PreErrorProc As String = General.g_ErrorProc
        General.g_ErrorProc = "NSC0000HA Get_KeyBoardKinmu"

        Dim rtnFlg As Boolean = False

        Try
            Select Case p_key
                Case Keys.NumPad0, Keys.NumPad1, Keys.NumPad2, Keys.NumPad3, Keys.NumPad4, _
                     Keys.NumPad5, Keys.NumPad6, Keys.NumPad7, Keys.NumPad8, Keys.NumPad9
                    '�e���L�[�̐���
                    rtnFlg = True

                Case Keys.D0, Keys.D1, Keys.D2, Keys.D3, Keys.D4, _
                     Keys.D5, Keys.D6, Keys.D7, Keys.D8, Keys.D9
                    '�L�[�{�[�h�̐���
                    rtnFlg = True

                Case Keys.F1, Keys.F2, Keys.F3, Keys.F4, Keys.F5, Keys.F6, _
                     Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12
                    '�t�@���N�V�����L�[
                    rtnFlg = True

                Case Else
            End Select

            Return rtnFlg
        Catch ex As Exception
            Throw
        End Try
    End Function

    '2014/04/23 Saijo add start P-06979---------------------------------------------------------------------------
    ''' <summary>
    ''' ���ڐݒ�u�Ζ��L���S�p�Q�����Ή��t���O�v�擾
    ''' </summary>
    ''' <param name="p_HospitalCD">�a�@CD</param>
    ''' <returns>String ("0"�F�Ή����Ȃ��A"1":�Ή�����)</returns>
    ''' <remarks>�Ζ��L���S�p�Q�����Ή��t���O(0�F�Ή����Ȃ��A1:�Ή�����)</remarks>
    Public Function Get_ItemValue(ByVal p_HospitalCD As String) As String
        Dim w_PreErrorProc As String = General.g_ErrorProc
        General.g_ErrorProc = "NSK0000H Get_ItemValue"

        Try
            Get_ItemValue = General.paGetItemValue( _
            General.G_STRMAINKEY1, General.G_STRSUBKEY1, "KINMUEMSECONDFLG", "0", p_HospitalCD)

        Catch ex As Exception
            Throw
        End Try
    End Function
    '2014/04/23 Saijo add end P-06979-------------------------------------------------------------------------------

    '2017/08/24 Angelo add st---------------------------------------------------------------------------------------
    '��ʕ\���p�̕ҏW����
    Public Function EditData(ByVal p_objValue As Object, ByVal p_intEditMode As Integer) As String
        Const W_SUBNAME As String = "BasNSK0000H EditData"

        Dim w_strEditValue As String

        EditData = ""
        Try
            '������
            w_strEditValue = ""

            If p_intEditMode = G_EDITMODE_NO Then
                '0��00�ɕϊ�
                w_strEditValue = General.paFormatZero(p_objValue, 2)
            ElseIf p_intEditMode = G_EDITMODE_DATETIME Then
                If p_objValue <> 0 Then
                    'yyyyMMddHHmmss��yyyy/MM/dd HH:mm:ss�ɕϊ�
                    w_strEditValue = Format(p_objValue, "0000/00/00 00:00:00")
                End If
            End If

            EditData = w_strEditValue
        Catch es As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
        End Try
    End Function
    Public Sub GetNenkyuContentsKbnAndHolCD(ByVal p_KinmuCD As String, ByRef p_GetContentsKbn As String, ByRef p_HolCD As String)
        Try
            p_GetContentsKbn = ""
            p_HolCD = ""
            If IsNumeric(p_KinmuCD) Then
                If 0 <= Integer.Parse(p_KinmuCD) AndAlso Integer.Parse(p_KinmuCD) <= UBound(g_KinmuM) Then
                    If g_KinmuM(Integer.Parse(p_KinmuCD)).HFlg = "2" Then
                        '�����Ζ��׸�
                        If IsNumeric(g_KinmuM(Integer.Parse(p_KinmuCD)).AMCD) Then
                            '�`�l�Ζ�CD
                            If 0 <= Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).AMCD) AndAlso Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).AMCD) <= UBound(g_KinmuM) Then
                                If Not String.IsNullOrEmpty(g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).AMCD)).HolBunruiCD) AndAlso
                                            g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).AMCD)).HolBunruiCD <> General.G_STRDAIKYUBUNRUI Then
                                    '�x�ݕ���CD
                                    p_GetContentsKbn = "2"
                                    p_HolCD = g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).AMCD)).HolBunruiCD
                                End If
                            End If
                        End If
                        If IsNumeric(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD) Then
                            '�o�l�Ζ�CD
                            If 0 <= Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD) AndAlso Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD) <= UBound(g_KinmuM) Then
                                If Not String.IsNullOrEmpty(g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD)).HolBunruiCD) AndAlso
                                            g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD)).HolBunruiCD <> General.G_STRDAIKYUBUNRUI Then
                                    '�x�ݕ���CD
                                    If p_GetContentsKbn = "2" Then
                                        p_GetContentsKbn = "2,3"
                                        p_HolCD = p_HolCD & "," & g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD)).HolBunruiCD
                                    Else
                                        p_GetContentsKbn = "3"
                                        p_HolCD = g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD)).HolBunruiCD
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If Not String.IsNullOrEmpty(g_KinmuM(Integer.Parse(p_KinmuCD)).HolBunruiCD) AndAlso
                                    g_KinmuM(Integer.Parse(p_KinmuCD)).HolBunruiCD <> General.G_STRDAIKYUBUNRUI Then
                            '�x�ݕ���CD
                            p_GetContentsKbn = "1"
                            p_HolCD = g_KinmuM(Integer.Parse(p_KinmuCD)).HolBunruiCD
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub
    '*****************************************************************************************************
    '   ���گ�ނ̎��ԔN�x��������e�f�[�^���ɕ�������
    '   ���Ұ��Fp_Str(I)�i���ԔN�x������j
    '           p_NenkyuDetail(O)�i���ԔN�x�ڍ׏��j
    '   �ҏW�d�l
    '       �x�ݕ��޺���(2)�{����FROM(4)�{����TO(4)�{����FLG(1)�{�Q����(�x�ݕ��޺���(2)�{����FROM(4)�{����TO(4)�{����FLG(1))�{�R���ڈȍ~�E�E�E
    '       �S�P�P�o�C�g�~������
    '           1) �x�ݕ���CD---�P�޲Ėڂ���Q�޲ĕ��iTrim�ŗ]���Ƚ�߰����Ȃ��j
    '           2) �擾���e�敪---34�޲Ėڂ���P�޲ĕ�
    '               �i"1":�S���A"2":�O���A"3":�㔼�A"4":���ԔN�x�j
    '           3) ����FROM---�R�޲Ėڂ���S�޲ĕ�
    '           4) ����TO---�V�޲Ėڂ���S�޲ĕ�
    '           5) ����FLG---11�޲Ėڂ���P�޲ĕ�
    '           6) �Q���ڈȍ~---12�޲Ėڂ���11�o�C�g��
    '*****************************************************************************************************
    Public Sub Get_NenkyuDetail(ByVal p_Str As String, ByVal p_Str2 As String, ByRef p_NenkyuDetail() As NenkyuDetail_Type, ByVal p_HolCD As String)
        Const W_SUBNAME As String = "NSK0000HA Get_NenkyuDetail"

        '2018/03/08 Yamanishi Upd -----------------------------------
        'Dim w_Loop As Integer
        'Dim w_Index As Integer
        'Dim w_RecCnt As Integer
        'Dim w_Pos As Integer
        'Dim w_StartTime As String 'FromTime
        'Dim w_EndTime As String 'ToTime
        'Dim w_NenkyuTime As Integer '���ԔN�x
        'Dim w_DayTime As Integer '���Ύ���
        'Dim w_NightTime As Integer '��Ύ���
        'Dim w_NextNightTime As Integer '������Ύ���
        Dim w_Index As Integer
        Dim w_RecCnt As Integer
        '------------------------------------------------------------
        Dim w_obj As Object
        Try
            '2018/03/08 Yamanishi Upd Start ------------------------------------------------------------------
            ''���ԔN�x�����鎞�̂�
            'If p_Str <> "" Then
            '    '���ԔN�x�����擾�i�P���ɂ��Q�W�޲āj
            '    w_RecCnt = Len(p_Str) / 28
            '    w_Index = UBound(p_NenkyuDetail)
            '    ReDim Preserve p_NenkyuDetail(w_Index + w_RecCnt)

            '    w_Pos = 1

            '    '���گ�ޓ\��t��������𕪊�����
            '    For w_Loop = 1 To w_RecCnt
            '        '�擾���e�敪(4:���ԔN�x)
            '        p_NenkyuDetail(w_Index + w_Loop).GetContentsKbn = "4"

            '        'HolidayBunruiCD
            '        p_NenkyuDetail(w_Index + w_Loop).HolidayBunruiCD = Mid(p_Str, w_Pos, 2)

            '        w_Pos = w_Pos + 2

            '        'FromTime
            '        w_StartTime = Mid(p_Str, w_Pos, 4)
            '        If IsNumeric(w_StartTime) Then
            '            p_NenkyuDetail(w_Index + w_Loop).FromTime = Integer.Parse(w_StartTime)
            '        Else
            '            p_NenkyuDetail(w_Index + w_Loop).FromTime = 0
            '        End If

            '        w_Pos = w_Pos + 4

            '        'ToTime
            '        w_EndTime = Mid(p_Str, w_Pos, 4)
            '        If IsNumeric(w_EndTime) Then
            '            p_NenkyuDetail(w_Index + w_Loop).ToTime = Integer.Parse(w_EndTime)
            '        Else
            '            p_NenkyuDetail(w_Index + w_Loop).ToTime = 0
            '        End If

            '        w_Pos = w_Pos + 4

            '        'DateKbn
            '        p_NenkyuDetail(w_Index + w_Loop).DateKbn = Mid(p_Str, w_Pos, 1)

            '        w_Pos = w_Pos + 1

            '        '���ԔN�x
            '        w_NenkyuTime = Integer.Parse(Mid(p_Str, w_Pos, 4))
            '        If IsNumeric(w_NenkyuTime) Then
            '            p_NenkyuDetail(w_Index + w_Loop).NenkyuTime = Integer.Parse(w_NenkyuTime)
            '        Else
            '            p_NenkyuDetail(w_Index + w_Loop).NenkyuTime = 0
            '        End If

            '        w_Pos = w_Pos + 4

            '        '�x�e���Z�t���O
            '        p_NenkyuDetail(w_Index + w_Loop).HolSubFlg = Mid(p_Str, w_Pos, 1)

            '        w_Pos = w_Pos + 1

            '        '���Ύ���
            '        w_DayTime = Integer.Parse(Mid(p_Str, w_Pos, 4))
            '        If IsNumeric(w_DayTime) Then
            '            p_NenkyuDetail(w_Index + w_Loop).DayTime = Integer.Parse(w_DayTime)
            '        Else
            '            p_NenkyuDetail(w_Index + w_Loop).DayTime = 0
            '        End If

            '        w_Pos = w_Pos + 4

            '        '��Ύ���
            '        w_NightTime = Integer.Parse(Mid(p_Str, w_Pos, 4))
            '        If IsNumeric(w_NightTime) Then
            '            p_NenkyuDetail(w_Index + w_Loop).NightTime = Integer.Parse(w_NightTime)
            '        Else
            '            p_NenkyuDetail(w_Index + w_Loop).NightTime = 0
            '        End If

            '        w_Pos = w_Pos + 4

            '        '������Ύ���
            '        w_NextNightTime = Integer.Parse(Mid(p_Str, w_Pos, 4))
            '        If IsNumeric(w_NextNightTime) Then
            '            p_NenkyuDetail(w_Index + w_Loop).NextNightTime = Integer.Parse(w_NextNightTime)
            '        Else
            '            p_NenkyuDetail(w_Index + w_Loop).NextNightTime = 0
            '        End If

            '        w_Pos = w_Pos + 4
            '    Next w_Loop
            'End If

            '���ԔN�x�����鎞�̂�
            While p_Str <> ""
                w_RecCnt = UBound(p_NenkyuDetail) + 1
                ReDim Preserve p_NenkyuDetail(w_RecCnt)

                With p_NenkyuDetail(w_RecCnt)
                    '�擾���e�敪(4:���ԔN�x)
                    .GetContentsKbn = General.G_GetContentsKbn_Time

                    p_Str = Get_NenkyuTime(p_Str,
                                           .HolidayBunruiCD, .FromTime, .ToTime,
                                           .DateKbn, .NenkyuTime, .HolSubFlg,
                                           .DayTime, .NightTime, .NextNightTime)

                End With
            End While
            '2018/03/08 Yamanishi Upd End --------------------------------------------------------------------

            '�S�����N�x�����鎞�̂�
            If p_Str2 <> "" Then
                '�N�x�����擾
                w_Index = UBound(p_NenkyuDetail)

                If p_Str2 = "2,3" Then
                    w_obj = General.paSplit(p_HolCD, ",")
                    ReDim Preserve p_NenkyuDetail(w_Index + 2)
                    '�擾���e�敪
                    p_NenkyuDetail(w_Index + 1).GetContentsKbn = "2"
                    'HolidayBunruiCD
                    p_NenkyuDetail(w_Index + 1).HolidayBunruiCD = w_obj(0)
                    'FromTime
                    p_NenkyuDetail(w_Index + 1).FromTime = 0
                    'ToTime
                    p_NenkyuDetail(w_Index + 1).ToTime = 0
                    'DateKbn
                    p_NenkyuDetail(w_Index + 1).DateKbn = "0"
                    '�擾���e�敪
                    p_NenkyuDetail(w_Index + 2).GetContentsKbn = "3"
                    'HolidayBunruiCD
                    p_NenkyuDetail(w_Index + 2).HolidayBunruiCD = w_obj(1)
                    'FromTime
                    p_NenkyuDetail(w_Index + 2).FromTime = 0
                    'ToTime
                    p_NenkyuDetail(w_Index + 2).ToTime = 0
                    'DateKbn
                    p_NenkyuDetail(w_Index + 2).DateKbn = "0"
                Else
                    ReDim Preserve p_NenkyuDetail(w_Index + 1)
                    '�擾���e�敪
                    p_NenkyuDetail(w_Index + 1).GetContentsKbn = p_Str2
                    'HolidayBunruiCD
                    p_NenkyuDetail(w_Index + 1).HolidayBunruiCD = p_HolCD
                    'FromTime
                    p_NenkyuDetail(w_Index + 1).FromTime = 0
                    'ToTime
                    p_NenkyuDetail(w_Index + 1).ToTime = 0
                    'DateKbn
                    p_NenkyuDetail(w_Index + 1).DateKbn = "0"
                End If
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub
    '2017/08/24 Angelo add en---------------------------------------------------------------------------------------
End Module