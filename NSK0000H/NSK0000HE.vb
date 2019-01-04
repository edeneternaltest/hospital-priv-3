Option Strict Off
Option Explicit On
Imports System.Text
Public Class frmNSK0000HE
    Inherits General.FormBase
    '/----------------------------------------------------------------------/
    '/
    '/    ���і��́F�Ō�x���V�X�e��(�Ζ��Ǘ�)
    '/ ��۸��і��́F�ۑ��ďo���
    '/        �h�c�FNSK0000HE
    '/        �T�v�F�ꎞ�ۑ��\��ꗗ�e�̈ꗗ�\���E�ۑ��E�ďo���s���B
    '/
    '/
    '/      �쐬�ҁF Angelo     CREATE 2017/08/04     REV 01.00
    '/      �X�V�ҁF            UPDATE     /  /      �y �z
    '/                                �X�V���e�F( )
    '/
    '/     Copyright (C) Inter co.,ltd 2000
    '/----------------------------------------------------------------------/
    '=======================================================
    '   �萔�錾
    '=======================================================
    '�X�v���b�h�̗�INDEX
    Private Const M_SAVESPR_COLIDX_NO As Integer = 0 '�ۑ��ԍ�
    Private Const M_SAVESPR_COLIDX_NAME As Integer = 1 '����҂̃��[�U�[
    Private Const M_SAVESPR_COLIDX_DATE As Integer = 2 '�ŏI�X�V����
    Private Const M_SAVESPR_COLIDX_BIKOU As Integer = 3 '���l

    Private Const M_SAVESPR_MAXROW As Integer = 5 '�ő�\���s��
    '=======================================================
    '   ��ײ�ްĕϐ�
    '=======================================================
    Private m_intIndexPreRow As Short '���ݑI������Ă���s
    Private m_intDefPlanNo As Short '�\���v����Ԃ̌v��ԍ�
    Private m_intSaveNo As Short '�ۑ��ԍ�
    Private m_StaffData() As StaffData_Type '�ΏېE�����

    Private m_intPlanStartDate As Integer '�J�n��
    Private m_intPlanEndDate As Integer '�I����

    Private m_ApplyEndFlg As Integer '�K�p�{�^���I���t���O�iTrue:�K�p�{�^���������j
    Private m_Sheet As FarPoint.Win.Spread.SheetView
    Private m_KinmuDataStCol As Integer
    Private m_KinmuDataEdCol As Integer
    Private m_StaffRowStRow As Integer
    Private m_StaffRowEdRow As Integer
    Private m_StaffMngIDCol As Integer
    Private m_DateLabelRow As Integer
    Private m_OuenStaffCnt As Integer
    Private m_MaxShowLine As Integer
    Private m_KinmuPlan As Integer
    Private m_PackageFLG As Integer
    Private m_ProgressForm As frmNSK0000HM

    Private Structure Save_Type
        Dim intSaveNo As Short '�ۑ��ԍ�
        Dim strBikou As String '���l
        Dim dblLastUpdTimeDate As Double '�ŏI�X�V����
        Dim strRegistID As String '����҂̃��[�U�[�h�c
        Sub init(ByVal p_SaveNo As Short)
            intSaveNo = p_SaveNo
            strBikou = ""
            dblLastUpdTimeDate = 0
            strRegistID = ""
        End Sub
    End Structure

    Private m_udtSaveYotei() As Save_Type '�ꎞ�ۑ��\��ꗗ

    '=======================================================
    '   Getter/Setter
    '=======================================================
    ''' <summary>�v����ԏ����󂯎��</summary>
    ''' <param name="p_FromYMD">�\���J�n��</param>
    ''' <param name="p_ToYMD">�\���I����</param>
    ''' <param name="Value">�v��ԍ�</param>
    ''' <remarks></remarks>
    Public WriteOnly Property pPlanInfo(ByVal p_FromYMD As String, ByVal p_ToYMD As String) As Short
        Set(ByVal Value As Short)
            m_intPlanStartDate = p_FromYMD '�J�n��
            m_intPlanEndDate = p_ToYMD '�I����
            m_intDefPlanNo = Value '�v��ԍ�
        End Set
    End Property

    ''' <summary>�E�������󂯎��</summary>
    ''' <param name="Value">�E�����</param>
    ''' <remarks></remarks>
    Public WriteOnly Property pStaffData() As Object
        Set(ByVal Value As Object)
            m_StaffData = Value
        End Set
    End Property

    Public WriteOnly Property pDataSorce() As Object
        Set(ByVal Value As Object)
            m_Sheet = Value
        End Set
    End Property

    ''' <summary>�K�p�{�^���I���t���O����ʉ�ʂɈ����n��</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property pTekiyouFlg() As Boolean
        Get
            pTekiyouFlg = m_ApplyEndFlg
        End Get
    End Property

    ''' <summary>�I�����ꂽ�ۑ��ԍ�����ʉ�ʂɈ����n��</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property pSelSaveNo() As String
        Get
            pSelSaveNo = m_intSaveNo
        End Get
    End Property
    Friend ReadOnly Property pProgressForm As frmNSK0000HM
        Get
            pProgressForm = m_ProgressForm
        End Get
    End Property

    Public WriteOnly Property pKinmuDataStCol As Integer
        Set(ByVal Value As Integer)
            m_KinmuDataStCol = Value
        End Set
    End Property
    Public WriteOnly Property pKinmuDataEdCol As Integer
        Set(ByVal Value As Integer)
            m_KinmuDataEdCol = Value
        End Set
    End Property
    Public WriteOnly Property pStaffRowStRow As Integer
        Set(ByVal Value As Integer)
            m_StaffRowStRow = Value
        End Set
    End Property
    Public WriteOnly Property pStaffRowEdRow As Integer
        Set(ByVal Value As Integer)
            m_StaffRowEdRow = Value
        End Set
    End Property
    Public WriteOnly Property pStaffMngIDCol As Integer
        Set(ByVal Value As Integer)
            m_StaffMngIDCol = Value
        End Set
    End Property
    Public WriteOnly Property pDateLabelRow As Integer
        Set(ByVal Value As Integer)
            m_DateLabelRow = Value
        End Set
    End Property
    Public WriteOnly Property pOuenStaffCnt As Integer
        Set(ByVal Value As Integer)
            m_OuenStaffCnt = Value
        End Set
    End Property
    Public WriteOnly Property pMaxShowLine As Integer
        Set(ByVal Value As Integer)
            m_MaxShowLine = Value
        End Set
    End Property
    Public WriteOnly Property pKinmuPlan As Integer
        Set(ByVal Value As Integer)
            m_KinmuPlan = Value
        End Set
    End Property
    Public WriteOnly Property pPackageFLG As Integer
        Set(ByVal Value As Integer)
            m_PackageFLG = Value
        End Set
    End Property

    ''' <summary>
    ''' frmNSK0000HE�t�H�[��Load�C�x���g
    ''' </summary>
    ''' <param name="eventSender">System.Object</param>
    ''' <param name="eventArgs">System.EventArgs</param>
    ''' <remarks>frmNSK0000HE��Load���A�\������</remarks>
    Private Sub frmNSK0000HE_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Const W_SUBNAME As String = "NSK0000HE Form_Load"
        Dim w_numWidth As Integer

        Try
            w_numWidth = sprSaveList_Sheet1.Columns(0).Width +
                                sprSaveList_Sheet1.Columns(1).Width +
                                sprSaveList_Sheet1.Columns(2).Width +
                                sprSaveList_Sheet1.Columns(3).Width

            '�X�v���b�h�̐ݒ�
            General.paSpreadSizeFit(sprSaveList,
                                    sprSaveList.VerticalScrollBarPolicy,
                                    sprSaveList.HorizontalScrollBarPolicy,
                                    M_SAVESPR_MAXROW,
                                    w_numWidth)

            '�ꎞ�ۑ��ꗗ�擾
            Call GetSaveData()

            '�ꎞ�ۑ��ꗗ�\��
            Call SetSprData()

            If m_intSaveNo = 0 Then
                '������
                m_intSaveNo = 1
                Call SetSelectData(0)
            Else
                '�ۑ���I������Ă���s�ڂ�I��
                Call SetSelectData(m_intIndexPreRow)
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    ''' <summary>
    ''' �ꎞ�ۑ��ꗗ�̎擾
    ''' �T�v:�Ζ��v���ʂ��J���Ă��镔���E�v����ԂŎ擾�B
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GetSaveData()
        Const W_SUBNAME As String = "NSK0000HE GetStoredData"

        Dim w_sbSql As New StringBuilder 'SQL��
        Dim w_objRs As ADODB.Recordset 'RecordSet ��޼ު��
        Dim w_objFields As ADODB.Fields '̨���� ��޼ު��
        Dim w_intDataCount As Short
        Dim w_intDataLoop As Short
        Dim w_intRowIdx As Short

        Try
            '�ꎞ�ۑ��\��ꗗ�̏�����
            ReDim m_udtSaveYotei(M_SAVESPR_MAXROW)
            For w_intRowIdx = 1 To M_SAVESPR_MAXROW
                m_udtSaveYotei(w_intRowIdx).init(w_intRowIdx)
            Next

            'Select�� �ҏW 
            With w_sbSql
                .AppendLine("SELECT")
                .AppendLine("  SAVENO")
                .AppendLine(", BIKOU")
                .AppendLine(", LASTUPDTIMEDATE")
                .AppendLine(", REGISTRANTID")
                .AppendLine("FROM")
                .AppendLine("  NS_TEMPPLANLIST_F")
                .AppendLine("WHERE")
                .AppendLine("    HOSPITALCD  = '" & General.g_strHospitalCD & "'")
                .AppendLine("AND PLANNO      =  " & m_intDefPlanNo)
                .AppendLine("AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'")
                .AppendLine("ORDER BY")
                .AppendLine("  SAVENO ASC")
            End With

            '�J�[�\���쐬
            w_objRs = General.paDBRecordSetOpen(w_sbSql.ToString())

            With w_objRs
                If .RecordCount <= 0 Then
                    '�ް������݂��Ȃ��Ƃ�
                Else
                    '�ް������݂���Ƃ�
                    .MoveLast()
                    '�ް������擾
                    w_intDataCount = .RecordCount
                    .MoveFirst()
                    '̨���޵�޼ު�Đ���
                    w_objFields = .Fields

                    '�ް�����Loop
                    For w_intDataLoop = 1 To w_intDataCount
                        '�ۑ��ԍ��ɑΉ�����s�Ɉꎞ�ۑ��̃f�[�^��ݒ肷��
                        w_intRowIdx = Short.Parse(General.paGetDbFieldVal(w_objFields("SAVENO"), 0))
                        '�ۑ��ԍ��擾
                        m_udtSaveYotei(w_intRowIdx).intSaveNo = General.paGetDbFieldVal(w_objFields("SAVENO"), 0)
                        '���l�擾
                        m_udtSaveYotei(w_intRowIdx).strBikou = General.paGetDbFieldVal(w_objFields("BIKOU"), "")
                        '�ŏI�X�V�����擾
                        m_udtSaveYotei(w_intRowIdx).dblLastUpdTimeDate = General.paGetDbFieldVal(w_objFields("LASTUPDTIMEDATE"), 0)
                        '����҂̃��[�U�[�h�c�擾
                        m_udtSaveYotei(w_intRowIdx).strRegistID = General.paGetDbFieldVal(w_objFields("REGISTRANTID"), "")

                        .MoveNext()
                    Next w_intDataLoop
                End If
            End With

            '���ݸ�����މ��
            w_sbSql.Clear()
            '��޼ު�Ẳ��
            w_objRs = Nothing
            w_objFields = Nothing
        Catch ex As Exception
            '���ݸ�����މ��
            w_sbSql.Clear()
            '��޼ު�Ẳ��
            w_objRs = Nothing
            w_objFields = Nothing

            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '�ꎞ�ۑ��ꗗ�̕\��
    Private Sub SetSprData()
        Const W_SUBNAME As String = "NSK0000HE SetSprData"

        Dim w_intRowIdx As Short
        Dim w_intSprLoop As Short
        Dim w_intSprRowCount As Short

        Try
            With sprSaveList_Sheet1
                '�X�v���b�h�̓��e���N���A����
                '�f�[�^�̃��Z�b�g
                sprSaveList_Sheet1.ClearRange(0, 0, sprSaveList_Sheet1.RowCount, sprSaveList_Sheet1.ColumnCount, False)

                '�X�^�C���̓K�p
                subSetStyles()

                '�X�v���b�h�̍ő�s������
                .RowCount = M_SAVESPR_MAXROW

                'spread�s��0����n�܂���ޯ��
                w_intSprRowCount = .RowCount - 1
                For w_intSprLoop = 0 To w_intSprRowCount
                    w_intRowIdx = w_intSprLoop + 1
                    '�ۑ��ԍ�
                    .Cells(w_intSprLoop, M_SAVESPR_COLIDX_NO).Text = EditData(m_udtSaveYotei(w_intRowIdx).intSaveNo, G_EDITMODE_NO)
                    '����҂̃��[�U�[�h�c
                    .Cells(w_intSprLoop, M_SAVESPR_COLIDX_NAME).Text = m_udtSaveYotei(w_intRowIdx).strRegistID
                    '�ŏI�X�V����
                    .Cells(w_intSprLoop, M_SAVESPR_COLIDX_DATE).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
                    .Cells(w_intSprLoop, M_SAVESPR_COLIDX_DATE).Text = EditData(m_udtSaveYotei(w_intRowIdx).dblLastUpdTimeDate, G_EDITMODE_DATETIME)
                    '���l
                    .Cells(w_intSprLoop, M_SAVESPR_COLIDX_BIKOU).Text = m_udtSaveYotei(w_intRowIdx).strBikou
                Next
            End With
        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '�X�^�C���̐ݒ���s���֐�
    Private Sub subSetStyles()
        Const W_SUBNAME As String = "NSK0000HE subSetStyles"

        Dim w_style As New FarPoint.Win.Spread.StyleInfo()
        Dim w_Font As New System.Drawing.Font("�l�r �S�V�b�N", 10.0!)
        Dim w_TextCellType As New FarPoint.Win.Spread.CellType.TextCellType

        Try
            '�X�^�C���̓K�p
            w_style.Font = w_Font
            w_style.CellType = w_TextCellType
            w_style.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Left
            w_style.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
            w_style.BackColor = Color.White

            '�X�v���b�h������
            sprSaveList_Sheet1.Models.Style.SetDirectInfo(-1, -1, w_style)

        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    ''' <summary>
    ''' Spread�Z���������̏���
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="EventArgs"></param>
    ''' <remarks></remarks>
    Private Sub sprSaveList_CellClick(ByVal sender As Object, ByVal EventArgs As FarPoint.Win.Spread.CellClickEventArgs) Handles sprSaveList.CellClick

        Try
            '�w�b�_�N���b�N���͏����𔲂���
            If EventArgs.ColumnHeader Then
                Exit Sub
            End If

            '�I���s�̔w�i�F�ύX
            Call SetSelectData(EventArgs.Row)

            '�ۑ��ԍ�
            m_intSaveNo = m_intIndexPreRow + 1 'CellClickEventArgs.row��0����n�܂���ޯ��
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' �I���s�̔w�i�F�ύX
    ''' </summary>
    ''' <param name="p_Row">�I���s</param>
    ''' <remarks></remarks>
    Private Sub SetSelectData(ByVal p_Row As Integer)
        Const W_SUBNAME As String = "NSK0000HE SetSelectData"

        Try
            If p_Row >= 0 Then
                With sprSaveList_Sheet1
                    '�X�v���b�h�̕\���̍X�V�͈ꊇ��
                    '�S�̂̕\���̏�����
                    .Rows(0, .RowCount - 1).BackColor = Color.White
                    .Rows(0, .RowCount - 1).ForeColor = Color.Black

                    If p_Row > -1 Then
                        '�w�肳�ꂽ�s�̔w�i��ύX
                        .Rows(p_Row).BackColor = Color.Cyan
                        .Rows(p_Row).ForeColor = Color.Black
                    End If

                    '�ꎞ�ۑ��ꗗspread�ɕۑ��҂����݂��邩�ǂ����̃`�F�b�N(�ۑ����ɖ��̂��K�v�ł�)
                    '�I�������s�Ƀf�[�^������ꍇ
                    If Not .Cells(p_Row, M_SAVESPR_COLIDX_NAME).Text = "" Then
                        '���l�̓��e���e�L�X�g�{�b�N�X�u���l�v�ɏo��
                        '�K�p�{�^����������
                        txtBikou.Text = .Cells(p_Row, M_SAVESPR_COLIDX_BIKOU).Text
                        cmdApply.Enabled = True
                    Else '�I�������s�Ƀf�[�^���Ȃ��ꍇ
                        '�e�L�X�g�{�b�N�X�u���l�v���N���A
                        '�K�p�{�^����񊈐���
                        txtBikou.Text = ""
                        cmdApply.Enabled = False
                    End If
                End With

                '����I�����ꂽ�s���m��
                m_intIndexPreRow = p_Row
            End If
        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    ''' <summary>
    ''' cmdSave�{�^��Click�C�x���g
    ''' </summary>
    ''' <param name="eventSender">System.Object</param>
    ''' <param name="eventArgs">System.EventArgs</param>
    ''' <remarks>�ꎞ�ۑ��f�[�^�̍폜�E�ۑ��E�ēǍ�</remarks>
    Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc '��ۼ��ެ���̈ꎞ�Ҕ�
        General.g_ErrorProc = "NSK0000HA cmdSave_Click"

        Try
            '��ݻ޸��݊J�n
            Call General.paBeginTrans()

            '�ް��폜
            Call DeleteSaveData()

            '�ް�������
            If InsertSaveData() Then
                '�Я�
                Call General.paCommit()

                '�ꎞ�ۑ��f�[�^�̍ēǍ�
                Call frmNSK0000HE_Load(eventSender, eventArgs)
            Else
                Call General.paRollBack()
            End If

            General.g_ErrorProc = w_PreErrorProc '�Ҕ���ۼ��ެ�������ɖ߂�
        Catch ex As Exception
            Call General.paRollBack()
            Call General.paTrpMsg(Convert.ToString(Err.Number), General.g_ErrorProc)
            End
        End Try
    End Sub

    ''' <summary>
    ''' �ꎞ�ۑ��f�[�^�̍폜
    ''' </summary>
    ''' <remarks>
    ''' �ȉ��̃e�[�u������f�[�^���폜����B
    '''    �E�ꎞ�ۑ��\��ꗗ�e
    '''    �E�ꎞ�ۑ��Ζ��\��e
    '''    �E�ꎞ�ۑ��Ζ��ڍׂe
    '''    �E�ꎞ�ۑ��N�x�e
    ''' </remarks>
    Private Sub DeleteSaveData()
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc '��ۼ��ެ���̈ꎞ�Ҕ�
        General.g_ErrorProc = "NSK0000HE DeleteSaveData"

        Dim w_sbSql As New StringBuilder 'SQL��
        '�e�[�u�����̃��X�g
        Dim w_arrTempSaveTables() As String = {"NS_TEMPPLANLIST_F", "NS_TEMPKINMUPLAN_F",
                                                "NS_TEMPKINMUDETAIL_F", "NS_TEMPNENKYU_F"}
        Try
            '���ׂẴe�[�u����񋓂���
            With w_sbSql
                For Each table As String In w_arrTempSaveTables
                    .AppendLine("DELETE FROM " & table)
                    .AppendLine("WHERE")
                    .AppendLine("    HOSPITALCD   = '" & General.g_strHospitalCD & "'")
                    .AppendLine("AND PLANNO      >=  " & m_intDefPlanNo)
                    .AppendLine("AND KINMUDEPTCD <= '" & General.g_strSelKinmuDeptCD & "'")
                    .AppendLine("AND SAVENO       =  " & m_intSaveNo)
                    '�X�V���s
                    Call General.paDBExecute(.ToString)
                    Call .Clear()
                Next table
            End With

            General.g_ErrorProc = w_PreErrorProc '�Ҕ���ۼ��ެ�������ɖ߂�
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' �ꎞ�ۑ��f�[�^�̕ۑ�
    ''' </summary>
    ''' <remarks>
    ''' �ȉ��̃e�[�u���Ƀf�[�^��o�^����B
    '''    �E�ꎞ�ۑ��\��ꗗ�e
    '''    �E�ꎞ�ۑ��Ζ��\��e
    '''    �E�ꎞ�ۑ��Ζ��ڍׂe
    '''    �E�ꎞ�ۑ��N�x�e
    '''    �E�ꎞ�ۑ���x�e
    ''' </remarks>
    Private Function InsertSaveData() As Boolean
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc '��ۼ��ެ���̈ꎞ�Ҕ�
        General.g_ErrorProc = "NSK0000HE InsertSaveData"

        Dim w_sbSql As New StringBuilder        'SQL��
        Dim w_strSysDate As String              '�X�V����
        Dim w_StaffMngID As String = String.Empty
        Dim w_strDate As String = String.Empty
        Dim w_KinmuCD As String = String.Empty  'KinmuCD
        Dim w_RiyuKBN As String = String.Empty  '���R�敪
        Dim w_KangoCD As String = String.Empty  '�����Ō�P��CD
        Dim w_Time As String = String.Empty     '���Ԑ�
        Dim w_Comment As String = String.Empty
        Dim w_Nenkyu() As NenkyuDetail_Type
        Dim w_strMsg() As String
        Try
            '�o�^������t���擾
            w_strSysDate = Format(Now, "yyyyMMddHHmmss")

            '*********************************************************************************************************'
            '                                    �ꎞ�ۑ��\��ꗗ�e�̍X�V     
            '*********************************************************************************************************'
            'Insert�� �ҏW 
            With w_sbSql
                .AppendLine("INSERT INTO NS_TEMPPLANLIST_F (")
                .AppendLine("  HOSPITALCD")
                .AppendLine(", PLANNO")
                .AppendLine(", KINMUDEPTCD")
                .AppendLine(", SAVENO")
                .AppendLine(", BIKOU")
                .AppendLine(", REGISTFIRSTTIMEDATE")
                .AppendLine(", LASTUPDTIMEDATE")
                .AppendLine(", REGISTRANTID")
                .AppendLine(") VALUES (")
                .AppendLine(" '" & General.g_strHospitalCD & "'")
                .AppendLine(", " & m_intDefPlanNo)
                .AppendLine(",'" & General.g_strSelKinmuDeptCD & "'")
                .AppendLine(", " & m_intSaveNo)
                .AppendLine(",'" & txtBikou.Text & "'")
                .AppendLine(", " & w_strSysDate)
                .AppendLine(", " & w_strSysDate)
                .AppendLine(",'" & General.g_strUserID & "')")
            End With

            '�X�V���s
            Call General.paDBExecute(w_sbSql.ToString())

            '���ݸ�����މ��
            w_sbSql.Clear()

            For i As Integer = 1 To UBound(m_StaffData)
                If General.g_lngDaikyuMng = 0 Then
                    '*********************************************************************************************************'
                    '                                    �ꎞ�ۑ���x�e�̍X�V     
                    '*********************************************************************************************************'
                    For j As Integer = 1 To UBound(m_StaffData(i).Daikyu)
                        For k As Integer = 1 To UBound(m_StaffData(i).Daikyu(j).DaikyuDetail)
                            If m_intPlanStartDate <= m_StaffData(i).Daikyu(j).DaikyuDetail(k).DaikyuDate AndAlso
                               m_StaffData(i).Daikyu(j).DaikyuDetail(k).DaikyuDate <= m_intPlanEndDate Then
                                '&1��&2��&3���Ă���ꍇ��&4�ł��܂���B~n&2��&5���Ă��������B
                                ReDim w_strMsg(5)
                                w_strMsg(1) = "�E��"
                                w_strMsg(2) = "��x"
                                w_strMsg(3) = "�擾"
                                w_strMsg(4) = "�ꎞ�ۑ�"
                                w_strMsg(5) = "�폜"
                                Call General.paMsgDsp("NS0412", w_strMsg)
                                General.g_ErrorProc = w_PreErrorProc '�Ҕ���ۼ��ެ�������ɖ߂�
                                Return False
                            End If
                        Next k
                        If m_intPlanStartDate <= m_StaffData(i).Daikyu(j).HolDate AndAlso m_StaffData(i).Daikyu(j).HolDate <= m_intPlanEndDate Then
                            '���ݸ�����މ��
                            w_sbSql.Clear()
                            'Insert�� �ҏW 
                            With w_sbSql
                                .AppendLine("INSERT INTO NS_TEMPDAIKYUMNG_F (")
                                .AppendLine("  HOSPITALCD")
                                .AppendLine(", PLANNO")
                                .AppendLine(", KINMUDEPTCD")
                                .AppendLine(", SAVENO")
                                .AppendLine(", STAFFMNGID")
                                .AppendLine(", GETKBN")
                                .AppendLine(", WORKHOLKINMUDATE")
                                .AppendLine(", WORKHOLKINMUCD")
                                .AppendLine(", TODOKEDENO")
                                .AppendLine(", REGISTFIRSTTIMEDATE")
                                .AppendLine(", LASTUPDTIMEDATE")
                                .AppendLine(", REGISTRANTID")
                                .AppendLine(") VALUES (")
                                .AppendLine(" '" & General.g_strHospitalCD & "'")
                                .AppendLine(", " & m_intDefPlanNo)
                                .AppendLine(",'" & General.g_strSelKinmuDeptCD & "'")
                                .AppendLine(", " & m_intSaveNo)
                                .AppendLine(",'" & m_StaffData(i).ID & "'")
                                .AppendLine(", " & m_StaffData(i).Daikyu(j).GetKbn)
                                .AppendLine(", " & m_StaffData(i).Daikyu(j).HolDate)
                                .AppendLine(",'" & m_StaffData(i).Daikyu(j).HolKinmuCD & "'")
                                .AppendLine(", 0")
                                .AppendLine(", " & w_strSysDate)
                                .AppendLine(", " & w_strSysDate)
                                .AppendLine(",'" & General.g_strUserID & "')")
                            End With
                            '�X�V���s
                            Call General.paDBExecute(w_sbSql.ToString)
                            '���ݸ�����މ��
                            w_sbSql.Clear()
                        End If
                    Next j
                End If
                '*********************************************************************************************************'
                '                                    �ꎞ�ۑ��Ζ��ڍׂe�̍X�V     
                '*********************************************************************************************************'
                For j As Integer = 1 To UBound(m_StaffData(i).Kojyo)
                    For k As Integer = 1 To UBound(m_StaffData(i).Kojyo(j).lngKinmuDetailTime)
                        '���ݸ�����މ��
                        w_sbSql.Clear()
                        'Insert�� �ҏW 
                        With w_sbSql
                            .AppendLine("INSERT INTO NS_TEMPKINMUDETAIL_F (")
                            .AppendLine("  HOSPITALCD")
                            .AppendLine(", PLANNO")
                            .AppendLine(", KINMUDEPTCD")
                            .AppendLine(", SAVENO")
                            .AppendLine(", STAFFMNGID")
                            .AppendLine(", DATEF")
                            .AppendLine(", SEQ")
                            .AppendLine(", KINMUDETAILDATE")
                            .AppendLine(", FROMTIME")
                            .AppendLine(", KINMUDETAILCD")
                            .AppendLine(", OUENKINMUDEPTCD") '2018/02/23 Yamanishi Add
                            .AppendLine(", TOTIME")
                            .AppendLine(", NEXTDAYFLG")
                            .AppendLine(", KINMUDETAILTIME")
                            .AppendLine(", HOLSUBFLG")
                            .AppendLine(", REZEPTCALCKBN")
                            .AppendLine(", DAYTIME")
                            .AppendLine(", NIGHTTIME")
                            .AppendLine(", NEXTNIGHTTIME")
                            .AppendLine(", UNIQUESEQNO")
                            .AppendLine(", REGISTFIRSTTIMEDATE")
                            .AppendLine(", LASTUPDTIMEDATE")
                            .AppendLine(", REGISTRANTID")
                            .AppendLine(") VALUES (")
                            .AppendLine(" '" & General.g_strHospitalCD & "'")
                            .AppendLine(", " & m_intDefPlanNo)
                            .AppendLine(",'" & General.g_strSelKinmuDeptCD & "'")
                            .AppendLine(", " & m_intSaveNo)
                            .AppendLine(",'" & m_StaffData(i).ID & "'")
                            .AppendLine(", " & m_StaffData(i).Kojyo(j).lngDate)
                            .AppendLine(", " & m_StaffData(i).Kojyo(j).Seq(k))
                            .AppendLine(", " & m_StaffData(i).Kojyo(j).lngKinmuDetailDate(k))
                            .AppendLine(", " & m_StaffData(i).Kojyo(j).lngTimeFrom(k))
                            .AppendLine(",'" & m_StaffData(i).Kojyo(j).strKinmuDetailCD(k) & "'")
                            .AppendLine(",'" & m_StaffData(i).Kojyo(j).OuenKinmuDeptCD(k) & "'") '2018/02/23 Yamanishi Add
                            .AppendLine(", " & m_StaffData(i).Kojyo(j).lngTimeTo(k))
                            .AppendLine(",'" & m_StaffData(i).Kojyo(j).strNextFlg(k) & "'")
                            .AppendLine(", " & m_StaffData(i).Kojyo(j).lngKinmuDetailTime(k))
                            .AppendLine(",'" & m_StaffData(i).Kojyo(j).strHolSubFlg(k) & "'")
                            .AppendLine(",'" & m_StaffData(i).Kojyo(j).strShinryoKbn(k) & "'")
                            .AppendLine(", " & m_StaffData(i).Kojyo(j).lngNikkinTime(k))
                            .AppendLine(", " & m_StaffData(i).Kojyo(j).lngYakinTime(k))
                            .AppendLine(", " & m_StaffData(i).Kojyo(j).lngYokuYakinTime(k))
                            .AppendLine(",'" & m_StaffData(i).Kojyo(j).UniqueseqNo(k) & "'")
                            .AppendLine(", " & w_strSysDate)
                            .AppendLine(", " & w_strSysDate)
                            .AppendLine(",'" & General.g_strUserID & "')")
                        End With
                        '�X�V���s
                        Call General.paDBExecute(w_sbSql.ToString)
                        '���ݸ�����މ��
                        w_sbSql.Clear()
                    Next k
                Next j
            Next i

            For w_Row As Integer = m_StaffRowStRow To m_StaffRowEdRow - (m_OuenStaffCnt * m_MaxShowLine)
                If IsDataRowAndGetMngID(w_Row, w_StaffMngID) Then
                    For w_Col As Integer = m_KinmuDataStCol To m_KinmuDataEdCol
                        If IsDataColAndGetKinmuData(w_Row, w_Col, w_strDate, w_KinmuCD, w_RiyuKBN, w_KangoCD, w_Time, w_Comment) Then
                            '*********************************************************************************************************'
                            '                                    �ꎞ�ۑ��Ζ��\��e�̍X�V     
                            '*********************************************************************************************************'
                            '���ݸ�����މ��
                            Call w_sbSql.Clear()
                            With w_sbSql
                                .AppendLine("INSERT INTO NS_TEMPKINMUPLAN_F (")
                                .AppendLine("  HOSPITALCD")
                                .AppendLine(", PLANNO")
                                .AppendLine(", KINMUDEPTCD")
                                .AppendLine(", SAVENO")
                                .AppendLine(", DATEF")
                                .AppendLine(", STAFFMNGID")
                                .AppendLine(", KINMUCD")
                                .AppendLine(", REASONKBN")
                                .AppendLine(", OUENKINMUDEPTCD")
                                .AppendLine(", HOPECOMMENT")
                                .AppendLine(", REGISTFIRSTTIMEDATE")
                                .AppendLine(", LASTUPDTIMEDATE")
                                .AppendLine(", REGISTRANTID")
                                .AppendLine(") VALUES (")
                                .AppendLine(" '" & General.g_strHospitalCD & "'")
                                .AppendLine(", " & m_intDefPlanNo)
                                .AppendLine(",'" & General.g_strSelKinmuDeptCD & "'")
                                .AppendLine(", " & m_intSaveNo)
                                .AppendLine(", " & w_strDate)
                                .AppendLine(",'" & w_StaffMngID & "'")
                                .AppendLine(",'" & w_KinmuCD & "'")
                                .AppendLine(",'" & w_RiyuKBN & "'")
                                .AppendLine(",'" & w_KangoCD & "'")
                                .AppendLine(",'" & w_Comment & "'")
                                .AppendLine(", " & w_strSysDate)
                                .AppendLine(", " & w_strSysDate)
                                .AppendLine(",'" & General.g_strUserID & "')")
                            End With

                            '�X�V���s
                            Call General.paDBExecute(w_sbSql.ToString)
                            '���ݸ�����މ��
                            Call w_sbSql.Clear()

                            '*********************************************************************************************************'
                            '                                    �ꎞ�ۑ��N�x�e�̍X�V     
                            '*********************************************************************************************************'
                            ReDim w_Nenkyu(0)
                            If ExistsNenkyuAndGetNenkyuData(w_KinmuCD, w_Time, w_Nenkyu) Then
                                For i As Integer = 1 To UBound(w_Nenkyu)
                                    '���ݸ�����މ��
                                    w_sbSql.Clear()
                                    'Insert�� �ҏW 
                                    With w_sbSql
                                        .AppendLine("INSERT INTO NS_TEMPNENKYU_F (")
                                        .AppendLine("  HOSPITALCD")
                                        .AppendLine(", PLANNO")
                                        .AppendLine(", KINMUDEPTCD")
                                        .AppendLine(", SAVENO")
                                        .AppendLine(", STAFFMNGID")
                                        .AppendLine(", DATEF")
                                        .AppendLine(", SEQ")
                                        .AppendLine(", GETCONTENTSKBN")
                                        .AppendLine(", HOLIDAYBUNRUICD")
                                        .AppendLine(", FROMTIME")
                                        .AppendLine(", TOTIME")
                                        .AppendLine(", NEXTDAYFLG")
                                        .AppendLine(", NENKYUTIME")
                                        .AppendLine(", HOLSUBFLG")
                                        .AppendLine(", DAYTIME")
                                        .AppendLine(", NIGHTTIME")
                                        .AppendLine(", NEXTNIGHTTIME")
                                        .AppendLine(", KINMUDATE")
                                        .AppendLine(", DATEKBN")
                                        .AppendLine(", UNIQUESEQNO")
                                        .AppendLine(", APPROVEFLG")
                                        .AppendLine(", DELFLG")
                                        .AppendLine(", REGISTFIRSTTIMEDATE")
                                        .AppendLine(", LASTUPDTIMEDATE")
                                        .AppendLine(", REGISTRANTID")
                                        .AppendLine(") VALUES (")
                                        .AppendLine(" '" & General.g_strHospitalCD & "'")       '�{��CD
                                        .AppendLine(", " & m_intDefPlanNo)                      '�\���v����Ԃ̌v��ԍ�
                                        .AppendLine(",'" & General.g_strSelKinmuDeptCD & "'")   '�I���Ζ�����CD
                                        .AppendLine(", " & m_intSaveNo)                         '�ۑ��ԍ�
                                        .AppendLine(",'" & w_StaffMngID & "'")                  '�E���Ǘ��ԍ�
                                        .AppendLine(", " & w_strDate)                           '���t
                                        .AppendLine(", " & i)                                   'SEQ
                                        .AppendLine(",'" & w_Nenkyu(i).GetContentsKbn & "'")    '�擾���e�敪
                                        .AppendLine(",'" & w_Nenkyu(i).HolidayBunruiCD & "'")   '�x�ݕ���CD
                                        .AppendLine(", " & w_Nenkyu(i).FromTime)                '�J�n����
                                        .AppendLine(", " & w_Nenkyu(i).ToTime)                  '�I������
                                        .AppendLine(",'" & w_Nenkyu(i).DateKbn & "'")           '����FLG
                                        .AppendLine(", " & w_Nenkyu(i).NenkyuTime)              '���ԔN�x
                                        .AppendLine(",'" & w_Nenkyu(i).HolSubFlg & "'")         '�x�e���Z�t���O
                                        .AppendLine(", " & w_Nenkyu(i).DayTime)                 '���Ύ���
                                        .AppendLine(", " & w_Nenkyu(i).NightTime)               '��Ύ���
                                        .AppendLine(", " & w_Nenkyu(i).NextNightTime)           '������Ύ���
                                        .AppendLine(", " & w_strDate)                           '�Ζ��N����
                                        .AppendLine(",'" & w_Nenkyu(i).DateKbn & "'")           '�N�����敪
                                        .AppendLine(",''")                                      '�N�xUNIQUESEQNO
                                        .AppendLine(",'1'")                                     '���F�ς�FLG
                                        .AppendLine(",''")                                      '�폜FLG
                                        .AppendLine(", " & w_strSysDate)                        '����o�^����
                                        .AppendLine(", " & w_strSysDate)                        '�ŏI�X�V����
                                        .AppendLine(",'" & General.g_strUserID & "')")          '�o�^��ID
                                    End With
                                    '�X�V���s
                                    Call General.paDBExecute(w_sbSql.ToString)
                                    '���ݸ�����މ��
                                    Call w_sbSql.Clear()
                                Next i
                            End If
                        End If
                    Next w_Col
                End If
            Next w_Row

            General.g_ErrorProc = w_PreErrorProc '�Ҕ���ۼ��ެ�������ɖ߂�
            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function IsDataRowAndGetMngID(ByVal p_Row As Integer, ByRef p_StaffMngID As String) As Boolean
        If (p_Row - m_StaffRowStRow) Mod m_MaxShowLine = m_KinmuPlan Then
            p_StaffMngID = m_Sheet.Cells(p_Row, m_StaffMngIDCol).Text
            Return IsNumeric(p_StaffMngID)
        End If
        Return False
    End Function
    Private Function IsDataColAndGetKinmuData(ByVal p_Row As Integer, ByVal p_Col As Integer,
                                              ByRef p_strDate As String, ByRef p_KinmuCD As String, ByRef p_RiyuKBN As String,
                                              ByRef p_KangoCD As String, ByRef p_Time As String, ByRef p_Comment As String) As Boolean
        Dim w_Var As String
        Dim w_iDate As Integer
        p_strDate = m_Sheet.Cells(m_DateLabelRow, p_Col).Text
        If Integer.TryParse(p_strDate, w_iDate) AndAlso IsDate(w_iDate.ToString("0000/00/00")) Then
            w_Var = m_Sheet.Cells(p_Row, p_Col).Text
            If Not String.IsNullOrEmpty(Trim(w_Var)) Then
                Call Get_KinmuMark(w_Var, p_KinmuCD, p_RiyuKBN, "", p_KangoCD, p_Time, p_Comment)
                Return IsNumeric(Trim(p_KinmuCD))
            End If
        End If
        Return False
    End Function
    Private Function ExistsNenkyuAndGetNenkyuData(ByVal p_KinmuCD As String, ByVal p_Time As String,
                                                  ByRef p_Nenkyu() As NenkyuDetail_Type) As Boolean
        Dim w_GetContentsKBN As String = String.Empty
        Dim w_HolCD As String = String.Empty
        If m_PackageFLG = 0 OrElse m_PackageFLG = 1 Then
            Call GetNenkyuContentsKbnAndHolCD(p_KinmuCD, w_GetContentsKBN, w_HolCD)
            If p_Time <> "" OrElse w_GetContentsKBN <> "" Then
                Call Get_NenkyuDetail(p_Time, w_GetContentsKBN, p_Nenkyu, w_HolCD)
                Return True
            End If
        End If
        Return False
    End Function

    ''' <summary>
    ''' cmdApply�{�^��Click�C�x���g
    ''' </summary>
    ''' <param name="eventSender">System.Object</param>
    ''' <param name="eventArgs">System.EventArgs</param>
    ''' <remarks></remarks>
    Private Sub cmdApply_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdApply.Click
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc '��ۼ��ެ���̈ꎞ�Ҕ�
        General.g_ErrorProc = "NSK0000HE cmdApply_Click"
        Try
            If General.paMsgDsp("NS0097", New String() {"", "�ҏW���̋Ζ�", "�j��"}) = MsgBoxResult.Yes Then

                m_ProgressForm = New frmNSK0000HM
                m_ProgressForm.pNumberDisp = False
                Call m_ProgressForm.Show(pProcessObj)
                m_ProgressForm.pForeColor = ColorTranslator.ToOle(Color.Black)
                m_ProgressForm.pSyoriText = "�I��������..."
                m_ProgressForm.pMaxValue = 3
                m_ProgressForm.pCountValue = 0

                '��޼ު�Ẳ��
                Erase m_StaffData
                Erase m_udtSaveYotei

                '�K�p
                m_ApplyEndFlg = True

                '̫�� ��۰��
                Me.Close()
            End If
            General.g_ErrorProc = w_PreErrorProc '�Ҕ���ۼ��ެ�������ɖ߂�
        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), General.g_ErrorProc)
            End
        End Try
    End Sub

    ''' <summary>
    ''' cmdClose�{�^��Click�C�x���g
    ''' </summary>
    ''' <param name="eventSender">System.Object</param>
    ''' <param name="eventArgs">System.EventArgs</param>
    ''' <remarks>��ʂ����</remarks>
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        Const W_SUBNAME As String = "NSK0000HE cmdClose_Click"
        Try
            '��޼ު�Ẳ��
            Erase m_StaffData
            Erase m_udtSaveYotei

            '����
            m_ApplyEndFlg = False

            '�������Ȃ��ŕ���
            Me.Close()
        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub
End Class