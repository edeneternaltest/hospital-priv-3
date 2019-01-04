Option Strict Off
Option Explicit On
Friend Class frmNSK0000HN
    Inherits General.FormBase
	
	Private Const M_ErrChkItem_Count As Short = 0 '�Ζ��^�x�������ٲ��ޯ��
	Private Const M_ErrChkItem_Interval As Short = 1 '�Ζ��^�x���Ԋu���ٲ��ޯ��
	Private Const M_ErrChkItem_Pattern As Short = 2 '�֎~�Ζ��p�^�[�����ٲ��ޯ��
	Private Const M_ErrChkItem_NotKinmu As Short = 3 '�֎~�Ζ����ٲ��ޯ��

	Private m_ViewFrom As Integer '�\�����ԊJ�n��
	Private m_ViewTo As Integer '�\�����ԏI����
	Private m_4WeekFrom As Object '4�T/1�����̏ꍇ�ł�4�T���Ԃ̊J�n���i�I�����͌v�Z�����߂�j	
	Private m_ChkCount As Boolean '�������H�iTrue:�����CFalse:�������j
	Private m_ChkInterval As Boolean '�Ԋu�����H�iTrue:�����CFalse:�������j
	Private m_ChkPattern As Boolean '�֎~����������H�iTrue:�����CFalse:�������j
	Private m_ChkNotKinmu As Boolean '�֎~�Ζ������H�iTrue:�����CFalse:�������j
    Private m_ChkStaffPattern As Boolean '�֎~�E������������H�iTrue:�����CFalse:�������j
    Private m_ChkGiryoType As Boolean '
    Private m_RenzokuKinmuCheck As Boolean '�A���Ζ�
    Private m_AbsoluteKinmuCheck As Boolean '�K�{�Ζ�
	
	Private Structure OutputType
        Dim Date_Renamed As Integer
		Dim StaffName As String
		Dim ErrorDetail As String
        Dim ErrorName As String
		Dim StaffIdx As Short '�Ώۂ̍s�C���f�b�N�X
		Dim ColIdx As Short '�Ώۂ̗�C���f�b�N�X
	End Structure
	
	Private m_ErrorList() As OutputType
	Private m_intSelIdx As Short
	
    Public WriteOnly Property pCountCheck() As Boolean
        Set(ByVal Value As Boolean)
            m_ChkCount = Value
        End Set
    End Property

	Public WriteOnly Property pIntervalCheck() As Boolean
		Set(ByVal Value As Boolean)
			m_ChkInterval = Value
		End Set
    End Property

	Public WriteOnly Property pNotKinmuCheck() As Boolean
		Set(ByVal Value As Boolean)
			m_ChkNotKinmu = Value
		End Set
    End Property

	Public WriteOnly Property pPatternCheck() As Boolean
		Set(ByVal Value As Boolean)
			m_ChkPattern = Value
		End Set
    End Property

    Public WriteOnly Property pStaffPatternCheck() As Boolean
        Set(ByVal Value As Boolean)
            m_ChkStaffPattern = Value
        End Set
    End Property

    Public WriteOnly Property pGiryoType() As Boolean
        Set(ByVal Value As Boolean)
            m_ChkGiryoType = Value
        End Set
    End Property

    Public WriteOnly Property pRenzokuKinmuCheck() As Boolean
        Set(ByVal Value As Boolean)
            m_RenzokuKinmuCheck = Value
        End Set
    End Property

    Public WriteOnly Property pAbsoluteKinmuCheck() As Boolean
        Set(ByVal Value As Boolean)
            m_AbsoluteKinmuCheck = Value
        End Set
    End Property

	Public ReadOnly Property pStaffIdxGet() As Short
		Get
			If m_intSelIdx <= UBound(m_ErrorList) Then
				pStaffIdxGet = m_ErrorList(m_intSelIdx).StaffIdx
			End If
		End Get
    End Property

	Public ReadOnly Property pDateIdxGet() As Short
		Get
			If m_intSelIdx <= UBound(m_ErrorList) Then
				pDateIdxGet = m_ErrorList(m_intSelIdx).ColIdx
			End If
		End Get
    End Property

	Public WriteOnly Property pViewFrom() As Integer
		Set(ByVal Value As Integer)
			m_ViewFrom = Value
		End Set
    End Property

	Public WriteOnly Property pViewTo() As Integer
		Set(ByVal Value As Integer)
			m_ViewTo = Value
		End Set
	End Property
	
	Private Sub Set_CountErrInf()
		On Error GoTo Set_CountErrInf
		Const W_SUBNAME As String = "NSK0000HN Set_CountErrInf"
		
		Dim w_Loop As Short
        Dim w_DataIndex As Short
		Dim w_CmbIndex As Short
		
        For w_Loop = 1 To UBound(g_KikanError2)
            For w_CmbIndex = 1 To UBound(g_KikanError2(w_Loop).CheckSpan)
                If g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).ErrorFlg = True And g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).ErrorDate >= m_ViewFrom And g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).ErrorDate <= m_ViewTo Then

                    '�G���[���\��
                    w_DataIndex = UBound(m_ErrorList) + 1
                    ReDim Preserve m_ErrorList(w_DataIndex)

                    '�G���[����
                    m_ErrorList(w_DataIndex).Date_Renamed = g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).ErrorDate

                    '�E���C���f�b�N�X
                    m_ErrorList(w_DataIndex).StaffIdx = g_KikanError2(w_Loop).StaffIdx

                    '����
                    m_ErrorList(w_DataIndex).StaffName = g_KikanError2(w_Loop).StaffName

                    '���t�C���f�b�N�X
                    m_ErrorList(w_DataIndex).ColIdx = g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).ColIdx

                    '�G���[���e
                    m_ErrorList(w_DataIndex).ErrorDetail = g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).KinmuName & Space(2) & String.Format("{0, 3}", g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).KinmuCount) & "��"

                    '�G���[����
                    m_ErrorList(w_DataIndex).ErrorName = "�Ζ��^�x����"
                End If
            Next w_CmbIndex
        Next w_Loop
		
		Exit Sub
Set_CountErrInf: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
		End
	End Sub

    '2015/05/13 Ishiga add start-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Set_RenzokuCountErrInf()
        On Error GoTo Set_RenzokuCountErrInf
        Const W_SUBNAME As String = "NSK0000HN Set_RenzokuCountErrInf"

        Dim w_Loop As Short
        Dim w_DataIndex As Short

        For w_Loop = 1 To UBound(g_RenzokuError2)
            If g_RenzokuError2(w_Loop).CheckSpan(0).ErrorFlg = True And g_RenzokuError2(w_Loop).CheckSpan(0).ErrorDate >= m_ViewFrom And g_RenzokuError2(w_Loop).CheckSpan(0).ErrorDate <= m_ViewTo Then

                '�G���[���\��
                w_DataIndex = UBound(m_ErrorList) + 1
                ReDim Preserve m_ErrorList(w_DataIndex)

                '�G���[����
                m_ErrorList(w_DataIndex).Date_Renamed = g_RenzokuError2(w_Loop).CheckSpan(0).ErrorDate

                '�E���C���f�b�N�X
                m_ErrorList(w_DataIndex).StaffIdx = g_RenzokuError2(w_Loop).StaffIdx

                '����
                m_ErrorList(w_DataIndex).StaffName = g_RenzokuError2(w_Loop).StaffName

                '���t�C���f�b�N�X
                m_ErrorList(w_DataIndex).ColIdx = g_RenzokuError2(w_Loop).CheckSpan(0).ColIdx

                '�G���[���e
                m_ErrorList(w_DataIndex).ErrorDetail = g_RenzokuError2(w_Loop).CheckSpan(0).KinmuName & Space(2) & String.Format("{0, 3}", g_RenzokuError2(w_Loop).CheckSpan(0).KinmuCount) & "��"

                '�G���[����
                m_ErrorList(w_DataIndex).ErrorName = "�Ζ��^�x���A��"
            End If
        Next w_Loop

        Exit Sub
Set_RenzokuCountErrInf:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub
    '2015/05/13 Ishiga add end---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

	Private Sub Set_ErrorInf(ByVal pIndex As Short)
		On Error GoTo Set_ErrorInf
		Const W_SUBNAME As String = "NSK0000HN Set_ErrorInf"
		
        Dim w_Int As Short
		Dim w_intCnt As Short
		Dim w_str As String
		Dim w_i As Short
        Dim w_LItem As ListViewItem
		
		lvwErrorList.Items.Clear()
		
        If m_ChkCount = True Then
            '�񐔑I��

            '�e�Ζ����i�W�v���ځj�̃G���[���\��
            Call Set_CountErrInf()
        End If

		If m_ChkInterval = True Then
            '�Ԋu�I��
			
			'�e�Ζ����i�W�v���ځj�̃G���[���\��
			Call Set_IntervalErrInf()
        End If

		If m_ChkPattern = True Then
            '�֎~�p�^�[��
			
			'�G���[�\���󋵕\��
			For w_Int = 1 To UBound(g_NotPatternError2)
                For w_i = 1 To UBound(g_NotPatternError2(w_Int).Data)
                    If g_NotPatternError2(w_Int).Data(w_i).ErrorFlg = True And (g_NotPatternError2(w_Int).Data(w_i).ErrorDate >= m_ViewFrom And g_NotPatternError2(w_Int).Data(w_i).ErrorDate <= m_ViewTo) Or (g_NotPatternError2(w_Int).Data(w_i).EndDate >= m_ViewFrom And g_NotPatternError2(w_Int).Data(w_i).EndDate <= m_ViewTo) Then

                        w_intCnt = UBound(m_ErrorList) + 1
                        ReDim Preserve m_ErrorList(w_intCnt)

                        '�G���[�̏ꍇ
                        m_ErrorList(w_intCnt).StaffIdx = g_NotPatternError2(w_Int).StaffIdx
                        m_ErrorList(w_intCnt).StaffName = g_NotPatternError2(w_Int).StaffName
                        m_ErrorList(w_intCnt).Date_Renamed = g_NotPatternError2(w_Int).Data(w_i).ErrorDate
                        m_ErrorList(w_intCnt).ErrorDetail = g_NotPatternError2(w_Int).Data(w_i).ErrorPattern
                        m_ErrorList(w_intCnt).ColIdx = g_NotPatternError2(w_Int).Data(w_i).ColIdx
                        '�G���[����
                        m_ErrorList(w_intCnt).ErrorName = "�֎~�Ζ��p�^�[��"
                    End If
                Next w_i
			Next w_Int
        End If

		If m_ChkNotKinmu = True Then
            '�֎~�Ζ�
			'�G���[�\���󋵕\��
			For w_Int = 1 To UBound(g_NotKinmuError2)
                For w_i = 1 To UBound(g_NotKinmuError2(w_Int).Data)
                    If g_NotKinmuError2(w_Int).Data(w_i).ErrorFlg = True And g_NotKinmuError2(w_Int).Data(w_i).ErrorDate >= m_ViewFrom And g_NotKinmuError2(w_Int).Data(w_i).ErrorDate <= m_ViewTo Then

                        w_intCnt = UBound(m_ErrorList) + 1
                        ReDim Preserve m_ErrorList(w_intCnt)

                        '�G���[�̏ꍇ
                        m_ErrorList(w_intCnt).StaffIdx = g_NotKinmuError2(w_Int).StaffIdx
                        m_ErrorList(w_intCnt).StaffName = g_NotKinmuError2(w_Int).StaffName
                        m_ErrorList(w_intCnt).Date_Renamed = g_NotKinmuError2(w_Int).Data(w_i).ErrorDate
                        m_ErrorList(w_intCnt).ErrorDetail = g_NotKinmuError2(w_Int).Data(w_i).KinmuName
                        m_ErrorList(w_intCnt).ColIdx = g_NotKinmuError2(w_Int).Data(w_i).ColIdx
                        '�G���[����
                        m_ErrorList(w_intCnt).ErrorName = "�֎~�Ζ�"
                    End If
                Next w_i
			Next w_Int
        End If

        If m_ChkStaffPattern = True Then
            '�֎~�E���p�^�[��

            '�G���[�\���󋵕\��
            For w_Int = 1 To UBound(g_NotStaffPatternError2)
                For w_i = 1 To UBound(g_NotStaffPatternError2(w_Int).Data)
                    If g_NotStaffPatternError2(w_Int).Data(w_i).ErrorFlg = True And g_NotStaffPatternError2(w_Int).Data(w_i).ErrorDate >= m_ViewFrom And g_NotStaffPatternError2(w_Int).Data(w_i).ErrorDate <= m_ViewTo Then

                        w_intCnt = UBound(m_ErrorList) + 1
                        ReDim Preserve m_ErrorList(w_intCnt)

                        '�G���[�̏ꍇ
                        m_ErrorList(w_intCnt).StaffIdx = g_NotStaffPatternError2(w_Int).StaffIdx
                        m_ErrorList(w_intCnt).StaffName = g_NotStaffPatternError2(w_Int).StaffName
                        m_ErrorList(w_intCnt).Date_Renamed = g_NotStaffPatternError2(w_Int).Data(w_i).ErrorDate
                        m_ErrorList(w_intCnt).ErrorDetail = g_NotStaffPatternError2(w_Int).Data(w_i).ErrorPattern
                        m_ErrorList(w_intCnt).ColIdx = g_NotStaffPatternError2(w_Int).Data(w_i).ColIdx
                        '�G���[����
                        m_ErrorList(w_intCnt).ErrorName = "�֎~�E���p�^�[��"
                    End If
                Next w_i
            Next w_Int
        End If

        If m_ChkGiryoType = True Then
            '�o���敪�̑g�ݍ��킹�`�F�b�N

            '�G���[�\���󋵕\��
            For w_Int = 1 To UBound(g_NotGiryoCheckError)
                For w_i = 1 To UBound(g_NotGiryoCheckError(w_Int).Data)
                    If g_NotGiryoCheckError(w_Int).Data(w_i).ErrorFlg = True And g_NotGiryoCheckError(w_Int).Data(w_i).ErrorDate >= m_ViewFrom And g_NotGiryoCheckError(w_Int).Data(w_i).ErrorDate <= m_ViewTo Then

                        w_intCnt = UBound(m_ErrorList) + 1
                        ReDim Preserve m_ErrorList(w_intCnt)

                        '�G���[�̏ꍇ
                        m_ErrorList(w_intCnt).StaffIdx = g_NotGiryoCheckError(w_Int).StaffIdx
                        m_ErrorList(w_intCnt).StaffName = g_NotGiryoCheckError(w_Int).StaffName
                        m_ErrorList(w_intCnt).Date_Renamed = g_NotGiryoCheckError(w_Int).Data(w_i).ErrorDate
                        m_ErrorList(w_intCnt).ErrorDetail = g_NotGiryoCheckError(w_Int).Data(w_i).ErrorPattern
                        m_ErrorList(w_intCnt).ColIdx = g_NotGiryoCheckError(w_Int).Data(w_i).ColIdx
                        '�G���[����
                        m_ErrorList(w_intCnt).ErrorName = "�o���敪�̑g�ݍ��킹"
                    End If
                Next w_i
            Next w_Int
        End If

        If m_RenzokuKinmuCheck = True Then
            '�񐔑I��

            '�e�Ζ����i�W�v���ځj�̃G���[���\��
            Call Set_RenzokuCountErrInf()
        End If

        If m_AbsoluteKinmuCheck = True Then
            '�K�{�Ζ��̑g�ݍ��킹�`�F�b�N

            '�G���[�\���󋵕\��
            For w_Int = 1 To UBound(g_NotAbsKinmuCheckError)
                For w_i = 1 To UBound(g_NotAbsKinmuCheckError(w_Int).Data)
                    If g_NotAbsKinmuCheckError(w_Int).Data(w_i).ErrorFlg = True And g_NotAbsKinmuCheckError(w_Int).Data(w_i).ErrorDate >= m_ViewFrom And g_NotAbsKinmuCheckError(w_Int).Data(w_i).ErrorDate <= m_ViewTo Then

                        w_intCnt = UBound(m_ErrorList) + 1
                        ReDim Preserve m_ErrorList(w_intCnt)

                        '�G���[�̏ꍇ
                        m_ErrorList(w_intCnt).StaffIdx = g_NotAbsKinmuCheckError(w_Int).StaffIdx
                        m_ErrorList(w_intCnt).StaffName = g_NotAbsKinmuCheckError(w_Int).StaffName
                        m_ErrorList(w_intCnt).Date_Renamed = g_NotAbsKinmuCheckError(w_Int).Data(w_i).ErrorDate
                        m_ErrorList(w_intCnt).ErrorDetail = g_NotAbsKinmuCheckError(w_Int).Data(w_i).ErrorPattern
                        m_ErrorList(w_intCnt).ColIdx = g_NotAbsKinmuCheckError(w_Int).Data(w_i).ColIdx
                        '�G���[����
                        m_ErrorList(w_intCnt).ErrorName = "�K�{�Ζ��̑g�ݍ��킹"
                    End If
                Next w_i
            Next w_Int
        End If

		'���t���\�����Ƀ\�[�g
		Call SortData()

		'---���X�g�ɕ\��---
		With lvwErrorList
			For w_i = 1 To UBound(m_ErrorList)
                w_str = Format(m_ErrorList(w_i).Date_Renamed, "0000/00/00")
                w_str = Format(CDate(w_str), "M/d")
				
				'Ͻ�����
				w_LItem = .Items.Add(w_str)
                If w_LItem.SubItems.Count > 1 Then
                    w_LItem.SubItems(1).Text = m_ErrorList(w_i).StaffName
                Else
                    w_LItem.SubItems.Insert(1, New ListViewItem.ListViewSubItem(Nothing, m_ErrorList(w_i).StaffName))
                End If

                If w_LItem.SubItems.Count > 2 Then
                    w_LItem.SubItems(2).Text = m_ErrorList(w_i).ErrorName
                Else
                    w_LItem.SubItems.Insert(2, New ListViewItem.ListViewSubItem(Nothing, m_ErrorList(w_i).ErrorName))
                End If

                If w_LItem.SubItems.Count > 3 Then
                    w_LItem.SubItems(3).Text = m_ErrorList(w_i).ErrorDetail
                Else
                    w_LItem.SubItems.Insert(3, New ListViewItem.ListViewSubItem(Nothing, m_ErrorList(w_i).ErrorDetail))
                End If
            Next w_i
			
			If UBound(m_ErrorList) > 0 Then
				'�P�Ԗڂ̱��т�I����Ԃ�
                .Items.Item(0).Selected = True
                .FocusedItem = .Items.Item(0)
			End If
		End With
		
		Exit Sub
Set_ErrorInf: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
	End Sub
	
	Private Sub Set_IntervalErrInf()
		On Error GoTo Set_IntervalErrInf
		Const W_SUBNAME As String = "NSK0000HN Set_IntervalErrInf"
		
		Dim w_Loop As Short
        Dim w_DataIndex As Short
		Dim w_Int As Short
		
        For w_Loop = 1 To UBound(g_KikanError2)
            For w_Int = 1 To UBound(g_KikanError2(w_Loop).InterValErr)
                If g_KikanError2(w_Loop).InterValErr(w_Int).ErrorFlg = True And g_KikanError2(w_Loop).InterValErr(w_Int).ErrorDate >= m_ViewFrom And g_KikanError2(w_Loop).InterValErr(w_Int).ErrorDate <= m_ViewTo Then

                    '�G���[���\��
                    w_DataIndex = UBound(m_ErrorList) + 1
                    ReDim Preserve m_ErrorList(w_DataIndex)

                    m_ErrorList(w_DataIndex).Date_Renamed = g_KikanError2(w_Loop).InterValErr(w_Int).ErrorDate
                    m_ErrorList(w_DataIndex).ErrorDetail = g_KikanError2(w_Loop).InterValErr(w_Int).ErrorName
                    m_ErrorList(w_DataIndex).StaffIdx = g_KikanError2(w_Loop).StaffIdx
                    m_ErrorList(w_DataIndex).StaffName = g_KikanError2(w_Loop).StaffName
                    '���t�C���f�b�N�X
                    m_ErrorList(w_DataIndex).ColIdx = g_KikanError2(w_Loop).InterValErr(w_Int).ColIdx
                    '�G���[����
                    m_ErrorList(w_DataIndex).ErrorName = "�Ζ��^�x���Ԋu"
                End If
            Next w_Int
        Next w_Loop
		
		Exit Sub
Set_IntervalErrInf: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
		End
	End Sub
	
	Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
		If lvwErrorList.Items.Count > 0 Then
            m_intSelIdx = lvwErrorList.FocusedItem.Index + 1
		Else
			m_intSelIdx = 0
		End If
        Call General.paSaveFieldWidth(lvwErrorList, General.G_STRMAINKEY2 & "\NSK0000H\", Me.Tag)
        Call General.paPutWindowPositon(Me, General.G_STRMAINKEY2 & "\NSK0000H\")
		
		Me.Close()
		
	End Sub
	
	Private Sub frmNSK0000HN_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo Form_Load
		Const W_SUBNAME As String = "NSK0000HN Form_Load"
		
        Dim w_Cnt As Short
        Dim clmX As ColumnHeader
        '2018/09/21 K.I Add Start-------------------------
        Dim w_Left As String
        Dim w_Top As String
        '2018/09/21 K.I Add End---------------------------

        ReDim m_ErrorList(0)
		
		If w_Cnt = 0 Then
        Else
            w_Cnt = 1
		End If
		
		'���X�g�̃w�b�_�\��
		With lvwErrorList
			.Columns.Clear()
			'ͯ�ް�쐬
            clmX = .Columns.Add("", "���t", Integer.Parse(General.paTwipsTopixels(700)))
            clmX = .Columns.Add("", "���@��", Integer.Parse(General.paTwipsTopixels(2500)))
            clmX = .Columns.Add("", "�G���[����", Integer.Parse(General.paTwipsTopixels(2437)))
            clmX = .Columns.Add("", "���@�e", Integer.Parse(General.paTwipsTopixels(2000)))
            Call General.paSetFieldWidth(lvwErrorList, General.G_STRMAINKEY2 & "\NSK0000H", Me.Tag, False)
			lvwErrorList.Visible = True
		End With
		
		'�G���[���̕\��
		Call Set_ErrorInf(w_Cnt)

        '����޳�̕\���߼޼�݂�ݒ肷��
        '2018/09/21 K.I Upd Start------------------------------------------------------------------------------------------
        '���W�X�g���擾���폜
        'Call General.paGetWindowPositon(Me, General.G_STRMAINKEY2 & "\" & "NSK0000H\")
        '��ʒ���
        w_Left = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - CSng(Me.Width)) / 2)
        w_Top = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - CSng(Me.Height)) / 2)
        Me.SetBounds((CSng(w_Left)), (CSng(w_Top)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
        '2018/09/21 K.I Upd End--------------------------------------------------------------------------------------------

        Exit Sub
Form_Load: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
		End
	End Sub
	
	Private Sub lvwErrorList_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvwErrorList.DoubleClick
		If lvwErrorList.Items.Count > 0 Then
			Call cmdClose_Click(cmdClose, New System.EventArgs())
		End If
    End Sub

	'���t /�\�����ɕ��ёւ�
	Private Sub SortData()
		On Error GoTo SortData
		Const W_SUBNAME As String = "NSK0000HN SortData"
		
		Dim w_Int As Short
		Dim w_Int2 As Short
		'��ėp�z��̊m��
		Dim w_WorkTbl As OutputType
		
		'�E���l���� �J��Ԃ�
		For w_Int = 1 To UBound(m_ErrorList)
			'(�E���l�� - ��ďI���l��) �J��Ԃ�
			For w_Int2 = 1 To UBound(m_ErrorList) - w_Int
				'���ёւ� ���s ?
				If (m_ErrorList(w_Int).Date_Renamed > m_ErrorList(w_Int + w_Int2).Date_Renamed) Then
					'�\��������҂����傫���ꍇ ����ւ�
                    w_WorkTbl = m_ErrorList(w_Int)
                    m_ErrorList(w_Int) = m_ErrorList(w_Int + w_Int2)
                    m_ErrorList(w_Int + w_Int2) = w_WorkTbl
					
				ElseIf (m_ErrorList(w_Int).Date_Renamed = m_ErrorList(w_Int + w_Int2).Date_Renamed) And m_ErrorList(w_Int).StaffIdx > m_ErrorList(w_Int + w_Int2).StaffIdx Then 
					'�\��������҂Ɠ����ŁA�E���Ǘ��ԍ����傫���ꍇ ����ւ�
                    w_WorkTbl = m_ErrorList(w_Int)
                    m_ErrorList(w_Int) = m_ErrorList(w_Int + w_Int2)
                    m_ErrorList(w_Int + w_Int2) = w_WorkTbl
				End If
			Next w_Int2
			'�҂�����
            Application.DoEvents()
		Next w_Int
		
		Exit Sub
SortData: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
		End
	End Sub
End Class