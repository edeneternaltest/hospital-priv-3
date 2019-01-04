Option Strict Off
Option Explicit On
Friend Class frmNSK0000HD
    Inherits General.FormBase
    Public NSK0000H_sql As New clsNSK0000H_sql
    '/ ��۸��і��́F�v��쐬/���ѓo�^���j���[
    '/        �h�c�FNSK0000HD
    '/        �T�v�F����/�u�����
    '/
    '/
    '/      �쐬�ҁF S.Y    CREATE 2000/08/14           REV 01.00
    '/      �X�V�ҁF        UPDATE     /  /             REV 01.01
    '/
    '/     Copyright (C) Inter co.,ltd 2000
    '/----------------------------------------------------------------------/
    '----- �ƭ��ް�ԍ� ��` --------------------------------------------------
    '�ƭ��ް�z��̃C���f�b�N�X��\���萔
    '[�\��]���j���[�z��̃C���f�b�N�X��\���萔
    Private Const M_MenuPalette As Short = 0 '��گ�
    '°��ް��Key�萔
    Private Const M_ToolBarKey_Palette As String = "Palette" '�p���b�g

    '-----------------------------------------------------------------
    '   �� �� �� ��
    '-----------------------------------------------------------------
    Private m_PgmFlg As String '�N��Ӱ��
    Private m_KenChiFlg As BasNSK0000H.geKenChiFlg '���� Or �u���i�񋓌^�Ő錾�j
    Private m_KinmuCD() As String 'KinmuCD
    Private m_KensakuKinmuCD As String '�����Ζ�CD
    Private m_KensakuTaisyo As Short '�Ώ�
    Private m_TaisyoOnly As Short '�Ώ�
    Private m_ChikanKinmuCD As String '�u���Ζ�CD
    Private m_RiyuKBN As String '�u����̗��R�敪
    Private m_FormShowFlg As Boolean '��ʂ��\������Ă��邩
    Private m_TargetKinmuCD() As String
    Private m_CmbHT As New Hashtable

    '2014/04/23 Saijo add start P-06979-----------------------------------------------------------------------
    Private m_strKinmuEmSecondFlg As String '�Ζ��L���S�p�Q�����Ή��t���O(0�F�Ή����Ȃ��A1:�Ή�����)
    '2014/04/23 Saijo add end P-06979-------------------------------------------------------------------------

    '------------------------------------------------------------------
    '  ����Đ錾
    '------------------------------------------------------------------
	Event PaletteEnabled()
	Event SelectKinmu()
	Event ChikanKinmu()
	Event SelectAllKinme()
	Event ChikanAllKinmu()
    Event SetEditMenu()
	
	'-- �Ζ��L�� �ޔ�z�� ----------------------
	Private Structure Kinmu_Type
		Dim CD As String 'KinmuCD
		Dim KinmuName As String '����
		Dim Mark As String '�L��
		Dim HolBunruiCD As String '�x�ݕ���CD
        Dim EffToDate As Integer '�L���I����
    End Structure

	Private m_KinmuM() As Kinmu_Type '�Ζ����z��
	
    Private m_StartDate As Integer
	
    '�J�n���擾
	Public WriteOnly Property pStartDate() As Integer
		Set(ByVal Value As Integer)
            m_StartDate = Value
		End Set
    End Property

    '����/�u�� �׸� ���擾����
	Public WriteOnly Property pKenChiFlg() As BasNSK0000H.geKenChiFlg
		Set(ByVal Value As BasNSK0000H.geKenChiFlg)
			m_KenChiFlg = Value
		End Set
	End Property
	
	Public WriteOnly Property pPgmFlg() As String
		Set(ByVal Value As String)
			m_PgmFlg = Value
		End Set
    End Property

	Public ReadOnly Property pKensakuKinmuCD() As String
		Get
            '�����Ζ�����
			pKensakuKinmuCD = m_KensakuKinmuCD
        End Get
	End Property
	
	Public ReadOnly Property pRiyuKBN() As String
		Get
            '���R�敪
			pRiyuKBN = m_RiyuKBN
        End Get
	End Property
	
	Public ReadOnly Property pKensakuTaisyo() As String
		Get
            '�Ώ�
			pKensakuTaisyo = CStr(m_KensakuTaisyo)
        End Get
	End Property

	Public ReadOnly Property pTaisyoOnly() As Short
		Get
            '�Ώ�
			pTaisyoOnly = m_TaisyoOnly
        End Get
	End Property
	
	Public ReadOnly Property pChikanKinmuCD() As String
		Get
            '�u���Ζ�����
			pChikanKinmuCD = m_ChikanKinmuCD
        End Get
	End Property

	Public Property pShowFlg() As Boolean
		Get
			pShowFlg = m_FormShowFlg
        End Get

		Set(ByVal Value As Boolean)
			m_FormShowFlg = Value
		End Set
	End Property
	
    Private Sub cboTaisyo_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboTaisyo.SelectedIndexChanged
        On Error GoTo cboTaisyo_Click
        Const W_SUBNAME As String = "NSK0000HD cboTaisyo_Click"

        '�Ώۂ�"���ׂ�"�̎��A"�Ώۂ݂̂Ō�������"�����ޯ�����g�p�s��
        If cboTaisyo.SelectedIndex = 0 Then
            If chkTaisyoOnly.Enabled = True Then
                chkTaisyoOnly.Enabled = False
            End If
        Else
            If chkTaisyoOnly.Enabled = False Then
                chkTaisyoOnly.Enabled = True
            End If
        End If

        Exit Sub
cboTaisyo_Click:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub

	Private Sub cmd_Next_Chikan_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Next_Chikan.Click
		On Error GoTo cmd_Next_Chikan_Click
		Const W_SUBNAME As String = "NSK0000HD cmd_Next_Chikan_Click"
		
        '����/�u���ް����擾
		Call Get_KenChiData()
		
		'������u�����s
		If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Kensaku Then
			
			'�������s
			RaiseEvent SelectKinmu()
			
		Else
			
			'�u�����s
			RaiseEvent ChikanKinmu()
			
		End If
		
		Exit Sub
cmd_Next_Chikan_Click: 
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
		End
	End Sub
	
	Private Sub Get_KenChiData()
		On Error GoTo Get_KenChiData
		Const W_SUBNAME As String = "NSK0000HD Get_KenChiData"
		
		Dim w_Index As Short
		
		'�����Ζ�����
        w_Index = GetItemData(cboKensakuKinmu, cboKensakuKinmu.SelectedIndex)
		If w_Index > 0 Then
			m_KensakuKinmuCD = m_KinmuCD(w_Index)
		Else
			m_KensakuKinmuCD = "000"
		End If
		
		'�Ώ�
		w_Index = cboTaisyo.SelectedIndex
		If w_Index >= 0 Then
			'�ψ���Ζ��̗��R�敪��"5"�̂��߁A�����Ζ��̋敪��"6"�ɂ���
			If w_Index = 5 Then
				m_KensakuTaisyo = CShort(CStr(6))
			Else
				m_KensakuTaisyo = CShort(CStr(w_Index))
			End If
		Else
			m_KensakuTaisyo = CShort("0")
		End If
		
		'�Ώۂ݂̂�����
		m_TaisyoOnly = chkTaisyoOnly.CheckState
		
		'�u���̏ꍇ
		If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Chikan Then
			
			'�u���Ζ�����
            w_Index = GetItemData(cboChikanKinmu, cboChikanKinmu.SelectedIndex)
			If w_Index > 0 Then
				m_ChikanKinmuCD = m_KinmuCD(w_Index)
			Else
				m_ChikanKinmuCD = "000"
            End If

			'�u����̗��R�敪
            If GetItemData(cboRiyu, cboRiyu.SelectedIndex) = 2 Then
                '�v��
                m_RiyuKBN = "2"
            ElseIf GetItemData(cboRiyu, cboRiyu.SelectedIndex) = 3 Then
                '��]
                m_RiyuKBN = "3"
            ElseIf GetItemData(cboRiyu, cboRiyu.SelectedIndex) = 4 Then
                '�Čf
                m_RiyuKBN = "4"
            Else
                m_RiyuKBN = "1"
            End If
        End If
		
		Exit Sub
Get_KenChiData: 
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
		End
	End Sub

	Private Sub cmd_Select_AllChikan_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Select_AllChikan.Click
		On Error GoTo cmd_Select_AllChikan_Click
		Const W_SUBNAME As String = "NSK0000HD cmd_Select_AllChikan_Click"
		
        '����/�u���ް����擾
		Call Get_KenChiData()
		
		'������u�����s
		If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Kensaku Then
			
			'��\���ɂ���
			Me.Hide()
			m_FormShowFlg = False
			'����(�I��)���s
			RaiseEvent SelectAllKinme()
			
			'�p���b�g�g�p��
			RaiseEvent PaletteEnabled()
			
            '�؂���A�R�s�[�A�\��t���̐���
			RaiseEvent SetEditMenu()
		Else
			
			'��\���ɂ���
			Me.Hide()
			m_FormShowFlg = False
			'�u��(���ׂ�)���s
			RaiseEvent ChikanAllKinmu()
			
			'�p���b�g�g�p��
			RaiseEvent PaletteEnabled()
			
		End If
		
		Exit Sub
cmd_Select_AllChikan_Click: 
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
		End
	End Sub
	
	Private Sub cmdEnd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEnd.Click
		On Error GoTo cmdEnd_Click
		Const W_SUBNAME As String = "NSK0000HD cmdEnd_Click"
		
		'�p���b�g�g�p��
		RaiseEvent PaletteEnabled()
		
        '�؂���A�R�s�[�A�\��t���̐���
		If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Kensaku Then
			RaiseEvent SetEditMenu()
		End If
		
		'��ʂ�Hide����
		Me.Hide()
		m_FormShowFlg = False
		
		Exit Sub
cmdEnd_Click: 
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
		End
    End Sub

    Private Sub frmNSK0000HD_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo Form_Load
        Const W_SUBNAME As String = "NSK0000HD Form_Activate"

        If Me.Visible = True Then
            '�ŏ�ʂɐݒ�
            Call General.paSetDialogPos(Me)
        End If

        Exit Sub
Form_Load:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub
	
    Public Sub frmNSK0000HD_Load()
        On Error GoTo Form_Load
        Const W_SUBNAME As String = "NSK0000HD Form_Load"

        Dim w_ImagePath As String
        Dim w_SystemPath As String
        '2018/09/21 K.I Add Start-------------------------
        Dim w_Left As String
        Dim w_Top As String
        '2018/09/21 K.I Add End---------------------------


        '̫�� ���� �ݒ�
        If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Chikan Then
            Me.Icon = New Icon(g_ImagePath & G_PERMUTATION_ICO)
        Else
            Me.Icon = New Icon(g_ImagePath & G_SEARCH_ICO)
        End If

        '�E�B���h�D����ʂ̍ŏ�ʂɐݒ�
        Call General.paSetDialogPos(Me)

        '�����ޯ���ݒ�
        '�\���ΏۋΖ�CD���擾
        Call Get_TargetKinmuCD()
        Call Set_ComboBox()
        '2014/04/23 Saijo add start P-06979------------------------------------
        '���ڐݒ�̎擾
        m_strKinmuEmSecondFlg = Get_ItemValue(General.g_strHospitalCD)
        '�Ζ��L���S�p�Q�����Ή��̃��C�A�E�g�ύX
        Call SetKinmuSecondView()
        '2014/04/23 Saijo add end P-06979--------------------------------------

        '�r�b�g�}�b�v�C���[�W�ۑ��p�X�擾
        w_SystemPath = My.Application.Info.DirectoryPath
        w_ImagePath = General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, "ImagePath", w_SystemPath & "image\")

        cmd_Next_Chikan.Text = ""
        cmd_Select_AllChikan.Text = ""

        '�����ݒ�
        If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Kensaku Then
            '�����̂Ƃ�
            _lblKomoku_1.Enabled = False
            _lblKomoku_1.Visible = False
            cboChikanKinmu.Enabled = False
            cboChikanKinmu.Visible = False
            _lblKomoku_3.Enabled = False
            _lblKomoku_3.Visible = False
            cboRiyu.Enabled = False
            cboRiyu.Visible = False
            cmd_Next_Chikan.Image = Image.FromFile(w_ImagePath & "��������.bmp")
            cmd_Select_AllChikan.Image = Image.FromFile(w_ImagePath & "�I��.bmp")
            Me.Text = "����"
            If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Then
                '�Ζ��ύX�̏ꍇ�A�Ώۂ��g�p�s�ɂ���
                '��̫�ĂőΏۂ�"���ׂ�"�Ƃ���
                _lblKomoku_2.Enabled = False
                cboTaisyo.Enabled = False
            End If
        Else
            '�u���̂Ƃ�
            _lblKomoku_1.Enabled = True
            _lblKomoku_1.Visible = True
            cboChikanKinmu.Enabled = True
            cboChikanKinmu.Visible = True
            _lblKomoku_3.Enabled = True
            _lblKomoku_3.Visible = True
            cboRiyu.Enabled = True
            cboRiyu.Visible = True
            cmd_Next_Chikan.Image = Image.FromFile(w_ImagePath & "�u�����s.bmp")
            cmd_Select_AllChikan.Image = Image.FromFile(w_ImagePath & "�S�Ēu��.bmp")
            Me.Text = "�u��"
            If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Then
                '�Ζ��ύX�̏ꍇ�A�Ώۂƒu����̗��R�敪���g�p�s�ɂ���
                '��̫�ĂőΏۂ�"���ׂ�"�Ƃ���
                '�u����̗��R�敪�͒ʏ�Ƃ���
                _lblKomoku_2.Enabled = False
                cboTaisyo.Enabled = False
                _lblKomoku_3.Enabled = False
                cboRiyu.Enabled = False
            End If
        End If

        '����޳�̕\���߼޼�݂�ݒ肷��
        '2018/09/21 K.I Upd Start------------------------------------------------------------------------------------------
        '���W�X�g���擾���폜
        'Call General.paGetWindowPositon(Me, General.G_STRMAINKEY2 & "\" & g_AppName)
        '��ʒ���
        w_Left = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - CSng(Me.Width)) / 2)
        w_Top = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - CSng(Me.Height)) / 2)
        Me.SetBounds((CSng(w_Left)), (CSng(w_Top)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
        '2018/09/21 K.I Upd End--------------------------------------------------------------------------------------------

        Exit Sub
Form_Load:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub
	
	Private Sub Set_ComboBox()
		On Error GoTo Set_ComboBox
		Const W_SUBNAME As String = "NSK0000HD Set_ComboBox"
		
		Dim w_Int As Short
		Dim w_Cnt As Short
		Dim w_str As String
        Dim w_RecCnt As Short
		Dim w_Sql As String
		Dim w_Rs As ADODB.Recordset
		Dim w_KinmuCD_F As ADODB.Field
		Dim w_�L��_F As ADODB.Field
		Dim w_����_F As ADODB.Field
		Dim w_�x�ݕ���CD_F As ADODB.Field
        Dim w_�L���I����_F As ADODB.Field
        Dim w_lngLoop As Integer
		Dim w_lngDataIdx As Integer
        Dim w_CmbCnt As Short
        Dim w_CmbCnt2 As Short
        '2017/05/02 Christopher Upd Start
        'Select���ҏW
        'w_Sql = "SELECT   KINMUCD "
        'w_Sql = w_Sql & ",MARKF "
        'w_Sql = w_Sql & ",NAME "
        'w_Sql = w_Sql & ",HOLIDAYBUNRUICD "
        '      w_Sql = w_Sql & ",EFFTODATE "
        'w_Sql = w_Sql & "FROM NS_KINMUNAME_M "
        '      w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
        'w_Sql = w_Sql & "ORDER BY DISPNO "
        ''RecordSet ��޼ު�Ă̐���
        '      w_Rs = General.paDBRecordSetOpen(w_Sql)

        Call NSK0000H_sql.select_NS_KINMUNAME_M_02(w_Rs)
        'Upd End
        If w_Rs.RecordCount <= 0 Then
            '�ް����Ȃ��Ƃ�
            w_Rs.Close()
            Exit Sub
        Else
            With w_Rs
                '�ް�������Ƃ�

                '�ް����� �擾
                .MoveLast()
                w_RecCnt = .RecordCount
                .MoveFirst()

                '̨���� ��޼ު�� �쐬
                w_KinmuCD_F = .Fields("KINMUCD")
                w_�L��_F = .Fields("MARKF")
                w_����_F = .Fields("NAME")
                w_�x�ݕ���CD_F = .Fields("HOLIDAYBUNRUICD")
                w_�L���I����_F = .Fields("EFFTODATE")
				
                For w_Int = 1 To w_RecCnt
                    w_str = w_KinmuCD_F.Value
                    '�\���ΏۋΖ�CD���`�F�b�N
                    For w_lngLoop = 1 To UBound(m_TargetKinmuCD)
                        If w_str = m_TargetKinmuCD(w_lngLoop) Then
                            w_lngDataIdx = w_lngDataIdx + 1
                            '�z��m��
                            ReDim Preserve m_KinmuM(w_lngDataIdx)

                            m_KinmuM(w_lngDataIdx).CD = w_str
                            m_KinmuM(w_lngDataIdx).KinmuName = IIf(IsDBNull(w_����_F.Value), "", w_����_F.Value)
                            m_KinmuM(w_lngDataIdx).Mark = IIf(IsDBNull(w_�L��_F.Value), "", w_�L��_F.Value)
                            m_KinmuM(w_lngDataIdx).HolBunruiCD = IIf(IsDBNull(w_�x�ݕ���CD_F.Value), "", w_�x�ݕ���CD_F.Value)
                            m_KinmuM(w_lngDataIdx).EffToDate = IIf(IsDBNull(w_�L���I����_F.Value), 0, w_�L���I����_F.Value) '-20080909-okamoto-Add

                            Exit For
                        End If
                    Next w_lngLoop

                    .MoveNext()
                Next w_Int
            End With
		End If
		w_Rs.Close()
		
		'--- �����ޯ���̐ݒ� ---
		cboKensakuKinmu.Items.Clear()
		cboChikanKinmu.Items.Clear()
		'ListIndex=0�͖�����
        '�����p�Ζ�����
        cboKensakuKinmu.Items.Add("������")
        SetItemData(cboKensakuKinmu, 0, 0)
		
        If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Then
            '�Ζ��ύX�ŋN������Ă���Ƃ��͖������ւ̒u���͍s��Ȃ�
        Else
            '�u���p�Ζ�����
            cboChikanKinmu.Items.Add("������")
            SetItemData(cboChikanKinmu, 0, 0)
        End If
		
		ReDim m_KinmuCD(0)
		
		For w_Int = 1 To UBound(m_KinmuM)
			If m_KinmuM(w_Int).CD <> "" Then
				If m_KinmuM(w_Int).EffToDate >= m_StartDate Or m_KinmuM(w_Int).EffToDate = 0 Or m_KinmuM(w_Int).EffToDate = 99999999 Then '-20080909-okamoto-Add
					w_Cnt = w_Cnt + 1
					w_str = Trim(m_KinmuM(w_Int).KinmuName)
                    w_str = w_str & Space(4 - General.paLenB(w_str))
					w_str = w_str & "(" & m_KinmuM(w_Int).Mark & ")"
					
					ReDim Preserve m_KinmuCD(w_Cnt)
					m_KinmuCD(w_Cnt) = m_KinmuM(w_Int).CD
					
					'�����p�Ζ�����
                    cboKensakuKinmu.Items.Add(w_str)
                    w_CmbCnt = w_CmbCnt + 1
                    SetItemData(cboKensakuKinmu, w_CmbCnt, w_Cnt)
                    
                    If General.g_lngDaikyuMng = 0 Then
                        If m_KinmuM(w_Int).HolBunruiCD <> General.G_STRDAIKYUBUNRUI Then '�x�ݕ���CD����x��CD�͑ΏۊO
                            '�u���p�Ζ�����
                            cboChikanKinmu.Items.Add(w_str)
                            w_CmbCnt2 = w_CmbCnt2 + 1
                            SetItemData(cboChikanKinmu, w_CmbCnt2, w_Cnt)
                        End If
                    Else
                        '�u���p�Ζ�����
                        cboChikanKinmu.Items.Add(w_str)
                        SetItemData(cboChikanKinmu, w_CmbCnt, w_Cnt)
                    End If
                End If
			End If
		Next w_Int
		
		'��̫�Ēl�ݒ�
		cboKensakuKinmu.SelectedIndex = 1
        If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Then
            cboChikanKinmu.SelectedIndex = 0
        Else
            cboChikanKinmu.SelectedIndex = 1
        End If
		
		'�u����̗��R�敪�����ޯ���ݒ�
		With cboRiyu
            .Items.Clear()
			If g_SaikeiFlg = True Then
                .Items.Add("�Čf")
                SetItemData(cboRiyu, 0, 4)
			Else
				.Items.Add("�ʏ�")
                .Items.Add("�v��")
                .Items.Add("��]")
                SetItemData(cboRiyu, 0, 1)
                SetItemData(cboRiyu, 1, 2)
                SetItemData(cboRiyu, 2, 3)
			End If
			.SelectedIndex = 0
        End With
		
		'�Ώۺ����ޯ���ݒ�
		With cboTaisyo
            .Items.Clear()
            .Items.Add("���ׂ�")
            .Items.Add("�ʏ�")
            .Items.Add("�v��")
            .Items.Add("��]")
            .Items.Add("�Čf")
            SetItemData(cboTaisyo, 0, 0)
            SetItemData(cboTaisyo, 1, 1)
            SetItemData(cboTaisyo, 2, 2)
            SetItemData(cboTaisyo, 3, 3)
            SetItemData(cboTaisyo, 4, 4)
			If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Kensaku Then
                .Items.Add("����")
                SetItemData(cboTaisyo, 5, 5)
			End If
			
			'�Ώۺ����ޯ������̫�Ēl�ݒ�
			.SelectedIndex = 0
        End With
		
		'�Ώۂ݂̂̌��������ޯ������̫�Ēl�ݒ�
        chkTaisyoOnly.CheckState = CheckState.Unchecked

		Exit Sub
Set_ComboBox: 
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
	End Sub
	
    Private Sub frmNSK0000HD_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As FormClosingEventArgs) Handles Me.FormClosing
        Dim UnloadMode As CloseReason = eventArgs.CloseReason
        On Error GoTo Form_QueryUnload
        Const W_SUBNAME As String = "NSK0000HD Form_QueryUnload"

        If UnloadMode = CloseReason.UserClosing Then
            '��گĎg�p��
            RaiseEvent PaletteEnabled()

            '���۰��ƭ��������ꂽ�ꍇ��Unload���Ȃ�
            eventArgs.Cancel = True
            Me.Hide()
            m_FormShowFlg = False
        End If

        Exit Sub
Form_QueryUnload:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub
	
    Public Sub frmNSK0000HD_FormClosed()
        On Error GoTo Form_Unload
        Const W_SUBNAME As String = "NSK0000HD Form_Unload"

        '����޳�̕\���߼޼�݂��i�[����
        Call General.paPutWindowPositon(Me, General.G_STRMAINKEY2 & "\" & g_AppName)

        Exit Sub
Form_Unload:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub
	
    Private Sub Get_TargetKinmuCD()
        On Error GoTo Get_TargetKinmuCD
        Const W_SUBNAME As String = "NSK0000HD Get_TargetKinmuCD"

        Dim w_Int As Short
        Dim w_RecCnt As Short
        Dim w_Sql As String
        Dim w_Rs As ADODB.Recordset
        Dim w_KinmuCD_F As ADODB.Field

        '������
        ReDim m_TargetKinmuCD(0)
        '2017/05/02 Christopher Upd Start
        'Select���ҏW
        'w_Sql = "SELECT   KINMUCD "
        'w_Sql = w_Sql & "FROM NS_SETKINMUNAME_F "
        'w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
        'w_Sql = w_Sql & "AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
        'w_Sql = w_Sql & "ORDER BY DISPNO "
        ''RecordSet ��޼ު�Ă̐���
        'w_Rs = General.paDBRecordSetOpen(w_Sql)

        Call NSK0000H_sql.select_NS_SETKINMUNAME_F_01(w_Rs)
        'Upd End
        If w_Rs.RecordCount <= 0 Then
            '�ް����Ȃ��Ƃ�
            w_Rs.Close()
            Exit Sub
        Else
            With w_Rs
                '�ް�������Ƃ�

                '�ް����� �擾
                .MoveLast()
                w_RecCnt = .RecordCount
                .MoveFirst()

                '̨���� ��޼ު�� �쐬
                w_KinmuCD_F = .Fields("KINMUCD")

                ReDim m_TargetKinmuCD(w_RecCnt)

                For w_Int = 1 To w_RecCnt
                    m_TargetKinmuCD(w_Int) = w_KinmuCD_F.Value

                    .MoveNext()
                Next w_Int

            End With
        End If
        w_Rs.Close()

        Exit Sub
Get_TargetKinmuCD:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub SetItemData(ByVal p_Cmb As ComboBox, ByVal p_Index As Integer, ByVal p_Item As Object)
        m_CmbHT(p_Cmb.Name & " " & p_Index) = p_Item
    End Sub

    Private Function GetItemData(ByVal p_Cmb As ComboBox, ByVal p_Index As Integer) As Object
        GetItemData = m_CmbHT(p_Cmb.Name & " " & p_Index)
    End Function

    '2014/04/23 Saijo add start P-06979--------------------------------------------------------------------------------------------------
    '/----------------------------------------------------------------------/
    '/  �T�v�@�@�@�@  : �Ζ��L���S�p�Q�����Ή��̃��C�A�E�g�ύX
    '/  �p�����[�^    : �Ȃ�
    '/  �߂�l        : �Ȃ�
    '/----------------------------------------------------------------------/
    Private Sub SetKinmuSecondView()
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc '��ۼ��ެ���̈ꎞ�Ҕ�
        General.g_ErrorProc = "frmNSC1000HE SetKinmuSecondView"

        Try
            '�Ζ��L���S�p�Q�����Ή��t���O����
            If m_strKinmuEmSecondFlg = "0" Then
                '0�F�Ή����Ȃ�(�]���̋Ζ��L�����̓T�C�Y�ƍő�2�o�C�g)
                cboKensakuKinmu.Size = New System.Drawing.Size(94, 23)
                cboChikanKinmu.Size = New System.Drawing.Size(94, 23)
            Else
                '1�F�Ή�����(�S�p�Q�������\���ł���Ζ��L�����̓T�C�Y�ƍő�4�o�C�g)
                cboKensakuKinmu.Size = New System.Drawing.Size(112, 23)
                cboChikanKinmu.Size = New System.Drawing.Size(112, 23)
            End If

            General.g_ErrorProc = w_PreErrorProc '�Ҕ���ۼ��ެ�������ɖ߂�

        Catch ex As Exception
            Err.Raise(Err.Number)
        End Try
    End Sub
    '2014/04/23 Saijo add end  P-06979----------------------------------------------------------------------------------------------------
End Class