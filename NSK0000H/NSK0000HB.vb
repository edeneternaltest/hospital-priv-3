Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Friend Class frmNSK0000HB
    Inherits General.FormBase
    Public NSK0000H_sql As New clsNSK0000H_sql
    '=====================================================================================
    '   ��  ��  ��  ��
    '=====================================================================================
    '[�\��]���j���[�z��̃C���f�b�N�X��\���萔
    Private Const M_MenuPalette As Short = 0 '��گ�
    '[�ҏW]���j���[�z��̃C���f�b�N�X��\���萔
    Private Const M_MenuKensaku As Short = 7 '����
    Private Const M_MenuChikan As Short = 8 '�u��
    '°��ް��Key�萔
    Private Const M_ToolBarKey_Search As String = "KinmuSerach" '����
    Private Const M_ToolBarKey_Tikan As String = "KinmuTikan" '�u��

    '=====================================================================================
    '   ��  ��  ��  ��
    '=====================================================================================
    Private m_OuenDispFlg As Integer '�����Ζ��敪�̃��W�I�{�^�����p���b�g�ɕ\�����邩(1:���Ȃ�,0:����)
    '2015/04/14 Bando Add Start ========================
    Private m_DispKinmuCd As String '��]���[�h���̕\���ΏۋΖ�CD
    '2015/04/14 Bando Add End   ========================
	
	Private m_PgmFlg As String '�N��Ӱ��
	Private m_BtnClickFlg As Boolean 'True:����������� Or �Ζ����݂��د�����Ă���Ƃ�,False:�د�����Ă��Ȃ��Ƃ�
	'���ݑI�� �Ζ��L��
	Private m_SelNowKinmuCD As String 'KinmuCD
	Private m_SelNowRiyuKbn As String '���R�敪
	Private m_SelSetIdx As Integer '�I�����ꂽ�Z�b�g�̲��ޯ��.
    Private m_SetCDIdx As Integer '�Z�b�gCD���ޯ��
    Private m_lstOptRiyu As New List(Of Object)
    Private m_lstCmdKinmu As New List(Of Object)
    Private m_lstCmdYasumi As New List(Of Object)
    Private m_lstCmdTokushu As New List(Of Object)
    Private m_lstCmdSet As New List(Of Object)

    '2014/04/23 Shimizu add start P-06979-------------------------------------------------------------------
    Private m_strKinmuEmSecondFlg As String '�Ζ��L���S�p�Q�����Ή��t���O(0�F�Ή����Ȃ��A1:�Ή�����)
    '2014/04/23 Shimizu add end P-06979---------------------------------------------------------------------

	Private Structure Kinmu_Type
		Dim CD As String 'KinmuCD
		Dim Mark As String '�Ζ��L��
		Dim KinmuName As String 'KinmuName
		Dim KBunruiCD As String '�Ζ�����CD
		Dim ClickFlg As Boolean '���݂̏��(True:�������܂�Ă���Ƃ�,False:��ɖ߂��Ă���Ƃ�)
		Dim Setumei As String '����
    End Structure

	Private m_KinmuMark() As Kinmu_Type
	Private m_YasumiMark() As Kinmu_Type
    Private m_TokushuMark() As Kinmu_Type

	Private Structure SetKinmu_Type
		Dim Mark As String
		<VBFixedArray(10)> Dim CD() As String
		Dim StrText As String
		Dim ClickFlg As Boolean '���݂̏��(True:�������܂�Ă���Ƃ�,False:��ɖ߂��Ă���Ƃ�)
		Dim KinmuCnt As Integer
        Dim blnKinmu As Boolean
		
        Public Sub Initialize()
            ReDim CD(10)
        End Sub
	End Structure
	
	Private m_SetKinmuMark() As SetKinmu_Type
    Private m_StartDate As Integer
	
	'����Đ錾
	Event KensakuEnabled()

	'�J�n���擾
	Public WriteOnly Property pStartDate() As Integer
		Set(ByVal Value As Integer)
			m_StartDate = Value
		End Set
	End Property

	'�I�������Z�b�g���ޯ�����擾
	Public WriteOnly Property pSelKinmuIdx() As Integer
		Set(ByVal Value As Integer)
			m_SelSetIdx = Value
		End Set
    End Property

	'�Z�b�gCD���ޯ��
	Public WriteOnly Property pSetCDIdx() As Integer
		Set(ByVal Value As Integer)
			m_SetCDIdx = Value
		End Set
    End Property

	Public WriteOnly Property pPgmFlg() As String
		Set(ByVal Value As String)
			m_PgmFlg = Value
		End Set
    End Property

	'True:����������� Or �Ζ����݂��د�����Ă���Ƃ�,False:�د�����Ă��Ȃ��Ƃ�
	Public ReadOnly Property pBtnClickFlg() As Boolean
		Get
            '���݂̏��
			pBtnClickFlg = m_BtnClickFlg
        End Get
    End Property

	Public ReadOnly Property pSelNowKinmuCD() As String
		Get
            '�Ζ��L��
			pSelNowKinmuCD = m_SelNowKinmuCD
        End Get
    End Property

	'�I�����ꂽ�Z�b�g�̋Ζ���
	Public ReadOnly Property pSetCnt() As Integer
		Get
            pSetCnt = m_SetKinmuMark(m_SelSetIdx).KinmuCnt
        End Get
	End Property
	
	'�Z�b�g�̋Ζ�CD
	Public ReadOnly Property pGetSetCD() As String
		Get
            pGetSetCD = m_SetKinmuMark(m_SelSetIdx).CD(m_SetCDIdx)
        End Get
	End Property
	
	Public ReadOnly Property pSelNowRiyuKbn() As String
		Get
            '���R�敪
			pSelNowRiyuKbn = m_SelNowRiyuKbn
        End Get
    End Property

    Private Sub CScmdClose_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CScmdClose.Click
        On Error GoTo CScmdClose_Click
        Const W_SUBNAME As String = "NSK0000HB CScmdClose_Click"

        RaiseEvent KensakuEnabled()

        '���ٳ���޳��\��
        Me.Hide()

        Exit Sub
CScmdClose_Click:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub
	
    Private Sub CScmdErase_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CScmdErase.Click
        On Error GoTo CScmdErase_Click
        Const W_SUBNAME As String = "NSK0000HB CScmdErase_Click"

        Static w_SelKinmuCD As String '�I������Ă���Ζ�CD
        Dim w_ForeColor As Integer
        Dim w_BackColor As Integer
        Dim w_Index As Short
        Dim w_Int As Short
        Dim w_LoopFlg As Boolean
        Dim w_RegKey As String
        Dim w_RegStr As String
        Dim w_SetKinmuFlg As Boolean
        Dim w_Font As Font

        'ڼ޽�؊i�[��
        w_RegStr = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

        '�Ζ����݂̏�Ԃ���ɖ߂��Ă����Ԃ�
        '�Ζ�
        For w_Int = 0 To 14
            w_Font = m_lstCmdKinmu(w_Int).Font
            m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdKinmu(w_Int).Checked = False
        Next w_Int

        '�x��
        For w_Int = 0 To 9
            w_Font = m_lstCmdYasumi(w_Int).Font
            m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdYasumi(w_Int).Checked = False
        Next w_Int

        '����
        For w_Int = 0 To 4
            w_Font = m_lstCmdTokushu(w_Int).Font
            m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdTokushu(w_Int).Checked = False
        Next w_Int

        '�Z�b�g�Ζ�
        For w_Int = 0 To 4
            w_Font = m_lstCmdSet(w_Int).Font
            m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdSet(w_Int).Checked = False
        Next w_Int

        '���݂̏��
        If CScmdErase.Checked = False Then
            '���݂�������Ă��Ȃ��ꍇ

            '�I������Ă����Ζ����݂������ꂽ��Ԃ�
            For w_Int = 0 To 14
                If w_Int <= UBound(m_KinmuMark) - 1 Then
                    If w_Int <= UBound(m_KinmuMark) - 1 And (w_Int + 1 + HscKinmu.Value * 3) <= UBound(m_KinmuMark) Then
                        If m_KinmuMark(w_Int + 1 + HscKinmu.Value * 3).ClickFlg = True Then
                            w_Font = m_lstCmdKinmu(w_Int).Font
                            m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
                            m_lstCmdKinmu(w_Int).Checked = True
                            w_LoopFlg = True
                            Exit For
                        End If
                    End If
                End If
            Next w_Int

            If w_LoopFlg = False Then
                For w_Int = 0 To 9
                    If w_Int <= UBound(m_YasumiMark) - 1 Then
                        If w_Int <= UBound(m_YasumiMark) - 1 And (w_Int + 1 + HscYasumi.Value * 3) <= UBound(m_YasumiMark) Then
                            If m_YasumiMark(w_Int + 1 + HscYasumi.Value * 2).ClickFlg = True Then
                                w_Font = m_lstCmdYasumi(w_Int).Font
                                m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
                                m_lstCmdYasumi(w_Int).Checked = True
                                w_LoopFlg = True
                                Exit For
                            End If
                        End If
                    End If
                Next w_Int
            End If

            If w_LoopFlg = False Then
                For w_Int = 0 To 4
                    If w_Int <= UBound(m_TokushuMark) - 1 Then
                        If w_Int <= UBound(m_TokushuMark) - 1 And (w_Int + 1 + HscTokushu.Value) <= UBound(m_TokushuMark) Then
                            If m_TokushuMark(w_Int + 1 + HscTokushu.Value).ClickFlg = True Then
                                w_Font = m_lstCmdTokushu(w_Int).Font
                                m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
                                m_lstCmdTokushu(w_Int).Checked = True
                                w_LoopFlg = True
                                Exit For
                            End If
                        End If
                    End If
                Next w_Int
            End If

            If w_LoopFlg = False Then
                For w_Int = 0 To 4
                    If w_Int <= UBound(m_SetKinmuMark) - 1 Then
                        If w_Int <= UBound(m_SetKinmuMark) - 1 And (w_Int + 1 + HscSet.Value) <= UBound(m_SetKinmuMark) Then
                            If m_SetKinmuMark(w_Int + 1 + HscSet.Value).ClickFlg = True Then
                                w_Font = m_lstCmdSet(w_Int).Font
                                m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
                                m_lstCmdSet(w_Int).Checked = True
                                Exit For
                            End If
                        End If
                    End If
                Next w_Int
            End If

            m_SelNowKinmuCD = w_SelKinmuCD
            '�Ζ��L�����قɐݒ�
            If w_SelKinmuCD = "" Then
                LblSelected.Text = ""
            Else
                If CShort(w_SelKinmuCD) < 1000 Then
                    LblSelected.Text = g_KinmuM(CShort(w_SelKinmuCD)).Mark
                Else
                    LblSelected.Text = m_SetKinmuMark(CShort(w_SelKinmuCD) / 1000).Mark
                    w_SetKinmuFlg = True
                End If
            End If

            For w_Int = 0 To m_lstOptRiyu.Count - 1
                If m_lstOptRiyu(w_Int).Checked Then
                    m_SelNowRiyuKbn = w_Int + 1
                    Exit For
                End If
            Next
            If m_SelNowRiyuKbn <> "" Then
                Select Case m_SelNowRiyuKbn
                    Case "1" '�ʏ�
                        '����/�w�i�F
                        w_ForeColor = ColorTranslator.ToOle(Color.Black)
                        w_BackColor = ColorTranslator.ToOle(Color.White)
                    Case "2" '�v��
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", CStr(General.G_PALEGREEN)))
                    Case "3" '��]
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", CStr(General.G_PLUM)))
                    Case "4" '�Čf
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", CStr(General.G_LIGHTCYAN)))
                    Case "5"
                        '����/�w�i�F
                        m_SelNowRiyuKbn = "6"
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", CStr(General.G_ORANGE)))
                    Case Else
                End Select

                '�Ζ��L�����ق̐F�ݒ�
                LblSelected.ForeColor = ColorTranslator.FromOle(w_ForeColor)
                LblSelected.BackColor = ColorTranslator.FromOle(w_BackColor)
            End If
        Else
            '���݂�������Ă���ꍇ
            '���ݑI������Ă���Ζ�����R��ޔ�
            w_SelKinmuCD = m_SelNowKinmuCD '�Ζ�����
            '����Ӱ�ނɐݒ�
            m_SelNowKinmuCD = ""
            m_SelNowRiyuKbn = ""
            LblSelected.Text = ""
            lblSetKinmuNm.Text = ""
            LblSelected.ForeColor = Color.Black
            LblSelected.BackColor = Color.White
        End If

        If g_LimitedFlg = False Then
            If g_SaikeiFlg = False Then
                '���R�敪��S���g�p�ɂ���
                If w_SetKinmuFlg = True Then
                    m_lstOptRiyu(0).Checked = True
                    m_lstOptRiyu(1).Enabled = False
                    m_lstOptRiyu(2).Enabled = False
                    m_lstOptRiyu(3).Enabled = False
                    m_lstOptRiyu(4).Enabled = False
                Else
                    m_lstOptRiyu(1).Enabled = True

                    '��]�񐔐�������@���@��]��0��@�̏ꍇ

                    '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                    'If g_HopeNumFlg = "1" And g_HopeNum = 0 Then
                    If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                        '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                        m_lstOptRiyu(2).Enabled = False
                    Else
                        m_lstOptRiyu(2).Enabled = True
                    End If

                    m_lstOptRiyu(3).Enabled = True
                    m_lstOptRiyu(4).Enabled = True
                End If
            End If
        End If

        Exit Sub
CScmdErase_Click:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub m_lstCmdKinmu_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _CScmdKinmu_0.Click, _CScmdKinmu_1.Click, _CScmdKinmu_2.Click, _
                                                                                                                            _CScmdKinmu_3.Click, _CScmdKinmu_4.Click, _CScmdKinmu_5.Click, _
                                                                                                                            _CScmdKinmu_6.Click, _CScmdKinmu_7.Click, _CScmdKinmu_8.Click, _
                                                                                                                            _CScmdKinmu_9.Click, _CScmdKinmu_10.Click, _CScmdKinmu_11.Click, _
                                                                                                                            _CScmdKinmu_12.Click, _CScmdKinmu_13.Click, _CScmdKinmu_14.Click

        Dim Index As Short = m_lstCmdKinmu.IndexOf(eventSender)
        On Error GoTo m_lstCmdKinmu_Click
        Const W_SUBNAME As String = "NSK0000HB m_lstCmdKinmu_Click"

        Dim w_str As String
        Dim w_ForeColor As Integer
        Dim w_BackColor As Integer
        Dim w_Index As Short
        Dim w_Int As Short
        Dim w_RegStr As String
        Dim w_Font As Font

        'ڼ޽�؊i�[��
        w_RegStr = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

        w_Font = m_lstCmdKinmu(Index).Font
        If m_lstCmdKinmu(Index).Checked Then
            m_lstCmdKinmu(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
        Else
            m_lstCmdKinmu(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
        End If

        '�I�����ꂽ���݈ȊO�̏�Ԃ���ɖ߂��Ă����ԁi������Ă��Ȃ���ԁj��
        '�Ζ�
        For w_Int = 0 To 14
            If w_Int <> Index Then
                w_Font = m_lstCmdKinmu(w_Int).Font
                m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
                m_lstCmdKinmu(w_Int).Checked = False
            End If
        Next w_Int

        '�x��
        For w_Int = 0 To 9
            w_Font = m_lstCmdYasumi(w_Int).Font
            m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdYasumi(w_Int).Checked = False
        Next w_Int

        '����
        For w_Int = 0 To 4
            w_Font = m_lstCmdTokushu(w_Int).Font
            m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdTokushu(w_Int).Checked = False
        Next w_Int

        '�Z�b�g�Ζ�
        For w_Int = 0 To 4
            w_Font = m_lstCmdSet(w_Int).Font
            m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdSet(w_Int).Checked = False
        Next w_Int

        '�Ζ��L�� �擾
        w_str = m_KinmuMark(Index + 1 + HscKinmu.Value * 3).Mark

        '2016/2/22 okamura add st --------------
        '���R�敪���Z�b�g����(�Ζ����݂���ɖ߂��Ă���Ƃ������s)
        For w_Int = 0 To m_lstOptRiyu.Count - 1
            If m_lstOptRiyu(w_Int).Checked Then
                m_SelNowRiyuKbn = w_Int + 1
                Exit For
            End If
        Next
        If m_SelNowRiyuKbn <> "" Then
            If m_SelNowRiyuKbn = "5" Then
                '�Ζ��̑I���̏ꍇ�͗��R�敪 "6"
                m_SelNowRiyuKbn = "6"
            End If
        End If
        '---------------------------------------

        '�Ζ��L�����قɐݒ�
        LblSelected.Text = w_str

        '���ʕϐ� �ޔ�
        'KinmuCD
        m_SelNowKinmuCD = m_KinmuMark(Index + 1 + HscKinmu.Value * 3).CD

        '���ׂĂ̋Ζ��L���z���ClickFlg��False��
        For w_Int = 1 To UBound(m_KinmuMark)
            m_KinmuMark(w_Int).ClickFlg = False
        Next w_Int

        For w_Int = 1 To UBound(m_YasumiMark)
            m_YasumiMark(w_Int).ClickFlg = False
        Next w_Int

        For w_Int = 1 To UBound(m_TokushuMark)
            m_TokushuMark(w_Int).ClickFlg = False
        Next w_Int

        For w_Int = 1 To UBound(m_SetKinmuMark)
            m_SetKinmuMark(w_Int).ClickFlg = False
        Next w_Int

        '�I�����ꂽ�Ζ��L����Ture��
        m_KinmuMark(Index + 1 + HscKinmu.Value * 3).ClickFlg = True

        '�������݂�������Ă���Ƃ�
        If CScmdErase.Checked Then
            For w_Int = 0 To m_lstOptRiyu.Count - 1
                If m_lstOptRiyu(w_Int).Checked Then
                    m_SelNowRiyuKbn = w_Int + 1
                    Exit For
                End If
            Next
            If m_SelNowRiyuKbn <> "" Then
                Select Case m_SelNowRiyuKbn
                    Case "1" '�ʏ�
                        '����/�w�i�F
                        w_ForeColor = ColorTranslator.ToOle(Color.Black)
                        w_BackColor = ColorTranslator.ToOle(Color.White)
                    Case "2" '�v��
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", CStr(General.G_PALEGREEN)))
                    Case "3" '��]
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", CStr(General.G_PLUM)))
                    Case "4" '�Čf
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", CStr(General.G_LIGHTCYAN)))
                    Case "5"
                        '����/�w�i�F
                        m_SelNowRiyuKbn = "6"
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", CStr(General.G_ORANGE)))
                    Case Else
                End Select

                '�Ζ��L�����ق̐F�ݒ�
                LblSelected.ForeColor = ColorTranslator.FromOle(w_ForeColor)
                LblSelected.BackColor = ColorTranslator.FromOle(w_BackColor)
            End If

            '�������݂�������Ă��Ȃ���Ԃ�
            CScmdErase.Checked = False
        End If

        m_BtnClickFlg = True
        '���ׂĂ̋Ζ����݂���ɖ߂��Ă���Ƃ�()
        If m_lstCmdKinmu(Index).Checked = False Then
            If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_PgmFlg = General.G_PGMSTARTFLG_CHANGEPLAN Then
                '�Ζ��ύX�̂Ƃ��\��t�����폜���s��Ȃ�
                m_BtnClickFlg = False
                m_SelNowKinmuCD = ""
                m_SelNowRiyuKbn = ""
                LblSelected.Text = ""
                LblSelected.ForeColor = Color.Black
                LblSelected.BackColor = Color.White
            Else
                CScmdErase.Checked = True
                Call CScmdErase_ClickEvent(CScmdErase, New System.EventArgs())
            End If
        End If

        If m_PgmFlg = General.G_PGMSTARTFLG_NEWPLAN Then
            If g_LimitedFlg = False Then
                If g_SaikeiFlg = False Then
                    '���R�敪��S���g�p�ɂ���
                    m_lstOptRiyu(1).Enabled = True

                    '��]�񐔐�������@���@��]��0��@�̏ꍇ
                    '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                    'If g_HopeNumFlg = "1" And g_HopeNum = 0 Then
                    If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                        '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                        m_lstOptRiyu(2).Enabled = False
                    Else
                        m_lstOptRiyu(2).Enabled = True
                    End If

                    m_lstOptRiyu(3).Enabled = True
                    m_lstOptRiyu(4).Enabled = True
                End If
            End If
        End If

        Exit Sub
m_lstCmdKinmu_Click:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub m_lstCmdSet_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _CScmdSet_0.Click, _
                                                                                                                        _CScmdSet_1.Click, _
                                                                                                                        _CScmdSet_2.Click, _
                                                                                                                        _CScmdSet_3.Click, _
                                                                                                                        _CScmdSet_4.Click

        Dim Index As Short = m_lstCmdSet.IndexOf(eventSender)
        On Error GoTo m_lstCmdSet_Click
        Const W_SUBNAME As String = "NSK0000HB m_lstCmdSet_Click"

        Dim w_str As String
        Dim w_ForeColor As Integer
        Dim w_BackColor As Integer
        Dim w_Index As Short
        Dim w_Int As Short
        Dim w_RegStr As String
        Dim w_Font As Font

        'ڼ޽�؊i�[��
        w_RegStr = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

        w_Font = m_lstCmdSet(Index).Font
        If m_lstCmdSet(Index).Checked Then
            m_lstCmdSet(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
        Else
            m_lstCmdSet(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
        End If

        '�I�����ꂽ���݈ȊO�̏�Ԃ���ɖ߂��Ă����Ԃ�
        '�Ζ�
        For w_Int = 0 To 14
            w_Font = m_lstCmdKinmu(w_Int).Font
            m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdKinmu(w_Int).Checked = False
        Next w_Int

        '�x��
        For w_Int = 0 To 9
            w_Font = m_lstCmdYasumi(w_Int).Font
            m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdYasumi(w_Int).Checked = False
        Next w_Int

        '����
        For w_Int = 0 To 4
            w_Font = m_lstCmdTokushu(w_Int).Font
            m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdTokushu(w_Int).Checked = False
        Next w_Int

        '�Z�b�g�Ζ�
        For w_Int = 0 To 4
            If w_Int <> Index Then
                w_Font = m_lstCmdSet(w_Int).Font
                m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
                m_lstCmdSet(w_Int).Checked = False
            End If
        Next w_Int

        '���ׂĂ̋Ζ��L���z���ClickFlg��False��
        For w_Int = 1 To UBound(m_KinmuMark)
            m_KinmuMark(w_Int).ClickFlg = False
        Next w_Int

        For w_Int = 1 To UBound(m_YasumiMark)
            m_YasumiMark(w_Int).ClickFlg = False
        Next w_Int

        For w_Int = 1 To UBound(m_TokushuMark)
            m_TokushuMark(w_Int).ClickFlg = False
        Next w_Int

        For w_Int = 1 To UBound(m_SetKinmuMark)
            m_SetKinmuMark(w_Int).ClickFlg = False
        Next w_Int

        '�I�����ꂽ�Ζ��L����Ture��
        m_SetKinmuMark(Index + 1 + HscSet.Value).ClickFlg = True

        '�Ζ��L�� �擾
        w_str = m_SetKinmuMark(Index + 1 + HscSet.Value).Mark

        '2016/2/22 okamura add st --------------
        '���R�敪���Z�b�g����(�Ζ����݂���ɖ߂��Ă���Ƃ������s)
        For w_Int = 0 To m_lstOptRiyu.Count - 1
            If m_lstOptRiyu(w_Int).Checked Then
                m_SelNowRiyuKbn = w_Int + 1
                Exit For
            End If
        Next
        If m_SelNowRiyuKbn <> "" Then
            If m_SelNowRiyuKbn = "5" Then
                '�Ζ��̑I���̏ꍇ�͗��R�敪 "6"
                m_SelNowRiyuKbn = "6"
            End If
        End If
        '---------------------------------------

        '�Ζ��L�����قɐݒ�
        LblSelected.Text = w_str
        '���̕������ݒ�
        lblSetKinmuNm.Text = m_SetKinmuMark(Index + 1 + HscSet.Value).StrText

        '���ʕϐ� �ޔ�
        'KinmuCD(�Z�b�g�Ζ��Ȃ̂ŋΖ�CD��1000�Ƃ��Ă���)
        m_SelNowKinmuCD = CStr(1000 * (Index + 1 + HscSet.Value))

        For w_Int = 0 To m_lstOptRiyu.Count - 1
            If m_lstOptRiyu(w_Int).Checked Then
                m_SelNowRiyuKbn = w_Int + 1
                Exit For
            End If
        Next
        If m_SelNowRiyuKbn <> "" Then
            If m_SelNowRiyuKbn = "5" Then
                '�Ζ��̑I���̏ꍇ�͗��R�敪 "6"
                m_SelNowRiyuKbn = "6"
            End If
        End If

        '�������݂�������Ă���Ƃ�
        If CScmdErase.Checked Then
            If m_SelNowRiyuKbn <> "" Then
                Select Case m_SelNowRiyuKbn
                    Case "1"
                        '����/�w�i�F
                        w_ForeColor = ColorTranslator.ToOle(Color.Black)
                        w_BackColor = ColorTranslator.ToOle(Color.White)
                    Case "3"
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", CStr(General.G_PLUM)))
                    Case "4" '�Čf
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", CStr(General.G_LIGHTCYAN)))
                End Select

                '�Ζ��L�����ق̐F�ݒ�
                LblSelected.ForeColor = ColorTranslator.FromOle(w_ForeColor)
                LblSelected.BackColor = ColorTranslator.FromOle(w_BackColor)
            End If

            '�������݂�������Ă��Ȃ���Ԃ�
            CScmdErase.Checked = False
        End If

        m_BtnClickFlg = True
        '���ׂĂ̋Ζ����݂���ɖ߂��Ă���Ƃ�
        w_Font = m_lstCmdSet(Index).Font
        If m_lstCmdSet(Index).Checked = False Then
            If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_PgmFlg = General.G_PGMSTARTFLG_CHANGEPLAN Then
                '�Ζ��ύX�̂Ƃ��\��t�����폜���s��Ȃ�
                m_BtnClickFlg = False
                m_SelNowKinmuCD = ""
                m_SelNowRiyuKbn = ""
                LblSelected.Text = ""
                lblSetKinmuNm.Text = ""
                LblSelected.ForeColor = Color.Black
                LblSelected.BackColor = Color.White
            Else
                CScmdErase.Checked = True
                Call CScmdErase_ClickEvent(CScmdErase, New System.EventArgs())
            End If

            If g_LimitedFlg = False Then
                If g_SaikeiFlg = False Then
                    m_lstOptRiyu(1).Enabled = True
                    '��]�񐔐�������@���@��]��0��@�̏ꍇ
                    '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                    'If g_HopeNumFlg = "1" And g_HopeNum = 0 Then
                    If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                        '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                        m_lstOptRiyu(2).Enabled = False
                    Else
                        m_lstOptRiyu(2).Enabled = True
                    End If

                    m_lstOptRiyu(3).Enabled = True
                    m_lstOptRiyu(4).Enabled = True
                End If
            End If
        Else
            If g_LimitedFlg = False Then
                If g_SaikeiFlg = False Then
                    '���R�敪�͒ʏ�̂�
                    m_lstOptRiyu(0).Checked = True
                    m_lstOptRiyu(1).Enabled = False
                    m_lstOptRiyu(2).Enabled = False
                    m_lstOptRiyu(3).Enabled = False
                    m_lstOptRiyu(4).Enabled = False
                End If
            End If
        End If

        Exit Sub
m_lstCmdSet_Click:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub m_lstCmdTokushu_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _CScmdTokushu_0.Click, _
                                                                                                                            _CScmdTokushu_1.Click, _
                                                                                                                            _CScmdTokushu_2.Click, _
                                                                                                                            _CScmdTokushu_3.Click, _
                                                                                                                            _CScmdTokushu_4.Click

        Dim Index As Short = m_lstCmdTokushu.IndexOf(eventSender)
        On Error GoTo m_lstCmdTokushu_Click
        Const W_SUBNAME As String = "NSK0000HB m_lstCmdTokushu_Click"

        Dim w_str As String
        Dim w_ForeColor As Integer
        Dim w_BackColor As Integer
        Dim w_Index As Short
        Dim w_Int As Short
        Dim w_RegStr As String
        Dim w_Font As Font

        'ڼ޽�؊i�[��
        w_RegStr = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

        w_Font = m_lstCmdTokushu(Index).Font
        If m_lstCmdTokushu(Index).Checked Then
            m_lstCmdTokushu(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
        Else
            m_lstCmdTokushu(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
        End If

        '�I�����ꂽ���݈ȊO�̏�Ԃ���ɖ߂��Ă����Ԃ�
        '�Ζ�
        For w_Int = 0 To 14
            w_Font = m_lstCmdKinmu(w_Int).Font
            m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdKinmu(w_Int).Checked = False
        Next w_Int

        '�x��
        For w_Int = 0 To 9
            w_Font = m_lstCmdYasumi(w_Int).Font
            m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdYasumi(w_Int).Checked = False
        Next w_Int

        '����
        For w_Int = 0 To 4
            If w_Int <> Index Then
                w_Font = m_lstCmdTokushu(w_Int).Font
                m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
                m_lstCmdTokushu(w_Int).Checked = False
            End If
        Next w_Int

        '�Z�b�g�Ζ�
        For w_Int = 0 To 4
            w_Font = m_lstCmdSet(w_Int).Font
            m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdSet(w_Int).Checked = False
        Next w_Int

        '���ׂĂ̋Ζ��L���z���ClickFlg��False��
        For w_Int = 1 To UBound(m_KinmuMark)
            m_KinmuMark(w_Int).ClickFlg = False
        Next w_Int

        For w_Int = 1 To UBound(m_YasumiMark)
            m_YasumiMark(w_Int).ClickFlg = False
        Next w_Int

        For w_Int = 1 To UBound(m_TokushuMark)
            m_TokushuMark(w_Int).ClickFlg = False
        Next w_Int

        For w_Int = 1 To UBound(m_SetKinmuMark)
            m_SetKinmuMark(w_Int).ClickFlg = False
        Next w_Int

        '�I�����ꂽ�Ζ��L����Ture��
        m_TokushuMark(Index + 1 + HscTokushu.Value).ClickFlg = True

        '�Ζ��L�� �擾
        w_str = m_TokushuMark(Index + 1 + HscTokushu.Value).Mark

        '2016/2/22 okamura add st --------------
        '���R�敪���Z�b�g����(�Ζ����݂���ɖ߂��Ă���Ƃ������s)
        For w_Int = 0 To m_lstOptRiyu.Count - 1
            If m_lstOptRiyu(w_Int).Checked Then
                m_SelNowRiyuKbn = w_Int + 1
                Exit For
            End If
        Next
        If m_SelNowRiyuKbn <> "" Then
            If m_SelNowRiyuKbn = "5" Then
                '�Ζ��̑I���̏ꍇ�͗��R�敪 "6"
                m_SelNowRiyuKbn = "6"
            End If
        End If
        '---------------------------------------

        '�Ζ��L�����قɐݒ�
        LblSelected.Text = w_str

        '���ʕϐ� �ޔ�
        'KinmuCD
        m_SelNowKinmuCD = m_TokushuMark(Index + 1 + HscTokushu.Value).CD

        '�������݂�������Ă���Ƃ�
        If CScmdErase.Checked Then
            For w_Int = 0 To m_lstOptRiyu.Count - 1
                If m_lstOptRiyu(w_Int).Checked Then
                    m_SelNowRiyuKbn = w_Int + 1
                    Exit For
                End If
            Next
            If m_SelNowRiyuKbn <> "" Then
                Select Case m_SelNowRiyuKbn
                    Case "1" '�ʏ�
                        '����/�w�i�F
                        w_ForeColor = ColorTranslator.ToOle(Color.Black)
                        w_BackColor = ColorTranslator.ToOle(Color.White)
                    Case "2" '�v��
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", CStr(General.G_PALEGREEN)))
                    Case "3" '��]
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", CStr(General.G_PLUM)))
                    Case "4" '�Čf
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", CStr(General.G_LIGHTCYAN)))
                    Case "5"
                        '����/�w�i�F
                        m_SelNowRiyuKbn = "6"
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", CStr(General.G_ORANGE)))
                    Case Else
                End Select

                '�Ζ��L�����ق̐F�ݒ�
                LblSelected.ForeColor = ColorTranslator.FromOle(w_ForeColor)
                LblSelected.BackColor = ColorTranslator.FromOle(w_BackColor)
            End If

            '�������݂�������Ă��Ȃ���Ԃ�
            CScmdErase.Checked = False
        End If

        m_BtnClickFlg = True
        '���ׂĂ̋Ζ����݂���ɖ߂��Ă���Ƃ�
        If m_lstCmdTokushu(Index).Checked = False Then
            If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_PgmFlg = General.G_PGMSTARTFLG_CHANGEPLAN Then
                '�Ζ��ύX�̂Ƃ��\��t�����폜���s��Ȃ�
                m_BtnClickFlg = False
                m_SelNowKinmuCD = ""
                m_SelNowRiyuKbn = ""
                LblSelected.Text = ""
                LblSelected.ForeColor = Color.Black
                LblSelected.BackColor = Color.White
            Else
                CScmdErase.Checked = True
                Call CScmdErase_ClickEvent(CScmdErase, New System.EventArgs())
            End If
        End If

        If m_PgmFlg = General.G_PGMSTARTFLG_NEWPLAN Then
            If g_LimitedFlg = False Then
                If g_SaikeiFlg = False Then
                    '���R�敪��S���g�p�ɂ���
                    m_lstOptRiyu(1).Enabled = True

                    '��]�񐔐�������@���@��]��0��@�̏ꍇ
                    '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                    'If g_HopeNumFlg = "1" And g_HopeNum = 0 Then
                    If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                        '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                        m_lstOptRiyu(2).Enabled = False
                    Else
                        m_lstOptRiyu(2).Enabled = True
                    End If

                    m_lstOptRiyu(3).Enabled = True
                    m_lstOptRiyu(4).Enabled = True
                End If
            End If
        End If

        Exit Sub
m_lstCmdTokushu_Click:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub m_lstCmdYasumi_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _CScmdYasumi_0.Click, _CScmdYasumi_1.Click, _
                                                                                                                            _CScmdYasumi_2.Click, _CScmdYasumi_3.Click, _
                                                                                                                            _CScmdYasumi_4.Click, _CScmdYasumi_5.Click, _
                                                                                                                            _CScmdYasumi_6.Click, _CScmdYasumi_7.Click, _
                                                                                                                            _CScmdYasumi_8.Click, _CScmdYasumi_9.Click

        Dim Index As Short = m_lstCmdYasumi.IndexOf(eventSender)
        On Error GoTo m_lstCmdYasumi_Click
        Const W_SUBNAME As String = "NSK0000HB m_lstCmdYasumi_Click"

        Dim w_str As String
        Dim w_ForeColor As Integer
        Dim w_BackColor As Integer
        Dim w_Index As Short
        Dim w_Int As Short
        Dim w_RegStr As String
        Dim w_Font As Font

        'ڼ޽�؊i�[��
        w_RegStr = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

        w_Font = m_lstCmdYasumi(Index).Font
        If m_lstCmdYasumi(Index).Checked Then
            m_lstCmdYasumi(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
        Else
            m_lstCmdYasumi(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
        End If

        '�I�����ꂽ���݈ȊO�̏�Ԃ���ɖ߂��Ă����Ԃ�
        '�Ζ�
        For w_Int = 0 To 14
            w_Font = m_lstCmdKinmu(w_Int).Font
            m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdKinmu(w_Int).Checked = False
        Next w_Int

        '�x��
        For w_Int = 0 To 9
            If w_Int <> Index Then
                w_Font = m_lstCmdYasumi(w_Int).Font
                m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
                m_lstCmdYasumi(w_Int).Checked = False
            End If
        Next w_Int

        '����
        For w_Int = 0 To 4
            w_Font = m_lstCmdTokushu(w_Int).Font
            m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdTokushu(w_Int).Checked = False
        Next w_Int

        '�Z�b�g�Ζ�
        For w_Int = 0 To 4
            w_Font = m_lstCmdSet(w_Int).Font
            m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdSet(w_Int).Checked = False
        Next w_Int

        '���ׂĂ̋Ζ��L���z���ClickFlg��False��
        For w_Int = 1 To UBound(m_KinmuMark)
            m_KinmuMark(w_Int).ClickFlg = False
        Next w_Int

        For w_Int = 1 To UBound(m_YasumiMark)
            m_YasumiMark(w_Int).ClickFlg = False
        Next w_Int

        For w_Int = 1 To UBound(m_TokushuMark)
            m_TokushuMark(w_Int).ClickFlg = False
        Next w_Int

        For w_Int = 1 To UBound(m_SetKinmuMark)
            m_SetKinmuMark(w_Int).ClickFlg = False
        Next w_Int

        '�I�����ꂽ�Ζ��L����Ture��
        m_YasumiMark(Index + 1 + HscYasumi.Value * 2).ClickFlg = True

        '�Ζ��L�� �擾
        w_str = m_YasumiMark(Index + 1 + HscYasumi.Value * 2).Mark

        '2016/2/22 okamura add st --------------
        '���R�敪���Z�b�g����(�Ζ����݂���ɖ߂��Ă���Ƃ������s)
        For w_Int = 0 To m_lstOptRiyu.Count - 1
            If m_lstOptRiyu(w_Int).Checked Then
                m_SelNowRiyuKbn = w_Int + 1
                Exit For
            End If
        Next
        If m_SelNowRiyuKbn <> "" Then
            If m_SelNowRiyuKbn = "5" Then
                '�Ζ��̑I���̏ꍇ�͗��R�敪 "6"
                m_SelNowRiyuKbn = "6"
            End If
        End If
        '---------------------------------------

        '�Ζ��L�����قɐݒ�
        LblSelected.Text = w_str

        '���ʕϐ� �ޔ�
        'KinmuCD
        m_SelNowKinmuCD = m_YasumiMark(Index + 1 + HscYasumi.Value).CD
        m_SelNowKinmuCD = m_YasumiMark(Index + 1 + HscYasumi.Value * 2).CD

        '�������݂�������Ă���Ƃ�
        If CScmdErase.Checked Then
            For w_Int = 0 To m_lstOptRiyu.Count - 1
                If m_lstOptRiyu(w_Int).Checked Then
                    m_SelNowRiyuKbn = w_Int + 1
                    Exit For
                End If
            Next
            If m_SelNowRiyuKbn <> "" Then
                Select Case m_SelNowRiyuKbn
                    Case "1" '�ʏ�
                        '����/�w�i�F
                        w_ForeColor = ColorTranslator.ToOle(Color.Black)
                        w_BackColor = ColorTranslator.ToOle(Color.White)
                    Case "2" '�v��
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", CStr(General.G_PALEGREEN)))
                    Case "3" '��]
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", CStr(General.G_PLUM)))
                    Case "4" '�Čf
                        '����/�w�i�F
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", CStr(General.G_LIGHTCYAN)))
                    Case "5"
                        '����/�w�i�F
                        m_SelNowRiyuKbn = "6"
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", CStr(General.G_ORANGE)))
                    Case Else
                End Select

                '�Ζ��L�����ق̐F�ݒ�
                LblSelected.ForeColor = ColorTranslator.FromOle(w_ForeColor)
                LblSelected.BackColor = ColorTranslator.FromOle(w_BackColor)
            End If
            '�������݂�������Ă��Ȃ���Ԃ�
            CScmdErase.Checked = False
        End If

        m_BtnClickFlg = True
        '���ׂĂ̋Ζ����݂���ɖ߂��Ă���Ƃ�()
        If m_lstCmdYasumi(Index).Checked = False Then
            If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_PgmFlg = General.G_PGMSTARTFLG_CHANGEPLAN Then
                '�Ζ��ύX�̂Ƃ��\��t�����폜���s��Ȃ�
                m_BtnClickFlg = False
                m_SelNowKinmuCD = ""
                m_SelNowRiyuKbn = ""
                LblSelected.Text = ""
                LblSelected.ForeColor = Color.Black
                LblSelected.BackColor = Color.White
            Else
                CScmdErase.Checked = True
                Call CScmdErase_ClickEvent(CScmdErase, New System.EventArgs())
            End If
        End If

        If m_PgmFlg = General.G_PGMSTARTFLG_NEWPLAN Then
            If g_LimitedFlg = False Then
                If g_SaikeiFlg = False Then
                    '���R�敪��S���g�p�ɂ���
                    m_lstOptRiyu(1).Enabled = True

                    '��]�񐔐�������@���@��]��0��@�̏ꍇ
                    '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                    'If g_HopeNumFlg = "1" And g_HopeNum = 0 Then
                    If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                        '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                        m_lstOptRiyu(2).Enabled = False
                    Else
                        m_lstOptRiyu(2).Enabled = True
                    End If

                    m_lstOptRiyu(3).Enabled = True
                    m_lstOptRiyu(4).Enabled = True
                End If
            End If
        End If

        Exit Sub
m_lstCmdYasumi_Click:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub frmNSK0000HB_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo Form_Load
        Const W_SUBNAME As String = "NSK0000HB Form_Activate"

        If Me.Visible = True Then
            '�ŏ�ʂɐݒ�
            Call General.paSetDialogPos(Me)
        End If

        Exit Sub
Form_Load:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Public Sub frmNSK0000HB_Load()
        On Error GoTo Form_Load
        Const W_SUBNAME As String = "NSK0000HB Form_Load"

        Dim w_Font As Font
        '2018/09/21 K.I Add Start-------------------------
        Dim w_Left As String
        Dim w_Top As String
        '2018/09/21 K.I Add End---------------------------

        Call subSetCtlList()

        '�����Ζ��敪�̕\��FLG
        m_OuenDispFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY7, "OUENDISPFLG", "1", General.g_strHospitalCD))

        '2015/04/14 Bando Add Start ========================================
        '��]���[�h���̕\���ΏۋΖ�CD
        m_DispKinmuCd = General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY15, "DISPKINMUCD", "", General.g_strHospitalCD)
        '2015/04/14 Bando Add End   ========================================

        '�ŏ�ʂɐݒ�
        Call General.paSetDialogPos(Me)

        '����޳�̕\���߼޼�݂�ݒ肷��
        '2018/09/21 K.I Upd Start------------------------------------------------------------------------------------------
        '���W�X�g���擾���폜
        'Call General.paGetWindowPositon(Me, General.G_STRMAINKEY2 & "\" & g_AppName)
        '��ʒ���
        w_Left = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - CSng(Me.Width)) / 2)
        w_Top = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - CSng(Me.Height)) / 2)
        Me.SetBounds((CSng(w_Left)), (CSng(w_Top)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
        '2018/09/21 K.I Upd End--------------------------------------------------------------------------------------------

        '{����}����
        CScmdErase.Image = Image.FromFile(g_ImagePath & G_ERASER_ICO)
        '{����}����
        CScmdClose.Image = Image.FromFile(g_ImagePath & G_CLOSE_ICO)

        '���ٳ���޳�ɋΖ��L�����
        Call Set_KinmuData(False)

        '�Ζ��L�����ق̐F�ݒ�
        LblSelected.ForeColor = Color.Black
        LblSelected.BackColor = Color.White

        '���тŎg�p����ꍇ�́A���R�敪�̓��͍͂s��Ȃ��B1:�v��A2:����
        If m_PgmFlg = General.G_PGMSTARTFLG_NEWPLAN Then
            '�v�� �̏ꍇ
            '�Čf�����̏ꍇ�́A�Čf�݂̂��g�p�ɂ���
            If g_SaikeiFlg = True Then
                m_lstOptRiyu(0).Enabled = False '�ʏ�
                m_lstOptRiyu(0).Visible = False
                m_lstOptRiyu(1).Enabled = False '�v��
                m_lstOptRiyu(2).Enabled = False '��]
                m_lstOptRiyu(3).Enabled = True '�Čf
                m_lstOptRiyu(3).Visible = True
                m_lstOptRiyu(3).Checked = True
                If m_OuenDispFlg = 0 Then
                    m_lstOptRiyu(4).Enabled = False '����
                Else
                    m_lstOptRiyu(4).Visible = False
                End If
            Else
                '�v�� �̏ꍇ
                m_lstOptRiyu(0).Enabled = True '�ʏ�
                m_lstOptRiyu(1).Enabled = True '�v��

                '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                'If g_HopeNumFlg = "1" And g_HopeNum = 0 Then
                If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                    '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                    '��]�񐔐�������@���@��]��0��@�̏ꍇ
                    m_lstOptRiyu(2).Enabled = False '��]
                Else
                    '�ȊO
                    m_lstOptRiyu(2).Enabled = True '��]
                End If

                m_lstOptRiyu(3).Enabled = True '�Čf
                m_lstOptRiyu(3).Visible = False
                m_lstOptRiyu(0).Checked = True
                If m_OuenDispFlg = 0 Then
                    m_lstOptRiyu(4).Enabled = True '����
                    m_lstOptRiyu(4).Visible = True
                Else
                    m_lstOptRiyu(4).Enabled = False '����
                    m_lstOptRiyu(4).Visible = False
                End If
            End If
        Else
            '���т̏ꍇ
            m_lstOptRiyu(0).Enabled = True '�ʏ�
            m_lstOptRiyu(0).Checked = True
            m_lstOptRiyu(1).Enabled = False '�v��
            m_lstOptRiyu(2).Enabled = False '��]
            m_lstOptRiyu(3).Enabled = False '�Čf
            m_lstOptRiyu(3).Visible = False
            '���т̏ꍇ�̓Z�b�g�g�p�s��
            _fra_4.Enabled = False
            If m_OuenDispFlg = 1 Then
                m_lstOptRiyu(4).Visible = False
            End If

        End If

        '�v��ύX�̏ꍇ�́A�����S���ƃZ�b�g�Ζ����\���ɂ���B
        If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEPLAN Then
            CScmdErase.Visible = False
            _fra_4.Visible = False
        End If

        '��]���͂̏ꍇ�͗��R�敪��]�̂ݎg�p��
        If g_LimitedFlg = True Then
            If g_SaikeiFlg = False Then
                m_lstOptRiyu(0).Enabled = False '�ʏ�
                m_lstOptRiyu(1).Enabled = False '�v��

                '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                'If g_HopeNumFlg = "1" And g_HopeNum = 0 Then
                If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                    '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                    '��]�񐔐�������@���@��]��0��@�̏ꍇ
                    m_lstOptRiyu(2).Enabled = False '��]
                    m_lstOptRiyu(2).Checked = False
                Else
                    '�ȊO
                    m_lstOptRiyu(2).Enabled = True '��]
                    m_lstOptRiyu(2).Checked = True
                End If

                m_lstOptRiyu(3).Enabled = False '�Čf
                m_lstOptRiyu(3).Visible = False
                m_lstOptRiyu(4).Enabled = False
                If m_OuenDispFlg = 1 Then
                    m_lstOptRiyu(4).Visible = False
                End If
                m_lstOptRiyu(2).Checked = True
            End If
        End If

        '��̫�Đݒ�
        If UBound(m_KinmuMark) > 0 Then
            '(�Ζ��̈�ԍŏ������݂������ꂽ��Ԃ�)
            w_Font = m_lstCmdKinmu(0).Font
            m_lstCmdKinmu(0).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
            m_lstCmdKinmu(0).Checked = True
            m_KinmuMark(1).ClickFlg = True
            LblSelected.Text = m_KinmuMark(1).Mark
            m_SelNowKinmuCD = m_KinmuMark(1).CD
            m_SelNowRiyuKbn = "1"
            m_BtnClickFlg = True
        ElseIf UBound(m_YasumiMark) > 0 Then
            '(�x�݂̈�ԍŏ������݂������ꂽ��Ԃ�)
            w_Font = m_lstCmdYasumi(0).Font
            m_lstCmdYasumi(0).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
            m_lstCmdYasumi(0).Checked = True
            m_YasumiMark(1).ClickFlg = True
            LblSelected.Text = m_YasumiMark(1).Mark
            m_SelNowKinmuCD = m_YasumiMark(1).CD
            m_SelNowRiyuKbn = "1"
            m_BtnClickFlg = True
        ElseIf UBound(m_TokushuMark) > 0 Then
            '(����̈�ԍŏ������݂������ꂽ��Ԃ�)
            w_Font = m_lstCmdTokushu(0).Font
            m_lstCmdTokushu(0).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
            m_lstCmdTokushu(0).Checked = True
            m_TokushuMark(1).ClickFlg = True
            LblSelected.Text = m_TokushuMark(1).Mark
            m_SelNowKinmuCD = m_TokushuMark(1).CD
            m_SelNowRiyuKbn = "1"
            m_BtnClickFlg = True
        End If

        If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Then
            '�Ζ��ύX�̂Ƃ�����������ݎg�p�s��
            CScmdErase.Enabled = False
            CScmdErase.Visible = False
        End If

        If g_LimitedFlg = True Then
            m_SelNowRiyuKbn = "3" '��]�̗��R�敪
        End If

        If g_SaikeiFlg = True Then
            '�Čf�����̏ꍇ�A���R�敪��"�Čf"��
            m_SelNowRiyuKbn = "4"
        End If

        '2014/04/23 Shimizu add start P-06979-----------------------------------
        '���ڐݒ�̎擾
        m_strKinmuEmSecondFlg = Get_ItemValue(General.g_strHospitalCD)
        '�Ζ��L���S�p�Q�����Ή��̃��C�A�E�g�ύX
        Call SetKinmuSecondView()
        '2014/04/23 Shimizu add end P-06979-------------------------------------

        Exit Sub
Form_Load:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    '���ٳ���޳�ɋΖ��L�����
    Public Sub Set_KinmuData(ByVal p_CallMainFlg As Boolean)
        On Error GoTo Set_KinmuData
        Const W_SUBNAME As String = "NSK0000HB Set_KinmuData"

        Dim w_Int As Short
        Dim w_KinmuCnt As Short
        Dim w_YasumiCnt As Short
        Dim w_TokushuCnt As Short
        Dim w_RecCnt As Short
        Dim w_Sql As String
        Dim w_Rs As ADODB.Recordset
        Dim w_SetKinmuCnt As Integer
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
        Dim w_Int2 As Integer
        Dim w_strKinmuBunruiCD As String
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

        '�Ζ��ް��i�[�z�񏉊���
        ReDim m_KinmuMark(0)
        ReDim m_YasumiMark(0)
        ReDim m_TokushuMark(0)
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
                w_RecCnt = .fKN_KinmuCount

                For w_Int = 1 To w_RecCnt

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
                                    w_KinmuCnt = w_KinmuCnt + 1
                                    ReDim Preserve m_KinmuMark(w_KinmuCnt)

                                    m_KinmuMark(w_KinmuCnt).CD = .fKN_KinmuCD
                                    m_KinmuMark(w_KinmuCnt).KinmuName = .fKN_Name
                                    m_KinmuMark(w_KinmuCnt).Mark = .fKN_MarkF
                                    m_KinmuMark(w_KinmuCnt).KBunruiCD = w_strKinmuBunruiCD
                                    m_KinmuMark(w_KinmuCnt).Setumei = .fKN_KinmuExplan
                                    m_KinmuMark(w_KinmuCnt).ClickFlg = False
                                End If
                            Else
                                w_KinmuCnt = w_KinmuCnt + 1
                                ReDim Preserve m_KinmuMark(w_KinmuCnt)

                                m_KinmuMark(w_KinmuCnt).CD = .fKN_KinmuCD
                                m_KinmuMark(w_KinmuCnt).KinmuName = .fKN_Name
                                m_KinmuMark(w_KinmuCnt).Mark = .fKN_MarkF
                                m_KinmuMark(w_KinmuCnt).KBunruiCD = w_strKinmuBunruiCD
                                m_KinmuMark(w_KinmuCnt).Setumei = .fKN_KinmuExplan
                                m_KinmuMark(w_KinmuCnt).ClickFlg = False
                            End If

                            '2015/04/14 Bando Upd End   ============================
                        ElseIf w_strKinmuBunruiCD = "2" Then
                            '-- �x�� --
                            '2015/04/14 Bando Upd Start ============================
                            '��]���[�h�̏ꍇ�A�\���ΏۋΖ��̂݃p���b�g�ɕ\��
                            'If g_HopeMode = 1 Then
                            If g_HopeMode = 1 AndAlso m_DispKinmuCd <> "" Then '2015/06/02 Bando Chg
                                If InStr(m_DispKinmuCd, .fKN_KinmuCD) > 0 Then
                                    w_YasumiCnt = w_YasumiCnt + 1
                                    ReDim Preserve m_YasumiMark(w_YasumiCnt)

                                    m_YasumiMark(w_YasumiCnt).CD = .fKN_KinmuCD
                                    m_YasumiMark(w_YasumiCnt).KinmuName = .fKN_Name
                                    m_YasumiMark(w_YasumiCnt).Mark = .fKN_MarkF
                                    m_YasumiMark(w_YasumiCnt).KBunruiCD = w_strKinmuBunruiCD
                                    m_YasumiMark(w_YasumiCnt).Setumei = .fKN_KinmuExplan
                                    m_YasumiMark(w_YasumiCnt).ClickFlg = False
                                End If
                            Else
                                w_YasumiCnt = w_YasumiCnt + 1
                                ReDim Preserve m_YasumiMark(w_YasumiCnt)

                                m_YasumiMark(w_YasumiCnt).CD = .fKN_KinmuCD
                                m_YasumiMark(w_YasumiCnt).KinmuName = .fKN_Name
                                m_YasumiMark(w_YasumiCnt).Mark = .fKN_MarkF
                                m_YasumiMark(w_YasumiCnt).KBunruiCD = w_strKinmuBunruiCD
                                m_YasumiMark(w_YasumiCnt).Setumei = .fKN_KinmuExplan
                                m_YasumiMark(w_YasumiCnt).ClickFlg = False
                            End If
                            '2015/04/14 Bando Upd End   ============================
                        ElseIf w_strKinmuBunruiCD = "3" Then
                            '-- ���� --
                            '2015/04/14 Bando Upd Start ============================
                            '��]���[�h�̏ꍇ�A�\���ΏۋΖ��̂݃p���b�g�ɕ\��
                            'If g_HopeMode = 1 Then
                            If g_HopeMode = 1 AndAlso m_DispKinmuCd <> "" Then '2015/06/02 Bando Chg
                                If InStr(m_DispKinmuCd, .fKN_KinmuCD) > 0 Then
                                    w_TokushuCnt = w_TokushuCnt + 1
                                    ReDim Preserve m_TokushuMark(w_TokushuCnt)

                                    m_TokushuMark(w_TokushuCnt).CD = .fKN_KinmuCD
                                    m_TokushuMark(w_TokushuCnt).KinmuName = .fKN_Name
                                    m_TokushuMark(w_TokushuCnt).Mark = .fKN_MarkF
                                    m_TokushuMark(w_TokushuCnt).KBunruiCD = w_strKinmuBunruiCD
                                    m_TokushuMark(w_TokushuCnt).Setumei = .fKN_KinmuExplan
                                    m_TokushuMark(w_TokushuCnt).ClickFlg = False
                                End If
                            Else
                                w_TokushuCnt = w_TokushuCnt + 1
                                ReDim Preserve m_TokushuMark(w_TokushuCnt)

                                m_TokushuMark(w_TokushuCnt).CD = .fKN_KinmuCD
                                m_TokushuMark(w_TokushuCnt).KinmuName = .fKN_Name
                                m_TokushuMark(w_TokushuCnt).Mark = .fKN_MarkF
                                m_TokushuMark(w_TokushuCnt).KBunruiCD = w_strKinmuBunruiCD
                                m_TokushuMark(w_TokushuCnt).Setumei = .fKN_KinmuExplan
                                m_TokushuMark(w_TokushuCnt).ClickFlg = False
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

        '�Ζ��i�p���b�g�ɋΖ��}�[�N��ݒ肷��j
        For w_Int = 0 To 14
            If w_Int <= w_KinmuCnt - 1 Then
                m_lstCmdKinmu(w_Int).Text = m_KinmuMark(w_Int + 1).Mark
                If m_KinmuMark(w_Int + 1).Setumei = "" Then
                    ToolTip1.SetToolTip(m_lstCmdKinmu(w_Int), Get_KinmuTipText(m_KinmuMark(w_Int + 1).CD))
                Else
                    ToolTip1.SetToolTip(m_lstCmdKinmu(w_Int), Get_KinmuTipText(m_KinmuMark(w_Int + 1).CD) & "�F" & m_KinmuMark(w_Int + 1).Setumei)
                End If
            Else
                Exit For
            End If
        Next w_Int

        '�x��
        For w_Int = 0 To 9
            If w_Int <= w_YasumiCnt - 1 Then
                m_lstCmdYasumi(w_Int).Text = m_YasumiMark(w_Int + 1).Mark
                If m_YasumiMark(w_Int + 1).Setumei = "" Then
                    ToolTip1.SetToolTip(m_lstCmdYasumi(w_Int), Get_KinmuTipText(m_YasumiMark(w_Int + 1).CD))
                Else
                    ToolTip1.SetToolTip(m_lstCmdYasumi(w_Int), Get_KinmuTipText(m_YasumiMark(w_Int + 1).CD) & "�F" & m_YasumiMark(w_Int + 1).Setumei)
                End If
            Else
                Exit For
            End If
        Next w_Int

        '����Ζ�
        For w_Int = 0 To 4
            If w_Int <= w_TokushuCnt - 1 Then
                m_lstCmdTokushu(w_Int).Text = m_TokushuMark(w_Int + 1).Mark
                If m_TokushuMark(w_Int + 1).Setumei = "" Then
                    ToolTip1.SetToolTip(m_lstCmdTokushu(w_Int), Get_KinmuTipText(m_TokushuMark(w_Int + 1).CD))
                Else
                    ToolTip1.SetToolTip(m_lstCmdTokushu(w_Int), Get_KinmuTipText(m_TokushuMark(w_Int + 1).CD) & "�F" & m_TokushuMark(w_Int + 1).Setumei)
                End If
            Else
                Exit For
            End If
        Next w_Int

        '��۰��ް�A��߼�����݂̐ݒ�
        '�Ζ�
        Select Case w_KinmuCnt
            Case 0
                For w_Int = 0 To 14
                    m_lstCmdKinmu(w_Int).Visible = False
                Next w_Int

                HscKinmu.Maximum = (0 + HscKinmu.LargeChange - 1)
                HscKinmu.Visible = False
            Case 1 To 14
                For w_Int = 14 To w_KinmuCnt Step -1
                    m_lstCmdKinmu(w_Int).Visible = False
                Next w_Int

                HscKinmu.Maximum = (0 + HscKinmu.LargeChange - 1)
                HscKinmu.Visible = False
            Case 15
                For w_Int = 0 To 14
                    m_lstCmdKinmu(w_Int).Visible = True
                Next w_Int

                HscKinmu.Maximum = (0 + HscKinmu.LargeChange - 1)
                HscKinmu.Visible = False
            Case Else
                For w_Int = 0 To 14
                    m_lstCmdKinmu(w_Int).Visible = True
                    m_lstCmdKinmu(w_Int).Enabled = True
                Next w_Int

                HscKinmu.Maximum = (((w_KinmuCnt - 15) \ 3) + IIf((w_KinmuCnt - 15) Mod 3 = 0, 0, 1) + HscKinmu.LargeChange - 1)
                HscKinmu.Visible = True
                HscKinmu.Enabled = True
        End Select

        '�x��
        Select Case w_YasumiCnt
            Case 0
                For w_Int = 0 To 9
                    m_lstCmdYasumi(w_Int).Visible = False
                Next w_Int

                HscYasumi.Maximum = (0 + HscYasumi.LargeChange - 1)
                HscYasumi.Visible = False
            Case 1 To 9
                For w_Int = 9 To w_YasumiCnt Step -1
                    m_lstCmdYasumi(w_Int).Visible = False
                Next w_Int

                HscYasumi.Maximum = (0 + HscYasumi.LargeChange - 1)
                HscYasumi.Visible = False
            Case 10
                For w_Int = 0 To 9
                    m_lstCmdYasumi(w_Int).Visible = True
                    m_lstCmdYasumi(w_Int).Enabled = True
                Next w_Int

                HscYasumi.Maximum = (0 + HscYasumi.LargeChange - 1)
                HscYasumi.Visible = False
            Case Else
                For w_Int = 0 To 9
                    m_lstCmdYasumi(w_Int).Visible = True
                    m_lstCmdYasumi(w_Int).Enabled = True
                Next w_Int

                HscYasumi.Maximum = (Int((w_YasumiCnt - 10) / 2 + 0.5) + HscYasumi.LargeChange - 1)
                HscYasumi.Visible = True
                HscYasumi.Enabled = True
        End Select

        '����Ζ�
        Select Case w_TokushuCnt
            Case 0
                For w_Int = 0 To 4
                    m_lstCmdTokushu(w_Int).Visible = False
                Next w_Int

                HscTokushu.Maximum = (0 + HscTokushu.LargeChange - 1)
                HscTokushu.Visible = False
            Case 1 To 4
                For w_Int = 4 To w_TokushuCnt Step -1
                    m_lstCmdTokushu(w_Int).Visible = False
                Next w_Int

                HscTokushu.Maximum = (0 + HscTokushu.LargeChange - 1)
                HscTokushu.Visible = False
            Case 5
                For w_Int = 0 To 4
                    m_lstCmdTokushu(w_Int).Visible = True
                    m_lstCmdTokushu(w_Int).Enabled = True
                Next w_Int

                HscTokushu.Maximum = (0 + HscTokushu.LargeChange - 1)
                HscTokushu.Visible = False
            Case Else
                For w_Int = 0 To 4
                    m_lstCmdTokushu(w_Int).Visible = True
                    m_lstCmdTokushu(w_Int).Enabled = True
                Next w_Int

                HscTokushu.Maximum = (w_TokushuCnt - 5 + HscTokushu.LargeChange - 1)
                HscTokushu.Visible = True
                HscTokushu.Enabled = True
        End Select

        '2015/7/6 okamura add st ----
        '�Z�b�g�Ζ��z�񏉊���
        ReDim m_SetKinmuMark(0)
        '----------------------------

        '2015/06/02 Bando Upd Start ==========================
        If g_HopeMode <> 1 Then
            '�Z�b�g�Ζ�
            '2017/05/22 Richard Upd Start
            ''SQL���ҏW
            'w_Sql = "SELECT * FROM NS_SETKINMU_M "
            'w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
            'w_Sql = w_Sql & "AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
            'w_Sql = w_Sql & "ORDER BY DISPNO "

            'w_Rs = General.paDBRecordSetOpen(w_Sql)
            '<1>
            Call NSK0000H_sql.select_NS_SETKINMU_M_01(w_Rs)
            'Upd End
            '2015/7/6 okamura del st ----
            ''�Z�b�g�Ζ��z�񏉊���
            'ReDim m_SetKinmuMark(0)
            '----------------------------

            If w_Rs.RecordCount <= 0 Then
            Else
                w_Int3 = 1

                With w_Rs
                    .MoveLast()
                    w_RecCnt = .RecordCount
                    .MoveFirst()

                    ReDim m_SetKinmuMark(w_RecCnt)
                    w_SetKinmuCnt = w_RecCnt

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

                    For w_Int = 1 To w_RecCnt
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
                                    w_SetKinmuCnt = w_SetKinmuCnt - 1
                                    Exit For
                                Case w_strKinmuCD2
                                    w_blnEndDate = False
                                    w_SetKinmuCnt = w_SetKinmuCnt - 1
                                    Exit For
                                Case w_strKinmuCD3
                                    w_blnEndDate = False
                                    w_SetKinmuCnt = w_SetKinmuCnt - 1
                                    Exit For
                                Case w_strKinmuCD4
                                    w_blnEndDate = False
                                    w_SetKinmuCnt = w_SetKinmuCnt - 1
                                    Exit For
                                Case w_strKinmuCD5
                                    w_blnEndDate = False
                                    w_SetKinmuCnt = w_SetKinmuCnt - 1
                                    Exit For
                                Case w_strKinmuCD6
                                    w_blnEndDate = False
                                    w_SetKinmuCnt = w_SetKinmuCnt - 1
                                    Exit For
                                Case w_strKinmuCD7
                                    w_blnEndDate = False
                                    w_SetKinmuCnt = w_SetKinmuCnt - 1
                                    Exit For
                                Case w_strKinmuCD8
                                    w_blnEndDate = False
                                    w_SetKinmuCnt = w_SetKinmuCnt - 1
                                    Exit For
                                Case w_strKinmuCD9
                                    w_blnEndDate = False
                                    w_SetKinmuCnt = w_SetKinmuCnt - 1
                                    Exit For
                                Case w_strKinmuCD10
                                    w_blnEndDate = False
                                    w_SetKinmuCnt = w_SetKinmuCnt - 1
                                    Exit For
                            End Select
                        Next w_Int2
                        If w_blnEndDate = True Then
                            m_SetKinmuMark(w_Int3).Initialize()

                            m_SetKinmuMark(w_Int3).Mark = w_�L��_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(1) = w_�Ζ�CD1_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(2) = w_�Ζ�CD2_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(3) = w_�Ζ�CD3_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(4) = w_�Ζ�CD4_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(5) = w_�Ζ�CD5_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(6) = w_�Ζ�CD6_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(7) = w_�Ζ�CD7_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(8) = w_�Ζ�CD8_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(9) = w_�Ζ�CD9_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(10) = w_�Ζ�CD10_F.Value & ""
                            m_SetKinmuMark(w_Int3).ClickFlg = False
                            m_SetKinmuMark(w_Int3).blnKinmu = True

                            '�Ζ����������邩(�Ԃɋ󔒂͂Ȃ����̂Ƃ���)
                            w_KinmuCnt = 0
                            For w_Int2 = 1 To 10
                                If m_SetKinmuMark(w_Int3).CD(w_Int2) <> "" Then
                                    w_KinmuCnt = w_KinmuCnt + 1
                                Else
                                    Exit For
                                End If
                            Next w_Int2

                            m_SetKinmuMark(w_Int3).KinmuCnt = w_KinmuCnt
                            w_Int3 = w_Int3 + 1
                        End If

                        .MoveNext()
                    Next w_Int
                End With
            End If

            w_Rs.Close()
        End If

        For w_Int = 0 To 4
            If w_Int <= w_SetKinmuCnt - 1 Then
                If m_SetKinmuMark(w_Int + 1).blnKinmu = True Then
                    m_lstCmdSet(w_Int).Text = m_SetKinmuMark(w_Int + 1).Mark
                    ToolTip1.SetToolTip(m_lstCmdSet(w_Int), Get_SetKinmuTipText(w_Int + 1))
                End If
            Else
                Exit For
            End If
        Next w_Int

        Select Case w_SetKinmuCnt
            Case 0
                For w_Int = 0 To 4
                    m_lstCmdSet(w_Int).Visible = False
                Next w_Int

                HscSet.Maximum = (0 + HscSet.LargeChange - 1)
                HscSet.Visible = False
            Case 1 To 4
                '�S��Visible=True�ɂ���
                For w_Int = 0 To 4
                    m_lstCmdSet(w_Int).Visible = True
                Next w_Int

                For w_Int = 4 To w_SetKinmuCnt Step -1
                    m_lstCmdSet(w_Int).Visible = False
                Next w_Int

                HscSet.Maximum = (0 + HscSet.LargeChange - 1)
                HscSet.Visible = False
            Case 5
                For w_Int = 0 To 4
                    m_lstCmdSet(w_Int).Visible = True
                    m_lstCmdSet(w_Int).Enabled = True
                Next w_Int

                HscSet.Maximum = (0 + HscSet.LargeChange - 1)
                HscSet.Visible = False
            Case Else
                For w_Int = 0 To 4
                    m_lstCmdSet(w_Int).Visible = True
                    m_lstCmdSet(w_Int).Enabled = True
                Next w_Int

                HscSet.Maximum = (w_SetKinmuCnt - 5 + HscSet.LargeChange - 1)
                HscSet.Visible = True
                HscSet.Enabled = True
        End Select

        '���C����ʂ���Ă΂ꂽ(�Z�b�g�Ζ����X�V����)�ꍇ�͑I���{�^����������
        If p_CallMainFlg = True And m_SelNowKinmuCD <> "" Then
            If Integer.Parse(m_SelNowKinmuCD) >= 1000 Then
                '���ݑI���Ζ����Z�b�g�Ζ��̏ꍇ
                If UBound(m_SetKinmuMark) > 0 Then
                    HscSet.Value = 0
                    m_lstCmdSet(0).Checked = False
                    m_BtnClickFlg = True
                    Call m_lstCmdSet_ClickEvent(m_lstCmdSet.Item(0), New System.EventArgs())
                Else
                    '���ݑI���Ζ����Z�b�g�Ζ��ȊO�̏ꍇ
                    lblSetKinmuNm.Text = ""
                    '��̫�Đݒ�
                    If UBound(m_KinmuMark) > 0 Then
                        '(�Ζ��̈�ԍŏ������݂������ꂽ��Ԃ�)
                        m_lstCmdKinmu(0).Checked = False
                        m_SelNowRiyuKbn = "1"
                        m_BtnClickFlg = True
                        Call m_lstCmdKinmu_ClickEvent(m_lstCmdKinmu.Item(0), New System.EventArgs())
                    ElseIf UBound(m_YasumiMark) > 0 Then
                        '(�x�݂̈�ԍŏ������݂������ꂽ��Ԃ�)
                        m_lstCmdYasumi(0).Checked = False
                        m_SelNowRiyuKbn = "1"
                        m_BtnClickFlg = True
                        Call m_lstCmdYasumi_ClickEvent(m_lstCmdYasumi.Item(0), New System.EventArgs())
                    ElseIf UBound(m_TokushuMark) > 0 Then
                        '(����̈�ԍŏ������݂������ꂽ��Ԃ�)
                        m_lstCmdTokushu(0).Checked = False
                        m_SelNowRiyuKbn = "1"
                        m_BtnClickFlg = True
                        Call m_lstCmdTokushu_ClickEvent(m_lstCmdTokushu.Item(0), New System.EventArgs())
                    Else
                        m_SelNowKinmuCD = ""
                        m_SelNowRiyuKbn = ""
                        LblSelected.Text = ""
                        lblSetKinmuNm.Text = ""
                        LblSelected.ForeColor = Color.Black
                        LblSelected.BackColor = Color.White
                        m_BtnClickFlg = False
                    End If
                End If
            End If
        End If


        Exit Sub
Set_KinmuData:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    '�Z�b�g�Ζ��c�[���`�b�v�p������擾
    Public Function Get_SetKinmuTipText(ByVal p_Int As Integer) As String
        On Error GoTo Get_SetKinmuTipText
        Const W_SUBNAME As String = "NSK0000HB Get_SetKinmuTipText"

        Dim w_str As String
        Dim w_strTEXT As String
        Dim w_Cnt As Integer
        Dim w_CD As String

        For w_Cnt = 1 To 10
            '�Ζ�CD���擾
            w_CD = m_SetKinmuMark(p_Int).CD(w_Cnt)

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

        m_SetKinmuMark(p_Int).StrText = w_strTEXT

        Get_SetKinmuTipText = w_str

        Exit Function
Get_SetKinmuTipText:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Function

    Private Sub frmNSK0000HB_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As FormClosingEventArgs) Handles Me.FormClosing
        Dim UnloadMode As CloseReason = eventArgs.CloseReason
        On Error GoTo Form_QueryUnload
        Const W_SUBNAME As String = "NSK0000HB Form_QueryUnload"

        If UnloadMode = CloseReason.UserClosing Then
            eventArgs.Cancel = True
            Me.Hide()

            RaiseEvent KensakuEnabled()
        End If

        Exit Sub
Form_QueryUnload:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Public Sub frmNSK0000HB_FormClosed()
        On Error GoTo Form_Unload
        Const W_SUBNAME As String = "NSK0000HB Form_Unload"

        '����޳�̕\���߼޼�݂��i�[����
        Call General.paPutWindowPositon(Me, General.G_STRMAINKEY2 & "\" & g_AppName)

        Exit Sub
Form_Unload:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub HscKinmu_Change(ByVal newScrollValue As Integer)
        On Error GoTo HscKinmu_Change
        Const W_SUBNAME As String = "NSK0000HB HscKinmu_Change"

        Dim w_Int As Short
        Dim w_Cnt As Short
        Dim w_Hsc_Cnt As Short
        Dim w_Font As Font

        '��������݂�Caption�ݒ�
        '�Ζ�
        w_Hsc_Cnt = newScrollValue

        For w_Int = 0 To 14
            '���݂̏�Ԃ���ɖ߂��Ă����Ԃ�
            w_Font = m_lstCmdKinmu(w_Int).Font
            m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdKinmu(w_Int).Checked = False
            w_Cnt = w_Int + 1 + w_Hsc_Cnt * 3

            If w_Cnt <= UBound(m_KinmuMark) Then

                m_lstCmdKinmu(w_Int).Text = m_KinmuMark(w_Int + 1 + w_Hsc_Cnt * 3).Mark
                If m_KinmuMark(w_Int + 1 + w_Hsc_Cnt * 3).Setumei = "" Then
                    ToolTip1.SetToolTip(m_lstCmdKinmu(w_Int), Get_KinmuTipText(m_KinmuMark(w_Int + 1 + w_Hsc_Cnt * 3).CD))
                Else
                    ToolTip1.SetToolTip(m_lstCmdKinmu(w_Int), Get_KinmuTipText(m_KinmuMark(w_Int + 1 + w_Hsc_Cnt * 3).CD) & "�F" & m_KinmuMark(w_Int + 1 + w_Hsc_Cnt * 3).Setumei)
                End If

                If CScmdErase.Checked = False Then
                    If m_KinmuMark(w_Int + 1 + w_Hsc_Cnt * 3).ClickFlg = True Then
                        '���݂�د����ꂽ��Ԃ�
                        w_Font = m_lstCmdKinmu(w_Int).Font
                        m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
                        m_lstCmdKinmu(w_Int).Checked = True
                    End If
                End If

                m_lstCmdKinmu(w_Int).Visible = True
                m_lstCmdKinmu(w_Int).Enabled = True
            Else
                m_lstCmdKinmu(w_Int).Visible = False
                m_lstCmdKinmu(w_Int).Enabled = False
            End If
        Next w_Int

        Exit Sub
HscKinmu_Change:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub HscYasumi_Change(ByVal newScrollValue As Integer)
        On Error GoTo HscYasumi_Change
        Const W_SUBNAME As String = "NSK0000HB HscYasumi_Change"

        Dim w_Int As Short
        Dim w_Cnt As Short
        Dim w_Hsc_Cnt As Short
        Dim w_Font As Font

        '��������݂�Caption�ݒ�
        '�x��
        w_Hsc_Cnt = newScrollValue

        For w_Int = 0 To 9
            '���݂̏�Ԃ���ɖ߂��Ă����Ԃ�

            w_Font = m_lstCmdYasumi(w_Int).Font
            m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdYasumi(w_Int).Checked = False
            w_Cnt = w_Int + 1 + w_Hsc_Cnt * 2

            If w_Cnt <= UBound(m_YasumiMark) And (w_Int + 1 + w_Hsc_Cnt * 2) <= UBound(m_YasumiMark) Then
                m_lstCmdYasumi(w_Int).Text = m_YasumiMark(w_Int + 1 + w_Hsc_Cnt * 2).Mark
                If m_YasumiMark(w_Int + 1 + w_Hsc_Cnt * 2).Setumei = "" Then
                    ToolTip1.SetToolTip(m_lstCmdYasumi(w_Int), Get_KinmuTipText(m_YasumiMark(w_Int + 1 + w_Hsc_Cnt * 2).CD))
                Else
                    ToolTip1.SetToolTip(m_lstCmdYasumi(w_Int), Get_KinmuTipText(m_YasumiMark(w_Int + 1 + w_Hsc_Cnt * 2).CD) & "�F" & m_YasumiMark(w_Int + 1 + w_Hsc_Cnt * 2).Setumei)
                End If

                If CScmdErase.Checked = False Then
                    If m_YasumiMark(w_Int + 1 + w_Hsc_Cnt * 2).ClickFlg = True Then
                        '���݂�د����ꂽ��Ԃ�
                        w_Font = m_lstCmdYasumi(w_Int).Font
                        m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
                        m_lstCmdYasumi(w_Int).Checked = True
                    End If
                End If

                m_lstCmdYasumi(w_Int).Visible = True
                m_lstCmdYasumi(w_Int).Enabled = True
            Else
                m_lstCmdYasumi(w_Int).Visible = False
                m_lstCmdYasumi(w_Int).Enabled = False
            End If
        Next w_Int

        Exit Sub
HscYasumi_Change:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub HscSet_Change(ByVal newScrollValue As Integer)
        On Error GoTo HscSet_Change
        Const W_SUBNAME As String = "NSK0000HB HscSet_Change"

        Dim w_Int As Short
        Dim w_Hsc_Cnt As Short
        Dim w_Font As Font

        '��������݂�Caption�ݒ�
        '�Z�b�g�Ζ�
        w_Hsc_Cnt = newScrollValue
        For w_Int = 0 To 4

            If UBound(m_SetKinmuMark) >= w_Int + 1 Then
                '���݂̏�Ԃ���ɖ߂��Ă����Ԃ�
                w_Font = m_lstCmdSet(w_Int).Font
                m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
                m_lstCmdSet(w_Int).Checked = False

                m_lstCmdSet(w_Int).Text = m_SetKinmuMark(w_Int + 1 + w_Hsc_Cnt).Mark
                ToolTip1.SetToolTip(m_lstCmdSet(w_Int), Get_SetKinmuTipText(w_Int + 1 + w_Hsc_Cnt))

                If CScmdErase.Checked = False Then
                    If m_SetKinmuMark(w_Int + 1 + w_Hsc_Cnt).ClickFlg = True Then
                        '���݂�د����ꂽ��Ԃ�
                        w_Font = m_lstCmdSet(w_Int).Font
                        m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
                        m_lstCmdSet(w_Int).Checked = True
                    End If
                End If
            End If
        Next w_Int

        Exit Sub
HscSet_Change:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub HscTokushu_Change(ByVal newScrollValue As Integer)
        On Error GoTo HscTokushu_Change
        Const W_SUBNAME As String = "NSK0000HB HscTokushu_Change"

        Dim w_Int As Short
        Dim w_Hsc_Cnt As Short
        Dim w_Font As Font

        '��������݂�Caption�ݒ�
        '����Ζ�
        w_Hsc_Cnt = newScrollValue
        For w_Int = 0 To 4
            '���݂̏�Ԃ���ɖ߂��Ă����Ԃ�
            w_Font = m_lstCmdTokushu(w_Int).Font
            m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdTokushu(w_Int).Checked = False

            m_lstCmdTokushu(w_Int).Text = m_TokushuMark(w_Int + 1 + w_Hsc_Cnt).Mark
            If m_TokushuMark(w_Int + 1 + w_Hsc_Cnt).Setumei = "" Then
                ToolTip1.SetToolTip(m_lstCmdTokushu(w_Int), Get_KinmuTipText(m_TokushuMark(w_Int + 1 + w_Hsc_Cnt).CD))
            Else
                ToolTip1.SetToolTip(m_lstCmdTokushu(w_Int), Get_KinmuTipText(m_TokushuMark(w_Int + 1 + w_Hsc_Cnt).CD) & "�F" & m_TokushuMark(w_Int + 1 + w_Hsc_Cnt).Setumei)
            End If

            If CScmdErase.Checked = False Then
                If m_TokushuMark(w_Int + 1 + w_Hsc_Cnt).ClickFlg = True Then
                    '���݂�د����ꂽ��Ԃ�
                    w_Font = m_lstCmdTokushu(w_Int).Font
                    m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
                    m_lstCmdTokushu(w_Int).Checked = True
                End If
            End If
        Next w_Int

        Exit Sub
HscTokushu_Change:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub m_lstOptRiyu_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _OptRiyu_0.CheckedChanged, _
                                                                                                                            _OptRiyu_1.CheckedChanged, _
                                                                                                                            _OptRiyu_2.CheckedChanged, _
                                                                                                                            _OptRiyu_3.CheckedChanged, _
                                                                                                                            _OptRiyu_4.CheckedChanged

        If eventSender.Checked Then
            Dim Index As Short = m_lstOptRiyu.IndexOf(eventSender)
            On Error GoTo m_lstOptRiyu_Click
            Const W_SUBNAME As String = "NSK0000HB m_lstOptRiyu_Click"

            Dim w_Index As Short
            Dim w_str As String
            Dim w_ForeColor As Integer
            Dim w_BackColor As Integer
            Dim w_RegStr As String
            Dim w_Font As Font

            'ڼ޽�؊i�[��
            w_RegStr = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            '�������݂�������Ă���Ƃ��ͤ�F�̕ύX�����Ȃ�
            w_Font = CScmdErase.Font
            If w_Font.Bold = False Then

                w_Index = Index
                '�I�v�V�����{�^���̃`�F�b�N
                If w_Index <> True Then
                    '���ʕϐ� �ޔ�
                    '���R�敪�i�ʏ�,�v��,��]�j
                    m_SelNowRiyuKbn = CStr(w_Index + 1)

                    '�Ζ��L�����ق̐F�ݒ�
                    '���R�敪 ?
                    Select Case m_SelNowRiyuKbn
                        Case "1" '�ʏ�
                            '����/�w�i�F
                            w_ForeColor = ColorTranslator.ToOle(Color.Black)
                            w_BackColor = ColorTranslator.ToOle(Color.White)
                        Case "2" '�v��
                            '����/�w�i�F
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", CStr(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", CStr(General.G_PALEGREEN)))
                        Case "3" '��]
                            '����/�w�i�F
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", CStr(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", CStr(General.G_PLUM)))
                        Case "4" '�Čf
                            '����/�w�i�F
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", CStr(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", CStr(General.G_LIGHTCYAN)))
                        Case "5" '����
                            m_SelNowRiyuKbn = "6"
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", CStr(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", CStr(General.G_ORANGE)))
                        Case Else
                    End Select

                    '�Ζ��L�����ق̐F�ݒ�
                    LblSelected.ForeColor = ColorTranslator.FromOle(w_ForeColor)
                    LblSelected.BackColor = ColorTranslator.FromOle(w_BackColor)
                End If
            End If

            Exit Sub
m_lstOptRiyu_Click:
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End If
    End Sub

    Private Sub HscKinmu_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscKinmu.Scroll
        HscKinmu_Change(eventArgs.NewValue)
    End Sub

    Private Sub HscYasumi_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscYasumi.Scroll
        HscYasumi_Change(eventArgs.NewValue)
    End Sub

    Private Sub HscSet_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscSet.Scroll
        HscSet_Change(eventArgs.NewValue)
    End Sub

    Private Sub HscTokushu_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscTokushu.Scroll
        HscTokushu_Change(eventArgs.NewValue)
    End Sub

    '�R���g���[���z��̑���Ƀ��X�g�Ɋi�[����
    Private Sub subSetCtlList()
        m_lstOptRiyu.Add(_OptRiyu_0)
        m_lstOptRiyu.Add(_OptRiyu_1)
        m_lstOptRiyu.Add(_OptRiyu_2)
        m_lstOptRiyu.Add(_OptRiyu_3)
        m_lstOptRiyu.Add(_OptRiyu_4)

        m_lstCmdKinmu.Add(_CScmdKinmu_0)
        m_lstCmdKinmu.Add(_CScmdKinmu_1)
        m_lstCmdKinmu.Add(_CScmdKinmu_2)
        m_lstCmdKinmu.Add(_CScmdKinmu_3)
        m_lstCmdKinmu.Add(_CScmdKinmu_4)
        m_lstCmdKinmu.Add(_CScmdKinmu_5)
        m_lstCmdKinmu.Add(_CScmdKinmu_6)
        m_lstCmdKinmu.Add(_CScmdKinmu_7)
        m_lstCmdKinmu.Add(_CScmdKinmu_8)
        m_lstCmdKinmu.Add(_CScmdKinmu_9)
        m_lstCmdKinmu.Add(_CScmdKinmu_10)
        m_lstCmdKinmu.Add(_CScmdKinmu_11)
        m_lstCmdKinmu.Add(_CScmdKinmu_12)
        m_lstCmdKinmu.Add(_CScmdKinmu_13)
        m_lstCmdKinmu.Add(_CScmdKinmu_14)

        m_lstCmdYasumi.Add(_CScmdYasumi_0)
        m_lstCmdYasumi.Add(_CScmdYasumi_1)
        m_lstCmdYasumi.Add(_CScmdYasumi_2)
        m_lstCmdYasumi.Add(_CScmdYasumi_3)
        m_lstCmdYasumi.Add(_CScmdYasumi_4)
        m_lstCmdYasumi.Add(_CScmdYasumi_5)
        m_lstCmdYasumi.Add(_CScmdYasumi_6)
        m_lstCmdYasumi.Add(_CScmdYasumi_7)
        m_lstCmdYasumi.Add(_CScmdYasumi_8)
        m_lstCmdYasumi.Add(_CScmdYasumi_9)

        m_lstCmdTokushu.Add(_CScmdTokushu_0)
        m_lstCmdTokushu.Add(_CScmdTokushu_1)
        m_lstCmdTokushu.Add(_CScmdTokushu_2)
        m_lstCmdTokushu.Add(_CScmdTokushu_3)
        m_lstCmdTokushu.Add(_CScmdTokushu_4)

        m_lstCmdSet.Add(_CScmdSet_0)
        m_lstCmdSet.Add(_CScmdSet_1)
        m_lstCmdSet.Add(_CScmdSet_2)
        m_lstCmdSet.Add(_CScmdSet_3)
        m_lstCmdSet.Add(_CScmdSet_4)
    End Sub

    '2014/04/23 Shimizu add start P-06979--------------------------------------------------------------------------------------------------
    '/----------------------------------------------------------------------/
    '/  �T�v�@�@�@�@  : �Ζ��L���S�p�Q�����Ή��̃��C�A�E�g�ύX
    '/  �p�����[�^    : �Ȃ�
    '/  �߂�l        : �Ȃ�
    '/----------------------------------------------------------------------/
    Private Sub SetKinmuSecondView()
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc '��ۼ��ެ���̈ꎞ�Ҕ�
        General.g_ErrorProc = "NSK0000HB SetKinmuSecondView"

        Const W_FRAME_FIRST_HEIGHT As Integer = 16 '1�s�ڂ̏c�ʒu
        Const W_FRAME_FIRST_WIDTH As Integer = 8 '1��ڂ̉��ʒu
        Const W_FRAME_ADD_HEIGHT As Integer = 24 '�s�̏c�ʒu������
        Const W_FRAME_ADD_WIDTH As Integer = 39 '�s�̉��ʒu������

        Const W_KINMU_HEIGHT As Integer = 25 '�t���[���̏c��
        Const W_FRAME_WIDTH As Integer = 213 '�t���[���̉���
        Const W_SCL_WIDTH As Integer = 196 '�X�N���[���̉���
        Const W_SCL_HEIGHT As Integer = 17 '�X�N���[���̏c��
        Const W_KINMU_WIDTH As Integer = 40 '�Ζ��̉���

        Const W_SET_HEIGHT_ADJUST As Integer = 30 '�Z�b�g�Ζ��̍�������

        Try
            '�Ζ��L���S�p�Q�����Ή��t���O����
            If m_strKinmuEmSecondFlg = "0" Then
                '0�F�Ή����Ȃ�(�]���̋Ζ��L�����̓T�C�Y�ƍő�2�o�C�g)

            Else
                '1�F�Ή�����(�S�p�Q�������\���ł���Ζ��L�����̓T�C�Y�ƍő�4�o�C�g)
                '�t�H�[��
                Me.Size = New System.Drawing.Size(330, 435)

                '�L��
                _fra_3.Location = New Point(240, 2)
                LblSelected.Location = New Point(10, 17)
                _fra_3.Size = New System.Drawing.Size(80, 47)
                LblSelected.Size = New System.Drawing.Size(61, 24)
                '�敪
                PnlRiyu.Location = New Point(240, 54)
                CScmdErase.Location = New Point(240, 155)
                CScmdClose.Location = New Point(240, 205)

                '�p���b�g���ڂ��Ă���p�l��
                SSPanel2.Size = New System.Drawing.Size(230, 387)

                '�t���[��
                _fra_0.Size = New System.Drawing.Size(W_FRAME_WIDTH, 113)
                _fra_1.Size = New System.Drawing.Size(W_FRAME_WIDTH, 89)
                _fra_2.Size = New System.Drawing.Size(W_FRAME_WIDTH, 65)
                _fra_4.Size = New System.Drawing.Size(W_FRAME_WIDTH, 95)

                '�X�N���[��
                HscKinmu.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)
                HscYasumi.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)
                HscTokushu.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)
                HscSet.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)


                '�Ζ�
                General.setSizeAndLocal(m_lstCmdKinmu, 3, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

                '�x��
                General.setSizeAndLocal(m_lstCmdYasumi, 2, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

                '����Ζ�
                General.setSizeAndLocal(m_lstCmdTokushu, 1, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

                '�Z�b�g
                lblSetKinmuNm.Width = 180
                lblSetKinmuNm.Height = 30
                CType(HscSet, System.Windows.Forms.Control).Location = New Point(W_FRAME_FIRST_WIDTH, 70)
                CType(m_lstCmdSet(0), System.Windows.Forms.Control).Location = New Point(W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT + W_SET_HEIGHT_ADJUST)
                General.setSizeAndLocal(m_lstCmdSet, 1, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT + W_SET_HEIGHT_ADJUST, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT + W_SET_HEIGHT_ADJUST)

            End If

            General.g_ErrorProc = w_PreErrorProc '�Ҕ���ۼ��ެ�������ɖ߂�

        Catch ex As Exception
            Err.Raise(Err.Number)
        End Try
    End Sub
    '2014/04/23 Shimizu add end P-06979----------------------------------------------------------------------------------------------------
End Class