Option Strict Off
Option Explicit On
Friend Class frmNSK0000HK
    Inherits General.FormBase
	
	Private m_KangoM() As KangoM_Type
	Private m_KangoCDFlg As Boolean
	Private m_SelKangoTCD As String '�I�����ꂽ�Ζ�����CD
	Private m_OKButtonFlg As Boolean
	
	Private Structure KangoM_Type
		Dim CD As String
        Dim Name_Renamed As String
    End Structure

	'�Ζ�����CD�擾�׸�
	Public ReadOnly Property pKangoCDFlg() As String
		Get
			pKangoCDFlg = CStr(m_KangoCDFlg)
		End Get
	End Property
	
	'���݉����׸�
	Public ReadOnly Property pOKFlg() As Boolean
		Get
			pOKFlg = m_OKButtonFlg
		End Get
	End Property
	
	'�I�����ꂽ�Ζ�����CD�擾�׸�
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

        '�z�񏉊���
        ReDim m_KangoM(0)

        With General.g_objGetMaster
            .pHospitalCD = General.g_strHospitalCD '�{�݃R�[�h
            .pKD_KinmuDeptCD = "" '��(�S��)
            .pKD_KijunDate = w_YYYYMMDD '���

            If .mGetKinmuDept = False Then
                '�ް����Ȃ��Ƃ�
            Else
                '�Ζ����������擾
                w_DataCnt = .fKD_KinmuDeptCount

                For w_Int = 1 To w_DataCnt
                    '���ޯ�����n��
                    .mKD_KinmuDeptIdx = w_Int
                    '�Ζ�����CD�擾
                    w_CD = .fKD_KinmuDeptCD

                    '�������Ζ�����CD�ȊO�̂Ƃ��͔z��Ɋi�[
                    If w_CD <> General.g_strSelKinmuDeptCD Then
                        '�z��g��
                        w_Index = UBound(m_KangoM) + 1
                        ReDim Preserve m_KangoM(w_Index)

                        '���Ζ�����CD
                        m_KangoM(w_Index).CD = w_CD
                        '�����̎擾
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
		
        '�I�����ꂽ�Ζ��n�̲��ޯ���擾
		w_SelectIndex = cboKangoTani.SelectedIndex + 1
		
		m_SelKangoTCD = m_KangoM(w_SelectIndex).CD '�I�����ꂽ�Ζ�����CD

		'----- ڼ޽�ؐݒ� ------------------
        w_StrReg = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY2
		'������Ζ��nCD
        Call General.paSaveSetting(w_StrReg, "NSK0000H", "OUENKINMUDEPTCD_" & General.g_strUserMngID, m_SelKangoTCD)
		
        'OK���݉���
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

        '�Ζ�����Ͻ��擾�׸ޏ�����
        m_KangoCDFlg = False
        'OK���ݔ��f�׸ޏ�����
        m_OKButtonFlg = False

        '������
        m_SelKangoTCD = ""

        '�Ζ������擾
        Call Get_KangoM()

        '�Ζ������l���ް����Ȃ������ꍇ
        If UBound(m_KangoM) <= 0 Then
            Exit Sub
        End If

        '�����ޯ���ɾ��
        cboKangoTani.Items.Clear()
        For w_Int = 1 To UBound(m_KangoM)
            cboKangoTani.Items.Add(m_KangoM(w_Int).Name_Renamed)
        Next w_Int

        '--- ڼ޽�ؐݒ�l �擾 ---
        '������Ζ��nCD
        w_StrReg = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY2

        w_strOuenKinmuDeptCd = General.paGetSetting(w_StrReg, "NSK0000H", "OUENKINMUDEPTCD_" & General.g_strUserMngID, "")

        w_lngOuenIndex = 0

        If w_strOuenKinmuDeptCd = "" Then
            cboKangoTani.SelectedIndex = 0
        Else
            For w_Int = 1 To UBound(m_KangoM)
                If w_strOuenKinmuDeptCd = m_KangoM(w_Int).CD Then
                    '������Ζ��nCD����v�����s��I��
                    w_lngOuenIndex = w_Int - 1
                End If
            Next w_Int

            cboKangoTani.SelectedIndex = w_lngOuenIndex
        End If

        '�ް��擾OK
        m_KangoCDFlg = True

        Exit Sub
Form_Load:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub
End Class