Option Strict Off
Option Explicit On
Friend Class frmNSK0000HD
    Inherits General.FormBase
    Public NSK0000H_sql As New clsNSK0000H_sql
    '/ ﾌﾟﾛｸﾞﾗﾑ名称：計画作成/実績登録メニュー
    '/        ＩＤ：NSK0000HD
    '/        概要：検索/置換画面
    '/
    '/
    '/      作成者： S.Y    CREATE 2000/08/14           REV 01.00
    '/      更新者：        UPDATE     /  /             REV 01.01
    '/
    '/     Copyright (C) Inter co.,ltd 2000
    '/----------------------------------------------------------------------/
    '----- ﾒﾆｭｰﾊﾞｰ番号 定義 --------------------------------------------------
    'ﾒﾆｭｰﾊﾞｰ配列のインデックスを表す定数
    '[表示]メニュー配列のインデックスを表す定数
    Private Const M_MenuPalette As Short = 0 'ﾊﾟﾚｯﾄ
    'ﾂｰﾙﾊﾞｰのKey定数
    Private Const M_ToolBarKey_Palette As String = "Palette" 'パレット

    '-----------------------------------------------------------------
    '   変 数 宣 言
    '-----------------------------------------------------------------
    Private m_PgmFlg As String '起動ﾓｰﾄﾞ
    Private m_KenChiFlg As BasNSK0000H.geKenChiFlg '検索 Or 置換（列挙型で宣言）
    Private m_KinmuCD() As String 'KinmuCD
    Private m_KensakuKinmuCD As String '検索勤務CD
    Private m_KensakuTaisyo As Short '対象
    Private m_TaisyoOnly As Short '対象
    Private m_ChikanKinmuCD As String '置換勤務CD
    Private m_RiyuKBN As String '置換後の理由区分
    Private m_FormShowFlg As Boolean '画面が表示されているか
    Private m_TargetKinmuCD() As String
    Private m_CmbHT As New Hashtable

    '2014/04/23 Saijo add start P-06979-----------------------------------------------------------------------
    Private m_strKinmuEmSecondFlg As String '勤務記号全角２文字対応フラグ(0：対応しない、1:対応する)
    '2014/04/23 Saijo add end P-06979-------------------------------------------------------------------------

    '------------------------------------------------------------------
    '  ｲﾍﾞﾝﾄ宣言
    '------------------------------------------------------------------
	Event PaletteEnabled()
	Event SelectKinmu()
	Event ChikanKinmu()
	Event SelectAllKinme()
	Event ChikanAllKinmu()
    Event SetEditMenu()
	
	'-- 勤務記号 退避配列 ----------------------
	Private Structure Kinmu_Type
		Dim CD As String 'KinmuCD
		Dim KinmuName As String '名称
		Dim Mark As String '記号
		Dim HolBunruiCD As String '休み分類CD
        Dim EffToDate As Integer '有効終了日
    End Structure

	Private m_KinmuM() As Kinmu_Type '勤務情報配列
	
    Private m_StartDate As Integer
	
    '開始日取得
	Public WriteOnly Property pStartDate() As Integer
		Set(ByVal Value As Integer)
            m_StartDate = Value
		End Set
    End Property

    '検索/置換 ﾌﾗｸﾞ を取得する
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
            '検索勤務ｺｰﾄﾞ
			pKensakuKinmuCD = m_KensakuKinmuCD
        End Get
	End Property
	
	Public ReadOnly Property pRiyuKBN() As String
		Get
            '理由区分
			pRiyuKBN = m_RiyuKBN
        End Get
	End Property
	
	Public ReadOnly Property pKensakuTaisyo() As String
		Get
            '対象
			pKensakuTaisyo = CStr(m_KensakuTaisyo)
        End Get
	End Property

	Public ReadOnly Property pTaisyoOnly() As Short
		Get
            '対象
			pTaisyoOnly = m_TaisyoOnly
        End Get
	End Property
	
	Public ReadOnly Property pChikanKinmuCD() As String
		Get
            '置換勤務ｺｰﾄﾞ
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

        '対象が"すべて"の時、"対象のみで検索する"ﾁｪｯｸﾎﾞｯｸｽを使用不可に
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
		
        '検索/置換ﾃﾞｰﾀを取得
		Call Get_KenChiData()
		
		'検索･置換実行
		If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Kensaku Then
			
			'検索実行
			RaiseEvent SelectKinmu()
			
		Else
			
			'置換実行
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
		
		'検索勤務ｺｰﾄﾞ
        w_Index = GetItemData(cboKensakuKinmu, cboKensakuKinmu.SelectedIndex)
		If w_Index > 0 Then
			m_KensakuKinmuCD = m_KinmuCD(w_Index)
		Else
			m_KensakuKinmuCD = "000"
		End If
		
		'対象
		w_Index = cboTaisyo.SelectedIndex
		If w_Index >= 0 Then
			'委員会勤務の理由区分が"5"のため、応援勤務の区分を"6"にする
			If w_Index = 5 Then
				m_KensakuTaisyo = CShort(CStr(6))
			Else
				m_KensakuTaisyo = CShort(CStr(w_Index))
			End If
		Else
			m_KensakuTaisyo = CShort("0")
		End If
		
		'対象のみのﾁｪｯｸ
		m_TaisyoOnly = chkTaisyoOnly.CheckState
		
		'置換の場合
		If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Chikan Then
			
			'置換勤務ｺｰﾄﾞ
            w_Index = GetItemData(cboChikanKinmu, cboChikanKinmu.SelectedIndex)
			If w_Index > 0 Then
				m_ChikanKinmuCD = m_KinmuCD(w_Index)
			Else
				m_ChikanKinmuCD = "000"
            End If

			'置換後の理由区分
            If GetItemData(cboRiyu, cboRiyu.SelectedIndex) = 2 Then
                '要請
                m_RiyuKBN = "2"
            ElseIf GetItemData(cboRiyu, cboRiyu.SelectedIndex) = 3 Then
                '希望
                m_RiyuKBN = "3"
            ElseIf GetItemData(cboRiyu, cboRiyu.SelectedIndex) = 4 Then
                '再掲
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
		
        '検索/置換ﾃﾞｰﾀを取得
		Call Get_KenChiData()
		
		'検索･置換実行
		If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Kensaku Then
			
			'非表示にする
			Me.Hide()
			m_FormShowFlg = False
			'検索(選択)実行
			RaiseEvent SelectAllKinme()
			
			'パレット使用可
			RaiseEvent PaletteEnabled()
			
            '切り取り、コピー、貼り付けの制御
			RaiseEvent SetEditMenu()
		Else
			
			'非表示にする
			Me.Hide()
			m_FormShowFlg = False
			'置換(すべて)実行
			RaiseEvent ChikanAllKinmu()
			
			'パレット使用可
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
		
		'パレット使用可
		RaiseEvent PaletteEnabled()
		
        '切り取り、コピー、貼り付けの制御
		If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Kensaku Then
			RaiseEvent SetEditMenu()
		End If
		
		'画面をHideする
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
            '最上位に設定
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


        'ﾌｫｰﾑ ｱｲｺﾝ 設定
        If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Chikan Then
            Me.Icon = New Icon(g_ImagePath & G_PERMUTATION_ICO)
        Else
            Me.Icon = New Icon(g_ImagePath & G_SEARCH_ICO)
        End If

        'ウィンドゥを画面の最上位に設定
        Call General.paSetDialogPos(Me)

        'ｺﾝﾎﾞﾎﾞｯｸｽ設定
        '表示対象勤務CDを取得
        Call Get_TargetKinmuCD()
        Call Set_ComboBox()
        '2014/04/23 Saijo add start P-06979------------------------------------
        '項目設定の取得
        m_strKinmuEmSecondFlg = Get_ItemValue(General.g_strHospitalCD)
        '勤務記号全角２文字対応のレイアウト変更
        Call SetKinmuSecondView()
        '2014/04/23 Saijo add end P-06979--------------------------------------

        'ビットマップイメージ保存パス取得
        w_SystemPath = My.Application.Info.DirectoryPath
        w_ImagePath = General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, "ImagePath", w_SystemPath & "image\")

        cmd_Next_Chikan.Text = ""
        cmd_Select_AllChikan.Text = ""

        '初期設定
        If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Kensaku Then
            '検索のとき
            _lblKomoku_1.Enabled = False
            _lblKomoku_1.Visible = False
            cboChikanKinmu.Enabled = False
            cboChikanKinmu.Visible = False
            _lblKomoku_3.Enabled = False
            _lblKomoku_3.Visible = False
            cboRiyu.Enabled = False
            cboRiyu.Visible = False
            cmd_Next_Chikan.Image = Image.FromFile(w_ImagePath & "次を検索.bmp")
            cmd_Select_AllChikan.Image = Image.FromFile(w_ImagePath & "選択.bmp")
            Me.Text = "検索"
            If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Then
                '勤務変更の場合、対象を使用不可にする
                'ﾃﾞﾌｫﾙﾄで対象は"すべて"とする
                _lblKomoku_2.Enabled = False
                cboTaisyo.Enabled = False
            End If
        Else
            '置換のとき
            _lblKomoku_1.Enabled = True
            _lblKomoku_1.Visible = True
            cboChikanKinmu.Enabled = True
            cboChikanKinmu.Visible = True
            _lblKomoku_3.Enabled = True
            _lblKomoku_3.Visible = True
            cboRiyu.Enabled = True
            cboRiyu.Visible = True
            cmd_Next_Chikan.Image = Image.FromFile(w_ImagePath & "置換実行.bmp")
            cmd_Select_AllChikan.Image = Image.FromFile(w_ImagePath & "全て置換.bmp")
            Me.Text = "置換"
            If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Then
                '勤務変更の場合、対象と置換後の理由区分を使用不可にする
                'ﾃﾞﾌｫﾙﾄで対象は"すべて"とする
                '置換後の理由区分は通常とする
                _lblKomoku_2.Enabled = False
                cboTaisyo.Enabled = False
                _lblKomoku_3.Enabled = False
                cboRiyu.Enabled = False
            End If
        End If

        'ｳｨﾝﾄﾞｳの表示ﾎﾟｼﾞｼｮﾝを設定する
        '2018/09/21 K.I Upd Start------------------------------------------------------------------------------------------
        'レジストリ取得を削除
        'Call General.paGetWindowPositon(Me, General.G_STRMAINKEY2 & "\" & g_AppName)
        '画面中央
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
		Dim w_記号_F As ADODB.Field
		Dim w_名称_F As ADODB.Field
		Dim w_休み分類CD_F As ADODB.Field
        Dim w_有効終了日_F As ADODB.Field
        Dim w_lngLoop As Integer
		Dim w_lngDataIdx As Integer
        Dim w_CmbCnt As Short
        Dim w_CmbCnt2 As Short
        '2017/05/02 Christopher Upd Start
        'Select文編集
        'w_Sql = "SELECT   KINMUCD "
        'w_Sql = w_Sql & ",MARKF "
        'w_Sql = w_Sql & ",NAME "
        'w_Sql = w_Sql & ",HOLIDAYBUNRUICD "
        '      w_Sql = w_Sql & ",EFFTODATE "
        'w_Sql = w_Sql & "FROM NS_KINMUNAME_M "
        '      w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
        'w_Sql = w_Sql & "ORDER BY DISPNO "
        ''RecordSet ｵﾌﾞｼﾞｪｸﾄの生成
        '      w_Rs = General.paDBRecordSetOpen(w_Sql)

        Call NSK0000H_sql.select_NS_KINMUNAME_M_02(w_Rs)
        'Upd End
        If w_Rs.RecordCount <= 0 Then
            'ﾃﾞｰﾀがないとき
            w_Rs.Close()
            Exit Sub
        Else
            With w_Rs
                'ﾃﾞｰﾀがあるとき

                'ﾃﾞｰﾀ件数 取得
                .MoveLast()
                w_RecCnt = .RecordCount
                .MoveFirst()

                'ﾌｨｰﾙﾄﾞ ｵﾌﾞｼﾞｪｸﾄ 作成
                w_KinmuCD_F = .Fields("KINMUCD")
                w_記号_F = .Fields("MARKF")
                w_名称_F = .Fields("NAME")
                w_休み分類CD_F = .Fields("HOLIDAYBUNRUICD")
                w_有効終了日_F = .Fields("EFFTODATE")
				
                For w_Int = 1 To w_RecCnt
                    w_str = w_KinmuCD_F.Value
                    '表示対象勤務CDかチェック
                    For w_lngLoop = 1 To UBound(m_TargetKinmuCD)
                        If w_str = m_TargetKinmuCD(w_lngLoop) Then
                            w_lngDataIdx = w_lngDataIdx + 1
                            '配列確保
                            ReDim Preserve m_KinmuM(w_lngDataIdx)

                            m_KinmuM(w_lngDataIdx).CD = w_str
                            m_KinmuM(w_lngDataIdx).KinmuName = IIf(IsDBNull(w_名称_F.Value), "", w_名称_F.Value)
                            m_KinmuM(w_lngDataIdx).Mark = IIf(IsDBNull(w_記号_F.Value), "", w_記号_F.Value)
                            m_KinmuM(w_lngDataIdx).HolBunruiCD = IIf(IsDBNull(w_休み分類CD_F.Value), "", w_休み分類CD_F.Value)
                            m_KinmuM(w_lngDataIdx).EffToDate = IIf(IsDBNull(w_有効終了日_F.Value), 0, w_有効終了日_F.Value) '-20080909-okamoto-Add

                            Exit For
                        End If
                    Next w_lngLoop

                    .MoveNext()
                Next w_Int
            End With
		End If
		w_Rs.Close()
		
		'--- ｺﾝﾎﾞﾎﾞｯｸｽの設定 ---
		cboKensakuKinmu.Items.Clear()
		cboChikanKinmu.Items.Clear()
		'ListIndex=0は未割当
        '検索用勤務ｺﾝﾎﾞ
        cboKensakuKinmu.Items.Add("未割当")
        SetItemData(cboKensakuKinmu, 0, 0)
		
        If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Then
            '勤務変更で起動されているときは未割当への置換は行わない
        Else
            '置換用勤務ｺﾝﾎﾞ
            cboChikanKinmu.Items.Add("未割当")
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
					
					'検索用勤務ｺﾝﾎﾞ
                    cboKensakuKinmu.Items.Add(w_str)
                    w_CmbCnt = w_CmbCnt + 1
                    SetItemData(cboKensakuKinmu, w_CmbCnt, w_Cnt)
                    
                    If General.g_lngDaikyuMng = 0 Then
                        If m_KinmuM(w_Int).HolBunruiCD <> General.G_STRDAIKYUBUNRUI Then '休み分類CDが代休のCDは対象外
                            '置換用勤務ｺﾝﾎﾞ
                            cboChikanKinmu.Items.Add(w_str)
                            w_CmbCnt2 = w_CmbCnt2 + 1
                            SetItemData(cboChikanKinmu, w_CmbCnt2, w_Cnt)
                        End If
                    Else
                        '置換用勤務ｺﾝﾎﾞ
                        cboChikanKinmu.Items.Add(w_str)
                        SetItemData(cboChikanKinmu, w_CmbCnt, w_Cnt)
                    End If
                End If
			End If
		Next w_Int
		
		'ﾃﾞﾌｫﾙﾄ値設定
		cboKensakuKinmu.SelectedIndex = 1
        If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Then
            cboChikanKinmu.SelectedIndex = 0
        Else
            cboChikanKinmu.SelectedIndex = 1
        End If
		
		'置換後の理由区分ｺﾝﾎﾞﾎﾞｯｸｽ設定
		With cboRiyu
            .Items.Clear()
			If g_SaikeiFlg = True Then
                .Items.Add("再掲")
                SetItemData(cboRiyu, 0, 4)
			Else
				.Items.Add("通常")
                .Items.Add("要請")
                .Items.Add("希望")
                SetItemData(cboRiyu, 0, 1)
                SetItemData(cboRiyu, 1, 2)
                SetItemData(cboRiyu, 2, 3)
			End If
			.SelectedIndex = 0
        End With
		
		'対象ｺﾝﾎﾞﾎﾞｯｸｽ設定
		With cboTaisyo
            .Items.Clear()
            .Items.Add("すべて")
            .Items.Add("通常")
            .Items.Add("要請")
            .Items.Add("希望")
            .Items.Add("再掲")
            SetItemData(cboTaisyo, 0, 0)
            SetItemData(cboTaisyo, 1, 1)
            SetItemData(cboTaisyo, 2, 2)
            SetItemData(cboTaisyo, 3, 3)
            SetItemData(cboTaisyo, 4, 4)
			If m_KenChiFlg = BasNSK0000H.geKenChiFlg.FormType_Kensaku Then
                .Items.Add("応援")
                SetItemData(cboTaisyo, 5, 5)
			End If
			
			'対象ｺﾝﾎﾞﾎﾞｯｸｽのﾃﾞﾌｫﾙﾄ値設定
			.SelectedIndex = 0
        End With
		
		'対象のみの検索ﾁｪｯｸﾎﾞｯｸｽのﾃﾞﾌｫﾙﾄ値設定
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
            'ﾊﾟﾚｯﾄ使用可
            RaiseEvent PaletteEnabled()

            'ｺﾝﾄﾛｰﾙﾒﾆｭｰから閉じられた場合はUnloadしない
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

        'ｳｨﾝﾄﾞｳの表示ﾎﾟｼﾞｼｮﾝを格納する
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

        '初期化
        ReDim m_TargetKinmuCD(0)
        '2017/05/02 Christopher Upd Start
        'Select文編集
        'w_Sql = "SELECT   KINMUCD "
        'w_Sql = w_Sql & "FROM NS_SETKINMUNAME_F "
        'w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
        'w_Sql = w_Sql & "AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
        'w_Sql = w_Sql & "ORDER BY DISPNO "
        ''RecordSet ｵﾌﾞｼﾞｪｸﾄの生成
        'w_Rs = General.paDBRecordSetOpen(w_Sql)

        Call NSK0000H_sql.select_NS_SETKINMUNAME_F_01(w_Rs)
        'Upd End
        If w_Rs.RecordCount <= 0 Then
            'ﾃﾞｰﾀがないとき
            w_Rs.Close()
            Exit Sub
        Else
            With w_Rs
                'ﾃﾞｰﾀがあるとき

                'ﾃﾞｰﾀ件数 取得
                .MoveLast()
                w_RecCnt = .RecordCount
                .MoveFirst()

                'ﾌｨｰﾙﾄﾞ ｵﾌﾞｼﾞｪｸﾄ 作成
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
    '/  概要　　　　  : 勤務記号全角２文字対応のレイアウト変更
    '/  パラメータ    : なし
    '/  戻り値        : なし
    '/----------------------------------------------------------------------/
    Private Sub SetKinmuSecondView()
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc 'ﾌﾟﾛｼｰｼﾞｬ名の一時待避
        General.g_ErrorProc = "frmNSC1000HE SetKinmuSecondView"

        Try
            '勤務記号全角２文字対応フラグ判定
            If m_strKinmuEmSecondFlg = "0" Then
                '0：対応しない(従来の勤務記号入力サイズと最大2バイト)
                cboKensakuKinmu.Size = New System.Drawing.Size(94, 23)
                cboChikanKinmu.Size = New System.Drawing.Size(94, 23)
            Else
                '1：対応する(全角２文字が表示できる勤務記号入力サイズと最大4バイト)
                cboKensakuKinmu.Size = New System.Drawing.Size(112, 23)
                cboChikanKinmu.Size = New System.Drawing.Size(112, 23)
            End If

            General.g_ErrorProc = w_PreErrorProc '待避ﾌﾟﾛｼｰｼﾞｬ名を元に戻す

        Catch ex As Exception
            Err.Raise(Err.Number)
        End Try
    End Sub
    '2014/04/23 Saijo add end  P-06979----------------------------------------------------------------------------------------------------
End Class