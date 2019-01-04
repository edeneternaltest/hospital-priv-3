Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Friend Class frmNSK0000HB
    Inherits General.FormBase
    Public NSK0000H_sql As New clsNSK0000H_sql
    '=====================================================================================
    '   定  数  宣  言
    '=====================================================================================
    '[表示]メニュー配列のインデックスを表す定数
    Private Const M_MenuPalette As Short = 0 'ﾊﾟﾚｯﾄ
    '[編集]メニュー配列のインデックスを表す定数
    Private Const M_MenuKensaku As Short = 7 '検索
    Private Const M_MenuChikan As Short = 8 '置換
    'ﾂｰﾙﾊﾞｰのKey定数
    Private Const M_ToolBarKey_Search As String = "KinmuSerach" '検索
    Private Const M_ToolBarKey_Tikan As String = "KinmuTikan" '置換

    '=====================================================================================
    '   変  数  宣  言
    '=====================================================================================
    Private m_OuenDispFlg As Integer '応援勤務区分のラジオボタンをパレットに表示するか(1:しない,0:する)
    '2015/04/14 Bando Add Start ========================
    Private m_DispKinmuCd As String '希望モード時の表示対象勤務CD
    '2015/04/14 Bando Add End   ========================
	
	Private m_PgmFlg As String '起動ﾓｰﾄﾞ
	Private m_BtnClickFlg As Boolean 'True:消しｺﾞﾑﾎﾞﾀﾝ Or 勤務ﾎﾞﾀﾝがｸﾘｯｸされているとき,False:ｸﾘｯｸされていないとき
	'現在選択 勤務記号
	Private m_SelNowKinmuCD As String 'KinmuCD
	Private m_SelNowRiyuKbn As String '理由区分
	Private m_SelSetIdx As Integer '選択されたセットのｲﾝﾃﾞｯｸｽ.
    Private m_SetCDIdx As Integer 'セットCDｲﾝﾃﾞｯｸｽ
    Private m_lstOptRiyu As New List(Of Object)
    Private m_lstCmdKinmu As New List(Of Object)
    Private m_lstCmdYasumi As New List(Of Object)
    Private m_lstCmdTokushu As New List(Of Object)
    Private m_lstCmdSet As New List(Of Object)

    '2014/04/23 Shimizu add start P-06979-------------------------------------------------------------------
    Private m_strKinmuEmSecondFlg As String '勤務記号全角２文字対応フラグ(0：対応しない、1:対応する)
    '2014/04/23 Shimizu add end P-06979---------------------------------------------------------------------

	Private Structure Kinmu_Type
		Dim CD As String 'KinmuCD
		Dim Mark As String '勤務記号
		Dim KinmuName As String 'KinmuName
		Dim KBunruiCD As String '勤務分類CD
		Dim ClickFlg As Boolean 'ﾎﾞﾀﾝの状態(True:押し込まれているとき,False:上に戻っているとき)
		Dim Setumei As String '説明
    End Structure

	Private m_KinmuMark() As Kinmu_Type
	Private m_YasumiMark() As Kinmu_Type
    Private m_TokushuMark() As Kinmu_Type

	Private Structure SetKinmu_Type
		Dim Mark As String
		<VBFixedArray(10)> Dim CD() As String
		Dim StrText As String
		Dim ClickFlg As Boolean 'ﾎﾞﾀﾝの状態(True:押し込まれているとき,False:上に戻っているとき)
		Dim KinmuCnt As Integer
        Dim blnKinmu As Boolean
		
        Public Sub Initialize()
            ReDim CD(10)
        End Sub
	End Structure
	
	Private m_SetKinmuMark() As SetKinmu_Type
    Private m_StartDate As Integer
	
	'ｲﾍﾞﾝﾄ宣言
	Event KensakuEnabled()

	'開始日取得
	Public WriteOnly Property pStartDate() As Integer
		Set(ByVal Value As Integer)
			m_StartDate = Value
		End Set
	End Property

	'選択したセットｲﾝﾃﾞｯｸｽを取得
	Public WriteOnly Property pSelKinmuIdx() As Integer
		Set(ByVal Value As Integer)
			m_SelSetIdx = Value
		End Set
    End Property

	'セットCDｲﾝﾃﾞｯｸｽ
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

	'True:消しｺﾞﾑﾎﾞﾀﾝ Or 勤務ﾎﾞﾀﾝがｸﾘｯｸされているとき,False:ｸﾘｯｸされていないとき
	Public ReadOnly Property pBtnClickFlg() As Boolean
		Get
            'ﾎﾞﾀﾝの状態
			pBtnClickFlg = m_BtnClickFlg
        End Get
    End Property

	Public ReadOnly Property pSelNowKinmuCD() As String
		Get
            '勤務記号
			pSelNowKinmuCD = m_SelNowKinmuCD
        End Get
    End Property

	'選択されたセットの勤務数
	Public ReadOnly Property pSetCnt() As Integer
		Get
            pSetCnt = m_SetKinmuMark(m_SelSetIdx).KinmuCnt
        End Get
	End Property
	
	'セットの勤務CD
	Public ReadOnly Property pGetSetCD() As String
		Get
            pGetSetCD = m_SetKinmuMark(m_SelSetIdx).CD(m_SetCDIdx)
        End Get
	End Property
	
	Public ReadOnly Property pSelNowRiyuKbn() As String
		Get
            '理由区分
			pSelNowRiyuKbn = m_SelNowRiyuKbn
        End Get
    End Property

    Private Sub CScmdClose_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CScmdClose.Click
        On Error GoTo CScmdClose_Click
        Const W_SUBNAME As String = "NSK0000HB CScmdClose_Click"

        RaiseEvent KensakuEnabled()

        'ﾊﾟﾈﾙｳｨﾝﾄﾞｳ非表示
        Me.Hide()

        Exit Sub
CScmdClose_Click:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub
	
    Private Sub CScmdErase_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CScmdErase.Click
        On Error GoTo CScmdErase_Click
        Const W_SUBNAME As String = "NSK0000HB CScmdErase_Click"

        Static w_SelKinmuCD As String '選択されている勤務CD
        Dim w_ForeColor As Integer
        Dim w_BackColor As Integer
        Dim w_Index As Short
        Dim w_Int As Short
        Dim w_LoopFlg As Boolean
        Dim w_RegKey As String
        Dim w_RegStr As String
        Dim w_SetKinmuFlg As Boolean
        Dim w_Font As Font

        'ﾚｼﾞｽﾄﾘ格納先
        w_RegStr = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

        '勤務ﾎﾞﾀﾝの状態を上に戻っている状態に
        '勤務
        For w_Int = 0 To 14
            w_Font = m_lstCmdKinmu(w_Int).Font
            m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdKinmu(w_Int).Checked = False
        Next w_Int

        '休み
        For w_Int = 0 To 9
            w_Font = m_lstCmdYasumi(w_Int).Font
            m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdYasumi(w_Int).Checked = False
        Next w_Int

        '特殊
        For w_Int = 0 To 4
            w_Font = m_lstCmdTokushu(w_Int).Font
            m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdTokushu(w_Int).Checked = False
        Next w_Int

        'セット勤務
        For w_Int = 0 To 4
            w_Font = m_lstCmdSet(w_Int).Font
            m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdSet(w_Int).Checked = False
        Next w_Int

        'ﾎﾞﾀﾝの状態
        If CScmdErase.Checked = False Then
            'ﾎﾞﾀﾝが押されていない場合

            '選択されていた勤務ﾎﾞﾀﾝを押された状態に
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
            '勤務記号ﾗﾍﾞﾙに設定
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
                    Case "1" '通常
                        '文字/背景色
                        w_ForeColor = ColorTranslator.ToOle(Color.Black)
                        w_BackColor = ColorTranslator.ToOle(Color.White)
                    Case "2" '要請
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", CStr(General.G_PALEGREEN)))
                    Case "3" '希望
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", CStr(General.G_PLUM)))
                    Case "4" '再掲
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", CStr(General.G_LIGHTCYAN)))
                    Case "5"
                        '文字/背景色
                        m_SelNowRiyuKbn = "6"
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", CStr(General.G_ORANGE)))
                    Case Else
                End Select

                '勤務記号ﾗﾍﾞﾙの色設定
                LblSelected.ForeColor = ColorTranslator.FromOle(w_ForeColor)
                LblSelected.BackColor = ColorTranslator.FromOle(w_BackColor)
            End If
        Else
            'ﾎﾞﾀﾝが押されている場合
            '現在選択されている勤務･理由を退避
            w_SelKinmuCD = m_SelNowKinmuCD '勤務ｺｰﾄﾞ
            '消去ﾓｰﾄﾞに設定
            m_SelNowKinmuCD = ""
            m_SelNowRiyuKbn = ""
            LblSelected.Text = ""
            lblSetKinmuNm.Text = ""
            LblSelected.ForeColor = Color.Black
            LblSelected.BackColor = Color.White
        End If

        If g_LimitedFlg = False Then
            If g_SaikeiFlg = False Then
                '理由区分を全部使用可にする
                If w_SetKinmuFlg = True Then
                    m_lstOptRiyu(0).Checked = True
                    m_lstOptRiyu(1).Enabled = False
                    m_lstOptRiyu(2).Enabled = False
                    m_lstOptRiyu(3).Enabled = False
                    m_lstOptRiyu(4).Enabled = False
                Else
                    m_lstOptRiyu(1).Enabled = True

                    '希望回数制限あり　かつ　希望回数0回　の場合

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

        'ﾚｼﾞｽﾄﾘ格納先
        w_RegStr = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

        w_Font = m_lstCmdKinmu(Index).Font
        If m_lstCmdKinmu(Index).Checked Then
            m_lstCmdKinmu(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
        Else
            m_lstCmdKinmu(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
        End If

        '選択されたﾎﾞﾀﾝ以外の状態を上に戻っている状態（押されていない状態）に
        '勤務
        For w_Int = 0 To 14
            If w_Int <> Index Then
                w_Font = m_lstCmdKinmu(w_Int).Font
                m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
                m_lstCmdKinmu(w_Int).Checked = False
            End If
        Next w_Int

        '休み
        For w_Int = 0 To 9
            w_Font = m_lstCmdYasumi(w_Int).Font
            m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdYasumi(w_Int).Checked = False
        Next w_Int

        '特殊
        For w_Int = 0 To 4
            w_Font = m_lstCmdTokushu(w_Int).Font
            m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdTokushu(w_Int).Checked = False
        Next w_Int

        'セット勤務
        For w_Int = 0 To 4
            w_Font = m_lstCmdSet(w_Int).Font
            m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdSet(w_Int).Checked = False
        Next w_Int

        '勤務記号 取得
        w_str = m_KinmuMark(Index + 1 + HscKinmu.Value * 3).Mark

        '2016/2/22 okamura add st --------------
        '理由区分をセットする(勤務ﾎﾞﾀﾝが上に戻っているときも実行)
        For w_Int = 0 To m_lstOptRiyu.Count - 1
            If m_lstOptRiyu(w_Int).Checked Then
                m_SelNowRiyuKbn = w_Int + 1
                Exit For
            End If
        Next
        If m_SelNowRiyuKbn <> "" Then
            If m_SelNowRiyuKbn = "5" Then
                '勤務の選択の場合は理由区分 "6"
                m_SelNowRiyuKbn = "6"
            End If
        End If
        '---------------------------------------

        '勤務記号ﾗﾍﾞﾙに設定
        LblSelected.Text = w_str

        '共通変数 退避
        'KinmuCD
        m_SelNowKinmuCD = m_KinmuMark(Index + 1 + HscKinmu.Value * 3).CD

        'すべての勤務記号配列のClickFlgをFalseに
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

        '選択された勤務記号をTureに
        m_KinmuMark(Index + 1 + HscKinmu.Value * 3).ClickFlg = True

        '消去ﾎﾞﾀﾝが押されているとき
        If CScmdErase.Checked Then
            For w_Int = 0 To m_lstOptRiyu.Count - 1
                If m_lstOptRiyu(w_Int).Checked Then
                    m_SelNowRiyuKbn = w_Int + 1
                    Exit For
                End If
            Next
            If m_SelNowRiyuKbn <> "" Then
                Select Case m_SelNowRiyuKbn
                    Case "1" '通常
                        '文字/背景色
                        w_ForeColor = ColorTranslator.ToOle(Color.Black)
                        w_BackColor = ColorTranslator.ToOle(Color.White)
                    Case "2" '要請
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", CStr(General.G_PALEGREEN)))
                    Case "3" '希望
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", CStr(General.G_PLUM)))
                    Case "4" '再掲
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", CStr(General.G_LIGHTCYAN)))
                    Case "5"
                        '文字/背景色
                        m_SelNowRiyuKbn = "6"
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", CStr(General.G_ORANGE)))
                    Case Else
                End Select

                '勤務記号ﾗﾍﾞﾙの色設定
                LblSelected.ForeColor = ColorTranslator.FromOle(w_ForeColor)
                LblSelected.BackColor = ColorTranslator.FromOle(w_BackColor)
            End If

            '消去ﾎﾞﾀﾝを押されていない状態に
            CScmdErase.Checked = False
        End If

        m_BtnClickFlg = True
        'すべての勤務ﾎﾞﾀﾝが上に戻っているとき()
        If m_lstCmdKinmu(Index).Checked = False Then
            If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_PgmFlg = General.G_PGMSTARTFLG_CHANGEPLAN Then
                '勤務変更のとき貼り付けも削除も行わない
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
                    '理由区分を全部使用可にする
                    m_lstOptRiyu(1).Enabled = True

                    '希望回数制限あり　かつ　希望回数0回　の場合
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

        'ﾚｼﾞｽﾄﾘ格納先
        w_RegStr = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

        w_Font = m_lstCmdSet(Index).Font
        If m_lstCmdSet(Index).Checked Then
            m_lstCmdSet(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
        Else
            m_lstCmdSet(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
        End If

        '選択されたﾎﾞﾀﾝ以外の状態を上に戻っている状態に
        '勤務
        For w_Int = 0 To 14
            w_Font = m_lstCmdKinmu(w_Int).Font
            m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdKinmu(w_Int).Checked = False
        Next w_Int

        '休み
        For w_Int = 0 To 9
            w_Font = m_lstCmdYasumi(w_Int).Font
            m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdYasumi(w_Int).Checked = False
        Next w_Int

        '特殊
        For w_Int = 0 To 4
            w_Font = m_lstCmdTokushu(w_Int).Font
            m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdTokushu(w_Int).Checked = False
        Next w_Int

        'セット勤務
        For w_Int = 0 To 4
            If w_Int <> Index Then
                w_Font = m_lstCmdSet(w_Int).Font
                m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
                m_lstCmdSet(w_Int).Checked = False
            End If
        Next w_Int

        'すべての勤務記号配列のClickFlgをFalseに
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

        '選択された勤務記号をTureに
        m_SetKinmuMark(Index + 1 + HscSet.Value).ClickFlg = True

        '勤務記号 取得
        w_str = m_SetKinmuMark(Index + 1 + HscSet.Value).Mark

        '2016/2/22 okamura add st --------------
        '理由区分をセットする(勤務ﾎﾞﾀﾝが上に戻っているときも実行)
        For w_Int = 0 To m_lstOptRiyu.Count - 1
            If m_lstOptRiyu(w_Int).Checked Then
                m_SelNowRiyuKbn = w_Int + 1
                Exit For
            End If
        Next
        If m_SelNowRiyuKbn <> "" Then
            If m_SelNowRiyuKbn = "5" Then
                '勤務の選択の場合は理由区分 "6"
                m_SelNowRiyuKbn = "6"
            End If
        End If
        '---------------------------------------

        '勤務記号ﾗﾍﾞﾙに設定
        LblSelected.Text = w_str
        '名称文字列を設定
        lblSetKinmuNm.Text = m_SetKinmuMark(Index + 1 + HscSet.Value).StrText

        '共通変数 退避
        'KinmuCD(セット勤務なので勤務CDを1000としておく)
        m_SelNowKinmuCD = CStr(1000 * (Index + 1 + HscSet.Value))

        For w_Int = 0 To m_lstOptRiyu.Count - 1
            If m_lstOptRiyu(w_Int).Checked Then
                m_SelNowRiyuKbn = w_Int + 1
                Exit For
            End If
        Next
        If m_SelNowRiyuKbn <> "" Then
            If m_SelNowRiyuKbn = "5" Then
                '勤務の選択の場合は理由区分 "6"
                m_SelNowRiyuKbn = "6"
            End If
        End If

        '消去ﾎﾞﾀﾝが押されているとき
        If CScmdErase.Checked Then
            If m_SelNowRiyuKbn <> "" Then
                Select Case m_SelNowRiyuKbn
                    Case "1"
                        '文字/背景色
                        w_ForeColor = ColorTranslator.ToOle(Color.Black)
                        w_BackColor = ColorTranslator.ToOle(Color.White)
                    Case "3"
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", CStr(General.G_PLUM)))
                    Case "4" '再掲
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", CStr(General.G_LIGHTCYAN)))
                End Select

                '勤務記号ﾗﾍﾞﾙの色設定
                LblSelected.ForeColor = ColorTranslator.FromOle(w_ForeColor)
                LblSelected.BackColor = ColorTranslator.FromOle(w_BackColor)
            End If

            '消去ﾎﾞﾀﾝを押されていない状態に
            CScmdErase.Checked = False
        End If

        m_BtnClickFlg = True
        'すべての勤務ﾎﾞﾀﾝが上に戻っているとき
        w_Font = m_lstCmdSet(Index).Font
        If m_lstCmdSet(Index).Checked = False Then
            If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_PgmFlg = General.G_PGMSTARTFLG_CHANGEPLAN Then
                '勤務変更のとき貼り付けも削除も行わない
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
                    '希望回数制限あり　かつ　希望回数0回　の場合
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
                    '理由区分は通常のみ
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

        'ﾚｼﾞｽﾄﾘ格納先
        w_RegStr = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

        w_Font = m_lstCmdTokushu(Index).Font
        If m_lstCmdTokushu(Index).Checked Then
            m_lstCmdTokushu(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
        Else
            m_lstCmdTokushu(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
        End If

        '選択されたﾎﾞﾀﾝ以外の状態を上に戻っている状態に
        '勤務
        For w_Int = 0 To 14
            w_Font = m_lstCmdKinmu(w_Int).Font
            m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdKinmu(w_Int).Checked = False
        Next w_Int

        '休み
        For w_Int = 0 To 9
            w_Font = m_lstCmdYasumi(w_Int).Font
            m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdYasumi(w_Int).Checked = False
        Next w_Int

        '特殊
        For w_Int = 0 To 4
            If w_Int <> Index Then
                w_Font = m_lstCmdTokushu(w_Int).Font
                m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
                m_lstCmdTokushu(w_Int).Checked = False
            End If
        Next w_Int

        'セット勤務
        For w_Int = 0 To 4
            w_Font = m_lstCmdSet(w_Int).Font
            m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdSet(w_Int).Checked = False
        Next w_Int

        'すべての勤務記号配列のClickFlgをFalseに
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

        '選択された勤務記号をTureに
        m_TokushuMark(Index + 1 + HscTokushu.Value).ClickFlg = True

        '勤務記号 取得
        w_str = m_TokushuMark(Index + 1 + HscTokushu.Value).Mark

        '2016/2/22 okamura add st --------------
        '理由区分をセットする(勤務ﾎﾞﾀﾝが上に戻っているときも実行)
        For w_Int = 0 To m_lstOptRiyu.Count - 1
            If m_lstOptRiyu(w_Int).Checked Then
                m_SelNowRiyuKbn = w_Int + 1
                Exit For
            End If
        Next
        If m_SelNowRiyuKbn <> "" Then
            If m_SelNowRiyuKbn = "5" Then
                '勤務の選択の場合は理由区分 "6"
                m_SelNowRiyuKbn = "6"
            End If
        End If
        '---------------------------------------

        '勤務記号ﾗﾍﾞﾙに設定
        LblSelected.Text = w_str

        '共通変数 退避
        'KinmuCD
        m_SelNowKinmuCD = m_TokushuMark(Index + 1 + HscTokushu.Value).CD

        '消去ﾎﾞﾀﾝが押されているとき
        If CScmdErase.Checked Then
            For w_Int = 0 To m_lstOptRiyu.Count - 1
                If m_lstOptRiyu(w_Int).Checked Then
                    m_SelNowRiyuKbn = w_Int + 1
                    Exit For
                End If
            Next
            If m_SelNowRiyuKbn <> "" Then
                Select Case m_SelNowRiyuKbn
                    Case "1" '通常
                        '文字/背景色
                        w_ForeColor = ColorTranslator.ToOle(Color.Black)
                        w_BackColor = ColorTranslator.ToOle(Color.White)
                    Case "2" '要請
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", CStr(General.G_PALEGREEN)))
                    Case "3" '希望
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", CStr(General.G_PLUM)))
                    Case "4" '再掲
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", CStr(General.G_LIGHTCYAN)))
                    Case "5"
                        '文字/背景色
                        m_SelNowRiyuKbn = "6"
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", CStr(General.G_ORANGE)))
                    Case Else
                End Select

                '勤務記号ﾗﾍﾞﾙの色設定
                LblSelected.ForeColor = ColorTranslator.FromOle(w_ForeColor)
                LblSelected.BackColor = ColorTranslator.FromOle(w_BackColor)
            End If

            '消去ﾎﾞﾀﾝを押されていない状態に
            CScmdErase.Checked = False
        End If

        m_BtnClickFlg = True
        'すべての勤務ﾎﾞﾀﾝが上に戻っているとき
        If m_lstCmdTokushu(Index).Checked = False Then
            If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_PgmFlg = General.G_PGMSTARTFLG_CHANGEPLAN Then
                '勤務変更のとき貼り付けも削除も行わない
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
                    '理由区分を全部使用可にする
                    m_lstOptRiyu(1).Enabled = True

                    '希望回数制限あり　かつ　希望回数0回　の場合
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

        'ﾚｼﾞｽﾄﾘ格納先
        w_RegStr = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

        w_Font = m_lstCmdYasumi(Index).Font
        If m_lstCmdYasumi(Index).Checked Then
            m_lstCmdYasumi(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
        Else
            m_lstCmdYasumi(Index).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
        End If

        '選択されたﾎﾞﾀﾝ以外の状態を上に戻っている状態に
        '勤務
        For w_Int = 0 To 14
            w_Font = m_lstCmdKinmu(w_Int).Font
            m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdKinmu(w_Int).Checked = False
        Next w_Int

        '休み
        For w_Int = 0 To 9
            If w_Int <> Index Then
                w_Font = m_lstCmdYasumi(w_Int).Font
                m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
                m_lstCmdYasumi(w_Int).Checked = False
            End If
        Next w_Int

        '特殊
        For w_Int = 0 To 4
            w_Font = m_lstCmdTokushu(w_Int).Font
            m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdTokushu(w_Int).Checked = False
        Next w_Int

        'セット勤務
        For w_Int = 0 To 4
            w_Font = m_lstCmdSet(w_Int).Font
            m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdSet(w_Int).Checked = False
        Next w_Int

        'すべての勤務記号配列のClickFlgをFalseに
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

        '選択された勤務記号をTureに
        m_YasumiMark(Index + 1 + HscYasumi.Value * 2).ClickFlg = True

        '勤務記号 取得
        w_str = m_YasumiMark(Index + 1 + HscYasumi.Value * 2).Mark

        '2016/2/22 okamura add st --------------
        '理由区分をセットする(勤務ﾎﾞﾀﾝが上に戻っているときも実行)
        For w_Int = 0 To m_lstOptRiyu.Count - 1
            If m_lstOptRiyu(w_Int).Checked Then
                m_SelNowRiyuKbn = w_Int + 1
                Exit For
            End If
        Next
        If m_SelNowRiyuKbn <> "" Then
            If m_SelNowRiyuKbn = "5" Then
                '勤務の選択の場合は理由区分 "6"
                m_SelNowRiyuKbn = "6"
            End If
        End If
        '---------------------------------------

        '勤務記号ﾗﾍﾞﾙに設定
        LblSelected.Text = w_str

        '共通変数 退避
        'KinmuCD
        m_SelNowKinmuCD = m_YasumiMark(Index + 1 + HscYasumi.Value).CD
        m_SelNowKinmuCD = m_YasumiMark(Index + 1 + HscYasumi.Value * 2).CD

        '消去ﾎﾞﾀﾝが押されているとき
        If CScmdErase.Checked Then
            For w_Int = 0 To m_lstOptRiyu.Count - 1
                If m_lstOptRiyu(w_Int).Checked Then
                    m_SelNowRiyuKbn = w_Int + 1
                    Exit For
                End If
            Next
            If m_SelNowRiyuKbn <> "" Then
                Select Case m_SelNowRiyuKbn
                    Case "1" '通常
                        '文字/背景色
                        w_ForeColor = ColorTranslator.ToOle(Color.Black)
                        w_BackColor = ColorTranslator.ToOle(Color.White)
                    Case "2" '要請
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", CStr(General.G_PALEGREEN)))
                    Case "3" '希望
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", CStr(General.G_PLUM)))
                    Case "4" '再掲
                        '文字/背景色
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", CStr(General.G_LIGHTCYAN)))
                    Case "5"
                        '文字/背景色
                        m_SelNowRiyuKbn = "6"
                        w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", CStr(General.G_BLACK)))
                        w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", CStr(General.G_ORANGE)))
                    Case Else
                End Select

                '勤務記号ﾗﾍﾞﾙの色設定
                LblSelected.ForeColor = ColorTranslator.FromOle(w_ForeColor)
                LblSelected.BackColor = ColorTranslator.FromOle(w_BackColor)
            End If
            '消去ﾎﾞﾀﾝを押されていない状態に
            CScmdErase.Checked = False
        End If

        m_BtnClickFlg = True
        'すべての勤務ﾎﾞﾀﾝが上に戻っているとき()
        If m_lstCmdYasumi(Index).Checked = False Then
            If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_PgmFlg = General.G_PGMSTARTFLG_CHANGEPLAN Then
                '勤務変更のとき貼り付けも削除も行わない
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
                    '理由区分を全部使用可にする
                    m_lstOptRiyu(1).Enabled = True

                    '希望回数制限あり　かつ　希望回数0回　の場合
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
            '最上位に設定
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

        '応援勤務区分の表示FLG
        m_OuenDispFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY7, "OUENDISPFLG", "1", General.g_strHospitalCD))

        '2015/04/14 Bando Add Start ========================================
        '希望モード時の表示対象勤務CD
        m_DispKinmuCd = General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY15, "DISPKINMUCD", "", General.g_strHospitalCD)
        '2015/04/14 Bando Add End   ========================================

        '最上位に設定
        Call General.paSetDialogPos(Me)

        'ｳｨﾝﾄﾞｳの表示ﾎﾟｼﾞｼｮﾝを設定する
        '2018/09/21 K.I Upd Start------------------------------------------------------------------------------------------
        'レジストリ取得を削除
        'Call General.paGetWindowPositon(Me, General.G_STRMAINKEY2 & "\" & g_AppName)
        '画面中央
        w_Left = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - CSng(Me.Width)) / 2)
        w_Top = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - CSng(Me.Height)) / 2)
        Me.SetBounds((CSng(w_Left)), (CSng(w_Top)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
        '2018/09/21 K.I Upd End--------------------------------------------------------------------------------------------

        '{消去}ﾎﾞﾀﾝ
        CScmdErase.Image = Image.FromFile(g_ImagePath & G_ERASER_ICO)
        '{閉じる}ﾎﾞﾀﾝ
        CScmdClose.Image = Image.FromFile(g_ImagePath & G_CLOSE_ICO)

        'ﾊﾟﾈﾙｳｨﾝﾄﾞｳに勤務記号ｾｯﾄ
        Call Set_KinmuData(False)

        '勤務記号ﾗﾍﾞﾙの色設定
        LblSelected.ForeColor = Color.Black
        LblSelected.BackColor = Color.White

        '実績で使用する場合は、理由区分の入力は行わない。1:計画、2:実績
        If m_PgmFlg = General.G_PGMSTARTFLG_NEWPLAN Then
            '計画 の場合
            '再掲部署の場合は、再掲のみを使用可にする
            If g_SaikeiFlg = True Then
                m_lstOptRiyu(0).Enabled = False '通常
                m_lstOptRiyu(0).Visible = False
                m_lstOptRiyu(1).Enabled = False '要請
                m_lstOptRiyu(2).Enabled = False '希望
                m_lstOptRiyu(3).Enabled = True '再掲
                m_lstOptRiyu(3).Visible = True
                m_lstOptRiyu(3).Checked = True
                If m_OuenDispFlg = 0 Then
                    m_lstOptRiyu(4).Enabled = False '応援
                Else
                    m_lstOptRiyu(4).Visible = False
                End If
            Else
                '計画 の場合
                m_lstOptRiyu(0).Enabled = True '通常
                m_lstOptRiyu(1).Enabled = True '要請

                '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                'If g_HopeNumFlg = "1" And g_HopeNum = 0 Then
                If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                    '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                    '希望回数制限あり　かつ　希望回数0回　の場合
                    m_lstOptRiyu(2).Enabled = False '希望
                Else
                    '以外
                    m_lstOptRiyu(2).Enabled = True '希望
                End If

                m_lstOptRiyu(3).Enabled = True '再掲
                m_lstOptRiyu(3).Visible = False
                m_lstOptRiyu(0).Checked = True
                If m_OuenDispFlg = 0 Then
                    m_lstOptRiyu(4).Enabled = True '応援
                    m_lstOptRiyu(4).Visible = True
                Else
                    m_lstOptRiyu(4).Enabled = False '応援
                    m_lstOptRiyu(4).Visible = False
                End If
            End If
        Else
            '実績の場合
            m_lstOptRiyu(0).Enabled = True '通常
            m_lstOptRiyu(0).Checked = True
            m_lstOptRiyu(1).Enabled = False '要請
            m_lstOptRiyu(2).Enabled = False '希望
            m_lstOptRiyu(3).Enabled = False '再掲
            m_lstOptRiyu(3).Visible = False
            '実績の場合はセット使用不可
            _fra_4.Enabled = False
            If m_OuenDispFlg = 1 Then
                m_lstOptRiyu(4).Visible = False
            End If

        End If

        '計画変更の場合は、消しゴムとセット勤務を非表示にする。
        If m_PgmFlg = General.G_PGMSTARTFLG_CHANGEPLAN Then
            CScmdErase.Visible = False
            _fra_4.Visible = False
        End If

        '希望入力の場合は理由区分希望のみ使用可
        If g_LimitedFlg = True Then
            If g_SaikeiFlg = False Then
                m_lstOptRiyu(0).Enabled = False '通常
                m_lstOptRiyu(1).Enabled = False '要請

                '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                'If g_HopeNumFlg = "1" And g_HopeNum = 0 Then
                If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                    '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                    '希望回数制限あり　かつ　希望回数0回　の場合
                    m_lstOptRiyu(2).Enabled = False '希望
                    m_lstOptRiyu(2).Checked = False
                Else
                    '以外
                    m_lstOptRiyu(2).Enabled = True '希望
                    m_lstOptRiyu(2).Checked = True
                End If

                m_lstOptRiyu(3).Enabled = False '再掲
                m_lstOptRiyu(3).Visible = False
                m_lstOptRiyu(4).Enabled = False
                If m_OuenDispFlg = 1 Then
                    m_lstOptRiyu(4).Visible = False
                End If
                m_lstOptRiyu(2).Checked = True
            End If
        End If

        'ﾃﾞﾌｫﾙﾄ設定
        If UBound(m_KinmuMark) > 0 Then
            '(勤務の一番最初のﾎﾞﾀﾝを押された状態に)
            w_Font = m_lstCmdKinmu(0).Font
            m_lstCmdKinmu(0).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
            m_lstCmdKinmu(0).Checked = True
            m_KinmuMark(1).ClickFlg = True
            LblSelected.Text = m_KinmuMark(1).Mark
            m_SelNowKinmuCD = m_KinmuMark(1).CD
            m_SelNowRiyuKbn = "1"
            m_BtnClickFlg = True
        ElseIf UBound(m_YasumiMark) > 0 Then
            '(休みの一番最初のﾎﾞﾀﾝを押された状態に)
            w_Font = m_lstCmdYasumi(0).Font
            m_lstCmdYasumi(0).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Bold)
            m_lstCmdYasumi(0).Checked = True
            m_YasumiMark(1).ClickFlg = True
            LblSelected.Text = m_YasumiMark(1).Mark
            m_SelNowKinmuCD = m_YasumiMark(1).CD
            m_SelNowRiyuKbn = "1"
            m_BtnClickFlg = True
        ElseIf UBound(m_TokushuMark) > 0 Then
            '(特殊の一番最初のﾎﾞﾀﾝを押された状態に)
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
            '勤務変更のとき消しｺﾞﾑﾎﾞﾀﾝ使用不可
            CScmdErase.Enabled = False
            CScmdErase.Visible = False
        End If

        If g_LimitedFlg = True Then
            m_SelNowRiyuKbn = "3" '希望の理由区分
        End If

        If g_SaikeiFlg = True Then
            '再掲部署の場合、理由区分を"再掲"に
            m_SelNowRiyuKbn = "4"
        End If

        '2014/04/23 Shimizu add start P-06979-----------------------------------
        '項目設定の取得
        m_strKinmuEmSecondFlg = Get_ItemValue(General.g_strHospitalCD)
        '勤務記号全角２文字対応のレイアウト変更
        Call SetKinmuSecondView()
        '2014/04/23 Shimizu add end P-06979-------------------------------------

        Exit Sub
Form_Load:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    'ﾊﾟﾈﾙｳｨﾝﾄﾞｳに勤務記号をｾｯﾄ
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
        Dim w_記号_F As ADODB.Field
        Dim w_勤務CD1_F As ADODB.Field
        Dim w_勤務CD2_F As ADODB.Field
        Dim w_勤務CD3_F As ADODB.Field
        Dim w_勤務CD4_F As ADODB.Field
        Dim w_勤務CD5_F As ADODB.Field
        Dim w_勤務CD6_F As ADODB.Field
        Dim w_勤務CD7_F As ADODB.Field
        Dim w_勤務CD8_F As ADODB.Field
        Dim w_勤務CD9_F As ADODB.Field
        Dim w_勤務CD10_F As ADODB.Field
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

        '勤務ﾃﾞｰﾀ格納配列初期化
        ReDim m_KinmuMark(0)
        ReDim m_YasumiMark(0)
        ReDim m_TokushuMark(0)
        ReDim w_strTaihiCD(0)
        w_Int2 = 1

        '勤務名称マスタ取得(勤務部署)
        With General.g_objGetMaster
            .pHospitalCD = General.g_strHospitalCD '施設コード
            .pKN_GetKbn = 1 '0:全件 1:指定勤務部署
            .pKN_KinmuDeptCD = General.g_strSelKinmuDeptCD '選択勤務部署

            If .mGetKinmuNameM = False Then
            Else
                'マスタ件数
                w_RecCnt = .fKN_KinmuCount

                For w_Int = 1 To w_RecCnt

                    '索引
                    .mKN_KinmuIdx = w_Int

                    .pKN_GetKbn = 0
                    w_lngEffEndDate = .fKN_EffToDate
                    .pKN_GetKbn = 1
                    If w_lngEffEndDate >= m_StartDate Or w_lngEffEndDate = 0 Or w_lngEffEndDate = 99999999 Then

                        '●勤務分類コード
                        w_strKinmuBunruiCD = .fKN_KinmuBunruiCD

                        If w_strKinmuBunruiCD = "1" Then
                            '-- 勤務 --
                            '2015/04/14 Bando Upd Start ============================
                            '希望モードの場合、表示対象勤務のみパレットに表示
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
                            '-- 休み --
                            '2015/04/14 Bando Upd Start ============================
                            '希望モードの場合、表示対象勤務のみパレットに表示
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
                            '-- 特殊 --
                            '2015/04/14 Bando Upd Start ============================
                            '希望モードの場合、表示対象勤務のみパレットに表示
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
                            '希望モードの場合、表示対象勤務のみパレットに表示
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

        '勤務（パレットに勤務マークを設定する）
        For w_Int = 0 To 14
            If w_Int <= w_KinmuCnt - 1 Then
                m_lstCmdKinmu(w_Int).Text = m_KinmuMark(w_Int + 1).Mark
                If m_KinmuMark(w_Int + 1).Setumei = "" Then
                    ToolTip1.SetToolTip(m_lstCmdKinmu(w_Int), Get_KinmuTipText(m_KinmuMark(w_Int + 1).CD))
                Else
                    ToolTip1.SetToolTip(m_lstCmdKinmu(w_Int), Get_KinmuTipText(m_KinmuMark(w_Int + 1).CD) & "：" & m_KinmuMark(w_Int + 1).Setumei)
                End If
            Else
                Exit For
            End If
        Next w_Int

        '休み
        For w_Int = 0 To 9
            If w_Int <= w_YasumiCnt - 1 Then
                m_lstCmdYasumi(w_Int).Text = m_YasumiMark(w_Int + 1).Mark
                If m_YasumiMark(w_Int + 1).Setumei = "" Then
                    ToolTip1.SetToolTip(m_lstCmdYasumi(w_Int), Get_KinmuTipText(m_YasumiMark(w_Int + 1).CD))
                Else
                    ToolTip1.SetToolTip(m_lstCmdYasumi(w_Int), Get_KinmuTipText(m_YasumiMark(w_Int + 1).CD) & "：" & m_YasumiMark(w_Int + 1).Setumei)
                End If
            Else
                Exit For
            End If
        Next w_Int

        '特殊勤務
        For w_Int = 0 To 4
            If w_Int <= w_TokushuCnt - 1 Then
                m_lstCmdTokushu(w_Int).Text = m_TokushuMark(w_Int + 1).Mark
                If m_TokushuMark(w_Int + 1).Setumei = "" Then
                    ToolTip1.SetToolTip(m_lstCmdTokushu(w_Int), Get_KinmuTipText(m_TokushuMark(w_Int + 1).CD))
                Else
                    ToolTip1.SetToolTip(m_lstCmdTokushu(w_Int), Get_KinmuTipText(m_TokushuMark(w_Int + 1).CD) & "：" & m_TokushuMark(w_Int + 1).Setumei)
                End If
            Else
                Exit For
            End If
        Next w_Int

        'ｽｸﾛｰﾙﾊﾞｰ、ｵﾌﾟｼｮﾝﾎﾞﾀﾝの設定
        '勤務
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

        '休み
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

        '特殊勤務
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
        'セット勤務配列初期化
        ReDim m_SetKinmuMark(0)
        '----------------------------

        '2015/06/02 Bando Upd Start ==========================
        If g_HopeMode <> 1 Then
            'セット勤務
            '2017/05/22 Richard Upd Start
            ''SQL文編集
            'w_Sql = "SELECT * FROM NS_SETKINMU_M "
            'w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
            'w_Sql = w_Sql & "AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
            'w_Sql = w_Sql & "ORDER BY DISPNO "

            'w_Rs = General.paDBRecordSetOpen(w_Sql)
            '<1>
            Call NSK0000H_sql.select_NS_SETKINMU_M_01(w_Rs)
            'Upd End
            '2015/7/6 okamura del st ----
            ''セット勤務配列初期化
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

                    w_記号_F = .Fields("SetMark")
                    w_勤務CD1_F = .Fields("KinmuCD1")
                    w_勤務CD2_F = .Fields("KinmuCD2")
                    w_勤務CD3_F = .Fields("KinmuCD3")
                    w_勤務CD4_F = .Fields("KinmuCD4")
                    w_勤務CD5_F = .Fields("KinmuCD5")
                    w_勤務CD6_F = .Fields("KinmuCD6")
                    w_勤務CD7_F = .Fields("KinmuCD7")
                    w_勤務CD8_F = .Fields("KinmuCD8")
                    w_勤務CD9_F = .Fields("KinmuCD9")
                    w_勤務CD10_F = .Fields("KinmuCD10")

                    For w_Int = 1 To w_RecCnt
                        w_blnEndDate = True
                        w_strKinmuCD1 = w_勤務CD1_F.Value & ""
                        w_strKinmuCD2 = w_勤務CD2_F.Value & ""
                        w_strKinmuCD3 = w_勤務CD3_F.Value & ""
                        w_strKinmuCD4 = w_勤務CD4_F.Value & ""
                        w_strKinmuCD5 = w_勤務CD5_F.Value & ""
                        w_strKinmuCD6 = w_勤務CD6_F.Value & ""
                        w_strKinmuCD7 = w_勤務CD7_F.Value & ""
                        w_strKinmuCD8 = w_勤務CD8_F.Value & ""
                        w_strKinmuCD9 = w_勤務CD9_F.Value & ""
                        w_strKinmuCD10 = w_勤務CD10_F.Value & ""

                        '勤務CD1～10までで退避していた勤務CDと一致しているものは期限切れ
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

                            m_SetKinmuMark(w_Int3).Mark = w_記号_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(1) = w_勤務CD1_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(2) = w_勤務CD2_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(3) = w_勤務CD3_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(4) = w_勤務CD4_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(5) = w_勤務CD5_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(6) = w_勤務CD6_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(7) = w_勤務CD7_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(8) = w_勤務CD8_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(9) = w_勤務CD9_F.Value & ""
                            m_SetKinmuMark(w_Int3).CD(10) = w_勤務CD10_F.Value & ""
                            m_SetKinmuMark(w_Int3).ClickFlg = False
                            m_SetKinmuMark(w_Int3).blnKinmu = True

                            '勤務がいくつあるか(間に空白はないものとする)
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
                '全件Visible=Trueにする
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

        'メイン画面から呼ばれた(セット勤務を更新した)場合は選択ボタンを初期化
        If p_CallMainFlg = True And m_SelNowKinmuCD <> "" Then
            If Integer.Parse(m_SelNowKinmuCD) >= 1000 Then
                '現在選択勤務がセット勤務の場合
                If UBound(m_SetKinmuMark) > 0 Then
                    HscSet.Value = 0
                    m_lstCmdSet(0).Checked = False
                    m_BtnClickFlg = True
                    Call m_lstCmdSet_ClickEvent(m_lstCmdSet.Item(0), New System.EventArgs())
                Else
                    '現在選択勤務がセット勤務以外の場合
                    lblSetKinmuNm.Text = ""
                    'ﾃﾞﾌｫﾙﾄ設定
                    If UBound(m_KinmuMark) > 0 Then
                        '(勤務の一番最初のﾎﾞﾀﾝを押された状態に)
                        m_lstCmdKinmu(0).Checked = False
                        m_SelNowRiyuKbn = "1"
                        m_BtnClickFlg = True
                        Call m_lstCmdKinmu_ClickEvent(m_lstCmdKinmu.Item(0), New System.EventArgs())
                    ElseIf UBound(m_YasumiMark) > 0 Then
                        '(休みの一番最初のﾎﾞﾀﾝを押された状態に)
                        m_lstCmdYasumi(0).Checked = False
                        m_SelNowRiyuKbn = "1"
                        m_BtnClickFlg = True
                        Call m_lstCmdYasumi_ClickEvent(m_lstCmdYasumi.Item(0), New System.EventArgs())
                    ElseIf UBound(m_TokushuMark) > 0 Then
                        '(特殊の一番最初のﾎﾞﾀﾝを押された状態に)
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

    'セット勤務ツールチップ用文字列取得
    Public Function Get_SetKinmuTipText(ByVal p_Int As Integer) As String
        On Error GoTo Get_SetKinmuTipText
        Const W_SUBNAME As String = "NSK0000HB Get_SetKinmuTipText"

        Dim w_str As String
        Dim w_strTEXT As String
        Dim w_Cnt As Integer
        Dim w_CD As String

        For w_Cnt = 1 To 10
            '勤務CDを取得
            w_CD = m_SetKinmuMark(p_Int).CD(w_Cnt)

            '空白でなく数値である場合
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

        'ｳｨﾝﾄﾞｳの表示ﾎﾟｼﾞｼｮﾝを格納する
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

        'ｺﾏﾝﾄﾞﾎﾞﾀﾝのCaption設定
        '勤務
        w_Hsc_Cnt = newScrollValue

        For w_Int = 0 To 14
            'ﾎﾞﾀﾝの状態を上に戻っている状態に
            w_Font = m_lstCmdKinmu(w_Int).Font
            m_lstCmdKinmu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdKinmu(w_Int).Checked = False
            w_Cnt = w_Int + 1 + w_Hsc_Cnt * 3

            If w_Cnt <= UBound(m_KinmuMark) Then

                m_lstCmdKinmu(w_Int).Text = m_KinmuMark(w_Int + 1 + w_Hsc_Cnt * 3).Mark
                If m_KinmuMark(w_Int + 1 + w_Hsc_Cnt * 3).Setumei = "" Then
                    ToolTip1.SetToolTip(m_lstCmdKinmu(w_Int), Get_KinmuTipText(m_KinmuMark(w_Int + 1 + w_Hsc_Cnt * 3).CD))
                Else
                    ToolTip1.SetToolTip(m_lstCmdKinmu(w_Int), Get_KinmuTipText(m_KinmuMark(w_Int + 1 + w_Hsc_Cnt * 3).CD) & "：" & m_KinmuMark(w_Int + 1 + w_Hsc_Cnt * 3).Setumei)
                End If

                If CScmdErase.Checked = False Then
                    If m_KinmuMark(w_Int + 1 + w_Hsc_Cnt * 3).ClickFlg = True Then
                        'ﾎﾞﾀﾝをｸﾘｯｸされた状態に
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

        'ｺﾏﾝﾄﾞﾎﾞﾀﾝのCaption設定
        '休み
        w_Hsc_Cnt = newScrollValue

        For w_Int = 0 To 9
            'ﾎﾞﾀﾝの状態を上に戻っている状態に

            w_Font = m_lstCmdYasumi(w_Int).Font
            m_lstCmdYasumi(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdYasumi(w_Int).Checked = False
            w_Cnt = w_Int + 1 + w_Hsc_Cnt * 2

            If w_Cnt <= UBound(m_YasumiMark) And (w_Int + 1 + w_Hsc_Cnt * 2) <= UBound(m_YasumiMark) Then
                m_lstCmdYasumi(w_Int).Text = m_YasumiMark(w_Int + 1 + w_Hsc_Cnt * 2).Mark
                If m_YasumiMark(w_Int + 1 + w_Hsc_Cnt * 2).Setumei = "" Then
                    ToolTip1.SetToolTip(m_lstCmdYasumi(w_Int), Get_KinmuTipText(m_YasumiMark(w_Int + 1 + w_Hsc_Cnt * 2).CD))
                Else
                    ToolTip1.SetToolTip(m_lstCmdYasumi(w_Int), Get_KinmuTipText(m_YasumiMark(w_Int + 1 + w_Hsc_Cnt * 2).CD) & "：" & m_YasumiMark(w_Int + 1 + w_Hsc_Cnt * 2).Setumei)
                End If

                If CScmdErase.Checked = False Then
                    If m_YasumiMark(w_Int + 1 + w_Hsc_Cnt * 2).ClickFlg = True Then
                        'ﾎﾞﾀﾝをｸﾘｯｸされた状態に
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

        'ｺﾏﾝﾄﾞﾎﾞﾀﾝのCaption設定
        'セット勤務
        w_Hsc_Cnt = newScrollValue
        For w_Int = 0 To 4

            If UBound(m_SetKinmuMark) >= w_Int + 1 Then
                'ﾎﾞﾀﾝの状態を上に戻っている状態に
                w_Font = m_lstCmdSet(w_Int).Font
                m_lstCmdSet(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
                m_lstCmdSet(w_Int).Checked = False

                m_lstCmdSet(w_Int).Text = m_SetKinmuMark(w_Int + 1 + w_Hsc_Cnt).Mark
                ToolTip1.SetToolTip(m_lstCmdSet(w_Int), Get_SetKinmuTipText(w_Int + 1 + w_Hsc_Cnt))

                If CScmdErase.Checked = False Then
                    If m_SetKinmuMark(w_Int + 1 + w_Hsc_Cnt).ClickFlg = True Then
                        'ﾎﾞﾀﾝをｸﾘｯｸされた状態に
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

        'ｺﾏﾝﾄﾞﾎﾞﾀﾝのCaption設定
        '特殊勤務
        w_Hsc_Cnt = newScrollValue
        For w_Int = 0 To 4
            'ﾎﾞﾀﾝの状態を上に戻っている状態に
            w_Font = m_lstCmdTokushu(w_Int).Font
            m_lstCmdTokushu(w_Int).Font = New Font(w_Font.FontFamily, w_Font.Size, FontStyle.Regular)
            m_lstCmdTokushu(w_Int).Checked = False

            m_lstCmdTokushu(w_Int).Text = m_TokushuMark(w_Int + 1 + w_Hsc_Cnt).Mark
            If m_TokushuMark(w_Int + 1 + w_Hsc_Cnt).Setumei = "" Then
                ToolTip1.SetToolTip(m_lstCmdTokushu(w_Int), Get_KinmuTipText(m_TokushuMark(w_Int + 1 + w_Hsc_Cnt).CD))
            Else
                ToolTip1.SetToolTip(m_lstCmdTokushu(w_Int), Get_KinmuTipText(m_TokushuMark(w_Int + 1 + w_Hsc_Cnt).CD) & "：" & m_TokushuMark(w_Int + 1 + w_Hsc_Cnt).Setumei)
            End If

            If CScmdErase.Checked = False Then
                If m_TokushuMark(w_Int + 1 + w_Hsc_Cnt).ClickFlg = True Then
                    'ﾎﾞﾀﾝをｸﾘｯｸされた状態に
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

            'ﾚｼﾞｽﾄﾘ格納先
            w_RegStr = General.G_SYSTEM_Win7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            '消去ﾎﾞﾀﾝが押されているときは､色の変更をしない
            w_Font = CScmdErase.Font
            If w_Font.Bold = False Then

                w_Index = Index
                'オプションボタンのチェック
                If w_Index <> True Then
                    '共通変数 退避
                    '理由区分（通常,要請,希望）
                    m_SelNowRiyuKbn = CStr(w_Index + 1)

                    '勤務記号ﾗﾍﾞﾙの色設定
                    '理由区分 ?
                    Select Case m_SelNowRiyuKbn
                        Case "1" '通常
                            '文字/背景色
                            w_ForeColor = ColorTranslator.ToOle(Color.Black)
                            w_BackColor = ColorTranslator.ToOle(Color.White)
                        Case "2" '要請
                            '文字/背景色
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", CStr(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", CStr(General.G_PALEGREEN)))
                        Case "3" '希望
                            '文字/背景色
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", CStr(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", CStr(General.G_PLUM)))
                        Case "4" '再掲
                            '文字/背景色
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", CStr(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", CStr(General.G_LIGHTCYAN)))
                        Case "5" '応援
                            m_SelNowRiyuKbn = "6"
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", CStr(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", CStr(General.G_ORANGE)))
                        Case Else
                    End Select

                    '勤務記号ﾗﾍﾞﾙの色設定
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

    'コントロール配列の代わりにリストに格納する
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
    '/  概要　　　　  : 勤務記号全角２文字対応のレイアウト変更
    '/  パラメータ    : なし
    '/  戻り値        : なし
    '/----------------------------------------------------------------------/
    Private Sub SetKinmuSecondView()
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc 'ﾌﾟﾛｼｰｼﾞｬ名の一時待避
        General.g_ErrorProc = "NSK0000HB SetKinmuSecondView"

        Const W_FRAME_FIRST_HEIGHT As Integer = 16 '1行目の縦位置
        Const W_FRAME_FIRST_WIDTH As Integer = 8 '1列目の横位置
        Const W_FRAME_ADD_HEIGHT As Integer = 24 '行の縦位置増え幅
        Const W_FRAME_ADD_WIDTH As Integer = 39 '行の横位置増え幅

        Const W_KINMU_HEIGHT As Integer = 25 'フレームの縦幅
        Const W_FRAME_WIDTH As Integer = 213 'フレームの横幅
        Const W_SCL_WIDTH As Integer = 196 'スクロールの横幅
        Const W_SCL_HEIGHT As Integer = 17 'スクロールの縦幅
        Const W_KINMU_WIDTH As Integer = 40 '勤務の横幅

        Const W_SET_HEIGHT_ADJUST As Integer = 30 'セット勤務の高さ調整

        Try
            '勤務記号全角２文字対応フラグ判定
            If m_strKinmuEmSecondFlg = "0" Then
                '0：対応しない(従来の勤務記号入力サイズと最大2バイト)

            Else
                '1：対応する(全角２文字が表示できる勤務記号入力サイズと最大4バイト)
                'フォーム
                Me.Size = New System.Drawing.Size(330, 435)

                '記号
                _fra_3.Location = New Point(240, 2)
                LblSelected.Location = New Point(10, 17)
                _fra_3.Size = New System.Drawing.Size(80, 47)
                LblSelected.Size = New System.Drawing.Size(61, 24)
                '区分
                PnlRiyu.Location = New Point(240, 54)
                CScmdErase.Location = New Point(240, 155)
                CScmdClose.Location = New Point(240, 205)

                'パレットを載せているパネル
                SSPanel2.Size = New System.Drawing.Size(230, 387)

                'フレーム
                _fra_0.Size = New System.Drawing.Size(W_FRAME_WIDTH, 113)
                _fra_1.Size = New System.Drawing.Size(W_FRAME_WIDTH, 89)
                _fra_2.Size = New System.Drawing.Size(W_FRAME_WIDTH, 65)
                _fra_4.Size = New System.Drawing.Size(W_FRAME_WIDTH, 95)

                'スクロール
                HscKinmu.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)
                HscYasumi.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)
                HscTokushu.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)
                HscSet.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)


                '勤務
                General.setSizeAndLocal(m_lstCmdKinmu, 3, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

                '休み
                General.setSizeAndLocal(m_lstCmdYasumi, 2, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

                '特殊勤務
                General.setSizeAndLocal(m_lstCmdTokushu, 1, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

                'セット
                lblSetKinmuNm.Width = 180
                lblSetKinmuNm.Height = 30
                CType(HscSet, System.Windows.Forms.Control).Location = New Point(W_FRAME_FIRST_WIDTH, 70)
                CType(m_lstCmdSet(0), System.Windows.Forms.Control).Location = New Point(W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT + W_SET_HEIGHT_ADJUST)
                General.setSizeAndLocal(m_lstCmdSet, 1, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT + W_SET_HEIGHT_ADJUST, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT + W_SET_HEIGHT_ADJUST)

            End If

            General.g_ErrorProc = w_PreErrorProc '待避ﾌﾟﾛｼｰｼﾞｬ名を元に戻す

        Catch ex As Exception
            Err.Raise(Err.Number)
        End Try
    End Sub
    '2014/04/23 Shimizu add end P-06979----------------------------------------------------------------------------------------------------
End Class