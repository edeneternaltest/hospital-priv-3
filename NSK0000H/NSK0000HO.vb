'****************************************************************************
'* Copyright (C) 2012 Nihon Inter Systems Corporation. All Rights Reserved. *
'****************************************************************************
Option Strict Off
Option Explicit On
Imports System.Collections.Generic

''' <summary>
''' 行事一覧画面(NSK0000HO)
''' </summary>
''' <remarks>
''' 規定値：画面ＩＤ　"NSK0000HO"
'''         画面名称　"行事一覧画面"
''' </remarks>
''' <history>
''' ===================================================================
''' 更新履歴
''' 項番       更新日付        担当者       更新内容
''' 0001       2012/07/**      ******　 　　P-*****（新規作成。パッケージバージョンアップ）
''' ===================================================================
''' </history>
Friend Class frmNSK0000HO
    Inherits Form

#Region "定数"
    ''' <summary>
    ''' 行事一覧のリスト定数クラス
    ''' </summary>
    ''' <remarks></remarks>
    Private Class EventList
        'キー
        Friend Const DATEF As String = "DATEF" '日程
        Friend Const TIME As String = "TIME" '時間
        Friend Const EVENTNAME As String = "EVENTNAME" '行事
        Friend Const ALL As String = "ALL" '全体
        '名称
        Friend DispNm As Dictionary(Of String, String)
        Friend Size As Dictionary(Of String, Integer)

        Sub New()
            DispNm = New Dictionary(Of String, String)
            DispNm.Item(DATEF) = "日程"
            DispNm.Item(TIME) = "時間"
            DispNm.Item(EVENTNAME) = "行事"
            DispNm.Item(ALL) = "全体"

            Size = New Dictionary(Of String, Integer)
            Size.Item(DATEF) = 800
            Size.Item(TIME) = 2000
            Size.Item(EVENTNAME) = 3000
            Size.Item(ALL) = 800
        End Sub
    End Class

    ''' <summary>
    ''' 行事参加者のリスト定数クラス
    ''' </summary>
    ''' <remarks></remarks>
    Private Class EventStaff
        'キー
        Friend Const EVENTNAME As String = "EVENTNAME"
        Friend Const POST As String = "POST"
        '名称
        Friend DispNm As Dictionary(Of String, String)
        Friend Size As Dictionary(Of String, Integer)

        Sub New()
            DispNm = New Dictionary(Of String, String)
            DispNm.Item(EVENTNAME) = "氏名"
            DispNm.Item(POST) = "役職"

            Size = New Dictionary(Of String, Integer)
            Size.Item(EVENTNAME) = 2500
            Size.Item(POST) = 2500
        End Sub
    End Class

    Private EvLst As EventList
    Private EvStf As EventStaff
#End Region

#Region "変数"
    Private m_errMsg As String

    Private m_EventData() As EventList_Type '行事情報
    Private m_selDate As Integer’選択日付
    Private m_FormShowFlg As Boolean '表示フラグ
#End Region

#Region "プロパティ"
    ''' <summary>
    ''' 表示フラグ
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property pShowFlg() As Boolean
        Get
            Return m_FormShowFlg
        End Get
        Set(ByVal Value As Boolean)
            m_FormShowFlg = Value
        End Set
    End Property

    ''' <summary>
    ''' 行事予定情報を設定
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property pEventData() As EventList_Type()
        Set(ByVal value As EventList_Type())
            m_EventData = General.paDeepCopy(value)
        End Set
    End Property

    ''' <summary>
    ''' 選択日付を設定
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property pSelDate() As Integer
        Set(ByVal value As Integer)
            m_selDate = value
        End Set
    End Property
#End Region

#Region "フォーム関連"
    Sub New()
        ' この呼び出しは、Windows フォーム デザイナで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        EvLst = New EventList
        EvStf = New EventStaff
    End Sub

    ''' <summary>
    ''' フォームロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmNSK0000HO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        m_errMsg = "frmNSK0000HO frmNSK0000HO_Load"
        Try
            '2018/09/21 K.I Add Start-------------------------
            Dim w_Left As String
            Dim w_Top As String
            '2018/09/21 K.I Add End---------------------------

            'ウィンドゥを画面の最上位に設定
            Call General.paSetDialogPos(Me)

            'リスト生成
            setEventListData()
            'ｳｨﾝﾄﾞｳの表示ﾎﾟｼﾞｼｮﾝを設定する
            '2018/09/21 K.I Upd Start------------------------------------------------------------------------------------------
            'レジストリ取得を削除
            'Call General.paGetWindowPositon(Me, General.G_STRMAINKEY2 & "\" & g_AppName)
            '画面中央
            w_Left = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - CSng(Me.Width)) / 2)
            w_Top = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - CSng(Me.Height)) / 2)
            Me.SetBounds((CSng(w_Left)), (CSng(w_Top)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
            '2018/09/21 K.I Upd End--------------------------------------------------------------------------------------------

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), m_errMsg)
        End Try
    End Sub

    ''' <summary>
    ''' フォームクロージング
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmNSK0000HO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        m_errMsg = "frmNSK0000HO frmNSK0000HO_FormClosing"

        Dim UnloadMode As CloseReason = e.CloseReason

        Try
            If UnloadMode = CloseReason.UserClosing Then
                'ｺﾝﾄﾛｰﾙﾒﾆｭｰから閉じられた場合はUnloadしない
                e.Cancel = True
                Me.Hide()
                m_FormShowFlg = False
            End If

            'ｳｨﾝﾄﾞｳの表示ﾎﾟｼﾞｼｮﾝを格納する
            Call General.paPutWindowPositon(Me, General.G_STRMAINKEY2 & "\" & g_AppName)
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), m_errMsg)
        End Try
    End Sub

    ''' <summary>
    ''' フォームクローズド
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmNSK0000HO_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs)
        m_errMsg = "frmNSK0000HO frmNSK0000HO_FormClosed"
        Try
            '非表示
            Me.Hide()
            m_FormShowFlg = False
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), m_errMsg)
        End Try
    End Sub

    ''' <summary>
    ''' フォームアクティブ
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmNSK0000HO_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        m_errMsg = "frmNSK0000HO frmNSK0000HO_Activated"
        Try
            '非表示
            If Me.Visible AndAlso Not m_FormShowFlg Then
                setEventListData()
                'フォーカスを当て選択状態にする
                RemoveHandler lstEventList.SelectedIndexChanged, AddressOf Me.lstEventList_SelectedIndexChanged
                lstEventList.Select()
                AddHandler lstEventList.SelectedIndexChanged, AddressOf Me.lstEventList_SelectedIndexChanged
            End If
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), m_errMsg)
        End Try
    End Sub
#End Region

#Region "イベント関連"
    ''' <summary>
    ''' 閉じるボタン押下
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        m_errMsg = "frmNSK0000HO cmdClose_Click"
        Try
            '非表示
            Me.Hide()
            m_FormShowFlg = False
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), m_errMsg)
        End Try
    End Sub

    ''' <summary>
    ''' 行事一覧選択
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub lstEventList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstEventList.SelectedIndexChanged
        m_errMsg = "frmNSK0000HO cmdClose_Click"

        Try
            If lstEventList.SelectedItems.Count <= 0 Then Exit Sub
            '行事参加者生成
            setEventStaffData()

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), m_errMsg)
        End Try
    End Sub
#End Region

#Region "共通処理"
    ''' <summary>
    ''' 行事一覧リストビュー生成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub setEventListData()
        Dim errMsg As String = m_errMsg
        m_errMsg = "frmNSK0000HO setEventListData"

        Dim lstItem() As String
        Dim w_day As Date
        Dim w_date As String
        Dim w_time As String
        Dim w_strAll As String
        Dim w_flg As Boolean
        Dim w_index As Short = 0

        Try
            '初期化
            lstEventList.BeginUpdate()
            lstEventList.Clear()

            '行事一覧の枠作成
            With lstEventList
                'ListViewの表示設定
                .View = View.Details
                .GridLines = True
                .FullRowSelect = True
                .TabStop = False
                .HideSelection = False   'リストビューがフォーカスを失っても、選択状態を保持する
                .MultiSelect = False    '複数選択不可
                .Scrollable = System.Windows.Forms.ScrollBars.Horizontal

                'ヘッダー部の追加
                .Columns.Add(EventList.DATEF, EvLst.DispNm(EventList.DATEF), (General.paTwipsTopixels(EvLst.Size(EventList.DATEF))), System.Windows.Forms.HorizontalAlignment.Left, "")
                .Columns.Add(EventList.TIME, EvLst.DispNm(EventList.TIME), (General.paTwipsTopixels(EvLst.Size(EventList.TIME))), System.Windows.Forms.HorizontalAlignment.Left, "")
                .Columns.Add(EventList.EVENTNAME, EvLst.DispNm(EventList.EVENTNAME), (General.paTwipsTopixels(EvLst.Size(EventList.EVENTNAME))), System.Windows.Forms.HorizontalAlignment.Left, "")
                .Columns.Add(EventList.ALL, EvLst.DispNm(EventList.ALL), (General.paTwipsTopixels(EvLst.Size(EventList.ALL))), System.Windows.Forms.HorizontalAlignment.Center, "")
            End With

            '行事一覧を生成する
            w_flg = False
            For i As Integer = 1 To UBound(m_EventData)
                '日程は日付のみ
                w_day = General.paGetDateFromDateInteger(m_EventData(i).DateF)
                w_date = General.paFormatSpace(General.paGetDateIntegerFromDate(w_day, General.G_DATE_ENUM.MM), 2) _
                         & "/" & General.paFormatSpace(General.paGetDateIntegerFromDate(w_day, General.G_DATE_ENUM.dd), 2)
                '時間
                w_time = getConvTime(m_EventData(i).Time_st) & "～" & getConvTime(m_EventData(i).Time_ed)
                '全体
                w_strAll = ""
                If m_EventData(i).allFlg Then w_strAll = "○"

                lstItem = New String() {w_date, _
                                        w_time, _
                                        m_EventData(i).EventName, _
                                        w_strAll}
                lstEventList.Items.Add(New ListViewItem(lstItem))

                '選択した日付の行事を初期選択
                If m_EventData(i).DateF = m_selDate AndAlso Not w_flg Then
                    w_index = i - 1
                    w_flg = True
                End If
            Next
            lstEventList.EndUpdate()

            lstEventList.Items(w_index).Selected = True '対象の項目を選択済にする
            lstEventList.Items(w_index).EnsureVisible()

            m_errMsg = errMsg
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 行事参加者リストビュー生成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub setEventStaffData()
        Dim errMsg As String = m_errMsg
        m_errMsg = "frmNSK0000HO setEventStaffData"

        Dim w_index As Integer
        Dim lstItem() As String

        Try
            '初期化
            lstEventStaff.BeginUpdate()
            lstEventStaff.Clear()
            w_index = lstEventList.SelectedItems(0).Index() + 1

            '行事参加者の枠作成
            With lstEventStaff
                'ListViewの表示設定
                .View = View.Details
                .GridLines = True
                .FullRowSelect = True
                .TabStop = False
                .HideSelection = False   'リストビューがフォーカスを失っても、選択状態を保持する
                .MultiSelect = False    '複数選択不可
                .Scrollable = System.Windows.Forms.ScrollBars.Vertical

                'ヘッダー部の追加
                .Columns.Add(EventStaff.EVENTNAME, EvStf.DispNm(EventStaff.EVENTNAME), (General.paTwipsTopixels(EvStf.Size(EventStaff.EVENTNAME))), System.Windows.Forms.HorizontalAlignment.Left, "")
                .Columns.Add(EventStaff.POST, EvStf.DispNm(EventStaff.POST), (General.paTwipsTopixels(EvStf.Size(EventStaff.POST))), System.Windows.Forms.HorizontalAlignment.Left, "")
            End With
            lstEventStaff.EndUpdate()

            If UBound(m_EventData(w_index).EventStaff) = 0 Then Exit Sub
            '行事参加者を生成する
            With m_EventData(w_index)
                For i As Short = 1 To UBound(.EventStaff)
                    lstItem = New String() {.EventStaff(i).staffNm, _
                                            .EventStaff(i).postNm}
                    lstEventStaff.Items.Add(New ListViewItem(lstItem))
                Next
            End With

            'ラベル変更
            lblEventName.Text = lstEventList.SelectedItems.Item(0).SubItems(2).Text

            m_errMsg = errMsg
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 時間フォーマット変更
    ''' </summary>
    ''' <param name="p_time"></param>
    ''' <returns>String型</returns>
    ''' <remarks>HHMM→HH:MM</remarks>
    Private Function getConvTime(ByVal p_time As Short) As String
        Dim errMsg As String = m_errMsg
        m_errMsg = "frmNSK0000HO getConvTime"

        Dim w_rtnTime As String

        Try
            '4桁フォーマット
            If General.paLenB(p_time) < 4 Then
                w_rtnTime = General.paFormatZero(p_time, 4)
            Else
                w_rtnTime = Convert.ToString(p_time)
            End If

            'セミコロンで区切る
            If w_rtnTime = "9999" Then
                w_rtnTime = "00:00"
            Else
                'セミコロンで区切る
                w_rtnTime = Strings.Left(w_rtnTime, 2) & ":" & Strings.Right(w_rtnTime, 2)
            End If

            m_errMsg = errMsg
            Return w_rtnTime
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

End Class