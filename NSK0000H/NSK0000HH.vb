Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Friend Class frmNSK0000HH
    Inherits General.FormBase
    Public NSK0000H_sql As New clsNSK0000H_sql
    '割当対象勤務選択ｵﾌﾟｼｮﾝｲﾝﾃﾞｯｸｽ
    Private Const M_OPT_STANDARD As Short = 0
    Private Const M_OPT_SELECT As Short = 1

    '勤務数(自動割当時に使用)
    Private Const M_KINMU_NUM As Short = 6

    '自動割当実行クラスオブジェクト変数
    Private m_AutoSched As Object

    '自動割当実行時ﾊﾟﾗﾒｰﾀ
    Private m_CalStartDate As Integer 'ｶﾚﾝﾀﾞｰ開始日
    Private m_CalEndDate As Integer 'ｶﾚﾝﾀﾞｰ終了日
    Private m_ScheduleStartDate As Integer '割当開始日
    Private m_ScheduleEndDate As Integer '割当終了日
    Private m_DisplayStartDate As Integer '表示期間開始日
    Private m_DisplayEndDate As Integer '表示期間終了日
    Private m_ScheduleStartCol As Integer '割当開始列
    Private m_ScheduleEndCol As Integer '割当終了列
    Private m_DisplayStartCol As Integer '表示期間開始列
    Private m_DisplayEndCol As Integer '表示期間終了列
    Private m_UserStartCol As Integer 'ﾕｰｻﾞｰ指定開始列
    Private m_UserEndCol As Integer 'ﾕｰｻﾞｰ指定終了列
    Private m_SelectDate As Long     '自動割当開始日
    '2012/11/16 Ishiga add start-------------
    Private m_PlanNO As Long     '計画番号
    '2012/11/16 Ishiga add end---------------

    '2014/06/12 TAKEBAYASHI P-07100 Add (変数の追加) START-->>
    Private m_SelTeamNo As Integer     '選択されたﾁｰﾑ番号
    Private m_OuenTeam As Integer     '応援者ﾁｰﾑ番号
    Private m_TeamCnt As Integer     'ﾁｰﾑ件数
    '2014/06/12 TAKEBAYASHI P-07100 Add (変数の追加) END<<--

    '自動割当実行済みﾌﾗｸﾞ
    Private m_SchedExecute As Boolean
    '自動割当ｷｬﾝｾﾙﾓｰﾄﾞ
    Private m_CancelMode As Boolean 'Trueの場合は自動ｼﾐｭﾚｰｼｮﾝのｷｬﾝｾﾙ
    '自動割当結果反映ﾌﾗｸﾞ（終了後割当結果を計画画面に反映させるか？）
    Private m_SchedSaveFlg As Boolean

    '割当勤務格納構造体・変数
    Private Structure KinmuMaster_Type
        Dim KinmuCD As String '勤務ｺｰﾄﾞ
        Dim KinmuName As String 'KinmuName
        Dim BunruiCD As String '分類ｺｰﾄﾞ
        Dim OnOffFlg As Boolean '選択状態ﾌﾗｸﾞ
    End Structure

    Private m_NsKinmuMData() As KinmuMaster_Type
    Private Const M_BackColor As String = "&H8080FF" '選択状態ﾊﾞｯｸｶﾗｰ
    Private Const M_NomalColor As String = "&H8000000F" '選択されていないときﾊﾞｯｸｶﾗｰ
    Private m_LoadError As Boolean
    Private m_KikanStartDate As String 'ｺﾝﾎﾞﾎﾞｯｸｽ選択開始年月日
    Private m_KikanEndDate As String 'ｺﾝﾎﾞﾎﾞｯｸｽ選択終了年月日
    Private m_SelectKikanNo As Short 'ｺﾝﾎﾞﾎﾞｯｸｽ選択期間ｲﾝﾃﾞｯｸｽ
    Private m_lstCmdKinmu As New List(Of Object)

    Private Sub GetKinmuList()
        On Error GoTo GetKinmuList
        Const W_SUBNAME As String = "NSK0000HH GetKinmuList"

        Dim w_Sql As String
        Dim w_Rs As ADODB.Recordset
        Dim w_i As Short
        Dim w_Cnt As Short
        Dim w_Code_F As ADODB.Field
        Dim w_Name_F As ADODB.Field
        Dim w_BCode_F As ADODB.Field
        Dim w_strMsg() As String
        '2017/05/02 Christopher Upd Start
        ''--- 勤務名称ﾏｽﾀ(割当対象のもの) -----
        'w_Sql = "Select KinmuCD, Name, AllocBunruiCD "
        'w_Sql = w_Sql & "From NS_KINMUNAME_M "
        'w_Sql = w_Sql & "Where AllocFlg = '2' "
        'w_Sql = w_Sql & "And HospitalCD = '" & General.g_strHospitalCD & "' "
        ''2014/06/13 TAKEBAYASHI P-07100 Add (有効期限終了日を条件に追加) START-->>
        ''ﾏｽﾀﾒﾝﾃﾅﾝｽで有効期限終了日を設定できない為(AllocFlg=2の場合)、ｺﾒﾝﾄ化(追加を取消し)
        ''w_Sql = w_Sql & "And (EFFTODATE > " & CInt(Convert.ToString(m_DisplayEndDate))
        ''w_Sql = w_Sql & "OR EFFTODATE = 0) "
        ''2014/06/13 TAKEBAYASHI P-07100 Add (有効期限終了日を条件に追加) END<<--
        'w_Sql = w_Sql & "Order By DispNo "
        ''ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ生成
        'w_Rs = General.paDBRecordSetOpen(w_Sql)

        Call NSK0000H_sql.select_NS_KINMUNAME_M_03(w_Rs)
        'Upd End
        With w_Rs

            'ﾚｺｰﾄﾞ件数を知るため最終行に移動する
            If .BOF = True And .EOF = True Then
                ReDim w_strMsg(1)
                w_strMsg(1) = "割当勤務"
                Call General.paMsgDsp("NS0010", w_strMsg)
                m_LoadError = True
                w_Cnt = 0
            Else
                'ﾚｺｰﾄﾞ件数格納
                .MoveLast()
                w_Cnt = .RecordCount
                .MoveFirst()
                w_Code_F = .Fields("KinmuCD")
                w_Name_F = .Fields("Name")
                w_BCode_F = .Fields("AllocBunruiCD")
            End If

            '配列に格納
            ReDim m_NsKinmuMData(w_Cnt)

            For w_i = 1 To w_Cnt
                m_NsKinmuMData(w_i - 1).KinmuCD = w_Code_F.Value
                m_NsKinmuMData(w_i - 1).KinmuName = w_Name_F.Value & ""
                m_NsKinmuMData(w_i - 1).BunruiCD = w_BCode_F.Value & ""
                m_NsKinmuMData(w_i - 1).OnOffFlg = False
                .MoveNext()
            Next w_i

        End With

        'ｵﾌﾞｼﾞｪｸﾄの解放
        w_Rs.Close()

        Exit Sub
GetKinmuList:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    'ﾌｫｰﾑﾛｰﾄﾞｲﾍﾞﾝﾄが正常に終了したか?
    Public ReadOnly Property pLoadState() As Boolean
        Get
            'True:正常.False:異常
            pLoadState = m_LoadError
        End Get
    End Property

    'ｶﾚﾝﾀﾞｰ終了日
    Public WriteOnly Property pAutoCalEndDate() As Integer
        Set(ByVal Value As Integer)
            m_CalEndDate = Value
        End Set
    End Property

    'ｶﾚﾝﾀﾞｰ開始日
    Public WriteOnly Property pAutoCalStartDate() As Integer
        Set(ByVal Value As Integer)
            m_CalStartDate = Value
        End Set
    End Property

    '表示期間終了列
    Public WriteOnly Property pAutoDisplayEndCol() As Integer
        Set(ByVal Value As Integer)
            m_DisplayEndCol = Value
        End Set
    End Property

    '表示期間終了日
    Public WriteOnly Property pAutoDisplayEndDate() As Integer
        Set(ByVal Value As Integer)
            m_DisplayEndDate = Value
        End Set
    End Property

    '表示期間開始列
    Public WriteOnly Property pAutoDisplayStartCol() As Integer
        Set(ByVal Value As Integer)
            m_DisplayStartCol = Value
        End Set
    End Property

    '表示期間開始日
    Public WriteOnly Property pAutoDisplayStartDate() As Integer
        Set(ByVal Value As Integer)
            m_DisplayStartDate = Value
        End Set
    End Property

    '割当終了列
    Public WriteOnly Property pAutoScheduleEndCol() As Integer
        Set(ByVal Value As Integer)
            m_ScheduleEndCol = Value
        End Set
    End Property

    '割当終了日
    Public WriteOnly Property pAutoScheduleEndDate() As Integer
        Set(ByVal Value As Integer)
            m_ScheduleEndDate = Value
        End Set
    End Property

    '割当開始列
    Public WriteOnly Property pAutoScheduleStartCol() As Integer
        Set(ByVal Value As Integer)
            m_ScheduleStartCol = Value
        End Set
    End Property

    '割当開始日
    Public WriteOnly Property pAutoScheduleStartDate() As Integer
        Set(ByVal Value As Integer)
            m_ScheduleStartDate = Value
        End Set
    End Property

    'ﾕｰｻﾞｰ指定終了列
    Public WriteOnly Property pAutoUserEndCol() As Integer
        Set(ByVal Value As Integer)
            m_UserEndCol = Value
        End Set
    End Property

    'ﾕｰｻﾞｰ指定開始列
    Public WriteOnly Property pAutoUserStartCol() As Integer
        Set(ByVal Value As Integer)
            m_UserStartCol = Value
        End Set
    End Property

    Public ReadOnly Property pSchedSaveFlg() As Boolean
        Get
            pSchedSaveFlg = m_SchedSaveFlg
        End Get
    End Property

    'ｺﾝﾎﾞﾎﾞｯｸｽ選択終了年月日
    Public WriteOnly Property pKikanEnd() As String
        Set(ByVal Value As String)
            m_KikanEndDate = Value
        End Set
    End Property

    'ｺﾝﾎﾞﾎﾞｯｸｽ選択開始年月日
    Public WriteOnly Property pKikanStart() As String
        Set(ByVal Value As String)
            m_KikanStartDate = Value
        End Set
    End Property

    '自動割当開始日
    Public ReadOnly Property p_SelectDate() As Long
        Get
            Return m_SelectDate
        End Get
    End Property

    '2014/06/12 TAKEBAYASHI P-07100 Add (ﾊﾟﾗﾒｰﾀ用ﾌﾟﾛﾊﾟﾃｨﾒｿｯﾄﾞ追加) START-->>
    '選択ﾁｰﾑ番号
    Public WriteOnly Property p_SelTeamNo() As Integer
        Set(ByVal Value As Integer)
            m_SelTeamNo = Value
        End Set
    End Property
    '応援ﾁｰﾑ番号
    Public WriteOnly Property p_OuenTeam() As Integer
        Set(ByVal Value As Integer)
            m_OuenTeam = Value
        End Set
    End Property
    'ﾁｰﾑ件数
    Public WriteOnly Property p_TeamCnt() As Integer
        Set(ByVal Value As Integer)
            m_TeamCnt = Value
        End Set
    End Property
    '2014/06/12 TAKEBAYASHI P-07100 Add (ﾊﾟﾗﾒｰﾀ用ﾌﾟﾛﾊﾟﾃｨﾒｿｯﾄﾞ追加) END<<--

    '2012/11/16 Ishiga add start----------------------------------
    '計画番号
    Public WriteOnly Property p_PlanNO() As String
        Set(ByVal Value As String)
            m_PlanNO = Value
        End Set
    End Property
    '2012/11/16 Ishiga add start----------------------------------

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub m_lstCmdKinmu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _cmdKinmu_0.Click, _cmdKinmu_1.Click, _
                                                                                                                    _cmdKinmu_2.Click, _cmdKinmu_3.Click, _
                                                                                                                    _cmdKinmu_4.Click, _cmdKinmu_5.Click
        Dim Index As Short = m_lstCmdKinmu.IndexOf(eventSender)
        Dim w_Cap As String
        Dim w_i As Short
        Dim w_Font As Font

        w_Cap = m_lstCmdKinmu(Index).Text

        For w_i = 1 To UBound(m_NsKinmuMData)
            If w_Cap = m_NsKinmuMData(w_i - 1).KinmuName Then
                If m_NsKinmuMData(w_i - 1).OnOffFlg = True Then
                    m_NsKinmuMData(w_i - 1).OnOffFlg = False
                Else
                    m_NsKinmuMData(w_i - 1).OnOffFlg = True
                End If
            End If
        Next w_i

        If ColorTranslator.ToOle(m_lstCmdKinmu(Index).BackColor) = CDbl(M_BackColor) Then
            m_lstCmdKinmu(Index).BackColor = SystemColors.Control
            w_Font = m_lstCmdKinmu(Index).Font
            m_lstCmdKinmu(Index).Font = New Font(w_Font, FontStyle.Regular)
        Else
            m_lstCmdKinmu(Index).BackColor = ColorTranslator.FromOle(CDbl(M_BackColor))
            w_Font = m_lstCmdKinmu(Index).Font
            m_lstCmdKinmu(Index).Font = New Font(w_Font, FontStyle.Bold)
        End If

    End Sub

    Private Sub cmdSchedExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSchedExit.Click
        On Error GoTo cmdSchedExit_Click
        Const W_SUBNAME As String = "Nske001 cmdSchedExit_Click"

        Dim w_Res As Short
        Dim w_Rtn As Boolean
        Dim w_strMsg() As String

        '開始ボタンを使用可能とします。
        cmdSchedStart.Enabled = True

        '閉じるボタンを使用不可とします。
        cmdSchedExit.Text = "終了(&E)"

        If m_SchedExecute <> True Then
            Me.Close()
            Exit Sub
        End If

        'ｷｬﾝｾﾙの場合
        If m_CancelMode = True Then
            '自動ｼﾐｭﾚｰｼｮﾝのｷｬﾝｾﾙ

            ReDim w_strMsg(2)
            w_strMsg(1) = "スケジューリング"
            w_strMsg(2) = "中止"
            w_Res = General.paMsgDsp("NS0097", w_strMsg)

            If w_Res = MsgBoxResult.Yes Then
                'ｷｬﾝｾﾙSWをONにする
                m_AutoSched.p_CancelSW = True
                '中止ﾒｯｾｰｼﾞを表示
                lblSchedMessage.Text = "ｽｹｼﾞｭｰﾘﾝｸﾞを中止しました．．．"
                '処理経過をクリアする
                prbSchedProcess.Minimum = 0
                prbSchedProcess.Maximum = 10
                prbSchedProcess.Value = 0
            End If

            Exit Sub

        End If

        '実行ｽﾃｰﾀｽ(ﾌﾟﾛﾊﾟﾃｨｰ)を参照し実行されていたらｲﾝﾀｰﾌｪｰｽﾃﾞｰﾀを作成
        If m_AutoSched.p_SchedStaus = "EXEC" Then
            '自動ｼﾐｭﾚｰｼｮﾝ終了ﾒｿｯﾄﾞの実行
            w_Rtn = m_AutoSched.mSchedExit
            If w_Rtn = False Then
                '自動ｼﾐｭﾚｰｼｮﾝ終了ﾒｿｯﾄﾞで実行時ｴﾗｰ発生！！
                Me.Close()
                Exit Sub
            End If

            '割当結果を保存します。
            w_Rtn = m_AutoSched.mMakeOutputIfTbl
            If w_Rtn = False Then
                '割当結果保存ﾒｿｯﾄﾞで実行時ｴﾗｰ発生！！
                Me.Close()
                Exit Sub
            End If

            '結果反映ﾌﾗｸﾞをONにする
            m_SchedSaveFlg = True
        End If

        Me.Close()

        Exit Sub
cmdSchedExit_Click:
        Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
        End

    End Sub

    Private Sub cmdSchedStart_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSchedStart.Click
        On Error GoTo cmdSchedStart_Click
        Const W_SUBNAME As String = "NSK0000HH cmdSchedStart_Click"

        '選択勤務格納ﾚｼﾞｽﾄﾘのKey
        Const M_RegKey As String = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY2 & "\" & "NSK0020B"
        Const M_RegSubKey As String = "List_Check"
        Const M_RegValue As String = "Field"

        Dim w_RegKey As String
        Dim w_i As Short
        Dim w_Res As Short
        Dim w_RegPath As String
        Dim w_Select As String
        Dim w_Rtn As Boolean
        Dim w_Date As Integer
        Dim w_SelectDate As Integer '入力された日付
        Dim w_strMsg() As String

        'Iniのｾｸｼｮﾝ名称
        w_RegPath = "NSK0000H"

        '自動シミュレーション実行確認メッセージ
        ReDim w_strMsg(2)
        w_strMsg(1) = "自動シミュレーション"
        w_strMsg(2) = "開始"
        w_Res = General.paMsgDsp("NS0097", w_strMsg)

        If w_Res = MsgBoxResult.Yes Then
            '[はい] ボタンを選択した場合

            Application.DoEvents()

            '割当開始日
            If imdDate.Text = "    /  /" Then
                ReDim w_strMsg(1)
                w_strMsg(1) = "割当開始日"
                Call General.paMsgDsp("NS0001", w_strMsg)
                imdDate.Focus()
                Exit Sub
            End If

            '入力日ﾁｪｯｸ
            w_Date = Integer.Parse(Format(CDate(imdDate.Text), "yyyyMMdd"))
            Select Case m_SelectKikanNo

                Case 1, 2
                    If m_DisplayStartDate <= w_Date And w_Date <= m_DisplayEndDate Then
                    Else
                        ReDim w_strMsg(1)
                        w_strMsg(1) = "割当開始日"
                        Call General.paMsgDsp("NS0003", w_strMsg)
                        imdDate.Focus()
                        Exit Sub
                    End If
                Case 3
                    If CDbl(m_KikanStartDate) <= w_Date And w_Date <= m_DisplayEndDate Then
                    Else
                        ReDim w_strMsg(1)
                        w_strMsg(1) = "割当開始日"
                        Call General.paMsgDsp("NS0003", w_strMsg)
                        imdDate.Focus()
                        Exit Sub
                    End If
            End Select

            '選択勤務があるの？
            w_Select = Convert.ToString(False) '選択勤務判定ﾌﾗｸﾞ初期化
            For w_i = 1 To UBound(m_NsKinmuMData)
                If m_NsKinmuMData(w_i - 1).OnOffFlg = True Then
                    '選択勤務あり
                    w_Select = Convert.ToString(True)
                    Exit For
                End If
            Next w_i

            If CBool(w_Select) = True Then

                '選択状態をﾚｼﾞｽﾄﾘに書き込み
                For w_i = 1 To UBound(m_NsKinmuMData)
                    '選択状態？
                    If m_NsKinmuMData(w_i - 1).OnOffFlg = True Then
                        '選択勤務
                        Call General.paSaveSetting(M_RegKey, M_RegSubKey, M_RegValue & Str(w_i), "1")
                    Else
                        '未選択勤務
                        Call General.paSaveSetting(M_RegKey, M_RegSubKey, M_RegValue & Str(w_i), "0")
                    End If

                Next w_i

            End If

            If CBool(w_Select) = False Then
                ReDim w_strMsg(1)
                w_strMsg(1) = "割当対象勤務"
                Call General.paMsgDsp("NS0002", w_strMsg)
                Exit Sub
            End If

            '2014/06/18 TAKEBAYASHI P-07100 Add (選択ﾁｰﾑ判定) START-->>
            If m_OuenTeam > 0 And m_OuenTeam = m_SelTeamNo Then
                ReDim w_strMsg(1)
                w_strMsg(1) = "チーム"
                Call General.paMsgDsp("NS0276", w_strMsg)
                Exit Sub
            End If
            '2014/06/18 TAKEBAYASHI P-07100 Add (選択ﾁｰﾑ判定) END<<--

            'Textに入力された日付を変数に格納
            w_SelectDate = General.paGetDateIntegerFromDate(imdDate.Value, General.G_DATE_ENUM.yyyyMMdd)
            m_SelectDate = w_SelectDate

            '開始ボタンを使用不可とします。
            cmdSchedStart.Enabled = False

            'ここから自動ｼﾐｭﾚｰｼｮﾝの各種ﾌﾟﾛﾊﾟﾃｨｰを設定します。
            '施設コード
            m_AutoSched.p_HospitalCD = General.g_strHospitalCD
            '勤務部署コード
            m_AutoSched.p_KangoTCD = General.g_strSelKinmuDeptCD
            'カレンダー開始年月日
            m_AutoSched.p_calstart_ymd = Convert.ToString(m_CalStartDate)
            'カレンダー終了年月日
            m_AutoSched.p_calend_ymd = Convert.ToString(m_CalEndDate)
            '割当開始年月日
            m_AutoSched.p_schedstart_ymd = Convert.ToString(m_ScheduleStartDate)
            '割当終了年月日
            m_AutoSched.p_schedend_ymd = Convert.ToString(m_ScheduleEndDate)
            '割当開始列
            m_AutoSched.p_schedstart_col = Convert.ToString(m_ScheduleStartCol)
            '割当終了列
            m_AutoSched.p_schedend_col = Convert.ToString(m_ScheduleEndCol)
            '入力された日付
            m_AutoSched.p_SelectDate = Convert.ToString(w_SelectDate)
            'ユーザー指定開始列
            m_AutoSched.p_usrstart_col = Convert.ToString(m_UserStartCol)
            'ユーザー指定終了列
            m_AutoSched.p_usrend_col = Convert.ToString(m_UserEndCol)

            '2014/06/12 TAKEBAYASHI P-07100 Add (ﾊﾟﾗﾒｰﾀの設定) START-->>
            m_AutoSched.p_SelTeamNo = m_SelTeamNo
            m_AutoSched.p_TeamCnt = m_TeamCnt
            '2014/06/12 TAKEBAYASHI P-07100 Add (ﾊﾟﾗﾒｰﾀの設定) END<<--

            '2012/11/16 Ishiga add start--------
            '計画番号
            m_AutoSched.p_PlanNO = m_PlanNO
            '2012/11/16 Ishiga add end----------

            '表示開始年月日
            m_AutoSched.p_dispstart_ymd = Convert.ToString(m_DisplayStartDate)
            '表示終了年月日
            m_AutoSched.p_dispend_ymd = Convert.ToString(m_DisplayEndDate)
            '表示開始列
            m_AutoSched.p_dispstart_col = Convert.ToString(m_DisplayStartCol)
            '表示終了列
            m_AutoSched.p_dispend_col = Convert.ToString(m_DisplayEndCol)

            'レジストリの取得→配列に対象／非対象を設定
            For w_i = 1 To UBound(m_NsKinmuMData)
                w_Select = General.paGetSetting(M_RegKey, M_RegSubKey, M_RegValue & Str(w_i), "0")
                If w_Select <> "0" Then
                    '2014/06/19 TAKEBAYASHI P-07100 Change (勤務ｺｰﾄﾞも文字列連結するように変更)
                    m_AutoSched.p_scheditem = m_AutoSched.p_scheditem & m_NsKinmuMData(w_i - 1).KinmuCD & m_NsKinmuMData(w_i - 1).BunruiCD
                End If
            Next w_i

            '再思考回数
            m_AutoSched.p_test_cnt = General.paGetItemValue(General.G_STRMAINKEY2, w_RegPath, "TESTCOUNT", Convert.ToString(5), General.g_strHospitalCD)
            '終了条件
            m_AutoSched.p_end_jyoken = General.paGetItemValue(General.G_STRMAINKEY2, w_RegPath, "ENDJYOUKENPOINT", Convert.ToString(30000), General.g_strHospitalCD)
            'I/F ファイルパス／ファイル名
            m_AutoSched.p_SchedDataPath = General.paGetIniSetting(General.G_STRININAME, General.G_STRSECTION1, "DataPath", "") & "Schedif.dat"
            '実行状況表示用ﾌﾟﾛｸﾞﾚｽﾊﾞｰｵﾌﾞｼﾞｪｸﾄ
            m_AutoSched.p_ObjProgressBar = prbSchedProcess
            '実行状況表示用ﾗﾍﾞﾙｵﾌﾞｼﾞｪｸﾄ
            m_AutoSched.p_ObjLabel = lblSchedMessage
            'ソート用 ﾘｽﾄﾎﾞｯｸｽｵﾌﾞｼﾞｪｸﾄ
            m_AutoSched.p_ObjListBox = Lst_SortList

            'ｷｬﾝｾﾙﾎﾞﾀﾝ制御ﾌﾗｸﾞ
            m_CancelMode = True

            '実行した
            m_SchedExecute = True

            'ここでは自動ｼﾐｭﾚｰｼｮﾝｸﾗｽの実行開始ﾒｿｯﾄﾞを実行します。ﾌﾟﾛﾊﾟﾃｨｰが正しく設定されて
            'いないと動作しません。又、ﾒｿｯﾄﾞが終了するまで制御は戻りません。
            w_Rtn = m_AutoSched.mSchedStart()

            'ｷｬﾝｾﾙﾎﾞﾀﾝ制御OFF
            m_CancelMode = False

            '自動シミュレーションの戻り値判定
            If w_Rtn = False Then
                '自動シミュレーションで実行時エラーが発生した場合

                '実行ﾌﾗｸﾞを初期化
                m_SchedExecute = False
                '割当結果反映ﾌﾗｸﾞを初期化
                m_SchedSaveFlg = False

                Me.Close()

            Else
                '開始ボタンを使用可能とします。
                cmdSchedStart.Enabled = True
            End If

            Call cmdSchedExit_Click(cmdSchedExit, New System.EventArgs())
        End If

        Exit Sub
cmdSchedStart_Click:
        Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub frmNSK0000HH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo Form_Load
        Const W_SUBNAME As String = "NSK0000HH Form_Load"

        Dim w_i As Short
        Dim w_Select As String
        Dim w_KinmuCnt As Short
        Dim w_Font As Font

        '選択勤務格納ﾚｼﾞｽﾄﾘのKey
        Const M_RegKey As String = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY2 & "\" & "NSK0020B"
        Const M_RegSubKey As String = "List_Check"
        Const M_RegValue As String = "Field"

        '初期化
        m_LoadError = False

        Call subSetCtlList()

        '割当勤務の選択ﾘｽﾄを取得します。
        Call GetKinmuList()

        '割当勤務数
        w_KinmuCnt = UBound(m_NsKinmuMData)

        '各種ﾌﾗｸﾞの初期化
        m_SchedExecute = False 'ｽｹｼﾞｭｰﾘﾝｸﾞ実行ﾌﾗｸﾞ
        m_CancelMode = False '終了ﾎﾞﾀﾝ
        m_SchedSaveFlg = False '割当結果反映ﾌﾗｸﾞ

        If m_LoadError = False Then
            '割当開始日
            m_SelectKikanNo = 0

            'ｺﾝﾎﾞﾎﾞｯｸｽの選択期間によってﾃﾞﾌｫﾙﾄの日付を変更
            If m_KikanStartDate = "0" Then
                w_Select = Convert.ToString(m_DisplayStartDate)
                m_SelectKikanNo = 1
            ElseIf CDbl(m_KikanStartDate) < m_DisplayStartDate Then
                w_Select = Convert.ToString(m_DisplayStartDate)
                m_SelectKikanNo = 2
            Else
                w_Select = m_KikanStartDate
                m_SelectKikanNo = 3
            End If

            imdDate.Text = Mid(w_Select, 1, 4) & "/" & Mid(w_Select, 5, 2) & "/" & Mid(w_Select, 7, 2)

            'ｺﾏﾝﾄﾞﾎﾞﾀﾝｷｬﾌﾟｼｮﾝ
            For w_i = 1 To M_KINMU_NUM
                If w_i <= w_KinmuCnt Then
                    m_lstCmdKinmu(w_i - 1).Visible = True
                    m_lstCmdKinmu(w_i - 1).Text = m_NsKinmuMData(w_i - 1).KinmuName
                Else
                    Exit For
                End If
            Next w_i

            For w_i = 1 To UBound(m_NsKinmuMData)
                'ﾚｼﾞｽﾄﾘより選択勤務かどうかを取得
                w_Select = General.paGetSetting(M_RegKey, M_RegSubKey, M_RegValue & Str(w_i), "0")
                If w_Select <> "0" Then
                    m_NsKinmuMData(w_i - 1).OnOffFlg = True

                    If w_i <= M_KINMU_NUM Then
                        '選択
                        m_lstCmdKinmu(w_i - 1).BackColor = ColorTranslator.FromOle(CDbl(M_BackColor))
                        w_Font = m_lstCmdKinmu(w_i - 1).Font
                        m_lstCmdKinmu(w_i - 1).Font = New Font(w_Font, FontStyle.Bold)
                    End If
                End If
            Next w_i
        End If

        'スクロールバー、オプションボタンの設定
        Select Case w_KinmuCnt
            Case 0 To M_KINMU_NUM
                For w_i = M_KINMU_NUM To (w_KinmuCnt + 1) Step -1
                    m_lstCmdKinmu(w_i - 1).Visible = False
                Next w_i
                HscKinmu.Maximum = (0 + HscKinmu.LargeChange - 1)
                HscKinmu.Visible = False
            Case Else
                For w_i = 1 To M_KINMU_NUM
                    m_lstCmdKinmu(w_i - 1).Visible = True
                    m_lstCmdKinmu(w_i - 1).Enabled = True
                Next w_i
                HscKinmu.Maximum = (w_KinmuCnt - M_KINMU_NUM + HscKinmu.LargeChange - 1)
                HscKinmu.Visible = True
                HscKinmu.Enabled = True
        End Select


        '自動ｼﾐｭﾚｰｼｮﾝｸﾗｽのｵﾌﾞｼﾞｪｸﾄ（ｲﾝｽﾀﾝｽ）を作成します。ｵﾌﾞｼﾞｪｸﾄ変数は
        'ﾓｼﾞｭｰﾙﾚﾍﾞﾙで宣言しています。ｲﾝｽﾀﾝｽの開放はﾌｫｰﾑｱﾝﾛｰﾄﾞ時に行っています。
        m_AutoSched = New NsAid_NSK0020B.ClsAutoSched

        '接続ｵﾌﾞｼﾞｪｸﾄ渡し

        'Inatalltype渡し
        m_AutoSched.pInstallType = General.g_InstallType
        'マスタ取得部品
        m_AutoSched.pGetMasterObj = General.g_objGetMaster

        'ｳｨﾝﾄﾞｳを中央に配置します。
        Me.StartPosition = FormStartPosition.CenterScreen

        Exit Sub
Form_Load:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub frmNSK0000HH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo Form_Unload
        Const W_SUBNAME As String = "NSK0000HH Form_Unload"

        '自動ｼﾐｭﾚｰｼｮﾝｸﾗｽのｲﾝｽﾀﾝｽを破棄します。
        m_AutoSched = Nothing

        Exit Sub
Form_Unload:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub HscKinmu_Change(ByVal newScrollValue As Integer)
        On Error GoTo HscKinmu_Change
        Const W_SUBNAME As String = "Nskk001d  HscKinmu_Change"

        'スクロールバーの更新
        Dim w_i As Short
        Dim w_Hsc_Cnt As Short
        Dim w_KinmuCnt As Short
        Dim w_Font As Font

        'コマンドボタンのＣＡＰＴＩＯＮ設定

        w_KinmuCnt = UBound(m_NsKinmuMData)

        '勤務
        w_Hsc_Cnt = newScrollValue
        For w_i = 1 To M_KINMU_NUM
            If w_i + w_Hsc_Cnt <= w_KinmuCnt Then
                m_lstCmdKinmu(w_i - 1).Text = m_NsKinmuMData(w_i + w_Hsc_Cnt - 1).KinmuName
                If m_NsKinmuMData(w_i + w_Hsc_Cnt - 1).OnOffFlg = True Then
                    m_lstCmdKinmu(w_i - 1).BackColor = ColorTranslator.FromOle(CDbl(M_BackColor))
                    w_Font = m_lstCmdKinmu(w_i - 1).Font
                    m_lstCmdKinmu(w_i - 1).Font = New Font(w_Font, FontStyle.Bold)
                Else
                    m_lstCmdKinmu(w_i - 1).BackColor = SystemColors.Control
                    w_Font = m_lstCmdKinmu(w_i - 1).Font
                    m_lstCmdKinmu(w_i - 1).Font = New Font(w_Font, FontStyle.Regular)
                End If
            Else
                m_lstCmdKinmu(w_i - 1).Text = ""
            End If
        Next w_i

        Exit Sub
HscKinmu_Change:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub

    Private Sub HscKinmu_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscKinmu.Scroll
        HscKinmu_Change(eventArgs.NewValue)
    End Sub

    Private Sub subSetCtlList()
        m_lstCmdKinmu.Add(_cmdKinmu_0)
        m_lstCmdKinmu.Add(_cmdKinmu_1)
        m_lstCmdKinmu.Add(_cmdKinmu_2)
        m_lstCmdKinmu.Add(_cmdKinmu_3)
        m_lstCmdKinmu.Add(_cmdKinmu_4)
        m_lstCmdKinmu.Add(_cmdKinmu_5)
    End Sub
End Class