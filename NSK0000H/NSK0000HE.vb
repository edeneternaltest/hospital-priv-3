Option Strict Off
Option Explicit On
Imports System.Text
Public Class frmNSK0000HE
    Inherits General.FormBase
    '/----------------------------------------------------------------------/
    '/
    '/    ｼｽﾃﾑ名称：看護支援システム(勤務管理)
    '/ ﾌﾟﾛｸﾞﾗﾑ名称：保存呼出画面
    '/        ＩＤ：NSK0000HE
    '/        概要：一時保存予定一覧Ｆの一覧表示・保存・呼出を行う。
    '/
    '/
    '/      作成者： Angelo     CREATE 2017/08/04     REV 01.00
    '/      更新者：            UPDATE     /  /      【 】
    '/                                更新内容：( )
    '/
    '/     Copyright (C) Inter co.,ltd 2000
    '/----------------------------------------------------------------------/
    '=======================================================
    '   定数宣言
    '=======================================================
    'スプレッドの列INDEX
    Private Const M_SAVESPR_COLIDX_NO As Integer = 0 '保存番号
    Private Const M_SAVESPR_COLIDX_NAME As Integer = 1 '操作者のユーザー
    Private Const M_SAVESPR_COLIDX_DATE As Integer = 2 '最終更新日時
    Private Const M_SAVESPR_COLIDX_BIKOU As Integer = 3 '備考

    Private Const M_SAVESPR_MAXROW As Integer = 5 '最大表示行数
    '=======================================================
    '   ﾌﾟﾗｲﾍﾞｰﾄ変数
    '=======================================================
    Private m_intIndexPreRow As Short '現在選択されている行
    Private m_intDefPlanNo As Short '表示計画期間の計画番号
    Private m_intSaveNo As Short '保存番号
    Private m_StaffData() As StaffData_Type '対象職員情報

    Private m_intPlanStartDate As Integer '開始日
    Private m_intPlanEndDate As Integer '終了日

    Private m_ApplyEndFlg As Integer '適用ボタン終了フラグ（True:適用ボタン押下時）
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
        Dim intSaveNo As Short '保存番号
        Dim strBikou As String '備考
        Dim dblLastUpdTimeDate As Double '最終更新日時
        Dim strRegistID As String '操作者のユーザーＩＤ
        Sub init(ByVal p_SaveNo As Short)
            intSaveNo = p_SaveNo
            strBikou = ""
            dblLastUpdTimeDate = 0
            strRegistID = ""
        End Sub
    End Structure

    Private m_udtSaveYotei() As Save_Type '一時保存予定一覧

    '=======================================================
    '   Getter/Setter
    '=======================================================
    ''' <summary>計画期間情報を受け取る</summary>
    ''' <param name="p_FromYMD">表示開始日</param>
    ''' <param name="p_ToYMD">表示終了日</param>
    ''' <param name="Value">計画番号</param>
    ''' <remarks></remarks>
    Public WriteOnly Property pPlanInfo(ByVal p_FromYMD As String, ByVal p_ToYMD As String) As Short
        Set(ByVal Value As Short)
            m_intPlanStartDate = p_FromYMD '開始日
            m_intPlanEndDate = p_ToYMD '終了日
            m_intDefPlanNo = Value '計画番号
        End Set
    End Property

    ''' <summary>職員情報を受け取る</summary>
    ''' <param name="Value">職員情報</param>
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

    ''' <summary>適用ボタン終了フラグを上位画面に引き渡す</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property pTekiyouFlg() As Boolean
        Get
            pTekiyouFlg = m_ApplyEndFlg
        End Get
    End Property

    ''' <summary>選択された保存番号を上位画面に引き渡す</summary>
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
    ''' frmNSK0000HEフォームLoadイベント
    ''' </summary>
    ''' <param name="eventSender">System.Object</param>
    ''' <param name="eventArgs">System.EventArgs</param>
    ''' <remarks>frmNSK0000HEをLoadし、表示する</remarks>
    Private Sub frmNSK0000HE_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Const W_SUBNAME As String = "NSK0000HE Form_Load"
        Dim w_numWidth As Integer

        Try
            w_numWidth = sprSaveList_Sheet1.Columns(0).Width +
                                sprSaveList_Sheet1.Columns(1).Width +
                                sprSaveList_Sheet1.Columns(2).Width +
                                sprSaveList_Sheet1.Columns(3).Width

            'スプレッドの設定
            General.paSpreadSizeFit(sprSaveList,
                                    sprSaveList.VerticalScrollBarPolicy,
                                    sprSaveList.HorizontalScrollBarPolicy,
                                    M_SAVESPR_MAXROW,
                                    w_numWidth)

            '一時保存一覧取得
            Call GetSaveData()

            '一時保存一覧表示
            Call SetSprData()

            If m_intSaveNo = 0 Then
                '初期化
                m_intSaveNo = 1
                Call SetSelectData(0)
            Else
                '保存後選択されている行目を選択
                Call SetSelectData(m_intIndexPreRow)
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    ''' <summary>
    ''' 一時保存一覧の取得
    ''' 概要:勤務計画画面を開いている部署・計画期間で取得。
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GetSaveData()
        Const W_SUBNAME As String = "NSK0000HE GetStoredData"

        Dim w_sbSql As New StringBuilder 'SQL文
        Dim w_objRs As ADODB.Recordset 'RecordSet ｵﾌﾞｼﾞｪｸﾄ
        Dim w_objFields As ADODB.Fields 'ﾌｨｰﾙﾄﾞ ｵﾌﾞｼﾞｪｸﾄ
        Dim w_intDataCount As Short
        Dim w_intDataLoop As Short
        Dim w_intRowIdx As Short

        Try
            '一時保存予定一覧の初期化
            ReDim m_udtSaveYotei(M_SAVESPR_MAXROW)
            For w_intRowIdx = 1 To M_SAVESPR_MAXROW
                m_udtSaveYotei(w_intRowIdx).init(w_intRowIdx)
            Next

            'Select文 編集 
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

            'カーソル作成
            w_objRs = General.paDBRecordSetOpen(w_sbSql.ToString())

            With w_objRs
                If .RecordCount <= 0 Then
                    'ﾃﾞｰﾀが存在しないとき
                Else
                    'ﾃﾞｰﾀが存在するとき
                    .MoveLast()
                    'ﾃﾞｰﾀ件数取得
                    w_intDataCount = .RecordCount
                    .MoveFirst()
                    'ﾌｨｰﾙﾄﾞｵﾌﾞｼﾞｪｸﾄ生成
                    w_objFields = .Fields

                    'ﾃﾞｰﾀ件数Loop
                    For w_intDataLoop = 1 To w_intDataCount
                        '保存番号に対応する行に一時保存のデータを設定する
                        w_intRowIdx = Short.Parse(General.paGetDbFieldVal(w_objFields("SAVENO"), 0))
                        '保存番号取得
                        m_udtSaveYotei(w_intRowIdx).intSaveNo = General.paGetDbFieldVal(w_objFields("SAVENO"), 0)
                        '備考取得
                        m_udtSaveYotei(w_intRowIdx).strBikou = General.paGetDbFieldVal(w_objFields("BIKOU"), "")
                        '最終更新日時取得
                        m_udtSaveYotei(w_intRowIdx).dblLastUpdTimeDate = General.paGetDbFieldVal(w_objFields("LASTUPDTIMEDATE"), 0)
                        '操作者のユーザーＩＤ取得
                        m_udtSaveYotei(w_intRowIdx).strRegistID = General.paGetDbFieldVal(w_objFields("REGISTRANTID"), "")

                        .MoveNext()
                    Next w_intDataLoop
                End If
            End With

            'ｽﾄﾘﾝｸﾞﾋﾞﾙﾀﾞ解放
            w_sbSql.Clear()
            'ｵﾌﾞｼﾞｪｸﾄの解放
            w_objRs = Nothing
            w_objFields = Nothing
        Catch ex As Exception
            'ｽﾄﾘﾝｸﾞﾋﾞﾙﾀﾞ解放
            w_sbSql.Clear()
            'ｵﾌﾞｼﾞｪｸﾄの解放
            w_objRs = Nothing
            w_objFields = Nothing

            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '一時保存一覧の表示
    Private Sub SetSprData()
        Const W_SUBNAME As String = "NSK0000HE SetSprData"

        Dim w_intRowIdx As Short
        Dim w_intSprLoop As Short
        Dim w_intSprRowCount As Short

        Try
            With sprSaveList_Sheet1
                'スプレッドの内容をクリアする
                'データのリセット
                sprSaveList_Sheet1.ClearRange(0, 0, sprSaveList_Sheet1.RowCount, sprSaveList_Sheet1.ColumnCount, False)

                'スタイルの適用
                subSetStyles()

                'スプレッドの最大行数制御
                .RowCount = M_SAVESPR_MAXROW

                'spread行は0から始まるｲﾝﾃﾞｯｸｽ
                w_intSprRowCount = .RowCount - 1
                For w_intSprLoop = 0 To w_intSprRowCount
                    w_intRowIdx = w_intSprLoop + 1
                    '保存番号
                    .Cells(w_intSprLoop, M_SAVESPR_COLIDX_NO).Text = EditData(m_udtSaveYotei(w_intRowIdx).intSaveNo, G_EDITMODE_NO)
                    '操作者のユーザーＩＤ
                    .Cells(w_intSprLoop, M_SAVESPR_COLIDX_NAME).Text = m_udtSaveYotei(w_intRowIdx).strRegistID
                    '最終更新日時
                    .Cells(w_intSprLoop, M_SAVESPR_COLIDX_DATE).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
                    .Cells(w_intSprLoop, M_SAVESPR_COLIDX_DATE).Text = EditData(m_udtSaveYotei(w_intRowIdx).dblLastUpdTimeDate, G_EDITMODE_DATETIME)
                    '備考
                    .Cells(w_intSprLoop, M_SAVESPR_COLIDX_BIKOU).Text = m_udtSaveYotei(w_intRowIdx).strBikou
                Next
            End With
        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    'スタイルの設定を行う関数
    Private Sub subSetStyles()
        Const W_SUBNAME As String = "NSK0000HE subSetStyles"

        Dim w_style As New FarPoint.Win.Spread.StyleInfo()
        Dim w_Font As New System.Drawing.Font("ＭＳ ゴシック", 10.0!)
        Dim w_TextCellType As New FarPoint.Win.Spread.CellType.TextCellType

        Try
            'スタイルの適用
            w_style.Font = w_Font
            w_style.CellType = w_TextCellType
            w_style.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Left
            w_style.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
            w_style.BackColor = Color.White

            'スプレッド初期化
            sprSaveList_Sheet1.Models.Style.SetDirectInfo(-1, -1, w_style)

        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    ''' <summary>
    ''' Spreadセル押下時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="EventArgs"></param>
    ''' <remarks></remarks>
    Private Sub sprSaveList_CellClick(ByVal sender As Object, ByVal EventArgs As FarPoint.Win.Spread.CellClickEventArgs) Handles sprSaveList.CellClick

        Try
            'ヘッダクリック時は処理を抜ける
            If EventArgs.ColumnHeader Then
                Exit Sub
            End If

            '選択行の背景色変更
            Call SetSelectData(EventArgs.Row)

            '保存番号
            m_intSaveNo = m_intIndexPreRow + 1 'CellClickEventArgs.rowは0から始まるｲﾝﾃﾞｯｸｽ
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 選択行の背景色変更
    ''' </summary>
    ''' <param name="p_Row">選択行</param>
    ''' <remarks></remarks>
    Private Sub SetSelectData(ByVal p_Row As Integer)
        Const W_SUBNAME As String = "NSK0000HE SetSelectData"

        Try
            If p_Row >= 0 Then
                With sprSaveList_Sheet1
                    'スプレッドの表示の更新は一括で
                    '全体の表示の初期化
                    .Rows(0, .RowCount - 1).BackColor = Color.White
                    .Rows(0, .RowCount - 1).ForeColor = Color.Black

                    If p_Row > -1 Then
                        '指定された行の背景を変更
                        .Rows(p_Row).BackColor = Color.Cyan
                        .Rows(p_Row).ForeColor = Color.Black
                    End If

                    '一時保存一覧spreadに保存者が存在するかどうかのチェック(保存時に名称が必要です)
                    '選択した行にデータがある場合
                    If Not .Cells(p_Row, M_SAVESPR_COLIDX_NAME).Text = "" Then
                        '備考の内容をテキストボックス「備考」に出力
                        '適用ボタンを活性化
                        txtBikou.Text = .Cells(p_Row, M_SAVESPR_COLIDX_BIKOU).Text
                        cmdApply.Enabled = True
                    Else '選択した行にデータがない場合
                        'テキストボックス「備考」をクリア
                        '適用ボタンを非活性化
                        txtBikou.Text = ""
                        cmdApply.Enabled = False
                    End If
                End With

                '今回選択された行を確保
                m_intIndexPreRow = p_Row
            End If
        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    ''' <summary>
    ''' cmdSaveボタンClickイベント
    ''' </summary>
    ''' <param name="eventSender">System.Object</param>
    ''' <param name="eventArgs">System.EventArgs</param>
    ''' <remarks>一時保存データの削除・保存・再読込</remarks>
    Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc 'ﾌﾟﾛｼｰｼﾞｬ名の一時待避
        General.g_ErrorProc = "NSK0000HA cmdSave_Click"

        Try
            'ﾄﾗﾝｻﾞｸｼｮﾝ開始
            Call General.paBeginTrans()

            'ﾃﾞｰﾀ削除
            Call DeleteSaveData()

            'ﾃﾞｰﾀ書込み
            If InsertSaveData() Then
                'ｺﾐｯﾄ
                Call General.paCommit()

                '一時保存データの再読込
                Call frmNSK0000HE_Load(eventSender, eventArgs)
            Else
                Call General.paRollBack()
            End If

            General.g_ErrorProc = w_PreErrorProc '待避ﾌﾟﾛｼｰｼﾞｬ名を元に戻す
        Catch ex As Exception
            Call General.paRollBack()
            Call General.paTrpMsg(Convert.ToString(Err.Number), General.g_ErrorProc)
            End
        End Try
    End Sub

    ''' <summary>
    ''' 一時保存データの削除
    ''' </summary>
    ''' <remarks>
    ''' 以下のテーブルからデータを削除する。
    '''    ・一時保存予定一覧Ｆ
    '''    ・一時保存勤務予定Ｆ
    '''    ・一時保存勤務詳細Ｆ
    '''    ・一時保存年休Ｆ
    ''' </remarks>
    Private Sub DeleteSaveData()
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc 'ﾌﾟﾛｼｰｼﾞｬ名の一時待避
        General.g_ErrorProc = "NSK0000HE DeleteSaveData"

        Dim w_sbSql As New StringBuilder 'SQL文
        'テーブル名のリスト
        Dim w_arrTempSaveTables() As String = {"NS_TEMPPLANLIST_F", "NS_TEMPKINMUPLAN_F",
                                                "NS_TEMPKINMUDETAIL_F", "NS_TEMPNENKYU_F"}
        Try
            'すべてのテーブルを列挙する
            With w_sbSql
                For Each table As String In w_arrTempSaveTables
                    .AppendLine("DELETE FROM " & table)
                    .AppendLine("WHERE")
                    .AppendLine("    HOSPITALCD   = '" & General.g_strHospitalCD & "'")
                    .AppendLine("AND PLANNO      >=  " & m_intDefPlanNo)
                    .AppendLine("AND KINMUDEPTCD <= '" & General.g_strSelKinmuDeptCD & "'")
                    .AppendLine("AND SAVENO       =  " & m_intSaveNo)
                    '更新実行
                    Call General.paDBExecute(.ToString)
                    Call .Clear()
                Next table
            End With

            General.g_ErrorProc = w_PreErrorProc '待避ﾌﾟﾛｼｰｼﾞｬ名を元に戻す
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 一時保存データの保存
    ''' </summary>
    ''' <remarks>
    ''' 以下のテーブルにデータを登録する。
    '''    ・一時保存予定一覧Ｆ
    '''    ・一時保存勤務予定Ｆ
    '''    ・一時保存勤務詳細Ｆ
    '''    ・一時保存年休Ｆ
    '''    ・一時保存代休Ｆ
    ''' </remarks>
    Private Function InsertSaveData() As Boolean
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc 'ﾌﾟﾛｼｰｼﾞｬ名の一時待避
        General.g_ErrorProc = "NSK0000HE InsertSaveData"

        Dim w_sbSql As New StringBuilder        'SQL文
        Dim w_strSysDate As String              '更新日時
        Dim w_StaffMngID As String = String.Empty
        Dim w_strDate As String = String.Empty
        Dim w_KinmuCD As String = String.Empty  'KinmuCD
        Dim w_RiyuKBN As String = String.Empty  '理由区分
        Dim w_KangoCD As String = String.Empty  '応援看護単位CD
        Dim w_Time As String = String.Empty     '時間数
        Dim w_Comment As String = String.Empty
        Dim w_Nenkyu() As NenkyuDetail_Type
        Dim w_strMsg() As String
        Try
            '登録する日付を取得
            w_strSysDate = Format(Now, "yyyyMMddHHmmss")

            '*********************************************************************************************************'
            '                                    一時保存予定一覧Ｆの更新     
            '*********************************************************************************************************'
            'Insert文 編集 
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

            '更新実行
            Call General.paDBExecute(w_sbSql.ToString())

            'ｽﾄﾘﾝｸﾞﾋﾞﾙﾀﾞ解放
            w_sbSql.Clear()

            For i As Integer = 1 To UBound(m_StaffData)
                If General.g_lngDaikyuMng = 0 Then
                    '*********************************************************************************************************'
                    '                                    一時保存代休Ｆの更新     
                    '*********************************************************************************************************'
                    For j As Integer = 1 To UBound(m_StaffData(i).Daikyu)
                        For k As Integer = 1 To UBound(m_StaffData(i).Daikyu(j).DaikyuDetail)
                            If m_intPlanStartDate <= m_StaffData(i).Daikyu(j).DaikyuDetail(k).DaikyuDate AndAlso
                               m_StaffData(i).Daikyu(j).DaikyuDetail(k).DaikyuDate <= m_intPlanEndDate Then
                                '&1が&2を&3している場合は&4できません。~n&2を&5してください。
                                ReDim w_strMsg(5)
                                w_strMsg(1) = "職員"
                                w_strMsg(2) = "代休"
                                w_strMsg(3) = "取得"
                                w_strMsg(4) = "一時保存"
                                w_strMsg(5) = "削除"
                                Call General.paMsgDsp("NS0412", w_strMsg)
                                General.g_ErrorProc = w_PreErrorProc '待避ﾌﾟﾛｼｰｼﾞｬ名を元に戻す
                                Return False
                            End If
                        Next k
                        If m_intPlanStartDate <= m_StaffData(i).Daikyu(j).HolDate AndAlso m_StaffData(i).Daikyu(j).HolDate <= m_intPlanEndDate Then
                            'ｽﾄﾘﾝｸﾞﾋﾞﾙﾀﾞ解放
                            w_sbSql.Clear()
                            'Insert文 編集 
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
                            '更新実行
                            Call General.paDBExecute(w_sbSql.ToString)
                            'ｽﾄﾘﾝｸﾞﾋﾞﾙﾀﾞ解放
                            w_sbSql.Clear()
                        End If
                    Next j
                End If
                '*********************************************************************************************************'
                '                                    一時保存勤務詳細Ｆの更新     
                '*********************************************************************************************************'
                For j As Integer = 1 To UBound(m_StaffData(i).Kojyo)
                    For k As Integer = 1 To UBound(m_StaffData(i).Kojyo(j).lngKinmuDetailTime)
                        'ｽﾄﾘﾝｸﾞﾋﾞﾙﾀﾞ解放
                        w_sbSql.Clear()
                        'Insert文 編集 
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
                        '更新実行
                        Call General.paDBExecute(w_sbSql.ToString)
                        'ｽﾄﾘﾝｸﾞﾋﾞﾙﾀﾞ解放
                        w_sbSql.Clear()
                    Next k
                Next j
            Next i

            For w_Row As Integer = m_StaffRowStRow To m_StaffRowEdRow - (m_OuenStaffCnt * m_MaxShowLine)
                If IsDataRowAndGetMngID(w_Row, w_StaffMngID) Then
                    For w_Col As Integer = m_KinmuDataStCol To m_KinmuDataEdCol
                        If IsDataColAndGetKinmuData(w_Row, w_Col, w_strDate, w_KinmuCD, w_RiyuKBN, w_KangoCD, w_Time, w_Comment) Then
                            '*********************************************************************************************************'
                            '                                    一時保存勤務予定Ｆの更新     
                            '*********************************************************************************************************'
                            'ｽﾄﾘﾝｸﾞﾋﾞﾙﾀﾞ解放
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

                            '更新実行
                            Call General.paDBExecute(w_sbSql.ToString)
                            'ｽﾄﾘﾝｸﾞﾋﾞﾙﾀﾞ解放
                            Call w_sbSql.Clear()

                            '*********************************************************************************************************'
                            '                                    一時保存年休Ｆの更新     
                            '*********************************************************************************************************'
                            ReDim w_Nenkyu(0)
                            If ExistsNenkyuAndGetNenkyuData(w_KinmuCD, w_Time, w_Nenkyu) Then
                                For i As Integer = 1 To UBound(w_Nenkyu)
                                    'ｽﾄﾘﾝｸﾞﾋﾞﾙﾀﾞ解放
                                    w_sbSql.Clear()
                                    'Insert文 編集 
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
                                        .AppendLine(" '" & General.g_strHospitalCD & "'")       '施設CD
                                        .AppendLine(", " & m_intDefPlanNo)                      '表示計画期間の計画番号
                                        .AppendLine(",'" & General.g_strSelKinmuDeptCD & "'")   '選択勤務部署CD
                                        .AppendLine(", " & m_intSaveNo)                         '保存番号
                                        .AppendLine(",'" & w_StaffMngID & "'")                  '職員管理番号
                                        .AppendLine(", " & w_strDate)                           '日付
                                        .AppendLine(", " & i)                                   'SEQ
                                        .AppendLine(",'" & w_Nenkyu(i).GetContentsKbn & "'")    '取得内容区分
                                        .AppendLine(",'" & w_Nenkyu(i).HolidayBunruiCD & "'")   '休み分類CD
                                        .AppendLine(", " & w_Nenkyu(i).FromTime)                '開始時間
                                        .AppendLine(", " & w_Nenkyu(i).ToTime)                  '終了時間
                                        .AppendLine(",'" & w_Nenkyu(i).DateKbn & "'")           '翌日FLG
                                        .AppendLine(", " & w_Nenkyu(i).NenkyuTime)              '時間年休
                                        .AppendLine(",'" & w_Nenkyu(i).HolSubFlg & "'")         '休憩減算フラグ
                                        .AppendLine(", " & w_Nenkyu(i).DayTime)                 '日勤時間
                                        .AppendLine(", " & w_Nenkyu(i).NightTime)               '夜勤時間
                                        .AppendLine(", " & w_Nenkyu(i).NextNightTime)           '翌日夜勤時間
                                        .AppendLine(", " & w_strDate)                           '勤務年月日
                                        .AppendLine(",'" & w_Nenkyu(i).DateKbn & "'")           '年月日区分
                                        .AppendLine(",''")                                      '年休UNIQUESEQNO
                                        .AppendLine(",'1'")                                     '承認済みFLG
                                        .AppendLine(",''")                                      '削除FLG
                                        .AppendLine(", " & w_strSysDate)                        '初回登録日時
                                        .AppendLine(", " & w_strSysDate)                        '最終更新日時
                                        .AppendLine(",'" & General.g_strUserID & "')")          '登録者ID
                                    End With
                                    '更新実行
                                    Call General.paDBExecute(w_sbSql.ToString)
                                    'ｽﾄﾘﾝｸﾞﾋﾞﾙﾀﾞ解放
                                    Call w_sbSql.Clear()
                                Next i
                            End If
                        End If
                    Next w_Col
                End If
            Next w_Row

            General.g_ErrorProc = w_PreErrorProc '待避ﾌﾟﾛｼｰｼﾞｬ名を元に戻す
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
    ''' cmdApplyボタンClickイベント
    ''' </summary>
    ''' <param name="eventSender">System.Object</param>
    ''' <param name="eventArgs">System.EventArgs</param>
    ''' <remarks></remarks>
    Private Sub cmdApply_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdApply.Click
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc 'ﾌﾟﾛｼｰｼﾞｬ名の一時待避
        General.g_ErrorProc = "NSK0000HE cmdApply_Click"
        Try
            If General.paMsgDsp("NS0097", New String() {"", "編集中の勤務", "破棄"}) = MsgBoxResult.Yes Then

                m_ProgressForm = New frmNSK0000HM
                m_ProgressForm.pNumberDisp = False
                Call m_ProgressForm.Show(pProcessObj)
                m_ProgressForm.pForeColor = ColorTranslator.ToOle(Color.Black)
                m_ProgressForm.pSyoriText = "終了処理中..."
                m_ProgressForm.pMaxValue = 3
                m_ProgressForm.pCountValue = 0

                'ｵﾌﾞｼﾞｪｸﾄの解放
                Erase m_StaffData
                Erase m_udtSaveYotei

                '適用
                m_ApplyEndFlg = True

                'ﾌｫｰﾑ ｱﾝﾛｰﾄﾞ
                Me.Close()
            End If
            General.g_ErrorProc = w_PreErrorProc '待避ﾌﾟﾛｼｰｼﾞｬ名を元に戻す
        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), General.g_ErrorProc)
            End
        End Try
    End Sub

    ''' <summary>
    ''' cmdCloseボタンClickイベント
    ''' </summary>
    ''' <param name="eventSender">System.Object</param>
    ''' <param name="eventArgs">System.EventArgs</param>
    ''' <remarks>画面を閉じる</remarks>
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        Const W_SUBNAME As String = "NSK0000HE cmdClose_Click"
        Try
            'ｵﾌﾞｼﾞｪｸﾄの解放
            Erase m_StaffData
            Erase m_udtSaveYotei

            '閉じる
            m_ApplyEndFlg = False

            '何もしないで閉じる
            Me.Close()
        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub
End Class