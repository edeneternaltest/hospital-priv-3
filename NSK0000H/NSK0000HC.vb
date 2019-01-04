Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Imports FarPoint.Win.Spread

Friend Class frmNSK0000HC
    Inherits General.FormBase
    Public NSK0000H_sql As New clsNSK0000H_sql
    '=======================================================
    '   定  数  宣  言
    '=======================================================
    'ﾊﾟﾚｯﾄ勤務最大表示数
    Private Const M_PARET_NUM As Short = 50 '勤務・休み・特殊用
    Private Const M_PARET_NUM_SET As Short = 25 'セット勤務用
	'勤務記号ﾃﾞｰﾀ 開始ｾﾙ 位置
    Private Const M_KinmuData_Row As Integer = 3
    Private Const M_KinmuData_Row_ChgJisseki As Integer = 4 '勤務変更画面よりﾛｰﾄﾞされた場合の修正可能行
    Private Const M_KinmuData_Col As Integer = 2
    '勤務入力不可能（前月実績・部署異動）
	Private m_MonthBefore_Fore As Integer '文字色
	Private m_MonthBefore_Back As Integer '背景色
	'計画期間外の4週部分
	Private m_Jisseki4W_Back As Integer '背景色
	'予定⇒実績変更勤務
    Private m_Comp_Fore As Integer '文字色
    Private m_WeekEnd_Back As Integer
    Private m_WeekEndColorFlg As String '土日背景色フラグ
    Private m_HolidayColorFlg As String '祝休日背景色フラグ
	'-------------------------------------------------------
	'  ﾌﾟﾗｲﾍﾞｰﾄ変数
	'-------------------------------------------------------
	'計画/実績画面ﾓｰﾄﾞの取得ﾓｰﾄﾞ("新規"，"計画変更"，"勤務変更")
	Private m_Mode As String
    '一段/二段表示ﾓｰﾄﾞの取得ﾓｰﾄﾞ("二段表示"，"一段表示")
	Private m_blnOneTwo As Boolean
	'職員ﾃﾞｰﾀ開始行
	Private m_StaffStartRow As Integer
	'最大表示段数
	Private m_MaxShowLine As Short
	'日当直表示行
	Private m_DutyData As Short
	'勤務予定表示行
	Private m_KinmuPlan As Short
	'勤務実績表示行
	Private m_KinmuJisseki As Short
	'届出表示行
	Private m_AppliData As Short
    '計画単位･表示期間の取得(1:４週／１ヶ月  2:４週／１ヶ月以外)
	Private m_DispKikan As String
	'更新ﾌﾗｸﾞ
	Private m_KosinFlg As Boolean
	'O.K.ﾎﾞﾀﾝ押下ﾌﾗｸﾞ
	Private m_OKFlg As Boolean
	'計画ﾃﾞｰﾀ 開始/終了 列番号情報
	Private m_KeikakuD_StartCol_Param As Integer
	Private m_KeikakuD_EndCol_Param As Integer
	'親計画ﾃﾞｰﾀ 開始/終了 列番号情報
	Private m_KeikakuD_StartCol As Integer
	Private m_KeikakuD_EndCol As Integer
	'親ｶﾚﾝﾄ行 位置
	Private m_CUR_ROW_Param As Integer
	'親ｺﾝﾄﾛｰﾙ
    Private m_Control_Param As FarPoint.Win.Spread.FpSpread
	'勤務用
	Private m_Kinmu() As KinmuM_Type
	Private m_KinmuCnt As Short
	'休み用
	Private m_Yasumi() As KinmuM_Type
	Private m_YasumiCnt As Short
	'特殊勤務用
	Private m_Tokusyu() As KinmuM_Type
	Private m_TokusyuCnt As Short
	'代休取得関連変数
	Private m_DaikyuMsgFlg As Integer '代休取得時の確認メッセージを表示するか(0:表示,1:非表示)
	Private m_SundayDaikyuFlg As Integer '代休取得可能日に日曜日を含めるか(0:含めない,1:含める)
	Private m_DaikyuAdvFlg As Integer '代休の先取りを可能にするか(0:しない,1:する)
    Private m_SaturdayDaikyuFlg As Integer '代休取得可能日に土曜日を含めるか(0:含めない,1:含める)
    Private m_DaikyuAdvThisMonthFlg As Integer '代休先取り当月制限フラグ(0:OFF,1:ON)
    Private m_OuenDispFlg As Integer '応援勤務区分のラジオボタンをパレットに表示するか(1:しない,0:する)
    '各日にち毎のﾃﾞｰﾀについてのﾌﾗｸﾞ（配列として受け取る）
	Private m_DataFlg As Object 'ﾃﾞｰﾀﾌﾗｸﾞ（"0":計画ﾃﾞｰﾀ，"1":実績ﾃﾞｰﾀ，その他:ﾃﾞｰﾀなし）
	Private m_KakuteiFlg As Object '確定ﾌﾗｸﾞ（"0":該当部署確定ﾃﾞｰﾀ，"1":他部署確定ﾃﾞｰﾀ）
	Private m_DataHideFlg As Object 'ﾃﾞｰﾀﾌﾗｸﾞ（"0":計画ﾃﾞｰﾀ，"1":実績ﾃﾞｰﾀ，その他:ﾃﾞｰﾀなし）
	Private m_KakuteiHideFlg As Object '確定ﾌﾗｸﾞ（"0":該当部署確定ﾃﾞｰﾀ，"1":他部署確定ﾃﾞｰﾀ）
    Private Const M_MenuKibouChk As Short = 4 '希望勤務への入力警告

    '2014/04/23 Saijo add start P-06979-------------------------------------------------------------------
    Private m_strKinmuEmSecondFlg As String '勤務記号全角２文字対応フラグ(0：対応しない、1:対応する)
    '2014/04/23 Saijo add end P-06979---------------------------------------------------------------------

    '2015/04/14 Bando Add Start ========================
    Private m_DispKinmuCd As String '希望モード時の表示対象勤務CD
    '2015/04/14 Bando Add End   ========================

    '代休詳細用構造体
    Private Structure DaikyuDetail_Type
        Dim DaikyuDate As Integer '代休取得日
        Dim DaikyuKinmuCD As String '代休取得勤務ＣＤ
        Dim GetFlg As String '代休取得タイプ(0:1日代休、1:0.5日代休)
    End Structure

    Private Structure Daikyu_Type
        Dim HolDate As Integer
        Dim HolKinmuCD As String
        Dim DaikyuDetail() As DaikyuDetail_Type
        Dim GetKbn As String '代休発生量タイプ(0:1日分,1:1.5日分)
        Dim RemainderHol As Double '代休未使用分
        Dim OutPutList As String '代休取得時にリストにあげるかどうかのﾌﾗｸﾞ("0":代休取得日非対象 "1":代休取得日対象)
    End Structure

    Private m_DaikyuData() As Daikyu_Type
    Private M_StaffID As String
    Private m_Index As Integer
    Private Const M_YYYYMMDDLabel_Row As Integer = 2 '日付隠しセル行
    Private Const M_PASTE As String = "1"
    Private Const M_DELETE As String = "2"
    Private Const M_SET As String = "3"
    Private m_DaikyuBackColorFlg As Boolean '代休勤務ﾊﾞｯｸｶﾗｰﾌﾗｸﾞ

    Private m_toolTipTxt As String 'ツールチップ表示文字列
    Private m_empRowDispFlg As Boolean '採用列表示フラグ
    Private Structure NightShortInfo
        Dim Date_St As Integer
        Dim Date_Ed As Integer
    End Structure
    Private m_nightWorkInfo() As NightShortInfo '夜勤専従
    Private m_shortWorkInfo() As NightShortInfo '短時間

    'セット勤務用
    Private Structure SetKinmu_Type
        Dim Mark As String
        <VBFixedArray(10)> Dim CD() As String
        Dim StrText As String
        Dim KinmuCnt As Integer
        Dim blnKinmu As Boolean

        Public Sub Initialize()
            ReDim CD(10)
        End Sub
    End Structure

    Private m_SetKinmu() As SetKinmu_Type
    Private m_SetCnt As Short
    Private m_lngDaikyuPastPeriod As Integer '過去の代休取得時の休日出勤日の有効範囲（何日前までの休日出勤は有効って感じ）
    Private m_FontSize As Short
    'ﾌｫﾝﾄｻｲｽﾞ定数
    Private Const M_FontSize_Big As Short = 14 '「大」ﾌｫﾝﾄｻｲｽﾞ=14
    Private Const M_FontSize_Middle As Short = 12 '「中」ﾌｫﾝﾄｻｲｽﾞ=12
    Private Const M_FontSize_Small As Short = 9 '「小」ﾌｫﾝﾄｻｲｽﾞ=9
    '2014/04/23 Saijo add start P-06979-----------------------------------------
    Private Const M_FontSize_Second_Big As Short = 10 '「大」ﾌｫﾝﾄｻｲｽﾞ=10
    Private Const M_FontSize_Second_Middle As Short = 9 '「中」ﾌｫﾝﾄｻｲｽﾞ=9
    Private Const M_FontSize_Second_Small As Short = 7 '「小」ﾌｫﾝﾄｻｲｽﾞ=7
    '2014/04/23 Saijo add end P-06979-------------------------------------------
    Private m_SpreadSize As Double
    '検索用文字列
    Private m_HolDateStr As String '祝日 検索用文字列
    Private m_OffDayStr As String '休日 検索用文字列
    Private m_Daikyu15KinmuCD() As String '代休が1.5日発生する勤務ＣＤ
    Private m_PackageFLG As Short 'パッケージマスタ(0:届出×日当直×,1:届出×日当直○,2:届出○日当直×,3:届出○日当直○)
    Private m_StartDate As Integer
    Private m_EndDate As Integer
    Private m_strUpdKojyoDate As String

    Private m_lstCmdKinmu As New List(Of Object)
    Private m_lstCmdYasumi As New List(Of Object)
    Private m_lstCmdTokusyu As New List(Of Object)
    Private m_lstCmdSet As New List(Of Object)

    '開始日取得
    Public WriteOnly Property pStartDate() As Integer
        Set(ByVal Value As Integer)
            m_StartDate = Value
        End Set
    End Property

    '終了日取得
    Public WriteOnly Property pEndDate() As Integer
        Set(ByVal Value As Integer)
            m_EndDate = Value
        End Set
    End Property

    Public WriteOnly Property pDataHideFlg() As Object
        Set(ByVal Value As Object)
            m_DataHideFlg = Value
            If IsArray(m_DataHideFlg) = False Then
                ReDim m_DataHideFlg(0)
            End If
        End Set
    End Property

    Public WriteOnly Property pKakuteiHideFlg() As Object
        Set(ByVal Value As Object)
            m_KakuteiHideFlg = Value
            If IsArray(m_KakuteiHideFlg) = False Then
                ReDim m_KakuteiHideFlg(0)
            End If
        End Set
    End Property

    Public WriteOnly Property pDataFlg() As Object
        Set(ByVal Value As Object)
            m_DataFlg = Value
            If IsArray(m_DataFlg) = False Then
                ReDim m_DataFlg(0)
            End If
        End Set
    End Property

    Public WriteOnly Property pKakuteiFlg() As Object
        Set(ByVal Value As Object)
            m_KakuteiFlg = Value
            If IsArray(m_KakuteiFlg) = False Then
                ReDim m_KakuteiFlg(0)
            End If
        End Set
    End Property

    '画面立ち上げﾓｰﾄﾞの取得(1:４週／１ヶ月  2:４週／１ヶ月以外)
    Public WriteOnly Property pDispKikan() As String
        Set(ByVal Value As String)
            m_DispKikan = Value
        End Set
    End Property

    'ﾌｫﾝﾄｻｲｽﾞ(ﾌｫｰﾑの幅を設定するときに使用)
    Public WriteOnly Property pFontSize() As Short
        Set(ByVal Value As Short)
            m_FontSize = Value
        End Set
    End Property

    'ﾌｫｰﾑの幅
    Public WriteOnly Property pSpreadSize() As Double
        Set(ByVal Value As Double)
            m_SpreadSize = Value
        End Set
    End Property

    '計画/実績画面ﾓｰﾄﾞの取得ﾓｰﾄﾞ(1:計画  2:実績)
    Public WriteOnly Property pMode() As String
        Set(ByVal Value As String)
            m_Mode = Value
        End Set
    End Property

    '一段/二段表示ﾓｰﾄﾞの取得ﾓｰﾄﾞ(True:二段  False:一段)
    Public WriteOnly Property pPlanTwo() As Boolean
        Set(ByVal Value As Boolean)
            m_blnOneTwo = Value
        End Set
    End Property

    '職員ﾃﾞｰﾀ開始行
    Public WriteOnly Property pStaffStartRow() As Integer
        Set(ByVal Value As Integer)
            m_StaffStartRow = Value
        End Set
    End Property

    '最大表示段数
    Public WriteOnly Property pMaxShowLine(ByVal p_DutyData As Short, ByVal p_KinmuPlan As Short, ByVal p_KinmuJisseki As Short, ByVal p_AppliData As Short) As Short
        Set(ByVal Value As Short)
            m_DutyData = p_DutyData
            m_KinmuPlan = p_KinmuPlan
            m_KinmuJisseki = p_KinmuJisseki
            m_AppliData = p_AppliData
            m_MaxShowLine = Value
        End Set
    End Property

    'ツールチップテキスト
    Public WriteOnly Property pToolTxt() As String
        Set(ByVal value As String)
            m_toolTipTxt = value
        End Set
    End Property

    '採用列の表示フラグ
    Public WriteOnly Property pEmpRowVisible() As Boolean
        Set(ByVal value As Boolean)
            m_empRowDispFlg = value
        End Set
    End Property

    '夜勤専従・育児短時間初期化
    Public WriteOnly Property pInitNightShortInfo(ByVal p_ngt As Integer, ByVal p_shr As Integer) As Boolean
        Set(ByVal value As Boolean)
            ReDim m_nightWorkInfo(p_ngt)
            ReDim m_shortWorkInfo(p_ngt)
        End Set
    End Property

    '夜勤専従情報
    Public WriteOnly Property pNightWork(ByVal p_st As Integer, ByVal p_ed As Integer) As Boolean
        Set(ByVal value As Boolean)
            Dim idx As Integer = UBound(m_nightWorkInfo)
            m_nightWorkInfo(idx).Date_St = p_st
            m_nightWorkInfo(idx).Date_Ed = p_ed

            If m_StartDate > p_st Then m_nightWorkInfo(idx).Date_St = m_StartDate
            If m_EndDate < p_ed Then m_nightWorkInfo(idx).Date_Ed = m_EndDate
        End Set
    End Property

    '短時間者情報
    Public WriteOnly Property pShortWork(ByVal p_st As Integer, ByVal p_ed As Integer) As Boolean
        Set(ByVal value As Boolean)
            Dim idx As Integer = UBound(m_shortWorkInfo)
            m_shortWorkInfo(idx).Date_St = p_st
            m_shortWorkInfo(idx).Date_Ed = p_ed

            If m_StartDate > p_st Then m_shortWorkInfo(idx).Date_St = m_StartDate
            If m_EndDate < p_ed Then m_shortWorkInfo(idx).Date_Ed = m_EndDate
        End Set
    End Property

    '代休データ受け取り
    Public WriteOnly Property pDaikyuData(ByVal p_HolDate As Integer, ByVal p_HolKinmuCD As String, ByVal p_GetKbn As String, ByVal p_RemainderHol As Double, ByVal p_OutPutList As String, ByVal p_DaikyuDate As Integer, ByVal p_DaikyuKinmuCD As String) As String
        Set(ByVal Value As String)
            Dim w_SeachLoop As Integer
            Dim w_TargetIdx As Integer
            Dim w_SubTargetIdx As Integer

            If m_Index = 0 Then
                '配列確保
                ReDim m_DaikyuData(0)
            End If

            '既に取得済みの日付かチェック
            w_TargetIdx = 0
            For w_SeachLoop = 1 To UBound(m_DaikyuData)
                If m_DaikyuData(w_SeachLoop).HolDate = p_HolDate Then
                    '取得済みの場合ｲﾝﾃﾞｯｸｽをセット
                    w_TargetIdx = w_SeachLoop
                End If
            Next w_SeachLoop

            If w_TargetIdx = 0 Then
                '新規の場合、配列拡張
                'ｲﾝﾃﾞｯｸｽｶｳﾝﾄｱｯﾌﾟ
                m_Index = m_Index + 1
                w_TargetIdx = m_Index
                ReDim Preserve m_DaikyuData(w_TargetIdx)
                m_DaikyuData(w_TargetIdx).HolDate = p_HolDate
                m_DaikyuData(w_TargetIdx).HolKinmuCD = p_HolKinmuCD
                m_DaikyuData(w_TargetIdx).GetKbn = p_GetKbn
                m_DaikyuData(w_TargetIdx).RemainderHol = p_RemainderHol
                m_DaikyuData(w_TargetIdx).OutPutList = p_OutPutList

                w_SubTargetIdx = 1
            Else
                w_SubTargetIdx = UBound(m_DaikyuData(w_TargetIdx).DaikyuDetail) + 1
            End If

            ReDim Preserve m_DaikyuData(w_TargetIdx).DaikyuDetail(w_SubTargetIdx)
            m_DaikyuData(w_TargetIdx).DaikyuDetail(w_SubTargetIdx).DaikyuDate = p_DaikyuDate
            m_DaikyuData(w_TargetIdx).DaikyuDetail(w_SubTargetIdx).DaikyuKinmuCD = p_DaikyuKinmuCD
            m_DaikyuData(w_TargetIdx).DaikyuDetail(w_SubTargetIdx).GetFlg = Value
        End Set
    End Property

    '職員管理番号を受け取る
    Public WriteOnly Property pStaffID() As String
        Set(ByVal Value As String)
            M_StaffID = Value

            '代休用配列初期化
            ReDim m_DaikyuData(0)

            '代休配列ｲﾝﾃﾞｯｸｽ初期化
            m_Index = 0
        End Set
    End Property

    '更新ﾌﾗｸﾞ(True:更新,False:未更新)
    Public ReadOnly Property pKosinFlg() As Boolean
        Get
            pKosinFlg = m_KosinFlg
        End Get
    End Property

    '代休データ親画面引渡し
    Public Function pDaikyuDataGet(ByRef p_HolKinmuCD As String, ByRef p_OutPutList As String, ByRef p_GetKbn As String, ByRef p_RemainderHol As Double, ByRef p_DetailDataCnt As Integer, ByVal p_Int As Integer) As Integer

        pDaikyuDataGet = m_DaikyuData(p_Int).HolDate
        p_HolKinmuCD = m_DaikyuData(p_Int).HolKinmuCD


        p_OutPutList = m_DaikyuData(p_Int).OutPutList
        p_GetKbn = m_DaikyuData(p_Int).GetKbn
        p_RemainderHol = m_DaikyuData(p_Int).RemainderHol
        p_DetailDataCnt = UBound(m_DaikyuData(p_Int).DaikyuDetail)
    End Function

    Public Function pDaikyuDetailDataGet(ByRef p_DaikyuKinmuCD As String, ByRef p_GetFlg As String, ByVal p_Int As Integer, ByVal p_SubInt As Integer) As Integer
        pDaikyuDetailDataGet = m_DaikyuData(p_Int).DaikyuDetail(p_SubInt).DaikyuDate
        p_DaikyuKinmuCD = m_DaikyuData(p_Int).DaikyuDetail(p_SubInt).DaikyuKinmuCD
        p_GetFlg = m_DaikyuData(p_Int).DaikyuDetail(p_SubInt).GetFlg
    End Function

    '代休データ件数
    Public ReadOnly Property pDaikyuCnt() As Integer
        Get
            pDaikyuCnt = UBound(m_DaikyuData)
        End Get
    End Property

    '休日情報を取得
    Public WriteOnly Property pHolData(ByVal p_HolDateStr As String) As String
        Set(ByVal Value As String)
            m_HolDateStr = p_HolDateStr
            m_OffDayStr = Value
        End Set
    End Property

    '計画ﾃﾞｰﾀ 開始/終了 列番号情報を受け取る
    Public WriteOnly Property pKeikakuD_EndCol() As Integer
        Set(ByVal Value As Integer)
            m_KeikakuD_EndCol = Value
        End Set
    End Property

    '計画ﾃﾞｰﾀ 開始/終了 列番号情報を受け取る
    Public WriteOnly Property pKeikakuD_StartCol() As Integer
        Set(ByVal Value As Integer)
            m_KeikakuD_StartCol = Value
        End Set
    End Property

    '親計画ﾃﾞｰﾀ 開始/終了 列番号情報を受け取る
    Public WriteOnly Property pKeikakuD_KinmuDataEndCol() As Integer
        Set(ByVal Value As Integer)
            m_KeikakuD_EndCol_Param = Value
        End Set
    End Property

    '親計画ﾃﾞｰﾀ 開始/終了 列番号情報を受け取る
    Public WriteOnly Property pKeikakuD_KinmuDataStartCol() As Integer
        Set(ByVal Value As Integer)
            m_KeikakuD_StartCol_Param = Value
        End Set
    End Property

    '親ｶﾚﾝﾄ行 位置をを受け取る
    Public WriteOnly Property pCUR_ROW() As Integer
        Set(ByVal Value As Integer)
            m_CUR_ROW_Param = Value
        End Set
    End Property

    '親画面で表示されているｽﾌﾟﾚｯﾄﾞｼｰﾄを受け取る
    Public WriteOnly Property pControl() As FarPoint.Win.Spread.FpSpread
        Set(ByVal Value As FarPoint.Win.Spread.FpSpread)
            m_Control_Param = Value
        End Set
    End Property

    Public ReadOnly Property pUpdKojyoDate() As String
        Get
            pUpdKojyoDate = m_strUpdKojyoDate
        End Get
    End Property

    '代休チェック
    Private Function Check_Daikyu(ByVal p_Ivent As String, ByVal p_Date As Object, ByVal p_KinmuCD As String) As Boolean

        Const W_SUBNAME As String = "NSK0000HC Check_Daikyu"

        Dim w_frmDaikyu As Object
        Dim w_Int As Integer
        Dim w_HolDate As Integer
        Dim w_HolKinmuCD As String
        Dim w_SelDate As Integer
        Dim w_STS As Integer
        Dim w_strMsg() As String
        Dim w_indIdx As Short
        Dim w_lngLoop As Integer
        Dim w_lngLoop2 As Integer
        Dim w_SelDate2 As Integer
        Dim w_lngIdx As Integer
        Dim w_lngDataCnt As Integer
        Dim w_HalfDaikyuFlg As Boolean

        '初期値
        Check_Daikyu = False
        Try
            m_DaikyuBackColorFlg = False

            'ペーストの場合
            If p_Ivent = "1" Then
                '祝日であるかどうか
                If General.pafncDaikyuCheck(General.g_strHospitalCD, p_Date, General.g_strSelKinmuDeptCD) = True Or (m_SundayDaikyuFlg = 1 And Weekday(CDate(Format(Integer.Parse(p_Date), "0000/00/00"))) = 1) Or (m_SaturdayDaikyuFlg = 1 And Weekday(CDate(Format(Integer.Parse(p_Date), "0000/00/00"))) = 7) Then
                    If p_KinmuCD <> "" Then

                        'すでにある配列の休日勤務日と選択された日付で同じ日があるか
                        For w_Int = 1 To UBound(m_DaikyuData)
                            'もしすでにある配列の中にあったら、代休取得済みであるかチェック
                            If m_DaikyuData(w_Int).HolDate = p_Date Then
                                For w_lngLoop = 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                    If m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate <> 0 Then
                                        ReDim w_strMsg(1)
                                        w_strMsg(1) = "代休取得済みの勤務です。~n"
                                        Call General.paMsgDsp("NS0110", w_strMsg)
                                        Exit Function
                                    End If
                                Next w_lngLoop
                            End If
                        Next w_Int

                        Select Case g_KinmuM(CShort(p_KinmuCD)).DaikyuFlg
                            Case "1"
                                '祝日に勤務分類=勤務の勤務CDを貼り付けた時は代休データを作成する
                                If m_DaikyuMsgFlg = 0 Then
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = ""
                                    w_STS = General.paMsgDsp("NS0111", w_strMsg)
                                Else
                                    '非表示の場合、必ず取得
                                    w_STS = MsgBoxResult.Yes
                                End If

                                If w_STS = MsgBoxResult.Yes Then
                                    '代休情報更新
                                    Call UpDate_DaikyuData(p_Date, p_KinmuCD, 0)

                                    m_DaikyuBackColorFlg = True
                                Else
                                    '代休情報更新
                                    Call UpDate_DaikyuData(p_Date, p_KinmuCD, 1)
                                End If
                            Case "2"
                                If g_KinmuM(CShort(p_KinmuCD)).HolBunruiCD = General.G_STRDAIKYUBUNRUI Then
                                    '祝日に代休を貼り付けた場合、代休は取得できない
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "この日に代休を"
                                    Call General.paMsgDsp("NS0112", w_strMsg)
                                    Exit Function
                                Else
                                    '代休情報更新
                                    Call UpDate_DaikyuData(p_Date, p_KinmuCD, 1)
                                End If
                        End Select
                    End If
                Else
                    '祝日でない場合
                    If g_KinmuM(CShort(p_KinmuCD)).HolBunruiCD = General.G_STRDAIKYUBUNRUI Then

                        w_frmDaikyu = New frmNSK0000HJ

                        w_frmDaikyu.pSelDate = Integer.Parse(p_Date)
                        w_frmDaikyu.pSelKinmuCD = p_KinmuCD
                        '起動元（"0":計画 それ以外:初期画面）
                        w_frmDaikyu.pKeikakuFlg = "0"
                        '代休データを渡す
                        For w_Int = 1 To UBound(m_DaikyuData)
                            'リスト対象データであるか
                            If m_DaikyuData(w_Int).OutPutList = "1" Then
                                If m_DaikyuAdvFlg = 0 Then
                                    '代休有効期限があるか？
                                    If m_lngDaikyuPastPeriod = -1 Then
                                        '代休有効期限なし
                                        '対象日付より過去のデータか？
                                        If m_DaikyuData(w_Int).HolDate < Integer.Parse(p_Date) Then
                                            w_HolDate = m_DaikyuData(w_Int).HolDate
                                            w_HolKinmuCD = m_DaikyuData(w_Int).HolKinmuCD
                                            '情報引渡し
                                            w_frmDaikyu.pDaikyuData(w_HolDate, w_HolKinmuCD) = m_DaikyuData(w_Int).RemainderHol
                                        End If
                                    Else
                                        '代休有効期限あり
                                        '対象日付より過去のデータかつ指定日(ﾃﾞﾌｫﾙﾄ56日)前までであるか
                                        If (m_DaikyuData(w_Int).HolDate < Integer.Parse(p_Date)) And (DateDiff(DateInterval.Day, CDate(Format(m_DaikyuData(w_Int).HolDate, "0000/00/00")), CDate(Format(Integer.Parse(p_Date), "0000/00/00"))) <= m_lngDaikyuPastPeriod) Then
                                            '対象日付より過去のデータであるか
                                            w_HolDate = m_DaikyuData(w_Int).HolDate
                                            w_HolKinmuCD = m_DaikyuData(w_Int).HolKinmuCD
                                            '情報引渡し
                                            w_frmDaikyu.pDaikyuData(w_HolDate, w_HolKinmuCD) = m_DaikyuData(w_Int).RemainderHol
                                        End If
                                    End If
                                Else
                                    '指定日前後であるかどうか
                                    '代休有効期限があるか？
                                    If m_lngDaikyuPastPeriod = -1 Then
                                        If m_DaikyuAdvThisMonthFlg = 0 Then
                                            '代休有効期限なし
                                            w_HolDate = m_DaikyuData(w_Int).HolDate
                                            w_HolKinmuCD = m_DaikyuData(w_Int).HolKinmuCD
                                            '情報引渡し
                                            w_frmDaikyu.pDaikyuData(w_HolDate, w_HolKinmuCD) = m_DaikyuData(w_Int).RemainderHol
                                        Else
                                            If m_DaikyuData(w_Int).HolDate <= CDbl(Format(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(Integer.Parse(General.paLeft(p_Date, 6) & "01"), "0000/00/00")))), "yyyyMMdd")) Then
                                                w_HolDate = m_DaikyuData(w_Int).HolDate
                                                w_HolKinmuCD = m_DaikyuData(w_Int).HolKinmuCD
                                                '情報引渡し
                                                '代休の未使用分も渡す
                                                w_frmDaikyu.pDaikyuData(w_HolDate, w_HolKinmuCD) = m_DaikyuData(w_Int).RemainderHol
                                            End If
                                        End If
                                    Else
                                        If m_DaikyuAdvThisMonthFlg = 0 Then
                                            '代休有効期限あり
                                            '代休有効期限範囲内か？
                                            If (DateDiff(DateInterval.Day, CDate(Format(m_DaikyuData(w_Int).HolDate, "0000/00/00")), CDate(Format(Integer.Parse(p_Date), "0000/00/00"))) <= m_lngDaikyuPastPeriod) And (DateDiff(DateInterval.Day, CDate(Format(Integer.Parse(p_Date), "0000/00/00")), CDate(Format(m_DaikyuData(w_Int).HolDate, "0000/00/00"))) <= m_lngDaikyuPastPeriod) Then
                                                w_HolDate = m_DaikyuData(w_Int).HolDate
                                                w_HolKinmuCD = m_DaikyuData(w_Int).HolKinmuCD
                                                '情報引渡し
                                                w_frmDaikyu.pDaikyuData(w_HolDate, w_HolKinmuCD) = m_DaikyuData(w_Int).RemainderHol
                                            End If
                                        Else
                                            If DateDiff(DateInterval.Day, CDate(Format(m_DaikyuData(w_Int).HolDate, "0000/00/00")), CDate(Format(Integer.Parse(p_Date), "0000/00/00"))) <= m_lngDaikyuPastPeriod And m_DaikyuData(w_Int).HolDate <= CDbl(Format(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(Integer.Parse(General.paLeft(p_Date, 6) & "01"), "0000/00/00")))), "yyyyMMdd")) And DateDiff(DateInterval.Day, CDate(Format(Integer.Parse(p_Date), "0000/00/00")), CDate(Format(m_DaikyuData(w_Int).HolDate, "0000/00/00"))) <= m_lngDaikyuPastPeriod Then
                                                w_HolDate = m_DaikyuData(w_Int).HolDate
                                                w_HolKinmuCD = m_DaikyuData(w_Int).HolKinmuCD
                                                '情報引渡し
                                                '代休の未使用分も渡す
                                                w_frmDaikyu.pDaikyuData(w_HolDate, w_HolKinmuCD) = m_DaikyuData(w_Int).RemainderHol
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next w_Int

                        If w_frmDaikyu.mfncDaikyuDate_Check = False Then
                            ReDim w_strMsg(1)
                            w_strMsg(1) = "この日に代休を"
                            Call General.paMsgDsp("NS0112", w_strMsg)
                            Exit Function
                        Else
                            '祝日以外に代休を貼り付けた場合、代休管理Ｆを取得し、取得できる代休の一覧画面を出力する。
                            w_frmDaikyu.ShowDialog(Me)
                            If w_frmDaikyu.pEndStatus = False Then
                                Exit Function
                            Else
                                'OKﾎﾞﾀﾝ押下時
                                w_SelDate = w_frmDaikyu.pSelDate

                                '半日 半日代休取得時の2つ目の取得年月日を取得
                                w_SelDate2 = w_frmDaikyu.pSelDate2
                                '半日代休Dlg
                                w_HalfDaikyuFlg = w_frmDaikyu.pGetDaikyuType

                                '代休の日にさらに代休を貼り付けた場合元の代休データをクリア
                                For w_Int = 1 To UBound(m_DaikyuData)
                                    w_lngDataCnt = UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                    For w_lngLoop = 1 To w_lngDataCnt
                                        If w_lngLoop > UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                            Exit For
                                        End If

                                        If m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = Integer.Parse(p_Date) Then
                                            'クリアするので代休未使用分に加算
                                            '半日代休かチェック
                                            Select Case m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg
                                                Case "0"
                                                    '１日
                                                    m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 1
                                                Case "1"
                                                    '半日
                                                    m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 0.5
                                            End Select

                                            m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = 0
                                            m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuKinmuCD = ""
                                            m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg = ""
                                            m_DaikyuData(w_Int).OutPutList = "1"

                                            If w_lngLoop < UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                                w_indIdx = w_lngLoop
                                                For w_lngLoop2 = w_lngLoop + 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                                    m_DaikyuData(w_Int).DaikyuDetail(w_indIdx) = m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop2)
                                                    w_indIdx = w_indIdx + 1
                                                Next w_lngLoop2

                                                '配列調整
                                                ReDim Preserve m_DaikyuData(w_Int).DaikyuDetail(UBound(m_DaikyuData(w_Int).DaikyuDetail) - 1)
                                            End If
                                        End If
                                    Next w_lngLoop
                                Next w_Int

                                '選択された日付のデータに代休CDと日付を格納
                                For w_Int = 1 To UBound(m_DaikyuData)
                                    If m_DaikyuData(w_Int).OutPutList = "1" Then
                                        If m_DaikyuData(w_Int).HolDate = w_SelDate Or m_DaikyuData(w_Int).HolDate = w_SelDate2 Then
                                            w_lngIdx = 1
                                            '１件目にデータがはいっているかチェック
                                            If m_DaikyuData(w_Int).DaikyuDetail(w_lngIdx).DaikyuDate <> 0 Then
                                                'はいっている場合、配列拡張
                                                w_lngIdx = UBound(m_DaikyuData(w_Int).DaikyuDetail) + 1
                                                ReDim Preserve m_DaikyuData(w_Int).DaikyuDetail(w_lngIdx)
                                            End If

                                            m_DaikyuData(w_Int).DaikyuDetail(w_lngIdx).DaikyuDate = Integer.Parse(p_Date)
                                            m_DaikyuData(w_Int).DaikyuDetail(w_lngIdx).DaikyuKinmuCD = p_KinmuCD

                                            '代休未使用分から差し引く
                                            If w_HalfDaikyuFlg = True Or w_SelDate2 <> 0 Then
                                                '半日代休の場合
                                                m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol - 0.5
                                                m_DaikyuData(w_Int).DaikyuDetail(w_lngIdx).GetFlg = "1"
                                            Else
                                                '１日代休の場合
                                                m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol - 1
                                                m_DaikyuData(w_Int).DaikyuDetail(w_lngIdx).GetFlg = "0"
                                            End If

                                            '代休の未使用分がなくなった場合
                                            If m_DaikyuData(w_Int).RemainderHol <= 0 Then
                                                m_DaikyuData(w_Int).OutPutList = "0"
                                            End If
                                        End If
                                    End If
                                Next w_Int
                            End If
                        End If
                    Else
                        '祝日以外に代休以外の勤務を貼り付けた場合、その日の元勤務が代休の場合,代休管理Ｆの代休取得年月日をNULLで更新する
                        '代休を取得しているかチェックし、取得していたら配列から代休データを削除
                        For w_Int = 1 To UBound(m_DaikyuData)
                            For w_lngLoop = 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                If w_lngLoop > UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                    Exit For
                                End If

                                If m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = Integer.Parse(p_Date) Then
                                    'クリアするので代休未使用分に加算
                                    '半日代休かチェック
                                    Select Case m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg
                                        Case "0"
                                            '１日
                                            m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 1
                                        Case "1"
                                            '半日
                                            m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 0.5
                                    End Select

                                    m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = 0
                                    m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuKinmuCD = ""
                                    m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg = ""
                                    m_DaikyuData(w_Int).OutPutList = "1"

                                    If w_lngLoop < UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                        w_indIdx = w_lngLoop
                                        For w_lngLoop2 = w_lngLoop + 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                            m_DaikyuData(w_Int).DaikyuDetail(w_indIdx) = m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop2)
                                            w_indIdx = w_indIdx + 1
                                        Next w_lngLoop2
                                        '配列調整
                                        ReDim Preserve m_DaikyuData(w_Int).DaikyuDetail(UBound(m_DaikyuData(w_Int).DaikyuDetail) - 1)
                                    End If
                                End If
                            Next w_lngLoop
                        Next w_Int
                    End If
                End If
            ElseIf p_Ivent = "2" Then
                '削除の場合
                '祝日であるかどうか
                If General.pafncDaikyuCheck(General.g_strHospitalCD, p_Date, General.g_strSelKinmuDeptCD) = True Then
                    '代休を取得している勤務でないかチェック
                    For w_Int = 1 To UBound(m_DaikyuData)
                        If m_DaikyuData(w_Int).HolDate = Integer.Parse(p_Date) Then
                            If m_DaikyuData(w_Int).OutPutList = "0" Then
                                ReDim w_strMsg(1)
                                w_strMsg(1) = "代休を取得しているので"
                                Call General.paMsgDsp("NS0098", w_strMsg)

                                Exit Function
                            End If
                        End If
                    Next w_Int

                    '代休情報更新
                    Call UpDate_DaikyuData(p_Date, p_KinmuCD, 1)
                Else
                    '祝日でない場合
                    '代休を削除しているかチェックし、削除する場合は配列から代休データを削除
                    For w_Int = 1 To UBound(m_DaikyuData)
                        For w_lngLoop = 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                            If w_lngLoop > UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                Exit For
                            End If

                            If m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = Integer.Parse(p_Date) Then
                                'クリアするので代休未使用分に加算
                                '半日代休かチェック
                                Select Case m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg
                                    Case "0"
                                        '１日
                                        m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 1
                                    Case "1"
                                        '半日
                                        m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 0.5
                                End Select

                                m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = 0
                                m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuKinmuCD = ""
                                m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg = ""
                                m_DaikyuData(w_Int).OutPutList = "1"

                                If w_lngLoop < UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                    w_indIdx = w_lngLoop
                                    For w_lngLoop2 = w_lngLoop + 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                        m_DaikyuData(w_Int).DaikyuDetail(w_indIdx) = m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop2)
                                        w_indIdx = w_indIdx + 1
                                    Next w_lngLoop2

                                    '配列調整
                                    ReDim Preserve m_DaikyuData(w_Int).DaikyuDetail(UBound(m_DaikyuData(w_Int).DaikyuDetail) - 1)
                                End If
                            End If
                        Next w_lngLoop
                    Next w_Int
                End If
            ElseIf p_Ivent = "3" Then
                'セット勤務貼り付け時
                '代休情報更新
                If General.pafncDaikyuCheck(General.g_strHospitalCD, p_Date, General.g_strSelKinmuDeptCD) = True Then
                    Call UpDate_DaikyuData(p_Date, p_KinmuCD, 1)
                Else
                    For w_Int = 1 To UBound(m_DaikyuData)
                        For w_lngLoop = 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                            If w_lngLoop > UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                Exit For
                            End If

                            If m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = Integer.Parse(p_Date) Then
                                'クリアするので代休未使用分に加算
                                '半日代休かチェック
                                Select Case m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg
                                    Case "0"
                                        '１日
                                        m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 1
                                    Case "1"
                                        '半日
                                        m_DaikyuData(w_Int).RemainderHol = m_DaikyuData(w_Int).RemainderHol + 0.5
                                End Select

                                m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuDate = 0
                                m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).DaikyuKinmuCD = ""
                                m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop).GetFlg = ""
                                m_DaikyuData(w_Int).OutPutList = "1"

                                If w_lngLoop < UBound(m_DaikyuData(w_Int).DaikyuDetail) Then
                                    w_indIdx = w_lngLoop
                                    For w_lngLoop2 = w_lngLoop + 1 To UBound(m_DaikyuData(w_Int).DaikyuDetail)
                                        m_DaikyuData(w_Int).DaikyuDetail(w_indIdx) = m_DaikyuData(w_Int).DaikyuDetail(w_lngLoop2)
                                        w_indIdx = w_indIdx + 1
                                    Next w_lngLoop2

                                    '配列調整
                                    ReDim Preserve m_DaikyuData(w_Int).DaikyuDetail(UBound(m_DaikyuData(w_Int).DaikyuDetail) - 1)
                                End If
                            End If
                        Next w_lngLoop
                    Next w_Int
                End If
            End If

            Check_Daikyu = True

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    '代休情報更新
    Private Sub UpDate_DaikyuData(ByVal p_Date As Integer, ByVal p_KinmuCD As String, ByVal p_Mode As Integer)

        Const W_SUBNAME As String = "NSK0000HC UpDate_DaikyuData"

        Dim w_Int As Integer
        Dim w_Int2 As Integer
        Dim w_DataFlg As Boolean
        Dim w_Index As Integer
        Dim w_WorkTbl As Daikyu_Type
        Dim w_WorkTbl2() As Daikyu_Type
        Try
            If p_Mode = CDbl("0") Then

                w_DataFlg = False

                '代休データ件数ループ
                For w_Int = 1 To UBound(m_DaikyuData)
                    '休日出勤年月日が今回貼り付けた日付と一致するか
                    If m_DaikyuData(w_Int).HolDate = p_Date Then

                        w_DataFlg = True

                        '勤務が変更になっていないか
                        If m_DaikyuData(w_Int).HolKinmuCD = p_KinmuCD Then
                            '代休取得日にしているか
                            If m_DaikyuData(w_Int).OutPutList = "0" Then
                                m_DaikyuData(w_Int).OutPutList = "1"
                                Exit For
                            End If
                        Else
                            m_DaikyuData(w_Int).HolKinmuCD = p_KinmuCD
                            m_DaikyuData(w_Int).OutPutList = "1"

                            '取得区分
                            'とりあえず１日としてセット
                            m_DaikyuData(w_Int).GetKbn = "0"
                            m_DaikyuData(w_Int).RemainderHol = 1

                            '代休1.5日分発生対象勤務かチェック
                            For w_Int2 = 1 To UBound(m_Daikyu15KinmuCD)
                                If p_KinmuCD = m_Daikyu15KinmuCD(w_Int2) Then
                                    m_DaikyuData(w_Int).GetKbn = "1"
                                    m_DaikyuData(w_Int).RemainderHol = 1.5
                                    Exit For
                                End If
                            Next w_Int2

                            Exit For
                        End If
                    End If
                Next w_Int

                '一致するデータがない場合は配列に追加
                If w_DataFlg = False Then
                    '配列拡張
                    w_Index = UBound(m_DaikyuData) + 1
                    ReDim Preserve m_DaikyuData(w_Index)
                    ReDim m_DaikyuData(w_Index).DaikyuDetail(1)

                    m_DaikyuData(w_Index).HolDate = p_Date
                    m_DaikyuData(w_Index).HolKinmuCD = p_KinmuCD
                    m_DaikyuData(w_Index).OutPutList = "1"

                    '取得区分
                    'とりあえず１日としてセット
                    m_DaikyuData(w_Index).GetKbn = "0"
                    m_DaikyuData(w_Index).RemainderHol = 1
                    '1.5日対象勤務かチェック
                    For w_Int = 1 To UBound(m_Daikyu15KinmuCD)
                        If p_KinmuCD = m_Daikyu15KinmuCD(w_Int) Then
                            m_DaikyuData(w_Index).GetKbn = "1"
                            m_DaikyuData(w_Index).RemainderHol = 1.5
                            Exit For
                        End If
                    Next w_Int

                    '休日出勤日でソート
                    For w_Int = 1 To UBound(m_DaikyuData)
                        For w_Int2 = 1 To UBound(m_DaikyuData) - w_Int
                            If (m_DaikyuData(w_Int2).HolDate > m_DaikyuData(w_Int2 + 1).HolDate) Then
                                w_WorkTbl = m_DaikyuData(w_Int2)
                                m_DaikyuData(w_Int2) = m_DaikyuData(w_Int2 + 1)
                                m_DaikyuData(w_Int2 + 1) = w_WorkTbl
                            End If
                        Next w_Int2
                    Next w_Int
                End If
            ElseIf p_Mode = CDbl("1") Then
                '代休データ件数ループ
                For w_Int = 1 To UBound(m_DaikyuData)
                    '休日出勤年月日が今回貼り付けた日付と一致するか
                    If m_DaikyuData(w_Int).HolDate = p_Date Then

                        'ワークテーブルにデータを退避
                        ReDim w_WorkTbl2(UBound(m_DaikyuData))
                        For w_Int2 = 1 To UBound(m_DaikyuData)
                            w_WorkTbl2(w_Int2) = m_DaikyuData(w_Int2)
                        Next w_Int2

                        '代休用配列初期化
                        ReDim m_DaikyuData(0)

                        'いいえを選択しているので配列から削除
                        For w_Int2 = 1 To UBound(w_WorkTbl2)
                            If w_Int <> w_Int2 Then
                                w_Index = UBound(m_DaikyuData) + 1
                                ReDim Preserve m_DaikyuData(w_Index)
                                m_DaikyuData(w_Index) = w_WorkTbl2(w_Int2)
                            End If
                        Next w_Int2

                        Exit For
                    End If
                Next w_Int
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Public Sub ShowWindow()

        Const W_SUBNAME As String = "NSK0000HC  ShowWindow"

        Dim w_Size As Double
        '2018/09/21 K.I Add Start---------
        Dim w_Left As String
        Dim w_Top As String
        '2018/09/21 K.I Add Start---------

        Try
            'ﾌｫｰﾑのｻｲｽﾞを調整
            '2014/04/23 Saijo upd start P-06979----------------------------
            'w_Size = m_SpreadSize + General.paTwipsTopixels(500)
            'If w_Size < General.paTwipsTopixels(9750) Then
            '    w_Size = General.paTwipsTopixels(9750)
            'Else
            '    'フォント小のときはサイズ固定
            '    If m_FontSize = M_FontSize_Small Then
            '       w_Size = General.paTwipsTopixels(12000)
            '    End If
            'End If
            If m_strKinmuEmSecondFlg = "0" Then
                w_Size = m_SpreadSize + General.paTwipsTopixels(500)
                If w_Size < General.paTwipsTopixels(9750) Then
                    w_Size = General.paTwipsTopixels(9750)
                Else
                    'フォント小のときはサイズ固定
                    If m_FontSize = M_FontSize_Small Then
                        w_Size = General.paTwipsTopixels(12000)
                    End If
                End If
            Else
                If m_FontSize = M_FontSize_Second_Big Then
                    w_Size = m_SpreadSize + General.paTwipsTopixels(700)
                Else
                    w_Size = 960.0 + General.paTwipsTopixels(2700)
                End If
            End If

            '2014/04/23 Saijo upd end P-06979------------------------------

            Me.Width = w_Size

            Call subSetCtlList()

            '2014/04/23 Saijo upd start P-06979----------------------------
            '勤務記号全角２文字対応のレイアウト変更
            Call SetKinmuSecondView()
            '2014/04/23 Saijo upd end P-06979------------------------------

            '-----------ﾊﾟﾚｯﾄ 設定-----------
            'NSKINMUNAMEM 取得
            Call GetKinmuName()

            'パネルウィンドウに記号をセット
            Call SetKinmuData()
            '--------------------------------

            '消しゴムICON設定
            cmdErase.Image = Image.FromFile(g_ImagePath & G_ERASER_ICO)

            '画面ｾﾝﾀﾘﾝｸﾞ
            Me.StartPosition = FormStartPosition.CenterScreen

            'ｳｨﾝﾄﾞｳの表示ﾎﾟｼﾞｼｮﾝを設定する
            '2018/09/21 K.I Upd Start------------------------------------------------------------------------------------------
            'レジストリ取得を削除
            'Call General.paGetWindowPositon(Me, General.G_STRMAINKEY2 & "\" & g_AppName)
            '画面中央
            w_Left = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - CSng(Me.Width)) / 2)
            w_Top = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - CSng(Me.Height)) / 2)
            Me.SetBounds((CSng(w_Left)), (CSng(w_Top)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
            '2018/09/21 K.I Upd End--------------------------------------------------------------------------------------------
            'ｳｨﾝﾄﾞｳの表示ﾎﾟｼﾞｼｮﾝを設定する

            'ｶｰｿﾙｾﾙ移動
            Call SetStartCursol()

            Me.ShowDialog(pProcessObj)

        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub GetKinmuName()
        Const W_SUBNAME As String = "NSK0000HC  GetKinmuName"

        Dim w_Int As Short
        Dim w_Int2 As Short
        Dim w_DataCnt As Short
        Dim w_Sql As String
        'DAOｵﾌﾞｼﾞｪｸﾄ
        Dim w_Rs As ADODB.Recordset
        'ﾌｨｰﾙﾄﾞ
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
        Dim w_KinmuCnt As Integer
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
        Dim w_strKinmuBunruiCD As String
        Try
            '初期化
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
                    w_DataCnt = .fKN_KinmuCount

                    For w_Int = 1 To w_DataCnt

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
                                        m_KinmuCnt = m_KinmuCnt + 1
                                        'NS_KINMUNAME_M格納用変数の再定義
                                        ReDim Preserve m_Kinmu(m_KinmuCnt)

                                        m_Kinmu(m_KinmuCnt - 1).CD = .fKN_KinmuCD
                                        m_Kinmu(m_KinmuCnt - 1).KinmuName = .fKN_Name
                                        m_Kinmu(m_KinmuCnt - 1).Mark = .fKN_MarkF
                                        m_Kinmu(m_KinmuCnt - 1).KBunruiCD = w_strKinmuBunruiCD
                                        m_Kinmu(m_KinmuCnt - 1).Setumei = .fKN_KinmuExplan
                                    End If
                                Else
                                    m_KinmuCnt = m_KinmuCnt + 1
                                    'NS_KINMUNAME_M格納用変数の再定義
                                    ReDim Preserve m_Kinmu(m_KinmuCnt)

                                    m_Kinmu(m_KinmuCnt - 1).CD = .fKN_KinmuCD
                                    m_Kinmu(m_KinmuCnt - 1).KinmuName = .fKN_Name
                                    m_Kinmu(m_KinmuCnt - 1).Mark = .fKN_MarkF
                                    m_Kinmu(m_KinmuCnt - 1).KBunruiCD = w_strKinmuBunruiCD
                                    m_Kinmu(m_KinmuCnt - 1).Setumei = .fKN_KinmuExplan
                                End If
                                '2015/04/14 Bando Upd End   ============================

                            ElseIf w_strKinmuBunruiCD = "2" Then
                                '-- 休み --
                                '2015/04/14 Bando Upd Start ============================
                                '希望モードの場合、表示対象勤務のみパレットに表示
                                'If g_HopeMode = 1 Then
                                If g_HopeMode = 1 AndAlso m_DispKinmuCd <> "" Then '2015/06/02 Bando Chg
                                    If InStr(m_DispKinmuCd, .fKN_KinmuCD) > 0 Then
                                        m_YasumiCnt = m_YasumiCnt + 1
                                        ReDim Preserve m_Yasumi(m_YasumiCnt)

                                        m_Yasumi(m_YasumiCnt - 1).CD = .fKN_KinmuCD
                                        m_Yasumi(m_YasumiCnt - 1).KinmuName = .fKN_Name
                                        m_Yasumi(m_YasumiCnt - 1).Mark = .fKN_MarkF
                                        m_Yasumi(m_YasumiCnt - 1).KBunruiCD = w_strKinmuBunruiCD
                                        m_Yasumi(m_YasumiCnt - 1).Setumei = .fKN_KinmuExplan
                                    End If
                                Else
                                    m_YasumiCnt = m_YasumiCnt + 1
                                    ReDim Preserve m_Yasumi(m_YasumiCnt)

                                    m_Yasumi(m_YasumiCnt - 1).CD = .fKN_KinmuCD
                                    m_Yasumi(m_YasumiCnt - 1).KinmuName = .fKN_Name
                                    m_Yasumi(m_YasumiCnt - 1).Mark = .fKN_MarkF
                                    m_Yasumi(m_YasumiCnt - 1).KBunruiCD = w_strKinmuBunruiCD
                                    m_Yasumi(m_YasumiCnt - 1).Setumei = .fKN_KinmuExplan
                                End If
                                '2015/04/14 Bando Upd End   ============================

                            ElseIf w_strKinmuBunruiCD = "3" Then
                                '-- 特殊 --
                                '2015/04/14 Bando Upd Start ============================
                                '希望モードの場合、表示対象勤務のみパレットに表示
                                'If g_HopeMode = 1 Then
                                If g_HopeMode = 1 AndAlso m_DispKinmuCd <> "" Then '2015/06/02 Bando Chg
                                    If InStr(m_DispKinmuCd, .fKN_KinmuCD) > 0 Then
                                        m_TokusyuCnt = m_TokusyuCnt + 1
                                        ReDim Preserve m_Tokusyu(m_TokusyuCnt)

                                        m_Tokusyu(m_TokusyuCnt - 1).CD = .fKN_KinmuCD
                                        m_Tokusyu(m_TokusyuCnt - 1).KinmuName = .fKN_Name
                                        m_Tokusyu(m_TokusyuCnt - 1).Mark = .fKN_MarkF
                                        m_Tokusyu(m_TokusyuCnt - 1).KBunruiCD = w_strKinmuBunruiCD
                                        m_Tokusyu(m_TokusyuCnt - 1).Setumei = .fKN_KinmuExplan
                                    End If
                                Else
                                    m_TokusyuCnt = m_TokusyuCnt + 1
                                    ReDim Preserve m_Tokusyu(m_TokusyuCnt)

                                    m_Tokusyu(m_TokusyuCnt - 1).CD = .fKN_KinmuCD
                                    m_Tokusyu(m_TokusyuCnt - 1).KinmuName = .fKN_Name
                                    m_Tokusyu(m_TokusyuCnt - 1).Mark = .fKN_MarkF
                                    m_Tokusyu(m_TokusyuCnt - 1).KBunruiCD = w_strKinmuBunruiCD
                                    m_Tokusyu(m_TokusyuCnt - 1).Setumei = .fKN_KinmuExplan
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

            '2015/06/02 Bando Upd Start =====================================
            If g_HopeMode <> 1 Then
                'セット勤務
                '2017/05/02 Christopher Upd Start
                ''SQL文編集
                'w_Sql = "SELECT * FROM NS_SETKINMU_M "
                'w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                'w_Sql = w_Sql & "AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                'w_Sql = w_Sql & "ORDER BY DISPNO "

                'w_Rs = General.paDBRecordSetOpen(w_Sql)
                '<1>
                Call NSK0000H_sql.select_NS_SETKINMU_M_01(w_Rs)
                'Upd End
                'セット勤務配列初期化
                ReDim m_SetKinmu(0)

                If w_Rs.RecordCount <= 0 Then
                    m_SetCnt = 0
                Else
                    w_Int3 = 0

                    With w_Rs
                        .MoveLast()
                        w_DataCnt = .RecordCount
                        .MoveFirst()

                        ReDim m_SetKinmu(w_DataCnt)
                        m_SetCnt = w_DataCnt

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

                        For w_Int = 1 To w_DataCnt

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
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD2
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD3
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD4
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD5
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD6
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD7
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD8
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD9
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                    Case w_strKinmuCD10
                                        w_blnEndDate = False
                                        m_SetCnt = m_SetCnt - 1
                                        Exit For
                                End Select
                            Next w_Int2

                            If w_blnEndDate = True Then
                                m_SetKinmu(w_Int3).Initialize()
                                m_SetKinmu(w_Int3).Mark = w_記号_F.Value & ""
                                m_SetKinmu(w_Int3).CD(1) = w_勤務CD1_F.Value & ""
                                m_SetKinmu(w_Int3).CD(2) = w_勤務CD2_F.Value & ""
                                m_SetKinmu(w_Int3).CD(3) = w_勤務CD3_F.Value & ""
                                m_SetKinmu(w_Int3).CD(4) = w_勤務CD4_F.Value & ""
                                m_SetKinmu(w_Int3).CD(5) = w_勤務CD5_F.Value & ""
                                m_SetKinmu(w_Int3).CD(6) = w_勤務CD6_F.Value & ""
                                m_SetKinmu(w_Int3).CD(7) = w_勤務CD7_F.Value & ""
                                m_SetKinmu(w_Int3).CD(8) = w_勤務CD8_F.Value & ""
                                m_SetKinmu(w_Int3).CD(9) = w_勤務CD9_F.Value & ""
                                m_SetKinmu(w_Int3).CD(10) = w_勤務CD10_F.Value & ""
                                m_SetKinmu(w_Int3).blnKinmu = True

                                '勤務がいくつあるか(間に空白はないものとする)
                                w_KinmuCnt = 0

                                For w_Int2 = 1 To 10
                                    If m_SetKinmu(w_Int3).CD(w_Int2) <> "" Then
                                        w_KinmuCnt = w_KinmuCnt + 1
                                    Else
                                        Exit For
                                    End If
                                Next w_Int2

                                m_SetKinmu(w_Int3).KinmuCnt = w_KinmuCnt
                                w_Int3 = w_Int3 + 1
                            End If
                            .MoveNext()
                        Next w_Int
                    End With
                End If
                '2015/06/02 Bando Upd End   =====================================

                w_Rs.Close()
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub SetKinmuData()

        Const W_SUBNAME As String = "NSK0000HC  SetKinmuData"

        Dim w_i As Short
        Try
            'コマンドボタンのＣＡＰＴＩＯＮ設定
            '勤務
            For w_i = 1 To M_PARET_NUM
                If w_i <= m_KinmuCnt Then
                    m_lstCmdKinmu(w_i - 1).Text = m_Kinmu(w_i - 1).Mark
                    If m_Kinmu(w_i - 1).Setumei = "" Then
                        ToolTip1.SetToolTip(m_lstCmdKinmu(w_i - 1), Get_KinmuTipText(m_Kinmu(w_i - 1).CD))
                    Else
                        ToolTip1.SetToolTip(m_lstCmdKinmu(w_i - 1), Get_KinmuTipText(m_Kinmu(w_i - 1).CD) & "：" & m_Kinmu(w_i - 1).Setumei)
                    End If
                Else
                    Exit For
                End If
            Next w_i

            '休み
            For w_i = 1 To M_PARET_NUM
                If w_i <= m_YasumiCnt Then
                    m_lstCmdYasumi(w_i - 1).Text = m_Yasumi(w_i - 1).Mark
                    If m_Yasumi(w_i - 1).Setumei = "" Then
                        ToolTip1.SetToolTip(m_lstCmdYasumi(w_i - 1), Get_KinmuTipText(m_Yasumi(w_i - 1).CD))
                    Else
                        ToolTip1.SetToolTip(m_lstCmdYasumi(w_i - 1), Get_KinmuTipText(m_Yasumi(w_i - 1).CD) & "：" & m_Yasumi(w_i - 1).Setumei)
                    End If
                Else
                    Exit For
                End If
            Next w_i

            '特殊勤務
            For w_i = 1 To M_PARET_NUM
                If w_i <= m_TokusyuCnt Then
                    m_lstCmdTokusyu(w_i - 1).Text = m_Tokusyu(w_i - 1).Mark
                    If m_Tokusyu(w_i - 1).Setumei = "" Then
                        ToolTip1.SetToolTip(m_lstCmdTokusyu(w_i - 1), Get_KinmuTipText(m_Tokusyu(w_i - 1).CD))
                    Else
                        ToolTip1.SetToolTip(m_lstCmdTokusyu(w_i - 1), Get_KinmuTipText(m_Tokusyu(w_i - 1).CD) & "：" & m_Tokusyu(w_i - 1).Setumei)
                    End If
                Else
                    Exit For
                End If
            Next w_i

            'セット勤務
            For w_i = 1 To M_PARET_NUM_SET
                If w_i <= m_SetCnt Then
                    If m_SetKinmu(w_i - 1).blnKinmu = True Then
                        m_lstCmdSet(w_i - 1).Text = m_SetKinmu(w_i - 1).Mark
                        ToolTip1.SetToolTip(m_lstCmdSet(w_i - 1), Get_SetKinmuTipText(w_i - 1))
                    End If
                Else
                    Exit For
                End If
            Next w_i

            'スクロールバー、オプションボタンの設定
            '勤務
            Select Case m_KinmuCnt
                Case 0 To M_PARET_NUM
                    For w_i = M_PARET_NUM To (m_KinmuCnt + 1) Step -1
                        m_lstCmdKinmu(w_i - 1).Visible = False
                    Next w_i

                    HscKinmu.Maximum = (0 + HscKinmu.LargeChange - 1)
                    HscKinmu.Visible = False
                Case Else
                    For w_i = 1 To M_PARET_NUM
                        m_lstCmdKinmu(w_i - 1).Visible = True
                        m_lstCmdKinmu(w_i - 1).Enabled = True
                    Next w_i

                    HscKinmu.Maximum = (General.paRoundUp((m_KinmuCnt - M_PARET_NUM) / 2, 0) + HscKinmu.LargeChange - 1)
                    HscKinmu.Visible = True
                    HscKinmu.Enabled = True
            End Select

            '休み
            Select Case m_YasumiCnt
                Case 0 To M_PARET_NUM
                    For w_i = M_PARET_NUM To (m_YasumiCnt + 1) Step -1
                        m_lstCmdYasumi(w_i - 1).Visible = False
                    Next w_i

                    HscYasumi.Maximum = (0 + HscYasumi.LargeChange - 1)
                    HscYasumi.Visible = False
                Case Else
                    For w_i = 1 To M_PARET_NUM
                        m_lstCmdYasumi(w_i - 1).Visible = True
                        m_lstCmdYasumi(w_i - 1).Enabled = True
                    Next w_i

                    HscYasumi.Maximum = (General.paRoundUp((m_YasumiCnt - M_PARET_NUM) / 2, 0) + HscYasumi.LargeChange - 1)
                    HscYasumi.Visible = True
                    HscYasumi.Enabled = True
            End Select

            '特殊勤務
            Select Case m_TokusyuCnt
                Case 0 To M_PARET_NUM
                    For w_i = M_PARET_NUM To (m_TokusyuCnt + 1) Step -1
                        m_lstCmdTokusyu(w_i - 1).Visible = False
                    Next w_i

                    HscTokusyu.Maximum = (0 + HscTokusyu.LargeChange - 1)
                    HscTokusyu.Visible = False
                Case Else
                    For w_i = 1 To M_PARET_NUM
                        m_lstCmdTokusyu(w_i - 1).Visible = True
                        m_lstCmdTokusyu(w_i - 1).Enabled = True
                    Next w_i

                    HscTokusyu.Maximum = (General.paRoundUp((m_TokusyuCnt - M_PARET_NUM) / 2, 0) + HscTokusyu.LargeChange - 1)
                    HscTokusyu.Visible = True
                    HscTokusyu.Enabled = True
            End Select

            'セット勤務
            Select Case m_SetCnt
                Case 0 To M_PARET_NUM_SET
                    For w_i = M_PARET_NUM_SET To (m_SetCnt + 1) Step -1
                        m_lstCmdSet(w_i - 1).Visible = False
                    Next w_i

                    HscSet.Maximum = (0 + HscSet.LargeChange - 1)
                    HscSet.Visible = False
                Case Else
                    For w_i = 1 To M_PARET_NUM_SET
                        m_lstCmdSet(w_i - 1).Visible = True
                        m_lstCmdSet(w_i - 1).Enabled = True
                    Next w_i

                    HscSet.Maximum = (m_SetCnt - M_PARET_NUM_SET + HscSet.LargeChange - 1)
                    HscSet.Visible = True
                    HscSet.Enabled = True
            End Select

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    'セット勤務ツールチップ用文字列取得
    Public Function Get_SetKinmuTipText(ByVal p_Int As Integer) As String

        Const W_SUBNAME As String = "NSK0000HC Get_SetKinmuTipText"

        Dim w_str As String = String.Empty
        Dim w_strTEXT As String = String.Empty
        Dim w_Cnt As Integer
        Dim w_CD As String
        Try
            For w_Cnt = 1 To 10
                '勤務CDを取得
                w_CD = m_SetKinmu(p_Int).CD(w_Cnt)

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

            m_SetKinmu(p_Int).StrText = w_strTEXT

            Get_SetKinmuTipText = w_str

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    Private Function Disp_Ouen(ByRef p_CD As String) As Boolean

        Const W_SUBNAME As String = "NSK0000HC  Disp_Ouen"

        Dim w_Form As frmNSK0000HK
        Dim w_strMsg() As String

        '初期化
        Disp_Ouen = False
        Try
            w_Form = New frmNSK0000HK
            w_Form.frmNSK0000HK_Load()
            If w_Form.pKangoCDFlg = True Then
                '表示
                w_Form.ShowDialog(Me)

                'OK押下ﾎﾞﾀﾝ時のみ処理続行
                If w_Form.pOKFlg = True Then
                    p_CD = w_Form.pSelKangoTCD
                    Disp_Ouen = True
                End If
            Else
                '勤務部署マスタ取得失敗
                ReDim w_strMsg(1)
                w_strMsg(1) = "勤務部署情報"
                Call General.paMsgDsp("NS0031", w_strMsg)
            End If

            '解放
            w_Form.Dispose()

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    '2015/04/13 Bando Add Start =======================================================
    Private Function Disp_Comment(ByRef p_com As String, ByRef p_riyu As String) As Boolean

        Const W_SUBNAME As String = "NSK0000HC  Disp_Comment"

        Dim w_Form As Object

        '初期化
        Disp_Comment = False
        Try
            w_Form = New frmNSK0000HP
            w_Form.pRiyuKbn = p_riyu
            'コメントが既に存在する場合はプロパティで受け渡す
            If p_com <> "" Then
                w_Form.p_com = p_com
            End If

            w_Form.frmNSK0000HP_Load()

            '表示
            w_Form.ShowDialog(Me)

            'OK押下ﾎﾞﾀﾝ時のみ処理続行
            If w_Form.pOKFlg = True Then
                p_com = w_Form.pComment
                Disp_Comment = True
            End If

            '解放
            w_Form = Nothing

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function
    '2015/04/13 Bando Add End   =======================================================

    Private Sub chkSet_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSet.CheckStateChanged

        Const W_SUBNAME As String = "NSK0000HC  chkSet_Click"
        Try
            If chkSet.CheckState = 1 Then
                '実績変更ではない
                If m_Mode <> General.G_PGMSTARTFLG_CHANGEJISSEKI Then
                    If g_LimitedFlg = False Then
                        If g_SaikeiFlg = True Then
                            _OptRiyu_0.Enabled = False
                            _OptRiyu_1.Enabled = False
                            _OptRiyu_2.Enabled = False
                            _OptRiyu_3.Enabled = True
                            _OptRiyu_3.Checked = True
                            _OptRiyu_4.Enabled = False
                            _Frame_3.Enabled = True
                            _Frame_0.Enabled = False
                            _Frame_1.Enabled = False
                            _Frame_2.Enabled = False
                        Else
                            picFrame.Enabled = False
                            _OptRiyu_0.Checked = True
                            _OptRiyu_0.Enabled = True
                            _OptRiyu_1.Enabled = False
                            '2014/05/14 Shimpo upd start P-06991-----------------------------------------------------------------------
                            '_OptRiyu_2.Enabled = False
                            If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                                _OptRiyu_2.Enabled = False
                            Else
                                _OptRiyu_2.Enabled = True
                            End If
                            '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                            _OptRiyu_4.Enabled = False
                            _Frame_3.Enabled = True
                            _Frame_0.Enabled = False
                            _Frame_1.Enabled = False
                            _Frame_2.Enabled = False
                        End If
                    Else
                        _Frame_3.Enabled = True
                        _Frame_0.Enabled = False
                        _Frame_1.Enabled = False
                        _Frame_2.Enabled = False
                    End If
                End If
            Else
                '実績変更ではない
                If m_Mode <> General.G_PGMSTARTFLG_CHANGEJISSEKI Then
                    If g_LimitedFlg = False Then
                        If g_SaikeiFlg = True Then
                            picFrame.Enabled = True
                            _OptRiyu_0.Enabled = False
                            _OptRiyu_1.Enabled = False
                            _OptRiyu_2.Enabled = False
                            _OptRiyu_3.Enabled = True
                            _OptRiyu_3.Checked = True
                            _OptRiyu_4.Enabled = False
                            _Frame_3.Enabled = False
                            _Frame_0.Enabled = True
                            _Frame_1.Enabled = True
                            _Frame_2.Enabled = True
                        Else
                            picFrame.Enabled = True
                            _OptRiyu_0.Enabled = True
                            _OptRiyu_1.Enabled = True
                            _OptRiyu_2.Enabled = True
                            _OptRiyu_4.Enabled = True
                            _Frame_3.Enabled = False
                            _Frame_0.Enabled = True
                            _Frame_1.Enabled = True
                            _Frame_2.Enabled = True
                        End If
                    Else
                        _Frame_3.Enabled = False
                        _Frame_0.Enabled = True
                        _Frame_1.Enabled = True
                        _Frame_2.Enabled = True
                    End If
                End If
            End If
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '閉じるﾎﾞﾀﾝ押下
    Private Sub cmdEnd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEnd.Click

        Const W_SUBNAME As String = "NSK0000HC  cmdEnd_Click"
        Try
            '画面消去
            Me.Close()

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '消しゴムﾎﾞﾀﾝ押下
    Private Sub cmdErase_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdErase.Click

        Const W_SUBNAME As String = "NSK0000HC  cmdErase_Click"

        Dim w_Lng As Integer
        Dim w_i As Integer
        Dim w_Cnt As Short
        Dim w_Row As Integer
        Dim w_YYYYMMDD As Integer
        Dim w_Var As Object
        Dim w_strMsg() As String
        Dim w_strUpdKojyoDate As String = String.Empty
        Dim w_CellRange() As Model.CellRange
        Dim w_KinmuCD As String   'KinmuCD
        Dim w_RiyuKBN As String   '理由区分
        Dim w_Time As String  '時間年休
        Dim w_Flg As String   '確定ﾌﾗｸﾞ
        Dim w_KinmuPlanCD As String
        Dim w_KangoCD As String
        Dim w_STS As Integer
        Dim w_ActiveCol As Long
        Dim w_ActiveRow As Long
        Dim w_MsgFlg1 As Boolean '希望勤務
        Dim w_MsgFlg2 As Boolean '再掲勤務
        Dim w_MsgFlg3 As Boolean '委員会勤務
        Dim w_MsgFlg4 As Boolean '応援勤務
        Dim w_MsgFlg5 As Boolean '要請勤務

        Try
            With sprSheet.Sheets(0)

                w_MsgFlg1 = False
                w_MsgFlg2 = False
                w_MsgFlg3 = False
                w_MsgFlg4 = False
                w_MsgFlg5 = False

                '消去位置ﾁｪｯｸ(列)
                If .ActiveColumn.Index < m_KeikakuD_StartCol Or .ActiveColumn.Index2 > m_KeikakuD_EndCol Then
                    Exit Sub
                End If

                '消去位置ﾁｪｯｸ(行)
                If .ActiveRow.Index < M_KinmuData_Row Or .ActiveRow.Index2 < M_KinmuData_Row Then
                    Exit Sub
                End If

                '連続したｾﾙﾌﾟﾛｯｸか？
                w_CellRange = .GetSelections

                For w_i = 0 To w_CellRange.Length - 1
                    '消去位置ﾁｪｯｸ(列)
                    If w_CellRange(w_i).Column < m_KeikakuD_StartCol Or (w_CellRange(w_i).Column + w_CellRange(w_i).ColumnCount - 1) > m_KeikakuD_EndCol Then
                        Exit Sub
                    End If

                    '消去位置ﾁｪｯｸ(行)
                    If w_CellRange(w_i).Row <> M_KinmuData_Row Or w_CellRange(w_i).RowCount <> 1 Then
                        Exit Sub
                    End If
                Next w_i

                If w_CellRange.Length = 0 Then
                    ReDim w_CellRange(0)
                    w_CellRange(0) = New Model.CellRange(.ActiveRow.Index, .ActiveColumn.Index, 1, 1)
                End If

                For w_i = 0 To w_CellRange.Length - 1
                    w_Row = w_CellRange(w_i).Row

                    For w_Lng = w_CellRange(w_i).Column To w_CellRange(w_i).Column + w_CellRange(w_i).ColumnCount - 1
                        If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_Row, w_Lng).BackColor) = m_MonthBefore_Back Then
                            '配属期間外(背景色がグレー)の場合入力不可
                            Exit Sub
                        End If

                        w_Cnt = CShort(w_Lng - M_KinmuData_Col + 1)

                        If UBound(m_DataFlg) >= w_Cnt Then
                            If g_SaikeiFlg = True Then
                                '再掲部署の場合
                                If m_DataFlg(w_Cnt) = "1" Then
                                    '実績ﾃﾞｰﾀの場合は消去不可
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "確定済み勤務を"
                                    Call General.paMsgDsp("NS0098", w_strMsg)
                                    Exit Sub
                                End If
                            Else
                                '再掲部署以外の場合
                                If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN And m_DataFlg(w_Cnt) = "1" And m_KakuteiFlg(w_Cnt) = "0" Then
                                    '計画変更で該当部署確定ﾃﾞｰﾀの場合は消去不可
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "確定済み勤務を"
                                    Call General.paMsgDsp("NS0098", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If

                        w_Var = .GetText(w_Row, w_CellRange(w_i).Column)
                        '2015/04/10 Bando Upd Start =======================
                        'Call Get_KinmuMark(w_Var, w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        Call Get_KinmuMark(w_Var, w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, "")
                        '2015/04/10 Bando Upd End   =======================
                        '警告表示
                        If g_SaikeiFlg = False Then
                            Select Case w_RiyuKBN
                                Case "2"
                                    If w_MsgFlg5 = False Then
                                        If frmNSK0000HA._mnuTool_5.Checked = True Then
                                            ReDim w_strMsg(2)
                                            w_strMsg(1) = "要請勤務"
                                            w_strMsg(2) = "削除"
                                            w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                            If w_STS = MsgBoxResult.No Then
                                                Exit Sub
                                            End If
                                            w_MsgFlg5 = True
                                        End If
                                    End If
                                Case "3"
                                    If w_MsgFlg1 = False Then
                                        If frmNSK0000HA._mnuTool_4.Checked = True Then
                                            ReDim w_strMsg(2)
                                            w_strMsg(1) = "希望勤務"
                                            w_strMsg(2) = "削除"
                                            w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                            If w_STS = MsgBoxResult.No Then
                                                Exit Sub
                                            End If
                                            w_MsgFlg1 = True
                                        End If
                                    End If
                                Case "4"
                                    If w_MsgFlg2 = False Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "再掲勤務"
                                        w_strMsg(2) = "削除"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                        w_MsgFlg2 = True
                                    End If
                                Case "5"
                                    If w_MsgFlg3 = False Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "委員会勤務"
                                        w_strMsg(2) = "削除"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                        w_MsgFlg3 = True
                                    End If
                                Case "6"
                                    If w_MsgFlg4 = False Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "応援勤務"
                                        w_strMsg(2) = "削除"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                        w_MsgFlg4 = True
                                    End If
                            End Select
                        End If

                        If General.g_lngDaikyuMng = 0 Then
                            '代休ﾁｪｯｸ
                            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_Lng)
                            w_YYYYMMDD = w_Var
                            If Check_Daikyu(M_DELETE, w_YYYYMMDD, "") = False Then
                                Exit Sub
                            End If
                        End If

                        w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_Lng)
                        If InStr(m_strUpdKojyoDate, w_Var) = 0 Then
                            w_strUpdKojyoDate = w_strUpdKojyoDate & w_Var & ","
                        End If
                    Next w_Lng

                    sprSheet.Sheets(0).ClearRange(w_CellRange(w_i).Row, w_CellRange(w_i).Column, w_CellRange(w_i).RowCount, w_CellRange(w_i).ColumnCount, True)

                    .Cells(w_CellRange(w_i).Row, w_CellRange(w_i).Column, w_CellRange(w_i).Row + w_CellRange(w_i).RowCount - 1, w_CellRange(w_i).Column + w_CellRange(w_i).ColumnCount - 1).ForeColor = Color.Black
                    .Cells(w_CellRange(w_i).Row, w_CellRange(w_i).Column, w_CellRange(w_i).Row + w_CellRange(w_i).RowCount - 1, w_CellRange(w_i).Column + w_CellRange(w_i).ColumnCount - 1).BackColor = Color.White
                Next w_i
            End With

            m_strUpdKojyoDate = m_strUpdKojyoDate & w_strUpdKojyoDate

            '更新ﾌﾗｸﾞ設定
            m_KosinFlg = True

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    'Private Sub m_lstCmdKinmu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _cmdKinmu_0.Click, _cmdKinmu_2.Click, _cmdKinmu_4.Click, _
    '                                                                                                            _cmdKinmu_6.Click, _cmdKinmu_8.Click, _cmdKinmu_10.Click, _cmdKinmu_12.Click, _
    '                                                                                                            _cmdKinmu_14.Click, _cmdKinmu_16.Click, _cmdKinmu_18.Click, _cmdKinmu_20.Click, _
    '                                                                                                            _cmdKinmu_22.Click, _cmdKinmu_24.Click, _cmdKinmu_26.Click, _cmdKinmu_28.Click, _
    '                                                                                                            _cmdKinmu_30.Click, _cmdKinmu_34.Click, _cmdKinmu_36.Click, _cmdKinmu_38.Click
    Private Sub m_lstCmdKinmu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim Index As Short = m_lstCmdKinmu.IndexOf(eventSender)
        Const W_SUBNAME As String = "NSK0000HC  m_lstCmdKinmu_Click"

        Dim w_RegStr As String
        Dim w_Var As Object
        Dim w_ActiveCol As Integer
        Dim w_ActiveRow As Integer
        Dim w_KinmuCD As String 'KinmuCD
        Dim w_RiyuKBN As String '理由区分
        Dim w_Time As String '時間年休
        Dim w_Flg As String '確定ﾌﾗｸﾞ
        Dim w_ForeColor As Integer '文字色
        Dim w_BackColor As Integer '背景色
        Dim w_InputFlg As Boolean '入力ﾌﾗｸﾞ
        Dim w_Cnt As Short
        Dim w_KinmuPlanCD As String = String.Empty 'KinmuCD(予定ﾃﾞｰﾀ)
        Dim w_STS As Short
        Dim w_KangoCD As String
        Dim w_RiyuPlanKbn As String
        Dim w_KangoPlanCD As String
        Dim w_YYYYMMDD As Integer
        Dim w_DaikyuInputFlg As Boolean
        Dim w_strMsg() As String
        Dim w_lngBackColor As Integer
        Dim w_IntCol As Short
        Dim w_KibouCnt As Short
        Dim w_KibouCol() As Integer
        Dim w_blnColChk As Boolean
        Dim w_Style As New StyleInfo
        '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
        Dim w_KibouCntDate As Integer
        '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------
        Dim w_Comment As String = String.Empty  '希望勤務時のコメント 2015/04/13 Band Add
        Try
            'ﾌｫｰｶｽ移動
            sprSheet.Focus()

            'ﾚｼﾞｽﾄﾘ格納先
            w_RegStr = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            '初期化
            w_KangoCD = ""

            '１件でも存在すれば・・・
            If m_KinmuCnt <> 0 Then

                w_ActiveCol = sprSheet.Sheets(0).ActiveColumn.Index
                w_ActiveRow = sprSheet.Sheets(0).ActiveRow.Index

                '入力場所ﾁｪｯｸ
                If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row_ChgJisseki Then
                        Exit Sub
                    End If
                Else
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row Then
                        Exit Sub
                    End If
                End If

                '部署異動ﾁｪｯｸ（配属範囲）
                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_ActiveCol).BackColor) = m_MonthBefore_Back Then
                    '配属期間外(背景色がグレー)の場合入力不可
                    Exit Sub
                End If

                '入力ﾌﾗｸﾞ
                w_InputFlg = True
                w_Cnt = CShort(w_ActiveCol - M_KinmuData_Col + 1)
                If g_SaikeiFlg = True Then
                    '再掲部署の場合
                    If m_DataFlg(w_Cnt) = "1" Then
                        '実績ﾃﾞｰﾀの場合
                        ReDim w_strMsg(2)
                        w_strMsg(1) = "確定済み勤務"
                        w_strMsg(2) = "再掲勤務"
                        Call General.paMsgDsp("NS0011", w_strMsg)

                        w_InputFlg = False
                    End If
                End If


                '勤務の曜日制限チェック
                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                '2013/10/02 Bando Chg Start ====================================================
                'If Check_YoubiLimit(w_YYYYMMDD, m_Kinmu(Index + HscKinmu.Value).CD) = False Then
                If Check_YoubiLimit(w_YYYYMMDD, m_Kinmu(Index + HscKinmu.Value * 2).CD) = False Then
                    '2013/10/02 Bando Chg End ====================================================
                    Exit Sub
                End If

                '2013/01/07 Ishiga add start------------------------------------------
                '超勤データの有無チェック
                If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If Check_OverKinmuData(w_YYYYMMDD) = False Then
                        Exit Sub
                    End If
                End If
                '2013/01/07 Ishiga add end--------------------------------------------

                '届出存在チェック
                If fncChkAppliData(w_YYYYMMDD) = False Then
                    Exit Sub
                End If

                If m_PackageFLG = 1 Or m_PackageFLG = 3 Then
                    '日当直のコンポーネントありの時
                    With General.g_objGetData
                        .p職員区分 = 0
                        .p職員番号 = M_StaffID
                        .pチェック基準日 = w_YYYYMMDD
                        .p処理区分 = 0
                        '2013/10/02 Bando Chg Start ====================================================
                        '.pチェック勤務CD = m_Kinmu(Index + HscKinmu.Value).CD
                        .pチェック勤務CD = m_Kinmu(Index + HscKinmu.Value * 2).CD
                        '2013/10/02 Bando Chg Start ====================================================

                        If .mChkKinmuDuty = False Then
                            '勤務変更不可
                            '*******ﾒｯｾｰｼﾞ***********************************
                            ReDim w_strMsg(1)
                            w_strMsg(1) = ""
                            Call General.paMsgDsp("NS0110", w_strMsg)
                            '************************************************
                            Exit Sub
                        End If
                    End With
                End If


                w_DaikyuInputFlg = True

                If General.g_lngDaikyuMng = 0 Then
                    '代休ﾁｪｯｸ
                    w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                    w_YYYYMMDD = w_Var
                    '2013/10/02 Bando Chg Start ======================================================================
                    'If Check_Daikyu(M_PASTE, w_YYYYMMDD, m_Kinmu(Index + HscKinmu.Value).CD) = False Then
                    If Check_Daikyu(M_PASTE, w_YYYYMMDD, m_Kinmu(Index + HscKinmu.Value * 2).CD) = False Then
                        '2013/10/02 Bando Chg End ======================================================================
                        w_DaikyuInputFlg = False
                        w_InputFlg = False
                    End If
                End If

                If w_InputFlg = True Then
                    '入力可能な場合
                    With sprSheet.Sheets(0)

                        '警告表示+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        '変更する勤務の値が空で計画変更の場合、一つ上の勤務を取得する。
                        w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                        If w_Var = "" And m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            w_Var = .GetText(w_ActiveRow - 1, w_ActiveCol)
                        End If

                        '2015/04/13 Bando Upd Start ========================
                        'Call Get_KinmuMark(w_Var, w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        Call Get_KinmuMark(w_Var, w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, "")
                        '2015/04/13 Bando Upd End   ========================

                        If g_SaikeiFlg = False Then
                            Select Case w_RiyuKBN
                                Case "2"
                                    If frmNSK0000HA._mnuTool_5.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "要請勤務"
                                        w_strMsg(2) = "編集"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "3"
                                    If frmNSK0000HA._mnuTool_4.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "希望勤務"
                                        w_strMsg(2) = "編集"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "4"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "再掲勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "5"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "委員会勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "6"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "応援勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                            End Select
                        Else
                            If w_RiyuKBN = "2" Then
                                If frmNSK0000HA._mnuTool_5.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "要請勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If

                            If w_RiyuKBN = "3" Then
                                If frmNSK0000HA._mnuTool_4.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "希望勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If

                        w_KinmuCD = m_Kinmu(Index + HscKinmu.Value * 2).CD
                        w_Time = ""
                        w_Flg = "0"
                        w_KangoCD = ""

                        '希望回数集計チェック
                        w_lngBackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                        If CDbl(g_HopeNumFlg) = 1 And _OptRiyu_2.Checked = True Then
                            w_KibouCnt = 0
                            ReDim w_KibouCol(0)
                            For w_IntCol = m_KeikakuD_StartCol To m_KeikakuD_EndCol

                                If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_IntCol).BackColor) = w_lngBackColor Then
                                    w_KibouCnt = w_KibouCnt + 1
                                    ReDim Preserve w_KibouCol(w_KibouCnt)
                                    w_KibouCol(w_KibouCnt) = w_IntCol
                                End If
                            Next w_IntCol

                            '同じ場所に同じ勤務を貼り付けた場合、スルー
                            w_blnColChk = False
                            For w_IntCol = 1 To UBound(w_KibouCol)
                                If w_ActiveCol = w_KibouCol(w_IntCol) And w_KinmuCD = w_KinmuPlanCD Then
                                    w_blnColChk = True
                                End If
                            Next w_IntCol

                            If g_HopeNum <= w_KibouCnt And w_blnColChk = False Then
                                '希望回数制限オーバーダイアログ表示
                                If g_KibouNumDiaLogFlg = 1 Then
                                    'ワーニング
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "希望勤務"
                                    w_strMsg(2) = "設定された希望勤務回数"
                                    '「&1が&2を超えています。~nこのまま登録してもよろしいですか。」
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    'エラー
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "設定された希望勤務回数を超えているため"
                                    '「&1入力できません。」
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If

                        '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
                        If CDbl(g_HopeNumDateFlg) = 1 And _OptRiyu_2.Checked = True Then
                            '日付別の希望勤務数チェック
                            w_KibouCntDate = frmNSK0000HA.Get_HopeNum_Of_Date(w_YYYYMMDD)
                            If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor) = w_lngBackColor Then
                                w_KibouCntDate = w_KibouCntDate - 1
                            End If

                            If g_HopeNumDate <= w_KibouCntDate Then
                                '希望回数制限オーバーダイアログ表示
                                If g_KibouNumDiaLogFlg = 1 Then
                                    'ワーニング
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "希望勤務"
                                    w_strMsg(2) = "設定された希望勤務(日付別)回数"
                                    '「&1が&2を超えています。~nこのまま登録してもよろしいですか。」
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    'エラー
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "設定された希望勤務(日付別)回数を超えているため"
                                    '「&1入力できません。」
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If
                        '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------

                        Select Case True
                            Case _OptRiyu_0.Checked '通常
                                '理由区分 通常
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                            Case _OptRiyu_1.Checked '要請
                                '理由区分 要請
                                w_RiyuKBN = "2"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", General.G_PALEGREEN))
                            Case _OptRiyu_2.Checked '希望
                                '理由区分 希望
                                w_RiyuKBN = "3"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                            Case _OptRiyu_3.Checked
                                '理由区分 再掲
                                w_RiyuKBN = "4"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", General.G_LIGHTCYAN))
                            Case _OptRiyu_4.Checked
                                w_RiyuKBN = "6"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", General.G_ORANGE))
                            Case Else
                                '理由区分 通常
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                        End Select

                        '区分が応援の場合のみ、応援先勤務地選択画面を表示
                        If w_RiyuKBN = "6" Then
                            If Disp_Ouen(w_KangoCD) = False Then
                                Exit Sub
                            End If
                        End If

                        '2015/04/13 Bando Add Start ========================
                        '希望の場合コメント入力画面表示
                        If w_RiyuKBN = "3" And g_InputHopeCommentFlg = "1" Then
                            If Disp_Comment(w_Comment, w_RiyuKBN) = False Then
                                Exit Sub
                            End If
                        End If
                        '2015/04/13 Bando Add End   ========================

                        '2015/07/22 Bando Add Start ========================
                        If w_Comment <> "" Then
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Fore", Convert.ToString(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Back", ColorTranslator.ToOle(Color.IndianRed)))
                        End If
                        '2015/07/22 Bando Add End   ========================

                        '代休発生勤務ﾊﾞｯｸｶﾗｰ
                        If General.g_lngDaikyuMng = 0 Then
                            If m_DaikyuBackColorFlg = True Then
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Back", General.G_LIGHTYELLOW))
                            End If
                        End If

                        'ｾﾙに記号を設定
                        '2015/04/13 Bando Upd Start =======================
                        'w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, w_Comment)
                        '2015/04/13 Bando Upd End   =======================
                        .SetText(w_ActiveRow, w_ActiveCol, w_Var)
                        '理由別色設定
                        .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                        .Cells(w_ActiveRow, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)

                        '勤務変更の時
                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then

                            '2015/04/13 Bando Upd Start ====================
                            'Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time)
                            Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time, "")
                            '2015/04/13 Bando Upd End   ====================

                            '予定と異なる場合色変更
                            If Trim(w_KinmuCD) <> Trim(w_KinmuPlanCD) Or w_RiyuKBN <> w_RiyuPlanKbn Or w_KangoCD <> w_KangoPlanCD Then
                                .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(m_Comp_Fore)
                            End If

                            '変更可能行の内容を格納
                            w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                            w_ForeColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).ForeColor)
                            w_BackColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor)

                            '変更可能行の内容を勤務変更画面に貼り付け行へコピー
                            .SetText(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol, w_Var)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        End If
                    End With
                End If
            End If

            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
            If InStr(m_strUpdKojyoDate, w_Var) = 0 Then
                m_strUpdKojyoDate = m_strUpdKojyoDate & w_Var & ","
            End If

            'ｶｰｿﾙｾﾙ移動
            If w_DaikyuInputFlg = True Then
                Call SetCursol()
            End If

            '更新ﾌﾗｸﾞｾｯﾄ
            If w_InputFlg = True Then
                m_KosinFlg = True
            End If
            '要修正 ed ********************

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '超勤データの有無チェック
    Private Function Check_OverKinmuData(ByVal p_Date As Object) As Boolean
        Const W_SUBNAME As String = "NSK0000HA Check_OverKinmuData"


        Dim w_strMsg() As String

        Check_OverKinmuData = False
        Try
            '超勤データ取得
            With General.g_objGetData
                .p病院CD = General.g_strHospitalCD
                .p承認区分 = 2 '0:調整F、1:明細F、2:両方
                .p職員区分 = 0 '0:職員管理番号
                .p職員番号 = M_StaffID '選択職員管理番号
                .p日付区分 = 0 '0:単一日
                .p開始年月日 = p_Date '開始年月日
                .p終了年月日 = 0 '終了年月日

                If .mGetOverKinmu = True Then
                    ReDim w_strMsg(1)
                    w_strMsg(1) = "時間外が既に登録されているため~n"
                    Call General.paMsgDsp("NS0110", w_strMsg)
                    Exit Function
                End If
            End With

            Check_OverKinmuData = True

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    '代休データチェック（背景色）
    Private Function Check_DaikyuBackColor(ByVal p_Date As Integer) As Boolean

        Const W_SUBNAME As String = "NSK0000HC Check_DaikyuBackColor"

        Dim w_Int As Integer

        '初期値
        Check_DaikyuBackColor = False
        Try
            '代休データループ
            For w_Int = 1 To UBound(m_DaikyuData)
                If m_DaikyuData(w_Int).HolDate = p_Date Or (m_SundayDaikyuFlg = 1 And Weekday(CDate(Format(p_Date, "0000/00/00"))) = 1) Or (m_SaturdayDaikyuFlg = 1 And Weekday(CDate(Format(p_Date, "0000/00/00"))) = 7) Then
                    Check_DaikyuBackColor = True
                    Exit Function
                End If
            Next w_Int

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    '勤務貼付け
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click

        Const W_SUBNAME As String = "NSK0000HC  cmdOK_Click"

        Dim w_Cnt As Short
        Dim w_Kinmu As Object
        Dim w_Color As Integer
        Dim w_Row As Integer '勤務記号ﾃﾞｰﾀ 開始ｾﾙ 位置
        Dim w_Kinmu_Param As Object
        Dim w_Row_Param As Integer
        Dim w_blnUpdFLG As Boolean
        Dim w_Style As New StyleInfo()
        Try
            If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                w_Row = M_KinmuData_Row_ChgJisseki + 1
            Else
                w_Row = M_KinmuData_Row
            End If

            '選択行を勤務予定の行に統一する
            m_CUR_ROW_Param = m_CUR_ROW_Param - ((m_CUR_ROW_Param - m_StaffStartRow) Mod m_MaxShowLine) + m_KinmuPlan

            For w_Cnt = 1 To (m_KeikakuD_EndCol - m_KeikakuD_StartCol + 1)

                '勤務取得
                w_Kinmu = sprSheet.Sheets(0).GetText(w_Row, m_KeikakuD_StartCol + w_Cnt - 1) '勤務変更の時

                '更新する？
                w_blnUpdFLG = False
                If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If m_blnOneTwo Then
                        '二段表示の時　下段を更新
                        w_Row_Param = m_CUR_ROW_Param + 1
                        '上段の背景色を確認
                        w_Style.Reset()
                        w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(m_CUR_ROW_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                        If w_Style.BackColor.A = 0 Then
                            w_Color = ColorTranslator.ToOle(Color.White)
                        Else
                            w_Color = ColorTranslator.ToOle(w_Style.BackColor)
                        End If
                    Else
                        '一段表示の時　上段を更新
                        w_Row_Param = m_CUR_ROW_Param
                        '下段の背景色を確認
                        w_Style.Reset()
                        w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                        If w_Style.BackColor.A = 0 Then
                            w_Color = ColorTranslator.ToOle(Color.White)
                        Else
                            w_Color = ColorTranslator.ToOle(w_Style.BackColor)
                        End If
                    End If

                    w_Kinmu_Param = m_Control_Param.Sheets(0).GetText(m_CUR_ROW_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1)
                    If m_Control_Param.Sheets(0).GetText(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1) <> "" Or w_Color = ColorTranslator.ToOle(Color.Black) Then
                        '下段にデータが存在するとき または 背景色が黒のとき　更新
                        w_blnUpdFLG = True
                    ElseIf w_Kinmu_Param <> "" Or (w_Kinmu_Param = "" And m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN) Then
                        '下段にデータが存在しないとき
                        If w_Kinmu_Param <> w_Kinmu Then
                            '上段のデータと異なる勤務のとき　更新
                            w_blnUpdFLG = True
                            If m_blnOneTwo = False Then
                                '一段表示の時　上段のデータを下段にコピー

                                '親画面勤務貼付け
                                If m_Control_Param.Sheets(0).GetText(m_CUR_ROW_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1) = "" And m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                                    m_Control_Param.Sheets(0).SetText(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Kinmu)
                                Else
                                    m_Control_Param.Sheets(0).SetText(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Kinmu_Param)
                                End If

                                '色取得
                                w_Style.Reset()
                                w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(m_CUR_ROW_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                                w_Color = ColorTranslator.ToOle(w_Style.ForeColor)

                                '色貼付け
                                w_Style.Reset()
                                w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                                w_Style.ForeColor = ColorTranslator.FromOle(w_Color)
                                m_Control_Param.Sheets(0).Models.Style.SetDirectInfo(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)

                                '色取得
                                w_Style.Reset()
                                w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(m_CUR_ROW_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                                If w_Style.BackColor.A = 0 Then
                                    w_Color = ColorTranslator.ToOle(Color.White)
                                Else
                                    w_Color = ColorTranslator.ToOle(w_Style.BackColor)
                                End If

                                '色貼付け
                                w_Style.Reset()
                                w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                                w_Style.BackColor = ColorTranslator.FromOle(w_Color)
                                m_Control_Param.Sheets(0).Models.Style.SetDirectInfo(m_CUR_ROW_Param + 1, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                            End If
                        End If
                    Else
                        '上段にデータが存在しないとき　上段を更新
                        If m_Mode <> General.G_PGMSTARTFLG_CHANGEPLAN Then
                            w_Row_Param = m_CUR_ROW_Param
                        End If

                        w_blnUpdFLG = True
                    End If
                Else
                    w_Row_Param = m_CUR_ROW_Param
                    w_blnUpdFLG = True
                End If

                If w_blnUpdFLG Then
                    '親画面勤務貼付け
                    m_Control_Param.Sheets(0).SetText(w_Row_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Kinmu)

                    '色取得
                    w_Color = ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_Row, m_KeikakuD_StartCol + w_Cnt - 1).ForeColor)

                    '色貼付け
                    w_Style.Reset()
                    w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(w_Row_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                    w_Style.ForeColor = ColorTranslator.FromOle(w_Color)
                    m_Control_Param.Sheets(0).Models.Style.SetDirectInfo(w_Row_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)

                    '色取得
                    If sprSheet.Sheets(0).Cells(w_Row, m_KeikakuD_StartCol + w_Cnt - 1).BackColor.A = 0 Then
                        w_Color = ColorTranslator.ToOle(Color.White)
                    Else
                        w_Color = ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_Row, m_KeikakuD_StartCol + w_Cnt - 1).BackColor)
                    End If

                    '色貼付け
                    w_Style.Reset()
                    w_Style = m_Control_Param.Sheets(0).Models.Style.GetDirectInfo(w_Row_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                    w_Style.BackColor = ColorTranslator.FromOle(w_Color)
                    m_Control_Param.Sheets(0).Models.Style.SetDirectInfo(w_Row_Param, m_KeikakuD_StartCol_Param + w_Cnt - 1, w_Style)
                End If
            Next w_Cnt

            m_OKFlg = True

            '画面消去
            Me.Close()
        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    'Private Sub m_lstCmdSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _CmdSet_0.Click, _CmdSet_1.Click, _CmdSet_2.Click, _CmdSet_3.Click, _
    '                                                                                                                _CmdSet_4.Click, _CmdSet_5.Click, _CmdSet_6.Click, _CmdSet_7.Click, _
    '                                                                                                                _CmdSet_8.Click, _CmdSet_9.Click, _CmdSet_10.Click, _CmdSet_11.Click, _
    '                                                                                                                _CmdSet_12.Click, _CmdSet_13.Click, _CmdSet_14.Click, _CmdSet_15.Click, _
    '                                                                                                                _CmdSet_16.Click, _CmdSet_17.Click, _CmdSet_18.Click, _CmdSet_19.Click
    Private Sub m_lstCmdSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim Index As Short = m_lstCmdSet.IndexOf(eventSender)
        Const W_SUBNAME As String = "NSK0000HC  m_lstCmdSet_Click"

        Dim w_RegStr As String
        Dim w_Var As Object
        Dim w_ActiveCol As Integer
        Dim w_ActiveRow As Integer
        Dim w_KinmuCD As String 'KinmuCD
        Dim w_RiyuKBN As String '理由区分
        Dim w_Time As String '時間年休
        Dim w_Flg As String '確定ﾌﾗｸﾞ
        Dim w_ForeColor As Integer '文字色
        Dim w_BackColor As Integer '背景色
        Dim w_KinmuPlanCD As String 'KinmuCD(予定ﾃﾞｰﾀ)
        Dim w_STS As Short
        Dim w_KangoCD As String
        Dim w_KangoPlanCD As String
        Dim w_RiyuPlanKbn As String
        Dim w_YYYYMMDD As Integer
        Dim w_Col As Integer
        Dim w_StopFlg As Boolean
        Dim w_DaikyuCnt As Integer
        Dim w_MsgFlg As Boolean
        Dim w_strMsg() As String
        Dim w_lngBackColor As Integer
        Dim w_IntCol As Short
        Dim w_KibouCnt As Short
        Dim w_objSyouninData As Object
        Dim w_ColCnt As Short
        Dim w_lngLoop As Integer
        Try
            'ﾌｫｰｶｽ移動
            sprSheet.Focus()

            ''ﾚｼﾞｽﾄﾘ格納先
            w_RegStr = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            '１件でも存在すれば・・・
            If m_SetCnt <> 0 Then

                w_ActiveCol = sprSheet.Sheets(0).ActiveColumn.Index
                w_ActiveRow = sprSheet.Sheets(0).ActiveRow.Index

                w_ColCnt = w_ActiveCol

                With sprSheet.Sheets(0)
                    '警告表示+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    If g_SaikeiFlg = False Then

                        w_MsgFlg = False
                        '2013/10/02 Bando Chg Start =========================================================
                        'For w_Col = 1 To m_SetKinmu(Index + HscSet.Value).KinmuCnt
                        For w_Col = 1 To m_SetKinmu(Index + HscSet.Value * 2).KinmuCnt
                            '2013/10/02 Bando Chg End =========================================================

                            '計画ﾃﾞｰﾀ列 範囲内か ?
                            If Not ((m_KeikakuD_StartCol <= w_ColCnt) And (w_ColCnt <= m_KeikakuD_EndCol)) Then
                                'ERROR
                                Exit For
                            End If

                            '2015/04/13 Bando Upd Start ==================
                            'Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                            Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, "")
                            '2015/04/13 Bando Upd End   ==================
                            Select Case w_RiyuKBN
                                Case "3"
                                    If frmNSK0000HA._mnuTool_4.Checked = True Then
                                        w_MsgFlg = True
                                        Exit For
                                    End If
                                Case "4", "5", "6"
                                    w_MsgFlg = True
                                    Exit For
                            End Select

                            '勤務の曜日制限チェック
                            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                            w_YYYYMMDD = w_Var
                            '2013/10/02 Bando Chg Start =======================================================
                            'If Check_YoubiLimit(w_YYYYMMDD, m_SetKinmu(Index + HscSet.Value).CD(w_Col)) = False Then
                            If Check_YoubiLimit(w_YYYYMMDD, m_SetKinmu(Index + HscSet.Value * 2).CD(w_Col)) = False Then
                                '2013/10/02 Bando Chg End =======================================================
                                Exit Sub
                            End If

                            '2013/01/07 Ishiga add start------------------------------------------
                            '超勤データの有無チェック
                            If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                                If Check_OverKinmuData(w_YYYYMMDD) = False Then
                                    Exit Sub
                                End If
                            End If
                            '2013/01/07 Ishiga add end--------------------------------------------

                            '届出存在チェック
                            If fncChkAppliData(w_YYYYMMDD) = False Then
                                Exit Sub
                            End If


                            If m_PackageFLG = 1 Or m_PackageFLG = 3 Then
                                '日当直のコンポーネントありの時
                                With General.g_objGetData
                                    .p職員区分 = 0
                                    .p職員番号 = M_StaffID
                                    .pチェック基準日 = w_YYYYMMDD
                                    .p処理区分 = 0
                                    '2013/10/02 Bando Chg Start =======================================================
                                    '.pチェック勤務CD = m_SetKinmu(Index + HscSet.Value).CD(w_Col)
                                    .pチェック勤務CD = m_SetKinmu(Index + HscSet.Value * 2).CD(w_Col)
                                    '2013/10/02 Bando Chg End   =======================================================

                                    If .mChkKinmuDuty = False Then
                                        '勤務変更不可
                                        '*******ﾒｯｾｰｼﾞ***********************************
                                        ReDim w_strMsg(1)
                                        w_strMsg(1) = ""
                                        Call General.paMsgDsp("NS0110", w_strMsg)
                                        '************************************************
                                        Exit Sub
                                    End If
                                End With
                            End If

                            w_ActiveCol = w_ActiveCol + 1

                            w_ColCnt = w_ColCnt + 1
                        Next w_Col

                        If w_MsgFlg = True Then
                            ReDim w_strMsg(0)
                            w_STS = General.paMsgDsp("NS0120", w_strMsg)
                            If w_STS = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                    '希望回数集計チェック
                    w_lngBackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                    If CDbl(g_HopeNumFlg) = 1 And _OptRiyu_2.Checked = True Then
                        w_KibouCnt = 0
                        For w_IntCol = m_KeikakuD_StartCol To m_KeikakuD_EndCol
                            If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_IntCol).BackColor) = w_lngBackColor Then
                                w_KibouCnt = w_KibouCnt + 1
                            End If
                        Next w_IntCol
                        '2013/10/02 Bando Chg Start ==============================================================
                        'If g_HopeNum < w_KibouCnt + m_SetKinmu(Index + HscSet.Value).KinmuCnt Then
                        If g_HopeNum < w_KibouCnt + m_SetKinmu(Index + HscSet.Value * 2).KinmuCnt Then
                            '2013/10/02 Bando Chg End ==============================================================
                            '希望回数制限オーバーダイアログ表示
                            If g_KibouNumDiaLogFlg = 1 Then
                                'ワーニング
                                ReDim w_strMsg(2)
                                w_strMsg(1) = "希望勤務"
                                w_strMsg(2) = "設定された希望勤務回数"
                                '「&1が&2を超えています。~nこのまま登録してもよろしいですか。」
                                w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                If w_STS = MsgBoxResult.No Then
                                    Exit Sub
                                End If
                            Else
                                'エラー
                                ReDim w_strMsg(1)
                                w_strMsg(1) = "設定された希望勤務回数を超えているため"
                                '「&1入力できません。」
                                w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                Exit Sub
                            End If
                        End If
                    End If

                    w_ActiveCol = sprSheet.Sheets(0).ActiveColumn.Index
                    w_ActiveRow = sprSheet.Sheets(0).ActiveRow.Index

                    '2013/10/02 Bando Chg Start ==========================================
                    'For w_Col = 1 To m_SetKinmu(Index + HscSet.Value).KinmuCnt
                    For w_Col = 1 To m_SetKinmu(Index + HscSet.Value * 2).KinmuCnt
                        '2013/10/02 Bando Chg End ==========================================

                        w_StopFlg = False

                        '貼り付けるｽﾍﾟｰｽがない場合は処理終了
                        If m_KeikakuD_EndCol < w_ActiveCol Then
                            w_ActiveCol = w_ActiveCol - 1
                            Exit For
                        End If

                        '入力場所ﾁｪｯｸ
                        If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row Then
                            w_ActiveCol = w_ActiveCol - 1
                            Exit For
                        End If

                        '部署異動ﾁｪｯｸ（配属範囲）
                        If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_ActiveCol).BackColor) = m_MonthBefore_Back Then
                            '配属期間外(背景色がグレー)の場合入力不可
                            Exit Sub
                        End If

                        If General.g_lngDaikyuMng = 0 Then
                            '代休ﾁｪｯｸ
                            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                            w_YYYYMMDD = w_Var
                            'すでに代休取得済みのデータがあるか
                            For w_DaikyuCnt = 1 To UBound(m_DaikyuData)
                                If m_DaikyuData(w_DaikyuCnt).HolDate = w_YYYYMMDD Then
                                    For w_lngLoop = 1 To UBound(m_DaikyuData(w_DaikyuCnt).DaikyuDetail)
                                        If m_DaikyuData(w_DaikyuCnt).DaikyuDetail(w_lngLoop).DaikyuDate <> 0 Then
                                            w_StopFlg = True
                                            Exit For
                                        End If
                                    Next w_lngLoop

                                    If w_StopFlg = True Then
                                        Exit For
                                    End If
                                End If
                            Next w_DaikyuCnt
                        End If

                        'もし代休取得済みのデータがあれば処理終了
                        If General.g_lngDaikyuMng = 0 Then
                            If w_StopFlg = True Then
                                w_ActiveCol = w_ActiveCol - 1
                                Exit For
                            Else
                                '代休取得済みのデータがない場合は、代休配列更新
                                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                                w_YYYYMMDD = w_Var
                                'セット勤務でも代休を貼る
                                '2013/10/02 Bando Chg Start ==========================================================
                                'Call Check_Daikyu(M_PASTE, w_YYYYMMDD, m_SetKinmu(Index + HscSet.Value).CD(w_Col))
                                Call Check_Daikyu(M_PASTE, w_YYYYMMDD, m_SetKinmu(Index + HscSet.Value * 2).CD(w_Col))
                                '2013/10/02 Bando Chg End   ==========================================================
                            End If
                        End If

                        '2013/10/02 Bando Chg Start ==========================================================
                        'w_KinmuCD = m_SetKinmu(Index + HscSet.Value).CD(w_Col)
                        w_KinmuCD = m_SetKinmu(Index + HscSet.Value * 2).CD(w_Col)
                        '2013/10/02 Bando Chg End   ==========================================================
                        w_Time = ""
                        w_Flg = "0"
                        w_KangoCD = ""

                        Select Case True
                            Case _OptRiyu_0.Checked '通常
                                '理由区分 通常
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                            Case _OptRiyu_1.Checked '要請
                                '理由区分 要請
                                w_RiyuKBN = "2"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", General.G_PALEGREEN))
                            Case _OptRiyu_2.Checked '希望
                                '理由区分 希望
                                w_RiyuKBN = "3"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                            Case _OptRiyu_3.Checked
                                '理由区分 再掲
                                w_RiyuKBN = "4"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", General.G_LIGHTCYAN))
                            Case Else
                                '理由区分 通常
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                        End Select

                        'セット勤務でも代休を貼る
                        If General.g_lngDaikyuMng = 0 Then
                            If m_DaikyuBackColorFlg = True Then
                                '代休発生勤務　文字/背景色
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Back", General.G_LIGHTYELLOW))
                            End If
                        End If

                        'ｾﾙに記号を設定する
                        '2015/04/13 Bando Upd Start =====================
                        'w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, "")
                        '2015/04/13 Bando Upd End   =====================
                        .SetText(w_ActiveRow, w_ActiveCol, w_Var)
                        '理由別の色設定
                        .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                        .Cells(w_ActiveRow, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)

                        '勤務変更の時
                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            '2015/04/13 Bando Upd Start =====================
                            'Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time)
                            Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time, "")
                            '2015/04/13 Bando Upd End   =====================
                            '予定と異なる場合色変更
                            If Trim(w_KinmuCD) <> Trim(w_KinmuPlanCD) Or w_RiyuKBN <> w_RiyuPlanKbn Or w_KangoCD <> w_KangoPlanCD Then
                                .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(m_Comp_Fore)
                            End If

                            '変更可能行の内容を格納
                            w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                            w_ForeColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).ForeColor)
                            w_BackColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor)
                            '変更可能行の内容を勤務変更画面に貼り付け行へコピー
                            .SetText(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol, w_Var)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        End If

                        '更新ﾌﾗｸﾞｾｯﾄ
                        m_KosinFlg = True

                        w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                        If InStr(m_strUpdKojyoDate, w_Var) = 0 Then
                            m_strUpdKojyoDate = m_strUpdKojyoDate & w_Var & ","
                        End If

                        '2013/10/02 Bando Chg Start ===================================
                        'If w_Col <> m_SetKinmu(Index + HscSet.Value).KinmuCnt Then
                        If w_Col <> m_SetKinmu(Index + HscSet.Value * 2).KinmuCnt Then
                            '2013/10/02 Bando Chg End ===================================
                            w_ActiveCol = w_ActiveCol + 1
                        End If
                    Next w_Col
                End With
            End If

            'ｶｰｿﾙｾﾙ移動
            'ｾﾙ位置設定(セット勤務の最後のセルにあわす)
            sprSheet.Sheets(0).SetActiveCell(w_ActiveRow, w_ActiveCol)

            Call SetCursol()

            If w_StopFlg = True Then
                ReDim w_strMsg(1)
                w_strMsg(1) = "代休取得済み勤務が存在するため"
                Call General.paMsgDsp("NS0107", w_strMsg)
            End If
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    'Private Sub m_lstCmdTokusyu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _CmdTokusyu_0.Click, _CmdTokusyu_2.Click, _CmdTokusyu_4.Click, _CmdTokusyu_6.Click, _
    '                                                                                                                    _CmdTokusyu_8.Click, _CmdTokusyu_10.Click, _CmdTokusyu_12.Click, _CmdTokusyu_14.Click, _
    '                                                                                                                    _CmdTokusyu_16.Click, _CmdTokusyu_18.Click, _CmdTokusyu_20.Click, _CmdTokusyu_22.Click, _
    '                                                                                                                    _CmdTokusyu_24.Click, _CmdTokusyu_26.Click, _CmdTokusyu_28.Click, _CmdTokusyu_30.Click, _
    '                                                                                                                    _CmdTokusyu_32.Click, _CmdTokusyu_34.Click, _CmdTokusyu_36.Click, _CmdTokusyu_38.Click
    Private Sub m_lstCmdTokusyu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim Index As Short = m_lstCmdTokusyu.IndexOf(eventSender)
        Const W_SUBNAME As String = "NSK0000HC  m_lstCmdTokusyu_Click"

        Dim w_RegStr As String
        Dim w_Var As Object
        Dim w_ActiveCol As Integer
        Dim w_ActiveRow As Integer
        Dim w_KinmuCD As String 'KinmuCD
        Dim w_RiyuKBN As String '理由区分
        Dim w_Time As String '時間年休
        Dim w_Flg As String '確定ﾌﾗｸﾞ
        Dim w_ForeColor As Integer '文字色
        Dim w_BackColor As Integer '背景色
        Dim w_InputFlg As Boolean '入力ﾌﾗｸﾞ
        Dim w_Cnt As Short
        Dim w_KinmuPlanCD As String 'KinmuCD(予定ﾃﾞｰﾀ)
        Dim w_STS As Short
        Dim w_KangoCD As String
        Dim w_KangoPlanCD As String
        Dim w_RiyuPlanKbn As String
        Dim w_YYYYMMDD As Integer
        Dim w_DaikyuInputFlg As Boolean
        Dim w_strMsg() As String
        Dim w_lngBackColor As Integer
        Dim w_IntCol As Short
        Dim w_KibouCnt As Short
        Dim w_KibouCol() As Integer
        Dim w_blnColChk As Boolean
        '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
        Dim w_KibouCntDate As Integer
        '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------
        Dim w_Comment As String = String.Empty  '希望勤務時のコメント 2015/04/13 Bando Add

        Try
            'ﾌｫｰｶｽ移動
            sprSheet.Focus()

            'ﾚｼﾞｽﾄﾘ格納先
            w_RegStr = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            '１件でも存在すれば・・・
            If m_TokusyuCnt <> 0 Then

                w_ActiveCol = sprSheet.Sheets(0).ActiveColumn.Index
                w_ActiveRow = sprSheet.Sheets(0).ActiveRow.Index

                '入力場所ﾁｪｯｸ
                If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row_ChgJisseki Then
                        Exit Sub
                    End If
                Else
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row Then
                        Exit Sub
                    End If
                End If

                '部署異動ﾁｪｯｸ（配属範囲）
                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_ActiveCol).BackColor) = m_MonthBefore_Back Then
                    '配属期間外(背景色がグレー)の場合入力不可
                    Exit Sub
                End If

                '入力ﾌﾗｸﾞ
                w_InputFlg = True
                w_Cnt = CShort(w_ActiveCol - M_KinmuData_Col + 1)
                If g_SaikeiFlg = True Then
                    '再掲部署の場合
                    If m_DataFlg(w_Cnt) = "1" Then
                        '実績ﾃﾞｰﾀの場合
                        ReDim w_strMsg(2)
                        w_strMsg(1) = "確定済み勤務"
                        w_strMsg(2) = "再掲勤務"
                        Call General.paMsgDsp("NS0011", w_strMsg)
                        w_InputFlg = False
                    End If
                End If


                '勤務の曜日制限チェック
                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                '2013/10/02 Bando Chg Start ===========================================================
                'If Check_YoubiLimit(w_YYYYMMDD, m_Tokusyu(Index + HscTokusyu.Value).CD) = False Then
                If Check_YoubiLimit(w_YYYYMMDD, m_Tokusyu(Index + HscTokusyu.Value * 2).CD) = False Then
                    '2013/10/02 Bando Chg End   ===========================================================
                    Exit Sub
                End If

                '2013/01/07 Ishiga add start------------------------------------------
                '超勤データの有無チェック
                If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If Check_OverKinmuData(w_YYYYMMDD) = False Then
                        Exit Sub
                    End If
                End If
                '2013/01/07 Ishiga add end--------------------------------------------

                '届出存在チェック
                If fncChkAppliData(w_YYYYMMDD) = False Then
                    Exit Sub
                End If

                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                If m_PackageFLG = 1 Or m_PackageFLG = 3 Then
                    '日当直のコンポーネントありの時
                    With General.g_objGetData
                        .p職員区分 = 0
                        .p職員番号 = M_StaffID
                        .pチェック基準日 = w_YYYYMMDD
                        .p処理区分 = 0
                        '2013/10/02 Bando Chg Start ======================================
                        '.pチェック勤務CD = m_Tokusyu(Index + HscTokusyu.Value).CD
                        .pチェック勤務CD = m_Tokusyu(Index + HscTokusyu.Value * 2).CD
                        '2013/10/02 Bando Chg End   ======================================

                        If .mChkKinmuDuty = False Then
                            '勤務変更不可
                            '*******ﾒｯｾｰｼﾞ***********************************
                            ReDim w_strMsg(1)
                            w_strMsg(1) = ""
                            Call General.paMsgDsp("NS0110", w_strMsg)
                            '************************************************
                            Exit Sub
                        End If
                    End With
                End If

                w_DaikyuInputFlg = True

                If General.g_lngDaikyuMng = 0 Then
                    '代休ﾁｪｯｸ
                    w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                    w_YYYYMMDD = w_Var
                    '2013/10/02 Bando Chg Start ======================================================
                    'If Check_Daikyu(M_PASTE, w_YYYYMMDD, m_Tokusyu(Index + HscTokusyu.Value).CD) = False Then
                    If Check_Daikyu(M_PASTE, w_YYYYMMDD, m_Tokusyu(Index + HscTokusyu.Value * 2).CD) = False Then
                        '2013/10/02 Bando Chg End   ======================================================
                        w_DaikyuInputFlg = False
                        w_InputFlg = False
                    End If
                End If

                If w_InputFlg = True Then
                    '入力可能な場合

                    With sprSheet.Sheets(0)

                        '警告表示+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        '変更する勤務の値が空で計画変更の場合、一つ上の勤務を取得する。
                        w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                        If w_Var = "" And m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            w_Var = .GetText(w_ActiveRow - 1, w_ActiveCol)
                        End If

                        '2015/04/13 Bando Upd Start =======================
                        'Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, w_Comment)
                        '2015/04/13 Bando Upd End   =======================
                        If g_SaikeiFlg = False Then
                            Select Case w_RiyuKBN
                                Case "2"
                                    If frmNSK0000HA._mnuTool_5.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "要請勤務"
                                        w_strMsg(2) = "編集"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "3"
                                    If frmNSK0000HA._mnuTool_4.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "希望勤務"
                                        w_strMsg(2) = "編集"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "4"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "再掲勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "5"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "委員会勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "6"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "応援勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                            End Select
                        Else
                            If w_RiyuKBN = "2" Then
                                If frmNSK0000HA._mnuTool_5.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "要請勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If

                            If w_RiyuKBN = "3" Then
                                If frmNSK0000HA._mnuTool_4.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "希望勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                        w_KinmuCD = m_Tokusyu(Index + HscTokusyu.Value * 2).CD
                        w_Time = ""
                        w_Flg = "0"
                        w_KangoCD = ""

                        '希望回数集計チェック
                        w_lngBackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                        If CDbl(g_HopeNumFlg) = 1 And _OptRiyu_2.Checked = True Then
                            w_KibouCnt = 0
                            ReDim w_KibouCol(0)
                            For w_IntCol = m_KeikakuD_StartCol To m_KeikakuD_EndCol
                                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_IntCol).BackColor) = w_lngBackColor Then
                                    w_KibouCnt = w_KibouCnt + 1
                                    ReDim Preserve w_KibouCol(w_KibouCnt)
                                    w_KibouCol(w_KibouCnt) = w_IntCol
                                End If
                            Next w_IntCol

                            '同じ場所に同じ勤務を貼り付けた場合、スルー
                            w_blnColChk = False
                            For w_IntCol = 1 To UBound(w_KibouCol)
                                If w_ActiveCol = w_KibouCol(w_IntCol) And w_KinmuCD = w_KinmuPlanCD Then
                                    w_blnColChk = True
                                End If
                            Next w_IntCol

                            If g_HopeNum <= w_KibouCnt And w_blnColChk = False Then
                                '希望回数制限オーバーダイアログ表示
                                If g_KibouNumDiaLogFlg = 1 Then
                                    'ワーニング
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "希望勤務"
                                    w_strMsg(2) = "設定された希望勤務回数"
                                    '「&1が&2を超えています。~nこのまま登録してもよろしいですか。」
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    'エラー
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "設定された希望勤務回数を超えているため"
                                    '「&1入力できません。」
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If

                        '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
                        If CDbl(g_HopeNumDateFlg) = 1 And _OptRiyu_2.Checked = True Then
                            '日付別の希望勤務数チェック
                            w_KibouCntDate = frmNSK0000HA.Get_HopeNum_Of_Date(w_YYYYMMDD)
                            If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor) = w_lngBackColor Then
                                w_KibouCntDate = w_KibouCntDate - 1
                            End If

                            If g_HopeNumDate <= w_KibouCntDate Then
                                '希望回数制限オーバーダイアログ表示
                                If g_KibouNumDiaLogFlg = 1 Then
                                    'ワーニング
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "希望勤務"
                                    w_strMsg(2) = "設定された希望勤務(日付別)回数"
                                    '「&1が&2を超えています。~nこのまま登録してもよろしいですか。」
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    'エラー
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "設定された希望勤務(日付別)回数を超えているため"
                                    '「&1入力できません。」
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If
                        '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------

                        Select Case True
                            Case _OptRiyu_0.Checked '通常
                                '理由区分 通常
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                            Case _OptRiyu_1.Checked '要請
                                '理由区分 要請
                                w_RiyuKBN = "2"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", General.G_PALEGREEN))
                            Case _OptRiyu_2.Checked '希望
                                '理由区分 希望
                                w_RiyuKBN = "3"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                            Case _OptRiyu_3.Checked
                                '理由区分 再掲
                                w_RiyuKBN = "4"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", General.G_LIGHTCYAN))
                            Case _OptRiyu_4.Checked
                                w_RiyuKBN = "6"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", General.G_ORANGE))
                            Case Else
                                '理由区分 その他（通常扱いとする）
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                        End Select

                        '区分が応援の場合のみ、応援先勤務地選択画面を表示
                        If w_RiyuKBN = "6" Then
                            If Disp_Ouen(w_KangoCD) = False Then
                                Exit Sub
                            End If
                        End If

                        '2015/04/13 Bando Upd Start ======================
                        '希望oの場合コメント入力画面表示
                        If w_RiyuKBN = "3" And g_InputHopeCommentFlg = "1" Then
                            If Disp_Comment(w_Comment, w_RiyuKBN) = False Then
                                Exit Sub
                            End If
                        End If
                        '2015/04/13 Bando Upd End   ======================

                        '2015/07/22 Bando Add Start ========================
                        If w_Comment <> "" Then
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Fore", Convert.ToString(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Back", ColorTranslator.ToOle(Color.IndianRed)))
                        End If
                        '2015/07/22 Bando Add End   ========================

                        '代休発生勤務ﾊﾞｯｸｶﾗｰ
                        If General.g_lngDaikyuMng = 0 Then
                            If m_DaikyuBackColorFlg = True Then
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Back", General.G_LIGHTYELLOW))
                            End If
                        End If

                        'ｾﾙに記号を設定する
                        '2015/04/13 Bando Upd Start ====================
                        'w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, w_Comment)
                        '2015/04/13 Bando Upd End   ====================
                        .SetText(w_ActiveRow, w_ActiveCol, w_Var)
                        '理由別の色設定
                        .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                        .Cells(w_ActiveRow, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        '勤務変更の時
                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            '2015/04/13 Bando Upd Start ====================
                            'Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time)
                            Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time, w_Comment)
                            '2015/04/13 Bando Upd End   ====================
                            '予定と異なる場合色変更
                            If Trim(w_KinmuCD) <> Trim(w_KinmuPlanCD) Or w_RiyuKBN <> w_RiyuPlanKbn Or w_KangoCD <> w_KangoPlanCD Then
                                .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(m_Comp_Fore)
                            End If

                            '変更可能行の内容を格納
                            w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                            w_ForeColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).ForeColor)
                            w_BackColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor)
                            '変更可能行の内容を勤務変更画面に貼り付け行へコピー
                            .SetText(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol, w_Var)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        End If
                    End With
                End If
            End If

            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
            If InStr(m_strUpdKojyoDate, w_Var) = 0 Then
                m_strUpdKojyoDate = m_strUpdKojyoDate & w_Var & ","
            End If

            'ｶｰｿﾙｾﾙ移動
            If w_DaikyuInputFlg = True Then
                Call SetCursol()
            End If

            '更新ﾌﾗｸﾞｾｯﾄ
            If w_InputFlg = True Then
                m_KosinFlg = True
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    'Private Sub m_lstCmdYasumi_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _cmdYasumi_0.Click, _cmdYasumi_2.Click, _cmdYasumi_4.Click, _cmdYasumi_6.Click, _
    '                                                                                                                    _cmdYasumi_8.Click, _cmdYasumi_10.Click, _cmdYasumi_12.Click, _cmdYasumi_14.Click, _
    '                                                                                                                    _cmdYasumi_16.Click, _cmdYasumi_18.Click, _cmdYasumi_20.Click, _cmdYasumi_22.Click, _
    '                                                                                                                    _cmdYasumi_24.Click, _cmdYasumi_26.Click, _cmdYasumi_28.Click, _cmdYasumi_30.Click, _
    '                                                                                                                    _cmdYasumi_32.Click, _cmdYasumi_34.Click, _cmdYasumi_36.Click, _cmdYasumi_38.Click
    Private Sub m_lstCmdYasumi_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim Index As Short = m_lstCmdYasumi.IndexOf(eventSender)
        Const W_SUBNAME As String = "NSK0000HC  m_lstCmdYasumi_Click"

        Dim w_RegStr As String
        Dim w_Var As Object
        Dim w_ActiveCol As Integer
        Dim w_ActiveRow As Integer
        Dim w_KinmuCD As String 'KinmuCD
        Dim w_RiyuKBN As String '理由区分
        Dim w_Time As String '時間年休
        Dim w_Flg As String '確定ﾌﾗｸﾞ
        Dim w_ForeColor As Integer '文字色
        Dim w_BackColor As Integer '背景色
        Dim w_InputFlg As Boolean '入力ﾌﾗｸﾞ
        Dim w_Cnt As Short
        Dim w_KinmuPlanCD As String 'KinmuCD(予定ﾃﾞｰﾀ)
        Dim w_STS As Short
        Dim w_KangoCD As String
        Dim w_KangoPlanCD As String
        Dim w_RiyuPlanKbn As String
        Dim w_YYYYMMDD As Integer
        Dim w_DaikyuInputFlg As Boolean
        Dim w_strMsg() As String
        Dim w_lngBackColor As Integer
        Dim w_IntCol As Short
        Dim w_KibouCnt As Short
        Dim w_KibouCol() As Integer
        Dim w_blnColChk As Boolean
        Dim w_objSyouninData As Object
        '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
        Dim w_KibouCntDate As Integer
        '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------
        Dim w_Comment As String = String.Empty  '希望勤務時のコメント 2015/04/13 Bando Add
        Try
            'ﾌｫｰｶｽ移動
            sprSheet.Focus()

            'ﾚｼﾞｽﾄﾘ格納先
            w_RegStr = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            '１件でも存在すれば・・・
            If m_YasumiCnt <> 0 Then

                w_ActiveCol = sprSheet.Sheets(0).ActiveColumn.Index
                w_ActiveRow = sprSheet.Sheets(0).ActiveRow.Index

                '入力場所ﾁｪｯｸ
                If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row_ChgJisseki Then
                        Exit Sub
                    End If
                Else
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row Then
                        Exit Sub
                    End If
                End If

                '部署異動ﾁｪｯｸ（配属範囲）
                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_ActiveCol).BackColor) = m_MonthBefore_Back Then
                    '配属期間外(背景色がグレー)の場合入力不可
                    Exit Sub
                End If

                '入力ﾌﾗｸﾞ
                w_InputFlg = True
                w_Cnt = CShort(w_ActiveCol - M_KinmuData_Col + 1)
                If g_SaikeiFlg = True Then
                    '再掲部署の場合
                    If m_DataFlg(w_Cnt) = "1" Then
                        '実績ﾃﾞｰﾀの場合
                        ReDim w_strMsg(2)
                        w_strMsg(1) = "確定済み勤務"
                        w_strMsg(2) = "再掲勤務"
                        Call General.paMsgDsp("NS0011", w_strMsg)

                        w_InputFlg = False
                    End If
                End If

                '勤務の曜日制限チェック
                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                '2013/10/02 Bando Chg Start ===================================================
                'If Check_YoubiLimit(w_YYYYMMDD, m_Yasumi(Index + HscYasumi.Value).CD) = False Then
                If Check_YoubiLimit(w_YYYYMMDD, m_Yasumi(Index + HscYasumi.Value * 2).CD) = False Then
                    '2013/10/02 Bando Chg End  ===================================================
                    Exit Sub
                End If

                '2013/01/07 Ishiga add start------------------------------------------
                '超勤データの有無チェック
                If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If Check_OverKinmuData(w_YYYYMMDD) = False Then
                        Exit Sub
                    End If
                End If
                '2013/01/07 Ishiga add end--------------------------------------------

                '届出存在チェック
                If fncChkAppliData(w_YYYYMMDD) = False Then
                    Exit Sub
                End If

                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                If m_PackageFLG = 1 Or m_PackageFLG = 3 Then
                    '日当直のコンポーネントありの時
                    With General.g_objGetData
                        .p職員区分 = 0
                        .p職員番号 = M_StaffID
                        .pチェック基準日 = w_YYYYMMDD
                        .p処理区分 = 0
                        '2013/10/02 Bando Chg Start ==============================
                        '.pチェック勤務CD = m_Yasumi(Index + HscYasumi.Value).CD
                        .pチェック勤務CD = m_Yasumi(Index + HscYasumi.Value * 2).CD
                        '2013/10/02 Bando Chg End   ==============================

                        If .mChkKinmuDuty = False Then
                            '勤務変更不可
                            '*******ﾒｯｾｰｼﾞ***********************************
                            ReDim w_strMsg(1)
                            w_strMsg(1) = ""
                            Call General.paMsgDsp("NS0110", w_strMsg)
                            '************************************************
                            Exit Sub
                        End If
                    End With
                End If

                w_DaikyuInputFlg = True

                If General.g_lngDaikyuMng = 0 Then
                    '代休ﾁｪｯｸ
                    w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                    w_YYYYMMDD = w_Var
                    '2013/10/02 Bando Chg Start ======================================================
                    'If Check_Daikyu(M_PASTE, w_YYYYMMDD, m_Yasumi(Index + HscYasumi.Value).CD) = False Then
                    If Check_Daikyu(M_PASTE, w_YYYYMMDD, m_Yasumi(Index + HscYasumi.Value * 2).CD) = False Then
                        '2013/10/02 Bando Chg End   ======================================================
                        w_DaikyuInputFlg = False
                        w_InputFlg = False
                    End If
                End If

                If w_InputFlg = True Then

                    With sprSheet.Sheets(0)
                        '警告表示+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        '変更する勤務の値が空で計画変更の場合、一つ上の勤務を取得する。
                        w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                        If w_Var = "" And m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            w_Var = .GetText(w_ActiveRow - 1, w_ActiveCol)
                        End If

                        '2015/04/13 Bando Upd Start ==========================
                        'Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, "")
                        '2015/04/13 Bando Upd End   ==========================

                        If g_SaikeiFlg = False Then
                            Select Case w_RiyuKBN
                                Case "2"
                                    If frmNSK0000HA._mnuTool_5.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "要請勤務"
                                        w_strMsg(2) = "編集"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "3"
                                    If frmNSK0000HA._mnuTool_4.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "希望勤務"
                                        w_strMsg(2) = "編集"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "4"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "再掲勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "5"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "委員会勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "6"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "応援勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                            End Select
                        Else
                            If w_RiyuKBN = "2" Then
                                If frmNSK0000HA._mnuTool_5.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "要請勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If

                            If w_RiyuKBN = "3" Then
                                If frmNSK0000HA._mnuTool_4.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "希望勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If

                        w_KinmuCD = m_Yasumi(Index + HscYasumi.Value * 2).CD
                        w_Time = ""
                        w_Flg = "0"
                        w_KangoCD = ""

                        '希望回数集計チェック
                        w_lngBackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                        If CDbl(g_HopeNumFlg) = 1 And _OptRiyu_2.Checked = True Then
                            w_KibouCnt = 0
                            ReDim w_KibouCol(0)
                            For w_IntCol = m_KeikakuD_StartCol To m_KeikakuD_EndCol
                                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_IntCol).BackColor) = w_lngBackColor Then
                                    w_KibouCnt = w_KibouCnt + 1
                                    ReDim Preserve w_KibouCol(w_KibouCnt)
                                    w_KibouCol(w_KibouCnt) = w_IntCol
                                End If
                            Next w_IntCol

                            '同じ場所に同じ勤務を貼り付けた場合、スルー
                            w_blnColChk = False
                            For w_IntCol = 1 To UBound(w_KibouCol)
                                If w_ActiveCol = w_KibouCol(w_IntCol) And w_KinmuCD = w_KinmuPlanCD Then
                                    w_blnColChk = True
                                End If
                            Next w_IntCol

                            If g_HopeNum <= w_KibouCnt And w_blnColChk = False Then
                                '希望回数制限オーバーダイアログ表示
                                If g_KibouNumDiaLogFlg = 1 Then
                                    'ワーニング
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "希望勤務"
                                    w_strMsg(2) = "設定された希望勤務回数"
                                    '「&1が&2を超えています。~nこのまま登録してもよろしいですか。」
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    'エラー
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "設定された希望勤務回数を超えているため"
                                    '「&1入力できません。」
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If

                        '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
                        If CDbl(g_HopeNumDateFlg) = 1 And _OptRiyu_2.Checked = True Then
                            '日付別の希望勤務数チェック
                            w_KibouCntDate = frmNSK0000HA.Get_HopeNum_Of_Date(w_YYYYMMDD)
                            If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor) = w_lngBackColor Then
                                w_KibouCntDate = w_KibouCntDate - 1
                            End If

                            If g_HopeNumDate <= w_KibouCntDate Then
                                '希望回数制限オーバーダイアログ表示
                                If g_KibouNumDiaLogFlg = 1 Then
                                    'ワーニング
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "希望勤務"
                                    w_strMsg(2) = "設定された希望勤務(日付別)回数"
                                    '「&1が&2を超えています。~nこのまま登録してもよろしいですか。」
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    'エラー
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "設定された希望勤務(日付別)回数を超えているため"
                                    '「&1入力できません。」
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If
                        '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------

                        Select Case True
                            Case _OptRiyu_0.Checked
                                '理由区分 通常
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                            Case _OptRiyu_1.Checked
                                '理由区分 要請
                                w_RiyuKBN = "2"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", General.G_PALEGREEN))
                            Case _OptRiyu_2.Checked
                                '理由区分 希望
                                w_RiyuKBN = "3"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                            Case _OptRiyu_3.Checked
                                '理由区分 再掲
                                w_RiyuKBN = "4"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", General.G_LIGHTCYAN))
                            Case _OptRiyu_4.Checked
                                w_RiyuKBN = "6"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", General.G_ORANGE))
                            Case Else
                                '理由区分 その他（通常扱いとする）
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                        End Select

                        '区分が応援の場合のみ、応援先勤務地選択画面を表示
                        If w_RiyuKBN = "6" Then
                            If Disp_Ouen(w_KangoCD) = False Then
                                Exit Sub
                            End If
                        End If

                        '2015/04/13 Bando Add Start =====================
                        '希望の場合コメント入力画面表示
                        If w_RiyuKBN = "3" And g_InputHopeCommentFlg = "1" Then
                            If Disp_Comment(w_Comment, w_RiyuKBN) = False Then
                                Exit Sub
                            End If
                        End If
                        '2015/04/13 Bando Add End   =====================

                        '2015/07/22 Bando Add Start ========================
                        If w_Comment <> "" Then
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Fore", Convert.ToString(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Back", ColorTranslator.ToOle(Color.IndianRed)))
                        End If
                        '2015/07/22 Bando Add End   ========================

                        '代休発生勤務ﾊﾞｯｸｶﾗｰ
                        If General.g_lngDaikyuMng = 0 Then
                            If m_DaikyuBackColorFlg = True Then
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Back", General.G_LIGHTYELLOW))
                            End If
                        End If

                        'ｾﾙに記号を設定する
                        '2015/04/13 Bando Upd Start ===========================
                        'w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        w_Var = Set_KinmuMark(w_KinmuCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, w_Comment)
                        '2015/04/13 Bando Upd End   ===========================
                        .SetText(w_ActiveRow, w_ActiveCol, w_Var)

                        '理由別の色設定
                        .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                        .Cells(w_ActiveRow, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)

                        '勤務変更の時
                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            '2015/04/13 Bando Upd Start =========================================
                            'Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time)
                            Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time, "")
                            '2015/04/13 Bando Upd End   =========================================

                            '予定と異なる場合色変更
                            If Trim(w_KinmuCD) <> Trim(w_KinmuPlanCD) Or w_RiyuKBN <> w_RiyuPlanKbn Or w_KangoCD <> w_KangoPlanCD Then
                                .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(m_Comp_Fore)
                            End If

                            '変更可能行の内容を格納
                            w_Var = .GetText(w_ActiveRow, w_ActiveCol)

                            w_ForeColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).ForeColor)
                            w_BackColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor)

                            '変更可能行の内容を勤務変更画面に貼り付け行へコピー
                            .SetText(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol, w_Var)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        End If
                    End With
                End If
            End If

            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
            If InStr(m_strUpdKojyoDate, w_Var) = 0 Then
                m_strUpdKojyoDate = m_strUpdKojyoDate & w_Var & ","
            End If

            'ｶｰｿﾙｾﾙ移動
            If w_DaikyuInputFlg = True Then
                Call SetCursol()
            End If

            If w_InputFlg = True Then
                '更新ﾌﾗｸﾞｾｯﾄ
                m_KosinFlg = True
            End If
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Public Sub frmNSK0000HC_Load()

        Const W_SUBNAME As String = "NSK0000HC  Form_Load"

        Dim w_Int As Short
        Dim w_str As String
        Dim w_varWork As Object
        Try
            Me.Hide()

            m_strUpdKojyoDate = ""

            'ﾚｼﾞｽﾄﾘ格納先
            Const w_RegStr As String = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            'スプレッド割り当てキー無効化
            subChgSpreadKeyMap()

            '前月 文字/背景色
            m_MonthBefore_Fore = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "MonthBefore_Fore", General.G_BLACK))
            m_MonthBefore_Back = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "MonthBefore_Back", General.G_LIGHTGRAY))

            '計画期間外の4週部分 背景色
            m_Jisseki4W_Back = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Jisseki4W_Back", ColorTranslator.ToOle(Color.Cyan).ToString))

            '実績が予定と異なる場合 文字/背景色
            m_Comp_Fore = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Comp_Fore", CStr(ColorTranslator.ToOle(Color.Red))))

            '土日祝 背景色
            m_WeekEnd_Back = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "WeekEnd_Back", ColorTranslator.ToOle(Color.LavenderBlush)))

            '--- 土日背景色フラグ
            m_WeekEndColorFlg = General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY7, "WEEKENDCOLORFLG", "0", General.g_strHospitalCD)
            '--- 祝休日背景色フラグ
            m_HolidayColorFlg = General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY7, "HOLIDAYCOLORFLG", "0", General.g_strHospitalCD)

            '代休の有効期間を求める(ﾃﾞﾌｫﾙﾄは８週間)
            m_lngDaikyuPastPeriod = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "PASTDAIKYUPERIOD", "56", General.g_strHospitalCD))
            m_DaikyuMsgFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "DAIKYUMSGFLG", CStr(0), General.g_strHospitalCD))
            m_SundayDaikyuFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "SUNDAYDAIKYUFLG", CStr(0), General.g_strHospitalCD))
            m_DaikyuAdvFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "DAIKYUADVANCEFLG", CStr(0), General.g_strHospitalCD))
            m_SaturdayDaikyuFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "SATURDAYDAIKYUFLG", CStr(0), General.g_strHospitalCD))

            '応援勤務区分の表示FLG
            m_OuenDispFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY7, "OUENDISPFLG", "1", General.g_strHospitalCD))

            '代休先取り当月制限フラグ
            m_DaikyuAdvThisMonthFlg = Integer.Parse(General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "DAIKYUADVANCETHISMONTHFLG", CStr(0), General.g_strHospitalCD))

            '1.5日分の代休が発生する勤務ＣＤ
            w_str = General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY5, "DAIKYU15KINNMUCD", "", General.g_strHospitalCD)
            w_varWork = General.paSplit(w_str, ",")
            ReDim m_Daikyu15KinmuCD(UBound(w_varWork) + 1)
            For w_Int = 0 To UBound(w_varWork)
                m_Daikyu15KinmuCD(w_Int + 1) = w_varWork(w_Int)
            Next w_Int

            '2014/04/23 Saijo add start P-06979-----------------------------------
            '勤務記号全角２文字対応フラグ(0：対応しない、1:対応する)
            m_strKinmuEmSecondFlg = Get_ItemValue(General.g_strHospitalCD)
            '2014/04/23 Saijo add end P-06979-------------------------------------

            '2015/04/14 Bando Add Start ========================================
            '希望モード時の表示対象勤務CD
            m_DispKinmuCd = General.paGetItemValue(General.G_STRMAINKEY2, General.G_STRSUBKEY15, "DISPKINMUCD", "", General.g_strHospitalCD)
            '2015/04/14 Bando Add End   ========================================

            Call Get_PackageUseFLG()

            '実績で使用する場合は、理由区分の入力は行わない。1:計画、2:実績
            If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                '勤務変更（実績修正）の場合
                _OptRiyu_0.Enabled = True '通常
                _OptRiyu_0.Checked = True
                _OptRiyu_1.Enabled = False '要請
                _OptRiyu_2.Enabled = False '希望
                _OptRiyu_3.Enabled = False '再掲
                _OptRiyu_3.Visible = False

                If m_OuenDispFlg = 0 Then
                    _OptRiyu_4.Enabled = True '応援
                Else
                    _OptRiyu_4.Enabled = False
                    _OptRiyu_4.Visible = False
                End If

                chkSet.CheckState = CheckState.Unchecked
                chkSet.Enabled = False
                _Frame_3.Enabled = False
                chkSet.Visible = False
                _Frame_3.Visible = False

                Me.Text = "個別勤務変更"

                '消しゴム使用不可
                cmdErase.Enabled = False
                cmdErase.Visible = False
            Else
                '計画 の場合
                '再掲部署の場合は、再掲のみを使用可にする
                If g_SaikeiFlg = True Then
                    _OptRiyu_0.Enabled = False '通常
                    _OptRiyu_0.Visible = False
                    _OptRiyu_1.Enabled = False '要請
                    _OptRiyu_2.Enabled = False '希望
                    _OptRiyu_3.Enabled = True '再掲
                    _OptRiyu_3.Visible = True
                    _OptRiyu_3.Checked = True
                    _OptRiyu_4.Enabled = False '応援

                    If m_OuenDispFlg = 1 Then
                        _OptRiyu_4.Visible = False
                    End If

                    chkSet.Visible = False
                    _Frame_3.Enabled = True
                Else
                    '計画 の場合
                    _OptRiyu_0.Enabled = True '通常
                    _OptRiyu_0.Checked = True
                    _OptRiyu_1.Enabled = True '要請


                    '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                    'If g_HopeNumFlg = "1" And g_HopeNum = 0 Then
                    If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                        '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                        '希望回数制限あり　かつ　希望回数0回　の場合
                        _OptRiyu_2.Enabled = False '希望
                    Else
                        '以外
                        _OptRiyu_2.Enabled = True '希望
                    End If

                    _OptRiyu_3.Enabled = True '再掲
                    _OptRiyu_3.Visible = False

                    If m_OuenDispFlg = 0 Then
                        _OptRiyu_4.Enabled = True '応援
                    Else
                        _OptRiyu_4.Enabled = False
                        _OptRiyu_4.Visible = False
                    End If

                    chkSet.CheckState = CheckState.Unchecked
                    _Frame_3.Enabled = False
                End If

                Me.Text = "個別勤務計画作成"
                '消しゴム使用可
                cmdErase.Enabled = True
            End If

            '希望入力モードの場合,理由区分希望のみ使用可
            If g_SaikeiFlg = False Then
                If g_LimitedFlg = True Then
                    _OptRiyu_0.Enabled = False '通常
                    _OptRiyu_1.Enabled = False '要請

                    '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                    'If g_HopeNumFlg = "1" And g_HopeNum = 0 Then
                    If (g_HopeNumFlg = "1" And g_HopeNum = 0) Or (g_HopeNumDateFlg = "1" And g_HopeNumDate = 0) Then
                        '2014/05/14 Shimpo upd end P-06991-------------------------------------------------------------------------
                        '希望回数制限あり　かつ　希望回数0回　の場合
                        _OptRiyu_2.Enabled = False '希望
                        _OptRiyu_2.Checked = False
                    Else
                        '以外
                        _OptRiyu_2.Enabled = True '希望
                        _OptRiyu_2.Checked = True
                    End If

                    _OptRiyu_2.Enabled = True '希望
                    _OptRiyu_2.Checked = True
                    _OptRiyu_3.Enabled = False '再掲
                    _OptRiyu_3.Visible = False
                    _OptRiyu_4.Enabled = False '応援

                    If m_OuenDispFlg = 1 Then
                        _OptRiyu_4.Visible = False
                    End If

                    chkSet.Visible = False
                    _Frame_3.Enabled = True
                End If
            End If


            '2014/04/23 Saijo upd start P-06979---------------------------
            ''ﾌｫﾝﾄｻｲｽﾞによってﾌｫｰﾑの幅を設定
            'Select Case m_FontSize
            'Case M_FontSize_Big
            '    Me.Width = General.paTwipsTopixels(14300)
            'Case M_FontSize_Middle
            '    Me.Width = General.paTwipsTopixels(12500)
            'Case M_FontSize_Small
            '    Me.Width = General.paTwipsTopixels(10750)
            'Case Else
            'End Select
            If m_strKinmuEmSecondFlg = "0" Then
                'ﾌｫﾝﾄｻｲｽﾞによってﾌｫｰﾑの幅を設定
                Select Case m_FontSize
                    Case M_FontSize_Big
                        Me.Width = General.paTwipsTopixels(14300)
                    Case M_FontSize_Middle
                        Me.Width = General.paTwipsTopixels(12500)
                    Case M_FontSize_Small
                        Me.Width = General.paTwipsTopixels(10750)
                    Case Else
                End Select
            Else
                'ﾌｫﾝﾄｻｲｽﾞによってﾌｫｰﾑの幅を設定
                Select Case m_FontSize
                    Case M_FontSize_Second_Big
                        Me.Width = General.paTwipsTopixels(14300)
                    Case M_FontSize_Second_Middle
                        Me.Width = General.paTwipsTopixels(12500)
                    Case M_FontSize_Second_Small
                        Me.Width = General.paTwipsTopixels(10750)
                    Case Else
                End Select
            End If
            '2014/04/23 Saijo upd end P-06979-----------------------------

            If Not m_empRowDispFlg Then
                '採用条件列非表示
                sprSheet.Sheets(0).SetColumnWidth(1, 0)
            End If
            '夜勤専従・短時間情報表示
            Call DispNightShortInfo()

            '更新ﾌﾗｸﾞ初期化
            m_KosinFlg = False
            m_OKFlg = False

        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub frmNSK0000HC_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As FormClosingEventArgs) Handles Me.FormClosing
        Dim UnloadMode As CloseReason = eventArgs.CloseReason
        Const W_SUBNAME As String = "NSK0000HC  Form_QueryUnload"

        Dim w_strMsg() As String
        Dim w_MsgRc As Short
        Try
            'O.K.ﾎﾞﾀﾝ押下ﾁｪｯｸ
            If m_OKFlg = False Then
                If m_KosinFlg Then
                    'ﾒｯｾｰｼﾞを表示
                    ReDim w_strMsg(0)
                    w_MsgRc = General.paMsgDsp("NS0041", w_strMsg)

                    Select Case w_MsgRc
                        Case MsgBoxResult.Yes
                            '勤務貼付け
                            Call cmdOK_Click(cmdOK, New System.EventArgs())

                        Case MsgBoxResult.No
                            '変更は破棄する
                            m_KosinFlg = False
                            m_strUpdKojyoDate = ""

                        Case MsgBoxResult.Cancel
                            '終了中止
                            eventArgs.Cancel = True
                            Exit Sub
                    End Select
                End If
            End If

        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '次のｾﾙへｶｰｿﾙ移動
    Private Sub SetCursol()

        Const W_SUBNAME As String = "NSK0000HC  SetCursol"

        Dim w_ActiveRow As Short
        Dim w_ActiveCol As Short
        Dim w_Col As Short
        Dim w_tmpCol As Short
        Dim w_flg As Boolean = False
        Try
            With sprSheet.Sheets(0)
                w_ActiveRow = .ActiveRow.Index
                w_ActiveCol = .ActiveColumn.Index

                For w_Col = w_ActiveCol + 1 To m_KeikakuD_EndCol
                    'ｾﾙ位置設定
                    If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_MonthBefore_Back And ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_Jisseki4W_Back Then
                        If Not IsExistBackColor(sprSheet, w_ActiveRow, w_Col) Then
                            '勤務貼付け済みﾁｪｯｸ
                            If Trim(.GetText(w_ActiveRow, w_Col)) = "" Then
                                .SetActiveCell(w_ActiveRow, w_Col)
                                Exit Sub
                            End If
                        End If
                    End If
                Next w_Col

                '移動ｾﾙが存在しない場合

                'ｾﾙ位置設定
                If w_ActiveCol + 1 <= m_KeikakuD_EndCol Then
                    For w_Col = w_ActiveCol + 1 To m_KeikakuD_EndCol
                        'ｾﾙ位置設定
                        If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_MonthBefore_Back And ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_Jisseki4W_Back Then
                            If Not IsExistBackColor(sprSheet, w_ActiveRow, w_Col) Then
                                .SetActiveCell(w_ActiveRow, w_Col)
                                Exit Sub
                            End If
                        End If
                    Next w_Col

                    For w_Col = m_KeikakuD_StartCol To m_KeikakuD_EndCol
                        'ｾﾙ位置設定
                        If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_MonthBefore_Back And ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_Jisseki4W_Back Then
                            w_tmpCol = w_Col
                            If Not IsExistBackColor(sprSheet, w_ActiveRow, w_Col) Then
                                w_flg = True
                                .SetActiveCell(w_ActiveRow, w_Col)
                                Exit For
                            End If
                        End If
                    Next w_Col

                    If Not w_flg Then
                        .SetActiveCell(w_ActiveRow, w_tmpCol)
                    End If

                    w_ActiveRow = .ActiveRow.Index
                    w_ActiveCol = .ActiveColumn.Index

                    If Trim(.GetText(w_ActiveRow, w_ActiveCol)) <> "" Then
                        For w_Col = w_ActiveCol + 1 To m_KeikakuD_EndCol
                            'ｾﾙ位置設定
                            If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_MonthBefore_Back And ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_Jisseki4W_Back Then
                                If Not IsExistBackColor(sprSheet, w_ActiveRow, w_Col) Then
                                    '勤務貼付け済みﾁｪｯｸ
                                    If Trim(.GetText(w_ActiveRow, w_Col)) = "" Then
                                        .SetActiveCell(w_ActiveRow, w_Col)
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next w_Col
                    End If
                Else
                    For w_Col = m_KeikakuD_StartCol To m_KeikakuD_EndCol
                        'ｾﾙ位置設定
                        If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_MonthBefore_Back And ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_Jisseki4W_Back Then
                            w_tmpCol = w_Col
                            If Not IsExistBackColor(sprSheet, w_ActiveRow, w_Col) Then
                                w_flg = True
                                .SetActiveCell(w_ActiveRow, w_Col)
                                Exit For
                            End If
                        End If
                    Next w_Col

                    If Not w_flg Then
                        .SetActiveCell(w_ActiveRow, w_tmpCol)
                    End If

                    w_ActiveRow = .ActiveRow.Index
                    w_ActiveCol = .ActiveColumn.Index

                    If Trim(.GetText(w_ActiveRow, w_ActiveCol)) <> "" Then
                        For w_Col = w_ActiveCol + 1 To m_KeikakuD_EndCol
                            'ｾﾙ位置設定
                            If ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_MonthBefore_Back And ColorTranslator.ToOle(.Cells(w_ActiveRow, w_Col).BackColor) <> m_Jisseki4W_Back Then
                                If Not IsExistBackColor(sprSheet, w_ActiveRow, w_Col) Then
                                    '勤務貼付け済みﾁｪｯｸ
                                    If Trim(.GetText(w_ActiveRow, w_Col)) = "" Then
                                        .SetActiveCell(w_ActiveRow, w_Col)
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next w_Col
                    End If
                End If

            End With

        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '初期ｶｰｿﾙ位置設定
    Private Sub SetStartCursol()

        Const W_SUBNAME As String = "NSK0000HC  SetStartCursol"

        Dim w_Col As Short
        Dim w_Row As Short
        Dim w_Color As Integer
        Dim w_ActiveRow As Short

        Try
            With sprSheet.Sheets(0)
                w_ActiveRow = .ActiveRow.Index

                'ｾﾙ位置設定（実績変更時は予定も参照する）
                If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    w_Row = M_KinmuData_Row_ChgJisseki
                Else
                    w_Row = w_ActiveRow
                End If

                For w_Col = M_KinmuData_Col To m_KeikakuD_EndCol
                    'セルカラー取得
                    w_Color = ColorTranslator.ToOle(.Cells(w_Row, w_Col).BackColor)
                    If w_Color <> m_MonthBefore_Back And w_Color <> m_Jisseki4W_Back Then
                        If Not IsExistBackColor(sprSheet, w_Row, w_Col) Then
                            '勤務貼付け済みﾁｪｯｸ
                            If Trim(.GetText(w_Row, w_Col)) = "" Then
                                .SetActiveCell(w_Row, w_Col)
                                Exit Sub
                            End If
                        End If
                    End If
                Next w_Col

                For w_Col = M_KinmuData_Col To m_KeikakuD_EndCol
                    'セルカラー取得
                    w_Color = ColorTranslator.ToOle(.Cells(w_Row, w_Col).BackColor)
                    If w_Color <> m_MonthBefore_Back And w_Color <> m_Jisseki4W_Back Then
                        If Not IsExistBackColor(sprSheet, w_Row, w_Col) Then
                            .SetActiveCell(w_Row, w_Col)
                            Exit Sub
                        End If
                    End If
                Next w_Col

                .SetActiveCell(w_Row, M_KinmuData_Col)
            End With

        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub frmNSK0000HC_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As FormClosedEventArgs) Handles Me.FormClosed

        Const W_SUBNAME As String = "NSK0000HC  Form_Unload"
        Try
            'ｳｨﾝﾄﾞｳの表示ﾎﾟｼﾞｼｮﾝを格納する
            Call General.paPutWindowPositon(Me, General.G_STRMAINKEY2 & "\" & g_AppName)

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End Try
    End Sub

    Private Sub HscKinmu_Change(ByVal newScrollValue As Integer)

        Const W_SUBNAME As String = "NSK0000HC  HscKinmu_Change"

        'スクロールバーの更新
        Dim w_i As Short
        Dim w_Hsc_Cnt As Short
        Dim w_int As Integer
        Try
            'コマンドボタンのＣＡＰＴＩＯＮ設定
            '勤務
            w_Hsc_Cnt = newScrollValue
            For w_i = 1 To M_PARET_NUM
                w_int = w_i + w_Hsc_Cnt * 2
                If w_int <= m_KinmuCnt Then
                    m_lstCmdKinmu(w_i - 1).Text = m_Kinmu(w_int - 1).Mark
                    If m_Kinmu(w_int - 1).Setumei = "" Then
                        ToolTip1.SetToolTip(m_lstCmdKinmu(w_i - 1), Get_KinmuTipText(m_Kinmu(w_int - 1).CD))
                    Else
                        ToolTip1.SetToolTip(m_lstCmdKinmu(w_i - 1), Get_KinmuTipText(m_Kinmu(w_int - 1).CD) & "：" & m_Kinmu(w_int - 1).Setumei)
                    End If
                    m_lstCmdKinmu(w_i - 1).Enabled = True
                    m_lstCmdKinmu(w_i - 1).visible = True
                Else
                    m_lstCmdKinmu(w_i - 1).Text = ""
                    m_lstCmdKinmu(w_i - 1).Enabled = False
                    m_lstCmdKinmu(w_i - 1).visible = False
                End If
            Next w_i

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub HscSet_Change(ByVal newScrollValue As Integer)

        Const W_SUBNAME As String = "NSK0000HC  HscSet_Change"

        'スクロールバーの更新
        Dim w_i As Short
        Dim w_Hsc_Cnt As Short
        Dim w_int As Integer
        Try
            'コマンドボタンのＣＡＰＴＩＯＮ設定
            '特殊勤務
            w_Hsc_Cnt = newScrollValue
            For w_i = 1 To M_PARET_NUM_SET
                w_int = w_i + w_Hsc_Cnt * 2
                If w_int <= m_SetCnt Then
                    m_lstCmdSet(w_i - 1).Text = m_SetKinmu(w_int - 1).Mark
                    ToolTip1.SetToolTip(m_lstCmdSet(w_i - 1), Get_SetKinmuTipText(w_int - 1))
                    m_lstCmdSet(w_i - 1).Enabled = True
                    m_lstCmdSet(w_i - 1).visible = True
                Else
                    m_lstCmdSet(w_i - 1).Text = ""
                    m_lstCmdSet(w_i - 1).Enabled = False
                    m_lstCmdSet(w_i - 1).visible = False
                End If
            Next w_i

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub HscTokusyu_Change(ByVal newScrollValue As Integer)

        Const W_SUBNAME As String = "NSK0000HC  HscTokusyu_Change"

        'スクロールバーの更新
        Dim w_i As Short
        Dim w_Hsc_Cnt As Short
        Dim w_int As Integer
        Try
            'コマンドボタンのＣＡＰＴＩＯＮ設定
            '特殊勤務
            w_Hsc_Cnt = newScrollValue
            For w_i = 1 To M_PARET_NUM
                w_int = w_i + w_Hsc_Cnt * 2
                If w_int <= m_TokusyuCnt Then
                    m_lstCmdTokusyu(w_i - 1).Text = m_Tokusyu(w_int - 1).Mark
                    If m_Tokusyu(w_int - 1).Setumei = "" Then
                        ToolTip1.SetToolTip(m_lstCmdTokusyu(w_i - 1), Get_KinmuTipText(m_Tokusyu(w_int - 1).CD))
                    Else
                        ToolTip1.SetToolTip(m_lstCmdTokusyu(w_i - 1), Get_KinmuTipText(m_Tokusyu(w_int - 1).CD) & "：" & m_Tokusyu(w_int - 1).Setumei)
                    End If
                    m_lstCmdTokusyu(w_i - 1).Enabled = True
                    m_lstCmdTokusyu(w_i - 1).visible = True
                Else
                    m_lstCmdTokusyu(w_i - 1).Text = ""
                    m_lstCmdTokusyu(w_i - 1).Enabled = False
                    m_lstCmdTokusyu(w_i - 1).visible = False
                End If
            Next w_i

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub HscYasumi_Change(ByVal newScrollValue As Integer)
        Const W_SUBNAME As String = "NSK0000HC  HscYasumi_Change"

        'スクロールバーの更新
        Dim w_i As Short
        Dim w_Hsc_Cnt As Short
        Dim w_int As Integer
        Try
            'コマンドボタンのＣＡＰＴＩＯＮ設定
            '休み
            w_Hsc_Cnt = newScrollValue
            For w_i = 1 To M_PARET_NUM
                w_int = w_i + w_Hsc_Cnt * 2
                If w_int <= m_YasumiCnt Then
                    m_lstCmdYasumi(w_i - 1).Text = m_Yasumi(w_int - 1).Mark
                    If m_Yasumi(w_int - 1).Setumei = "" Then
                        ToolTip1.SetToolTip(m_lstCmdYasumi(w_i - 1), Get_KinmuTipText(m_Yasumi(w_int - 1).CD))
                    Else
                        ToolTip1.SetToolTip(m_lstCmdYasumi(w_i - 1), Get_KinmuTipText(m_Yasumi(w_int - 1).CD) & "：" & m_Yasumi(w_int - 1).Setumei)
                    End If
                    m_lstCmdYasumi(w_i - 1).Enabled = True
                    m_lstCmdYasumi(w_i - 1).visible = True
                Else
                    m_lstCmdYasumi(w_i - 1).Text = ""
                    m_lstCmdYasumi(w_i - 1).Enabled = False
                    m_lstCmdYasumi(w_i - 1).visible = False
                End If
            Next w_i

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '採用CDを取得 (時間年休取得時使用)
    Private Function Get_SaiyoCD(ByVal p_Date As Integer, ByVal p_StaffID As String) As String
        Const W_SUBNAME As String = "NSK0000HC Get_SaiyoCD"


        Dim w_RecCnt As Integer
        Try
            '採用CDを取得
            General.g_objGetData.p病院CD = General.g_strHospitalCD
            General.g_objGetData.p職員番号 = p_StaffID '職員管理番号
            General.g_objGetData.p日付区分 = 0 '日付は単一日を指定
            General.g_objGetData.p開始年月日 = p_Date '開始年月日
            General.g_objGetData.p履歴ソートFLG = 1 '降順

            If General.g_objGetData.mStaffInit = False Then
                Get_SaiyoCD = ""
            Else
                w_RecCnt = General.g_objGetData.f職員管理件数

                '単一日指定なので必ず１件になる
                General.g_objGetData.p職員管理索引 = 1

                Get_SaiyoCD = General.g_objGetData.f採用条件CD
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    Private Sub sprSheet_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles sprSheet.CellClick
        Const W_SUBNAME As String = "NSK0000HC sprSheet_CellClick"

        Dim w_ToolTipText As String = ""

        Try
            If e.Column = 1 AndAlso e.Row = M_KinmuData_Row Then
                '採用条件列押下時、ツールチップを表示する
                If TypeOf (sprSheet.Sheets(0).GetCellType(e.Row, e.Column)) Is CellType.ButtonCellType Then
                    ToolTip1.Show(m_toolTipTxt, sprSheet)
                End If
            End If
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub sprSheet_MouseMoveEvent(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles sprSheet.MouseMove
        Const W_SUBNAME As String = "NSK0000HC sprSheet_MouseMove"

        Dim w_Col As Integer
        Dim w_Row As Integer
        Dim w_Kinmu As Object
        Dim w_str As String
        Dim w_StrData As Object
        Dim w_KBN As String
        Dim w_KangoCD As String
        Dim w_CellRange As Model.CellRange

        Static s_Col As Integer
        Static s_Row As Integer
        Try
            'ﾏｳｽ位置列／行取得
            w_CellRange = sprSheet.GetCellFromPixel(0, 0, eventArgs.X, eventArgs.Y)
            w_Row = w_CellRange.Row
            w_Col = w_CellRange.Column

            If (w_Col = s_Col) And (w_Row = s_Row) Then
                Exit Sub
            End If

            s_Col = w_Col
            s_Row = w_Row

            '計画入力範囲内（列）
            If (w_Col < m_KeikakuD_StartCol) Or (w_Col > m_KeikakuD_EndCol) Then
                ToolTip1.SetToolTip(sprSheet, "")
                Exit Sub
            End If

            '計画入力範囲内（行）
            If (w_Row < M_KinmuData_Row) Then
                ToolTip1.SetToolTip(sprSheet, "")
                Exit Sub
            End If

            '勤務取得
            w_StrData = sprSheet.Sheets(0).GetText(w_Row, w_Col)

            w_str = General.paRight(w_StrData, 11)

            '勤務存在ﾁｪｯｸ
            w_Kinmu = Trim(General.paLeft(w_str, 3))
            If w_Kinmu = "" Then
                ToolTip1.SetToolTip(sprSheet, "")
                Exit Sub
            End If

            '初期化
            ToolTip1.SetToolTip(sprSheet, "")

            '区分が"6"の場合のみ応援看護単位名称を取得しﾂｰﾙﾁｯﾌﾟに設定
            w_KBN = Mid(w_str, 4, 1)
            If w_KBN = "6" Then
                w_KangoCD = Trim(Mid(w_str, 8))
                ToolTip1.SetToolTip(sprSheet, "応援先 : " & w_KangoCD)
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    '曜日制限チェック
    Private Function Check_YoubiLimit(ByVal p_Date As Object, ByVal p_KinmuCD As String) As Boolean
        Const W_SUBNAME As String = "NSK0000HA Check_YoubiLimit"

        Dim w_intLoop1 As Short
        Dim w_intLoop2 As Short
        Dim w_chkValue As String
        Dim w_strYoubi As String
        Dim w_strMsg() As String

        Check_YoubiLimit = False
        Try
            '曜日を取得
            If InStr(m_HolDateStr, p_Date) > 0 Then
                '祝日・休日の場合は曜日エラーチェックを行わない
                If InStr(m_OffDayStr, p_Date) > 0 Then
                    w_chkValue = "8"
                    w_strYoubi = "休日"
                Else
                    w_chkValue = "7"
                    w_strYoubi = "祝日"
                End If

                '曜日制限チェック
                For w_intLoop1 = 1 To UBound(g_KinmuM)
                    If g_KinmuM(w_intLoop1).CD = p_KinmuCD Then
                        For w_intLoop2 = 1 To UBound(g_KinmuM(w_intLoop1).YoubiLimit)
                            If w_chkValue = g_KinmuM(w_intLoop1).YoubiLimit(w_intLoop2) Then
                                '制限対象の場合、処理抜け
                                ReDim w_strMsg(2)
                                w_strMsg(1) = w_strYoubi
                                w_strMsg(2) = g_KinmuM(w_intLoop1).KinmuName
                                Call General.paMsgDsp("NS0011", w_strMsg)
                                Exit Function
                            End If
                        Next w_intLoop2
                        Exit For
                    End If
                Next w_intLoop1
            Else

                Select Case Weekday(CDate(Format(Integer.Parse(p_Date), "0000/00/00")))
                    Case FirstDayOfWeek.Monday
                        w_chkValue = "0"
                        w_strYoubi = "月曜日"
                    Case FirstDayOfWeek.Tuesday
                        w_chkValue = "1"
                        w_strYoubi = "火曜日"
                    Case FirstDayOfWeek.Wednesday
                        w_chkValue = "2"
                        w_strYoubi = "水曜日"
                    Case FirstDayOfWeek.Thursday
                        w_chkValue = "3"
                        w_strYoubi = "木曜日"
                    Case FirstDayOfWeek.Friday
                        w_chkValue = "4"
                        w_strYoubi = "金曜日"
                    Case FirstDayOfWeek.Saturday
                        w_chkValue = "5"
                        w_strYoubi = "土曜日"
                    Case FirstDayOfWeek.Sunday
                        w_chkValue = "6"
                        w_strYoubi = "日曜日"
                End Select

                '曜日制限チェック
                For w_intLoop1 = 1 To UBound(g_KinmuM)
                    If g_KinmuM(w_intLoop1).CD = p_KinmuCD Then
                        For w_intLoop2 = 1 To UBound(g_KinmuM(w_intLoop1).YoubiLimit)
                            If w_chkValue = g_KinmuM(w_intLoop1).YoubiLimit(w_intLoop2) Then
                                '制限対象の場合、処理抜け
                                ReDim w_strMsg(2)
                                w_strMsg(1) = w_strYoubi
                                w_strMsg(2) = g_KinmuM(w_intLoop1).KinmuName
                                Call General.paMsgDsp("NS0011", w_strMsg)
                                Exit Function
                            End If
                        Next w_intLoop2
                        Exit For
                    End If
                Next w_intLoop1
            End If
            Check_YoubiLimit = True

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    'パッケージ情報 ﾃﾞｰﾀ取得
    Private Function Get_PackageUseFLG() As Boolean

        Const W_SUBNAME As String = "NSK0000HA Get_PackageUseFLG"

        Dim w_lngRecCnt As Integer
        Dim w_lngLoop As Integer
        Dim w_strSql As String
        Dim w_Rs As ADODB.Recordset
        Dim w_CD_F As ADODB.Field
        Dim w_FLG_F As ADODB.Field
        Dim w_strAppliUseFlg As String
        Dim w_strDutyUseFlg As String
        Dim w_strMsg() As String
        Try
            w_strAppliUseFlg = "0"
            w_strDutyUseFlg = "0"

            ''パッケージMより届出と日当直のUSEFLGを取得
            'w_strSql = ""
            'w_strSql = "Select PACKAGECD, USEFLG"
            'w_strSql = w_strSql & " From NS_PACKAGE_M"
            'w_strSql = w_strSql & " Where HospitalCD = '" & General.g_strHospitalCD & "'"

            ''ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ 生成
            'w_Rs = General.paDBRecordSetOpen(w_strSql)

            Call NSK0000H_sql.select_NS_PACKAGE_M_01(w_Rs)

            With w_Rs
                If .RecordCount <= 0 Then
                    'ﾃﾞｰﾀがないとき
                    ReDim w_strMsg(1)
                    w_strMsg(1) = "パッケージマスタ"
                    Call General.paMsgDsp("NS0032", w_strMsg)
                    .Close()
                    Exit Function
                Else
                    .MoveLast()
                    w_lngRecCnt = .RecordCount
                    .MoveFirst()

                    w_CD_F = .Fields("PACKAGECD")
                    w_FLG_F = .Fields("USEFLG")

                    For w_lngLoop = 1 To w_lngRecCnt
                        Select Case w_CD_F.Value
                            Case "A"
                                w_strAppliUseFlg = w_FLG_F.Value & ""
                            Case "D"
                                w_strDutyUseFlg = w_FLG_F.Value & ""
                        End Select

                        .MoveNext()
                    Next w_lngLoop
                End If

                .Close()
            End With

            w_Rs = Nothing

            '判定
            If w_strAppliUseFlg = "0" Then
            Else
                '項目設定
                w_strAppliUseFlg = General.paGetItemValue(General.G_STRMAINKEY1, General.G_STRSUBKEY1, "USEAPPLIFLG", "0", General.g_strHospitalCD)
            End If

            'パッケージマスタ(0:届出×日当直×,1:届出×日当直○,2:届出○日当直×,3:届出○日当直○)
            m_PackageFLG = 0

            If w_strAppliUseFlg = "1" Then
                '届出あり
                m_PackageFLG = m_PackageFLG + 2
            End If

            If w_strDutyUseFlg = "1" Then
                '日当直あり
                m_PackageFLG = m_PackageFLG + 1
            End If

            '正常終了
            Get_PackageUseFLG = True

        Catch ex As Exception
            Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    Private Sub HscKinmu_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscKinmu.Scroll
        HscKinmu_Change(eventArgs.NewValue)
    End Sub

    Private Sub HscSet_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscSet.Scroll
        HscSet_Change(eventArgs.NewValue)
    End Sub

    Private Sub HscTokusyu_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscTokusyu.Scroll
        HscTokusyu_Change(eventArgs.NewValue)
    End Sub

    Private Sub HscYasumi_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As ScrollEventArgs) Handles HscYasumi.Scroll
        HscYasumi_Change(eventArgs.NewValue)
    End Sub

    'コントロール配列の代わりにリストに格納する
    Private Sub subSetCtlList()

        General.paSetControlList(_Frame_0, "_cmdKinmu_", m_lstCmdKinmu)
        General.paSetControlList(_Frame_1, "_cmdYasumi_", m_lstCmdYasumi)
        General.paSetControlList(_Frame_2, "_CmdTokusyu_", m_lstCmdTokusyu)
        General.paSetControlList(_Frame_3, "_CmdSet_", m_lstCmdSet)

        '勤務
        For Each w_control As Button In m_lstCmdKinmu
            AddHandler w_control.Click, AddressOf m_lstCmdKinmu_Click
        Next
        '休み
        For Each w_control As Button In m_lstCmdYasumi
            AddHandler w_control.Click, AddressOf m_lstCmdYasumi_Click
        Next
        '特殊
        For Each w_control As Button In m_lstCmdTokusyu
            AddHandler w_control.Click, AddressOf m_lstCmdTokusyu_Click
        Next
        'セット
        For Each w_control As Button In m_lstCmdSet
            AddHandler w_control.Click, AddressOf m_lstCmdSet_Click
        Next
    End Sub


    ''' <summary>
    ''' 届出データ存在チェック
    ''' </summary>
    ''' <param name="p_YYYYMMDD"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncChkAppliData(ByVal p_YYYYMMDD As Integer) As Boolean

        Dim w_strMsg As String()
        Try
            If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then

                With General.g_objSyouninData
                    .pHospitalCD = General.g_strHospitalCD
                    .mKC_UNIQUESEQNO = ""
                    .mKC_StaffMngID = M_StaffID
                    .mKC_TargetDate = p_YYYYMMDD
                    .pSecurityObj = General.g_objSecurity
                    If .mChkKinmuChange = False Then
                        ReDim w_strMsg(1)
                        w_strMsg(1) = "時間外が既に登録されているため~n"
                        Call General.paMsgDsp("NS0110", w_strMsg)
                        Return False
                    End If
                End With
            End If
            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' スプレッドデフォルトキー無効化
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subChgSpreadKeyMap()
        Dim w_PreErrorProc As String = General.g_ErrorProc
        General.g_ErrorProc = "NSC0000HA subChgSpreadKeyMap"

        Dim im_m As New FarPoint.Win.Spread.InputMap

        Try
            'デフォルトで設定されているF2、F3、F4をとりあえず無効化
            '「F2」：編集モードが有効になっている場合は、アクティブセル内の値を消去
            im_m = sprSheet.GetInputMap(FarPoint.Win.Spread.InputMapMode.WhenFocused)
            im_m.Put(New FarPoint.Win.Spread.Keystroke(Keys.F2, Keys.None), FarPoint.Win.Spread.SpreadActions.None)

            im_m = sprSheet.GetInputMap(FarPoint.Win.Spread.InputMapMode.WhenAncestorOfFocused)
            im_m.Put(New FarPoint.Win.Spread.Keystroke(Keys.F2, Keys.None), FarPoint.Win.Spread.SpreadActions.None)

            '「F3」：編集モードが有効になっている場合は、日付時刻型セルに現在の日時を入力
            im_m = sprSheet.GetInputMap(FarPoint.Win.Spread.InputMapMode.WhenFocused)
            im_m.Put(New FarPoint.Win.Spread.Keystroke(Keys.F3, Keys.None), FarPoint.Win.Spread.SpreadActions.None)

            im_m = sprSheet.GetInputMap(FarPoint.Win.Spread.InputMapMode.WhenAncestorOfFocused)
            im_m.Put(New FarPoint.Win.Spread.Keystroke(Keys.F3, Keys.None), FarPoint.Win.Spread.SpreadActions.None)

            '「F4」：日付時刻型セルで編集モードが有効になっている場合は、日付を選択するためのポップアップカレンダーをスプレッドシートに表示
            im_m = sprSheet.GetInputMap(FarPoint.Win.Spread.InputMapMode.WhenFocused)
            im_m.Put(New FarPoint.Win.Spread.Keystroke(Keys.F4, Keys.None), FarPoint.Win.Spread.SpreadActions.None)

            im_m = sprSheet.GetInputMap(FarPoint.Win.Spread.InputMapMode.WhenAncestorOfFocused)
            im_m.Put(New FarPoint.Win.Spread.Keystroke(Keys.F4, Keys.None), FarPoint.Win.Spread.SpreadActions.None)

            General.g_ErrorProc = w_PreErrorProc
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' キーダウン判定
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub sprSheet_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles sprSheet.KeyDown
        Const W_SUBNAME As String = "NSK0000HC  sprSheet_KeyDown"

        Try
            If Not e.Control Then
                'コントロールキー同時押しはスルー
                If IsNumOrFuncKey(e.KeyCode) Then
                    '設定がなければスルー
                    If Not g_objKeyBoard.ContainsKey(e.KeyCode) Then Exit Sub
                    '勤務貼り付け
                    pasteKeyBoardKinmu(g_objKeyBoard(e.KeyCode))
                End If

                If e.KeyCode = Keys.Delete Then
                    cmdErase.PerformClick()
                End If
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    ''' <summary>
    ''' 勤務貼り付け（指定勤務コード）
    ''' </summary>
    ''' <param name="p_kinmuCd"></param>
    ''' <remarks>キーボード対応で使用</remarks>
    Private Sub pasteKeyBoardKinmu(ByVal p_kinmuCd As String)
        Const W_SUBNAME As String = "NSK0000HC  sprSheet_KeyDown"

        Dim w_RegStr As String
        Dim w_Var As Object
        Dim w_ActiveCol As Integer
        Dim w_ActiveRow As Integer
        Dim w_RiyuKBN As String '理由区分
        Dim w_Time As String '時間年休
        Dim w_Flg As String '確定ﾌﾗｸﾞ
        Dim w_ForeColor As Integer '文字色
        Dim w_BackColor As Integer '背景色
        Dim w_InputFlg As Boolean '入力ﾌﾗｸﾞ
        Dim w_Cnt As Short
        Dim w_KinmuPlanCD As String 'KinmuCD(予定ﾃﾞｰﾀ)
        Dim w_STS As Short
        Dim w_KangoCD As String
        Dim w_KangoPlanCD As String
        Dim w_RiyuPlanKbn As String
        Dim w_YYYYMMDD As Integer
        Dim w_DaikyuInputFlg As Boolean
        Dim w_strMsg() As String
        Dim w_lngBackColor As Integer
        Dim w_IntCol As Short
        Dim w_KibouCnt As Short
        Dim w_KibouCol() As Integer
        Dim w_blnColChk As Boolean
        Dim w_Comment As String = String.Empty  '希望勤務時のコメント 2015/04/13 Bando Add

        Try
            'ﾌｫｰｶｽ移動
            sprSheet.Focus()

            'ﾚｼﾞｽﾄﾘ格納先
            w_RegStr = General.G_SYSTEM_WIN7 & "\" & General.G_STRMAINKEY1 & "\" & "Current"

            '１件でも存在すれば・・・
            If m_TokusyuCnt <> 0 Then

                w_ActiveCol = sprSheet.Sheets(0).ActiveColumn.Index
                w_ActiveRow = sprSheet.Sheets(0).ActiveRow.Index

                '入力場所ﾁｪｯｸ
                If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row_ChgJisseki Then
                        Exit Sub
                    End If
                Else
                    If w_ActiveCol < M_KinmuData_Col Or w_ActiveRow < M_KinmuData_Row Then
                        Exit Sub
                    End If
                End If

                '部署異動ﾁｪｯｸ（配属範囲）
                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_ActiveCol).BackColor) = m_MonthBefore_Back Then
                    '配属期間外(背景色がグレー)の場合入力不可
                    Exit Sub
                End If

                '入力ﾌﾗｸﾞ
                w_InputFlg = True
                w_Cnt = CShort(w_ActiveCol - M_KinmuData_Col + 1)
                If g_SaikeiFlg = True Then
                    '再掲部署の場合
                    If m_DataFlg(w_Cnt) = "1" Then
                        '実績ﾃﾞｰﾀの場合
                        ReDim w_strMsg(2)
                        w_strMsg(1) = "確定済み勤務"
                        w_strMsg(2) = "再掲勤務"
                        Call General.paMsgDsp("NS0011", w_strMsg)
                        w_InputFlg = False
                    End If
                End If


                '勤務の曜日制限チェック
                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                If Check_YoubiLimit(w_YYYYMMDD, p_kinmuCd) = False Then
                    Exit Sub
                End If

                '2013/01/07 Ishiga add start------------------------------------------
                '超勤データの有無チェック
                If m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                    If Check_OverKinmuData(w_YYYYMMDD) = False Then
                        Exit Sub
                    End If
                End If
                '2013/01/07 Ishiga add end--------------------------------------------

                '届出存在チェック
                If fncChkAppliData(w_YYYYMMDD) = False Then
                    Exit Sub
                End If

                w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                w_YYYYMMDD = w_Var
                If m_PackageFLG = 1 Or m_PackageFLG = 3 Then
                    '日当直のコンポーネントありの時
                    With General.g_objGetData
                        .p職員区分 = 0
                        .p職員番号 = M_StaffID
                        .pチェック基準日 = w_YYYYMMDD
                        .p処理区分 = 0
                        .pチェック勤務CD = p_kinmuCd

                        If .mChkKinmuDuty = False Then
                            '勤務変更不可
                            '*******ﾒｯｾｰｼﾞ***********************************
                            ReDim w_strMsg(1)
                            w_strMsg(1) = ""
                            Call General.paMsgDsp("NS0110", w_strMsg)
                            '************************************************
                            Exit Sub
                        End If
                    End With
                End If

                w_DaikyuInputFlg = True

                If General.g_lngDaikyuMng = 0 Then
                    '代休ﾁｪｯｸ
                    w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
                    w_YYYYMMDD = w_Var
                    If Check_Daikyu(M_PASTE, w_YYYYMMDD, p_kinmuCd) = False Then
                        w_DaikyuInputFlg = False
                        w_InputFlg = False
                    End If
                End If

                If w_InputFlg = True Then
                    '入力可能な場合

                    With sprSheet.Sheets(0)

                        '警告表示+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        '変更する勤務の値が空で計画変更の場合、一つ上の勤務を取得する。
                        w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                        If w_Var = "" And m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            w_Var = .GetText(w_ActiveRow - 1, w_ActiveCol)
                        End If

                        '2015/04/13 Bando Upd Start =========================================
                        'Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        Call Get_KinmuMark(.GetText(w_ActiveRow, w_ActiveCol), w_KinmuPlanCD, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, "")
                        '2015/04/13 Bando Upd End   =========================================
                        If g_SaikeiFlg = False Then
                            Select Case w_RiyuKBN
                                Case "2"
                                    If frmNSK0000HA._mnuTool_5.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "要請勤務"
                                        w_strMsg(2) = "編集"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "3"
                                    If frmNSK0000HA._mnuTool_4.Checked = True Then
                                        ReDim w_strMsg(2)
                                        w_strMsg(1) = "希望勤務"
                                        w_strMsg(2) = "編集"
                                        w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                        If w_STS = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Case "4"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "再掲勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "5"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "委員会勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Case "6"
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "応援勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                            End Select
                        Else
                            If w_RiyuKBN = "2" Then
                                If frmNSK0000HA._mnuTool_5.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "要請勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If

                            If w_RiyuKBN = "3" Then
                                If frmNSK0000HA._mnuTool_4.Checked = True Then
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "希望勤務"
                                    w_strMsg(2) = "編集"
                                    w_STS = General.paMsgDsp("NS0097", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                        w_Time = ""
                        w_Flg = "0"
                        w_KangoCD = ""

                        '希望回数集計チェック
                        w_lngBackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                        If CDbl(g_HopeNumFlg) = 1 And _OptRiyu_2.Checked = True Then
                            w_KibouCnt = 0
                            ReDim w_KibouCol(0)
                            For w_IntCol = m_KeikakuD_StartCol To m_KeikakuD_EndCol
                                If ColorTranslator.ToOle(sprSheet.Sheets(0).Cells(w_ActiveRow, w_IntCol).BackColor) = w_lngBackColor Then
                                    w_KibouCnt = w_KibouCnt + 1
                                    ReDim Preserve w_KibouCol(w_KibouCnt)
                                    w_KibouCol(w_KibouCnt) = w_IntCol
                                End If
                            Next w_IntCol

                            '同じ場所に同じ勤務を貼り付けた場合、スルー
                            w_blnColChk = False
                            For w_IntCol = 1 To UBound(w_KibouCol)
                                If w_ActiveCol = w_KibouCol(w_IntCol) And p_kinmuCd = w_KinmuPlanCD Then
                                    w_blnColChk = True
                                End If
                            Next w_IntCol

                            If g_HopeNum <= w_KibouCnt And w_blnColChk = False Then
                                '希望回数制限オーバーダイアログ表示
                                If g_KibouNumDiaLogFlg = 1 Then
                                    'ワーニング
                                    ReDim w_strMsg(2)
                                    w_strMsg(1) = "希望勤務"
                                    w_strMsg(2) = "設定された希望勤務回数"
                                    '「&1が&2を超えています。~nこのまま登録してもよろしいですか。」
                                    w_STS = General.paMsgDsp("NS0150", w_strMsg)
                                    If w_STS = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    'エラー
                                    ReDim w_strMsg(1)
                                    w_strMsg(1) = "設定された希望勤務回数を超えているため"
                                    '「&1入力できません。」
                                    w_STS = General.paMsgDsp("NS0099", w_strMsg)
                                    Exit Sub
                                End If
                            End If
                        End If

                        Select Case True
                            Case _OptRiyu_0.Checked '通常
                                '理由区分 通常
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                            Case _OptRiyu_1.Checked '要請
                                '理由区分 要請
                                w_RiyuKBN = "2"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Yousei_Back", General.G_PALEGREEN))
                            Case _OptRiyu_2.Checked '希望
                                '理由区分 希望
                                w_RiyuKBN = "3"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Kibou_Back", General.G_PLUM))
                            Case _OptRiyu_3.Checked
                                '理由区分 再掲
                                w_RiyuKBN = "4"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Saikei_Back", General.G_LIGHTCYAN))
                            Case _OptRiyu_4.Checked
                                w_RiyuKBN = "6"
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Ouen_Back", General.G_ORANGE))
                            Case Else
                                '理由区分 その他（通常扱いとする）
                                w_RiyuKBN = "1"
                                w_ForeColor = ColorTranslator.ToOle(Color.Black)
                                w_BackColor = ColorTranslator.ToOle(Color.White)
                        End Select

                        '区分が応援の場合のみ、応援先勤務地選択画面を表示
                        If w_RiyuKBN = "6" Then
                            If Disp_Ouen(w_KangoCD) = False Then
                                Exit Sub
                            End If
                        End If

                        '2015/04/13 Bando Add Start =============================
                        '希望or要請の場合コメント入力画面表示
                        If w_RiyuKBN = "3" And g_InputHopeCommentFlg = "1" Then
                            If Disp_Comment(w_Comment, w_RiyuKBN) = False Then
                                Exit Sub
                            End If
                        End If
                        '2015/04/13 Bando Add End   =============================

                        '2015/07/22 Bando Add Start ========================
                        If w_Comment <> "" Then
                            w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Fore", Convert.ToString(General.G_BLACK)))
                            w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "HopeCom_Back", ColorTranslator.ToOle(Color.IndianRed)))
                        End If
                        '2015/07/22 Bando Add End   ========================

                        '代休発生勤務ﾊﾞｯｸｶﾗｰ
                        If General.g_lngDaikyuMng = 0 Then
                            If m_DaikyuBackColorFlg = True Then
                                w_ForeColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Fore", General.G_BLACK))
                                w_BackColor = Integer.Parse(General.paGetSetting(w_RegStr, "Color", "Daikyu_Back", General.G_LIGHTYELLOW))
                            End If
                        End If

                        'ｾﾙに記号を設定する
                        '2015/04/13 Bando Upd Start =========================
                        'w_Var = Set_KinmuMark(p_kinmuCd, w_RiyuKBN, w_Flg, w_KangoCD, w_Time)
                        w_Var = Set_KinmuMark(p_kinmuCd, w_RiyuKBN, w_Flg, w_KangoCD, w_Time, w_Comment)
                        '2015/04/13 Bando Upd End =========================
                        .SetText(w_ActiveRow, w_ActiveCol, w_Var)
                        '理由別の色設定
                        .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                        .Cells(w_ActiveRow, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        '勤務変更の時
                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            '2015/04/13 Bando Upd Start =========================================
                            'Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time)
                            Call Get_KinmuMark(.GetText(M_KinmuData_Row, w_ActiveCol), w_KinmuPlanCD, w_RiyuPlanKbn, w_Flg, w_KangoPlanCD, w_Time, "")
                            '2015/04/13 Bando Upd End   =========================================
                            '予定と異なる場合色変更
                            If Trim(p_kinmuCd) <> Trim(w_KinmuPlanCD) Or w_RiyuKBN <> w_RiyuPlanKbn Or w_KangoCD <> w_KangoPlanCD Then
                                .Cells(w_ActiveRow, w_ActiveCol).ForeColor = ColorTranslator.FromOle(m_Comp_Fore)
                            End If

                            '変更可能行の内容を格納
                            w_Var = .GetText(w_ActiveRow, w_ActiveCol)
                            w_ForeColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).ForeColor)
                            w_BackColor = ColorTranslator.ToOle(.Cells(w_ActiveRow, w_ActiveCol).BackColor)
                            '変更可能行の内容を勤務変更画面に貼り付け行へコピー
                            .SetText(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol, w_Var)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).ForeColor = ColorTranslator.FromOle(w_ForeColor)
                            .Cells(M_KinmuData_Row_ChgJisseki + 1, w_ActiveCol).BackColor = ColorTranslator.FromOle(w_BackColor)
                        End If
                    End With
                End If
            End If

            w_Var = sprSheet.Sheets(0).GetText(M_YYYYMMDDLabel_Row, w_ActiveCol)
            If InStr(m_strUpdKojyoDate, w_Var) = 0 Then
                m_strUpdKojyoDate = m_strUpdKojyoDate & w_Var & ","
            End If

            'ｶｰｿﾙｾﾙ移動
            If w_DaikyuInputFlg = True Then
                Call SetCursol()
            End If

            '更新ﾌﾗｸﾞｾｯﾄ
            If w_InputFlg = True Then
                m_KosinFlg = True
            End If
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    ''' <summary>
    ''' 背景色チェック
    ''' </summary>
    ''' <param name="p_spr"></param>
    ''' <param name="p_row"></param>
    ''' <param name="p_col"></param>
    ''' <returns></returns>
    ''' <remarks>通常カラーかどうか判定</remarks>
    Private Function IsExistBackColor(ByVal p_spr As FarPoint.Win.Spread.FpSpread, ByVal p_row As Integer, ByVal p_col As Integer) As Boolean
        Const W_SUBNAME As String = "NSK0000HC  IsExistBackColor"

        Dim rtnFlg As Boolean = True
        Dim frCl_Normal As Integer
        Dim bkCl_Normal As Integer

        Try
            '通常カラー
            frCl_Normal = ColorTranslator.ToOle(Color.Black)
            bkCl_Normal = ColorTranslator.ToOle(Color.White)

            With p_spr.Sheets(0)
                If ColorTranslator.ToOle(.Cells(p_row, p_col).BackColor) = bkCl_Normal _
                        OrElse .Cells(p_row, p_col).BackColor.A = 0 _
                        OrElse ColorTranslator.ToOle(.Cells(p_row, p_col).BackColor) = m_WeekEnd_Back Then
                    rtnFlg = False
                End If
            End With

            Return rtnFlg
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Function

    ''' <summary>
    ''' 夜勤専従・短時間情報表示
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DispNightShortInfo()
        Const W_SUBNAME As String = "NSK0000HC  DispNightShortInfo"

        Dim cnt As Integer
        Dim bigDate As Integer
        Dim smallDate As Integer

        Try
            '初期化（非表示）
            lblText1.Visible = False
            lblTermS1.Visible = False
            lblTermS2.Visible = False
            lblText2.Visible = False
            lblTermN1.Visible = False
            lblTermN2.Visible = False
            Panel1.Visible = False
            Panel2.Visible = False

            '短時間チェック
            cnt = 0
            For i As Integer = 1 To UBound(m_shortWorkInfo)
                '期間内か判定
                If m_shortWorkInfo(i).Date_St <= m_EndDate AndAlso m_StartDate <= m_shortWorkInfo(i).Date_Ed Then
                    cnt += 1
                    '開始日
                    smallDate = Integer.Parse(General.paGetDateStringFromInteger(m_shortWorkInfo(i).Date_St, General.G_DATE_ENUM.dd))

                    '終了日
                    bigDate = Integer.Parse(General.paGetDateStringFromInteger(m_shortWorkInfo(i).Date_Ed, General.G_DATE_ENUM.dd))

                    If cnt = 1 Then
                        lblTermS1.Text = General.paFormatSpace(smallDate, 2) & "日～" & General.paFormatSpace(bigDate, 2) & "日"
                        lblTermS1.Visible = True
                    Else
                        lblTermS2.Text = General.paFormatSpace(smallDate, 2) & "日～" & General.paFormatSpace(bigDate, 2) & "日"
                        lblTermS2.Visible = True
                    End If

                    '2件以上は処理抜け
                    If cnt >= 2 Then Exit For
                End If
            Next
            '1件以上あれば
            If cnt >= 1 Then Panel1.Visible = True

            '夜勤専従
            cnt = 0
            For i As Integer = 1 To UBound(m_nightWorkInfo)
                '期間内か判定
                If m_nightWorkInfo(i).Date_St <= m_EndDate AndAlso m_StartDate <= m_nightWorkInfo(i).Date_Ed Then
                    cnt += 1
                    '開始日
                    smallDate = Integer.Parse(General.paGetDateStringFromInteger(m_nightWorkInfo(i).Date_St, General.G_DATE_ENUM.dd))
                    '終了日
                    bigDate = Integer.Parse(General.paGetDateStringFromInteger(m_nightWorkInfo(i).Date_Ed, General.G_DATE_ENUM.dd))

                    If cnt = 1 Then
                        lblTermN1.Text = General.paFormatSpace(smallDate, 2) & "日～" & General.paFormatSpace(bigDate, 2) & "日"
                        lblTermN1.Visible = True
                    Else
                        lblTermN2.Text = General.paFormatSpace(smallDate, 2) & "日～" & General.paFormatSpace(bigDate, 2) & "日"
                        lblTermN2.Visible = True
                    End If

                    '2件以上は処理抜け
                    If cnt >= 2 Then Exit For
                End If
            Next
            '1件以上あれば
            If cnt >= 1 Then lblText2.Visible = True

            If Not Panel1.Visible AndAlso Panel2.Visible Then
                '短時間非表示で夜勤専従があれば詰める
                Panel2.Top = Panel2.Top - Panel1.Height
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub

    Private Sub sprSheet_Paint(ByVal sender As Object, ByVal e As PaintEventArgs) Handles sprSheet.Paint
        Dim colLoop As Integer
        Dim weekDayStr As String

        Try
            With sprSheet.Sheets(0)
                For colLoop = M_KinmuData_Col To m_KeikakuD_EndCol
                    weekDayStr = .GetText(1, colLoop)

                    If (m_WeekEndColorFlg = "1" AndAlso (weekDayStr = "土" OrElse weekDayStr = "日")) OrElse _
                            (m_HolidayColorFlg = "1" AndAlso (weekDayStr = "祝" OrElse weekDayStr = "休")) Then

                        If .Cells(M_KinmuData_Row, colLoop).BackColor.A = 0 Then
                            .Cells(M_KinmuData_Row, colLoop).BackColor = ColorTranslator.FromOle(m_WeekEnd_Back)
                        ElseIf .Cells(M_KinmuData_Row, colLoop).BackColor = Color.White Then
                            .Cells(M_KinmuData_Row, colLoop).BackColor = ColorTranslator.FromOle(m_WeekEnd_Back)
                        End If

                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            If .Cells(M_KinmuData_Row_ChgJisseki, colLoop).BackColor.A = 0 Then
                                .Cells(M_KinmuData_Row_ChgJisseki, colLoop).BackColor = ColorTranslator.FromOle(m_WeekEnd_Back)
                            ElseIf .Cells(M_KinmuData_Row_ChgJisseki, colLoop).BackColor = Color.White Then
                                .Cells(M_KinmuData_Row_ChgJisseki, colLoop).BackColor = ColorTranslator.FromOle(m_WeekEnd_Back)
                            End If
                        End If

                    Else

                        If .Cells(M_KinmuData_Row, colLoop).BackColor = ColorTranslator.FromOle(m_WeekEnd_Back) Then
                            .Cells(M_KinmuData_Row, colLoop).BackColor = Color.White
                        End If

                        If m_Mode = General.G_PGMSTARTFLG_CHANGEJISSEKI Or m_Mode = General.G_PGMSTARTFLG_CHANGEPLAN Then
                            If .Cells(M_KinmuData_Row_ChgJisseki, colLoop).BackColor = ColorTranslator.FromOle(m_WeekEnd_Back) Then
                                .Cells(M_KinmuData_Row_ChgJisseki, colLoop).BackColor = Color.White
                            End If
                        End If

                    End If
                Next
            End With
        Catch ex As Exception
            Call General.paTrpMsg(Str(Err.Number), "")
            End
        End Try
    End Sub

    '2014/04/23 Saijo add start P-06979--------------------------------------------------------------------------------------------------
    '/----------------------------------------------------------------------/
    '/  概要　　　　  : 勤務記号全角２文字対応のレイアウト変更
    '/  パラメータ    : なし
    '/  戻り値        : なし
    '/----------------------------------------------------------------------/
    Private Sub SetKinmuSecondView()
        Dim w_PreErrorProc As String
        w_PreErrorProc = General.g_ErrorProc 'ﾌﾟﾛｼｰｼﾞｬ名の一時待避
        General.g_ErrorProc = "NSK0000HC SetKinmuSecondView"

        Const W_FRAME_FIRST_HEIGHT As Integer = 20 '1行目の縦位置
        Const W_FRAME_ADD_HEIGHT As Integer = 27 '行の縦位置増え幅

        Const W_FRAME_FIRST_WIDTH As Integer = 8 '1列目の横位置
        Const W_FRAME_ADD_WIDTH As Integer = 39 '列の横位置増え幅

        Const W_FRAME_HEIGHT As Integer = 95 'フレームの縦幅
        Const W_FRAME_WIDTH As Integer = 990 'フレームの横幅
        Const W_SCL_WIDTH As Integer = 976 'スクロールの横幅
        Const W_SCL_HEIGHT As Integer = 16 'スクロールの縦幅
        Const W_KINMU_WIDTH As Integer = 40 '勤務の横幅
        Const W_KINMU_HEIGHT As Integer = 25 '勤務の縦幅

        Try
            '勤務記号全角２文字対応フラグ判定
            If m_strKinmuEmSecondFlg = "0" Then
                '0：対応しない(従来の勤務記号入力サイズと最大2バイト)
            Else
                '1：対応する(全角２文字が表示できる勤務記号入力サイズと最大4バイト)
                'パレットを載せているフレーム
                FramAll.Size = New System.Drawing.Size(1200, 400)

                'フレーム
                _Frame_0.Size = New System.Drawing.Size(W_FRAME_WIDTH, W_FRAME_HEIGHT)
                _Frame_1.Size = New System.Drawing.Size(W_FRAME_WIDTH, W_FRAME_HEIGHT)
                _Frame_2.Size = New System.Drawing.Size(W_FRAME_WIDTH, W_FRAME_HEIGHT)
                _Frame_3.Size = New System.Drawing.Size(W_FRAME_WIDTH, W_FRAME_HEIGHT)

                'スクロール
                HscKinmu.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)
                HscYasumi.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)
                HscTokusyu.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)
                HscSet.Size = New System.Drawing.Size(W_SCL_WIDTH, W_SCL_HEIGHT)

                '勤務
                General.setSizeAndLocal(m_lstCmdKinmu, 2, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

                '勤務(休み)
                General.setSizeAndLocal(m_lstCmdYasumi, 2, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

                '勤務(特殊勤務)
                General.setSizeAndLocal(m_lstCmdTokusyu, 2, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

                '勤務(セット)
                General.setSizeAndLocal(m_lstCmdSet, 1, W_KINMU_WIDTH, W_KINMU_HEIGHT, W_FRAME_FIRST_WIDTH, W_FRAME_FIRST_HEIGHT, W_FRAME_ADD_WIDTH, W_FRAME_ADD_HEIGHT)

            End If

            General.g_ErrorProc = w_PreErrorProc '待避ﾌﾟﾛｼｰｼﾞｬ名を元に戻す

        Catch ex As Exception
            Err.Raise(Err.Number)
        End Try
    End Sub
    '2014/04/23 Saijo add end P-06979----------------------------------------------------------------------------------------------------
End Class