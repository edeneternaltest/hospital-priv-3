Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Module BasNSK0000H
    '/----------------------------------------------------------------------/
    '/
    '/    ｼｽﾃﾑ名称：看護支援システム(勤務管理)
    '/ ﾌﾟﾛｸﾞﾗﾑ名称：計画作成メニュー
    '/        ＩＤ：NSK0000H
    '/        概要：対象月の計画を立案する
    '/
    '/
    '/      作成者： S.Y    CREATE 2000/07/24           REV 01.00
    '/      更新者： M.N           2008/11/25           【P-00859】
    '/      更新者： M.I           2008/12/08           【P-00931】
    '/      更新者： M.I           2008/12/09           【P-00947】
    '/      更新者： M.I           2008/12/09           【P-00958】
    '/      更新者： M.I           2008/12/09           【PRE-0314】
    '/      更新者： M.I           2008/12/16           【P-01000】
    '/      更新者： M.I           2008/12/17           【P-01006】
    '/      更新者： M.I           2008/12/24           【P-01053】
    '/      更新者： M.I           2008/12/25           【P-01082】
    '/      更新者： M.I           2009/01/08           【P-01127】
    '/      更新者： M.I           2009/01/08           【P-01132】
    '/      更新者： M.I           2009/01/15           【P-01172】
    '/      更新者： M.I           2009/02/09           【P-01424】
    '/      更新者： M.I           2009/06/10           【PKG-0215】
    '/      更新者： M.I           2009/06/10           【PKG-0129】
    '/      更新者： M.I           2009/06/15           【PKG-0089】
    '/      更新者： M.I           2009/06/16           【PRE-0683】
    '/      更新者： T.I           2009/06/18           【PRE-0706】
    '/      更新者： M.I           2009/06/19           【PRE-0709】
    '/      更新者： M.I           2009/06/19           【PRE-0713】
    '/      更新者： M.I           2009/06/19           【PRE-0721】
    '/      更新者： M.I           2009/06/19           【PRE-0726】
    '/      更新者： M.I           2009/06/25           【PRE-0772】
    '/      更新者： M.I           2009/07/07           【P-01967】
    '/      更新者： M.I           2009/07/07           【PRE-0906】
    '/      更新者： M.I           2009/07/13           【PRE-0914】
    '/      更新者： M.I           2009/07/15           【P-02030】
    '/      更新者： okamura       2009/07/17           【P-01981】
    '/      更新者： okamoto       2009/07/23           【P-02050】
    '/      更新者： okamura       2009/08/07           【PRE-1013】
    '/      更新者： okamura       2009/09/02           【P-02215】
    '/      更新者： M.I           2009/11/12           【P-02390】
    '/      更新者： Y.I           2012/10/24           【P-*****】（PKGバージョンUP_7.0）
    '/      更新者： T.Ishiga      2013/01/07           【P-05697】
    '/      更新者： Y.Bando       2015/04/10           【P-07830】 (PKG7.5)希望勤務時コメント入力機能追加
    '/      更新者： Angelo        2017/08/24           【PKGバージョンアップ】
    '/     Copyright (C) Inter co.,ltd 2000
    '/----------------------------------------------------------------------/

    '--------------------------------------------------------------------------------
    '       NSK0000H 定数 宣言
    '--------------------------------------------------------------------------------
    'アイコン
    Public Const G_FORM_ICO As String = "kinmu.ico" 'フォームアイコン
	Public Const G_ERASER_ICO As String = "Eraser.ico" '消しゴム(NSK0000HB)
	Public Const G_CLOSE_ICO As String = "Close.ico" '閉じる(NSK0000HB)
	Public Const G_SEARCH_ICO As String = "Search.ico" '検索(NSK0000HD)
	Public Const G_PERMUTATION_ICO As String = "Permutation.ico" '置換(NSK0000HD)
	
	'文字列
	Public Const G_LOAD_STR As String = "一時ファイル読み込み中..." '(1002)
	Public Const G_SAVE_STR As String = "一時ファイル保存中..." '(1005)
	Public Const G_KEIKAKUSAVE_STR As String = "計画データ 保存中..." '(1010)
	Public Const G_PICKUP_STR As String = "データ抽出中…" '(1011)
	Public Const G_SORT_STR As String = "データ並び替え中..." '(1012)
	Public Const G_DELTE_STR As String = "消去中..." '(1014)
	Public Const G_CUT_STR As String = "切り取り中..." '(1015)
    Public Const G_ROLLBACK_STR As String = "元に戻し中..." '(1016)
    Public Const G_REDO_STR As String = "やり直し中..." '2016/04/05 Ishiga add
	Public Const G_PASTE_STR As String = "勤務記号貼り付け中..." '(1017)
	Public Const G_COPY_STR As String = "コピー中..." '(1018)
	Public Const G_JISSEKISAVE_STR As String = "実績データ 更新中..." '(1020)
	Public Const G_SEARCH_STR As String = "検索中..." '(1023)
	Public Const G_TEAM_STR As String = "チーム" '(5008)
    Public Const G_KINMUDEPT_STR As String = "勤務部署" '(5009)

    '2017/08/24 Angelo add st---------------------------------
    '表示内容の編集モード
    Public Const G_EDITMODE_NO As Short = 0
    Public Const G_EDITMODE_DATETIME As Short = 1
    '2017/08/24 Angelo add en---------------------------------

    '2017/09/08 Angelo add st----------------------------------------------
    'ﾃﾞｰﾀﾀｲﾌﾟ
    Public Const M_PLANDATA As String = "0" '計画ﾃﾞｰﾀ
    Public Const M_JISSEKIDATA As String = "1" '実績ﾃﾞｰﾀ
    '2017/09/08 Angelo add en----------------------------------------------

    '--------------------------------------------------------------------------------
    '       NSK0000H 変数 宣言
    '--------------------------------------------------------------------------------
    Public g_AppName As String 'ﾌﾟﾛｸﾞﾗﾑID 格納
    Public g_KinmuM() As KinmuM_Type '勤務情報配列
    Public g_HolidayBunruiM() As HolidayM_Type '休み分類情報配列
	Public w_KinmuCDCount As Short
	Public g_SaikeiFlg As Boolean 'True:再掲部署,False:再掲部署以外
	Public g_LimitedFlg As Boolean '(True：勤務者  False：管理者)
	Public g_ImagePath As String 'イメージパス
	
    Public g_HopeNum As Short '希望回数
    '2014/05/14 Shimpo add start P-06991-----------------------------------------------------------------------
    Public g_HopeNumDate As Short '希望回数(日付別)
    '2014/05/14 Shimpo add end P-06991-------------------------------------------------------------------------
    Public g_HopeNumFlg As String '希望回数制限フラグ(1:制限する　2:制限しない)
    '2014/05/22 Shimpo add start P-06991-----------------------------------------------------------------------
    Public g_HopeNumDateFlg As String '希望回数(日付別)制限フラグ(1:制限する　2:制限しない)
    '2014/05/22 Shimpo add end P-06991-------------------------------------------------------------------------
    Public g_KibouNumDiaLogFlg As Integer '希望回数制限ダイアログ（1:ワーニング　以外:エラー）
    Public g_HopeMode As String = "0"        '（1:希望モード、0:それ以外）
    Public g_objKeyBoard As Dictionary(Of Integer, String)
    '2015/04/13 Bando Add Start ===================
    Public g_InputHopeCommentFlg As String '希望コメントフラグ(1：コメント可 2:コメント不可)
    '2015/04/13 Bando Add Start ===================

    '2017/09/04 Angelo Add st-----------------------------------------------------------------------------------------------------------------------
    '勤務地異動情報
    Public Structure IdoData_Type
        Dim CD As String '追加仕様：応援者・月中異動者の総夜勤を全部出す
        Dim IdoYMD As Integer '異動年月日
        Dim IdoYMD2 As Integer '異動年月日２
        Dim SyuryoYMD As Integer '終了年月日
    End Structure

    '勤務情報
    Public Structure KinmuData_Type
        Dim KinmuCD As String 'KinmuCD
        Dim Date_Renamed As Integer '日付
        Dim RiyuKBN As String '理由区分
        Dim Time As String '時間年休（最大４件まで）
        Dim DataFlg As String '計画ﾃﾞｰﾀ,実績ﾃﾞｰﾀ判別用ﾌﾗｸﾞ(0:計画ﾃﾞｰﾀ,1:実績ﾃﾞｰﾀ)
        Dim KakuteiFlg As String '確定判別用ﾌﾗｸﾞ(0:該当部署,1:他部署)
        Dim DataChk As String 'DBの存在の有無("1":DBﾃﾞｰﾀあり，"":DBﾃﾞｰﾀなし）
        Dim OuenKangoCD As String '応援先看護単位CD
        Dim FirstRegistTimeDate As Double
        Dim LastUpdTimeDate As Double
        Dim RegistID As String
        Dim Comment As String '希望コメント　2015/04/10 Bando Add
    End Structure

    '年休詳細情報
    Public Structure NenkyuDetail_Type
        Dim GetContentsKbn As String '取得内容区分(1:全日,2:前半,3:後半,4:時間年休)
        Dim HolidayBunruiCD As String '休み分類CD
        Dim FromTime As Integer '開始時間
        Dim ToTime As Integer '終了時間
        Dim DateKbn As String '年月日区分(0:当日,1:翌日)
        Dim NenkyuTime As Integer '時間年休
        Dim HolSubFlg As String '休憩減算フラグ
        Dim DayTime As Integer '日勤時間
        Dim NightTime As Integer '夜勤時間
        Dim NextNightTime As Integer '翌日夜勤時間
    End Structure

    '年休情報
    Public Structure NenkyuData_Type
        Dim Date_Renamed As Integer '日付
        '*=*=*=*=*=*=*=*=*=*=*=*=排他処理対応の為の一時保存スペース*=*=*=*=*=*=*=*=*=*=*=*=*
        'この構造体は更新処理(「Save_PlanData」「Save_PlanData_4W」「Save_JissekiData」)の時
        'に値を入れ「UpDate_JikanNenkyu」でＤＢ上(NS_NENKYU_F)に保存しにいく為に使用してます
        Dim Detail() As NenkyuDetail_Type
        '*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
        Dim DataFlg As Boolean 'データ更新FLG(True:更新する，False:更新しない)
        Dim DataChk As Boolean 'DBの存在の有無(True:DBﾃﾞｰﾀあり，False:DBﾃﾞｰﾀなし）
        Dim FirstRegistTimeDate As Double
        Dim LastUpdTimeDate As Double
        Dim RegistID As String
    End Structure

    '個人別勤務条件
    Public Structure PersonalCondition_Type
        Dim NotKinmu As Boolean 'True:割当不可，False:割当可
        Dim CountMax As Integer
        Dim CountMin As Integer
        Dim IntervalMax As Integer
        Dim IntervalMin As Integer
        Dim MondayNot As Boolean 'True:割当不可，False:割当可
        Dim TuesdayNot As Boolean 'True:割当不可，False:割当可
        Dim WednesdayNot As Boolean 'True:割当不可，False:割当可
        Dim ThursdayNot As Boolean 'True:割当不可，False:割当可
        Dim FridayNot As Boolean 'True:割当不可，False:割当可
        Dim SaturdayNot As Boolean 'True:割当不可，False:割当可
        Dim SundayNot As Boolean 'True:割当不可，False:割当可
        '2015/05/13 Ishiga add start---------------------
        Dim RenzokuCountMax As Integer '連続勤務上限回数
        Dim RenkyuCountMax As Integer '連休上限回数
        '2015/05/13 Ishiga add end-----------------------
    End Structure

    '看護加算計算用データ構造体
    Public Structure Kangokasan_Type
        Dim PostCD As String
        Dim JobCD As String
        Dim SaiyoCD As String
        Dim YakinKBN As String
        Dim ChildShort As Boolean '短時間
    End Structure

    '代休詳細用構造体
    'Private Structure DaikyuDetail_Type
    <Serializable()> Public Structure DaikyuDetail_Type '2016/04/06 Yamanishi Upd
        Dim DaikyuDate As Integer '代休消化日
        Dim DaikyuKinmuCD As String '代休消化勤務ＣＤ
        Dim GetFlg As String '取得タイプ(0:1日、1:0.5日)
    End Structure

    '代休用構造体
    'Private Structure Daikyu_Type
    <Serializable()> Public Structure Daikyu_Type '2016/04/06 Yamanishi Upd
        Dim HolDate As Integer
        Dim HolKinmuCD As String
        Dim DaikyuDetail() As DaikyuDetail_Type '代休詳細
        Dim GetKbn As String '代休発生量タイプ(0:1日分,1:1.5日分)
        Dim RemainderHol As Double '残り代休
        Dim OutPutList As String '代休取得時にリストにあげるかどうかのﾌﾗｸﾞ("0":代休取得日非対象 "1":代休取得日対象)
        Dim FirstRegistTimeDate As Double
        Dim LastUpdTimeDate As Double
        Dim RegistID As String
    End Structure

    '時間年休用構造体
    Public Structure JikanNenkyu_Type
        Dim Date_Renamed As Integer
        Dim StrData As String
    End Structure

    '採用履歴情報
    Public Structure SaiyoData_Type
        Dim SaiyoDate As Integer
        Dim TentaiDate As Integer
        Dim SaiyoCD As String
        Dim strStaffNo As String
        Dim lngHaizoku As Integer
        Dim blnHaizokuFLG As Boolean
        Dim strSecName As String
    End Structure

    'Private Structure KojyoData_Type
    <Serializable()> Public Structure KojyoData_Type '2016/04/06 Yamanishi Upd
        Dim lngDate As Integer '年月日
        Dim lngKinmuDetailDate() As Integer '勤務詳細年月日
        Dim strKinmuDetailCD() As String '勤務詳細CD
        Dim lngTimeFrom() As Integer '時間帯From
        Dim lngTimeTo() As Integer '時間帯To
        Dim lngNikkinTime() As Integer '日勤時間
        Dim lngYakinTime() As Integer '夜勤時間
        Dim lngYokuYakinTime() As Integer '翌日夜勤時間
        Dim strNextFlg() As String '翌週FLG
        Dim lngKinmuDetailTime() As Integer '控除時間
        Dim strHolSubFlg() As String '休憩減算フラグ
        Dim strShinryoKbn() As String '診療報酬計算区分
        Dim UniqueseqNo() As String 'ユニークNO
        Dim Seq() As Integer
        '2018/02/23 Yamanishi Upd Start ------------------------------------------
        'Sub init()
        '    ReDim lngKinmuDetailDate(0)
        '    ReDim strKinmuDetailCD(0)
        '    ReDim lngTimeFrom(0)
        '    ReDim lngTimeTo(0)
        '    ReDim lngNikkinTime(0)
        '    ReDim lngYakinTime(0)
        '    ReDim lngYokuYakinTime(0)
        '    ReDim strNextFlg(0)
        '    ReDim lngKinmuDetailTime(0)
        '    ReDim strHolSubFlg(0)
        '    ReDim strShinryoKbn(0)
        '    ReDim UniqueseqNo(0)
        '    ReDim Seq(0)
        'End Sub

        Dim OuenKinmuDeptCD() As String

        Sub init(Optional ByVal p_Cnt As Integer = 0)
            ReDim lngKinmuDetailDate(p_Cnt)
            ReDim strKinmuDetailCD(p_Cnt)
            ReDim lngTimeFrom(p_Cnt)
            ReDim lngTimeTo(p_Cnt)
            ReDim lngNikkinTime(p_Cnt)
            ReDim lngYakinTime(p_Cnt)
            ReDim lngYokuYakinTime(p_Cnt)
            ReDim strNextFlg(p_Cnt)
            ReDim lngKinmuDetailTime(p_Cnt)
            ReDim strHolSubFlg(p_Cnt)
            ReDim strShinryoKbn(p_Cnt)
            ReDim UniqueseqNo(p_Cnt)
            ReDim Seq(p_Cnt)
            ReDim OuenKinmuDeptCD(p_Cnt)
        End Sub
        '2018/02/23 Yamanishi Upd End --------------------------------------------
    End Structure

    Public Structure SumCntDeteil_Type
        Dim GetFlg As Boolean
        Dim Cnt As Double
    End Structure

    Public Structure SumCntData_Type
        Dim SumSeq() As SumCntDeteil_Type
    End Structure

    '対象職員情報
    Public Structure StaffData_Type
        Dim ID As String '職員管理番号
        Dim PreID As String '職員番号
        Dim StaffName As String '氏名
        Dim SaiyoYMD As Integer '採用年月日
        Dim TentaiYMD As Integer '転退年月日
        Dim IdoData() As IdoData_Type '異動情報（該当部署分のみ）
        Dim InitialHyojiNo As Integer '表示No(初期値)
        Dim HyojiNo As Integer '表示No
        Dim HyojiNo1 As Integer '表示No1
        Dim HyojiNo2 As Integer '表示No2
        Dim HyojiNo3 As Integer '表示No3
        Dim HyojiNo4 As Integer '表示No4
        Dim HyojiNo5 As Integer '表示No5
        Dim Team As Integer 'ﾁｰﾑ
        Dim AutoKBN As String '自動割当区分
        Dim JobHyojiNo As Integer '職種表示No
        Dim PostHyojiNo As Integer '役職表示No
        Dim GiryoHyojiNo As Integer '技量表示No
        Dim GiryoLvCD As String 'SkillLvlCD
        Dim GiryoBunruiCD As String 'SkillBunruiCD
        Dim KinmuData() As KinmuData_Type '勤務ﾃﾞｰﾀ
        Dim SaikeiData() As KinmuData_Type '再掲ﾃﾞｰﾀ（再掲部署のみ）
        Dim CompKinmuData() As KinmuData_Type '勤務ﾃﾞｰﾀ(予定ﾃﾞｰﾀ[実績ﾃﾞｰﾀとの比較用])
        Dim NenkyuData() As NenkyuData_Type '年休ﾃﾞｰﾀ
        Dim KinmuCondition() As PersonalCondition_Type '個人別勤務条件データ
        Dim PersonalError As Boolean 'True:エラー，False:エラーなし
        Dim JobCode As String '職種CD     (I/F)
        Dim PostCode As String '役職CD     (I/F)
        Dim PostName As String '役職名称     (I/F)　
        Dim HaizokuIF As Integer '配属日     (I/F)
        Dim TensyutuIF As Integer '転出日     (I/F)
        Dim Syokai As Double 'RegistFirstTimeDate
        Dim Saisyu As Double 'LastUpdTimeDate
        Dim UpdateFlg As String '職員歴Ｆ（ﾃﾞﾌｫﾙﾄ値:"0" Or 該当計画番号:"1" Or 前計画番号:"2"）
        Dim YakinKBN As String '夜勤専従者区分
        Dim PatternCD As String 'パターンコード
        Dim OuenStaffFlg As Integer '応援勤務者かどうかの判断（0：対象部署所属　1：応援勤務者）
        Dim TargetKikanFlg As Boolean '対象期間内（１ヶ月）の間に在職しているか（職員取得は１ヶ月で行っていないため、職員情報設定画面にデータを渡す際に使用） (True:在職, False:在職してない)
        Dim KangoKasanData() As Kangokasan_Type '看護加算計算用データ
        Dim Daikyu() As Daikyu_Type '代休データ
        Dim LoadDaikyu() As Daikyu_Type '画面ﾛｰﾄﾞ時の代休データ（勤務変更個人別画面時に使用）
        'Dim BackDaikyu() As Daikyu_Type '代休データ退避用 '2016/04/06 Yamanishi Del
        Dim WariateJikanNenkyu() As JikanNenkyu_Type '自動割り当て時の時間年休退避用配列
        Dim WariateOuenKinmu() As JikanNenkyu_Type '自動割り当て時の応援勤務退避用配列
        Dim WariateComment() As JikanNenkyu_Type '自動割り当て時の希望コメント退避用配列 2015/04/10 Bando Add
        Dim SaiyoData() As SaiyoData_Type '採用履歴情報
        Dim RuikeiTime As Integer '累計時間(該当計画番号の２つ前まで)
        Dim RuikeiTime_Jisseki As Integer '実績期間(該当計画番号の１つ前の時間)
        Dim Kojyo() As KojyoData_Type
        Dim blnEndDayChangeFlg As Boolean '強制的に異動暦のENDDATEを期間の最終日に変換したか(99999999や0を20080331へ)
        Dim NightWork() As NghtShrtData_Type '夜勤専従情報
        Dim ShortWork() As NghtShrtData_Type '短時間制度情報
        Dim SumCntData As SumCntData_Type
        Dim SumCntData_4W() As SumCntData_Type
        Dim ResultCnt() As Integer
        '追加仕様：応援者・月中異動者の総夜勤を全部出す---------------------------------------------------------------------
        Dim BeforeKinmuData() As KinmuData_Type
        Dim IdoHistory() As IdoData_Type
        'Dim BeforeJikanNenkyu() As JikanNenkyu_Type '2016/06/15 Yamanishi Del
        '追加仕様：応援者・月中異動者の総夜勤を全部出す---------------------------------------------------------------------
    End Structure
    '2017/09/04 Angelo Add en-----------------------------------------------------------------------------------------------------------------------

    '-- 勤務記号 退避配列 ----------------------
    Public Structure KinmuM_Type
		Dim CD As String 'KinmuCD
		Dim KinmuName As String '名称
		Dim Mark As String '記号
		Dim KBunruiCD As String '勤務分類CD
		Dim WBunruiCD As String 'AllocBunruiCD
        Dim WFlg As String '割当ﾌﾗｸﾞ
		Dim HFlg As String '半日勤務ﾌﾗｸﾞ
		Dim AMCD As String 'ＡＭ勤CD
		Dim PMCD As String 'ＰＭ勤CD
		Dim From As Short '勤務時間帯FROM
        Dim To_Renamed As Short '勤務時間帯TO
		Dim Time As Short '勤務時間
		Dim TimeFlg As String '時間休ﾌﾗｸﾞ
		Dim Setumei As String '説明
        Dim DaikyuFlg As String '代休取得ﾌﾗｸﾞ(1:可能 2:不可)
        Dim HolBunruiCD As String '休み分類CD
        Dim KinmuKBN As String '勤務区分(0:勤務 1:勤務以外)
        Dim YoubiLimit() As String '曜日制限
        Dim EfftoDate As String '有効終了日 2015/05/22 Bando Add
	End Structure
	
    '-- 休み分類記号 退避配列 --
	Public Structure HolidayM_Type
		Dim CD As String 'HolidayBunruiCD
		Dim HolidayName As String '名称
		Dim SecName As String '略称
		Dim Mark As String '記号
		Dim Setumei As String '説明
		Dim DivFlg As String '分割休暇取得ﾌﾗｸﾞ(0:不可 1:分割可能)
		Dim GetPosFrom As Short '取得可能日FROM
		Dim GetPosTo As Short '取得可能日TO
		Dim AppliCD As String '届出CD
	End Structure
	
	'ｴﾗｰﾁｪｯｸ用
	'*** 回数チェック用 ***
    Public Structure SpanCount_Type
        Dim KinmuCount As Single
        Dim ErrorFlg As Boolean 'True:ｴﾗｰあり，False:ｴﾗｰなし
    End Structure

	'*** 間隔チェック用 ***
    Public Structure IntervalErr_Type
        Dim ErrorDate As Integer
        Dim ErrorFlg As Boolean 'True:ｴﾗｰあり，False:ｴﾗｰなし
    End Structure

	'回数／間隔エラーチェック用
    Public Structure CountErr_Type
        Dim ErrName As String
        Dim ErrBunrui As String
        Dim CheckSpan() As SpanCount_Type '回数エラー（配列は計画期間 0:表示期間，n:各計画期間ごと）
        Dim InterValErr As IntervalErr_Type '間隔エラー
    End Structure

    Public g_KikanError() As CountErr_Type '配列は集計勤務数
    Public g_RenzokuError() As CountErr_Type
	
	'*** 禁止パターンチェック用 ***
    Public Structure NotPatternErr_Type
        Dim ErrorPattern As String
        Dim ErrorDate As Integer
        Dim ErrorFlg As Boolean 'True:ｴﾗｰあり，False:ｴﾗｰなし
    End Structure

    Public g_NotPatternError As NotPatternErr_Type

	'*** 否定勤務チェック用 ***
    Public Structure NotKinmuErr_Type
        Dim KinmuName As String
        Dim ErrorDate As Integer
        Dim ErrorFlg As Boolean 'True:ｴﾗｰあり，False:ｴﾗｰなし
    End Structure

    Public g_NotKinmuError As NotKinmuErr_Type

    '計画単位用列挙型（４週／１ヶ月）
	Public Enum gePlanType
		PlanType_Month '１ヶ月（ﾃﾞｰﾀ："1"）
		PlanType_Week '４週（ﾃﾞｰﾀ："2"）
    End Enum

	'表示期間用列挙型（４週／１ヶ月）
	Public Enum geViewType
		ViewType_Month '１ヶ月（ﾃﾞｰﾀ："1"）
		ViewType_Week '４週（ﾃﾞｰﾀ："2"）
    End Enum

	'Window位置･大きさ設定用列挙型
	Public Enum geWindowPosition
		GetSettingValue 'ｳｨﾝﾄﾞｳの位置、大きさ設定
		SaveSettingValue 'ｳｨﾝﾄﾞｳの位置、大きさ保存
    End Enum

    '検索/置換区別
    Public Enum geKenChiFlg
        FormType_Kensaku '検索画面
        FormType_Chikan '置換画面
    End Enum

    'ｴﾗｰﾁｪｯｸ用
    '*** 回数チェック用 ***
    Public Structure SpanCount2_Type
        Dim KinmuCount As Single
        Dim ErrorDate As Integer 'エラー開始日
        Dim ErrorFlg As Boolean 'True:ｴﾗｰあり，False:ｴﾗｰなし
        Dim ColIdx As Short '日付(列)インデックス
        Dim KinmuName As String 'エラー勤務
    End Structure

	'*** 間隔チェック用 ***
    Public Structure IntervalErr2_Type
        Dim ErrorDate As Integer
        Dim ErrorFlg As Boolean 'True:ｴﾗｰあり，False:ｴﾗｰなし
        Dim ColIdx As Short '日付(列)インデックス
        Dim ErrorName As String
    End Structure

	'回数／間隔エラーチェック用
    Public Structure CountErr2_Type
        Dim ErrName As String
        Dim ErrBunrui As String
        Dim StaffName As String '職員氏名
        Dim StaffIdx As Short '職員(行)インデックス
        Dim CheckSpan() As SpanCount2_Type '回数エラー（配列は計画期間 0:表示期間，n:各計画期間ごと）
        Dim InterValErr() As IntervalErr2_Type '間隔エラー
    End Structure

    Public g_KikanError2() As CountErr2_Type '配列は集計勤務数
    Public g_RenzokuError2() As CountErr2_Type

    '夜勤専従・短時間者情報
    <Serializable()> Public Structure NghtShrtData_Type
        Dim Date_from As Integer
        Dim Date_to As Integer
        Dim ReasonCd As String
        Dim ReasonRNm As String
    End Structure

    '行事参加者
    <Serializable()> Public Structure EventStaff_Type
        Dim staffMngId As String
        Dim staffNm As String
        Dim postNm As String
    End Structure

    '行事予定
    <Serializable()> Public Structure EventList_Type
        Dim DateF As Integer
        Dim Time_st As Short
        Dim Time_ed As Short
        Dim EventName As String
        Dim allFlg As Boolean
        Dim uniqNo As String
        Dim EventStaff() As EventStaff_Type
        Sub init()
            DateF = 0
            Time_st = 0
            Time_ed = 0
            EventName = ""
            allFlg = False
            uniqNo = ""
            ReDim EventStaff(0)
        End Sub
    End Structure

    '*** 禁止パターンチェック用 ***
    Public Structure NotPatternDetail_Type
        Dim ErrorPattern As String
        Dim ErrorDate As Integer
        Dim ErrorFlg As Boolean 'True:ｴﾗｰあり，False:ｴﾗｰなし
        Dim ColIdx As Short '日付(列)インデックス
        Dim EndDate As Integer
    End Structure

    Public Structure NotPatternErr2_Type
        Dim StaffName As String
        Dim StaffIdx As Short '職員(行)インデックス
        Dim Data() As NotPatternDetail_Type
    End Structure

    Public g_NotPatternError2() As NotPatternErr2_Type

    '*** 否定勤務チェック用 ***
    Public Structure NotKinmuDetail_Type
        Dim KinmuName As String
        Dim ErrorDate As Integer
        Dim ErrorFlg As Boolean 'True:ｴﾗｰあり，False:ｴﾗｰなし
        Dim ColIdx As Short '日付(列)インデックス
    End Structure

    Public Structure NotKinmuErr2_Type
        Dim StaffName As String
        Dim StaffIdx As Short '職員(行)インデックス
        Dim Data() As NotKinmuDetail_Type
    End Structure
    Public g_NotKinmuError2() As NotKinmuErr2_Type

    '*** 禁止職員パターンチェック用 ***
    Public g_NotStaffPatternError2() As NotPatternErr2_Type

    '経験区分の組み合わせチェック用
    Public g_NotGiryoCheckError() As NotPatternErr2_Type

    Public g_NotAbsKinmuCheckError() As NotPatternErr2_Type '必須勤務
    '2017/09/29 Yamanishi Add ---------------------------------------------------------------------------------------------------------
    Private m_dicRandomKey2TimeNenkyu As New Dictionary(Of String, String)
    ''' <summary>
    ''' ランダムな文字列を生成し、それをKeyとしてDictionaryにValueとなるデータを格納
    ''' </summary>
    ''' <param name="p_Len">生成する文字列の長さ</param>
    ''' <param name="p_Val">Valueとなるデータ</param>
    ''' <param name="p_Dic">Dictionary</param>
    ''' <returns>生成された文字列</returns>
    Public Function GenerateRandomKeyAndSetDictionary(ByVal p_Len As Integer, ByVal p_Val As String, ByRef p_Dic As Dictionary(Of String, String)) As String
        '使用する文字
        Const W_Chars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ@[]{};:*;-/.,?!#$%&()=^~|"
        Dim w_RetVal As New System.Text.StringBuilder(p_Len)
        Dim w_Random As New Random
        For i As Integer = 0 To p_Len - 1
            '選択された位置の文字を取得・追加
            w_RetVal.Append(W_Chars(w_Random.Next(W_Chars.Length)))
        Next i
        GenerateRandomKeyAndSetDictionary = w_RetVal.ToString
        If Not p_Dic.ContainsKey(GenerateRandomKeyAndSetDictionary) Then
            p_Dic.Add(GenerateRandomKeyAndSetDictionary, p_Val)
        Else
            '被ってたらやり直し
            Return GenerateRandomKeyAndSetDictionary(p_Len, p_Val, p_Dic)
        End If
    End Function
    '----------------------------------------------------------------------------------------------------------------------------------
    '*****************************************************************************************************
    '   ｽﾌﾟﾚｯﾄﾞの勤務記号文字列を各データ毎に分割する
    '   ﾊﾟﾗﾒｰﾀ：p_Val(I)（勤務ｺｰﾄﾞ）
    '           p_KinmuCD(O)（勤務ｺｰﾄﾞ）
    '           p_RiyuKBN(O)（理由区分）
    '           p_Time(O)（時間年休）
    '           p_Flg(0)（確定部署フラグ）
    '           p_KangoCD（応援先看護単位CD）
    '   編集仕様
    '       勤務記号(2)＋Space(5)＋勤務ｺｰﾄﾞ(3)＋理由区分(1)＋確定部署フラグ(1)+応援先看護単位CD(4) + 希望コメント(20) +時間年休(44)
    '       全６０バイト分
    '           1) 勤務記号---１ﾊﾞｲﾄ目から７ﾊﾞｲﾄ分（Trimで余分なｽﾍﾟｰｽを省く）
    '           2) KinmuCD---８ﾊﾞｲﾄ目から３ﾊﾞｲﾄ分（Trimで余分なｽﾍﾟｰｽを省く）
    '           3) 理由区分---11ﾊﾞｲﾄ目から１ﾊﾞｲﾄ分
    '           4) 確定部署FLG---12ﾊﾞｲﾄ目から１バイト分
    '             （"1"：他部署確定ﾃﾞｰﾀ，"0"("1"以外):該当部署確定ﾃﾞｰﾀ、又は、予定ﾃﾞｰﾀ）
    '           5) 応援先看護単位CD---13ﾊﾞｲﾄ目から４バイト分     'Add Tanaka 2003/06/23
    '　　　　　↓2015/04/10 Bando Mod
    '           6) 希望コメント 16ﾊﾞｲﾄ目から２０バイト分
    '           7) 時間年休---36ﾊﾞｲﾄ目から４４ﾊﾞｲﾄ分
    '*****************************************************************************************************
    '2015/04/10 Bando Upd Start ========================================
    'Public Sub Get_KinmuMark(ByVal p_Val As Object, ByRef p_KinmuCD As String, ByRef p_RiyuKBN As String, ByRef p_Flg As String, ByRef p_KangoCD As String, ByRef p_Time As String)
    Public Sub Get_KinmuMark(ByVal p_Val As Object, ByRef p_KinmuCD As String, ByRef p_RiyuKBN As String, ByRef p_Flg As String, ByRef p_KangoCD As String, ByRef p_Time As String, ByRef p_Comment As String)
        'On Error GoTo Get_KinmuMark
        Const W_SUBNAME As String = "BasNSK0000H Get_KinmuMark"

        Dim w_str As String

        Try
            'セルの値が設定されていない場合
            If Trim(p_Val) = "" Then
                'If IsDBNull(p_Val) = True Then
                p_KinmuCD = ""
                p_RiyuKBN = ""
                p_Flg = "0"
                p_KangoCD = ""
                p_Time = ""
                p_Comment = ""
                Exit Sub
            End If

            '勤務記号以外は半角とする
            w_str = CStr(p_Val)

            'w_str = Right(w_str, 127)
            w_str = General.paRightB(w_str, 147)

            'ｽﾌﾟﾚｯﾄﾞ貼り付け文字列を分割する
            p_KinmuCD = Trim(Left(w_str, 3)) '勤務記号は将来を見据え3バイト分確保する
            p_RiyuKBN = Trim(Mid(w_str, 4, 1))
            p_Flg = Trim(Mid(w_str, 5, 1))
            p_KangoCD = Trim(Mid(w_str, 6, 10))
            p_Comment = Trim(General.paMidB(w_str, 16, 20))
            'p_Time = Trim(Mid(w_str, 16, 112))
            p_Time = Trim(General.paMidB(w_str, 36, 112))
            '2017/09/29 Yamanishi Add ---------------------------------------------------------------------------------------------------------
            If m_dicRandomKey2TimeNenkyu.ContainsKey(p_Time) Then
                p_Time = Trim(m_dicRandomKey2TimeNenkyu(p_Time))
            End If
            '----------------------------------------------------------------------------------------------------------------------------------
            'Get_KinmuMark:
            '        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
            '        End
        Catch ex As Exception
            End
        End Try
    End Sub
    '2015/04/10 Bando Upd End   ========================================

    '******************************************************************************************
    '   ｽﾌﾟﾚｯﾄﾞに貼り付ける勤務記号を編集する
    '   ﾊﾟﾗﾒｰﾀ：p_KinmuCD（勤務ｺｰﾄﾞ）
    '           p_RiyuKBN（理由区分）
    '           p_Time（時間年休）
    '           p_Flg（確定部署フラグ）
    '           p_KangoCD (応援先看護単位CD)
    '   編集仕様
    '       勤務記号(2)＋Space(5)＋勤務ｺｰﾄﾞ(3)＋理由区分(1)＋確定部署フラグ(1)+応援先看護単位CD(4) + 希望コメント(20) +時間年休(44)
    '       全６０バイト分
    '           1) 勤務記号---１ﾊﾞｲﾄ目から７ﾊﾞｲﾄ分（Trimで余分なｽﾍﾟｰｽを省く）
    '           2) KinmuCD---８ﾊﾞｲﾄ目から３ﾊﾞｲﾄ分（Trimで余分なｽﾍﾟｰｽを省く）
    '           3) 理由区分---11ﾊﾞｲﾄ目から１ﾊﾞｲﾄ分
    '           4) 確定部署FLG---12ﾊﾞｲﾄ目から１バイト分
    '               （"1"：他部署確定ﾃﾞｰﾀ，"0"("1"以外):該当部署確定ﾃﾞｰﾀ、又は、予定ﾃﾞｰﾀ）
    '           5) 応援先看護単位CD---13ﾊﾞｲﾄ目から4ﾊﾞｲﾄ     'Add Tanaka 2003/06/23
    '　　　　　↓2014/04/10 Bando Mod
    '           6) 希望コメント 16ﾊﾞｲﾄ目から２０バイト分
    '           7) 時間年休---36ﾊﾞｲﾄ目から４４ﾊﾞｲﾄ分
    '******************************************************************************************
    '2015/04/10 Bando Upd Start =======================================
    'Public Function Set_KinmuMark(ByVal p_KinmuCD As String, ByVal p_RiyuKBN As String, ByVal p_Flg As String, ByVal p_KangoCD As String, ByVal p_Time As String) As Object
    Public Function Set_KinmuMark(ByVal p_KinmuCD As String, ByVal p_RiyuKBN As String, ByVal p_Flg As String, ByVal p_KangoCD As String, ByVal p_Time As String, ByVal p_Comment As String) As Object
        On Error GoTo Set_KinmuMark
        Const W_SUBNAME As String = "BasNSK0000H Set_KinmuMark"

        Dim w_SprText As String

        If IsNumeric(p_KinmuCD) = False Then
            Set_KinmuMark = CObj(Space(69))
            Exit Function
            'ElseIf CShort(p_KinmuCD) < 0 Or CShort(p_KinmuCD) > UBound(g_KinmuM) Then
        ElseIf CShort(p_KinmuCD) <= 0 Or CShort(p_KinmuCD) > UBound(g_KinmuM) Then
            Set_KinmuMark = CObj(Space(69))
            Exit Function
        Else
            '記号
            w_SprText = g_KinmuM(CShort(p_KinmuCD)).Mark
            If 7 - General.paLenB(w_SprText) >= 0 Then
                w_SprText = w_SprText & Space(7 - General.paLenB(w_SprText))
            End If
            'ｺｰﾄﾞ（勤務ｺｰﾄﾞは将来を見据え3ﾊﾞｲﾄ分確保する）
            w_SprText = w_SprText & p_KinmuCD
            If 3 - General.paLenB(p_KinmuCD) >= 0 Then
                w_SprText = w_SprText & Space(3 - General.paLenB(p_KinmuCD))
            End If
            '理由区分
            w_SprText = w_SprText & Left(p_RiyuKBN & Space(1), 1)
            '確定部署ﾌﾗｸﾞ
            If p_Flg = "" Then
                w_SprText = w_SprText & "0"
            Else
                w_SprText = w_SprText & Left(p_Flg, 1)
            End If
            '応援勤務の看護単位CD
            w_SprText = w_SprText & Left(p_KangoCD & Space(10), 10)

            '希望
            w_SprText = w_SprText & General.paLeftB(p_Comment & Space(20), 20)

            '時間数
            '2017/09/29 Yamanishi Upd ---------------------------------------------------------------------------------------------------------
            'w_SprText = w_SprText & Left(p_Time & Space(112), 112)
            p_Time = Trim(p_Time)
            If p_Time.Length <= 112 Then
                w_SprText = w_SprText & Left(p_Time & Space(112), 112)
            Else
                w_SprText = w_SprText & GenerateRandomKeyAndSetDictionary(112, p_Time, m_dicRandomKey2TimeNenkyu)
            End If
            '----------------------------------------------------------------------------------------------------------------------------------
            '編集文字列を返却する
            Set_KinmuMark = CObj(w_SprText)
        End If

        Exit Function
Set_KinmuMark:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Function
    '2015/04/10 Bando Upd End   =======================================

    '2018/03/08 Yamanishi Add Start ------------------------------------------------------------------------------------------
    ''' <summary>
    ''' 時間休の各項目から時間休文字列作成
    ''' </summary>
    ''' <param name="p_BunruiCD">休暇分類コード</param>
    ''' <param name="p_FromTime">開始時刻</param>
    ''' <param name="p_ToTime">終了時刻</param>
    ''' <param name="p_DateKbn">翌日フラグ</param>
    ''' <param name="p_NenkyuTime">年休取得時間</param>
    ''' <param name="p_HolSubFlg">休憩減算フラグ</param>
    ''' <param name="p_DayTime">日勤時間</param>
    ''' <param name="p_NightTime">夜勤時間</param>
    ''' <param name="p_NextNightTime">翌日夜勤時間</param>
    ''' <returns>時間休文字列</returns>
    Public Function Set_NenkyuTime(ByVal p_BunruiCD As Object, ByVal p_FromTime As Object, ByVal p_ToTime As Object,
                                   ByVal p_DateKbn As Object, ByVal p_NenkyuTime As Object, ByVal p_HolSubFlg As Object,
                                   ByVal p_DayTime As Object, ByVal p_NightTime As Object, ByVal p_NextNightTime As Object) As String
        Const W_SUBNAME As String = "BasNSK0000H Set_NenkyuTime"

        Const WC_BunuruiCDLength As Integer = 2
        Const WC_TimeLength As Integer = 4
        Const WC_KbnLength As Integer = 1

        Dim w_Time As String
        Try
            '休暇分類コード
            w_Time = General.paFormatSpace(Convert.ToString(p_BunruiCD), WC_BunuruiCDLength)

            '開始時刻
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_FromTime), WC_TimeLength)

            '終了時刻
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_ToTime), WC_TimeLength)

            '翌日フラグ
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_DateKbn), WC_KbnLength)

            '年休取得時間
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_NenkyuTime), WC_TimeLength)

            '休憩減算フラグ
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_HolSubFlg), WC_KbnLength)

            '日勤時間
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_DayTime), WC_TimeLength)

            '夜勤時間
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_NightTime), WC_TimeLength)

            '翌日夜勤時間
            w_Time = w_Time & General.paFormatSpace(Convert.ToString(p_NextNightTime), WC_TimeLength)

            Return w_Time

        Catch ex As Exception
            Call General.paTrpMsg(Err.Number, W_SUBNAME)
            End
        End Try
    End Function

    ''' <summary>
    ''' 時間休文字列から休暇分類と取得時間を取得する
    ''' </summary>
    ''' <param name="p_Time">時間休文字列</param>
    ''' <param name="p_BunruiCD">休暇分類コード</param>
    ''' <param name="p_NenkyuTime">年休取得時間</param>
    ''' <returns>復元済分を取り除いた時間休文字列</returns>
    Public Function Get_NenkyuTime(ByVal p_Time As String,
                                   ByRef p_BunruiCD As String, ByRef p_NenkyuTime As Integer) As String
        Return Get_NenkyuTime(p_Time, p_BunruiCD, 0, 0, "", p_NenkyuTime, "", 0, 0, 0)
    End Function

    ''' <summary>
    ''' 時間休文字列から日勤・夜勤・翌夜勤時間を取得する
    ''' </summary>
    ''' <param name="p_Time">時間休文字列</param>
    ''' <param name="p_DayTime">日勤時間</param>
    ''' <param name="p_NightTime">夜勤時間</param>
    ''' <param name="p_NextNightTime">翌日夜勤時間</param>
    ''' <returns>復元済分を取り除いた時間休文字列</returns>
    Public Function Get_NenkyuTime(ByVal p_Time As String,
                                   ByRef p_DayTime As Integer, ByRef p_NightTime As Integer, ByRef p_NextNightTime As Integer) As String
        Return Get_NenkyuTime(p_Time, "", 0, 0, "", 0, "", p_DayTime, p_NightTime, p_NextNightTime)
    End Function

    ''' <summary>
    ''' 時間休文字列から各種パラメータを復元する
    ''' </summary>
    ''' <param name="p_Time">時間休文字列</param>
    ''' <param name="p_BunruiCD">休暇分類コード</param>
    ''' <param name="p_FromTime">開始時刻</param>
    ''' <param name="p_ToTime">終了時刻</param>
    ''' <param name="p_DateKbn">翌日フラグ</param>
    ''' <param name="p_NenkyuTime">年休取得時間</param>
    ''' <param name="p_HolSubFlg">休憩減算フラグ</param>
    ''' <param name="p_DayTime">日勤時間</param>
    ''' <param name="p_NightTime">夜勤時間</param>
    ''' <param name="p_NextNightTime">翌日夜勤時間</param>
    ''' <returns>復元済分を取り除いた時間休文字列</returns>
    Public Function Get_NenkyuTime(ByVal p_Time As String,
                                   ByRef p_BunruiCD As String, ByRef p_FromTime As Integer, ByRef p_ToTime As Integer,
                                   ByRef p_DateKbn As String, ByRef p_NenkyuTime As Integer, ByRef p_HolSubFlg As String,
                                   ByRef p_DayTime As Integer, ByRef p_NightTime As Integer, ByRef p_NextNightTime As Integer) As String
        Const W_SUBNAME As String = "BasNSK0000H Get_NenkyuTime"

        Const WC_BunuruiCDLength As Integer = 2
        Const WC_TimeLength As Integer = 4
        Const WC_KbnLength As Integer = 1

        Dim w_Length As Integer
        Dim w_Position As Integer
        Try
            w_Position = 1

            '休暇分類コード
            w_Length = WC_BunuruiCDLength
            p_BunruiCD = Mid(p_Time, w_Position, w_Length)
            w_Position = w_Position + w_Length

            '開始時刻
            w_Length = WC_TimeLength
            If Not Integer.TryParse(Mid(p_Time, w_Position, w_Length), p_FromTime) Then
                p_FromTime = 0
            End If
            w_Position = w_Position + w_Length

            '終了時刻
            w_Length = WC_TimeLength
            If Not Integer.TryParse(Mid(p_Time, w_Position, w_Length), p_ToTime) Then
                p_ToTime = 0
            End If
            w_Position = w_Position + w_Length

            '翌日フラグ
            w_Length = WC_KbnLength
            p_DateKbn = Mid(p_Time, w_Position, w_Length)
            w_Position = w_Position + w_Length

            '年休取得時間
            w_Length = WC_TimeLength
            If Not Integer.TryParse(Mid(p_Time, w_Position, w_Length), p_NenkyuTime) Then
                p_NenkyuTime = 0
            End If
            w_Position = w_Position + w_Length

            '休憩減算フラグ
            w_Length = WC_KbnLength
            p_HolSubFlg = Mid(p_Time, w_Position, w_Length)
            w_Position = w_Position + w_Length

            '日勤時間
            w_Length = WC_TimeLength
            If Not Integer.TryParse(Mid(p_Time, w_Position, w_Length), p_DayTime) Then
                p_DayTime = 0
            End If
            w_Position = w_Position + w_Length

            '夜勤時間
            w_Length = WC_TimeLength
            If Not Integer.TryParse(Mid(p_Time, w_Position, w_Length), p_NightTime) Then
                p_NightTime = 0
            End If
            w_Position = w_Position + w_Length

            '翌日夜勤時間
            w_Length = WC_TimeLength
            If Not Integer.TryParse(Mid(p_Time, w_Position, w_Length), p_NextNightTime) Then
                p_NextNightTime = 0
            End If
            w_Position = w_Position + w_Length

            Return Mid(p_Time, w_Position)

        Catch ex As Exception
            Call General.paTrpMsg(Err.Number, W_SUBNAME)
            End
        End Try
    End Function

    ''' <summary>
    ''' HHmm形式を分数に変換
    ''' </summary>
    ''' <param name="p_HHmm">HHmm</param>
    ''' <returns>分数</returns>
    Public Function HHmmToMin(ByVal p_HHmm As Integer) As Integer
        Return (p_HHmm \ 100) * 60 + (p_HHmm Mod 100)
    End Function
    '2018/03/08 Yamanishi Add End --------------------------------------------------------------------------------------------

    Public Function Get_KinmuTipText(ByVal p_KinmuCD As String) As String
        On Error GoTo Get_KinmuTipText
        Const W_SUBNAME As String = "BasNSK0000H Get_KinmuTipText"

        Dim w_str As String

        If IsNumeric(p_KinmuCD) = False Then
            Get_KinmuTipText = ""
            Exit Function
        ElseIf CShort(p_KinmuCD) <= 0 Or UBound(g_KinmuM) < CShort(p_KinmuCD) Then
            Get_KinmuTipText = ""
            Exit Function
        End If

        '半日勤務か？
        Select Case g_KinmuM(CShort(p_KinmuCD)).HFlg
            Case "1" '全日
                w_str = g_KinmuM(CShort(p_KinmuCD)).KinmuName
            Case "2" '半日
                w_str = g_KinmuM(CShort(g_KinmuM(CShort(p_KinmuCD)).AMCD)).KinmuName & "／" & g_KinmuM(CShort(g_KinmuM(CShort(p_KinmuCD)).PMCD)).KinmuName
            Case Else
                w_str = ""
        End Select

        Get_KinmuTipText = w_str

        Exit Function
Get_KinmuTipText:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Function

    '*****************************************************************************************************
    '   ｽﾌﾟﾚｯﾄﾞの休み分類記号文字列を各データ毎に分割する
    '   ﾊﾟﾗﾒｰﾀ：p_Val(I)（勤務ｺｰﾄﾞ）
    '           p_UniqueSeqNO(O)（UNIQUESEQNO）
    '           p_AppliCD(O)（届出ｺｰﾄﾞ）
    '           p_HolBunruiCD(O)（休み分類ｺｰﾄﾞ）
    '           p_GetContentsKBN(O)（取得内容区分）
    '           p_intTimeFrom(O)（時間FROM）
    '           p_intTimeTo(O)（時間TO）
    '           p_strNextDayFlg(O)（翌日FLG）
    '   編集仕様
    '       休み分類記号(2)＋Space(5)＋休み分類ｺｰﾄﾞ(2)＋UNIQUESEQNO(18)＋届出ｺｰﾄﾞ(6)+取得内容区分(1)＋時間FROM(4)＋時間TO(4)＋翌日FLG(1)
    '       全４３バイト分
    '           1) 休み分類記号---１ﾊﾞｲﾄ目から７ﾊﾞｲﾄ分（Trimで余分なｽﾍﾟｰｽを省く）
    '           2) 休み分類CD---８ﾊﾞｲﾄ目から２ﾊﾞｲﾄ分（Trimで余分なｽﾍﾟｰｽを省く）
    '           3) 届出UniqueSeqNO---10ﾊﾞｲﾄ目から１８ﾊﾞｲﾄ分
    '           4) 届出CD---28ﾊﾞｲﾄ目から６ﾊﾞｲﾄ分
    '           5) 取得内容区分---34ﾊﾞｲﾄ目から１ﾊﾞｲﾄ分
    '               （"1":全日、"2":前半、"3":後半、"4":時間年休）
    '           6) 時間FROM---35ﾊﾞｲﾄ目から４ﾊﾞｲﾄ分
    '           7) 時間TO---39ﾊﾞｲﾄ目から４ﾊﾞｲﾄ分
    '           8) 翌日FLG---43ﾊﾞｲﾄ目から１ﾊﾞｲﾄ分
    '*****************************************************************************************************
    Public Sub Get_AppliMark(ByVal p_Val As Object, ByRef p_UniqueSeqNO As String, ByRef p_AppliCD As String, ByRef p_HolBunruiCD As String, ByRef p_GetContentsKBN As String, ByRef p_intTimeFrom As Short, ByRef p_intTimeTo As Short, ByRef p_strNextDayFlg As String)
        On Error GoTo Get_AppliMark
        Const W_SUBNAME As String = "BasNSK0000H Get_AppliMark"

        Dim w_str As String

        'セルの値が設定されていない場合
        If IsDBNull(p_Val) = True Then
            p_HolBunruiCD = ""
            p_UniqueSeqNO = ""
            p_AppliCD = ""
            p_GetContentsKBN = ""
            p_intTimeFrom = 0
            p_intTimeTo = 0
            p_strNextDayFlg = ""
            Exit Sub
        End If

        '勤務記号以外は半角とする
        w_str = CStr(p_Val)
        w_str = Right(w_str, 36)

        'ｽﾌﾟﾚｯﾄﾞ貼り付け文字列を分割する
        p_HolBunruiCD = Trim(Left(w_str, 2))
        p_UniqueSeqNO = Trim(Mid(w_str, 3, 18))
        p_AppliCD = Trim(Mid(w_str, 21, 6))
        p_GetContentsKBN = Trim(Mid(w_str, 27, 1))
        If IsNumeric(Mid(w_str, 28, 4)) Then
            p_intTimeFrom = CShort(Mid(w_str, 28, 4))
        Else
            p_intTimeFrom = 0
        End If
        If IsNumeric(Mid(w_str, 32, 4)) Then
            p_intTimeTo = CShort(Mid(w_str, 32, 4))
        Else
            p_intTimeTo = 0
        End If
        p_strNextDayFlg = Trim(Mid(w_str, 36, 1))

        Exit Sub
Get_AppliMark:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub

    '******************************************************************************************
    '   ｽﾌﾟﾚｯﾄﾞに貼り付ける休み分類記号を編集する
    '   ﾊﾟﾗﾒｰﾀ：p_UniqueSeqNO（UNIQUESEQNO）
    '           p_AppliCD（届出ｺｰﾄﾞ）
    '           p_HolBunruiCD（休み分類ｺｰﾄﾞ）
    '           p_GetContentsKBN（取得内容区分）
    '           p_intTimeFrom（時間FROM）
    '           p_intTimeTo（時間TO）
    '           p_strNextDayFlg（翌日FLG）
    '   編集仕様
    '       休み分類記号(2)＋Space(5)＋休み分類ｺｰﾄﾞ(2)＋UNIQUESEQNO(18)＋届出ｺｰﾄﾞ(6)+取得内容区分(1)＋時間FROM(4)＋時間TO(4)＋翌日FLG(1)
    '       全４３バイト分
    '           1) 休み分類記号---１ﾊﾞｲﾄ目から７ﾊﾞｲﾄ分（Trimで余分なｽﾍﾟｰｽを省く）
    '           2) 休み分類CD---８ﾊﾞｲﾄ目から２ﾊﾞｲﾄ分（Trimで余分なｽﾍﾟｰｽを省く）
    '           3) 届出UniqueSeqNO---10ﾊﾞｲﾄ目から１８ﾊﾞｲﾄ分
    '           4) 届出CD---28ﾊﾞｲﾄ目から６ﾊﾞｲﾄ分
    '           5) 取得内容区分---34ﾊﾞｲﾄ目から１ﾊﾞｲﾄ分
    '               （"1":全日、"2":前半、"3":後半、"4":時間年休）
    '           6) 時間FROM---35ﾊﾞｲﾄ目から４ﾊﾞｲﾄ分
    '           7) 時間TO---39ﾊﾞｲﾄ目から４ﾊﾞｲﾄ分
    '           8) 翌日FLG---43ﾊﾞｲﾄ目から１ﾊﾞｲﾄ分
    '******************************************************************************************
    Public Function Set_AppliMark(ByVal p_UniqueSeqNO As String, ByVal p_AppliCD As String, ByVal p_HolBunruiCD As String, ByVal p_GetContentsKBN As String, ByVal p_intTimeFrom As Short, ByVal p_intTimeTo As Short, ByVal p_strNextDayFlg As String) As Object
        On Error GoTo Set_AppliMark
        Const W_SUBNAME As String = "BasNSK0000H Set_AppliMark"

        Dim w_SprText As String
        Dim w_intLoop As Short
        Dim w_intIndex As Short
        Dim w_blnFLG As Boolean

        w_blnFLG = False
        For w_intLoop = 1 To UBound(g_HolidayBunruiM)
            If g_HolidayBunruiM(w_intLoop).CD = p_HolBunruiCD Then
                w_blnFLG = True
                w_intIndex = w_intLoop
                Exit For
            End If
        Next w_intLoop

        If w_blnFLG = False Then
            w_SprText = Space(34)
            '時間FROM
            w_SprText = w_SprText & "0000"
            '時間TO
            w_SprText = w_SprText & "0000"
            w_SprText = w_SprText & Space(1)
            Set_AppliMark = CObj(w_SprText)
            Exit Function
        Else
            '記号
            w_SprText = g_HolidayBunruiM(w_intIndex).Mark
            If 7 - General.paLenB(w_SprText) >= 0 Then
                w_SprText = w_SprText & Space(7 - General.paLenB(w_SprText))
            End If
            'ｺｰﾄﾞ
            w_SprText = w_SprText & p_HolBunruiCD
            If 2 - General.paLenB(p_HolBunruiCD) >= 0 Then
                w_SprText = w_SprText & Space(2 - General.paLenB(p_HolBunruiCD))
            End If
            'UNIQUESEQNO
            w_SprText = w_SprText & p_UniqueSeqNO
            If 18 - General.paLenB(p_UniqueSeqNO) >= 0 Then
                w_SprText = w_SprText & Space(18 - General.paLenB(p_UniqueSeqNO))
            End If
            '届出ｺｰﾄﾞ
            w_SprText = w_SprText & p_AppliCD
            If 6 - General.paLenB(p_AppliCD) >= 0 Then
                w_SprText = w_SprText & Space(6 - General.paLenB(p_AppliCD))
            End If
            '取得内容区分
            w_SprText = w_SprText & Left(p_GetContentsKBN & Space(1), 1)
            '時間FROM
            w_SprText = w_SprText & Left(Format(p_intTimeFrom, "0000"), 4)
            '時間TO
            w_SprText = w_SprText & Left(Format(p_intTimeTo, "0000"), 4)
            '翌日FLG
            w_SprText = w_SprText & Left(p_strNextDayFlg & Space(1), 1)

            '編集文字列を返却する
            Set_AppliMark = CObj(w_SprText)

        End If

        Exit Function
Set_AppliMark:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Function

    Public Function Get_AppliTipText(ByVal p_HolBunruiCD As String) As String
        On Error GoTo Get_AppliTipText
        Const W_SUBNAME As String = "BasNSK0000H Get_AppliTipText"

        Dim w_intLoop As Short
        Dim w_intIndex As Short

        w_intIndex = -1
        For w_intLoop = 1 To UBound(g_HolidayBunruiM)
            If g_HolidayBunruiM(w_intLoop).CD = p_HolBunruiCD Then
                w_intIndex = w_intLoop
                Exit For
            End If
        Next w_intLoop

        If w_intIndex = -1 Then
            Get_AppliTipText = ""
            Exit Function
        End If

        Get_AppliTipText = g_HolidayBunruiM(w_intIndex).HolidayName

        Exit Function
Get_AppliTipText:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Function

    '*****************************************************************************************************
    '   ｽﾌﾟﾚｯﾄﾞの勤務記号文字列を各データ毎に分割する
    '   ﾊﾟﾗﾒｰﾀ：p_Val(I)（勤務ｺｰﾄﾞ）
    '           p_KinmuCD(O)（日当直勤務ｺｰﾄﾞ）
    '           p_GroupCD(O)（日当直グループｺｰﾄﾞ）
    '   編集仕様
    '       日当直勤務記号(2)＋Space(5)＋日当直勤務ｺｰﾄﾞ(3)＋日当直グループｺｰﾄﾞ(2)＋２件目(日当直勤務記号(2)＋Space(5)＋日当直勤務ｺｰﾄﾞ(3)＋日当直グループｺｰﾄﾞ(2))＋３件目以降・・・
    '       全１２バイト×件数分
    '           1) 日当直勤務記号---１ﾊﾞｲﾄ目から７ﾊﾞｲﾄ分（Trimで余分なｽﾍﾟｰｽを省く）
    '           2) 日当直勤務CD---８ﾊﾞｲﾄ目から３ﾊﾞｲﾄ分（Trimで余分なｽﾍﾟｰｽを省く）
    '           3) 日当直グループCD---11ﾊﾞｲﾄ目から２ﾊﾞｲﾄ分
    '           4) ２件目以降---13ﾊﾞｲﾄ目から12バイト分
    '*****************************************************************************************************
    Public Sub Get_DutyMark(ByVal p_Val As Object, ByRef p_KinmuCD As Object, ByRef p_GroupCD As Object)
        On Error GoTo Get_DutyMark
        Const W_SUBNAME As String = "BasNSK0000H Get_DutyMark"

        Dim w_str As String
        Dim w_Str2 As String
        Dim w_intCount As Short
        Dim w_intLoop As Short

        'セルの値が設定されていない場合
        If IsDBNull(p_Val) = True Then
            ReDim p_KinmuCD(0)
            ReDim p_GroupCD(0)
            Exit Sub
        End If

        '勤務記号以外は半角とする
        w_str = CStr(p_Val)
        w_intCount = General.paLenB(w_str) / 12
        ReDim p_KinmuCD(w_intCount)
        ReDim p_GroupCD(w_intCount)
        For w_intLoop = 1 To w_intCount
            w_Str2 = General.paLeftB(w_str, w_intLoop * 12)
            w_Str2 = Right(w_Str2, 5)

            'ｽﾌﾟﾚｯﾄﾞ貼り付け文字列を分割する
            p_KinmuCD(w_intLoop) = Trim(Left(w_Str2, 3)) '勤務記号は将来を見据え3バイト分確保する
            p_GroupCD(w_intLoop) = Trim(Mid(w_Str2, 4, 2))
        Next w_intLoop

        Exit Sub
Get_DutyMark:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Sub

    '******************************************************************************************
    '   ｽﾌﾟﾚｯﾄﾞに貼り付ける勤務記号を編集する
    '   ﾊﾟﾗﾒｰﾀ：p_KinmuCD（日当直勤務ｺｰﾄﾞ）
    '           p_GroupCD（日当直グループｺｰﾄﾞ）
    '   編集仕様
    '       日当直勤務記号(2)＋Space(5)＋日当直勤務ｺｰﾄﾞ(3)＋日当直グループｺｰﾄﾞ(2)＋２件目(日当直勤務記号(2)＋Space(5)＋日当直勤務ｺｰﾄﾞ(3)＋日当直グループｺｰﾄﾞ(2))＋３件目以降・・・
    '       全１２バイト×件数分
    '           1) 日当直勤務記号---１ﾊﾞｲﾄ目から７ﾊﾞｲﾄ分（Trimで余分なｽﾍﾟｰｽを省く）
    '           2) 日当直勤務CD---８ﾊﾞｲﾄ目から３ﾊﾞｲﾄ分（Trimで余分なｽﾍﾟｰｽを省く）
    '           3) 日当直グループCD---11ﾊﾞｲﾄ目から２ﾊﾞｲﾄ分
    '           4) ２件目以降---13ﾊﾞｲﾄ目から12バイト分
    '******************************************************************************************
    Public Function Set_DutyMark(ByVal p_KinmuCD As Object, ByVal p_GroupCD As Object) As Object
        On Error GoTo Set_DutyMark
        Const W_SUBNAME As String = "BasNSK0000H Set_DutyMark"

        Dim w_SprText As String
        Dim w_intLoop As Short

        Set_DutyMark = ""

        For w_intLoop = 1 To UBound(p_KinmuCD)
            If IsNumeric(p_KinmuCD(w_intLoop)) = False Then
            ElseIf CShort(p_KinmuCD(w_intLoop)) < 0 Or CShort(p_KinmuCD(w_intLoop)) > UBound(g_KinmuM) Then
            Else
                '記号
                w_SprText = g_KinmuM(CShort(p_KinmuCD(w_intLoop))).Mark
                If 7 - General.paLenB(w_SprText) >= 0 Then
                    w_SprText = w_SprText & Space(7 - General.paLenB(w_SprText))
                End If
                'ｺｰﾄﾞ（勤務ｺｰﾄﾞは将来を見据え3ﾊﾞｲﾄ分確保する）
                w_SprText = w_SprText & p_KinmuCD(w_intLoop)
                If 3 - General.paLenB(p_KinmuCD(w_intLoop)) >= 0 Then
                    w_SprText = w_SprText & Space(3 - General.paLenB(p_KinmuCD(w_intLoop)))
                End If
                'グループｺｰﾄﾞ
                w_SprText = w_SprText & Left(p_GroupCD(w_intLoop) & Space(2), 2)

                '編集文字列を返却する
                Set_DutyMark = Set_DutyMark & CObj(w_SprText)
            End If
        Next w_intLoop

        If Set_DutyMark = "" Then
            Set_DutyMark = CObj(Space(12))
        End If

        Exit Function
Set_DutyMark:
        Call General.paTrpMsg(CStr(Err.Number), W_SUBNAME)
        End
    End Function

    '表示順名格納用
    Public m_HyoujijunMeiDate() As HyoujijunMeiDate_Type

    Public Structure HyoujijunMeiDate_Type
        Dim HName As String '表示順名称
        Dim HMasterCD As String '表示順CD
    End Structure

    ''' <summary>
    ''' 押下キー判定
    ''' </summary>
    ''' <param name="p_key"></param>
    ''' <returns></returns>
    ''' <remarks>キーボード対応キーかどうか判定する</remarks>
    Public Function IsNumOrFuncKey(ByVal p_key As System.Windows.Forms.Keys) As Boolean
        Dim w_PreErrorProc As String = General.g_ErrorProc
        General.g_ErrorProc = "NSC0000HA Get_KeyBoardKinmu"

        Dim rtnFlg As Boolean = False

        Try
            Select Case p_key
                Case Keys.NumPad0, Keys.NumPad1, Keys.NumPad2, Keys.NumPad3, Keys.NumPad4, _
                     Keys.NumPad5, Keys.NumPad6, Keys.NumPad7, Keys.NumPad8, Keys.NumPad9
                    'テンキーの数字
                    rtnFlg = True

                Case Keys.D0, Keys.D1, Keys.D2, Keys.D3, Keys.D4, _
                     Keys.D5, Keys.D6, Keys.D7, Keys.D8, Keys.D9
                    'キーボードの数字
                    rtnFlg = True

                Case Keys.F1, Keys.F2, Keys.F3, Keys.F4, Keys.F5, Keys.F6, _
                     Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12
                    'ファンクションキー
                    rtnFlg = True

                Case Else
            End Select

            Return rtnFlg
        Catch ex As Exception
            Throw
        End Try
    End Function

    '2014/04/23 Saijo add start P-06979---------------------------------------------------------------------------
    ''' <summary>
    ''' 項目設定「勤務記号全角２文字対応フラグ」取得
    ''' </summary>
    ''' <param name="p_HospitalCD">病院CD</param>
    ''' <returns>String ("0"：対応しない、"1":対応する)</returns>
    ''' <remarks>勤務記号全角２文字対応フラグ(0：対応しない、1:対応する)</remarks>
    Public Function Get_ItemValue(ByVal p_HospitalCD As String) As String
        Dim w_PreErrorProc As String = General.g_ErrorProc
        General.g_ErrorProc = "NSK0000H Get_ItemValue"

        Try
            Get_ItemValue = General.paGetItemValue( _
            General.G_STRMAINKEY1, General.G_STRSUBKEY1, "KINMUEMSECONDFLG", "0", p_HospitalCD)

        Catch ex As Exception
            Throw
        End Try
    End Function
    '2014/04/23 Saijo add end P-06979-------------------------------------------------------------------------------

    '2017/08/24 Angelo add st---------------------------------------------------------------------------------------
    '画面表示用の編集処理
    Public Function EditData(ByVal p_objValue As Object, ByVal p_intEditMode As Integer) As String
        Const W_SUBNAME As String = "BasNSK0000H EditData"

        Dim w_strEditValue As String

        EditData = ""
        Try
            '初期化
            w_strEditValue = ""

            If p_intEditMode = G_EDITMODE_NO Then
                '0⇒00に変換
                w_strEditValue = General.paFormatZero(p_objValue, 2)
            ElseIf p_intEditMode = G_EDITMODE_DATETIME Then
                If p_objValue <> 0 Then
                    'yyyyMMddHHmmss⇒yyyy/MM/dd HH:mm:ssに変換
                    w_strEditValue = Format(p_objValue, "0000/00/00 00:00:00")
                End If
            End If

            EditData = w_strEditValue
        Catch es As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
        End Try
    End Function
    Public Sub GetNenkyuContentsKbnAndHolCD(ByVal p_KinmuCD As String, ByRef p_GetContentsKbn As String, ByRef p_HolCD As String)
        Try
            p_GetContentsKbn = ""
            p_HolCD = ""
            If IsNumeric(p_KinmuCD) Then
                If 0 <= Integer.Parse(p_KinmuCD) AndAlso Integer.Parse(p_KinmuCD) <= UBound(g_KinmuM) Then
                    If g_KinmuM(Integer.Parse(p_KinmuCD)).HFlg = "2" Then
                        '半日勤務ﾌﾗｸﾞ
                        If IsNumeric(g_KinmuM(Integer.Parse(p_KinmuCD)).AMCD) Then
                            'ＡＭ勤務CD
                            If 0 <= Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).AMCD) AndAlso Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).AMCD) <= UBound(g_KinmuM) Then
                                If Not String.IsNullOrEmpty(g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).AMCD)).HolBunruiCD) AndAlso
                                            g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).AMCD)).HolBunruiCD <> General.G_STRDAIKYUBUNRUI Then
                                    '休み分類CD
                                    p_GetContentsKbn = "2"
                                    p_HolCD = g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).AMCD)).HolBunruiCD
                                End If
                            End If
                        End If
                        If IsNumeric(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD) Then
                            'ＰＭ勤務CD
                            If 0 <= Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD) AndAlso Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD) <= UBound(g_KinmuM) Then
                                If Not String.IsNullOrEmpty(g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD)).HolBunruiCD) AndAlso
                                            g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD)).HolBunruiCD <> General.G_STRDAIKYUBUNRUI Then
                                    '休み分類CD
                                    If p_GetContentsKbn = "2" Then
                                        p_GetContentsKbn = "2,3"
                                        p_HolCD = p_HolCD & "," & g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD)).HolBunruiCD
                                    Else
                                        p_GetContentsKbn = "3"
                                        p_HolCD = g_KinmuM(Integer.Parse(g_KinmuM(Integer.Parse(p_KinmuCD)).PMCD)).HolBunruiCD
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If Not String.IsNullOrEmpty(g_KinmuM(Integer.Parse(p_KinmuCD)).HolBunruiCD) AndAlso
                                    g_KinmuM(Integer.Parse(p_KinmuCD)).HolBunruiCD <> General.G_STRDAIKYUBUNRUI Then
                            '休み分類CD
                            p_GetContentsKbn = "1"
                            p_HolCD = g_KinmuM(Integer.Parse(p_KinmuCD)).HolBunruiCD
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub
    '*****************************************************************************************************
    '   ｽﾌﾟﾚｯﾄﾞの時間年休文字列を各データ毎に分割する
    '   ﾊﾟﾗﾒｰﾀ：p_Str(I)（時間年休文字列）
    '           p_NenkyuDetail(O)（時間年休詳細情報）
    '   編集仕様
    '       休み分類ｺｰﾄﾞ(2)＋時間FROM(4)＋時間TO(4)＋翌日FLG(1)＋２件目(休み分類ｺｰﾄﾞ(2)＋時間FROM(4)＋時間TO(4)＋翌日FLG(1))＋３件目以降・・・
    '       全１１バイト×件数分
    '           1) 休み分類CD---１ﾊﾞｲﾄ目から２ﾊﾞｲﾄ分（Trimで余分なｽﾍﾟｰｽを省く）
    '           2) 取得内容区分---34ﾊﾞｲﾄ目から１ﾊﾞｲﾄ分
    '               （"1":全日、"2":前半、"3":後半、"4":時間年休）
    '           3) 時間FROM---３ﾊﾞｲﾄ目から４ﾊﾞｲﾄ分
    '           4) 時間TO---７ﾊﾞｲﾄ目から４ﾊﾞｲﾄ分
    '           5) 翌日FLG---11ﾊﾞｲﾄ目から１ﾊﾞｲﾄ分
    '           6) ２件目以降---12ﾊﾞｲﾄ目から11バイト分
    '*****************************************************************************************************
    Public Sub Get_NenkyuDetail(ByVal p_Str As String, ByVal p_Str2 As String, ByRef p_NenkyuDetail() As NenkyuDetail_Type, ByVal p_HolCD As String)
        Const W_SUBNAME As String = "NSK0000HA Get_NenkyuDetail"

        '2018/03/08 Yamanishi Upd -----------------------------------
        'Dim w_Loop As Integer
        'Dim w_Index As Integer
        'Dim w_RecCnt As Integer
        'Dim w_Pos As Integer
        'Dim w_StartTime As String 'FromTime
        'Dim w_EndTime As String 'ToTime
        'Dim w_NenkyuTime As Integer '時間年休
        'Dim w_DayTime As Integer '日勤時間
        'Dim w_NightTime As Integer '夜勤時間
        'Dim w_NextNightTime As Integer '翌日夜勤時間
        Dim w_Index As Integer
        Dim w_RecCnt As Integer
        '------------------------------------------------------------
        Dim w_obj As Object
        Try
            '2018/03/08 Yamanishi Upd Start ------------------------------------------------------------------
            ''時間年休がある時のみ
            'If p_Str <> "" Then
            '    '時間年休件数取得（１件につき２８ﾊﾞｲﾄ）
            '    w_RecCnt = Len(p_Str) / 28
            '    w_Index = UBound(p_NenkyuDetail)
            '    ReDim Preserve p_NenkyuDetail(w_Index + w_RecCnt)

            '    w_Pos = 1

            '    'ｽﾌﾟﾚｯﾄﾞ貼り付け文字列を分割する
            '    For w_Loop = 1 To w_RecCnt
            '        '取得内容区分(4:時間年休)
            '        p_NenkyuDetail(w_Index + w_Loop).GetContentsKbn = "4"

            '        'HolidayBunruiCD
            '        p_NenkyuDetail(w_Index + w_Loop).HolidayBunruiCD = Mid(p_Str, w_Pos, 2)

            '        w_Pos = w_Pos + 2

            '        'FromTime
            '        w_StartTime = Mid(p_Str, w_Pos, 4)
            '        If IsNumeric(w_StartTime) Then
            '            p_NenkyuDetail(w_Index + w_Loop).FromTime = Integer.Parse(w_StartTime)
            '        Else
            '            p_NenkyuDetail(w_Index + w_Loop).FromTime = 0
            '        End If

            '        w_Pos = w_Pos + 4

            '        'ToTime
            '        w_EndTime = Mid(p_Str, w_Pos, 4)
            '        If IsNumeric(w_EndTime) Then
            '            p_NenkyuDetail(w_Index + w_Loop).ToTime = Integer.Parse(w_EndTime)
            '        Else
            '            p_NenkyuDetail(w_Index + w_Loop).ToTime = 0
            '        End If

            '        w_Pos = w_Pos + 4

            '        'DateKbn
            '        p_NenkyuDetail(w_Index + w_Loop).DateKbn = Mid(p_Str, w_Pos, 1)

            '        w_Pos = w_Pos + 1

            '        '時間年休
            '        w_NenkyuTime = Integer.Parse(Mid(p_Str, w_Pos, 4))
            '        If IsNumeric(w_NenkyuTime) Then
            '            p_NenkyuDetail(w_Index + w_Loop).NenkyuTime = Integer.Parse(w_NenkyuTime)
            '        Else
            '            p_NenkyuDetail(w_Index + w_Loop).NenkyuTime = 0
            '        End If

            '        w_Pos = w_Pos + 4

            '        '休憩減算フラグ
            '        p_NenkyuDetail(w_Index + w_Loop).HolSubFlg = Mid(p_Str, w_Pos, 1)

            '        w_Pos = w_Pos + 1

            '        '日勤時間
            '        w_DayTime = Integer.Parse(Mid(p_Str, w_Pos, 4))
            '        If IsNumeric(w_DayTime) Then
            '            p_NenkyuDetail(w_Index + w_Loop).DayTime = Integer.Parse(w_DayTime)
            '        Else
            '            p_NenkyuDetail(w_Index + w_Loop).DayTime = 0
            '        End If

            '        w_Pos = w_Pos + 4

            '        '夜勤時間
            '        w_NightTime = Integer.Parse(Mid(p_Str, w_Pos, 4))
            '        If IsNumeric(w_NightTime) Then
            '            p_NenkyuDetail(w_Index + w_Loop).NightTime = Integer.Parse(w_NightTime)
            '        Else
            '            p_NenkyuDetail(w_Index + w_Loop).NightTime = 0
            '        End If

            '        w_Pos = w_Pos + 4

            '        '翌日夜勤時間
            '        w_NextNightTime = Integer.Parse(Mid(p_Str, w_Pos, 4))
            '        If IsNumeric(w_NextNightTime) Then
            '            p_NenkyuDetail(w_Index + w_Loop).NextNightTime = Integer.Parse(w_NextNightTime)
            '        Else
            '            p_NenkyuDetail(w_Index + w_Loop).NextNightTime = 0
            '        End If

            '        w_Pos = w_Pos + 4
            '    Next w_Loop
            'End If

            '時間年休がある時のみ
            While p_Str <> ""
                w_RecCnt = UBound(p_NenkyuDetail) + 1
                ReDim Preserve p_NenkyuDetail(w_RecCnt)

                With p_NenkyuDetail(w_RecCnt)
                    '取得内容区分(4:時間年休)
                    .GetContentsKbn = General.G_GetContentsKbn_Time

                    p_Str = Get_NenkyuTime(p_Str,
                                           .HolidayBunruiCD, .FromTime, .ToTime,
                                           .DateKbn, .NenkyuTime, .HolSubFlg,
                                           .DayTime, .NightTime, .NextNightTime)

                End With
            End While
            '2018/03/08 Yamanishi Upd End --------------------------------------------------------------------

            '全半日年休がある時のみ
            If p_Str2 <> "" Then
                '年休件数取得
                w_Index = UBound(p_NenkyuDetail)

                If p_Str2 = "2,3" Then
                    w_obj = General.paSplit(p_HolCD, ",")
                    ReDim Preserve p_NenkyuDetail(w_Index + 2)
                    '取得内容区分
                    p_NenkyuDetail(w_Index + 1).GetContentsKbn = "2"
                    'HolidayBunruiCD
                    p_NenkyuDetail(w_Index + 1).HolidayBunruiCD = w_obj(0)
                    'FromTime
                    p_NenkyuDetail(w_Index + 1).FromTime = 0
                    'ToTime
                    p_NenkyuDetail(w_Index + 1).ToTime = 0
                    'DateKbn
                    p_NenkyuDetail(w_Index + 1).DateKbn = "0"
                    '取得内容区分
                    p_NenkyuDetail(w_Index + 2).GetContentsKbn = "3"
                    'HolidayBunruiCD
                    p_NenkyuDetail(w_Index + 2).HolidayBunruiCD = w_obj(1)
                    'FromTime
                    p_NenkyuDetail(w_Index + 2).FromTime = 0
                    'ToTime
                    p_NenkyuDetail(w_Index + 2).ToTime = 0
                    'DateKbn
                    p_NenkyuDetail(w_Index + 2).DateKbn = "0"
                Else
                    ReDim Preserve p_NenkyuDetail(w_Index + 1)
                    '取得内容区分
                    p_NenkyuDetail(w_Index + 1).GetContentsKbn = p_Str2
                    'HolidayBunruiCD
                    p_NenkyuDetail(w_Index + 1).HolidayBunruiCD = p_HolCD
                    'FromTime
                    p_NenkyuDetail(w_Index + 1).FromTime = 0
                    'ToTime
                    p_NenkyuDetail(w_Index + 1).ToTime = 0
                    'DateKbn
                    p_NenkyuDetail(w_Index + 1).DateKbn = "0"
                End If
            End If

        Catch ex As Exception
            Call General.paTrpMsg(Convert.ToString(Err.Number), W_SUBNAME)
            End
        End Try
    End Sub
    '2017/08/24 Angelo add en---------------------------------------------------------------------------------------
End Module