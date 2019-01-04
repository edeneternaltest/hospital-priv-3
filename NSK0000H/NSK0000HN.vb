Option Strict Off
Option Explicit On
Friend Class frmNSK0000HN
    Inherits General.FormBase
	
	Private Const M_ErrChkItem_Count As Short = 0 '勤務／休日回数ﾗﾍﾞﾙｲﾝﾃﾞｯｸｽ
	Private Const M_ErrChkItem_Interval As Short = 1 '勤務／休日間隔ﾗﾍﾞﾙｲﾝﾃﾞｯｸｽ
	Private Const M_ErrChkItem_Pattern As Short = 2 '禁止勤務パターンﾗﾍﾞﾙｲﾝﾃﾞｯｸｽ
	Private Const M_ErrChkItem_NotKinmu As Short = 3 '禁止勤務ﾗﾍﾞﾙｲﾝﾃﾞｯｸｽ

	Private m_ViewFrom As Integer '表示期間開始日
	Private m_ViewTo As Integer '表示期間終了日
	Private m_4WeekFrom As Object '4週/1ヶ月の場合での4週期間の開始日（終了日は計算し求める）	
	Private m_ChkCount As Boolean '回数ﾁｪｯｸ？（True:ﾁｪｯｸ，False:未ﾁｪｯｸ）
	Private m_ChkInterval As Boolean '間隔ﾁｪｯｸ？（True:ﾁｪｯｸ，False:未ﾁｪｯｸ）
	Private m_ChkPattern As Boolean '禁止ﾊﾟﾀｰﾝﾁｪｯｸ？（True:ﾁｪｯｸ，False:未ﾁｪｯｸ）
	Private m_ChkNotKinmu As Boolean '禁止勤務ﾁｪｯｸ？（True:ﾁｪｯｸ，False:未ﾁｪｯｸ）
    Private m_ChkStaffPattern As Boolean '禁止職員ﾊﾟﾀｰﾝﾁｪｯｸ？（True:ﾁｪｯｸ，False:未ﾁｪｯｸ）
    Private m_ChkGiryoType As Boolean '
    Private m_RenzokuKinmuCheck As Boolean '連続勤務
    Private m_AbsoluteKinmuCheck As Boolean '必須勤務
	
	Private Structure OutputType
        Dim Date_Renamed As Integer
		Dim StaffName As String
		Dim ErrorDetail As String
        Dim ErrorName As String
		Dim StaffIdx As Short '対象の行インデックス
		Dim ColIdx As Short '対象の列インデックス
	End Structure
	
	Private m_ErrorList() As OutputType
	Private m_intSelIdx As Short
	
    Public WriteOnly Property pCountCheck() As Boolean
        Set(ByVal Value As Boolean)
            m_ChkCount = Value
        End Set
    End Property

	Public WriteOnly Property pIntervalCheck() As Boolean
		Set(ByVal Value As Boolean)
			m_ChkInterval = Value
		End Set
    End Property

	Public WriteOnly Property pNotKinmuCheck() As Boolean
		Set(ByVal Value As Boolean)
			m_ChkNotKinmu = Value
		End Set
    End Property

	Public WriteOnly Property pPatternCheck() As Boolean
		Set(ByVal Value As Boolean)
			m_ChkPattern = Value
		End Set
    End Property

    Public WriteOnly Property pStaffPatternCheck() As Boolean
        Set(ByVal Value As Boolean)
            m_ChkStaffPattern = Value
        End Set
    End Property

    Public WriteOnly Property pGiryoType() As Boolean
        Set(ByVal Value As Boolean)
            m_ChkGiryoType = Value
        End Set
    End Property

    Public WriteOnly Property pRenzokuKinmuCheck() As Boolean
        Set(ByVal Value As Boolean)
            m_RenzokuKinmuCheck = Value
        End Set
    End Property

    Public WriteOnly Property pAbsoluteKinmuCheck() As Boolean
        Set(ByVal Value As Boolean)
            m_AbsoluteKinmuCheck = Value
        End Set
    End Property

	Public ReadOnly Property pStaffIdxGet() As Short
		Get
			If m_intSelIdx <= UBound(m_ErrorList) Then
				pStaffIdxGet = m_ErrorList(m_intSelIdx).StaffIdx
			End If
		End Get
    End Property

	Public ReadOnly Property pDateIdxGet() As Short
		Get
			If m_intSelIdx <= UBound(m_ErrorList) Then
				pDateIdxGet = m_ErrorList(m_intSelIdx).ColIdx
			End If
		End Get
    End Property

	Public WriteOnly Property pViewFrom() As Integer
		Set(ByVal Value As Integer)
			m_ViewFrom = Value
		End Set
    End Property

	Public WriteOnly Property pViewTo() As Integer
		Set(ByVal Value As Integer)
			m_ViewTo = Value
		End Set
	End Property
	
	Private Sub Set_CountErrInf()
		On Error GoTo Set_CountErrInf
		Const W_SUBNAME As String = "NSK0000HN Set_CountErrInf"
		
		Dim w_Loop As Short
        Dim w_DataIndex As Short
		Dim w_CmbIndex As Short
		
        For w_Loop = 1 To UBound(g_KikanError2)
            For w_CmbIndex = 1 To UBound(g_KikanError2(w_Loop).CheckSpan)
                If g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).ErrorFlg = True And g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).ErrorDate >= m_ViewFrom And g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).ErrorDate <= m_ViewTo Then

                    'エラー情報表示
                    w_DataIndex = UBound(m_ErrorList) + 1
                    ReDim Preserve m_ErrorList(w_DataIndex)

                    'エラー期間
                    m_ErrorList(w_DataIndex).Date_Renamed = g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).ErrorDate

                    '職員インデックス
                    m_ErrorList(w_DataIndex).StaffIdx = g_KikanError2(w_Loop).StaffIdx

                    '氏名
                    m_ErrorList(w_DataIndex).StaffName = g_KikanError2(w_Loop).StaffName

                    '日付インデックス
                    m_ErrorList(w_DataIndex).ColIdx = g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).ColIdx

                    'エラー内容
                    m_ErrorList(w_DataIndex).ErrorDetail = g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).KinmuName & Space(2) & String.Format("{0, 3}", g_KikanError2(w_Loop).CheckSpan(w_CmbIndex).KinmuCount) & "回"

                    'エラー項目
                    m_ErrorList(w_DataIndex).ErrorName = "勤務／休日回数"
                End If
            Next w_CmbIndex
        Next w_Loop
		
		Exit Sub
Set_CountErrInf: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
		End
	End Sub

    '2015/05/13 Ishiga add start-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Set_RenzokuCountErrInf()
        On Error GoTo Set_RenzokuCountErrInf
        Const W_SUBNAME As String = "NSK0000HN Set_RenzokuCountErrInf"

        Dim w_Loop As Short
        Dim w_DataIndex As Short

        For w_Loop = 1 To UBound(g_RenzokuError2)
            If g_RenzokuError2(w_Loop).CheckSpan(0).ErrorFlg = True And g_RenzokuError2(w_Loop).CheckSpan(0).ErrorDate >= m_ViewFrom And g_RenzokuError2(w_Loop).CheckSpan(0).ErrorDate <= m_ViewTo Then

                'エラー情報表示
                w_DataIndex = UBound(m_ErrorList) + 1
                ReDim Preserve m_ErrorList(w_DataIndex)

                'エラー期間
                m_ErrorList(w_DataIndex).Date_Renamed = g_RenzokuError2(w_Loop).CheckSpan(0).ErrorDate

                '職員インデックス
                m_ErrorList(w_DataIndex).StaffIdx = g_RenzokuError2(w_Loop).StaffIdx

                '氏名
                m_ErrorList(w_DataIndex).StaffName = g_RenzokuError2(w_Loop).StaffName

                '日付インデックス
                m_ErrorList(w_DataIndex).ColIdx = g_RenzokuError2(w_Loop).CheckSpan(0).ColIdx

                'エラー内容
                m_ErrorList(w_DataIndex).ErrorDetail = g_RenzokuError2(w_Loop).CheckSpan(0).KinmuName & Space(2) & String.Format("{0, 3}", g_RenzokuError2(w_Loop).CheckSpan(0).KinmuCount) & "回"

                'エラー項目
                m_ErrorList(w_DataIndex).ErrorName = "勤務／休日連続"
            End If
        Next w_Loop

        Exit Sub
Set_RenzokuCountErrInf:
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
    End Sub
    '2015/05/13 Ishiga add end---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

	Private Sub Set_ErrorInf(ByVal pIndex As Short)
		On Error GoTo Set_ErrorInf
		Const W_SUBNAME As String = "NSK0000HN Set_ErrorInf"
		
        Dim w_Int As Short
		Dim w_intCnt As Short
		Dim w_str As String
		Dim w_i As Short
        Dim w_LItem As ListViewItem
		
		lvwErrorList.Items.Clear()
		
        If m_ChkCount = True Then
            '回数選択

            '各勤務毎（集計項目）のエラー情報表示
            Call Set_CountErrInf()
        End If

		If m_ChkInterval = True Then
            '間隔選択
			
			'各勤務毎（集計項目）のエラー情報表示
			Call Set_IntervalErrInf()
        End If

		If m_ChkPattern = True Then
            '禁止パターン
			
			'エラー表示状況表示
			For w_Int = 1 To UBound(g_NotPatternError2)
                For w_i = 1 To UBound(g_NotPatternError2(w_Int).Data)
                    If g_NotPatternError2(w_Int).Data(w_i).ErrorFlg = True And (g_NotPatternError2(w_Int).Data(w_i).ErrorDate >= m_ViewFrom And g_NotPatternError2(w_Int).Data(w_i).ErrorDate <= m_ViewTo) Or (g_NotPatternError2(w_Int).Data(w_i).EndDate >= m_ViewFrom And g_NotPatternError2(w_Int).Data(w_i).EndDate <= m_ViewTo) Then

                        w_intCnt = UBound(m_ErrorList) + 1
                        ReDim Preserve m_ErrorList(w_intCnt)

                        'エラーの場合
                        m_ErrorList(w_intCnt).StaffIdx = g_NotPatternError2(w_Int).StaffIdx
                        m_ErrorList(w_intCnt).StaffName = g_NotPatternError2(w_Int).StaffName
                        m_ErrorList(w_intCnt).Date_Renamed = g_NotPatternError2(w_Int).Data(w_i).ErrorDate
                        m_ErrorList(w_intCnt).ErrorDetail = g_NotPatternError2(w_Int).Data(w_i).ErrorPattern
                        m_ErrorList(w_intCnt).ColIdx = g_NotPatternError2(w_Int).Data(w_i).ColIdx
                        'エラー項目
                        m_ErrorList(w_intCnt).ErrorName = "禁止勤務パターン"
                    End If
                Next w_i
			Next w_Int
        End If

		If m_ChkNotKinmu = True Then
            '禁止勤務
			'エラー表示状況表示
			For w_Int = 1 To UBound(g_NotKinmuError2)
                For w_i = 1 To UBound(g_NotKinmuError2(w_Int).Data)
                    If g_NotKinmuError2(w_Int).Data(w_i).ErrorFlg = True And g_NotKinmuError2(w_Int).Data(w_i).ErrorDate >= m_ViewFrom And g_NotKinmuError2(w_Int).Data(w_i).ErrorDate <= m_ViewTo Then

                        w_intCnt = UBound(m_ErrorList) + 1
                        ReDim Preserve m_ErrorList(w_intCnt)

                        'エラーの場合
                        m_ErrorList(w_intCnt).StaffIdx = g_NotKinmuError2(w_Int).StaffIdx
                        m_ErrorList(w_intCnt).StaffName = g_NotKinmuError2(w_Int).StaffName
                        m_ErrorList(w_intCnt).Date_Renamed = g_NotKinmuError2(w_Int).Data(w_i).ErrorDate
                        m_ErrorList(w_intCnt).ErrorDetail = g_NotKinmuError2(w_Int).Data(w_i).KinmuName
                        m_ErrorList(w_intCnt).ColIdx = g_NotKinmuError2(w_Int).Data(w_i).ColIdx
                        'エラー項目
                        m_ErrorList(w_intCnt).ErrorName = "禁止勤務"
                    End If
                Next w_i
			Next w_Int
        End If

        If m_ChkStaffPattern = True Then
            '禁止職員パターン

            'エラー表示状況表示
            For w_Int = 1 To UBound(g_NotStaffPatternError2)
                For w_i = 1 To UBound(g_NotStaffPatternError2(w_Int).Data)
                    If g_NotStaffPatternError2(w_Int).Data(w_i).ErrorFlg = True And g_NotStaffPatternError2(w_Int).Data(w_i).ErrorDate >= m_ViewFrom And g_NotStaffPatternError2(w_Int).Data(w_i).ErrorDate <= m_ViewTo Then

                        w_intCnt = UBound(m_ErrorList) + 1
                        ReDim Preserve m_ErrorList(w_intCnt)

                        'エラーの場合
                        m_ErrorList(w_intCnt).StaffIdx = g_NotStaffPatternError2(w_Int).StaffIdx
                        m_ErrorList(w_intCnt).StaffName = g_NotStaffPatternError2(w_Int).StaffName
                        m_ErrorList(w_intCnt).Date_Renamed = g_NotStaffPatternError2(w_Int).Data(w_i).ErrorDate
                        m_ErrorList(w_intCnt).ErrorDetail = g_NotStaffPatternError2(w_Int).Data(w_i).ErrorPattern
                        m_ErrorList(w_intCnt).ColIdx = g_NotStaffPatternError2(w_Int).Data(w_i).ColIdx
                        'エラー項目
                        m_ErrorList(w_intCnt).ErrorName = "禁止職員パターン"
                    End If
                Next w_i
            Next w_Int
        End If

        If m_ChkGiryoType = True Then
            '経験区分の組み合わせチェック

            'エラー表示状況表示
            For w_Int = 1 To UBound(g_NotGiryoCheckError)
                For w_i = 1 To UBound(g_NotGiryoCheckError(w_Int).Data)
                    If g_NotGiryoCheckError(w_Int).Data(w_i).ErrorFlg = True And g_NotGiryoCheckError(w_Int).Data(w_i).ErrorDate >= m_ViewFrom And g_NotGiryoCheckError(w_Int).Data(w_i).ErrorDate <= m_ViewTo Then

                        w_intCnt = UBound(m_ErrorList) + 1
                        ReDim Preserve m_ErrorList(w_intCnt)

                        'エラーの場合
                        m_ErrorList(w_intCnt).StaffIdx = g_NotGiryoCheckError(w_Int).StaffIdx
                        m_ErrorList(w_intCnt).StaffName = g_NotGiryoCheckError(w_Int).StaffName
                        m_ErrorList(w_intCnt).Date_Renamed = g_NotGiryoCheckError(w_Int).Data(w_i).ErrorDate
                        m_ErrorList(w_intCnt).ErrorDetail = g_NotGiryoCheckError(w_Int).Data(w_i).ErrorPattern
                        m_ErrorList(w_intCnt).ColIdx = g_NotGiryoCheckError(w_Int).Data(w_i).ColIdx
                        'エラー項目
                        m_ErrorList(w_intCnt).ErrorName = "経験区分の組み合わせ"
                    End If
                Next w_i
            Next w_Int
        End If

        If m_RenzokuKinmuCheck = True Then
            '回数選択

            '各勤務毎（集計項目）のエラー情報表示
            Call Set_RenzokuCountErrInf()
        End If

        If m_AbsoluteKinmuCheck = True Then
            '必須勤務の組み合わせチェック

            'エラー表示状況表示
            For w_Int = 1 To UBound(g_NotAbsKinmuCheckError)
                For w_i = 1 To UBound(g_NotAbsKinmuCheckError(w_Int).Data)
                    If g_NotAbsKinmuCheckError(w_Int).Data(w_i).ErrorFlg = True And g_NotAbsKinmuCheckError(w_Int).Data(w_i).ErrorDate >= m_ViewFrom And g_NotAbsKinmuCheckError(w_Int).Data(w_i).ErrorDate <= m_ViewTo Then

                        w_intCnt = UBound(m_ErrorList) + 1
                        ReDim Preserve m_ErrorList(w_intCnt)

                        'エラーの場合
                        m_ErrorList(w_intCnt).StaffIdx = g_NotAbsKinmuCheckError(w_Int).StaffIdx
                        m_ErrorList(w_intCnt).StaffName = g_NotAbsKinmuCheckError(w_Int).StaffName
                        m_ErrorList(w_intCnt).Date_Renamed = g_NotAbsKinmuCheckError(w_Int).Data(w_i).ErrorDate
                        m_ErrorList(w_intCnt).ErrorDetail = g_NotAbsKinmuCheckError(w_Int).Data(w_i).ErrorPattern
                        m_ErrorList(w_intCnt).ColIdx = g_NotAbsKinmuCheckError(w_Int).Data(w_i).ColIdx
                        'エラー項目
                        m_ErrorList(w_intCnt).ErrorName = "必須勤務の組み合わせ"
                    End If
                Next w_i
            Next w_Int
        End If

		'日付＞表示順にソート
		Call SortData()

		'---リストに表示---
		With lvwErrorList
			For w_i = 1 To UBound(m_ErrorList)
                w_str = Format(m_ErrorList(w_i).Date_Renamed, "0000/00/00")
                w_str = Format(CDate(w_str), "M/d")
				
				'ﾏｽﾀｺｰﾄﾞ
				w_LItem = .Items.Add(w_str)
                If w_LItem.SubItems.Count > 1 Then
                    w_LItem.SubItems(1).Text = m_ErrorList(w_i).StaffName
                Else
                    w_LItem.SubItems.Insert(1, New ListViewItem.ListViewSubItem(Nothing, m_ErrorList(w_i).StaffName))
                End If

                If w_LItem.SubItems.Count > 2 Then
                    w_LItem.SubItems(2).Text = m_ErrorList(w_i).ErrorName
                Else
                    w_LItem.SubItems.Insert(2, New ListViewItem.ListViewSubItem(Nothing, m_ErrorList(w_i).ErrorName))
                End If

                If w_LItem.SubItems.Count > 3 Then
                    w_LItem.SubItems(3).Text = m_ErrorList(w_i).ErrorDetail
                Else
                    w_LItem.SubItems.Insert(3, New ListViewItem.ListViewSubItem(Nothing, m_ErrorList(w_i).ErrorDetail))
                End If
            Next w_i
			
			If UBound(m_ErrorList) > 0 Then
				'１番目のｱｲﾃﾑを選択状態に
                .Items.Item(0).Selected = True
                .FocusedItem = .Items.Item(0)
			End If
		End With
		
		Exit Sub
Set_ErrorInf: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
        End
	End Sub
	
	Private Sub Set_IntervalErrInf()
		On Error GoTo Set_IntervalErrInf
		Const W_SUBNAME As String = "NSK0000HN Set_IntervalErrInf"
		
		Dim w_Loop As Short
        Dim w_DataIndex As Short
		Dim w_Int As Short
		
        For w_Loop = 1 To UBound(g_KikanError2)
            For w_Int = 1 To UBound(g_KikanError2(w_Loop).InterValErr)
                If g_KikanError2(w_Loop).InterValErr(w_Int).ErrorFlg = True And g_KikanError2(w_Loop).InterValErr(w_Int).ErrorDate >= m_ViewFrom And g_KikanError2(w_Loop).InterValErr(w_Int).ErrorDate <= m_ViewTo Then

                    'エラー情報表示
                    w_DataIndex = UBound(m_ErrorList) + 1
                    ReDim Preserve m_ErrorList(w_DataIndex)

                    m_ErrorList(w_DataIndex).Date_Renamed = g_KikanError2(w_Loop).InterValErr(w_Int).ErrorDate
                    m_ErrorList(w_DataIndex).ErrorDetail = g_KikanError2(w_Loop).InterValErr(w_Int).ErrorName
                    m_ErrorList(w_DataIndex).StaffIdx = g_KikanError2(w_Loop).StaffIdx
                    m_ErrorList(w_DataIndex).StaffName = g_KikanError2(w_Loop).StaffName
                    '日付インデックス
                    m_ErrorList(w_DataIndex).ColIdx = g_KikanError2(w_Loop).InterValErr(w_Int).ColIdx
                    'エラー項目
                    m_ErrorList(w_DataIndex).ErrorName = "勤務／休日間隔"
                End If
            Next w_Int
        Next w_Loop
		
		Exit Sub
Set_IntervalErrInf: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
		End
	End Sub
	
	Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
		If lvwErrorList.Items.Count > 0 Then
            m_intSelIdx = lvwErrorList.FocusedItem.Index + 1
		Else
			m_intSelIdx = 0
		End If
        Call General.paSaveFieldWidth(lvwErrorList, General.G_STRMAINKEY2 & "\NSK0000H\", Me.Tag)
        Call General.paPutWindowPositon(Me, General.G_STRMAINKEY2 & "\NSK0000H\")
		
		Me.Close()
		
	End Sub
	
	Private Sub frmNSK0000HN_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo Form_Load
		Const W_SUBNAME As String = "NSK0000HN Form_Load"
		
        Dim w_Cnt As Short
        Dim clmX As ColumnHeader
        '2018/09/21 K.I Add Start-------------------------
        Dim w_Left As String
        Dim w_Top As String
        '2018/09/21 K.I Add End---------------------------

        ReDim m_ErrorList(0)
		
		If w_Cnt = 0 Then
        Else
            w_Cnt = 1
		End If
		
		'リストのヘッダ表示
		With lvwErrorList
			.Columns.Clear()
			'ﾍｯﾀﾞｰ作成
            clmX = .Columns.Add("", "日付", Integer.Parse(General.paTwipsTopixels(700)))
            clmX = .Columns.Add("", "氏　名", Integer.Parse(General.paTwipsTopixels(2500)))
            clmX = .Columns.Add("", "エラー項目", Integer.Parse(General.paTwipsTopixels(2437)))
            clmX = .Columns.Add("", "内　容", Integer.Parse(General.paTwipsTopixels(2000)))
            Call General.paSetFieldWidth(lvwErrorList, General.G_STRMAINKEY2 & "\NSK0000H", Me.Tag, False)
			lvwErrorList.Visible = True
		End With
		
		'エラー情報の表示
		Call Set_ErrorInf(w_Cnt)

        'ｳｨﾝﾄﾞｳの表示ﾎﾟｼﾞｼｮﾝを設定する
        '2018/09/21 K.I Upd Start------------------------------------------------------------------------------------------
        'レジストリ取得を削除
        'Call General.paGetWindowPositon(Me, General.G_STRMAINKEY2 & "\" & "NSK0000H\")
        '画面中央
        w_Left = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - CSng(Me.Width)) / 2)
        w_Top = CStr((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - CSng(Me.Height)) / 2)
        Me.SetBounds((CSng(w_Left)), (CSng(w_Top)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
        '2018/09/21 K.I Upd End--------------------------------------------------------------------------------------------

        Exit Sub
Form_Load: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
		End
	End Sub
	
	Private Sub lvwErrorList_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvwErrorList.DoubleClick
		If lvwErrorList.Items.Count > 0 Then
			Call cmdClose_Click(cmdClose, New System.EventArgs())
		End If
    End Sub

	'日付 /表示順に並び替え
	Private Sub SortData()
		On Error GoTo SortData
		Const W_SUBNAME As String = "NSK0000HN SortData"
		
		Dim w_Int As Short
		Dim w_Int2 As Short
		'ｿｰﾄ用配列の確保
		Dim w_WorkTbl As OutputType
		
		'職員人数分 繰り返し
		For w_Int = 1 To UBound(m_ErrorList)
			'(職員人数 - ｿｰﾄ終了人数) 繰り返し
			For w_Int2 = 1 To UBound(m_ErrorList) - w_Int
				'並び替え 実行 ?
				If (m_ErrorList(w_Int).Date_Renamed > m_ErrorList(w_Int + w_Int2).Date_Renamed) Then
					'表示順が後者よりも大きい場合 入れ替え
                    w_WorkTbl = m_ErrorList(w_Int)
                    m_ErrorList(w_Int) = m_ErrorList(w_Int + w_Int2)
                    m_ErrorList(w_Int + w_Int2) = w_WorkTbl
					
				ElseIf (m_ErrorList(w_Int).Date_Renamed = m_ErrorList(w_Int + w_Int2).Date_Renamed) And m_ErrorList(w_Int).StaffIdx > m_ErrorList(w_Int + w_Int2).StaffIdx Then 
					'表示順が後者と同じで、職員管理番号が大きい場合 入れ替え
                    w_WorkTbl = m_ErrorList(w_Int)
                    m_ErrorList(w_Int) = m_ErrorList(w_Int + w_Int2)
                    m_ErrorList(w_Int + w_Int2) = w_WorkTbl
				End If
			Next w_Int2
			'待ち処理
            Application.DoEvents()
		Next w_Int
		
		Exit Sub
SortData: 
        Call General.paTrpMsg(Str(Err.Number), W_SUBNAME)
		End
	End Sub
End Class