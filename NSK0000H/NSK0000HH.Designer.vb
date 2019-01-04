<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNSK0000HH
#Region "Windows フォーム デザイナによって生成されたコード "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'この呼び出しは、Windows フォーム デザイナで必要です。
		InitializeComponent()
	End Sub
	'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows フォーム デザイナで必要です。
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents cmdSchedExit As System.Windows.Forms.Button
    Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents HscKinmu As System.Windows.Forms.HScrollBar
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents Picture1 As System.Windows.Forms.PictureBox
	Public WithEvents Lst_SortList As System.Windows.Forms.ListBox
	Public WithEvents cmdSchedStart As System.Windows.Forms.Button
	Public WithEvents prbSchedProcess As System.Windows.Forms.ProgressBar
	Public WithEvents lblSchedMessage As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更できます。
    'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNSK0000HH))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdSchedExit = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.Frame2 = New System.Windows.Forms.GroupBox
        Me.HscKinmu = New System.Windows.Forms.HScrollBar
        Me._cmdKinmu_5 = New System.Windows.Forms.Button
        Me._cmdKinmu_4 = New System.Windows.Forms.Button
        Me._cmdKinmu_3 = New System.Windows.Forms.Button
        Me._cmdKinmu_2 = New System.Windows.Forms.Button
        Me._cmdKinmu_1 = New System.Windows.Forms.Button
        Me._cmdKinmu_0 = New System.Windows.Forms.Button
        Me.Picture1 = New System.Windows.Forms.PictureBox
        Me.Lst_SortList = New System.Windows.Forms.ListBox
        Me.cmdSchedStart = New System.Windows.Forms.Button
        Me.prbSchedProcess = New System.Windows.Forms.ProgressBar
        Me.lblSchedMessage = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.imdDate = New CustomText.NiszMaskedText(Me.components)
        Me.Frame2.SuspendLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdSchedExit
        '
        Me.cmdSchedExit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSchedExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSchedExit.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmdSchedExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSchedExit.Location = New System.Drawing.Point(232, 264)
        Me.cmdSchedExit.Name = "cmdSchedExit"
        Me.cmdSchedExit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSchedExit.Size = New System.Drawing.Size(57, 25)
        Me.cmdSchedExit.TabIndex = 18
        Me.cmdSchedExit.Text = "終了(&E)"
        Me.cmdSchedExit.UseVisualStyleBackColor = False
        Me.cmdSchedExit.Visible = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.Location = New System.Drawing.Point(227, 200)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(100, 33)
        Me.cmdCancel.TabIndex = 8
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.HscKinmu)
        Me.Frame2.Controls.Add(Me._cmdKinmu_5)
        Me.Frame2.Controls.Add(Me._cmdKinmu_4)
        Me.Frame2.Controls.Add(Me._cmdKinmu_3)
        Me.Frame2.Controls.Add(Me._cmdKinmu_2)
        Me.Frame2.Controls.Add(Me._cmdKinmu_1)
        Me.Frame2.Controls.Add(Me._cmdKinmu_0)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(8, 80)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(321, 81)
        Me.Frame2.TabIndex = 0
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "割当勤務"
        '
        'HscKinmu
        '
        Me.HscKinmu.Cursor = System.Windows.Forms.Cursors.Default
        Me.HscKinmu.LargeChange = 1
        Me.HscKinmu.Location = New System.Drawing.Point(16, 62)
        Me.HscKinmu.Maximum = 32767
        Me.HscKinmu.Name = "HscKinmu"
        Me.HscKinmu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HscKinmu.Size = New System.Drawing.Size(290, 17)
        Me.HscKinmu.TabIndex = 9
        '
        '_cmdKinmu_5
        '
        Me._cmdKinmu_5.Location = New System.Drawing.Point(256, 22)
        Me._cmdKinmu_5.Name = "_cmdKinmu_5"
        Me._cmdKinmu_5.Size = New System.Drawing.Size(49, 41)
        Me._cmdKinmu_5.TabIndex = 6
        Me._cmdKinmu_5.UseVisualStyleBackColor = True
        '
        '_cmdKinmu_4
        '
        Me._cmdKinmu_4.Location = New System.Drawing.Point(208, 22)
        Me._cmdKinmu_4.Name = "_cmdKinmu_4"
        Me._cmdKinmu_4.Size = New System.Drawing.Size(49, 41)
        Me._cmdKinmu_4.TabIndex = 5
        Me._cmdKinmu_4.UseVisualStyleBackColor = True
        '
        '_cmdKinmu_3
        '
        Me._cmdKinmu_3.Location = New System.Drawing.Point(160, 22)
        Me._cmdKinmu_3.Name = "_cmdKinmu_3"
        Me._cmdKinmu_3.Size = New System.Drawing.Size(49, 41)
        Me._cmdKinmu_3.TabIndex = 4
        Me._cmdKinmu_3.UseVisualStyleBackColor = True
        '
        '_cmdKinmu_2
        '
        Me._cmdKinmu_2.Location = New System.Drawing.Point(112, 22)
        Me._cmdKinmu_2.Name = "_cmdKinmu_2"
        Me._cmdKinmu_2.Size = New System.Drawing.Size(49, 41)
        Me._cmdKinmu_2.TabIndex = 3
        Me._cmdKinmu_2.UseVisualStyleBackColor = True
        '
        '_cmdKinmu_1
        '
        Me._cmdKinmu_1.Location = New System.Drawing.Point(64, 22)
        Me._cmdKinmu_1.Name = "_cmdKinmu_1"
        Me._cmdKinmu_1.Size = New System.Drawing.Size(49, 41)
        Me._cmdKinmu_1.TabIndex = 2
        Me._cmdKinmu_1.UseVisualStyleBackColor = True
        '
        '_cmdKinmu_0
        '
        Me._cmdKinmu_0.Location = New System.Drawing.Point(16, 22)
        Me._cmdKinmu_0.Name = "_cmdKinmu_0"
        Me._cmdKinmu_0.Size = New System.Drawing.Size(49, 41)
        Me._cmdKinmu_0.TabIndex = 1
        Me._cmdKinmu_0.UseVisualStyleBackColor = True
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Image = CType(resources.GetObject("Picture1.Image"), System.Drawing.Image)
        Me.Picture1.Location = New System.Drawing.Point(8, 8)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(33, 33)
        Me.Picture1.TabIndex = 12
        Me.Picture1.TabStop = False
        '
        'Lst_SortList
        '
        Me.Lst_SortList.BackColor = System.Drawing.SystemColors.Window
        Me.Lst_SortList.Cursor = System.Windows.Forms.Cursors.Default
        Me.Lst_SortList.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Lst_SortList.ItemHeight = 15
        Me.Lst_SortList.Location = New System.Drawing.Point(8, 200)
        Me.Lst_SortList.Name = "Lst_SortList"
        Me.Lst_SortList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Lst_SortList.Size = New System.Drawing.Size(49, 19)
        Me.Lst_SortList.Sorted = True
        Me.Lst_SortList.TabIndex = 11
        Me.Lst_SortList.Visible = False
        '
        'cmdSchedStart
        '
        Me.cmdSchedStart.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSchedStart.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSchedStart.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmdSchedStart.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSchedStart.Image = CType(resources.GetObject("cmdSchedStart.Image"), System.Drawing.Image)
        Me.cmdSchedStart.Location = New System.Drawing.Point(112, 200)
        Me.cmdSchedStart.Name = "cmdSchedStart"
        Me.cmdSchedStart.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSchedStart.Size = New System.Drawing.Size(100, 33)
        Me.cmdSchedStart.TabIndex = 7
        Me.cmdSchedStart.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdSchedStart.UseVisualStyleBackColor = False
        '
        'prbSchedProcess
        '
        Me.prbSchedProcess.Location = New System.Drawing.Point(8, 168)
        Me.prbSchedProcess.Name = "prbSchedProcess"
        Me.prbSchedProcess.Size = New System.Drawing.Size(321, 25)
        Me.prbSchedProcess.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.prbSchedProcess.TabIndex = 10
        '
        'lblSchedMessage
        '
        Me.lblSchedMessage.BackColor = System.Drawing.Color.Transparent
        Me.lblSchedMessage.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSchedMessage.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSchedMessage.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSchedMessage.Location = New System.Drawing.Point(112, 264)
        Me.lblSchedMessage.Name = "lblSchedMessage"
        Me.lblSchedMessage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSchedMessage.Size = New System.Drawing.Size(73, 25)
        Me.lblSchedMessage.TabIndex = 17
        Me.lblSchedMessage.Text = "開始ボタンをクリックしてください"
        Me.lblSchedMessage.Visible = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(16, 53)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(71, 17)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "割当開始"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(47, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(225, 17)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "勤務表作成支援システム"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(57, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(278, 17)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Copyright(C)2002 Nihon InterSystems Co.,LTD."
        '
        'imdDate
        '
        Me.imdDate.AutoSize = False
        Me.imdDate.EnabledBackColor = System.Drawing.SystemColors.Window
        Me.imdDate.Format = "D"
        Me.imdDate.HighlightText = True
        Me.imdDate.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.imdDate.InsertKeyMode = System.Windows.Forms.InsertKeyMode.Overwrite
        Me.imdDate.Location = New System.Drawing.Point(93, 50)
        Me.imdDate.Mask = "0000/00/00"
        Me.imdDate.MaxDate = "20991231"
        Me.imdDate.MinDate = "18680101"
        Me.imdDate.Name = "imdDate"
        Me.imdDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.imdDate.Size = New System.Drawing.Size(89, 22)
        Me.imdDate.TabIndex = 19
        '
        'frmNSK0000HH
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(337, 250)
        Me.ControlBox = False
        Me.Controls.Add(Me.imdDate)
        Me.Controls.Add(Me.cmdSchedExit)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.Lst_SortList)
        Me.Controls.Add(Me.cmdSchedStart)
        Me.Controls.Add(Me.prbSchedProcess)
        Me.Controls.Add(Me.lblSchedMessage)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(323, 265)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNSK0000HH"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Tag = "NSK0000HH"
        Me.Text = "自動シミュレーション"
        Me.Frame2.ResumeLayout(False)
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents _cmdKinmu_5 As System.Windows.Forms.Button
    Friend WithEvents _cmdKinmu_4 As System.Windows.Forms.Button
    Friend WithEvents _cmdKinmu_3 As System.Windows.Forms.Button
    Friend WithEvents _cmdKinmu_2 As System.Windows.Forms.Button
    Friend WithEvents _cmdKinmu_1 As System.Windows.Forms.Button
    Friend WithEvents _cmdKinmu_0 As System.Windows.Forms.Button
    Friend WithEvents imdDate As CustomText.NiszMaskedText
#End Region
End Class