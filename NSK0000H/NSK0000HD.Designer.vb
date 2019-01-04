<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNSK0000HD
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
	Public WithEvents chkTaisyoOnly As System.Windows.Forms.CheckBox
	Public WithEvents cboKensakuKinmu As System.Windows.Forms.ComboBox
	Public WithEvents cboChikanKinmu As System.Windows.Forms.ComboBox
	Public WithEvents cboTaisyo As System.Windows.Forms.ComboBox
	Public WithEvents cmd_Next_Chikan As System.Windows.Forms.Button
	Public WithEvents cmdEnd As System.Windows.Forms.Button
	Public WithEvents cmd_Select_AllChikan As System.Windows.Forms.Button
	Public WithEvents cboRiyu As System.Windows.Forms.ComboBox
	Public WithEvents _lblKomoku_0 As System.Windows.Forms.Label
	Public WithEvents _lblKomoku_1 As System.Windows.Forms.Label
	Public WithEvents _lblKomoku_2 As System.Windows.Forms.Label
	Public WithEvents _lblKomoku_3 As System.Windows.Forms.Label
    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更できます。
    'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNSK0000HD))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkTaisyoOnly = New System.Windows.Forms.CheckBox
        Me.cboKensakuKinmu = New System.Windows.Forms.ComboBox
        Me.cboChikanKinmu = New System.Windows.Forms.ComboBox
        Me.cboTaisyo = New System.Windows.Forms.ComboBox
        Me.cmd_Next_Chikan = New System.Windows.Forms.Button
        Me.cmdEnd = New System.Windows.Forms.Button
        Me.cmd_Select_AllChikan = New System.Windows.Forms.Button
        Me.cboRiyu = New System.Windows.Forms.ComboBox
        Me._lblKomoku_0 = New System.Windows.Forms.Label
        Me._lblKomoku_1 = New System.Windows.Forms.Label
        Me._lblKomoku_2 = New System.Windows.Forms.Label
        Me._lblKomoku_3 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'chkTaisyoOnly
        '
        Me.chkTaisyoOnly.BackColor = System.Drawing.SystemColors.Control
        Me.chkTaisyoOnly.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTaisyoOnly.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTaisyoOnly.Location = New System.Drawing.Point(14, 139)
        Me.chkTaisyoOnly.Name = "chkTaisyoOnly"
        Me.chkTaisyoOnly.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTaisyoOnly.Size = New System.Drawing.Size(217, 19)
        Me.chkTaisyoOnly.TabIndex = 8
        Me.chkTaisyoOnly.Text = "対象のみで検索する(&O)"
        Me.chkTaisyoOnly.UseVisualStyleBackColor = False
        '
        'cboKensakuKinmu
        '
        Me.cboKensakuKinmu.BackColor = System.Drawing.SystemColors.Window
        Me.cboKensakuKinmu.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboKensakuKinmu.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKensakuKinmu.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboKensakuKinmu.Items.AddRange(New Object() {"未割当", "日勤(・)", "深夜(○)", "準夜(△)", "週休(週)"})
        Me.cboKensakuKinmu.Location = New System.Drawing.Point(186, 10)
        Me.cboKensakuKinmu.Name = "cboKensakuKinmu"
        Me.cboKensakuKinmu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboKensakuKinmu.Size = New System.Drawing.Size(94, 23)
        Me.cboKensakuKinmu.TabIndex = 1
        '
        'cboChikanKinmu
        '
        Me.cboChikanKinmu.BackColor = System.Drawing.SystemColors.Window
        Me.cboChikanKinmu.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboChikanKinmu.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChikanKinmu.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboChikanKinmu.Items.AddRange(New Object() {"未割当", "日勤(・)", "深夜(○)", "準夜(△)", "週休(週)"})
        Me.cboChikanKinmu.Location = New System.Drawing.Point(186, 38)
        Me.cboChikanKinmu.Name = "cboChikanKinmu"
        Me.cboChikanKinmu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboChikanKinmu.Size = New System.Drawing.Size(94, 23)
        Me.cboChikanKinmu.TabIndex = 3
        '
        'cboTaisyo
        '
        Me.cboTaisyo.BackColor = System.Drawing.SystemColors.Window
        Me.cboTaisyo.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboTaisyo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTaisyo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboTaisyo.Items.AddRange(New Object() {"通常", "要請", "希望", "再掲"})
        Me.cboTaisyo.Location = New System.Drawing.Point(105, 104)
        Me.cboTaisyo.Name = "cboTaisyo"
        Me.cboTaisyo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboTaisyo.Size = New System.Drawing.Size(94, 23)
        Me.cboTaisyo.TabIndex = 7
        '
        'cmd_Next_Chikan
        '
        Me.cmd_Next_Chikan.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_Next_Chikan.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_Next_Chikan.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_Next_Chikan.Location = New System.Drawing.Point(320, 10)
        Me.cmd_Next_Chikan.Name = "cmd_Next_Chikan"
        Me.cmd_Next_Chikan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_Next_Chikan.Size = New System.Drawing.Size(100, 33)
        Me.cmd_Next_Chikan.TabIndex = 9
        Me.cmd_Next_Chikan.Text = "次を検索(&F)"
        Me.cmd_Next_Chikan.UseVisualStyleBackColor = False
        '
        'cmdEnd
        '
        Me.cmdEnd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEnd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdEnd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEnd.Image = CType(resources.GetObject("cmdEnd.Image"), System.Drawing.Image)
        Me.cmdEnd.Location = New System.Drawing.Point(320, 120)
        Me.cmdEnd.Name = "cmdEnd"
        Me.cmdEnd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdEnd.Size = New System.Drawing.Size(100, 33)
        Me.cmdEnd.TabIndex = 11
        Me.cmdEnd.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdEnd.UseVisualStyleBackColor = False
        '
        'cmd_Select_AllChikan
        '
        Me.cmd_Select_AllChikan.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_Select_AllChikan.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_Select_AllChikan.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_Select_AllChikan.Location = New System.Drawing.Point(320, 56)
        Me.cmd_Select_AllChikan.Name = "cmd_Select_AllChikan"
        Me.cmd_Select_AllChikan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_Select_AllChikan.Size = New System.Drawing.Size(100, 33)
        Me.cmd_Select_AllChikan.TabIndex = 10
        Me.cmd_Select_AllChikan.Text = "選択(&S)"
        Me.cmd_Select_AllChikan.UseVisualStyleBackColor = False
        '
        'cboRiyu
        '
        Me.cboRiyu.BackColor = System.Drawing.SystemColors.Window
        Me.cboRiyu.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboRiyu.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRiyu.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboRiyu.Items.AddRange(New Object() {"通常", "要請", "希望", "再掲"})
        Me.cboRiyu.Location = New System.Drawing.Point(186, 67)
        Me.cboRiyu.Name = "cboRiyu"
        Me.cboRiyu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboRiyu.Size = New System.Drawing.Size(94, 23)
        Me.cboRiyu.TabIndex = 5
        '
        '_lblKomoku_0
        '
        Me._lblKomoku_0.BackColor = System.Drawing.Color.Transparent
        Me._lblKomoku_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblKomoku_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lblKomoku_0.Location = New System.Drawing.Point(14, 13)
        Me._lblKomoku_0.Name = "_lblKomoku_0"
        Me._lblKomoku_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblKomoku_0.Size = New System.Drawing.Size(169, 17)
        Me._lblKomoku_0.TabIndex = 0
        Me._lblKomoku_0.Text = "検索する勤務(&N)　　："
        '
        '_lblKomoku_1
        '
        Me._lblKomoku_1.BackColor = System.Drawing.Color.Transparent
        Me._lblKomoku_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblKomoku_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lblKomoku_1.Location = New System.Drawing.Point(14, 42)
        Me._lblKomoku_1.Name = "_lblKomoku_1"
        Me._lblKomoku_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblKomoku_1.Size = New System.Drawing.Size(169, 17)
        Me._lblKomoku_1.TabIndex = 2
        Me._lblKomoku_1.Text = "置換後の勤務(&E)　　："
        '
        '_lblKomoku_2
        '
        Me._lblKomoku_2.BackColor = System.Drawing.Color.Transparent
        Me._lblKomoku_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblKomoku_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lblKomoku_2.Location = New System.Drawing.Point(14, 108)
        Me._lblKomoku_2.Name = "_lblKomoku_2"
        Me._lblKomoku_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblKomoku_2.Size = New System.Drawing.Size(93, 17)
        Me._lblKomoku_2.TabIndex = 6
        Me._lblKomoku_2.Text = "対象(&L)  ："
        '
        '_lblKomoku_3
        '
        Me._lblKomoku_3.BackColor = System.Drawing.Color.Transparent
        Me._lblKomoku_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblKomoku_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lblKomoku_3.Location = New System.Drawing.Point(14, 70)
        Me._lblKomoku_3.Name = "_lblKomoku_3"
        Me._lblKomoku_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblKomoku_3.Size = New System.Drawing.Size(169, 17)
        Me._lblKomoku_3.TabIndex = 4
        Me._lblKomoku_3.Text = "置換後の理由区分(&K)："
        '
        'frmNSK0000HD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(431, 168)
        Me.Controls.Add(Me.chkTaisyoOnly)
        Me.Controls.Add(Me.cboKensakuKinmu)
        Me.Controls.Add(Me.cboChikanKinmu)
        Me.Controls.Add(Me.cboTaisyo)
        Me.Controls.Add(Me.cmd_Next_Chikan)
        Me.Controls.Add(Me.cmdEnd)
        Me.Controls.Add(Me.cmd_Select_AllChikan)
        Me.Controls.Add(Me.cboRiyu)
        Me.Controls.Add(Me._lblKomoku_0)
        Me.Controls.Add(Me._lblKomoku_1)
        Me.Controls.Add(Me._lblKomoku_2)
        Me.Controls.Add(Me._lblKomoku_3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(246, 223)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNSK0000HD"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Tag = "NSK0000HD"
        Me.Text = "検索"
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class