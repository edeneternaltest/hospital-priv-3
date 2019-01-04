<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNSK0000HG
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
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents _optSort_5 As System.Windows.Forms.RadioButton
	Public WithEvents _optSort_4 As System.Windows.Forms.RadioButton
	Public WithEvents _cmbSortKey_2 As System.Windows.Forms.ComboBox
	Public WithEvents _fraSort_2 As System.Windows.Forms.GroupBox
	Public WithEvents _optSort_3 As System.Windows.Forms.RadioButton
	Public WithEvents _optSort_2 As System.Windows.Forms.RadioButton
	Public WithEvents _cmbSortKey_1 As System.Windows.Forms.ComboBox
	Public WithEvents _fraSort_1 As System.Windows.Forms.GroupBox
	Public WithEvents _optSort_1 As System.Windows.Forms.RadioButton
	Public WithEvents _optSort_0 As System.Windows.Forms.RadioButton
	Public WithEvents _cmbSortKey_0 As System.Windows.Forms.ComboBox
	Public WithEvents _fraSort_0 As System.Windows.Forms.GroupBox
    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更できます。
    'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNSK0000HG))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me._fraSort_2 = New System.Windows.Forms.GroupBox
        Me._optSort_5 = New System.Windows.Forms.RadioButton
        Me._optSort_4 = New System.Windows.Forms.RadioButton
        Me._cmbSortKey_2 = New System.Windows.Forms.ComboBox
        Me._fraSort_1 = New System.Windows.Forms.GroupBox
        Me._optSort_3 = New System.Windows.Forms.RadioButton
        Me._optSort_2 = New System.Windows.Forms.RadioButton
        Me._cmbSortKey_1 = New System.Windows.Forms.ComboBox
        Me._fraSort_0 = New System.Windows.Forms.GroupBox
        Me._optSort_1 = New System.Windows.Forms.RadioButton
        Me._optSort_0 = New System.Windows.Forms.RadioButton
        Me._cmbSortKey_0 = New System.Windows.Forms.ComboBox
        Me._fraSort_2.SuspendLayout()
        Me._fraSort_1.SuspendLayout()
        Me._fraSort_0.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.Location = New System.Drawing.Point(275, 64)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(100, 33)
        Me.cmdCancel.TabIndex = 13
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Image = CType(resources.GetObject("cmdOK.Image"), System.Drawing.Image)
        Me.cmdOK.Location = New System.Drawing.Point(275, 16)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(100, 33)
        Me.cmdOK.TabIndex = 12
        Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        '_fraSort_2
        '
        Me._fraSort_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraSort_2.Controls.Add(Me._optSort_5)
        Me._fraSort_2.Controls.Add(Me._optSort_4)
        Me._fraSort_2.Controls.Add(Me._cmbSortKey_2)
        Me._fraSort_2.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me._fraSort_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraSort_2.Location = New System.Drawing.Point(11, 159)
        Me._fraSort_2.Name = "_fraSort_2"
        Me._fraSort_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraSort_2.Size = New System.Drawing.Size(253, 67)
        Me._fraSort_2.TabIndex = 8
        Me._fraSort_2.TabStop = False
        Me._fraSort_2.Text = "３番目のソート"
        '
        '_optSort_5
        '
        Me._optSort_5.BackColor = System.Drawing.SystemColors.Control
        Me._optSort_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._optSort_5.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me._optSort_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optSort_5.Location = New System.Drawing.Point(182, 40)
        Me._optSort_5.Name = "_optSort_5"
        Me._optSort_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optSort_5.Size = New System.Drawing.Size(60, 20)
        Me._optSort_5.TabIndex = 11
        Me._optSort_5.TabStop = True
        Me._optSort_5.Text = "降順"
        Me._optSort_5.UseVisualStyleBackColor = False
        '
        '_optSort_4
        '
        Me._optSort_4.BackColor = System.Drawing.SystemColors.Control
        Me._optSort_4.Checked = True
        Me._optSort_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._optSort_4.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me._optSort_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optSort_4.Location = New System.Drawing.Point(182, 20)
        Me._optSort_4.Name = "_optSort_4"
        Me._optSort_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optSort_4.Size = New System.Drawing.Size(60, 20)
        Me._optSort_4.TabIndex = 10
        Me._optSort_4.TabStop = True
        Me._optSort_4.Text = "昇順"
        Me._optSort_4.UseVisualStyleBackColor = False
        '
        '_cmbSortKey_2
        '
        Me._cmbSortKey_2.BackColor = System.Drawing.SystemColors.Window
        Me._cmbSortKey_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmbSortKey_2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cmbSortKey_2.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me._cmbSortKey_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._cmbSortKey_2.Location = New System.Drawing.Point(11, 26)
        Me._cmbSortKey_2.Name = "_cmbSortKey_2"
        Me._cmbSortKey_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmbSortKey_2.Size = New System.Drawing.Size(153, 23)
        Me._cmbSortKey_2.TabIndex = 9
        '
        '_fraSort_1
        '
        Me._fraSort_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraSort_1.Controls.Add(Me._optSort_3)
        Me._fraSort_1.Controls.Add(Me._optSort_2)
        Me._fraSort_1.Controls.Add(Me._cmbSortKey_1)
        Me._fraSort_1.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me._fraSort_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraSort_1.Location = New System.Drawing.Point(11, 83)
        Me._fraSort_1.Name = "_fraSort_1"
        Me._fraSort_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraSort_1.Size = New System.Drawing.Size(253, 67)
        Me._fraSort_1.TabIndex = 4
        Me._fraSort_1.TabStop = False
        Me._fraSort_1.Text = "２番目のソート"
        '
        '_optSort_3
        '
        Me._optSort_3.BackColor = System.Drawing.SystemColors.Control
        Me._optSort_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._optSort_3.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me._optSort_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optSort_3.Location = New System.Drawing.Point(182, 40)
        Me._optSort_3.Name = "_optSort_3"
        Me._optSort_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optSort_3.Size = New System.Drawing.Size(60, 20)
        Me._optSort_3.TabIndex = 7
        Me._optSort_3.TabStop = True
        Me._optSort_3.Text = "降順"
        Me._optSort_3.UseVisualStyleBackColor = False
        '
        '_optSort_2
        '
        Me._optSort_2.BackColor = System.Drawing.SystemColors.Control
        Me._optSort_2.Checked = True
        Me._optSort_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optSort_2.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me._optSort_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optSort_2.Location = New System.Drawing.Point(182, 20)
        Me._optSort_2.Name = "_optSort_2"
        Me._optSort_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optSort_2.Size = New System.Drawing.Size(60, 20)
        Me._optSort_2.TabIndex = 6
        Me._optSort_2.TabStop = True
        Me._optSort_2.Text = "昇順"
        Me._optSort_2.UseVisualStyleBackColor = False
        '
        '_cmbSortKey_1
        '
        Me._cmbSortKey_1.BackColor = System.Drawing.SystemColors.Window
        Me._cmbSortKey_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmbSortKey_1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cmbSortKey_1.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me._cmbSortKey_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._cmbSortKey_1.Location = New System.Drawing.Point(11, 26)
        Me._cmbSortKey_1.Name = "_cmbSortKey_1"
        Me._cmbSortKey_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmbSortKey_1.Size = New System.Drawing.Size(153, 23)
        Me._cmbSortKey_1.TabIndex = 5
        '
        '_fraSort_0
        '
        Me._fraSort_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraSort_0.Controls.Add(Me._optSort_1)
        Me._fraSort_0.Controls.Add(Me._optSort_0)
        Me._fraSort_0.Controls.Add(Me._cmbSortKey_0)
        Me._fraSort_0.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me._fraSort_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraSort_0.Location = New System.Drawing.Point(11, 8)
        Me._fraSort_0.Name = "_fraSort_0"
        Me._fraSort_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraSort_0.Size = New System.Drawing.Size(253, 67)
        Me._fraSort_0.TabIndex = 0
        Me._fraSort_0.TabStop = False
        Me._fraSort_0.Text = "１番目のソート"
        '
        '_optSort_1
        '
        Me._optSort_1.BackColor = System.Drawing.SystemColors.Control
        Me._optSort_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optSort_1.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me._optSort_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optSort_1.Location = New System.Drawing.Point(182, 40)
        Me._optSort_1.Name = "_optSort_1"
        Me._optSort_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optSort_1.Size = New System.Drawing.Size(60, 20)
        Me._optSort_1.TabIndex = 3
        Me._optSort_1.TabStop = True
        Me._optSort_1.Text = "降順"
        Me._optSort_1.UseVisualStyleBackColor = False
        '
        '_optSort_0
        '
        Me._optSort_0.BackColor = System.Drawing.SystemColors.Control
        Me._optSort_0.Checked = True
        Me._optSort_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optSort_0.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me._optSort_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optSort_0.Location = New System.Drawing.Point(182, 20)
        Me._optSort_0.Name = "_optSort_0"
        Me._optSort_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optSort_0.Size = New System.Drawing.Size(60, 20)
        Me._optSort_0.TabIndex = 2
        Me._optSort_0.TabStop = True
        Me._optSort_0.Text = "昇順"
        Me._optSort_0.UseVisualStyleBackColor = False
        '
        '_cmbSortKey_0
        '
        Me._cmbSortKey_0.BackColor = System.Drawing.SystemColors.Window
        Me._cmbSortKey_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmbSortKey_0.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cmbSortKey_0.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me._cmbSortKey_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._cmbSortKey_0.Location = New System.Drawing.Point(11, 26)
        Me._cmbSortKey_0.Name = "_cmbSortKey_0"
        Me._cmbSortKey_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmbSortKey_0.Size = New System.Drawing.Size(153, 23)
        Me._cmbSortKey_0.TabIndex = 1
        '
        'frmNSK0000HG
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(380, 234)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me._fraSort_2)
        Me.Controls.Add(Me._fraSort_1)
        Me.Controls.Add(Me._fraSort_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Location = New System.Drawing.Point(299, 158)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNSK0000HG"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Tag = "NSK0000HG"
        Me.Text = "並び替え"
        Me._fraSort_2.ResumeLayout(False)
        Me._fraSort_1.ResumeLayout(False)
        Me._fraSort_0.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class