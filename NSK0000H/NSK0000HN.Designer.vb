<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNSK0000HN
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
	Public WithEvents cmdClose As System.Windows.Forms.Button
	Public WithEvents lvwErrorList As System.Windows.Forms.ListView
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNSK0000HN))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdClose = New System.Windows.Forms.Button
        Me.lvwErrorList = New System.Windows.Forms.ListView
        Me.SuspendLayout()
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Image = CType(resources.GetObject("cmdClose.Image"), System.Drawing.Image)
        Me.cmdClose.Location = New System.Drawing.Point(416, 272)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(100, 33)
        Me.cmdClose.TabIndex = 1
        Me.cmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'lvwErrorList
        '
        Me.lvwErrorList.BackColor = System.Drawing.SystemColors.Window
        Me.lvwErrorList.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvwErrorList.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lvwErrorList.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lvwErrorList.FullRowSelect = True
        Me.lvwErrorList.HideSelection = False
        Me.lvwErrorList.Location = New System.Drawing.Point(8, 8)
        Me.lvwErrorList.Name = "lvwErrorList"
        Me.lvwErrorList.Size = New System.Drawing.Size(513, 257)
        Me.lvwErrorList.TabIndex = 0
        Me.lvwErrorList.UseCompatibleStateImageBehavior = False
        Me.lvwErrorList.View = System.Windows.Forms.View.Details
        '
        'frmNSK0000HN
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(529, 313)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.lvwErrorList)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Location = New System.Drawing.Point(3, 18)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNSK0000HN"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.Tag = "NSK0000HN"
        Me.Text = "勤務情報"
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class