<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNSK0000HM
#Region "Windows �t�H�[�� �f�U�C�i�ɂ���Đ������ꂽ�R�[�h "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'���̌Ăяo���́AWindows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
		InitializeComponent()
	End Sub
	'Form �́A�R���|�[�l���g�ꗗ�Ɍ㏈�������s���邽�߂� dispose ���I�[�o�[���C�h���܂��B
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents pgbProcess As System.Windows.Forms.ProgressBar
	Public WithEvents lblCount As System.Windows.Forms.Label
	Public WithEvents lblSyori As System.Windows.Forms.Label
	'����: �ȉ��̃v���V�[�W���� Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
	'Windows �t�H�[�� �f�U�C�i���g���ĕύX�ł��܂��B
	'�R�[�h �G�f�B�^���g�p���āA�ύX���Ȃ��ł��������B
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.pgbProcess = New System.Windows.Forms.ProgressBar
        Me.lblCount = New System.Windows.Forms.Label
        Me.lblSyori = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'pgbProcess
        '
        Me.pgbProcess.Location = New System.Drawing.Point(40, 80)
        Me.pgbProcess.Name = "pgbProcess"
        Me.pgbProcess.Size = New System.Drawing.Size(243, 25)
        Me.pgbProcess.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.pgbProcess.TabIndex = 2
        '
        'lblCount
        '
        Me.lblCount.BackColor = System.Drawing.Color.Transparent
        Me.lblCount.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCount.Font = New System.Drawing.Font("�l�r �S�V�b�N", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblCount.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblCount.Location = New System.Drawing.Point(0, 56)
        Me.lblCount.Name = "lblCount"
        Me.lblCount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCount.Size = New System.Drawing.Size(313, 20)
        Me.lblCount.TabIndex = 1
        Me.lblCount.Text = "999 / 999"
        Me.lblCount.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblSyori
        '
        Me.lblSyori.BackColor = System.Drawing.Color.Transparent
        Me.lblSyori.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSyori.Font = New System.Drawing.Font("�l�r �S�V�b�N", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSyori.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblSyori.Location = New System.Drawing.Point(0, 16)
        Me.lblSyori.Name = "lblSyori"
        Me.lblSyori.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSyori.Size = New System.Drawing.Size(313, 25)
        Me.lblSyori.TabIndex = 0
        Me.lblSyori.Text = "�Ζ��v��f�[�^�擾��..."
        Me.lblSyori.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmNSK0000HM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(313, 123)
        Me.ControlBox = False
        Me.Controls.Add(Me.pgbProcess)
        Me.Controls.Add(Me.lblCount)
        Me.Controls.Add(Me.lblSyori)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("�l�r �S�V�b�N", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Location = New System.Drawing.Point(96, 293)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNSK0000HM"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "������..."
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class