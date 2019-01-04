<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNSK0000HO
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使用して変更できます。  
    'コード エディタを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNSK0000HO))
        Me.lstEventList = New System.Windows.Forms.ListView
        Me.lstEventStaff = New System.Windows.Forms.ListView
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblEventName = New System.Windows.Forms.Label
        Me.cmdClose = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lstEventList
        '
        Me.lstEventList.GridLines = True
        Me.lstEventList.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lstEventList.Location = New System.Drawing.Point(28, 33)
        Me.lstEventList.Name = "lstEventList"
        Me.lstEventList.Size = New System.Drawing.Size(472, 140)
        Me.lstEventList.TabIndex = 1
        Me.lstEventList.UseCompatibleStateImageBehavior = False
        Me.lstEventList.View = System.Windows.Forms.View.Details
        '
        'lstEventStaff
        '
        Me.lstEventStaff.GridLines = True
        Me.lstEventStaff.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lstEventStaff.Location = New System.Drawing.Point(28, 204)
        Me.lstEventStaff.Name = "lstEventStaff"
        Me.lstEventStaff.Size = New System.Drawing.Size(358, 228)
        Me.lstEventStaff.TabIndex = 4
        Me.lstEventStaff.UseCompatibleStateImageBehavior = False
        Me.lstEventStaff.View = System.Windows.Forms.View.Details
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(28, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(103, 15)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "＜行事一覧＞"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(41, 183)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(87, 15)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "＜参加者＞"
        '
        'lblEventName
        '
        Me.lblEventName.AutoSize = True
        Me.lblEventName.Location = New System.Drawing.Point(150, 183)
        Me.lblEventName.Name = "lblEventName"
        Me.lblEventName.Size = New System.Drawing.Size(55, 15)
        Me.lblEventName.TabIndex = 3
        Me.lblEventName.Text = "研修Ⅲ"
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.75!)
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Image = CType(resources.GetObject("cmdClose.Image"), System.Drawing.Image)
        Me.cmdClose.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdClose.Location = New System.Drawing.Point(399, 401)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(101, 31)
        Me.cmdClose.TabIndex = 6
        Me.cmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'frmNSK0000HO
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(517, 444)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.lblEventName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lstEventStaff)
        Me.Controls.Add(Me.lstEventList)
        Me.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmNSK0000HO"
        Me.Text = "今月の行事予定"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lstEventStaff As System.Windows.Forms.ListView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblEventName As System.Windows.Forms.Label
    Public WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents lstEventList As System.Windows.Forms.ListView
End Class
