<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNSK0000HE
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
    Public WithEvents cmdClose As Button
    Public WithEvents cmdApply As Button
    Public WithEvents cmdSave As Button
    Public WithEvents lblBikou As Label
    Friend WithEvents txtBikou As CustomText.NiszText
    Friend WithEvents sprSaveList As FarPoint.Win.Spread.FpSpread
    Friend WithEvents sprSaveList_Sheet1 As FarPoint.Win.Spread.SheetView
    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更できます。
    'コード エディタを使用して、変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNSK0000HE))
        Dim DefaultFocusIndicatorRenderer1 As FarPoint.Win.Spread.DefaultFocusIndicatorRenderer = New FarPoint.Win.Spread.DefaultFocusIndicatorRenderer()
        Dim DefaultScrollBarRenderer1 As FarPoint.Win.Spread.DefaultScrollBarRenderer = New FarPoint.Win.Spread.DefaultScrollBarRenderer()
        Dim DefaultScrollBarRenderer2 As FarPoint.Win.Spread.DefaultScrollBarRenderer = New FarPoint.Win.Spread.DefaultScrollBarRenderer()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.cmdApply = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.lblBikou = New System.Windows.Forms.Label()
        Me.txtBikou = New CustomText.NiszText()
        Me.sprSaveList = New FarPoint.Win.Spread.FpSpread()
        Me.sprSaveList_Sheet1 = New FarPoint.Win.Spread.SheetView()
        CType(Me.sprSaveList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sprSaveList_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdClose.Font = New System.Drawing.Font("MS Gothic", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Image = CType(resources.GetObject("cmdClose.Image"), System.Drawing.Image)
        Me.cmdClose.Location = New System.Drawing.Point(548, 217)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(100, 36)
        Me.cmdClose.TabIndex = 6
        Me.cmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'cmdApply
        '
        Me.cmdApply.BackColor = System.Drawing.SystemColors.Control
        Me.cmdApply.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdApply.Font = New System.Drawing.Font("MS Gothic", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmdApply.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdApply.Image = CType(resources.GetObject("cmdApply.Image"), System.Drawing.Image)
        Me.cmdApply.Location = New System.Drawing.Point(425, 217)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdApply.Size = New System.Drawing.Size(100, 36)
        Me.cmdApply.TabIndex = 5
        Me.cmdApply.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdApply.UseVisualStyleBackColor = False
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSave.Font = New System.Drawing.Font("MS Gothic", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.Image = CType(resources.GetObject("cmdSave.Image"), System.Drawing.Image)
        Me.cmdSave.Location = New System.Drawing.Point(302, 217)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSave.Size = New System.Drawing.Size(100, 36)
        Me.cmdSave.TabIndex = 4
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'lblBikou
        '
        Me.lblBikou.BackColor = System.Drawing.SystemColors.Control
        Me.lblBikou.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBikou.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBikou.Location = New System.Drawing.Point(37, 177)
        Me.lblBikou.Name = "lblBikou"
        Me.lblBikou.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBikou.Size = New System.Drawing.Size(49, 20)
        Me.lblBikou.TabIndex = 2
        Me.lblBikou.Text = "備考"
        Me.lblBikou.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtBikou
        '
        Me.txtBikou.EnabledBackColor = System.Drawing.SystemColors.Window
        Me.txtBikou.Format = "HZ"
        Me.txtBikou.HighlightText = True
        Me.txtBikou.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.txtBikou.Location = New System.Drawing.Point(97, 173)
        Me.txtBikou.MaxLength = 20
        Me.txtBikou.MaxLengthUnit = "B"
        Me.txtBikou.Name = "txtBikou"
        Me.txtBikou.NumType = False
        Me.txtBikou.Size = New System.Drawing.Size(325, 22)
        Me.txtBikou.TabIndex = 3
        '
        'sprSaveList
        '
        Me.sprSaveList.AccessibleDescription = "sprDateList, Sheet1, Row 0, Column 0, "
        Me.sprSaveList.AllowUserZoom = False
        Me.sprSaveList.ColumnSplitBoxPolicy = FarPoint.Win.Spread.SplitBoxPolicy.Never
        Me.sprSaveList.FocusRenderer = DefaultFocusIndicatorRenderer1
        Me.sprSaveList.Font = New System.Drawing.Font("MS Gothic", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.sprSaveList.HorizontalScrollBar.Buttons = New FarPoint.Win.Spread.FpScrollBarButtonCollection("BackwardLineButton,ThumbTrack,ForwardLineButton")
        Me.sprSaveList.HorizontalScrollBar.Name = ""
        Me.sprSaveList.HorizontalScrollBar.Renderer = DefaultScrollBarRenderer1
        Me.sprSaveList.HorizontalScrollBar.TabIndex = 3
        Me.sprSaveList.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.sprSaveList.Location = New System.Drawing.Point(21, 21)
        Me.sprSaveList.Name = "sprSaveList"
        Me.sprSaveList.RowSplitBoxPolicy = FarPoint.Win.Spread.SplitBoxPolicy.Never
        Me.sprSaveList.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.sprSaveList_Sheet1})
        Me.sprSaveList.Size = New System.Drawing.Size(645, 124)
        Me.sprSaveList.Skin = FarPoint.Win.Spread.DefaultSpreadSkins.Classic
        Me.sprSaveList.TabIndex = 1
        Me.sprSaveList.TabStop = False
        Me.sprSaveList.VerticalScrollBar.Buttons = New FarPoint.Win.Spread.FpScrollBarButtonCollection("BackwardLineButton,ThumbTrack,ForwardLineButton")
        Me.sprSaveList.VerticalScrollBar.Name = ""
        Me.sprSaveList.VerticalScrollBar.Renderer = DefaultScrollBarRenderer2
        Me.sprSaveList.VerticalScrollBar.TabIndex = 4
        Me.sprSaveList.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.sprSaveList.VisualStyles = FarPoint.Win.VisualStyles.Off
        '
        'sprSaveList_Sheet1
        '
        Me.sprSaveList_Sheet1.Reset()
        Me.sprSaveList_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.sprSaveList_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.sprSaveList_Sheet1.ColumnCount = 4
        Me.sprSaveList_Sheet1.RowHeader.ColumnCount = 0
        Me.sprSaveList_Sheet1.ColumnFooter.Columns.Default.Resizable = False
        Me.sprSaveList_Sheet1.ColumnFooter.DefaultStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.sprSaveList_Sheet1.ColumnFooter.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.sprSaveList_Sheet1.ColumnFooter.DefaultStyle.Parent = "HeaderDefault"
        Me.sprSaveList_Sheet1.ColumnFooter.DefaultStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.sprSaveList_Sheet1.ColumnFooterSheetCornerStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.sprSaveList_Sheet1.ColumnFooterSheetCornerStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.sprSaveList_Sheet1.ColumnFooterSheetCornerStyle.Parent = "RowHeaderDefault"
        Me.sprSaveList_Sheet1.ColumnFooterSheetCornerStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.sprSaveList_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "No"
        Me.sprSaveList_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "保存者"
        Me.sprSaveList_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "保存日時"
        Me.sprSaveList_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "備考"
        Me.sprSaveList_Sheet1.ColumnHeader.Columns.Default.Resizable = False
        Me.sprSaveList_Sheet1.ColumnHeader.DefaultStyle.Font = New System.Drawing.Font("MS Gothic", 11.25!)
        Me.sprSaveList_Sheet1.ColumnHeader.DefaultStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.sprSaveList_Sheet1.ColumnHeader.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.sprSaveList_Sheet1.ColumnHeader.DefaultStyle.Parent = "HeaderDefault"
        Me.sprSaveList_Sheet1.ColumnHeader.DefaultStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.sprSaveList_Sheet1.ColumnHeader.Rows.Get(0).Height = 21.0!
        Me.sprSaveList_Sheet1.Columns.Default.Resizable = False
        Me.sprSaveList_Sheet1.Columns.Get(0).Label = "No"
        Me.sprSaveList_Sheet1.Columns.Get(0).Locked = False
        Me.sprSaveList_Sheet1.Columns.Get(0).Width = 25.0!
        Me.sprSaveList_Sheet1.Columns.Get(1).Label = "保存者"
        Me.sprSaveList_Sheet1.Columns.Get(1).Locked = False
        Me.sprSaveList_Sheet1.Columns.Get(1).Width = 100.0!
        Me.sprSaveList_Sheet1.Columns.Get(2).Label = "保存日時"
        Me.sprSaveList_Sheet1.Columns.Get(2).Locked = False
        Me.sprSaveList_Sheet1.Columns.Get(2).Width = 175.0!
        Me.sprSaveList_Sheet1.Columns.Get(3).Label = "備考"
        Me.sprSaveList_Sheet1.Columns.Get(3).Locked = False
        Me.sprSaveList_Sheet1.Columns.Get(3).Width = 325.0!
        Me.sprSaveList_Sheet1.DefaultStyle.Font = New System.Drawing.Font("MS Gothic", 11.25!)
        Me.sprSaveList_Sheet1.DefaultStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.sprSaveList_Sheet1.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.sprSaveList_Sheet1.DefaultStyle.Parent = "DataAreaDefault"
        Me.sprSaveList_Sheet1.DefaultStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.sprSaveList_Sheet1.FilterBar.DefaultStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.sprSaveList_Sheet1.FilterBar.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.sprSaveList_Sheet1.FilterBar.DefaultStyle.Parent = "FilterBarDefault"
        Me.sprSaveList_Sheet1.FilterBar.DefaultStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.sprSaveList_Sheet1.FilterBarHeaderStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.sprSaveList_Sheet1.FilterBarHeaderStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.sprSaveList_Sheet1.FilterBarHeaderStyle.Parent = "RowHeaderDefault"
        Me.sprSaveList_Sheet1.FilterBarHeaderStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.sprSaveList_Sheet1.OperationMode = FarPoint.Win.Spread.OperationMode.[ReadOnly]
        Me.sprSaveList_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.sprSaveList_Sheet1.RowHeader.DefaultStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.sprSaveList_Sheet1.RowHeader.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.sprSaveList_Sheet1.RowHeader.DefaultStyle.Parent = "RowHeaderDefault"
        Me.sprSaveList_Sheet1.RowHeader.DefaultStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.sprSaveList_Sheet1.Rows.Default.Height = 21.0!
        Me.sprSaveList_Sheet1.SelectionBackColor = System.Drawing.Color.Cyan
        Me.sprSaveList_Sheet1.SelectionStyle = FarPoint.Win.Spread.SelectionStyles.None
        Me.sprSaveList_Sheet1.SelectionUnit = FarPoint.Win.Spread.Model.SelectionUnit.Row
        Me.sprSaveList_Sheet1.SheetCornerStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.sprSaveList_Sheet1.SheetCornerStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.sprSaveList_Sheet1.SheetCornerStyle.Parent = "RowHeaderDefault"
        Me.sprSaveList_Sheet1.SheetCornerStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.sprSaveList_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'frmNSK0000HE
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(670, 274)
        Me.Controls.Add(Me.sprSaveList)
        Me.Controls.Add(Me.lblBikou)
        Me.Controls.Add(Me.txtBikou)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdApply)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("MS Gothic", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Location = New System.Drawing.Point(246, 223)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNSK0000HE"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "保存呼出"
        CType(Me.sprSaveList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sprSaveList_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class
