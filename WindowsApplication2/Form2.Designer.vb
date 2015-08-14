<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form2))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.hierarchicalDataGridView1 = New ApplicationAspect.Controls.HierarchicalDataGridView()
        Me.TjDataSet = New TechJournalLoader.tjDataSet()
        Me.LogsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.LogsTableAdapter = New TechJournalLoader.tjDataSetTableAdapters.logsTableAdapter()
        Me.id = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.name_col = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.description = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EstimateRows = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EstimateIO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EstimateCPU = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AvgRowSize = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TotalSubtreeCost = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EstimateExecutions = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.StmtText = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.id_parent = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SaveFileDialog = New System.Windows.Forms.SaveFileDialog()
        CType(Me.hierarchicalDataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TjDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LogsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(502, 410)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Закрыть"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'hierarchicalDataGridView1
        '
        Me.hierarchicalDataGridView1.AllowUserToAddRows = False
        Me.hierarchicalDataGridView1.AllowUserToDeleteRows = False
        Me.hierarchicalDataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.hierarchicalDataGridView1.CellBrush = Nothing
        Me.hierarchicalDataGridView1.CellFont = Nothing
        Me.hierarchicalDataGridView1.CellImage = Nothing
        Me.hierarchicalDataGridView1.CellImageOffset = New System.Drawing.Point(10, 0)
        Me.hierarchicalDataGridView1.CellStringFormat = Nothing
        Me.hierarchicalDataGridView1.CollapsedImage = CType(resources.GetObject("hierarchicalDataGridView1.CollapsedImage"), System.Drawing.Image)
        Me.hierarchicalDataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.hierarchicalDataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.id, Me.name_col, Me.description, Me.EstimateRows, Me.EstimateIO, Me.EstimateCPU, Me.AvgRowSize, Me.TotalSubtreeCost, Me.EstimateExecutions, Me.StmtText, Me.id_parent})
        Me.hierarchicalDataGridView1.DisplayCellImage = False
        Me.hierarchicalDataGridView1.DisplayHierarchy = True
        Me.hierarchicalDataGridView1.ExpandedImage = CType(resources.GetObject("hierarchicalDataGridView1.ExpandedImage"), System.Drawing.Image)
        Me.hierarchicalDataGridView1.HierarchicalColumnName = Nothing
        Me.hierarchicalDataGridView1.Location = New System.Drawing.Point(3, 2)
        Me.hierarchicalDataGridView1.Name = "hierarchicalDataGridView1"
        Me.hierarchicalDataGridView1.ParentCellBrush = Nothing
        Me.hierarchicalDataGridView1.ParentCellFont = Nothing
        Me.hierarchicalDataGridView1.ParentCellStringFormat = Nothing
        Me.hierarchicalDataGridView1.PrimaryKeyColumnName = Nothing
        Me.hierarchicalDataGridView1.ReadOnly = True
        Me.hierarchicalDataGridView1.RootValue = Nothing
        Me.hierarchicalDataGridView1.Size = New System.Drawing.Size(574, 403)
        Me.hierarchicalDataGridView1.TabIndex = 3
        '
        'TjDataSet
        '
        Me.TjDataSet.DataSetName = "tjDataSet"
        Me.TjDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'LogsBindingSource
        '
        Me.LogsBindingSource.DataMember = "logs"
        Me.LogsBindingSource.DataSource = Me.TjDataSet
        '
        'LogsTableAdapter
        '
        Me.LogsTableAdapter.ClearBeforeFill = True
        '
        'id
        '
        Me.id.DataPropertyName = "id"
        Me.id.HeaderText = "id"
        Me.id.Name = "id"
        Me.id.ReadOnly = True
        Me.id.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.id.Visible = False
        '
        'name_col
        '
        Me.name_col.DataPropertyName = "Rows"
        Me.name_col.HeaderText = "Rows"
        Me.name_col.Name = "name_col"
        Me.name_col.ReadOnly = True
        Me.name_col.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.name_col.Width = 50
        '
        'description
        '
        Me.description.DataPropertyName = "Executes"
        Me.description.HeaderText = "Executes"
        Me.description.Name = "description"
        Me.description.ReadOnly = True
        Me.description.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.description.Width = 60
        '
        'EstimateRows
        '
        Me.EstimateRows.DataPropertyName = "EstimateRows"
        Me.EstimateRows.HeaderText = "EstimateRows"
        Me.EstimateRows.Name = "EstimateRows"
        Me.EstimateRows.ReadOnly = True
        Me.EstimateRows.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.EstimateRows.Width = 75
        '
        'EstimateIO
        '
        Me.EstimateIO.DataPropertyName = "EstimateIO"
        Me.EstimateIO.HeaderText = "EstimateIO"
        Me.EstimateIO.Name = "EstimateIO"
        Me.EstimateIO.ReadOnly = True
        Me.EstimateIO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.EstimateIO.Width = 65
        '
        'EstimateCPU
        '
        Me.EstimateCPU.DataPropertyName = "EstimateCPU"
        Me.EstimateCPU.HeaderText = "EstimateCPU"
        Me.EstimateCPU.Name = "EstimateCPU"
        Me.EstimateCPU.ReadOnly = True
        Me.EstimateCPU.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.EstimateCPU.Width = 75
        '
        'AvgRowSize
        '
        Me.AvgRowSize.DataPropertyName = "AvgRowSize"
        Me.AvgRowSize.HeaderText = "AvgRowSize"
        Me.AvgRowSize.Name = "AvgRowSize"
        Me.AvgRowSize.ReadOnly = True
        Me.AvgRowSize.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.AvgRowSize.Width = 75
        '
        'TotalSubtreeCost
        '
        Me.TotalSubtreeCost.DataPropertyName = "TotalSubtreeCost"
        Me.TotalSubtreeCost.HeaderText = "TotalSubtreeCost"
        Me.TotalSubtreeCost.Name = "TotalSubtreeCost"
        Me.TotalSubtreeCost.ReadOnly = True
        Me.TotalSubtreeCost.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'EstimateExecutions
        '
        Me.EstimateExecutions.DataPropertyName = "EstimateExecutions"
        Me.EstimateExecutions.HeaderText = "EstimateExecutions"
        Me.EstimateExecutions.Name = "EstimateExecutions"
        Me.EstimateExecutions.ReadOnly = True
        Me.EstimateExecutions.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'StmtText
        '
        Me.StmtText.DataPropertyName = "StmtText"
        Me.StmtText.HeaderText = "StmtText"
        Me.StmtText.Name = "StmtText"
        Me.StmtText.ReadOnly = True
        Me.StmtText.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.StmtText.Width = 250
        '
        'id_parent
        '
        Me.id_parent.DataPropertyName = "id_parent"
        Me.id_parent.HeaderText = "id_parent"
        Me.id_parent.Name = "id_parent"
        Me.id_parent.ReadOnly = True
        Me.id_parent.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.id_parent.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(416, 413)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(58, 20)
        Me.TextBox1.TabIndex = 4
        Me.TextBox1.Visible = False
        '
        'Button2
        '
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Button2.Location = New System.Drawing.Point(13, 410)
        Me.Button2.Name = "Button2"
        Me.Button2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button2.Size = New System.Drawing.Size(130, 23)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Выгрузить в Excel"
        Me.Button2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button2.UseVisualStyleBackColor = True
        '
        'SaveFileDialog
        '
        Me.SaveFileDialog.Filter = "Excel file|*.xlsx"
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(579, 436)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.hierarchicalDataGridView1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form2"
        Me.Text = "Анализ плана запроса"
        CType(Me.hierarchicalDataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TjDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LogsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TjDataSet As TechJournalLoader.tjDataSet
    Friend WithEvents LogsBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents LogsTableAdapter As TechJournalLoader.tjDataSetTableAdapters.logsTableAdapter
    Private WithEvents hierarchicalDataGridView1 As ApplicationAspect.Controls.HierarchicalDataGridView
    Friend WithEvents id As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents name_col As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EstimateRows As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EstimateIO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EstimateCPU As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AvgRowSize As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TotalSubtreeCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EstimateExecutions As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StmtText As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents id_parent As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog As System.Windows.Forms.SaveFileDialog
End Class
