<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCodeGenerator
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
    Me.btnGenerateCode = New System.Windows.Forms.Button
    Me.btnValidateEntityItemsWithDatabase = New System.Windows.Forms.Button
    Me.cboPAProjects = New System.Windows.Forms.ComboBox
    Me.dgvPrjEntities = New System.Windows.Forms.DataGridView
    Me.dgvPrjEntityItems = New System.Windows.Forms.DataGridView
    Me.dgvDBCheck = New System.Windows.Forms.DataGridView
    Me.txtMissingFields = New System.Windows.Forms.TextBox
    Me.btnShowMainForm = New System.Windows.Forms.Button
    Me.btnCreateEntityItemsFromDatabase = New System.Windows.Forms.Button
    CType(Me.dgvPrjEntities, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.dgvPrjEntityItems, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.dgvDBCheck, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'btnGenerateCode
    '
    Me.btnGenerateCode.Location = New System.Drawing.Point(42, 18)
    Me.btnGenerateCode.Name = "btnGenerateCode"
    Me.btnGenerateCode.Size = New System.Drawing.Size(125, 27)
    Me.btnGenerateCode.TabIndex = 0
    Me.btnGenerateCode.Text = "Generate Code"
    Me.btnGenerateCode.UseVisualStyleBackColor = True
    '
    'btnValidateEntityItemsWithDatabase
    '
    Me.btnValidateEntityItemsWithDatabase.Location = New System.Drawing.Point(556, 18)
    Me.btnValidateEntityItemsWithDatabase.Name = "btnValidateEntityItemsWithDatabase"
    Me.btnValidateEntityItemsWithDatabase.Size = New System.Drawing.Size(125, 54)
    Me.btnValidateEntityItemsWithDatabase.TabIndex = 1
    Me.btnValidateEntityItemsWithDatabase.Text = "Validate Entity Items with  Database"
    Me.btnValidateEntityItemsWithDatabase.UseVisualStyleBackColor = True
    '
    'cboPAProjects
    '
    Me.cboPAProjects.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboPAProjects.DisplayMember = "ProjectName"
    Me.cboPAProjects.FormattingEnabled = True
    Me.cboPAProjects.Location = New System.Drawing.Point(42, 51)
    Me.cboPAProjects.Name = "cboPAProjects"
    Me.cboPAProjects.Size = New System.Drawing.Size(323, 21)
    Me.cboPAProjects.TabIndex = 19
    '
    'dgvPrjEntities
    '
    Me.dgvPrjEntities.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvPrjEntities.Location = New System.Drawing.Point(42, 88)
    Me.dgvPrjEntities.Name = "dgvPrjEntities"
    Me.dgvPrjEntities.Size = New System.Drawing.Size(240, 265)
    Me.dgvPrjEntities.TabIndex = 20
    '
    'dgvPrjEntityItems
    '
    Me.dgvPrjEntityItems.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvPrjEntityItems.Location = New System.Drawing.Point(321, 88)
    Me.dgvPrjEntityItems.Name = "dgvPrjEntityItems"
    Me.dgvPrjEntityItems.Size = New System.Drawing.Size(240, 265)
    Me.dgvPrjEntityItems.TabIndex = 21
    '
    'dgvDBCheck
    '
    Me.dgvDBCheck.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvDBCheck.Location = New System.Drawing.Point(584, 88)
    Me.dgvDBCheck.Name = "dgvDBCheck"
    Me.dgvDBCheck.Size = New System.Drawing.Size(221, 70)
    Me.dgvDBCheck.TabIndex = 22
    '
    'txtMissingFields
    '
    Me.txtMissingFields.Location = New System.Drawing.Point(584, 187)
    Me.txtMissingFields.Multiline = True
    Me.txtMissingFields.Name = "txtMissingFields"
    Me.txtMissingFields.ScrollBars = System.Windows.Forms.ScrollBars.Both
    Me.txtMissingFields.Size = New System.Drawing.Size(221, 166)
    Me.txtMissingFields.TabIndex = 23
    '
    'btnShowMainForm
    '
    Me.btnShowMainForm.AllowDrop = True
    Me.btnShowMainForm.Location = New System.Drawing.Point(697, 18)
    Me.btnShowMainForm.Name = "btnShowMainForm"
    Me.btnShowMainForm.Size = New System.Drawing.Size(125, 27)
    Me.btnShowMainForm.TabIndex = 24
    Me.btnShowMainForm.Text = "Show Main Form"
    Me.btnShowMainForm.UseVisualStyleBackColor = True
    '
    'btnCreateEntityItemsFromDatabase
    '
    Me.btnCreateEntityItemsFromDatabase.Location = New System.Drawing.Point(393, 12)
    Me.btnCreateEntityItemsFromDatabase.Name = "btnCreateEntityItemsFromDatabase"
    Me.btnCreateEntityItemsFromDatabase.Size = New System.Drawing.Size(125, 60)
    Me.btnCreateEntityItemsFromDatabase.TabIndex = 25
    Me.btnCreateEntityItemsFromDatabase.Text = "Create Entity Items from Database"
    Me.btnCreateEntityItemsFromDatabase.UseVisualStyleBackColor = True
    '
    'frmCodeGenerator
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(848, 384)
    Me.Controls.Add(Me.btnCreateEntityItemsFromDatabase)
    Me.Controls.Add(Me.btnShowMainForm)
    Me.Controls.Add(Me.txtMissingFields)
    Me.Controls.Add(Me.dgvDBCheck)
    Me.Controls.Add(Me.dgvPrjEntityItems)
    Me.Controls.Add(Me.dgvPrjEntities)
    Me.Controls.Add(Me.cboPAProjects)
    Me.Controls.Add(Me.btnValidateEntityItemsWithDatabase)
    Me.Controls.Add(Me.btnGenerateCode)
    Me.Name = "frmCodeGenerator"
    Me.Text = "Form3"
    CType(Me.dgvPrjEntities, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.dgvPrjEntityItems, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.dgvDBCheck, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents btnGenerateCode As System.Windows.Forms.Button
  Friend WithEvents btnValidateEntityItemsWithDatabase As System.Windows.Forms.Button
  Friend WithEvents cboPAProjects As System.Windows.Forms.ComboBox
  Friend WithEvents dgvPrjEntities As System.Windows.Forms.DataGridView
  Friend WithEvents dgvPrjEntityItems As System.Windows.Forms.DataGridView
  Friend WithEvents dgvDBCheck As System.Windows.Forms.DataGridView
  Friend WithEvents txtMissingFields As System.Windows.Forms.TextBox
  Friend WithEvents btnShowMainForm As System.Windows.Forms.Button
  Friend WithEvents btnCreateEntityItemsFromDatabase As System.Windows.Forms.Button
End Class
