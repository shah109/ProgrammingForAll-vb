<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.EntitiesItems = New System.Windows.Forms.TabControl
        Me.ProjectEntities = New System.Windows.Forms.TabPage
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnRemove = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader11 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader12 = New System.Windows.Forms.ColumnHeader
        Me.btnMoveUp = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.lvProjectEntities = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.tabEntityItems = New System.Windows.Forms.TabPage
        Me.tabGenerateCode = New System.Windows.Forms.TabPage
        Me.cboProjectName = New System.Windows.Forms.ComboBox
        Me.lblProjects = New System.Windows.Forms.Label
        Me.EntitiesItems.SuspendLayout()
        Me.ProjectEntities.SuspendLayout()
        Me.SuspendLayout()
        '
        'EntitiesItems
        '
        Me.EntitiesItems.Appearance = System.Windows.Forms.TabAppearance.Buttons
        Me.EntitiesItems.Controls.Add(Me.ProjectEntities)
        Me.EntitiesItems.Controls.Add(Me.tabEntityItems)
        Me.EntitiesItems.Controls.Add(Me.tabGenerateCode)
        Me.EntitiesItems.Location = New System.Drawing.Point(12, 61)
        Me.EntitiesItems.Name = "EntitiesItems"
        Me.EntitiesItems.SelectedIndex = 0
        Me.EntitiesItems.Size = New System.Drawing.Size(821, 384)
        Me.EntitiesItems.TabIndex = 0
        '
        'ProjectEntities
        '
        Me.ProjectEntities.BackColor = System.Drawing.Color.Transparent
        Me.ProjectEntities.Controls.Add(Me.Label6)
        Me.ProjectEntities.Controls.Add(Me.Label5)
        Me.ProjectEntities.Controls.Add(Me.btnRemove)
        Me.ProjectEntities.Controls.Add(Me.btnAdd)
        Me.ProjectEntities.Controls.Add(Me.ListView1)
        Me.ProjectEntities.Controls.Add(Me.btnMoveUp)
        Me.ProjectEntities.Controls.Add(Me.Label4)
        Me.ProjectEntities.Controls.Add(Me.Label3)
        Me.ProjectEntities.Controls.Add(Me.Label2)
        Me.ProjectEntities.Controls.Add(Me.Label1)
        Me.ProjectEntities.Controls.Add(Me.TextBox6)
        Me.ProjectEntities.Controls.Add(Me.TextBox5)
        Me.ProjectEntities.Controls.Add(Me.TextBox4)
        Me.ProjectEntities.Controls.Add(Me.TextBox3)
        Me.ProjectEntities.Controls.Add(Me.TextBox2)
        Me.ProjectEntities.Controls.Add(Me.TextBox1)
        Me.ProjectEntities.Controls.Add(Me.lvProjectEntities)
        Me.ProjectEntities.Location = New System.Drawing.Point(4, 25)
        Me.ProjectEntities.Name = "ProjectEntities"
        Me.ProjectEntities.Padding = New System.Windows.Forms.Padding(3)
        Me.ProjectEntities.Size = New System.Drawing.Size(813, 355)
        Me.ProjectEntities.TabIndex = 0
        Me.ProjectEntities.Text = "Project Entities"
        Me.ProjectEntities.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(443, 9)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(85, 13)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Included Entities"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 9)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(87, 13)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Available Entities"
        '
        'btnRemove
        '
        Me.btnRemove.Location = New System.Drawing.Point(374, 118)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(51, 23)
        Me.btnRemove.TabIndex = 14
        Me.btnRemove.Text = "<"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(375, 89)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(51, 23)
        Me.btnAdd.TabIndex = 13
        Me.btnAdd.Text = ">"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader7, Me.ColumnHeader8, Me.ColumnHeader9, Me.ColumnHeader10, Me.ColumnHeader11, Me.ColumnHeader12})
        Me.ListView1.FullRowSelect = True
        Me.ListView1.GridLines = True
        Me.ListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(437, 25)
        Me.ListView1.MultiSelect = False
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(336, 152)
        Me.ListView1.TabIndex = 12
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.List
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Entity ID"
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Entity Collection"
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Entity Short Name"
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "Input File"
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "Last Code Generated"
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = "Last Updated"
        '
        'btnMoveUp
        '
        Me.btnMoveUp.Location = New System.Drawing.Point(379, 46)
        Me.btnMoveUp.Name = "btnMoveUp"
        Me.btnMoveUp.Size = New System.Drawing.Size(51, 37)
        Me.btnMoveUp.TabIndex = 11
        Me.btnMoveUp.Text = "Move Up"
        Me.btnMoveUp.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(426, 252)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(113, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Entity Collection Name"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(426, 213)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(113, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Entity Collection Name"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(255, 252)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(113, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Entity Collection Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(255, 213)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Entity Name"
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(624, 268)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(149, 20)
        Me.TextBox6.TabIndex = 6
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(624, 229)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(149, 20)
        Me.TextBox5.TabIndex = 5
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(429, 268)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(149, 20)
        Me.TextBox4.TabIndex = 4
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(429, 229)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(149, 20)
        Me.TextBox3.TabIndex = 3
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(258, 268)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(149, 20)
        Me.TextBox2.TabIndex = 2
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(258, 229)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(149, 20)
        Me.TextBox1.TabIndex = 1
        '
        'lvProjectEntities
        '
        Me.lvProjectEntities.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.lvProjectEntities.FullRowSelect = True
        Me.lvProjectEntities.GridLines = True
        Me.lvProjectEntities.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.lvProjectEntities.HideSelection = False
        Me.lvProjectEntities.Location = New System.Drawing.Point(15, 25)
        Me.lvProjectEntities.MultiSelect = False
        Me.lvProjectEntities.Name = "lvProjectEntities"
        Me.lvProjectEntities.Size = New System.Drawing.Size(349, 152)
        Me.lvProjectEntities.TabIndex = 0
        Me.lvProjectEntities.UseCompatibleStateImageBehavior = False
        Me.lvProjectEntities.View = System.Windows.Forms.View.List
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Entity ID"
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Entity Collection"
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Entity Short Name"
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Input File"
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Last Code Generated"
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Last Updated"
        '
        'tabEntityItems
        '
        Me.tabEntityItems.Location = New System.Drawing.Point(4, 25)
        Me.tabEntityItems.Name = "tabEntityItems"
        Me.tabEntityItems.Padding = New System.Windows.Forms.Padding(3)
        Me.tabEntityItems.Size = New System.Drawing.Size(813, 355)
        Me.tabEntityItems.TabIndex = 1
        Me.tabEntityItems.Text = "Entitiy Items"
        Me.tabEntityItems.UseVisualStyleBackColor = True
        '
        'tabGenerateCode
        '
        Me.tabGenerateCode.Location = New System.Drawing.Point(4, 25)
        Me.tabGenerateCode.Name = "tabGenerateCode"
        Me.tabGenerateCode.Padding = New System.Windows.Forms.Padding(3)
        Me.tabGenerateCode.Size = New System.Drawing.Size(813, 355)
        Me.tabGenerateCode.TabIndex = 2
        Me.tabGenerateCode.Text = "Generate Code"
        Me.tabGenerateCode.UseVisualStyleBackColor = True
        '
        'cboProjectName
        '
        Me.cboProjectName.FormattingEnabled = True
        Me.cboProjectName.Location = New System.Drawing.Point(42, 21)
        Me.cboProjectName.Name = "cboProjectName"
        Me.cboProjectName.Size = New System.Drawing.Size(194, 21)
        Me.cboProjectName.TabIndex = 1
        '
        'lblProjects
        '
        Me.lblProjects.AutoSize = True
        Me.lblProjects.Location = New System.Drawing.Point(39, 5)
        Me.lblProjects.Name = "lblProjects"
        Me.lblProjects.Size = New System.Drawing.Size(40, 13)
        Me.lblProjects.TabIndex = 11
        Me.lblProjects.Text = "Project"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(845, 452)
        Me.Controls.Add(Me.lblProjects)
        Me.Controls.Add(Me.cboProjectName)
        Me.Controls.Add(Me.EntitiesItems)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.EntitiesItems.ResumeLayout(False)
        Me.ProjectEntities.ResumeLayout(False)
        Me.ProjectEntities.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents EntitiesItems As System.Windows.Forms.TabControl
    Friend WithEvents ProjectEntities As System.Windows.Forms.TabPage
    Friend WithEvents lvProjectEntities As System.Windows.Forms.ListView
    Public WithEvents tabEntityItems As System.Windows.Forms.TabPage
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cboProjectName As System.Windows.Forms.ComboBox
    Friend WithEvents lblProjects As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader12 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnMoveUp As System.Windows.Forms.Button
    Friend WithEvents tabGenerateCode As System.Windows.Forms.TabPage
End Class
