<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSettings
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
    Me.txtAccessDBPath = New System.Windows.Forms.TextBox
    Me.Label1 = New System.Windows.Forms.Label
    Me.TabSettings = New System.Windows.Forms.TabControl
    Me.TabPage1 = New System.Windows.Forms.TabPage
    Me.Label14 = New System.Windows.Forms.Label
    Me.Label7 = New System.Windows.Forms.Label
    Me.TextBox2 = New System.Windows.Forms.TextBox
    Me.TextBox1 = New System.Windows.Forms.TextBox
    Me.chkMultiUserAccess = New System.Windows.Forms.CheckBox
    Me.Label4 = New System.Windows.Forms.Label
    Me.Label3 = New System.Windows.Forms.Label
    Me.txtAccessPassword = New System.Windows.Forms.TextBox
    Me.txtAccessLogin = New System.Windows.Forms.TextBox
    Me.btnAccessDBPath = New System.Windows.Forms.Button
    Me.Label2 = New System.Windows.Forms.Label
    Me.cboDBType = New System.Windows.Forms.ComboBox
    Me.TabPage2 = New System.Windows.Forms.TabPage
    Me.CheckBox1 = New System.Windows.Forms.CheckBox
    Me.cboMonthsOfAudit = New System.Windows.Forms.ComboBox
    Me.Label6 = New System.Windows.Forms.Label
    Me.chkAskForUpdateConfirmation = New System.Windows.Forms.CheckBox
    Me.chkDontShowSplash = New System.Windows.Forms.CheckBox
    Me.Label13 = New System.Windows.Forms.Label
    Me.txtAuditEntriesToShow = New System.Windows.Forms.TextBox
    Me.TabPage3 = New System.Windows.Forms.TabPage
    Me.Label5 = New System.Windows.Forms.Label
    Me.txtAccessRights = New System.Windows.Forms.TextBox
    Me.Label12 = New System.Windows.Forms.Label
    Me.Label11 = New System.Windows.Forms.Label
    Me.Label10 = New System.Windows.Forms.Label
    Me.txtWinDomain = New System.Windows.Forms.TextBox
    Me.txtWinComputerName = New System.Windows.Forms.TextBox
    Me.txtWinUserName = New System.Windows.Forms.TextBox
    Me.txtLastUpdateID = New System.Windows.Forms.TextBox
    Me.Label9 = New System.Windows.Forms.Label
    Me.txtLoginNumber = New System.Windows.Forms.TextBox
    Me.Label8 = New System.Windows.Forms.Label
    Me.btnCancel = New System.Windows.Forms.Button
    Me.btnApply = New System.Windows.Forms.Button
    Me.txtCurrentDirectory = New System.Windows.Forms.TextBox
    Me.Label15 = New System.Windows.Forms.Label
    Me.TabSettings.SuspendLayout()
    Me.TabPage1.SuspendLayout()
    Me.TabPage2.SuspendLayout()
    Me.TabPage3.SuspendLayout()
    Me.SuspendLayout()
    '
    'txtAccessDBPath
    '
    Me.txtAccessDBPath.Location = New System.Drawing.Point(28, 123)
    Me.txtAccessDBPath.Name = "txtAccessDBPath"
    Me.txtAccessDBPath.Size = New System.Drawing.Size(422, 20)
    Me.txtAccessDBPath.TabIndex = 0
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(25, 107)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(91, 13)
    Me.Label1.TabIndex = 1
    Me.Label1.Text = "Access Database"
    '
    'TabSettings
    '
    Me.TabSettings.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
    Me.TabSettings.Controls.Add(Me.TabPage1)
    Me.TabSettings.Controls.Add(Me.TabPage2)
    Me.TabSettings.Controls.Add(Me.TabPage3)
    Me.TabSettings.Dock = System.Windows.Forms.DockStyle.Top
    Me.TabSettings.Location = New System.Drawing.Point(0, 0)
    Me.TabSettings.Name = "TabSettings"
    Me.TabSettings.SelectedIndex = 0
    Me.TabSettings.Size = New System.Drawing.Size(525, 260)
    Me.TabSettings.TabIndex = 2
    '
    'TabPage1
    '
    Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
    Me.TabPage1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.TabPage1.Controls.Add(Me.Label14)
    Me.TabPage1.Controls.Add(Me.Label7)
    Me.TabPage1.Controls.Add(Me.TextBox2)
    Me.TabPage1.Controls.Add(Me.TextBox1)
    Me.TabPage1.Controls.Add(Me.chkMultiUserAccess)
    Me.TabPage1.Controls.Add(Me.Label4)
    Me.TabPage1.Controls.Add(Me.Label3)
    Me.TabPage1.Controls.Add(Me.txtAccessPassword)
    Me.TabPage1.Controls.Add(Me.txtAccessLogin)
    Me.TabPage1.Controls.Add(Me.btnAccessDBPath)
    Me.TabPage1.Controls.Add(Me.Label2)
    Me.TabPage1.Controls.Add(Me.cboDBType)
    Me.TabPage1.Controls.Add(Me.txtAccessDBPath)
    Me.TabPage1.Controls.Add(Me.Label1)
    Me.TabPage1.Location = New System.Drawing.Point(4, 25)
    Me.TabPage1.Name = "TabPage1"
    Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage1.Size = New System.Drawing.Size(517, 231)
    Me.TabPage1.TabIndex = 0
    Me.TabPage1.Text = "Database"
    Me.TabPage1.UseVisualStyleBackColor = True
    '
    'Label14
    '
    Me.Label14.AutoSize = True
    Me.Label14.Location = New System.Drawing.Point(281, 47)
    Me.Label14.Name = "Label14"
    Me.Label14.Size = New System.Drawing.Size(53, 13)
    Me.Label14.TabIndex = 14
    Me.Label14.Text = "DB Name"
    '
    'Label7
    '
    Me.Label7.AutoSize = True
    Me.Label7.Location = New System.Drawing.Point(281, 23)
    Me.Label7.Name = "Label7"
    Me.Label7.Size = New System.Drawing.Size(69, 13)
    Me.Label7.TabIndex = 13
    Me.Label7.Text = "Server Name"
    '
    'TextBox2
    '
    Me.TextBox2.Enabled = False
    Me.TextBox2.Location = New System.Drawing.Point(367, 42)
    Me.TextBox2.Name = "TextBox2"
    Me.TextBox2.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
    Me.TextBox2.Size = New System.Drawing.Size(135, 20)
    Me.TextBox2.TabIndex = 12
    '
    'TextBox1
    '
    Me.TextBox1.Enabled = False
    Me.TextBox1.Location = New System.Drawing.Point(367, 16)
    Me.TextBox1.Name = "TextBox1"
    Me.TextBox1.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
    Me.TextBox1.Size = New System.Drawing.Size(135, 20)
    Me.TextBox1.TabIndex = 11
    '
    'chkMultiUserAccess
    '
    Me.chkMultiUserAccess.AutoSize = True
    Me.chkMultiUserAccess.Location = New System.Drawing.Point(29, 67)
    Me.chkMultiUserAccess.Name = "chkMultiUserAccess"
    Me.chkMultiUserAccess.Size = New System.Drawing.Size(109, 17)
    Me.chkMultiUserAccess.TabIndex = 10
    Me.chkMultiUserAccess.Text = "Multi-user Access"
    Me.chkMultiUserAccess.UseVisualStyleBackColor = True
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(256, 152)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(53, 13)
    Me.Label4.TabIndex = 9
    Me.Label4.Text = "Password"
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(26, 151)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(47, 13)
    Me.Label3.TabIndex = 8
    Me.Label3.Text = "Login ID"
    '
    'txtAccessPassword
    '
    Me.txtAccessPassword.Location = New System.Drawing.Point(315, 149)
    Me.txtAccessPassword.Name = "txtAccessPassword"
    Me.txtAccessPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
    Me.txtAccessPassword.Size = New System.Drawing.Size(135, 20)
    Me.txtAccessPassword.TabIndex = 7
    '
    'txtAccessLogin
    '
    Me.txtAccessLogin.Location = New System.Drawing.Point(79, 149)
    Me.txtAccessLogin.Name = "txtAccessLogin"
    Me.txtAccessLogin.Size = New System.Drawing.Size(135, 20)
    Me.txtAccessLogin.TabIndex = 6
    '
    'btnAccessDBPath
    '
    Me.btnAccessDBPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.btnAccessDBPath.Location = New System.Drawing.Point(465, 123)
    Me.btnAccessDBPath.Name = "btnAccessDBPath"
    Me.btnAccessDBPath.Size = New System.Drawing.Size(37, 20)
    Me.btnAccessDBPath.TabIndex = 5
    Me.btnAccessDBPath.Text = "---"
    Me.btnAccessDBPath.TextAlign = System.Drawing.ContentAlignment.TopCenter
    Me.btnAccessDBPath.UseVisualStyleBackColor = True
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(25, 23)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(80, 13)
    Me.Label2.TabIndex = 3
    Me.Label2.Text = "Database Type"
    '
    'cboDBType
    '
    Me.cboDBType.FormattingEnabled = True
    Me.cboDBType.Items.AddRange(New Object() {"MS Access", "SQL Server", "Oracle", "MySQL"})
    Me.cboDBType.Location = New System.Drawing.Point(28, 39)
    Me.cboDBType.Name = "cboDBType"
    Me.cboDBType.Size = New System.Drawing.Size(136, 21)
    Me.cboDBType.TabIndex = 2
    '
    'TabPage2
    '
    Me.TabPage2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.TabPage2.Controls.Add(Me.CheckBox1)
    Me.TabPage2.Controls.Add(Me.cboMonthsOfAudit)
    Me.TabPage2.Controls.Add(Me.Label6)
    Me.TabPage2.Controls.Add(Me.chkAskForUpdateConfirmation)
    Me.TabPage2.Controls.Add(Me.chkDontShowSplash)
    Me.TabPage2.Controls.Add(Me.Label13)
    Me.TabPage2.Controls.Add(Me.txtAuditEntriesToShow)
    Me.TabPage2.Location = New System.Drawing.Point(4, 25)
    Me.TabPage2.Name = "TabPage2"
    Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage2.Size = New System.Drawing.Size(517, 231)
    Me.TabPage2.TabIndex = 2
    Me.TabPage2.Text = "Options"
    Me.TabPage2.UseVisualStyleBackColor = True
    '
    'CheckBox1
    '
    Me.CheckBox1.AutoSize = True
    Me.CheckBox1.Location = New System.Drawing.Point(26, 199)
    Me.CheckBox1.Name = "CheckBox1"
    Me.CheckBox1.Size = New System.Drawing.Size(149, 17)
    Me.CheckBox1.TabIndex = 20
    Me.CheckBox1.Text = "Post Update Confirmation "
    Me.CheckBox1.UseVisualStyleBackColor = True
    '
    'cboMonthsOfAudit
    '
    Me.cboMonthsOfAudit.AutoCompleteCustomSource.AddRange(New String() {"1", "2", "3"})
    Me.cboMonthsOfAudit.FormattingEnabled = True
    Me.cboMonthsOfAudit.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "18", "18", "19", "20"})
    Me.cboMonthsOfAudit.Location = New System.Drawing.Point(412, 54)
    Me.cboMonthsOfAudit.Name = "cboMonthsOfAudit"
    Me.cboMonthsOfAudit.Size = New System.Drawing.Size(70, 21)
    Me.cboMonthsOfAudit.TabIndex = 19
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.Location = New System.Drawing.Point(236, 54)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(158, 13)
    Me.Label6.TabIndex = 18
    Me.Label6.Text = "Months of Audit Entries to Show"
    '
    'chkAskForUpdateConfirmation
    '
    Me.chkAskForUpdateConfirmation.AutoSize = True
    Me.chkAskForUpdateConfirmation.Location = New System.Drawing.Point(26, 176)
    Me.chkAskForUpdateConfirmation.Name = "chkAskForUpdateConfirmation"
    Me.chkAskForUpdateConfirmation.Size = New System.Drawing.Size(194, 17)
    Me.chkAskForUpdateConfirmation.TabIndex = 16
    Me.chkAskForUpdateConfirmation.Text = "Ask For Confirmation before Update"
    Me.chkAskForUpdateConfirmation.UseVisualStyleBackColor = True
    '
    'chkDontShowSplash
    '
    Me.chkDontShowSplash.AutoSize = True
    Me.chkDontShowSplash.Location = New System.Drawing.Point(26, 153)
    Me.chkDontShowSplash.Name = "chkDontShowSplash"
    Me.chkDontShowSplash.Size = New System.Drawing.Size(163, 17)
    Me.chkDontShowSplash.TabIndex = 15
    Me.chkDontShowSplash.Text = "Dont Show Splash at Startup"
    Me.chkDontShowSplash.UseVisualStyleBackColor = True
    '
    'Label13
    '
    Me.Label13.AutoSize = True
    Me.Label13.Location = New System.Drawing.Point(236, 28)
    Me.Label13.Name = "Label13"
    Me.Label13.Size = New System.Drawing.Size(140, 13)
    Me.Label13.TabIndex = 9
    Me.Label13.Text = "No. of Audit Entries to Show"
    '
    'txtAuditEntriesToShow
    '
    Me.txtAuditEntriesToShow.Location = New System.Drawing.Point(382, 25)
    Me.txtAuditEntriesToShow.Name = "txtAuditEntriesToShow"
    Me.txtAuditEntriesToShow.Size = New System.Drawing.Size(100, 20)
    Me.txtAuditEntriesToShow.TabIndex = 8
    '
    'TabPage3
    '
    Me.TabPage3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.TabPage3.Controls.Add(Me.Label15)
    Me.TabPage3.Controls.Add(Me.txtCurrentDirectory)
    Me.TabPage3.Controls.Add(Me.Label5)
    Me.TabPage3.Controls.Add(Me.txtAccessRights)
    Me.TabPage3.Controls.Add(Me.Label12)
    Me.TabPage3.Controls.Add(Me.Label11)
    Me.TabPage3.Controls.Add(Me.Label10)
    Me.TabPage3.Controls.Add(Me.txtWinDomain)
    Me.TabPage3.Controls.Add(Me.txtWinComputerName)
    Me.TabPage3.Controls.Add(Me.txtWinUserName)
    Me.TabPage3.Controls.Add(Me.txtLastUpdateID)
    Me.TabPage3.Controls.Add(Me.Label9)
    Me.TabPage3.Controls.Add(Me.txtLoginNumber)
    Me.TabPage3.Controls.Add(Me.Label8)
    Me.TabPage3.Location = New System.Drawing.Point(4, 25)
    Me.TabPage3.Name = "TabPage3"
    Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage3.Size = New System.Drawing.Size(517, 231)
    Me.TabPage3.TabIndex = 3
    Me.TabPage3.Text = "Login Info"
    Me.TabPage3.UseVisualStyleBackColor = True
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(309, 58)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(75, 13)
    Me.Label5.TabIndex = 11
    Me.Label5.Text = "Access Rights"
    '
    'txtAccessRights
    '
    Me.txtAccessRights.Location = New System.Drawing.Point(393, 52)
    Me.txtAccessRights.Name = "txtAccessRights"
    Me.txtAccessRights.ReadOnly = True
    Me.txtAccessRights.Size = New System.Drawing.Size(72, 20)
    Me.txtAccessRights.TabIndex = 10
    '
    'Label12
    '
    Me.Label12.AutoSize = True
    Me.Label12.Location = New System.Drawing.Point(340, 108)
    Me.Label12.Name = "Label12"
    Me.Label12.Size = New System.Drawing.Size(43, 13)
    Me.Label12.TabIndex = 9
    Me.Label12.Text = "Domain"
    '
    'Label11
    '
    Me.Label11.AutoSize = True
    Me.Label11.Location = New System.Drawing.Point(299, 84)
    Me.Label11.Name = "Label11"
    Me.Label11.Size = New System.Drawing.Size(83, 13)
    Me.Label11.TabIndex = 8
    Me.Label11.Text = "Computer Name"
    '
    'Label10
    '
    Me.Label10.AutoSize = True
    Me.Label10.Location = New System.Drawing.Point(276, 31)
    Me.Label10.Name = "Label10"
    Me.Label10.Size = New System.Drawing.Size(107, 13)
    Me.Label10.TabIndex = 7
    Me.Label10.Text = "Windows User Name"
    '
    'txtWinDomain
    '
    Me.txtWinDomain.Location = New System.Drawing.Point(395, 105)
    Me.txtWinDomain.Name = "txtWinDomain"
    Me.txtWinDomain.Size = New System.Drawing.Size(100, 20)
    Me.txtWinDomain.TabIndex = 6
    '
    'txtWinComputerName
    '
    Me.txtWinComputerName.Location = New System.Drawing.Point(394, 78)
    Me.txtWinComputerName.Name = "txtWinComputerName"
    Me.txtWinComputerName.Size = New System.Drawing.Size(100, 20)
    Me.txtWinComputerName.TabIndex = 5
    '
    'txtWinUserName
    '
    Me.txtWinUserName.Enabled = False
    Me.txtWinUserName.Location = New System.Drawing.Point(393, 27)
    Me.txtWinUserName.Name = "txtWinUserName"
    Me.txtWinUserName.Size = New System.Drawing.Size(100, 20)
    Me.txtWinUserName.TabIndex = 4
    '
    'txtLastUpdateID
    '
    Me.txtLastUpdateID.Location = New System.Drawing.Point(107, 51)
    Me.txtLastUpdateID.Name = "txtLastUpdateID"
    Me.txtLastUpdateID.Size = New System.Drawing.Size(100, 20)
    Me.txtLastUpdateID.TabIndex = 3
    '
    'Label9
    '
    Me.Label9.AutoSize = True
    Me.Label9.Location = New System.Drawing.Point(28, 58)
    Me.Label9.Name = "Label9"
    Me.Label9.Size = New System.Drawing.Size(79, 13)
    Me.Label9.TabIndex = 2
    Me.Label9.Text = "Last Update ID"
    '
    'txtLoginNumber
    '
    Me.txtLoginNumber.Location = New System.Drawing.Point(107, 27)
    Me.txtLoginNumber.Name = "txtLoginNumber"
    Me.txtLoginNumber.Size = New System.Drawing.Size(100, 20)
    Me.txtLoginNumber.TabIndex = 1
    '
    'Label8
    '
    Me.Label8.AutoSize = True
    Me.Label8.Location = New System.Drawing.Point(28, 31)
    Me.Label8.Name = "Label8"
    Me.Label8.Size = New System.Drawing.Size(73, 13)
    Me.Label8.TabIndex = 0
    Me.Label8.Text = "Login Number"
    '
    'btnCancel
    '
    Me.btnCancel.Location = New System.Drawing.Point(308, 262)
    Me.btnCancel.Name = "btnCancel"
    Me.btnCancel.Size = New System.Drawing.Size(78, 31)
    Me.btnCancel.TabIndex = 3
    Me.btnCancel.Text = "Cancel"
    Me.btnCancel.UseVisualStyleBackColor = True
    '
    'btnApply
    '
    Me.btnApply.Location = New System.Drawing.Point(415, 262)
    Me.btnApply.Name = "btnApply"
    Me.btnApply.Size = New System.Drawing.Size(93, 31)
    Me.btnApply.TabIndex = 4
    Me.btnApply.Text = "Apply / Close"
    Me.btnApply.UseVisualStyleBackColor = True
    '
    'txtCurrentDirectory
    '
    Me.txtCurrentDirectory.Enabled = False
    Me.txtCurrentDirectory.Location = New System.Drawing.Point(120, 169)
    Me.txtCurrentDirectory.Name = "txtCurrentDirectory"
    Me.txtCurrentDirectory.Size = New System.Drawing.Size(373, 20)
    Me.txtCurrentDirectory.TabIndex = 12
    '
    'Label15
    '
    Me.Label15.AutoSize = True
    Me.Label15.Location = New System.Drawing.Point(28, 172)
    Me.Label15.Name = "Label15"
    Me.Label15.Size = New System.Drawing.Size(86, 13)
    Me.Label15.TabIndex = 13
    Me.Label15.Text = "Current Directory"
    '
    'frmSettings
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(525, 303)
    Me.Controls.Add(Me.btnApply)
    Me.Controls.Add(Me.btnCancel)
    Me.Controls.Add(Me.TabSettings)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
    Me.MinimizeBox = False
    Me.Name = "frmSettings"
    Me.Text = "Settings"
    Me.TabSettings.ResumeLayout(False)
    Me.TabPage1.ResumeLayout(False)
    Me.TabPage1.PerformLayout()
    Me.TabPage2.ResumeLayout(False)
    Me.TabPage2.PerformLayout()
    Me.TabPage3.ResumeLayout(False)
    Me.TabPage3.PerformLayout()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents txtAccessDBPath As System.Windows.Forms.TextBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents TabSettings As System.Windows.Forms.TabControl
  Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
  Friend WithEvents cboDBType As System.Windows.Forms.ComboBox
  Friend WithEvents btnCancel As System.Windows.Forms.Button
  Friend WithEvents btnApply As System.Windows.Forms.Button
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents txtAccessPassword As System.Windows.Forms.TextBox
  Friend WithEvents txtAccessLogin As System.Windows.Forms.TextBox
  Friend WithEvents btnAccessDBPath As System.Windows.Forms.Button
  Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
  Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
  Friend WithEvents Label12 As System.Windows.Forms.Label
  Friend WithEvents Label11 As System.Windows.Forms.Label
  Friend WithEvents Label10 As System.Windows.Forms.Label
  Friend WithEvents txtWinDomain As System.Windows.Forms.TextBox
  Friend WithEvents txtWinComputerName As System.Windows.Forms.TextBox
  Friend WithEvents txtWinUserName As System.Windows.Forms.TextBox
  Friend WithEvents txtLastUpdateID As System.Windows.Forms.TextBox
  Friend WithEvents Label9 As System.Windows.Forms.Label
  Friend WithEvents txtLoginNumber As System.Windows.Forms.TextBox
  Friend WithEvents Label8 As System.Windows.Forms.Label
  Friend WithEvents Label13 As System.Windows.Forms.Label
  Friend WithEvents txtAuditEntriesToShow As System.Windows.Forms.TextBox
  Friend WithEvents chkAskForUpdateConfirmation As System.Windows.Forms.CheckBox
  Friend WithEvents chkDontShowSplash As System.Windows.Forms.CheckBox
  Friend WithEvents chkMultiUserAccess As System.Windows.Forms.CheckBox
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents txtAccessRights As System.Windows.Forms.TextBox
  Friend WithEvents cboMonthsOfAudit As System.Windows.Forms.ComboBox
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
  Friend WithEvents Label14 As System.Windows.Forms.Label
  Friend WithEvents Label7 As System.Windows.Forms.Label
  Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
  Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
  Friend WithEvents Label15 As System.Windows.Forms.Label
  Friend WithEvents txtCurrentDirectory As System.Windows.Forms.TextBox
End Class
