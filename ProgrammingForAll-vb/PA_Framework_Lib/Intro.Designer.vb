<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmIntro
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
    Me.Label1 = New System.Windows.Forms.Label
    Me.btnOK = New System.Windows.Forms.Button
    Me.Label2 = New System.Windows.Forms.Label
    Me.txtLicenseStatus = New System.Windows.Forms.TextBox
    Me.Label3 = New System.Windows.Forms.Label
    Me.txtLicenseFile = New System.Windows.Forms.TextBox
    Me.Label4 = New System.Windows.Forms.Label
    Me.txtLicensedTo = New System.Windows.Forms.TextBox
    Me.btnOpenLicFile = New System.Windows.Forms.Button
    Me.btnValidateLicense = New System.Windows.Forms.Button
    Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
    Me.SuspendLayout()
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label1.Location = New System.Drawing.Point(97, 11)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(189, 20)
    Me.Label1.TabIndex = 0
    Me.Label1.Text = "PA Framework Library "
    Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
    '
    'btnOK
    '
    Me.btnOK.Location = New System.Drawing.Point(267, 192)
    Me.btnOK.Name = "btnOK"
    Me.btnOK.Size = New System.Drawing.Size(75, 23)
    Me.btnOK.TabIndex = 1
    Me.btnOK.Text = "OK"
    Me.btnOK.UseVisualStyleBackColor = True
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(13, 49)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(80, 13)
    Me.Label2.TabIndex = 2
    Me.Label2.Text = "License Status:"
    '
    'txtLicenseStatus
    '
    Me.txtLicenseStatus.Location = New System.Drawing.Point(93, 46)
    Me.txtLicenseStatus.Name = "txtLicenseStatus"
    Me.txtLicenseStatus.ReadOnly = True
    Me.txtLicenseStatus.Size = New System.Drawing.Size(245, 20)
    Me.txtLicenseStatus.TabIndex = 3
    Me.txtLicenseStatus.Text = "Trial"
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(13, 140)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(66, 13)
    Me.Label3.TabIndex = 4
    Me.Label3.Text = "License File:"
    '
    'txtLicenseFile
    '
    Me.txtLicenseFile.Location = New System.Drawing.Point(93, 134)
    Me.txtLicenseFile.Multiline = True
    Me.txtLicenseFile.Name = "txtLicenseFile"
    Me.txtLicenseFile.Size = New System.Drawing.Size(214, 41)
    Me.txtLicenseFile.TabIndex = 5
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(13, 72)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(65, 13)
    Me.Label4.TabIndex = 6
    Me.Label4.Text = "Licensed to:"
    '
    'txtLicensedTo
    '
    Me.txtLicensedTo.Location = New System.Drawing.Point(93, 72)
    Me.txtLicensedTo.Multiline = True
    Me.txtLicensedTo.Name = "txtLicensedTo"
    Me.txtLicensedTo.ReadOnly = True
    Me.txtLicensedTo.Size = New System.Drawing.Size(245, 56)
    Me.txtLicensedTo.TabIndex = 7
    Me.txtLicensedTo.Text = "Evaluation User " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Production version of the Library will be available soon."
    '
    'btnOpenLicFile
    '
    Me.btnOpenLicFile.Enabled = False
    Me.btnOpenLicFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.btnOpenLicFile.Location = New System.Drawing.Point(313, 142)
    Me.btnOpenLicFile.Name = "btnOpenLicFile"
    Me.btnOpenLicFile.Size = New System.Drawing.Size(28, 30)
    Me.btnOpenLicFile.TabIndex = 8
    Me.btnOpenLicFile.Text = "..."
    Me.btnOpenLicFile.UseVisualStyleBackColor = True
    '
    'btnValidateLicense
    '
    Me.btnValidateLicense.Location = New System.Drawing.Point(139, 193)
    Me.btnValidateLicense.Name = "btnValidateLicense"
    Me.btnValidateLicense.Size = New System.Drawing.Size(103, 23)
    Me.btnValidateLicense.TabIndex = 9
    Me.btnValidateLicense.Text = "Validate License"
    Me.btnValidateLicense.UseVisualStyleBackColor = True
    '
    'frmIntro
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(357, 229)
    Me.Controls.Add(Me.btnValidateLicense)
    Me.Controls.Add(Me.btnOpenLicFile)
    Me.Controls.Add(Me.txtLicensedTo)
    Me.Controls.Add(Me.Label4)
    Me.Controls.Add(Me.txtLicenseFile)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.txtLicenseStatus)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.btnOK)
    Me.Controls.Add(Me.Label1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.Name = "frmIntro"
    Me.Text = "PA Framework"
    Me.TopMost = True
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents btnOK As System.Windows.Forms.Button
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents txtLicenseStatus As System.Windows.Forms.TextBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents txtLicenseFile As System.Windows.Forms.TextBox
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents txtLicensedTo As System.Windows.Forms.TextBox
  Friend WithEvents btnOpenLicFile As System.Windows.Forms.Button
  Friend WithEvents btnValidateLicense As System.Windows.Forms.Button
  Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
End Class
