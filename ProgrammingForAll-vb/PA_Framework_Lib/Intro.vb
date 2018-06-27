Public Class frmIntro
    Public Shared gen As New GenEBADBUpdates
    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim sCode As String
        GenEBADBUpdates.strLicenseKey = Me.txtLicenseFile.Text.Trim
        sCode = GenEBADBUpdates.Decode(GenEBADBUpdates.strLicenseKey)
        'If sCode.ToLower <> "error" Then
        '    gen.SaveinReg("lic", GenEBADBUpdates.strLicenseKey)
        '    gen.Initialize(GenEBADBUpdates.ExApp)
        'End If
        'GenEBADBUpdates.ValidateGeneba(GenEBADBUpdates.strLicenseKey)
        ''to call a function like gen.validategeneba in genebaupdates.
        Me.Close()
    End Sub

    'Private Sub frmIntro_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    '    Call gen.ValidateGeneba(Me.txtLicenseKey.Text)
    'End Sub

    Private Sub frmIntro_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'If Me.txtLicenseStatus.Text.Contains("Valid") Then
        'Me.txtLicenseFile.Text = GenEBADBUpdates.strLicenseKey
        'End If
    End Sub

  Private Sub btnOpenLicFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenLicFile.Click
    OpenFileDialog1.Title = "Please Select a File"
    OpenFileDialog1.InitialDirectory = "C:\"

    If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
      Me.txtLicenseFile.Text = OpenFileDialog1.FileName
    End If

  End Sub

  Private Sub btnValidateLicense_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnValidateLicense.Click

  End Sub

End Class