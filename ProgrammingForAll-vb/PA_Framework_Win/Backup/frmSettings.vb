Imports PA_Framework_OM
Public Class frmSettings

  Private Overloads Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click
    Call Me.ApplyIntoSettings()
    Me.Close()
  End Sub

  Private Overloads Sub frmSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Me.LoadFromSettings()  'load AppSettings from the AppSettings object this is the correct way
    'Me.bFirstLaunch = False  'Now set the first launch to false because the sitecoll has been created

  End Sub

  Private Overloads Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
    Me.Close()
  End Sub

  Function LoadFromSettings() As nullable
    LoadFromSettings = Nothing
    Dim objNet As Object
    Dim strUserName, strComputerName, strUserDomain As String
    objNet = CreateObject("WScript.NetWork")

    strUserName = objNet.UserName
    strComputerName = objNet.ComputerName
    strUserDomain = objNet.UserDomain

    Me.cboDBType.Text = AppSettings.GetSetting(AppSettings.DBType)

    Me.txtAccessDBPath.Text = AppSettings.GetSetting("AccessDBPath")
    Me.txtAccessRights.Text = AppSettings.GetSetting("AccessRights")
    Me.txtAccessLogin.Text = AppSettings.GetSetting("AccessLoginID")
    Me.txtAccessPassword.Text = AppSettings.GetSetting("AcccessLoginPassword")

    Me.chkMultiUserAccess.Checked = AppSettings.GetSetting("MultiUserAccess")

    Me.txtAuditEntriesToShow.Text = AppSettings.GetSetting("AuditEntriesToShow")
    Me.cboMonthsOfAudit.Text = AppSettings.GetSetting("MonthsToAccessHistory")

    Me.chkAskForUpdateConfirmation.Checked = AppSettings.GetSetting("bAskForUpdateConfirmation")

    Me.txtWinUserName.Text = strUserName
    Me.txtWinComputerName.Text = strComputerName
    Me.txtWinDomain.Text = strUserDomain
    Me.txtLoginNumber.Text = AppSettings.GetSetting("LoginNumber")
    Me.txtLastUpdateID.Text = AppSettings.GetSetting("LastUpdateID")
    Me.txtCurrentDirectory.Text = AppSettings.GetDLLPath
  End Function

  Public Sub ApplyIntoSettings()
    'save AppSettings in the AppSettings object

    AppSettings.SetSetting(AppSettings.AccessDBPath, Me.cboDBType.Text)
    AppSettings.SetSetting(AppSettings.AccessDBPath, Me.txtAccessDBPath.Text)
    AppSettings.SetSetting(AppSettings.AccessLoginID, Me.txtAccessLogin.Text)
    AppSettings.SetSetting(AppSettings.AcccessLoginPassword, Me.txtAccessPassword.Text)
    AppSettings.SetSetting(AppSettings.MultiUserAccess, Me.chkMultiUserAccess.Checked)
    ''''''''''''''Start Framework
    'PA_Framework_OM.AppSettings.sProjectFilePath = Me.txtProjectFilePath.Text
    'PA_Framework_OM.AppSettings.sImportPath = Me.txtImportPath.Text
    'PA_Framework_OM.AppSettings.sExportPath = Me.txtExportPath.Text

    ' Options
    AppSettings.SetSetting("AuditEntriesToShow", Me.txtAuditEntriesToShow.Text)
    AppSettings.SetSetting("MonthsToAccessHistory", Me.cboMonthsOfAudit.Text)

    AppSettings.SetSetting("bAskForUpdateConfirmation", CStr(Me.chkAskForUpdateConfirmation.Checked))
    'PA_Framework_OM.AppSettings.bDontShowSplash = Me.chkDontShowSplash.Enabled

    ' LoginInfo
    AppSettings.SetSetting(AppSettings.LoginName, Me.txtWinUserName.Text)
    AppSettings.SetSetting(AppSettings.AccessRights, Me.txtAccessRights.Text)
    AppSettings.SetSetting(AppSettings.ComputerName, Me.txtWinComputerName.Text)
    AppSettings.SetSetting(AppSettings.DomainName, Me.txtWinDomain.Text)
    AppSettings.SetSetting(AppSettings.LoginNumber, Me.txtLoginNumber.Text)
    AppSettings.SetSetting(AppSettings.LastUpdateID, Me.txtLastUpdateID.Text)

  End Sub

  Private Sub btnAccessDBPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAccessDBPath.Click
    Dim folderdlg As New OpenFileDialog
    folderdlg.InitialDirectory = Me.txtAccessDBPath.Text
    folderdlg.ShowDialog()
    If folderdlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
      Me.txtAccessDBPath.Text = folderdlg.FileName.ToString
    End If
  End Sub
End Class