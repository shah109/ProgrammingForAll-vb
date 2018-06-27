Imports Microsoft.Win32

Public Class AppSettings
  Public Shared sDBType As String
  Public Shared sAccessDBPath As String
  Public Shared sAccessLoginID As String
  Public Shared sAcccessLoginPassword As String
  Public Shared EnableAccessRights As Boolean
  Public Shared AccessRights As Integer

  Public Shared sProjectFilePath As String
  Public Shared sImportPath As String
  Public Shared sExportPath As String

  Public Shared LastUpdateID As Integer
  Public Shared nLastUpdateID As Integer
  Public Shared bFirstLaunch As Boolean
  Public Shared sAuditEntriesToShow As String
  Public Shared bAskForUpdateConfirmation As Boolean
  Public Shared bDontShowSplash As Boolean
  Public Shared sLoginName As String
  Public Shared sComputerName As String
  Public Shared sDomainName As String
  Public Shared sLoginNumber As String
  Public Shared sLastUpdateID As String

  'Public Shared abc As String
  'Public Shared abc As String
  'Public Shared abc As String
  'Public Shared abc As String
  'Public Shared abc As String
  'Public Shared abc As String
  'Public Shared abc As String
  'Public Shared abc As String
  'Public Shared abc As String
  'Public Shared abc As String

  Public Shared Sub SaveIntoRegistry()
    Dim Reg As RegistryKey
    Dim keyValue As String

    keyValue = "SOFTWARE\PAFramework\Database"
    Reg = Registry.CurrentUser.OpenSubKey(keyValue, True)
    Reg.SetValue("DBType", sDBType)
    Reg.SetValue("DBPath", sAccessDBPath)
    Reg.SetValue("Login", sAccessLoginID)
    Reg.SetValue("LoginPW", sAcccessLoginPassword)
    Reg.Close()

    '''''''''''''Start Framework
    keyValue = "SOFTWARE\PAFramework\Framework"
    Reg = Registry.CurrentUser.OpenSubKey(keyValue, True)
    Reg.SetValue("ProjectPath", sProjectFilePath)
    Reg.SetValue("ImportPath", sImportPath)
    Reg.SetValue("ExportPath", sExportPath)
    Reg.Close()
    ' Options

    keyValue = "SOFTWARE\PAFramework\Options"
    Reg = Registry.CurrentUser.OpenSubKey(keyValue, True)
    Reg.SetValue("AuditEntries", sAuditEntriesToShow)
    Reg.SetValue("AskUpdateConfirmation", bAskForUpdateConfirmation)
    Reg.SetValue("DontShowAboutDlg", bDontShowSplash)
    Reg.Close()

    ' LoginInfo
    keyValue = "SOFTWARE\PAFramework\LoginInfo"
    Reg = Registry.CurrentUser.OpenSubKey(keyValue, True)
    Reg.SetValue("UserName", sLoginName)
    Reg.SetValue("ComputerName", sComputerName)
    Reg.SetValue("DomainName", sDomainName)
    Reg.SetValue("LoginNumber", sLoginNumber)
    Reg.SetValue("LastUpdateID", sLastUpdateID)
    Reg.Close()
  End Sub

  Public Shared Sub LoadFromRegistry()
    Dim Reg As RegistryKey
    Dim keyValue As String

    keyValue = "SOFTWARE\PAFramework\Database"
    Reg = Registry.CurrentUser.OpenSubKey(keyValue, False)
    sDBType = Reg.GetValue("DBType")
    sAccessDBPath = Reg.GetValue("DBPath")
    sAccessLoginID = Reg.GetValue("Login")
    sAcccessLoginPassword = Reg.GetValue("LoginPW")
    Reg.Close()

    '''''''''''''Start Framework
    keyValue = "SOFTWARE\PAFramework\Framework"
    Reg = Registry.CurrentUser.OpenSubKey(keyValue, False)
    sProjectFilePath = Reg.GetValue("ProjectPath")
    sImportPath = Reg.GetValue("ImportPath")
    sExportPath = Reg.GetValue("ExportPath")
    Reg.Close()
    ' Options

    keyValue = "SOFTWARE\PAFramework\Options"
    Reg = Registry.CurrentUser.OpenSubKey(keyValue, False)
    sAuditEntriesToShow = Reg.GetValue("AuditEntries")
    bAskForUpdateConfirmation = Reg.GetValue("AskUpdateConfirmation")
    bDontShowSplash = Reg.GetValue("DontShowAboutDlg")
    Reg.Close()

    ' LoginInfo
    keyValue = "SOFTWARE\\PAFramework\\LoginInfo"
    Reg = Registry.CurrentUser.OpenSubKey(keyValue, False)
    sLoginName = Reg.GetValue("UserName")
    sComputerName = Reg.GetValue("ComputerName")
    sDomainName = Reg.GetValue("DomainName")
    sLoginNumber = Reg.GetValue("LoginNumber")
    sLastUpdateID = Reg.GetValue("LastUpdateID")
    Reg.Close()
  End Sub



  Public Shared Sub InitializeSettings()
        'this function initializes the registry settings if Framework is being launched for the first time.
    Dim objNet As Object
    'Dim sSite As String
    Dim strUserName, strComputerName, strUserDomain As String
    objNet = CreateObject("WScript.NetWork")

    strUserName = objNet.UserName
    strComputerName = objNet.ComputerName
    strUserDomain = objNet.UserDomain

    Dim reg As RegistryKey

    'start database

    reg = Registry.CurrentUser.CreateSubKey("SOFTWARE\PAFramework\Database")
    reg.SetValue("DBType", "")
    reg.SetValue("DBPath", "")
    reg.SetValue("Login", "")
    reg.SetValue("LoginPW", "")
    reg.Close()

    '''''''''''''Start Framework
    reg = Registry.CurrentUser.CreateSubKey("SOFTWARE\\PAFramework\\Framework")
    reg.SetValue("ProjectPath", "")
    reg.SetValue("ImportPath", "")
    reg.SetValue("ExportPath", "")
    reg.Close()        ' Options
    reg = Registry.CurrentUser.CreateSubKey("SOFTWARE\\PAFramework\\Options")
    reg.SetValue("AuditEntries", "")
        reg.SetValue("AskUpdateConfirmation", "1")
    reg.SetValue("DontShowAboutDlg", "1")
    reg.Close()

    ' LoginInfo
    reg = Registry.CurrentUser.CreateSubKey("SOFTWARE\\PAFramework\\LoginInfo")
    reg.SetValue("UserName", strUserDomain)
    reg.SetValue("ComputerName", strComputerName)
    reg.SetValue("DomainName", strUserDomain)
    reg.SetValue("LoginNumber", "")
    reg.SetValue("LastUpdateID", "")
    reg.Close()

  End Sub
End Class
