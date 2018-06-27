Imports Microsoft.Win32
'Imports System.Runtime.InteropServices
Imports System.IO
'Imports System.Text
Imports System.Windows.Forms
Imports System.Data
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class AppSettings

  Public Shared cnn As New ADODB.Connection
  Public Shared rst As New ADODB.Recordset
  Public Shared cmd As New ADODB.Command

  Public Shared DBType As String = "DBType"
  Public Shared AccessDBPath As String = "AccessDBPath"
  Public Shared AccessLoginID As String = "AccessLoginID"
  Public Shared AcccessLoginPassword As String = "AcccessLoginPassword"
  Public Shared MultiUserAccess As String = "MultiUserAccess"
  Public Shared AccessRights As String = "AccessRights"

  Public Shared AuditEntriesToShow As String = "AuditEntriesToShow"
  Public Shared bAskForUpdateConfirm As String = "bAskForUpdateConfirm"

  Public Shared LoginName As String = "LoginName"
  Public Shared ComputerName As String = "ComputerName"
  Public Shared DomainName As String = "DomainName"
  Public Shared LoginNumber As String = "LoginNumber"
  Public Shared LastUpdateID As String = "LastUpdateID"

  Public Shared sLoginName As String
  Public Shared sComputerName As String
  Public Shared sDomainName As String
  Public Shared sLoginNumber As String

  Public Shared RegDatabase As RegistryKey
  Public Shared RegFramework As RegistryKey
  Public Shared RegOptions As RegistryKey
  Public Shared RegLoginInfo As RegistryKey

  Private Shared Reg As RegistryKey
  Private Shared keyValue As String
  Private Shared currPath As String
  'Private Shared currPath As String = Environment.CurrentDirectory
  Private Shared SettingsPath As String

  Shared Function GetDLLPath()
    Dim ass As System.Reflection.Assembly = System.Reflection.Assembly.GetCallingAssembly
    'Dim s As String = ass.CodeBase
    Dim nLoc As Integer
    Dim sCurrPath As String
    Dim FullPath As String = ass.Location
    nLoc = InStrRev(FullPath, "\")
    sCurrPath = Mid(FullPath, 1, nLoc - 1)
    'Debug.Print(ass.GetName.ToString)
    'SetSetting("CurrPath", sCurrPath)
    Return sCurrPath
  End Function
  Public Shared Sub SetUserDetails()
    Dim objNet As Object
    Dim strUserName, strComputerName, strUserDomain As String
    objNet = CreateObject("WScript.NetWork")

    strUserName = objNet.UserName
    strComputerName = objNet.ComputerName
    strUserDomain = objNet.UserDomain

    SetSetting("LoginName", strUserName)
    SetSetting("ComputerName", strComputerName)
    SetSetting("DomainName", strUserDomain)
  End Sub

  Public Shared Sub InitializeSettings()
    SetSetting("DBType", "MSAccess")
    SetSetting("AccessDBPath", "C:\" & OMGlobals.APPLongName & "\" & "PA_Institute.mdb")
    SetSetting("AccessLoginID", "")
    SetSetting("AcccessLoginPassword", "")

    SetSetting("MonthsToAccessHistory", "2")
    SetSetting("AuditEntriesToShow", "1000")
    SetSetting("bAskForUpdateConfirmation", "False")
  End Sub

  Public Shared Function GetAppConnString() As String
    'Gets the connection string for the database based on AppSettings.
    'Dim strTMUser As String
    'Dim strTMpw As String

    'strTMUser = "TMAS"
    'strTMpw = AppSettings.sAcccessLoginPassword
    'GetAppConnString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    '           "Data Source= " & AppSettings.sAccessDBPath & ";User Id=admin" & _
    '                 ";Jet OLEDB:Database Password=" & AppSettings.sAcccessLoginPassword & ";"
    'GetAppConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\PA_Institute\PA_Institute.mdb"
    If GetSetting("AccessDBPath") = "" Then
      SetSetting("AccessDBPath", currPath & "\PA_Institute.mdb")
    End If
    GetAppConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GetSetting("AccessDBPath")

    strDBO = ""
    'nd If
    'AccessDBPath
  End Function

  Public Shared Sub WriteToErrorLog(ByVal sMessage As String)
    'check and make the directory if necessary; this is set to look in 
    'the application folder, you may wish to place the error log in 
    'another location depending upon the user's role and write access to 
    'different areas of the file system
    If Not System.IO.Directory.Exists("C:\" & OMGlobals.APPLongName & _
    "\Errors\") Then
      System.IO.Directory.CreateDirectory("C:\" & OMGlobals.APPLongName & _
        "\Errors\")
    End If

    'check the file
    Dim fs As FileStream = New FileStream("C:\" & OMGlobals.APPLongName & _
    "\Errors\errlog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
    Dim s As StreamWriter = New StreamWriter(fs)
    s.Close()
    fs.Close()

    'log it
    Dim fs1 As FileStream = New FileStream("C:\" & OMGlobals.APPLongName & _
    "\Errors\errlog.txt", FileMode.Append, FileAccess.Write)
    Dim s1 As StreamWriter = New StreamWriter(fs1)
    s1.Write(Now & ":, " & sMessage & vbCrLf)
    's1.Write("Title: " & title & vbCrLf)
    's1.Write("Message: " & msg & vbCrLf)
    's1.Write("StackTrace: " & stkTrace & vbCrLf)
    's1.Write("Date/Time: " & DateTime.Now.ToString() & vbCrLf)
    's1.Write("================================================" & vbCrLf)
    s1.Close()
    fs1.Close()

  End Sub

  Public Shared Function GetSetting(ByVal Key As String) As String
    Dim sReturn As String = String.Empty
    Dim dsSettings As New DataSet
    currPath = "C:" 'GetDLLPath()
    If System.IO.File.Exists(currPath & "\Settings.xml") Then
      SettingsPath = currPath & "\Settings.xml"
      dsSettings.ReadXml(SettingsPath)
    ElseIf System.IO.File.Exists("C:\" & OMGlobals.APPLongName & "\Settings.xml") Then
      SettingsPath = "C:\" & OMGlobals.APPLongName & "\Settings.xml"
      dsSettings.ReadXml(SettingsPath)
    Else
      dsSettings.Tables.Add("Settings")
      dsSettings.Tables(0).Columns.Add("Key", GetType(String))
      dsSettings.Tables(0).Columns.Add("Value", GetType(String))
    End If

    Dim dr() As DataRow = dsSettings.Tables("Settings").Select("Key = '" & Key & "'")
    If dr.Length = 1 Then sReturn = dr(0)("Value").ToString

    Return sReturn
  End Function

  Public Shared Sub SetSetting(ByVal Key As String, ByVal Value As String)
    Dim dsSettings As New DataSet
    currPath = "C:" 'GetDLLPath()
    If System.IO.File.Exists(currPath & "\Settings.xml") Then
      SettingsPath = currPath & "\Settings.xml"
      dsSettings.ReadXml(SettingsPath)
    ElseIf System.IO.File.Exists("C:\" & OMGlobals.APPLongName & "\Settings.xml") Then
      SettingsPath = "C:\" & OMGlobals.APPLongName & "\Settings.xml"
      dsSettings.ReadXml(SettingsPath)
    Else
      dsSettings.Tables.Add("Settings")
      dsSettings.Tables(0).Columns.Add("Key", GetType(String))
      dsSettings.Tables(0).Columns.Add("Value", GetType(String))
    End If

    Dim dr() As DataRow = dsSettings.Tables(0).Select("Key = '" & Key & "'")
    If dr.Length = 1 Then
      dr(0)("Value") = Value
    Else
      Dim drSetting As DataRow = dsSettings.Tables("Settings").NewRow
      drSetting("Key") = Key
      drSetting("Value") = Value
      dsSettings.Tables("Settings").Rows.Add(drSetting)
    End If
    dsSettings.WriteXml(SettingsPath)
  End Sub

  'Public Function GetPASetting(ByVal key As String) As String
  '  Return GetSetting(key)
  'End Function

  'Public Sub SetPASetting(ByVal key As String, ByVal value As String)
  '  Call SetSetting(key, value)
  'End Sub

End Class
