Option Strict Off
Option Explicit On
Imports System.Data
Public Class DBStrings

  ' Dim Cnn As New OleDb.OleDbConnection
  ' Dim Cmd As New OleDb.OleDbCommand
  Public Shared cnn1 As New OleDb.OleDbConnection
  Public Shared cmd1 As New OleDb.OleDbCommand
  Public Shared dr1 As OleDb.OleDbDataReader
  'Public Shared strConnFrameworkString As String
  'Const adOpenStatic As Short = 3
  'Const adLockOptimistic As Short = 3
  'Const adUseClient As Short = 3

    'Public Shared rst As ADODB.Recordset
    'Public Shared cnn As ADODB.Connection
    'Public Shared cmd As New ADODB.Command
  Public Shared gstrSQLCall As String
  Public Shared gConnectionString As String
  Public Shared strDBO As String
  Dim Dr As OleDb.OleDbDataReader
  'Const adOpenStatic As Short = 3

  '  Public Shared Function GetAppConnString() As String
  '    'Gets the connection string for the database based on settings.
  '    Dim strTMpw As String

  '    strTMpw = AppSettings.sAcccessLoginPassword

  '    'Else    'use mdb database
  '    'GetAppConnString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
  '    '"Data Source= " & AppSettings.sAccessDBPath & ";" & _
  '    '"Jet OLEDB:Database Password=" & strTMpw

  '        GetAppConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\PA_FrameWork\PA_Framework_VB.mdb;User Id=admin;Password=;"
  '        gConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\PA_FrameWork\PA_Framework_VB.mdb;User Id=admin;Password=;"
  '    strDBO = ""
  '    'End If
  '  End Function
End Class
