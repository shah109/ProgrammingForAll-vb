Imports Microsoft.Win32
Imports PA_Framework_OM

Public Class Form2


  Private Sub Form2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Dim keyValue As String = "SOFTWARE\PAFramework\"
    If (Registry.CurrentUser.OpenSubKey(keyValue, False)) Is Nothing Then
      Call PA_Framework_Lib.AppSettings.InitializeSettings()  'Launching for the first time - hence initialize
    End If
    AppSettings.LoadFromRegistry()
    Call PA_Framework_OM.DBLoad()
  End Sub
End Class