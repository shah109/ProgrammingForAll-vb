Imports Microsoft.Win32
Imports PA_Framework_OM
'Imports adodb
Imports PA_Framework_OM.OMGlobals

Public Class frmCodeGenerator

  Dim omg As New OMGlobals
  Dim WithEvents PAProjectChldProjectEntitiesBindingSource As New BindingSource
  Dim ProjectEntityChildEntityItemsBindingSource As New BindingSource
  Dim dbCheckBindingSource As New BindingSource

  Private Sub frmCodeGenerator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Dim sUser As String
    'If Not System.IO.File.Exists("C:\" & OMGlobals.APPLongName & "\Settings.xml") Then
    '  AppSettings.InitializeSettings()
    'End If
    AppSettings.SetUserDetails()
    Call omg.DBLoad()
    Call omg.PassFormToOM(Me)
    'StatusStrip1.
    sUser = AppSettings.GetSetting("LoginName")

    cboPAProjects.DataSource = cPAProjects

    Me.dgvPrjEntityItems.DataSource = ProjectEntityChildEntityItemsBindingSource
    Me.dgvPrjEntities.DataSource = PAProjectChldProjectEntitiesBindingSource
    Me.dgvDBCheck.DataSource = dbCheckBindingSource

  End Sub

  Private Sub cboPAProjects_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPAProjects.SelectedIndexChanged
    Dim strSelection As String
    Dim pr As PAProject
    strSelection = cboPAProjects.Text
    pr = cPAProjects.GetProjectByName(strSelection)
    PAProjectChldProjectEntitiesBindingSource.DataSource = pr.ChildprojectEntities
  End Sub

  Private Sub PAProjectChldProjectEntitiesBindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles PAProjectChldProjectEntitiesBindingSource.CurrentChanged
    currePrjEnt_ = PAProjectChldProjectEntitiesBindingSource.Current
    ProjectEntityChildEntityItemsBindingSource.DataSource = currePrjEnt_.ChildEntities("ProjectEntityItems")
  End Sub

  Public Sub CreateDatabaseTable()
    'Dim conn As New OleDbConnection
    'Dim cmd As OleDbCommand

    'Dim res As New List(Of String)
    'conn.ConnectionString = PA_Framework_OM.PASettings.GetAppConnString
    'Dim cmdString As String = "Create table aabb"
    'cmd = New OleDbCommand(cmdString, conn)
    'conn.Open()
    'res = conn.GetSchema(
    'cmd.ExecuteNonQuery()
    ''Dim myreader As OleDbDataReader
    ''myreader = cmd.ExecuteReader(Data.CommandBehavior.CloseConnection)
    ''While myreader.Read
    ''  res.Add(myreader("ProjectName").ToString)

    ''End While
  End Sub

  Private Sub btnValidateEntityItemsWithDatabase_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnValidateEntityItemsWithDatabase.Click
    Me.txtMissingFields.Clear()
    Me.txtMissingFields.Text = PA_DBCreation.ValidateTablesAndColumns(cboPAProjects.Text)
  End Sub
  

  Private Sub btnShowMainForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowMainForm.Click
    Dim frmM As New frmPAInstitute
    frmM.Show()
    'frmPAInstitute.
  End Sub

  Private Sub btnCreateEntityItemsFromDatabase_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateEntityItemsFromDatabase.Click
    'Call PA_DBCreation.UpdateDBField()
    Call PA_DBCreation.CreateEntityItemsFromDatabase(frmPAInstitute.nPrevUpdateID)
  End Sub
End Class

