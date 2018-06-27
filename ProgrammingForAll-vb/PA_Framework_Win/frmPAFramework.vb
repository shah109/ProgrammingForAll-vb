Option Explicit On
Imports Microsoft.Win32
Imports PA_Framework_OM
Imports PA_Framework_Lib

Public Class frmPAFramework

  Dim aa As TabControl
  Dim ctr As Control

  Private Sub SettingsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SettingsToolStripMenuItem.Click
    frmSettings.Show()
  End Sub

  Private Sub frmPAFramework_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Dim keyValue As String = "SOFTWARE\PAFramework\"
    If (Registry.CurrentUser.OpenSubKey(keyValue, False)) Is Nothing Then
      Call PA_Framework_Lib.AppSettings.InitializeSettings()  'Launching for the first time - hence initialize
    End If
    PA_Framework_Lib.AppSettings.LoadFromRegistry()
    'Dim bds As New Binding

    'For Each ctr In Me.Controls
    '  Debug.Print(ctr.Name & "  " & ctr.GetType.ToString)

    'Next

  End Sub

  Private Sub LoadFromDatabaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadFromDatabaseToolStripMenuItem.Click
    OMGlobals.DBLoad()
    'Me.EntityAsDataGridView.Update()
    'Me.dgvProjects.Update()
  End Sub

  Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
    Me.Close()
  End Sub
  'EntityAs Seciton

  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'page functions
  'Project functions
  Private Sub TabProjects_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabProjects.Enter
    Call PA_Framework_Lib.MGlobals.CallDBLoadIfNeeded()
    ProjectsBindingSource.DataSource = cProjects
    Me.dgvProjects.DataSource = ProjectsBindingSource
    OMGlobals.currObjCollection = OMGlobals.cProjects
  End Sub

  Private Sub dgvProjects_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvProjects.SelectionChanged

  End Sub

  Private Sub btn_ePrj_Update_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ePrjt_Update.Click
    btn_ePrjtEnt_Update.Text = "Updating"
    'currePrjt_ = cProjects.item(Me.txt_ePrj_ID.Text)
    currePrjt_ = ProjectsBindingSource.Current
    Call UIFunctions.DBUpdate("Update", cProjects, currePrjt_)
    btn_ePrjtEnt_Update.Text = "Update"
  End Sub

  Private Sub btn_ePrjt_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ePrjt_Add.Click
    Dim np = New Project
    'np.ID = "12"
    np.ProjectName = "new"
    'cProjects.Add(np)
    '
    ProjectsBindingSource.Add(np)
    Debug.Print(cProjects.Count)
    'MGlobals.currePrjEntItm_ = cProjectEntityItems.item(Me.txt_ePrjEntItm_ID.Text)
    Call UIFunctions.DBUpdate("Add", cProjects, np)
    ProjectsBindingSource.MoveLast()

  End Sub


  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Project entity functions

  Private Sub TabProjectEntities_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabProjectEntities.Enter
    Call PA_Framework_Lib.MGlobals.CallDBLoadIfNeeded()
    ProjectEntitiesBindingSource.DataSource = cProjectEntities
    Me.ProjectEntitiesDataGridView.DataSource = ProjectEntitiesBindingSource
    OMGlobals.currObjCollection = OMGlobals.cProjectEntities
  End Sub

  Private Sub btn_ePrjtEnt_Update_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ePrjtEnt_Update.Click
    btn_ePrjtEnt_Update.Text = "Updating"
    OMGlobals.currePrjEnt_ = ProjectEntitiesBindingSource.Current
    Call UIFunctions.DBUpdate("Update", cProjectEntities, OMGlobals.currePrjEnt_)
    'btn_ePrjtEnt_Update.Text = "Update"
  End Sub

  Private Sub btn_ePrjtEnt_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ePrjtEnt_Add.Click
    Dim np = New ProjectEntity
    np.EntityName = "new"
    ProjectEntitiesBindingSource.Add(np)
    ProjectEntitiesBindingSource.MoveLast()
    Call UIFunctions.DBUpdate("Add", cProjectEntities, np)
  End Sub
 

  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Project Entity Items functions

  Private Sub TabEntityItems_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabEntityItems.Enter
    Call PA_Framework_Lib.MGlobals.CallDBLoadIfNeeded()
    ProjectEntityItemsBindingSource.DataSource = cProjectEntityItems
    Me.ProjectEntityItemsDataGridView.DataSource = ProjectEntityItemsBindingSource

    OMGlobals.currObjCollection = OMGlobals.cProjectEntityItems
  End Sub

  Private Sub btnAdd__ePrjEntItm__Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ePrjEntItm_Add.Click
    Dim np = New ProjectEntityItem
    np.ItemDBName = "new"
    ProjectEntityItemsBindingSource.Add(np)
    ProjectEntityItemsBindingSource.MoveLast()
    Call UIFunctions.DBUpdate("Add", cProjectEntityItems, np)
  End Sub

  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  Private Sub TabEntityAs_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabEntityAs.Enter

    'Call PA_Framework_Lib.MGlobals.CallDBLoadIfNeeded()
    'oMe.EntityAsBindingSource.DataSource = cEntityAs
    'Me.EntityAsDataGridView.DataSource = EntityAsBindingSource
    'OMGlobals.currObjCollection = MGlobals.cEntityAs
  End Sub


  'Private Sub btn_eA_Update_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_eA_Update.Click
  '  MGlobals.curreA_ = cEntityAs.Item(Me.txt_eA_ID.Text)
  '  Call UIFunctions.DBUpdate("Update", cEntityAs, MGlobals.curreA_)
  'End Sub

  Private Sub btnUpdate_eHist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_eHist.Click
    Call TabPageChangeHistory_Enter(sender, e)
  End Sub

  Private Sub TabPageChangeHistory_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPageChangeHistory.Enter
    'ChangeHistorysBindingSource.Clear()
    PA_Framework_OM.cChangeHistorys = New ChangeHistorys
    OMGlobals.cChangeHistorys.Load()
    'Dim Change As New BindingSource
    ChangeHistorysBindingSource.DataSource = PA_Framework_OM.cChangeHistorys
    Me.dgvChangeHistory.DataSource = ChangeHistorysBindingSource
    'MGlobals.currObjCollection = MGlobals.cChangeHistorys

    'MGlobals.currObjCollection = MGlobals.cProjects
  End Sub

  
  
 

  
 
 
End Class