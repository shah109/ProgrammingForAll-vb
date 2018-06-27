Option Explicit On
Imports Microsoft.Win32
Imports PA_Framework_OM
Imports PA_Framework_OM.OMGlobals

Partial Public Class frmPAInstitute
  ''Project Entities

  Private Sub TabProjectEntities_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPrjEnts.Enter
    Me.txtProgress.Text = AppSettings.GetSetting("Comments")
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    ProjectEntityBindingSource.DataSource = cProjectEntities
    cboPAProjects.DataSource = cPAProjects
    CurrentProjectEntityChanged()
  End Sub

  Private Sub ProjectEntityBindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ProjectEntityBindingSource.CurrentChanged
    currePrjEnt_ = ProjectEntityBindingSource.Current
    CurrentProjectEntityChanged()
  End Sub

  Private Sub CurrentProjectEntityChanged()
    If currePrjEnt_ Is Nothing Then Exit Sub
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    Dim nCount As Integer, strParDetails As String = "", strChldDetails = ""
    nCount = cPAProjects.GetEntityDependencies(currePrjEnt_, strParDetails, strChldDetails)
    Me.txtProjectEntityDependencies.Text = CStr(nCount)
    If nCount <> 0 Then
      Me.btnDeleteProjectEntity.Enabled = False
    Else
      Me.btnDeleteProjectEntity.Enabled = True
    End If

    Me.txtProjectEntityParDeps.Text = strParDetails
    Me.txtProjectEntityChldDeps.Text = strChldDetails

    'ProjectEntity child Project Entity Items
    ProjectEntityChldPrjEntItemBindingSource.DataSource = currePrjEnt_.ChildEntities("ProjectEntityItems")
    Me.dgvPrjEntsChldEntItems.DataSource = ProjectEntityChldPrjEntItemBindingSource
    ProjectEntityChldPrjEntItemBindingSource.ResetBindings(True)

    'ProjectEntity available child Entity items
    Dim pei As New ProjectEntityItem
    Call omg.FillAvailableChildEntities(cPAProjects, cProjectEntityItems, currePrjEnt_, pei, "ProjectEntityItems")
    AvProjectEntityChldPrjEntItemBindingSource.DataSource = currePrjEnt_.AvailableChildEntities(pei)
    Me.dgvAvPrjEntsChldEntItems.DataSource = AvProjectEntityChldPrjEntItemBindingSource
    AvProjectEntityChldPrjEntItemBindingSource.ResetBindings(True)

  End Sub

  Private Sub btnAddAvProjectEntityChldPerson_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvPrjEntChldPrjEntItem.Click
    Dim pei As ProjectEntityItem
    currePrjEnt_ = ProjectEntityBindingSource.Current()
    pei = AvProjectEntityChldPrjEntItemBindingSource.Current()
    If pei Is Nothing Then Exit Sub
    omg.AddChildEntity(cPAProjects, currePrjEnt_, pei, "ProjectEntityItems")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.CurrentProjectEntityChanged()
  End Sub

  Private Sub btnRemProjectEntityChldPrjEntItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemPrjEntChldPrjEntItem.Click
    Dim pei As ProjectEntityItem
    currePrjEnt_ = ProjectEntityBindingSource.Current
    pei = ProjectEntityChldPrjEntItemBindingSource.Current
    If pei Is Nothing Then Exit Sub
    omg.RemoveChildEntity(cPAProjects, currePrjEnt_, pei, "ProjectEntityItems")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.CurrentProjectEntityChanged()
  End Sub

  Private Sub btnAddProjectEntity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddProjectEntity.Click
    Dim pe = New ProjectEntity
    pe.EntityName = "newEntity"
    'ProjectEntityBindingSource.Add(pe)

    Call omg.DBUpdate("Add", cProjectEntities, pe)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    ProjectEntityBindingSource.ResetBindings(True)
    ProjectEntityBindingSource.MoveLast()
  End Sub

  Private Sub btnDeleteProjectEntity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteProjectEntity.Click
    Dim nCount As Integer, strParDetails As String = "", strChldDetails As String = ""
    currePrjEnt_ = ProjectEntityBindingSource.Current
    nCount = cPAProjects.GetEntityDependencies(currePrjEnt_, strParDetails, strChldDetails)
    If nCount = 0 Then
      Call omg.DBUpdate("Delete", cProjectEntities, currePrjEnt_)
      nPrevUpdateID = omg.GetSetting("LastUpdateID")
    Else
      MsgBox("Can not delete because of dependencies ")
    End If
    ProjectEntityBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnUpdateProjectEntity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateProjectEntity.Click
    currePrjEnt_ = ProjectEntityBindingSource.Current
    Call omg.DBUpdate("Update", cProjectEntities, currePrjEnt_)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
  End Sub

  Private Sub btnUpdateAvPrjEntChldEntItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateAvPrjEntChldEntItems.Click
    Dim pei As ProjectEntityItem
    pei = AvProjectEntityChldPrjEntItemBindingSource.Current
    Call omg.DBUpdate("Update", cProjectEntityItems, pei)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
  End Sub

  Private Sub btnUpdatePrjEntChldEntItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdatePrjEntChldEntItems.Click
    Dim pei As ProjectEntityItem
    pei = ProjectEntityChldPrjEntItemBindingSource.Current
    Call omg.DBUpdate("Update", cProjectEntityItems, pei)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
  End Sub

  Private Sub btnMoveUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMoveUp.Click
    Dim pei As ProjectEntityItem
    pei = ProjectEntityChldPrjEntItemBindingSource.Current
    currePrjEnt_.ChildEntities("ProjectEntityItems").MoveUp(pei)
    currePrjEnt_ = ProjectEntityBindingSource.Current
    Call omg.ReOrderChildEntities(currePrjEnt_, currePrjEnt_.ChildEntities("ProjectEntityItems"), "ProjectEntityItems")
    'Call DBUpdate("Update", cProjectEntities, currePrjEnt_)
    ProjectEntityChldPrjEntItemBindingSource.MovePrevious()
    Call Me.CurrentProjectEntityChanged()
  End Sub

  Private Sub cboPAProjects_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPAProjects.SelectedIndexChanged
    Dim strSelection As String
    Dim pr As PAProject
    strSelection = cboPAProjects.Text
    pr = cPAProjects.GetProjectByName(strSelection)
    ProjectEntityBindingSource.DataSource = pr.ChildprojectEntities
    'MsgBox(cboPAProjects.SelectedIndex.ToString)
  End Sub

  Private Sub btnGenerateCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerateCode.Click
    'Code for generating code
    '1. Create Entitys class collection
    PACode_EntityCollectionClass.GenerateEntityCollectionClass(cboPAProjects.Text)
    '2. Create Entity base class
    PACode_EntityBaseClass.GenerateEntityBaseClass(cboPAProjects.Text)
    '3. Create Entity class 
    PACode_EntityClass.GenerateEntityClass(cboPAProjects.Text)
    '4 Create Global custom code
    Dim sPAGlobalCustom As New PACode_Global
    '5 Create Metadata custom code.
    PACode_Metadata.GenerateMetaDataCustom(cboPAProjects.Text)

    Call PACode_ExcelUI.GenerateExcelUI(cboPAProjects.Text)
  End Sub

  Private Sub TabPrjEnts_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPrjEnts.Leave
    Call AppSettings.SetSetting("Comments", Me.txtProgress.Text)
  End Sub
End Class
