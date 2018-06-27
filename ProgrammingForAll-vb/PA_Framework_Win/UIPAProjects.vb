Option Explicit On
Imports Microsoft.Win32
Imports PA_Framework_OM
Imports PA_Framework_OM.OMGlobals

Partial Public Class frmPAInstitute
  Private Sub TabPAProjects_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPAProjects.Enter
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    PAProjectBindingSource.DataSource = cPAProjects
    CurrentPAProjectChanged()
  End Sub

  Private Sub PAProjectBindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles PAProjectBindingSource.CurrentChanged
    currePrj_ = PAProjectBindingSource.Current
    CurrentPAProjectChanged()
  End Sub

  Sub CurrentPAProjectChanged()
    If currePrj_ Is Nothing Then Exit Sub
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    Dim nCount As Integer, strParDetails As String = "", strChldDetails = ""
    nCount = cPAProjects.GetEntityDependencies(currePrj_, strParDetails, strChldDetails)
    Me.txtPAProjectDependencies.Text = CStr(nCount)
    If nCount <> 0 Then
      Me.btnDeletePAProject.Enabled = False
    Else
      Me.btnDeletePAProject.Enabled = True
    End If
    Me.txtPAProjectParDeps.Text = strParDetails
    Me.txtPAProjectChldDeps.Text = strChldDetails

    'PAProject child Project Entities
    PAProjectChldProjectEntitiesBindingSource.DataSource = currePrj_.ChildEntities("ProjectEntities")
    Me.dgvPAProjectChldProjectEntity.DataSource = PAProjectChldProjectEntitiesBindingSource
    PAProjectChldProjectEntitiesBindingSource.ResetBindings(True)

    'PAProject available Project Entities
    Dim pe As New ProjectEntity
    Call omg.FillAvailableChildEntities(cPAProjects, cProjectEntities, currePrj_, pe, "ProjectEntities")
    AvPAProjectChldProjectEntitiesBindingSource.DataSource = currePrj_.AvailableChildEntities(pe)
    Me.dgvAvPAProjectChldProjectEntity.DataSource = AvPAProjectChldProjectEntitiesBindingSource
    AvPAProjectChldProjectEntitiesBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnNewPAProject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewPAProject.Click
    Dim pap = New PAProject
    pap.ProjectName = "new"

    Call omg.DBUpdate("Add", cPAProjects, pap)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    PAProjectBindingSource.ResetBindings(True)
    PAProjectBindingSource.MoveLast()

  End Sub

  Private Sub btnDeletePAProject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeletePAProject.Click
    Dim nCount As Integer, strParDetails As String = "", strChldDetails As String = ""
    currePrj_ = PAProjectBindingSource.Current
    nCount = cPAProjects.GetEntityDependencies(currePrj_, strParDetails, strChldDetails)
    If nCount = 0 Then
      Call omg.DBUpdate("Delete", cPAProjects, currePrj_)
      nPrevUpdateID = omg.GetSetting("LastUpdateID")
    Else
      MsgBox("Can not delete because of dependencies ")
    End If

    PAProjectBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnUpdatePAProject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdatePAProject.Click
    currePrj_ = PAProjectBindingSource.Current
    omg.DBUpdate("Update", cPAProjects, currePrj_)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
  End Sub

  Private Sub btnAddAvPAProjectChldProjectEntity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvPAProjectChldProjectEntity.Click
    Dim pe As ProjectEntity
    currePrj_ = PAProjectBindingSource.Current
    pe = AvPAProjectChldProjectEntitiesBindingSource.Current()
    If pe Is Nothing Then Exit Sub
    omg.AddChildEntity(cPAProjects, currePrj_, pe, "ProjectEntities")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.CurrentPAProjectChanged()
  End Sub

  Private Sub btnRemPAProjectChldProjectEntity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemPAProjectChldProjectEntity.Click
    Dim pe As ProjectEntity
    pe = PAProjectChldProjectEntitiesBindingSource.Current
    currePrj_ = PAProjectBindingSource.Current
    If pe Is Nothing Then Exit Sub
    omg.RemoveChildEntity(cPAProjects, currePrj_, pe, "ProjectEntities")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.CurrentPAProjectChanged()
  End Sub

  Private Sub btnMoveUpPAEnts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMoveUpPAEnts.Click
    Dim pe As ProjectEntity
    pe = PAProjectChldProjectEntitiesBindingSource.Current
    currePrj_.ChildEntities("ProjectEntities").MoveUp(pe)
    currePrj_ = PAProjectBindingSource.Current
    Call omg.ReOrderChildEntities(currePrj_, currePrj_.ChildEntities("ProjectEntities"), "ProjectEntities")
    PAProjectChldProjectEntitiesBindingSource.MovePrevious()
    Call Me.CurrentPAProjectChanged()
  End Sub

End Class
