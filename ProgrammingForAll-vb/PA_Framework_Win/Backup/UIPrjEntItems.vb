Option Explicit On
Imports Microsoft.Win32
Imports PA_Framework_OM
Imports PA_Framework_OM.OMGlobals

Partial Public Class frmPAInstitute

  Private Sub TabPrjEntItems_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPrjEntItems.Enter
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    ProjectEntityItemBindingSource.DataSource = cProjectEntityItems
    ProjectEntityItemBindingSourceChanged()
  End Sub

  Private Sub ProjectEntityItemBindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ProjectEntityItemBindingSource.CurrentChanged
    currePrjEntItm_ = ProjectEntityItemBindingSource.Current
    ProjectEntityItemBindingSourceChanged()
  End Sub

  Sub ProjectEntityItemBindingSourceChanged()
    If ProjectEntityItemBindingSource.Current Is Nothing Then Exit Sub
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    Dim nCount As Integer, strParDetails As String = "", strChldDetails = ""
    nCount = cPAProjects.GetEntityDependencies(currePrjEntItm_, strParDetails, strChldDetails)
    Me.txtPrjEntItemDependencies.Text = CStr(nCount)
    If nCount <> 0 Then
      Me.btnDeletePerson.Enabled = False
    Else
      Me.btnDeletePerson.Enabled = True
    End If

    Me.txtPrjEntItemParDeps.Text = strParDetails
    Me.txtPrjEntItemChldDeps.Text = strChldDetails

    'ProjectEntityItem parent ProjectEntities
    Dim pe As New ProjectEntity
    'PersonParInstructorBindingSource.Clear()
    Call omg.FillParentEntities(cPAProjects, pe, currePrjEntItm_, "ProjectEntityItems")
    PrjEntItemParProjectEntityBindingSource.DataSource = currePrjEntItm_.ParentEntities(pe)
    Me.dgvPrjEntItemParProjectEntity.DataSource = PrjEntItemParProjectEntityBindingSource
    PrjEntItemParProjectEntityBindingSource.ResetBindings(True)

    'InstructorParentAvailableDepartments
    'AvPersonParInstructorBindingSource.Clear()
    Call omg.FillAvailableParentEntities(cPAProjects, pe, currePrjEntItm_, "ProjectEntityItems")
    AvPrjEntItemParProjectEntityBindingSource.DataSource = currePrjEntItm_.AvailableParentEntities(pe)
    Me.dgvAvPrjEntItemParProjectEntity.DataSource = AvPrjEntItemParProjectEntityBindingSource
    AvPrjEntItemParProjectEntityBindingSource.ResetBindings(True)

  End Sub

  Private Sub btnRemPrjEntItemParInstructor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemPrjEntItemParProjectEntity.Click
    Dim pe As ProjectEntity
    Dim bAddEntity As Integer
    currePrjEntItm_ = ProjectEntityItemBindingSource.Current
    pe = PrjEntItemParProjectEntityBindingSource.Current
    If pe Is Nothing Then Exit Sub

    bAddEntity = omg.RemoveChildEntity(cPAProjects, pe, currePrjEntItm_, "ProjectEntityItems")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bAddEntity <> 1 Then
      MsgBox("Could not remove the entity from db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.ProjectEntityItemBindingSourceChanged()
  End Sub

  Private Sub btnAddAvPrjEntItemParInstructor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvPrjEntItemParProjectEntity.Click
    Dim pe As ProjectEntity
    Dim bAddEntity As Integer
    currePrjEntItm_ = ProjectEntityItemBindingSource.Current
    pe = AvPrjEntItemParProjectEntityBindingSource.Current
    If pe Is Nothing Then Exit Sub

    bAddEntity = omg.AddChildEntity(cPAProjects, pe, currePrjEntItm_, "ProjectEntityItems")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bAddEntity <> 1 Then
      MsgBox("Could not add the entity to db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.ProjectEntityItemBindingSourceChanged()
  End Sub

  Private Sub btnAddPrjEntItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddPrjEntItem.Click
    Dim pei = New ProjectEntityItem
    pei.PropertyName = "new"
    'ProjectEntityItemBindingSource.Add(pei)
    Call omg.DBUpdate("Add", cProjectEntityItems, pei)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    ProjectEntityItemBindingSource.ResetBindings(True)
    ProjectEntityItemBindingSource.MoveLast()

  End Sub

  Private Sub btnDeletePrjEntItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelPrjEntItem.Click
    Dim nCount As Integer, strParDetails As String = "", strChldDetails As String = ""
    currePrjEntItm_ = ProjectEntityItemBindingSource.Current
    nCount = cPAProjects.GetEntityDependencies(currePrjEntItm_, strParDetails, strChldDetails)
    If nCount = 0 Then
      Call omg.DBUpdate("Delete", cProjectEntityItems, currePrjEntItm_)
      nPrevUpdateID = omg.GetSetting("LastUpdateID")
    Else
      MsgBox("Can not delete because of dependencies ")
    End If
    ProjectEntityItemBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnUpdatePrjEntItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdatePrjEntItem.Click
    currePrjEntItm_ = ProjectEntityItemBindingSource.Current
    Call omg.DBUpdate("Update", cProjectEntityItems, currePrjEntItm_)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
  End Sub
End Class


