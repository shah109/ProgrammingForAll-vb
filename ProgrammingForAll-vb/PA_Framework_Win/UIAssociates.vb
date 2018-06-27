Option Explicit On
Imports Microsoft.Win32
'Imports PA_Framework_Lib
Imports PA_Framework_OM
Imports PA_Framework_OM.OMGlobals

Partial Public Class frmPAInstitute
  ''Associates

  Private Sub TabAssociates_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabAssociates.Enter
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    AssociateBindingSource.DataSource = cAssociates
    CurrentAssociateChanged()
  End Sub

  Private Sub AssociateBindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AssociateBindingSource.CurrentChanged
    curreAssoc_ = AssociateBindingSource.Current
    CurrentAssociateChanged()
  End Sub

  Private Sub CurrentAssociateChanged()
    If curreAssoc_ Is Nothing Then Exit Sub
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    Dim nCount As Integer, strParDetails As String = "", strChldDetails = ""
    nCount = cPAProjects.GetEntityDependencies(curreAssoc_, strParDetails, strChldDetails)
    Me.txtAssociateDependencies.Text = CStr(nCount)
    If nCount <> 0 Then
      Me.btnDeleteAssociate.Enabled = False
    Else
      Me.btnDeleteAssociate.Enabled = True
    End If

    Me.txtAssociateParDeps.Text = strParDetails
    Me.txtAssociateChldDeps.Text = strChldDetails

    'Associate child Persons
    AssociateChldPersonBindingSource.DataSource = curreAssoc_.ChildEntities("Persons")
    Me.dgvAssocChldPersons.DataSource = AssociateChldPersonBindingSource
    AssociateChldPersonBindingSource.ResetBindings(True)

    'Associate available child Persons
    Dim prs As New Person
    Call omg.FillAvailableChildEntities(cPAProjects, cPersons, curreAssoc_, prs, "Persons")
    AvAssociateChldPersonBindingSource.DataSource = curreAssoc_.AvailableChildEntities(prs)
    Me.dgvAvAssocChldPersons.DataSource = AvAssociateChldPersonBindingSource
    AvAssociateChldPersonBindingSource.ResetBindings(True)

  End Sub

  Private Sub btnAddAvAssociateChldPerson_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvAssocChldPersons.Click
    Dim std As Person
    curreAssoc_ = AssociateBindingSource.Current()
    std = AvAssociateChldPersonBindingSource.Current()
    If std Is Nothing Then Exit Sub
    omg.AddChildEntity(cPAProjects, curreAssoc_, std, "Persons")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")

    Call Me.CurrentAssociateChanged()
  End Sub

  Private Sub btnRemAssociateChldPerson_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemAssocChldPersons.Click
    Dim std As Person
    curreAssoc_ = AssociateBindingSource.Current
    std = AssociateChldPersonBindingSource.Current
    If std Is Nothing Then Exit Sub
    omg.RemoveChildEntity(cPAProjects, curreAssoc_, std, "Persons")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.CurrentAssociateChanged()
  End Sub

  Private Sub btnAddAssociate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAssociate.Click
    Dim inst = New Associate
    inst.Comments = "new"
    Call omg.DBUpdate("Add", cAssociates, inst)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    AssociateBindingSource.ResetBindings(True)
    AssociateBindingSource.MoveLast()

  End Sub

  Private Sub btnDeleteAssociate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteAssociate.Click
    Dim nCount As Integer, strParDetails As String = "", strChldDetails As String = ""
    curreAssoc_ = AssociateBindingSource.Current
    nCount = cPAProjects.GetEntityDependencies(curreAssoc_, strParDetails, strChldDetails)
    If nCount = 0 Then
      omg.DBUpdate("Delete", cAssociates, curreAssoc_)
      nPrevUpdateID = omg.GetSetting("LastUpdateID")

    Else
      MsgBox("Can not delete because of dependencies ")
    End If
    AssociateBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnUpdateAssociate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateAssociate.Click
    curreAssoc_ = AssociateBindingSource.Current
    Call omg.DBUpdate("Update", cAssociates, curreAssoc_)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
  End Sub
End Class
