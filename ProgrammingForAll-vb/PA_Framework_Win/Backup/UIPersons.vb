Option Explicit On
Imports Microsoft.Win32
Imports PA_Framework_OM
Imports PA_Framework_OM.OMGlobals
Partial Public Class frmPAInstitute

  Private Sub TabPersons_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPersons.Enter
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    PersonBindingSource.DataSource = cPersons
    PersonBindingSourceChanged()
  End Sub

  Private Sub PersonBindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles PersonBindingSource.CurrentChanged
    currePrsn_ = PersonBindingSource.Current
    PersonBindingSourceChanged()
  End Sub

  Sub PersonBindingSourceChanged()
    If PersonBindingSource.Current Is Nothing Then Exit Sub
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    Dim nCount As Integer, strParDetails As String = "", strChldDetails = ""
    nCount = cPAProjects.GetEntityDependencies(currePrsn_, strParDetails, strChldDetails)
    Me.txtPersonDependencies.Text = CStr(nCount)
    If nCount <> 0 Then
      Me.btnDeletePerson.Enabled = False
    Else
      Me.btnDeletePerson.Enabled = True
    End If

    Me.txtPersonParDeps.Text = strParDetails
    Me.txtPersonChldDeps.Text = strChldDetails

    'Person parent Instructors
    Dim crs As New Instructor
    'PersonParInstructorBindingSource.Clear()
    Call omg.FillParentEntities(cPAProjects, crs, currePrsn_, "Persons")
    PersonParInstructorBindingSource.DataSource = currePrsn_.ParentEntities(crs)
    Me.dgvPersonParInstructor.DataSource = PersonParInstructorBindingSource
    PersonParInstructorBindingSource.ResetBindings(True)

    'InstructorParentAvailableDepartments
    'AvPersonParInstructorBindingSource.Clear()
    Call omg.FillAvailableParentEntities(cPAProjects, crs, currePrsn_, "Persons")
    AvPersonParInstructorBindingSource.DataSource = currePrsn_.AvailableParentEntities(crs)
    Me.dgvAvPersonParInstructor.DataSource = AvPersonParInstructorBindingSource
    AvPersonParInstructorBindingSource.ResetBindings(True)

    'Person parent Students
    Dim std As New Student
    'PersonParInstructorBindingSource.Clear()
    Call omg.FillParentEntities(cPAProjects, std, currePrsn_, "Persons")
    PersonParStudentBindingSource.DataSource = currePrsn_.ParentEntities(std)
    Me.dgvPersonParStudent.DataSource = PersonParStudentBindingSource
    PersonParStudentBindingSource.ResetBindings(True)

    'Person Available parent Students
    Call omg.FillAvailableParentEntities(cPAProjects, std, currePrsn_, "Persons")
    AvPersonParStudentBindingSource.DataSource = currePrsn_.AvailableParentEntities(std)
    Me.dgvAvPersonParStudent.DataSource = AvPersonParStudentBindingSource
    AvPersonParStudentBindingSource.ResetBindings(True)

    'Person parent Associates
    Dim assoc As New Associate

    Call omg.FillParentEntities(cPAProjects, assoc, currePrsn_, "Persons")
    PersonParAssociateBindingSource.DataSource = currePrsn_.ParentEntities(assoc)
    Me.dgvPersonParAssociate.DataSource = PersonParAssociateBindingSource
    PersonParAssociateBindingSource.ResetBindings(True)

    'Person Available parent Associates
    Call omg.FillAvailableParentEntities(cPAProjects, assoc, currePrsn_, "Persons")
    AvPersonParAssociateBindingSource.DataSource = currePrsn_.AvailableParentEntities(assoc)
    Me.dgvAvPersonParAssociate.DataSource = AvPersonParAssociateBindingSource
    AvPersonParAssociateBindingSource.ResetBindings(True)

  End Sub

  Private Sub btnRemPersonParAssociate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemPersonParAssociate.Click
    Dim assoc As Associate
    Dim bAddEntity As Integer
    currePrsn_ = PersonBindingSource.Current
    assoc = PersonParAssociateBindingSource.Current
    If assoc Is Nothing Then Exit Sub

    bAddEntity = omg.RemoveChildEntity(cPAProjects, assoc, currePrsn_, "Persons")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bAddEntity <> 1 Then
      MsgBox("Could not remove the entity from db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.PersonBindingSourceChanged()
  End Sub

  Private Sub btnAddAvPersonParAssociate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvPersonParAssociate.Click
    Dim assoc As Associate
    Dim bAddEntity As Integer
    currePrsn_ = PersonBindingSource.Current
    assoc = AvPersonParAssociateBindingSource.Current
    If assoc Is Nothing Then Exit Sub

    bAddEntity = omg.AddChildEntity(cPAProjects, assoc, currePrsn_, "Persons")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bAddEntity <> 1 Then
      MsgBox("Could not add the entity to db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.PersonBindingSourceChanged()
  End Sub

  Private Sub btnRemPersonParStudent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemPersonParStudent.Click
    Dim std As Student
    Dim bAddEntity As Integer
    currePrsn_ = PersonBindingSource.Current
    std = PersonParStudentBindingSource.Current
    If std Is Nothing Then Exit Sub

    bAddEntity = omg.RemoveChildEntity(cPAProjects, std, currePrsn_, "Persons")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bAddEntity <> 1 Then
      MsgBox("Could not remove the entity from db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.PersonBindingSourceChanged()
  End Sub

  Private Sub btnAddAvPersonParStudent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvPersonParStudent.Click
    Dim std As Student
    Dim bAddEntity As Integer
    currePrsn_ = PersonBindingSource.Current
    std = AvPersonParStudentBindingSource.Current
    If std Is Nothing Then Exit Sub

    bAddEntity = omg.AddChildEntity(cPAProjects, std, currePrsn_, "Persons")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bAddEntity <> 1 Then
      MsgBox("Could not add the entity to db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.PersonBindingSourceChanged()
  End Sub

  Private Sub btnRemPersonParInstructor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemPersonParInstructor.Click
    Dim crs As Instructor
    Dim bAddEntity As Integer
    currePrsn_ = PersonBindingSource.Current
    crs = PersonParInstructorBindingSource.Current
    If crs Is Nothing Then Exit Sub

    bAddEntity = omg.RemoveChildEntity(cPAProjects, crs, currePrsn_, "Persons")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bAddEntity <> 1 Then
      MsgBox("Could not remove the entity from db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.PersonBindingSourceChanged()
  End Sub

  Private Sub btnAddAvPersonParInstructor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvPersonParInstructor.Click
    Dim crs As Instructor
    Dim bAddEntity As Integer
    currePrsn_ = PersonBindingSource.Current
    crs = AvPersonParInstructorBindingSource.Current
    If crs Is Nothing Then Exit Sub

    bAddEntity = omg.AddChildEntity(cPAProjects, crs, currePrsn_, "Persons")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bAddEntity <> 1 Then
      MsgBox("Could not add the entity to db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.PersonBindingSourceChanged()
  End Sub

  Private Sub btnAddPerson_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddPerson.Click
    Dim prs = New Person
    prs.FirstName = "new"

    Call omg.DBUpdate("Add", cPersons, prs)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    PersonBindingSource.ResetBindings(True)
    PersonBindingSource.MoveLast()

  End Sub

  Private Sub btnDeletePerson_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeletePerson.Click
    Dim nCount As Integer, strParDetails As String = "", strChldDetails As String = ""
    currePrsn_ = PersonBindingSource.Current
    nCount = cPAProjects.GetEntityDependencies(currePrsn_, strParDetails, strChldDetails)
    If nCount = 0 Then
      Call omg.DBUpdate("Delete", cPersons, currePrsn_)
      nPrevUpdateID = omg.GetSetting("LastUpdateID")
    Else
      MsgBox("Can not delete because of dependencies ")
    End If
    PersonBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnUpdatePerson_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdatePerson.Click
    currePrsn_ = PersonBindingSource.Current
    Call omg.DBUpdate("Update", cPersons, currePrsn_)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
  End Sub

End Class


