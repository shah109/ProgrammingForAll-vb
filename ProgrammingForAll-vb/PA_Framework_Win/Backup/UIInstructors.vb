Option Explicit On
Imports Microsoft.Win32
Imports PA_Framework_OM
Imports PA_Framework_OM.OMGlobals

Partial Public Class frmPAInstitute
  ''Instructors

  Private Sub TabInstructors_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabInstructors.Enter
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    InstructorBindingSource.DataSource = cInstructors
    CurrentInstructorChanged()
  End Sub

  Private Sub InstructorBindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles InstructorBindingSource.CurrentChanged
    curreInst_ = InstructorBindingSource.Current
    CurrentInstructorChanged()
  End Sub

  Private Sub CurrentInstructorChanged()
    If curreInst_ Is Nothing Then Exit Sub
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    Dim nCount As Integer, strParDetails As String = "", strChldDetails = ""
    nCount = cPAProjects.GetEntityDependencies(curreInst_, strParDetails, strChldDetails)
    Me.txtInstructorDependencies.Text = CStr(nCount)
    If nCount <> 0 Then
      Me.btnDeleteInstructor.Enabled = False
    Else
      Me.btnDeleteInstructor.Enabled = True
    End If

    Me.txtInstructorParDeps.Text = strParDetails
    Me.txtInstructorChldDeps.Text = strChldDetails

    'If Not curreInst_.mChildPersons.ItemByLO(1) Is Nothing Then
    '  Me.txtInstructorFirstName.Text = curreInst_.mChildPersons.ItemByLO(1).Fullname

    '  Me.cboIstructorChildPersons.SelectedText = curreInst_.mChildPersons.ItemByLO(1).Fullname
    'Else
    '  Me.txtInstructorFirstName.Text = ""
    'End If



    'Instructor child Persons
    InstructorChldPersonBindingSource.DataSource = curreInst_.ChildEntities("Persons")
    Me.dgvInstructorsChldPersons.DataSource = InstructorChldPersonBindingSource
    InstructorChldPersonBindingSource.ResetBindings(True)



    'Instructor available child Persons
    Dim prs As New Person
    Call omg.FillAvailableChildEntities(cPAProjects, cPersons, curreInst_, prs, "Persons")
    AvInstructorChldPersonBindingSource.DataSource = curreInst_.AvailableChildEntities(prs)
    Me.dgvAvInstructorChldPersons.DataSource = AvInstructorChldPersonBindingSource
    AvInstructorChldPersonBindingSource.ResetBindings(True)

    'Me.cboIstructorChildPersons.DataSource = AvInstructorChldPersonBindingSource

    'Instructor parent Courses
    Dim crs As New Course
    'InstructorParCourseBindingSource.Clear()
    Call omg.FillParentEntities(cPAProjects, crs, curreInst_, "Instructors")
    InstructorParCourseBindingSource.DataSource = curreInst_.ParentEntities(crs)
    Me.dgvInstructorParCourses.DataSource = InstructorParCourseBindingSource
    InstructorParCourseBindingSource.ResetBindings(True)

    'Available Instructor parent Courses
    'AvInstructorParCourseBindingSource.Clear()
    Call omg.FillAvailableParentEntities(cPAProjects, crs, curreInst_, "Instructors")
    AvInstructorParCourseBindingSource.DataSource = curreInst_.AvailableParentEntities(crs)
    Me.dgvAvInstructorParCourses.DataSource = AvInstructorParCourseBindingSource
    AvInstructorParCourseBindingSource.ResetBindings(True)

  End Sub

  Private Sub btnAddAvInstructorChldPerson_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvInstructorChldPerson.Click
    Dim std As Person
    curreInst_ = InstructorBindingSource.Current()
    std = AvInstructorChldPersonBindingSource.Current()
    If std Is Nothing Then Exit Sub
    omg.AddChildEntity(cPAProjects, curreInst_, std, "Persons")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.CurrentInstructorChanged()
  End Sub

  Private Sub btnRemInstructorChldPerson_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemInstructorChldPerson.Click
    Dim std As Person
    curreInst_ = InstructorBindingSource.Current
    std = InstructorChldPersonBindingSource.Current
    If std Is Nothing Then Exit Sub
    omg.RemoveChildEntity(cPAProjects, curreInst_, std, "Persons")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.CurrentInstructorChanged()
  End Sub

  Private Sub btnRemInstructorParCourses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemInstructorParCourses.Click
    Dim crs As Course
    Dim bAddEntity As Integer
    curreInst_ = InstructorBindingSource.Current
    crs = PersonParInstructorBindingSource.Current
    If crs Is Nothing Then Exit Sub
    bAddEntity = omg.RemoveChildEntity(cPAProjects, crs, curreInst_, "Instructors")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bAddEntity <> 1 Then
      MsgBox("Could not remove the entity from db")
      Exit Sub
    End If
    'refresh dgv's
    Call CurrentInstructorChanged()
  End Sub

  Private Sub btnAddAvInstructorParCourses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvInstructorParCourses.Click
    Dim crs As Course
    Dim bAddEntity As Integer
    curreInst_ = InstructorBindingSource.Current()
    crs = AvInstructorParCourseBindingSource.Current
    If crs Is Nothing Then Exit Sub
    bAddEntity = omg.AddChildEntity(cPAProjects, crs, curreInst_, "Instructors")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bAddEntity <> 1 Then
      MsgBox("Could not add the entityto db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.CurrentInstructorChanged()
  End Sub

  Private Sub btnAddInstructor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddInstructor.Click
    Dim inst = New Instructor
    inst.Comments = "new"

    Call omg.DBUpdate("Add", cInstructors, inst)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    InstructorBindingSource.ResetBindings(True)
    InstructorBindingSource.MoveLast()

  End Sub

  Private Sub btnDeleteInstructor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteInstructor.Click
    Dim nCount As Integer, strParDetails As String = "", strChldDetails As String = ""
    curreInst_ = InstructorBindingSource.Current
    nCount = cPAProjects.GetEntityDependencies(curreInst_, strParDetails, strChldDetails)
    If nCount = 0 Then
      Call omg.DBUpdate("Delete", cInstructors, curreInst_)
      nPrevUpdateID = omg.GetSetting("LastUpdateID")
    Else
      MsgBox("Can not delete because of dependencies ")
    End If
    InstructorBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnUpdateInstructor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateInstructor.Click
    curreInst_ = InstructorBindingSource.Current
    Call omg.DBUpdate("Update", cInstructors, curreInst_)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
  End Sub

End Class
