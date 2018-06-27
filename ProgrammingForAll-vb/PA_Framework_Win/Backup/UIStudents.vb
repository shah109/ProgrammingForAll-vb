Option Explicit On
Imports Microsoft.Win32
Imports PA_Framework_OM
Imports PA_Framework_OM.OMGlobals

Partial Public Class frmPAInstitute
  Private Sub TabStudents_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabStudents.Enter
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    StudentBindingSource.DataSource = cStudents
    StudentBindingSourceChanged()
  End Sub

  Private Sub StudentBindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles StudentBindingSource.CurrentChanged
    curreStdt_ = StudentBindingSource.Current
    StudentBindingSourceChanged()
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)

  End Sub

  Sub StudentBindingSourceChanged()
    If curreStdt_ Is Nothing Then Exit Sub
    Dim nCount As Integer, strParDetails As String = "", strChldDetails = ""
    nCount = cPAProjects.GetEntityDependencies(curreStdt_, strParDetails, strChldDetails)
    Me.txtStudentDependencies.Text = CStr(nCount)
    If nCount <> 0 Then
      Me.btnDeleteStudent.Enabled = False
    Else
      Me.btnDeleteStudent.Enabled = True
    End If
    Me.txtStudentParDeps.Text = strParDetails
    Me.txtStudentChldDeps.Text = strChldDetails

    'student parent courses
    Dim crs As New Course
    Call omg.FillParentEntities(cPAProjects, crs, curreStdt_, "Students")
    StudentParCourseBindingSource.DataSource = curreStdt_.ParentEntities(crs)
    Me.dgvStudentParCourses.DataSource = StudentParCourseBindingSource
    StudentParCourseBindingSource.ResetBindings(True)

    'CourseParentAvailableDepartments
    Call omg.FillAvailableParentEntities(cPAProjects, crs, curreStdt_, "Students")
    AvStudentParCourseBindingSource.DataSource = curreStdt_.AvailableParentEntities(crs)
    Me.dgvAvStudentParCourses.DataSource = AvStudentParCourseBindingSource
    AvStudentParCourseBindingSource.ResetBindings(True)

    'student parent calendar
    Dim cal As New Calendar
    Call omg.FillParentEntities(cPAProjects, cal, curreStdt_, "Students")
    StudentParCalendarBindingSource.DataSource = curreStdt_.ParentEntities(cal)
    Me.dgvStudentParCalendars.DataSource = StudentParCalendarBindingSource
    StudentParCalendarBindingSource.ResetBindings(True)

    'CourseParent available Calendar
    Call omg.FillAvailableParentEntities(cPAProjects, cal, curreStdt_, "Students")
    AvStudentParCalendarBindingSource.DataSource = curreStdt_.AvailableParentEntities(cal)
    Me.dgvAvStudentParCalendars.DataSource = AvStudentParCalendarBindingSource
    AvStudentParCalendarBindingSource.ResetBindings(True)


  End Sub

  Private Sub btnRemStudentParCourse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemStudentParCourse.Click
    Dim crs As Course
    Dim bAddEntity As Integer
    curreStdt_ = StudentBindingSource.Current
    crs = StudentParCourseBindingSource.Current
    bAddEntity = omg.RemoveChildEntity(cPAProjects, crs, curreStdt_, "Students")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bAddEntity = 0 Then
      MsgBox("Could not remove the entity from db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.StudentBindingSourceChanged()
  End Sub

  Private Sub btnAddAvStudentParCourse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvStudentParCourse.Click
    Dim crs As Course
    Dim bAddEntity As Integer
    curreStdt_ = StudentBindingSource.Current
    crs = AvStudentParCourseBindingSource.Current

    bAddEntity = omg.AddChildEntity(cPAProjects, crs, curreStdt_, "Students")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bAddEntity = 0 Then
      MsgBox("Could not add the entity to db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.StudentBindingSourceChanged()
  End Sub

  Private Sub btnAddStudent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddStudent.Click
    Dim std = New Student
    std.EntityItem_1 = "new"
    'StudentBindingSource.Add(std)
    Call omg.DBUpdate("Add", cStudents, std)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    StudentBindingSource.ResetBindings(True)
    StudentBindingSource.MoveLast()
  End Sub

  Private Sub btnDeleteStudent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteStudent.Click
    Dim nCount As Integer, strParDetails As String = "", strChldDetails As String = ""
    curreStdt_ = StudentBindingSource.Current
    'StudentBindingSource.s()

    nCount = cPAProjects.GetEntityDependencies(curreStdt_, strParDetails, strChldDetails)
    If nCount = 0 Then
      Call omg.DBUpdate("Delete", cStudents, curreStdt_)
      nPrevUpdateID = omg.GetSetting("LastUpdateID")
    Else
      MsgBox("Can not delete because of dependencies ")
    End If
    StudentBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnUpdateStudent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateStudent.Click
    curreStdt_ = StudentBindingSource.Current
    omg.DBUpdate("Update", cStudents, curreStdt_)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    StudentBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnRemStudentParCalendar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemStudentParCalendar.Click
    Dim cal As Calendar
    Dim bAddEntity As Integer
    curreStdt_ = StudentBindingSource.Current
    cal = StudentParCalendarBindingSource.Current
    If cal Is Nothing Then Exit Sub
    bAddEntity = omg.RemoveChildEntity(cPAProjects, cal, curreStdt_, "Students")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bAddEntity = 0 Then
      MsgBox("Could not remove the entity from db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.StudentBindingSourceChanged()
  End Sub

  Private Sub btnAddAvStudentParCalendar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvStudentParCalendar.Click
    Dim cal As Calendar
    Dim bAddEntity As Integer
    curreStdt_ = StudentBindingSource.Current
    cal = AvStudentParCalendarBindingSource.Current
    If cal Is Nothing Then Exit Sub
    bAddEntity = omg.AddChildEntity(cPAProjects, cal, curreStdt_, "Students")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.StudentBindingSourceChanged()
  End Sub


End Class
