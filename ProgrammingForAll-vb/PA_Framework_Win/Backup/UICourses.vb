Option Explicit On
Imports Microsoft.Win32
Imports PA_Framework_OM
Imports PA_Framework_OM.OMGlobals
Partial Public Class frmPAInstitute

  Private Sub TabCourses_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabCourses.Enter
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    CourseBindingSource.DataSource = omg.GetCollection(cCourses)
    CourseSelectionChanged()
    CourseBindingSource.ResetBindings(True)
  End Sub

  Private Sub CourseBindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CourseBindingSource.CurrentChanged
    curreCrse_ = CourseBindingSource.Current
    CourseSelectionChanged()
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)

  End Sub

  Private Sub CourseSelectionChanged()
    If curreCrse_ Is Nothing Then Exit Sub
    'Refresh all children and parents
    Dim nCount As Integer, strParDetails As String = "", strChldDetails = ""
    nCount = cPAProjects.GetEntityDependencies(curreCrse_, strParDetails, strChldDetails)
    Me.txtCourseDependencies.Text = CStr(nCount)
    If nCount <> 0 Then
      Me.btnDeleteCourse.Enabled = False
    Else
      Me.btnDeleteCourse.Enabled = True
    End If
    Me.TxtCourseParDeps.Text = strParDetails
    Me.txtCourseChldDeps.Text = strChldDetails

    'Course child Students
    CoursechldStudentBindingSource.DataSource = curreCrse_.ChildEntities("Students")
    Me.dgvCourseChldStudents.DataSource = CoursechldStudentBindingSource
    CoursechldStudentBindingSource.ResetBindings(True)

    'Course available child Students
    Dim std As New Student
    Call omg.FillAvailableChildEntities(cPAProjects, cStudents, curreCrse_, std, "Students")
    AvCourseChldStudentBindingSource.DataSource = curreCrse_.AvailableChildEntities(std)
    Me.dgvAvCourseChldStudents.DataSource = AvCourseChldStudentBindingSource
    AvCourseChldStudentBindingSource.ResetBindings(True)

    'Course child department
    Dim dpt As New Department
    CourseChldDepartmentBindingSource.DataSource = curreCrse_.ChildEntities("Departments")
    Me.dgvCourseChldDepartments.DataSource = CourseChldDepartmentBindingSource
    CourseChldDepartmentBindingSource.ResetBindings(True)

    'Course child Available Departments
    Call omg.FillAvailableChildEntities(cPAProjects, cDepartments, curreCrse_, dpt, "Departments")
    AvCourseChldDepartmentBindingSource.DataSource = curreCrse_.AvailableChildEntities(dpt)
    Me.dgvAvCourseChldDepartment.DataSource = AvCourseChldDepartmentBindingSource
    AvCourseChldDepartmentBindingSource.ResetBindings(True)

    'Course parent Calendars
    Dim cal As New Calendar
    Call omg.FillParentEntities(cPAProjects, cal, curreCrse_, "Courses")
    CourseParCalendarBindingSourse.DataSource = curreCrse_.ParentEntities(cal)
    Me.dgvCourseParCalendars.DataSource = CourseParCalendarBindingSourse
    CourseParCalendarBindingSourse.ResetBindings(True)

    'Course Parent Available Calendars
    Call omg.FillAvailableParentEntities(cPAProjects, cal, curreCrse_, "Courses")
    AvCourseParCalendarBindingSource.DataSource = curreCrse_.AvailableParentEntities(cal)
    Me.dgvAvCourseParCalendars.DataSource = AvCourseParCalendarBindingSource
    AvCourseParCalendarBindingSource.ResetBindings(True)

    'Course child Instructors
    CourseChldInstructorBindingSource.DataSource = curreCrse_.ChildEntities("Instructors")
    Me.dgvCourseChldInstructors.DataSource = CourseChldInstructorBindingSource
    CourseChldInstructorBindingSource.ResetBindings(True)

    'Course available child Instructors
    Dim inst As New Instructor
    Call omg.FillAvailableChildEntities(cPAProjects, cInstructors, curreCrse_, inst, "Instructors")
    AvCourseChldInstructorBindingSource.DataSource = curreCrse_.AvailableChildEntities(inst)
    Me.dgvAvCourseChldInstructors.DataSource = AvCourseChldInstructorBindingSource
    AvCourseChldInstructorBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnAddCourse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddCourse.Click
    Dim np = New Course
    np.Name = "new"

    nUpdate = omg.DBUpdate("Add", cCourses, np)
    If nUpdate <> 0 Then nPrevUpdateID = omg.GetSetting("LastUpdateID")
    CourseBindingSource.ResetBindings(True)
    CourseBindingSource.MoveLast()
    nPrevUpdateID = AppSettings.LastUpdateID
  End Sub

  Private Sub btnDeleteCourse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteCourse.Click
    Dim nCount As Integer, strParDetails As String = "", strChldDetails As String = ""
    curreCrse_ = CourseBindingSource.Current
    nCount = cPAProjects.GetEntityDependencies(curreCrse_, strParDetails, strChldDetails)
    If nCount = 0 Then
      nUpdate = omg.DBUpdate("Delete", cCourses, curreCrse_)
      If nUpdate <> 0 Then nPrevUpdateID = omg.GetSetting("LastUpdateID")
    Else
      MsgBox("Can not delete because of dependencies ")
    End If
    CourseBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnUpateCourse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpateCourse.Click
    curreCrse_ = CourseBindingSource.Current
    nUpdate = omg.DBUpdate("Update", cCourses, curreCrse_)
    If nUpdate <> 0 Then nPrevUpdateID = omg.GetSetting("LastUpdateID")
    CourseBindingSource.ResetCurrentItem()
  End Sub

  Private Sub btnAddAvCourseChldStudent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvCourseChldStudent.Click
    Dim std As Student
    curreCrse_ = CourseBindingSource.Current()
    std = AvCourseChldStudentBindingSource.Current()
    If std Is Nothing Then Exit Sub
    nUpdate = omg.AddChildEntity(cPAProjects, curreCrse_, std, "Students")
    If nUpdate <> 0 Then nPrevUpdateID = omg.GetSetting("LastUpdateID") 'refresh dgv's
    Call Me.CourseSelectionChanged()

  End Sub

  Private Sub btnRemCourseChldStudent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemCourseChldStudent.Click
    Dim std As Student
    curreCrse_ = CourseBindingSource.Current
    std = CoursechldStudentBindingSource.Current
    If std Is Nothing Then Exit Sub
    nUpdate = omg.RemoveChildEntity(cPAProjects, curreCrse_, std, "Students")
    If nUpdate <> 0 Then nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.CourseSelectionChanged()
    'nPrevUpdateID = omg.GetSetting("LastUpdateID")

  End Sub

  Private Sub btnAddCourseChldDepartment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddCourseChldDepartment.Click
    Dim dpt As Department
    Dim bAddEntity As Integer
    curreCrse_ = CourseBindingSource.Current()
    dpt = AvCourseChldDepartmentBindingSource.Current
    bAddEntity = omg.AddChildEntity(cPAProjects, curreCrse_, dpt, "Departments")
    If bAddEntity <> 1 Then
      MsgBox("Could not add the entityto db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.CourseSelectionChanged()
    nPrevUpdateID = omg.GetSetting("LastUpdateID")

  End Sub

  Private Sub btnRemCourseChldDepartment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemCourseChldDepartment.Click
    Dim dpt As Department
    Dim bAddEntity As Integer
    curreCrse_ = CourseBindingSource.Current
    dpt = CourseChldDepartmentBindingSource.Current
    bAddEntity = omg.RemoveChildEntity(cPAProjects, curreCrse_, dpt, "Departments")
    If bAddEntity <> 1 Then
      MsgBox("Could not remove the entity from db")
      Exit Sub
    End If
    nPrevUpdateID = omg.GetSetting("LastUpdateID")

    'refresh dgv's
    Call Me.CourseSelectionChanged()
    nPrevUpdateID = omg.GetSetting("LastUpdateID")

  End Sub

  Private Sub btnAddAvCourseParCalendar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvCourseParCalendar.Click
    Dim cal As Calendar
    Dim bAddEntity As Integer
    curreCrse_ = CourseBindingSource.Current()
    cal = AvCourseParCalendarBindingSource.Current
    If cal Is Nothing Then Exit Sub
    bAddEntity = omg.AddChildEntity(cPAProjects, cal, curreCrse_, "Courses")
    If bAddEntity <> 1 Then
      MsgBox("Could not add the entityto db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.CourseSelectionChanged()
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
  End Sub

  Private Sub btnRemCourseParCalendars_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemCourseParCalendars.Click
    Dim cal As Calendar
    Dim bAddEntity As Integer
    curreCrse_ = CourseBindingSource.Current
    cal = CourseParCalendarBindingSourse.Current
    If cal Is Nothing Then Exit Sub

    bAddEntity = omg.RemoveChildEntity(cPAProjects, cal, curreCrse_, "Courses")
    If bAddEntity <> 1 Then
      MsgBox("Could not remove the entity from db")
      Exit Sub
    End If
    'refresh dgv's
    Call Me.CourseSelectionChanged()
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
  End Sub

  Private Sub btnRemCourseChldInstructor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemCourseChldInstructor.Click
    Dim inst As Instructor
    curreCrse_ = CourseBindingSource.Current
    inst = CourseChldInstructorBindingSource.Current
    If inst Is Nothing Then Exit Sub

    nUpdate = omg.RemoveChildEntity(cPAProjects, curreCrse_, inst, "Instructors")
    If nUpdate <> 0 Then nPrevUpdateID = omg.GetSetting("LastUpdateID") 'refresh dgv's
    Call Me.CourseSelectionChanged()
  End Sub

  Private Sub btnAddAvCourseChldInstructor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvCourseChldInstructor.Click
    Dim inst As Instructor
    curreCrse_ = CourseBindingSource.Current()
    inst = AvCourseChldInstructorBindingSource.Current()
    If inst Is Nothing Then Exit Sub

    nUpdate = omg.AddChildEntity(cPAProjects, curreCrse_, inst, "Instructors")
    If nUpdate <> 0 Then nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.CourseSelectionChanged()
  End Sub

End Class
