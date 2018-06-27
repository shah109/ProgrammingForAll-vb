Option Explicit On
Imports Microsoft.Win32
'Imports PA_Framework_Lib
Imports PA_Framework_OM
Imports PA_Framework_OM.OMGlobals

Partial Public Class frmPAInstitute

  Private Sub TabCalendar_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabCalendars.Enter
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    CalendarBindingSource.DataSource = omg.getCalendars
    CalendarBindingSourceChanged()


  End Sub

  Private Sub CalendarBindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CalendarBindingSource.CurrentChanged
    curreCal_ = CalendarBindingSource.Current
    CalendarBindingSourceChanged()
  End Sub

  Sub CalendarBindingSourceChanged()
    If curreCal_ Is Nothing Then Exit Sub
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    Dim nCount As Integer, strParDetails As String = "", strChldDetails = ""
    nCount = cPAProjects.GetEntityDependencies(curreCal_, strParDetails, strChldDetails)
    Me.txtCalendarDependencies.Text = CStr(nCount)
    If nCount <> 0 Then
      Me.btnDeleteCalendar.Enabled = False
    Else
      Me.btnDeleteCalendar.Enabled = True
    End If

    Me.txtCalendarParDeps.Text = strParDetails
    Me.txtCalendarChldDeps.Text = strChldDetails

    'Calendar child Courses 
    Dim crs As New Course
    Call omg.FillChildEntities(cPAProjects, curreCal_, crs, "Courses")
    CalendarChldCoursesBindingSource.DataSource = curreCal_.ChildEntities("Courses")
    Me.dgvCalendarChldCourses.DataSource = CalendarChldCoursesBindingSource
    CalendarChldCoursesBindingSource.ResetBindings(True)

    'Calendar available child courses

    Call omg.FillAvailableChildEntities(cPAProjects, cCourses, curreCal_, crs, "Courses")
    AvCalendarChldCoursesBindingSource.DataSource = curreCal_.AvailableChildEntities(crs)
    Me.dgvAvCalendarChldCourses.DataSource = AvCalendarChldCoursesBindingSource
    AvCalendarChldCoursesBindingSource.ResetBindings(True)

    'Calendar child Students
    Dim std As New Student
    Call omg.FillChildEntities(cPAProjects, curreCal_, std, "Students")
    CalendarChldStudentsBindingSource.DataSource = curreCal_.ChildEntities("Students")
    Me.dgvCalendarChldStudents.DataSource = CalendarChldStudentsBindingSource
    CalendarChldStudentsBindingSource.ResetBindings(True)
    
    'Calendar available child Students
    'Available child students should only be the ones who are in the associated course
    Dim CrsForCurrLecture As New Course
    If curreCal_.ChildEntities("Courses").Count <> 0 Then
      CrsForCurrLecture = curreCal_.ChildEntities("Courses").LoadOrder(1)
      Console.WriteLine("CrsForCurrLecture child students Count: {0:d}", CrsForCurrLecture.mChildStudents.Count)
    End If
    ''''''''''''''''''''''''''

    'Calendar available child Students. Dont have to iterate through all students, only the students in the associated course.
    Call omg.FillAvailableChildEntities(cPAProjects, CrsForCurrLecture.mChildStudents, curreCal_, std, "Students")
    'Call omg.FillAvailableChildEntities(cPAProjects, curreCal_, std, "Students")
    AvCalendarChldStudentsBindingSource.DataSource = curreCal_.AvailableChildEntities(std)
    Me.dgvAvCalendarChldStudents.DataSource = AvCalendarChldStudentsBindingSource
    AvCalendarChldStudentsBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnAddCalendar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddCalendar.Click
    Dim np = New Calendar
    np.Comments = "new"
    Call omg.DBUpdate("Add", cCalendars, np)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    CalendarBindingSource.ResetBindings(True)
    CalendarBindingSource.MoveLast()
  End Sub

  Private Sub btnDeleteCalendar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteCalendar.Click
    Dim nCount As Integer, strChldDetails As String = "", strParDetails As String = ""
    curreCal_ = CalendarBindingSource.Current
    nCount = cPAProjects.GetEntityDependencies(curreCal_, strParDetails, strChldDetails)
    If nCount = 0 Then
      omg.DBUpdate("Delete", cCalendars, curreCal_)
      nPrevUpdateID = omg.GetSetting("LastUpdateID")
    Else
      MsgBox("Can not delete because of dependencies ")
    End If
    CalendarBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnUpdateCalendar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateCalendar.Click
    curreCal_ = CalendarBindingSource.Current
    Call omg.DBUpdate("Update", cCalendars, curreCal_)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
  End Sub

  Private Sub btnRemCalendarChldCourse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemCalendarChldCourse.Click
    Dim crs As Course
    curreCal_ = CalendarBindingSource.Current
    crs = CalendarChldCoursesBindingSource.Current
    If crs Is Nothing Then Exit Sub
    omg.RemoveChildEntity(cPAProjects, curreCal_, crs, "Courses")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.CalendarBindingSourceChanged()
  End Sub

  Private Sub btnAddAvCalendarChldCourse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvCalendarChldCourse.Click
    Dim crs As Course
    curreCal_ = CalendarBindingSource.Current()
    crs = AvCalendarChldCoursesBindingSource.Current()
    If crs Is Nothing Then Exit Sub
    omg.AddChildEntity(cPAProjects, curreCal_, crs, "Courses")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.CalendarBindingSourceChanged()
  End Sub

  Private Sub btnAddAvCalendarChldStudent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvCalendarChldStudent.Click
    Dim std As Student
    curreCal_ = CalendarBindingSource.Current()
    std = AvCalendarChldStudentsBindingSource.Current()
    If std Is Nothing Then Exit Sub
    omg.AddChildEntity(cPAProjects, curreCal_, std, "Students")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.CalendarBindingSourceChanged()
  End Sub

  Private Sub btnRemCalendarChldStudent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemCalendarChldStudent.Click
    Dim std As Student
    curreCal_ = CalendarBindingSource.Current
    std = CalendarChldStudentsBindingSource.Current
    If std Is Nothing Then Exit Sub
    omg.RemoveChildEntity(cPAProjects, curreCal_, std, "Students")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.CalendarBindingSourceChanged()
  End Sub

End Class

