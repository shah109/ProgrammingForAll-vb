Option Explicit On
Imports Microsoft.Win32
'Imports PA_Framework_Lib
Imports PA_Framework_OM
Imports PA_Framework_OM.OMGlobals

'Partial Public Class frmPAInstitute

'  Private Sub TabAttendances_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabAttendances.Enter
'    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
'    AttendanceBindingSource.DataSource = cAttendances
'    AttendanceBindingSourceChanged()
'  End Sub

'  Private Sub AttendanceBindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AttendanceBindingSource.CurrentChanged
'    curreAtt_ = AttendanceBindingSource.Current
'    AttendanceBindingSourceChanged()
'  End Sub

'  Sub AttendanceBindingSourceChanged()
'    If curreAtt_ Is Nothing Then Exit Sub
'    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
'    'Attendance child Calendars
'    Dim cal As New Calendar
'    Call omg.FillChildEntities(cPAProjects, curreAtt_, cal, "CalendarItems")
'    AttendanceChldCalendarsBindingSource.DataSource = curreAtt_.ChildEntities("CalendarItems")
'    Me.dgvAttendanceChldCalendars.DataSource = AttendanceChldCalendarsBindingSource
'    AttendanceChldCalendarsBindingSource.ResetBindings(True)

'    'Attendance available child Calendars
'    Call omg.FillAvailableChildEntities(cPAProjects, curreAtt_, cal, "CalendarItems")
'    AvAttendanceChldCalendarsBindingSource.DataSource = curreAtt_.AvailableChildEntities(cal)
'    Me.dgvAvAttendanceChldCalendars.DataSource = AvAttendanceChldCalendarsBindingSource
'    AvAttendanceChldCalendarsBindingSource.ResetBindings(True)

'    'Attendance child Students
'    Dim std As New Student
'    Call omg.FillChildEntities(cPAProjects, curreAtt_, std, "Students")
'    AttendanceChldStudentsBindingSource.DataSource = curreAtt_.ChildEntities("Students")
'    Me.dgvAttendanceChldStudents.DataSource = AttendanceChldStudentsBindingSource
'    AttendanceChldStudentsBindingSource.ResetBindings(True)

'    'Attendance available child Students
'    Call omg.FillAvailableChildEntities(cPAProjects, curreAtt_, std, "Students")
'    avAttendanceChldStudentsBindingSource.DataSource = curreAtt_.AvailableChildEntities(std)
'    Me.dgvAvAttendanceChldStudents.DataSource = avAttendanceChldStudentsBindingSource
'    avAttendanceChldStudentsBindingSource.ResetBindings(True)

'    'AttendanceBindingSource.ResetBindings(True)
'  End Sub

'  Private Sub btnAddAttendance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAttendance.Click
'    Dim att = New Attendance
'    att.Comments = "new"
'    Call omg.DBUpdate("Add", cAttendances, att)
'    nPrevUpdateID = omg.GetSetting("LastUpdateID")
'    AttendanceBindingSource.ResetBindings(True)
'    AttendanceBindingSource.MoveLast()

'  End Sub

'  Private Sub btnDeleteAttendance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteAttendance.Click
'    'no deletetion allowed
'  End Sub

'  Private Sub btnUpdateAttendance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateAttendance.Click
'    curreAtt_ = AttendanceBindingSource.Current
'    Call omg.DBUpdate("Update", cAttendances, curreAtt_)
'    nPrevUpdateID = omg.GetSetting("LastUpdateID")
'  End Sub

'  Private Sub btnRemAttendanceChldCalendar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemAttendanceChldCalendar.Click
'    Dim cal As Calendar
'    curreAtt_ = AttendanceBindingSource.Current
'    cal = AttendanceChldCalendarsBindingSource.Current
'    omg.RemoveChildEntity(cPAProjects, curreAtt_, cal, "CalendarItems")
'    nPrevUpdateID = omg.GetSetting("LastUpdateID")
'    'refresh dgv's
'    Call Me.AttendanceBindingSourceChanged()
'  End Sub

'  Private Sub btbnAddAvAttendanceChldCalendar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAvAttendanceChldCalendar.Click
'    Dim crs As Calendar
'    curreAtt_ = AttendanceBindingSource.Current()
'    crs = AvAttendanceChldCalendarsBindingSource.Current()
'    omg.AddChildEntity(cPAProjects, curreAtt_, crs, "CalendarItems")
'    nPrevUpdateID = omg.GetSetting("LastUpdateID")

'    'refresh dgv's
'    Call Me.AttendanceBindingSourceChanged()
'  End Sub

'    Private Sub btnAddAvAttendanceChldStudent_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddAvAttendanceChldStudent.Click
'        Dim std As Student
'        curreAtt_ = AttendanceBindingSource.Current()
'        std = avAttendanceChldStudentsBindingSource.Current()
'    omg.AddChildEntity(cPAProjects, curreAtt_, std, "Students")
'    nPrevUpdateID = omg.GetSetting("LastUpdateID")
'        'refresh dgv's
'        Call Me.AttendanceBindingSourceChanged()
'    End Sub

'    Private Sub btnRemAttendanceChldStudent_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRemAttendanceChldStudent.Click
'        Dim std As Student
'        curreAtt_ = AttendanceBindingSource.Current
'        std = AttendanceChldStudentsBindingSource.Current
'    omg.RemoveChildEntity(cPAProjects, curreAtt_, std, "Students")
'    nPrevUpdateID = omg.GetSetting("LastUpdateID")
'        'refresh dgv's
'        Call Me.AttendanceBindingSourceChanged()
'    End Sub
'End Class
