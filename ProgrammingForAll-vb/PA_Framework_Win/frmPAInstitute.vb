Imports Microsoft.Win32
Imports PA_Framework_OM
'Imports PA_Framework_Lib
Imports System.IO.File

Public Class frmPAInstitute
  Dim WithEvents omg As New OMGlobals
  Dim bEditBegines As Boolean
  Dim nUpdate As Integer ' return value from updates
  Dim sCurrentCellText As String
  Dim arrChangedRows As ArrayList
  Public Shared nPrevUpdateID As Integer  ' local update id
  'Departments
  Dim DepartmentParCourseBindingSource As New BindingSource
  Dim AvDepartmentParCourseBindingSource As New BindingSource

  'Courses
  Dim CoursechldStudentBindingSource As New BindingSource
  Dim AvCourseChldStudentBindingSource As New BindingSource
  'Dim CourseParDepartmentBindingSource As New BindingSource
  'Dim AvCourseParDepartmentBindingSource As New BindingSource
  Dim CourseChldInstructorBindingSource As New BindingSource
  Dim AvCourseChldInstructorBindingSource As New BindingSource
  Dim CourseParCalendarBindingSourse As New BindingSource
  Dim AvCourseParCalendarBindingSource As New BindingSource
  Dim CourseChldDepartmentBindingSource As New BindingSource
  Dim AvCourseChldDepartmentBindingSource As New BindingSource

  'Students
  Dim StudentParCourseBindingSource As New BindingSource
  Dim AvStudentParCourseBindingSource As New BindingSource
  Dim StudentParCalendarBindingSource As New BindingSource
  Dim AvStudentParCalendarBindingSource As New BindingSource
  'Calendar
  Dim CalendarChldCoursesBindingSource As New BindingSource
  Dim AvCalendarChldCoursesBindingSource As New BindingSource
  Dim CalendarChldStudentsBindingSource As New BindingSource
  Dim AvCalendarChldStudentsBindingSource As New BindingSource

  'Attendance
  Dim AttendanceChldCalendarsBindingSource As New BindingSource
  Dim AvAttendanceChldCalendarsBindingSource As New BindingSource
  Dim AttendanceChldStudentsBindingSource As New BindingSource
  Dim avAttendanceChldStudentsBindingSource As New BindingSource

  'Instructor
  Dim InstructorChldPersonBindingSource As New BindingSource
  Dim AvInstructorChldPersonBindingSource As New BindingSource
  Dim InstructorParCourseBindingSource As New BindingSource
  Dim AvInstructorParCourseBindingSource As New BindingSource

  'Associates
  'Dim AssociateBindingSource As New BindingSource
  Dim AssociateChldPersonBindingSource As New BindingSource
  Dim AvAssociateChldPersonBindingSource As New BindingSource
  'Person
  Dim PersonParInstructorBindingSource As New BindingSource
  Dim AvPersonParInstructorBindingSource As New BindingSource
  Dim PersonParStudentBindingSource As New BindingSource
  Dim AvPersonParStudentBindingSource As New BindingSource
  Dim PersonParAssociateBindingSource As New BindingSource
  Dim AvPersonParAssociateBindingSource As New BindingSource

  'PA Projects
  Dim PAProjectChldProjectEntitiesBindingSource As New BindingSource
  Dim AvPAProjectChldProjectEntitiesBindingSource As New BindingSource

  'project Entity
  Dim ProjectEntityChldPrjEntItemBindingSource As New BindingSource
  Dim AvProjectEntityChldPrjEntItemBindingSource As New BindingSource

  'PrjEntItem
  Dim PrjEntItemParProjectEntityBindingSource As New BindingSource
  Dim AvPrjEntItemParProjectEntityBindingSource As New BindingSource

  'ChangeHistory
  Dim entColl As New ChangeHistorys
  Dim ChangeRowHistoryBindingSource As New BindingSource

  Private Sub PA_Institute_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    'omg = OMGlobals.GetInstance
    'omg = omg.GetOMGInstance 'OMGlobals.GetInstance()
    'Debug.Print(Application.StartupPath)
    Dim sUser As String, sAccessRight As String
    If Not System.IO.File.Exists("C:\" & OMGlobals.APPLongName & "\Settings.xml") Then
      AppSettings.InitializeSettings()
    End If
    AppSettings.SetUserDetails()
    'Call omg.DBLoad()
    Call omg.PassFormToOM(Me)

    sUser = AppSettings.GetSetting("LoginName")
    sAccessRight = AppSettings.GetSetting("AccessRights")
    Me.StatusStrip1.Items("stsLoggedInUser").Text = sUser & " / " & sAccessRight
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
  End Sub

  Private Sub LoadDataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadDataToolStripMenuItem.Click
    Call omg.DBLoad()
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
  End Sub

  Private Sub TabPageChangeHistory_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPageChangeHistory.Enter
    'ChangeHistorysBindingSource.Clear()
    OMGlobals.cChangeHistorys = New ChangeHistorys
    OMGlobals.cChangeHistorys.Load()

    ChangeHistoryBindingSource.DataSource = OMGlobals.cChangeHistorys
    Me.dgvChangeHistory.DataSource = ChangeHistoryBindingSource

  End Sub

  Private Sub TabHistory_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabHistory.Enter
    'ChangeHistorysBindingSource.Clear()
    OMGlobals.cChangeHistorys = New ChangeHistorys
    OMGlobals.cChangeHistorys.Load()

    ChangeHistoryBindingSource.DataSource = OMGlobals.cChangeHistorys
    Me.dgvHistory.DataSource = ChangeHistoryBindingSource
  End Sub

  Private Sub ChangeHistoryBindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChangeHistoryBindingSource.CurrentChanged
    Dim ent As ChangeHistory

    Dim sText As String = Nothing
    Dim currHist As ChangeHistory
    currHist = ChangeHistoryBindingSource.Current
    entColl.RemoveAll()
    For Each ent In OMGlobals.cChangeHistorys
      If currHist.KeyField = ent.KeyField And currHist.Table = ent.Table Then
        entColl.Add(ent)
        'sText = sText & ent.DateTime & "," & ent.User & "," & ent.Changes & vbCrLf
      End If
      'txtChangeHistory.Text=sText
    Next
    txtTable.Text = currHist.Table
    txtKeyField.Text = currHist.KeyField
    ChangeRowHistoryBindingSource.DataSource = entColl
    Me.dgvRowChangeHistory.DataSource = ChangeRowHistoryBindingSource
    ChangeRowHistoryBindingSource.ResetBindings(True)

  End Sub

  Private Sub SettingsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SettingsToolStripMenuItem.Click
    Dim frmSet As New frmSettings
    frmSet.ShowDialog()
  End Sub

  Private Sub UsersGuideToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsersGuideToolStripMenuItem.Click
    'System.IO.File.Open(Application.StartupPath & "\PA_Framework_Users_Guide.ppt", IO.FileMode.Open, IO.FileAccess.ReadWrite, IO.FileShare.None)
    'FileOpen(FreeFile, Application.StartupPath & "\" & "PA Framework Users Guide.ppt", OpenMode.Binary, OpenAccess.Read, OpenShare.Default)
    Process.Start("C:\" & OMGlobals.APPLongName & "\PA_Framework_Users_Guide_1.0.pdf")
  End Sub

  Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
    'Dim ffff As New frmIntro
    'ffff.Show()
  End Sub

  Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
    Me.Close()
  End Sub

  Private Sub dgvAvPrjEntsChldEntItems_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvAvPrjEntsChldEntItems.DataError

  End Sub

  Private Sub LoadDBFormToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadDBFormToolStripMenuItem.Click
    Dim frmDB As New frmCodeGenerator
    frmDB.Show()
  End Sub

  Private Sub omg_EntityUpdated(ByVal sUpdateType As String, ByRef oEntity As Object, ByVal sReturn As Integer) Handles omg.EntityUpdated
    'EntityUpdated event raised in DBUpdate function in OMGlobals class. All update messages are defined in this function.
    Select Case sUpdateType
      Case "Update"
        Select Case sReturn
          Case 0
            MsgBox("This record (ID:" & oEntity.ID & ") has been updated since you last refreshed. Please load data again and then update.")
          Case 1
            MsgBox(TypeName(oEntity) & " " & oEntity.id & " has been updated")
          Case -1
            MsgBox("Nothing to Update")
        End Select
      Case "Add"
        Select Case sReturn
          Case 0
            MsgBox("A record has been added to the DB since you last refreshed. Please Load Data again")
        End Select

      Case "Delete"
        Select Case sReturn
          Case 0
            MsgBox("Deletion of " & TypeName(oEntity) & " " & oEntity.ID & " Failed")
          Case 1
            MsgBox("Successully deleted " & TypeName(oEntity) & " " & oEntity.ID)
        End Select
    End Select
  End Sub
End Class

