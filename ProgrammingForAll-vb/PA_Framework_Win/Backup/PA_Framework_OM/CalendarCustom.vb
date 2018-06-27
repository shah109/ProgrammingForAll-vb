Partial Public Class calendar
  Public Function DidStudentAttend(ByVal std As Student) As Boolean
    Dim st As Student
    DidStudentAttend = False
    For Each st In Me.ChildEntities("Students")
      If std.ID = st.ID Then
        Return True
        Exit Function
      End If
    Next st
  End Function

  Public Function SortbyDate(ByRef objCal As Calendars) As Calendars
    SortbyDate = Nothing
    Dim ocal As Calendar
    Dim ocal1 As Calendar
    For Each ocal In Me.mContainer
      For Each ocal1 In Me.mContainer
      Next
    Next
  End Function
End Class
