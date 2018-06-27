Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Associate
  Inherits PAEnt
  Dim sComments As String

  'Dim mChildPersons As New Persons
  'Dim sChildPersonsString As String = ""
  'Dim mAvChildPersons As New Persons

  'Dim mParentCourses As New Courses
  'Dim mAvParentCourses As New Courses

  'Dim mParentCalendars As New Calendars

  Public Property Comments() As String
    Get
      Comments = sComments
    End Get
    Set(ByVal value As String)
      sComments = value
    End Set
  End Property

  Public Sub New()
    mContainer = cAssociates
  End Sub
End Class
