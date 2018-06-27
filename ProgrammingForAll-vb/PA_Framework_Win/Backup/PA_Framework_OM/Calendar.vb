Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Calendar
  Inherits PAEnt

  Dim sLectureDate As Date
  Dim sLectureLocation As String
  Dim sComments As String

 
  Public ReadOnly Property CourseID() As String
    Get
      CourseID = sChildCoursesString
    End Get
  End Property

  Public Property LectureDate() As Date
    Get
      LectureDate = sLectureDate
    End Get
    Set(ByVal value As Date)
      sLectureDate = value
    End Set
  End Property

  Public Property Location() As String
    Get
      Location = sLectureLocation
    End Get
    Set(ByVal value As String)
      sLectureLocation = value
    End Set
  End Property
  Public Property Comments() As String
    Get
      Comments = sComments
    End Get
    Set(ByVal value As String)
      sComments = value
    End Set
  End Property

  Sub New()
    mContainer = cCalendars
  End Sub

  Public Function GetItemProperty(ByVal sName As String) As String
    Select Case sName
      Case "Location"
        Return Me.Location
      Case "LectureDate"
        Return Me.LectureDate
    End Select
    Return String.Empty
  End Function


End Class
