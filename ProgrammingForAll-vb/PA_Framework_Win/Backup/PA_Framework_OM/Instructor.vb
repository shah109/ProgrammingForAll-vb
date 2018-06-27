Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Instructor
  Inherits PAEnt
  Dim sInstructorName As String
  Dim dDateStarted As DateTime
  Dim sComments As String

  Public Property InstructorName() As String
    Get
      InstructorName = sInstructorName
    End Get
    Set(ByVal value As String)
      sInstructorName = value
    End Set
  End Property


  Public Property DateStarted() As DateTime
    Get
      DateStarted = dDateStarted
    End Get
    Set(ByVal value As DateTime)
      dDateStarted = value
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

  Public Sub New()
    mContainer = cInstructors
  End Sub

  Public Function GetItemProperty(ByVal sName As String) As String
    Select Case sName
      Case "InstructorName"
        Return Me.InstructorName
      Case "DateStarted"
        Return Me.DateStarted
    End Select
    Return String.Empty
  End Function

End Class
