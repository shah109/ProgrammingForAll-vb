﻿Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Attendance
  Inherits PAEnt
  Dim sComments As String

  Public ReadOnly Property CalendarID() As String
    Get
      CalendarID = sChildCalendarItemsString
    End Get
  End Property

  Public ReadOnly Property StudentID() As String
    Get
      StudentID = sChildStudentsString
    End Get
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
    mContainer = cAttendances
  End Sub

End Class
