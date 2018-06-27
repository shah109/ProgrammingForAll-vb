Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Course
  Inherits PAEnt

  Dim sEntityItem_1 As String
  Dim sEntityItem_2 As String

  Public Property Name() As String
    Get
      Name = sEntityItem_1
    End Get
    Set(ByVal value As String)
      sEntityItem_1 = value
    End Set
  End Property

  Public Property Description() As String
    Get
      Description = sEntityItem_2
    End Get
    Set(ByVal value As String)
      sEntityItem_2 = value
    End Set
  End Property

  Public Sub New()
    mContainer = cCourses
  End Sub

  Public ReadOnly Property PersonName() As String
    Get
      If mChildPersons.Count <> 0 Then
        PersonName = mChildPersons.Item(0).FirstName
      Else
        PersonName = ""
      End If
    End Get
  End Property

  Public Function GetItemProperty(ByVal sName As String) As String
    Select Case sName
      Case "Name"
        Return Me.Name
      Case "Description"
        Return Me.Description
    End Select
    Return String.Empty
  End Function

End Class
