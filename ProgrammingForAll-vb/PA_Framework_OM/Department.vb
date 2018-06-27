Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Department
  Inherits PAEnt

  Dim sEntityItem_1 As String
  Dim sEntityItem_2 As String

  Public Property EntityItem_1() As String
    Get
      EntityItem_1 = sEntityItem_1
    End Get
    Set(ByVal value As String)
      sEntityItem_1 = value
    End Set
  End Property

  Public Property EntityItem_2() As String
    Get
      EntityItem_2 = sEntityItem_2
    End Get
    Set(ByVal value As String)
      sEntityItem_2 = value
    End Set
  End Property

  'Public ReadOnly Property Courses() As String
  '  Get
  '    Courses = sChildCoursesString
  '  End Get
  'End Property

  Public Sub New()
    mContainer = cDepartments
  End Sub

  Public Function GetItemProperty(ByVal sName As String) As String
    Select Case sName
      Case "EntityItem_1"
        Return Me.EntityItem_1
      Case "EntityItem_2"
        Return Me.EntityItem_2
    End Select
    Return String.Empty
  End Function


End Class
