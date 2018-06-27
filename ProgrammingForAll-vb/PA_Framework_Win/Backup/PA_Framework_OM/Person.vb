Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Person
  Inherits PAEnt

  Dim sFirstName As String
  Dim sMiddleName As String
  Dim sLastName As String
  Dim sLoginID As String
  Dim sEmail As String
  Dim sPhone As String
  Dim sAccessRight As Integer
  Dim sDateJoined As Date
  Dim sRemarks As String
  Dim sEntityItem_1 As String
  Dim sEntityItem_2 As String

  Public Property FirstName() As String
    Get
      FirstName = sFirstName
    End Get
    Set(ByVal value As String)
      sFirstName = value
    End Set
  End Property
  Public Property MiddleName() As String
    Get
      MiddleName = sMiddleName
    End Get
    Set(ByVal value As String)
      sMiddleName = value
    End Set
  End Property

  Public Property LastName() As String
    Get
      LastName = sLastName
    End Get
    Set(ByVal value As String)
      sLastName = value
    End Set
  End Property
  Public Property LoginID() As String
    Get
      LoginID = sLoginID
    End Get
    Set(ByVal value As String)
      sLoginID = value
    End Set
  End Property

  Public Property Email() As String
    Get
      Email = sEmail
    End Get
    Set(ByVal value As String)
      sEmail = value
    End Set
  End Property

  Public Property Phone() As String
    Get
      Phone = sPhone
    End Get
    Set(ByVal value As String)
      sPhone = value
    End Set
  End Property

  Public Property AccessRight() As String
    Get
      AccessRight = sAccessRight
    End Get
    Set(ByVal value As String)
      sAccessRight = value
    End Set
  End Property

  Public Property DateJoined() As Date
    Get
      DateJoined = sDateJoined
    End Get
    Set(ByVal value As Date)
      sDateJoined = value
    End Set
  End Property

  Public Property Remarks() As String
    Get
      Remarks = sRemarks
    End Get
    Set(ByVal value As String)
      sRemarks = value
    End Set
  End Property

  Public Sub New()
    mContainer = cPersons
  End Sub

  Public ReadOnly Property FullName()
    Get
      FullName = FirstName & " " & LastName
    End Get
  End Property

  Public Function GetItemProperty(ByVal sName As String) As String
    Select Case sName
      Case "FirstName"
        Return Me.FirstName
      Case "LastName"
        Return Me.LastName
    End Select
    Return String.Empty
  End Function
End Class
