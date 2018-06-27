Option Explicit On

Public Class Attendance
  'sPH_cls_DateTime:
  'Framework File Attendance Copied : 9/22/2011 6:18:12 PM
  'sPH_cls_DateTime:End
  'sh-Attendance

  Public Loadorder As Integer  'Load order
  'PH:For Each Entity Item. Create fields with clear naming conventions and with the required data types
  'sPH_cls_Decl:
  Dim sid As String
  Dim sComments As String
  'sPH_cls_Decl:End
  'PH: End

  ''PH:For Each Child. Add two fields for each child entity this entity supports.
  'Replace all variable of Entity3 with your own <EntityName> variables
  'Dim mChildEntity3s As New Entity3s
  'Dim sChildEntity3sString As String
  ''PH: End
  'sPH_cls_ChildDecl:
  Dim mChildStudents As New Students
  Dim sChildStudentsString As String
  Dim mChildCalendarItems As New Calendars
  Dim sChildCalendarItemsString As String
  'sPH_cls_ChildDecl:End

  Public mContainer As Object
  Dim sLastUpdate As Date

  Public Property ID() As String
    Get
      ID = sid
    End Get
    Set(ByVal value As String)
      sid = value
    End Set
  End Property

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

  'sPH_cls_Properties:End

  Public Property Lastupdate() As Date
    Get
      Lastupdate = sLastUpdate
    End Get
    Set(ByVal value As Date)
      sLastUpdate = value
    End Set
  End Property

  'Child Entity Methods
  Public Function ChildEntities(ByVal sEnt As String) As Object
    Select Case sEnt
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3", "Entity3s"
      '    Set ChildEntities = mChildEntity3s
      '  'PH: End
      'sPH_cls_ChildEntities:
      Case "Students"
        ChildEntities = mChildStudents
      Case "CalendarItems"
        ChildEntities = mChildCalendarItems
        'sPH_cls_ChildEntities:End
    End Select
  End Function

  Public Function ChildEntityString(ByVal sEnt As String) As String
    Select Case sEnt
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3", "Entity3s":
      '    ChildEntityString = sChildEntity3sString
      '  'PH: End
      'sPH_cls_GetChildString:
      Case "Students"
        ChildEntityString = sChildStudentsString
      Case "CalendarItems"
        ChildEntityString = sChildCalendarItemsString
        'sPH_cls_GetChildString:End
    End Select
  End Function

  Public Function ChildEntityString(ByVal sEnt As String, ByVal strEnt As String)
    Select Case sEnt
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3s", "Entity3":
      '    sChildEntity3sString = strEnt
      '  'PH: End
      'sPH_cls_LetChildString:
      Case "Students"
        sChildStudentsString = strEnt
      Case "CalendarItems"
        sChildCalendarItemsString = strEnt
        'sPH_cls_LetChildString:End
    End Select
  End Function

  Public Function BuildChildEntityObjects(ByVal strPar As String, ByVal strEnt As String) As Object
    Call BLFunctions.gBuildChildEntityObjects(Me, strPar, strEnt)
  End Function

  Public Function BuildChildEntityString(ByRef enChlds As String) As String
    Call gBuildChildEntityString(Me, enChlds)
  End Function

  ' Parent Entity Methods.
  Public Function ParentEntities(ByVal objPar As Object)
    Select Case TypeName(objPar)
      ''PH: For Each Parent entity, add a case statement
      '  Case "Entity1", "Entity1s":
      '    Set ParentEntities = mParentEntity1s
      ' 'PH: End
      'sPH_cls_ParentEntities:
      ''No Entries
      'sPH_cls_ParentEntities:End
    End Select
  End Function

  Sub New()
    mContainer = cAttendances
  End Sub

  Public Function GetItemProperty(ByVal strP As String) As String
    Select Case strP
      'sPH_cls_GetItemProperty:
      Case "ID"
        GetItemProperty = Me.ID
      Case "Comments"
        GetItemProperty = Me.Comments
        'sPH_cls_GetItemProperty:End
    End Select
  End Function



End Class
