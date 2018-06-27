Imports System
Imports System.Collections.Generic
Imports System.Text

Public Class OMForExcel
  Dim strAssoc As String
  Dim nUpdate As Integer

  Public Function DotNetMethod(ByVal input1 As String) As String
    Return "Hello " & input1
  End Function

  Public Function loadObjectModel()
    Call MGlobals.DBLoad()
  End Function

  Public Function GetCollections(ByVal strObj As String) As Object
    Select Case strObj
      Case "Departments"
        GetCollections = cDepartments
      Case "Courses"
        GetCollections = cCourses
      Case "Students"
        GetCollections = cStudents
      Case "Persons"
        GetCollections = cPersons
      Case "Calendars"
        GetCollections = cCalendars
      Case "Attendances"
        GetCollections = cAttendances
      Case "EntityDataItems"
        GetCollections = cEntityDataItems
      Case "UpdateLogItems"
        GetCollections = cUpdateLogItems
      Case "AppArrays"
        GetCollections = cAppArrays
    End Select
  End Function

  Public Function gAddChildEntity(ByVal enParent As Object, ByVal enChld As Object, ByVal sChildPropertyName As String) As Integer
    'returns 0 if successful, 1 if the record is updated since last read, 2 if the entity attempted to add has already been added (1-M)
    'Adds a child entity ehchld to entity enParent
    Dim strAssoc As String
    Dim nUpdate As Integer
    Dim bAddEntity As Boolean
    Call MGlobals.CallDBLoadIfNeeded()  ' syncs with db before calling the next function.
    bAddEntity = True

    strAssoc = cEntityDataItems.GetAssociation(TypeName(enParent), TypeName(enChld), gsChildPropertyName, gsJoinTable)
    If strAssoc = "1-M" Then    ' check to see if the Child entity is still available.
      bAddEntity = Not BLFunctions.gCheckAllParentsForChildPresence(enParent.mContainer, enChld, sChildPropertyName)
    End If
    If strAssoc = "M-1" And enParent.ChildEntities(sChildPropertyName).Count <> 0 Then      ' only one child allowed for M-1.
      gAddChildEntity = 3
      Exit Function
    End If
    If bAddEntity = False Then
      MsgBox("The entity you are attempting to Add is not available any more.It has been added to another entity", vbOKOnly)
      gAddChildEntity = 2  'entity is no more available because it has been added to another entity (1-M)
      Exit Function
    End If

    If strAssoc = "M-M" And gsJoinTable <> "CSR" Then  'JTFK's
      mObjectForJoinTable = CreateObjectFromString(gsJoinTable)
      gAddChildEntity = gAddChildEntityViaJoin(gsJoinTable, enParent, enChld, gsChildPropertyName)
      Exit Function
    End If
    ' start of CSFK's
    Call enParent.ChildEntities(sChildPropertyName).Add(enChld)
    Call enParent.BuildChildEntityString(gsChildPropertyName)
    nUpdate = enParent.mContainer.UpdateDB(enParent)
    If nUpdate = 0 Then  'revert the child addition
      Call enParent.ChildEntities(enChld).Remove(enChld)
      Call enParent.BuildChildEntityString(enParent.ChildEntities(enChld))
      MsgBox("This record is not present or has been updated since you last refreshed. Please Load Data again and then update.")
      gAddChildEntity = 0
      Exit Function   'db not updated because the record has been updated since last refresh.
    End If
    gAddChildEntity = 1  'success
  End Function

  Public Function gRemoveChildEntity(ByVal enParent As Object, ByVal enChld As Object, ByVal sChildPropertyName As String) As Integer
    'Removes a child entity enChld from the Parent entity enParent

    strAssoc = cEntityDataItems.GetAssociation(TypeName(enParent), TypeName(enChld), sChildPropertyName, gsJoinTable)
    If strAssoc = "M-M" And gsJoinTable <> "CSR" Then  'JTFK's
      mObjectForJoinTable = CreateObjectFromString(gsJoinTable)
      gRemoveChildEntity = gRemoveChildEntityViaJoin(gsJoinTable, enParent, enChld, sChildPropertyName)
      Exit Function
    End If
    Call enParent.ChildEntities(sChildPropertyName).Remove(enChld)
    Call enParent.BuildChildEntityString(sChildPropertyName)
    nUpdate = enParent.mContainer.UpdateDB(enParent)
    If nUpdate = 1 Then
      gRemoveChildEntity = nUpdate
      Exit Function
    End If
    Call enParent.ChildEntities(sChildPropertyName).Add(enChld)  'revert
    Call enParent.BuildChildEntityString(enChld)
    MsgBox("This record has been updated since you last refreshed. Please refresh again and then update.")
    gRemoveChildEntity = nUpdate
  End Function
  Public Sub ShowHello()
    MsgBox("Hello")
  End Sub
End Class
