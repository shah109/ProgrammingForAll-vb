Imports PA_Framework_Lib
Public Module BLFunctions
  'sPH_mod_DateTime:
  'Framework File BLFunctions Copied : 9/15/2011 6:19:53 PM
  'sPH_mod_DateTime:End

  Dim nUpdate As Integer
  Dim strAssoc As String
  Dim bAddEntity As Boolean
  Dim strParentPropertyName As String

  Function gAddChildEntityViaJoin(ByVal sEntJt As String, ByVal enParent As Object, ByVal enChld As Object, ByVal sChildPropertyName As String) As Integer
    Call cEntityDataItems.GetAssociation(sEntJt, TypeName(enParent), strParentPropertyName, "")
    'Call cEntityDataItems.GetAssociation(sEntJt, TypeName(enChld), strChildPropertyName, "")
    mObjectForJoinTable.ChildEntityString(strParentPropertyName) = enParent.ID
    mObjectForJoinTable.ChildEntityString(sChildPropertyName) = enChld.ID

    Call mObjectForJoinTable.ChildEntities(strParentPropertyName).Add(enParent)
    Call mObjectForJoinTable.ChildEntities(sChildPropertyName).Add(enChld)

    Call enParent.ChildEntities(sChildPropertyName).Add(enChld)

    nUpdate = mObjectForJoinTable.mContainer.AddtoDB(mObjectForJoinTable)
    If nUpdate = 0 Then  'revert
      Call enParent.ChildEntities(enChld).Remove(enChld)
      MsgBox("Update Error: Parent ID:" & enParent.ID)
    End If

    gAddChildEntityViaJoin = nUpdate
  End Function
  '
  Function gRemoveChildEntityViaJoin(ByVal sEntJt As String, ByVal enParent As Object, ByVal enChld As Object, ByVal sChildPropertyName As String) As Integer
    Dim entJt As Object
    Dim entCont As New Object
    entJt = GetItemForJoinTable(sEntJt, enParent, enChld, sChildPropertyName)

    'Call cEntityDataItems.GetAssociation(sEntJt, TypeName(enChld), strChildPropertyName, "")
    Call enParent.ChildEntities(sChildPropertyName).Remove(enChld)
    entCont = entJt.mContainer
    nUpdate = entCont.DeleteFromDB(entJt)
    If nUpdate = 0 Then  'revert
      Call enParent.ChildEntities(enChld).Add(enChld)
      MsgBox("Update Error: Parent ID:" & enParent.ID)
    End If
    gRemoveChildEntityViaJoin = nUpdate
  End Function


  'Function gAddChildEntity(ByVal enParent As Object, ByVal enChld As Object, ByVal sChildPropertyName As String) As Integer
  '  'returns 0 if successful, 1 if the record is updated since last read, 2 if the entity attempted to add has already been added (1-M)
  '  'Adds a child entity ehchld to entity enParent
  '  Dim strAssoc As String
  '  Dim bAddEntity As Boolean
  '  Call MGlobals.CallDBLoadIfNeeded()  ' syncs with db before calling the next function.
  '  bAddEntity = True

  '  strAssoc = cEntityDataItems.GetAssociation(TypeName(enParent), TypeName(enChld), gsChildPropertyName, gsJoinTable)
  '  If strAssoc = "1-M" Then    ' check to see if the Child entity is still available.
  '    bAddEntity = Not BLFunctions.gCheckAllParentsForChildPresence(enParent.mContainer, enChld, sChildPropertyName)
  '  End If
  '  If strAssoc = "M-1" And enParent.ChildEntities(sChildPropertyName).Count <> 0 Then      ' only one child allowed for M-1.
  '    gAddChildEntity = 3
  '    Exit Function
  '  End If
  '  If bAddEntity = False Then
  '    MsgBox("The entity you are attempting to Add is not available any more.It has been added to another entity", vbOKOnly)
  '    gAddChildEntity = 2  'entity is no more available because it has been added to another entity (1-M)
  '    Exit Function
  '  End If

  '  If strAssoc = "M-M" And gsJoinTable <> "CSR" Then  'JTFK's
  '    mObjectForJoinTable = CreateObjectFromString(gsJoinTable)
  '    gAddChildEntity = gAddChildEntityViaJoin(gsJoinTable, enParent, enChld, gsChildPropertyName)
  '    Exit Function
  '  End If
  '  ' start of CSFK's
  '  Call enParent.ChildEntities(sChildPropertyName).Add(enChld)
  '  Call enParent.BuildChildEntityString(gsChildPropertyName)
  '  nUpdate = enParent.mContainer.UpdateDB(enParent)
  '  If nUpdate = 0 Then  'revert the child addition
  '    Call enParent.ChildEntities(enChld).Remove(enChld)
  '    Call enParent.BuildChildEntityString(enParent.ChildEntities(enChld))
  '    MsgBox("This record is not present or has been updated since you last refreshed. Please Load Data again and then update.")
  '    gAddChildEntity = 0
  '    Exit Function   'db not updated because the record has been updated since last refresh.
  '  End If
  '  gAddChildEntity = 1  'success
  'End Function

  'global
  'Public Function gRemoveChildEntity(ByVal enParent As Object, ByVal enChld As Object, ByVal sChildPropertyName As String) As Integer
  '  'Removes a child entity enChld from the Parent entity enParent

  '  strAssoc = cEntityDataItems.GetAssociation(TypeName(enParent), TypeName(enChld), sChildPropertyName, gsJoinTable)
  '  If strAssoc = "M-M" And gsJoinTable <> "CSR" Then  'JTFK's
  '    mObjectForJoinTable = CreateObjectFromString(gsJoinTable)
  '    gRemoveChildEntity = gRemoveChildEntityViaJoin(gsJoinTable, enParent, enChld, sChildPropertyName)
  '    Exit Function
  '  End If
  '  Call enParent.ChildEntities(sChildPropertyName).Remove(enChld)
  '  Call enParent.BuildChildEntityString(sChildPropertyName)
  '  nUpdate = enParent.mContainer.UpdateDB(enParent)
  '  If nUpdate = 1 Then
  '    gRemoveChildEntity = nUpdate
  '    Exit Function
  '  End If
  '  Call enParent.ChildEntities(sChildPropertyName).Add(enChld)  'revert
  '  Call enParent.BuildChildEntityString(enChld)
  '  MsgBox("This record has been updated since you last refreshed. Please refresh again and then update.")
  '  gRemoveChildEntity = nUpdate
  'End Function

  'global
  Public Function gReOrderChildEntities(ByVal enParent As Object, ByVal enChlds As Object, ByVal strChild As String) As Integer
    Dim strChildString As String
    strChildString = enParent.ChildEntityString(strChild)  'preserve in case needing to revert
    Call enParent.BuildChildEntityString(strChild)
    nUpdate = enParent.mContainer.UpdateDB(enParent)
    If nUpdate = 0 Then  'revert back the entity string, the childentities need not be reverted
      Call enParent.BuildChildEntityString(strChildString)
    End If
    gReOrderChildEntities = nUpdate
  End Function
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  'UI
  Function gCheckAllParentsForChildPresence(ByVal objParents As Object, ByVal objChild As Object, ByVal sChildPropertyName As String) As Boolean
    'checks if a Child entity is still available in a 1-M relationship.
    Dim m As Object
    Dim nPresent As Integer
    nPresent = 0
    gCheckAllParentsForChildPresence = False
    For Each m In objParents
      If gCheckThisParentForChildPresence(m, objChild, sChildPropertyName) Then
        nPresent = nPresent + 1
      End If
    Next m
    If nPresent > 0 Then
      gCheckAllParentsForChildPresence = True
    Else
      gCheckAllParentsForChildPresence = False
    End If
  End Function

  'UI
  Function gCheckThisParentForChildPresence(ByVal objParent As Object, ByVal objChild As Object, ByVal sChildPropertyName As String) As Boolean
    gCheckThisParentForChildPresence = False
    Dim sp As Object
    For Each sp In objParent.ChildEntities(sChildPropertyName)
      If objChild.ID = sp.ID Then
        gCheckThisParentForChildPresence = True
        Exit Function
      End If
    Next sp
    gCheckThisParentForChildPresence = False
    'End Select
  End Function

  Public Function gBuildChildEntityObjects(ByVal objEnt As Object, ByVal strChld As String, ByVal strEnt As String) As Object    'ok
    'objEnt: parent object
    'strChld: Name of the child entities object
    'strEnt: string id's to be converted into objects

    Dim n As Integer
    Dim i As Integer
    Dim ent As Object
    'Dim ents As Object
    'Dim sAssoc As String
    Dim sJoinTable As String
    Dim sArray() As String
    Dim mObjectForJoinTable As Object
    objEnt.ChildEntities(strChld).removeall()
    sJoinTable = cEntityDataItems.GetJoinTable(TypeName(objEnt), strChld, "")
    If sJoinTable <> "CSR" Then
      Call cEntityDataItems.GetAssociation(sJoinTable, TypeName(objEnt), strParentPropertyName, "")
      mObjectForJoinTable = CreateObjectFromString(sJoinTable)
      Dim obj As Object
      For Each obj In mObjectForJoinTable.mContainer.Items
        If objEnt.ID = obj.ChildEntityString(strParentPropertyName) Then
          ent = MGlobals.CreateObjectFromString(strChld)
          'Set ent = ent.mContainer.Item(strEnt)
          ent = ent.mContainer.item(obj.ChildEntityString(strChld))
          Call objEnt.ChildEntities(strChld).Add(ent)
        End If
      Next obj
      Exit Function
    End If

    sArray = Split(strEnt, ";")
    n = UBound(sArray)
    For i = 0 To n  'removed ';' from both ends.
      If (strEnt = "") Then Exit For
      'Set ent = objEnt.ChildEntities(strChld).mContained.mContainer.Item(sArray(i))
      'Call objEnt.ChildEntities(strChld).Add(ent)
      'hopefully this should work , objchld is a string
      ent = MGlobals.CreateObjectFromString(strChld)
      ent = ent.mContainer.item(sArray(i))
      ' ent=ent.GetType.
      objEnt.ChildEntityString(strChld, strEnt)
      Call objEnt.ChildEntities(strChld).Add(ent)
    Next i
    gBuildChildEntityObjects = objEnt.ChildEntities(strChld)
  End Function

  Public Function gBuildChildEntityString(ByVal entPar As Object, ByVal enChlds As String) As String
    Dim en As Object
    Dim sSp As String
    sSp = ""
    For Each en In entPar.ChildEntities(enChlds)
      sSp = sSp & en.ID & ";"
    Next en
    If Len(sSp) > 0 Then sSp = Mid(sSp, 1, Len(sSp) - 1)
    entPar.ChildEntityString(enChlds, sSp)
  End Function

  Function GetItemForJoinTable(ByVal enJT As String, ByVal enPar As Object, ByVal enChld As Object, ByVal strChldPropertyName As String) As Object
    'enJT is the join table
    'enPar: parent entity
    'enChld: Child Entity
    Call cEntityDataItems.GetAssociation(enJT, TypeName(enPar), strParentPropertyName, "")
    Dim oJT As Object
    oJT = MGlobals.CreateObjectFromString(enJT)
    Dim eP As Object
    For Each eP In oJT.mContainer
      If eP.ChildEntityString(strParentPropertyName) = enPar.ID And eP.ChildEntityString(strChldPropertyName) = enChld.ID Then
        GetItemForJoinTable = eP
        Exit Function
      End If
    Next eP
    GetItemForJoinTable = Nothing
  End Function


  'Public Sub gRecordChanges(ByVal rstField As String, ByVal objVar As Object, ByVal abbr As String)
  '  If CStr("" & rst.Fields(rstField).Value) <> CStr(objVar) Then
  '    rst(rstField).Value = objVar
  '    gstrChanges = gstrChanges & abbr & CStr(objVar) & ","
  '    gbChangesMade = True
  '  Else
  '    gstrChanges = gstrChanges & ","
  '  End If
  'End Sub

  '  Sub gAddChanges(ByVal rstField As String, ByVal objVar As Object, ByVal abbr As String)
  '      rst(rstField).Value = objVar
  '      gstrChanges = gstrChanges & abbr & CStr(objVar) & ","
  '  End Sub

  Public Sub gRecordChanges1(ByVal sOperation As String, ByVal rstField As String, ByVal objVar As Object, ByVal abbr As String)
    Select Case sOperation
      Case "Update"
        If CStr("" & rst.Fields(rstField).Value) <> CStr(objVar) Then
          rst(rstField).Value = objVar
          gstrChanges = gstrChanges & abbr & CStr(objVar) & ","
          gbChangesMade = True
        Else
          gstrChanges = gstrChanges & ","
        End If
      Case "Add"
        rst(rstField).Value = objVar
        gstrChanges = gstrChanges & abbr & CStr(objVar) & ","
    End Select
  End Sub

End Module
