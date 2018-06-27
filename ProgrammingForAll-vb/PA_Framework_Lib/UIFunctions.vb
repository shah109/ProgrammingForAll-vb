
Public Class UIFunctions
  Private Shared bUpdate As Boolean

  Public Shared Sub FillParentEntities(ByRef objEDI As Object, ByRef entParent As Object, ByRef entChild As Object, ByVal sChildPropertyName As String)
    'Fills in the Parent entities for an entity list box. 

    Dim sp As Object
    Try
      Call entChild.ParentEntities(entParent).RemoveAll()
    Catch
      MsgBox("Error from FillParentEntities(): Child:" & TypeName(entChild) & "; Parent:" & TypeName(entParent) & "; ChildEntityName:" & sChildPropertyName)
      PASettings_Lib.WriteToErrorLog("Error from FillParentEntities(): Child:" & TypeName(entChild) & "; Parent:" & TypeName(entParent) & "; ChildEntityName:" & sChildPropertyName)
      Exit Sub
    End Try
    Dim bFlag As Boolean
    For Each sp In entParent.mContainer
      bFlag = CheckThisParentForChildPresence(sp, entChild, sChildPropertyName)
      If bFlag = True Then
        Call entChild.ParentEntities(entParent).Add(sp)
      End If
    Next sp
  End Sub

  Public Shared Sub FillAvailableParentEntities(ByRef objEDI As Object, ByRef entParent As Object, ByRef entChild As Object, ByVal sChildPropertyName As String)
    'Fills in the available Parent entities in the available Parent entities list box.
    Dim bParentEntity_Available As Boolean
    Dim nParentCount As Integer
    Dim sJoinTable As String = String.Empty
    Try
      Call entChild.AvailableParentEntities(entParent).RemoveAll()
    Catch ex As Exception
      MsgBox(ex.Message & vbCrLf & "Error from FillAvailableParentEntities(): Child:" & TypeName(entChild) & "; Parent:" & TypeName(entParent) & "; ChildEntityName:" & sChildPropertyName)
      PASettings_Lib.WriteToErrorLog("Error from FillAvailableParentEntities(): Child:" & TypeName(entChild) & "; Parent:" & TypeName(entParent) & "; ChildEntityName:" & sChildPropertyName)
      Exit Sub
    End Try
    Dim sp As Object

    bParentEntity_Available = False
    gStrAssoc = objEDI.GetAssociation(TypeName(entParent), TypeName(entChild), sChildPropertyName, sJoinTable)

    nParentCount = entChild.ParentEntities(entParent).Count
    If (gStrAssoc = "1-1") And (nParentCount <> 0) Then Exit Sub ' only one parent entry because it a 1-1 relationship.

    If (gStrAssoc = "1-M") Then
      bParentEntity_Available = Not CheckAllParentsForChildPresence(entParent.mContainer, entChild, sChildPropertyName)
    End If
    For Each sp In entParent.mContainer
      If Trim(sp.ID) = "" Then GoTo nextloop1

      If gStrAssoc = "M-M" Then
        bParentEntity_Available = Not CheckThisParentForChildPresence(sp, entChild, sChildPropertyName)
      End If
      If (gStrAssoc = "M-1") Or (gStrAssoc = "1-1") Then
        bParentEntity_Available = Not CBool(sp.ChildEntities(sChildPropertyName).Count)   'The should be no child entities for the parent to show in the list
      End If
      If bParentEntity_Available = True Then
        Call entChild.AvailableParentEntities(entParent).Add(sp)

      End If
nextloop1:
    Next sp
  End Sub

  Public Shared Sub FillAvailableParentEntitiesForCBO(ByRef objEDI As Object, ByRef entParent As Object, ByRef entChild As Object, ByVal sChildPropertyName As String)
    'Fills in the available Parent entities in the available Parent entities list box.
    Dim bParentEntity_Available As Boolean
    Dim nParentCount As Integer
    Dim sJoinTable As String = String.Empty
    Try
      Call entChild.AvailableParentEntities(entParent).RemoveAll()
    Catch ex As Exception
      MsgBox(ex.Message & vbCrLf & "Error from FillAvailableParentEntities(): Child:" & TypeName(entChild) & "; Parent:" & TypeName(entParent) & "; ChildEntityName:" & sChildPropertyName)
      PASettings_Lib.WriteToErrorLog("Error from FillAvailableParentEntities(): Child:" & TypeName(entChild) & "; Parent:" & TypeName(entParent) & "; ChildEntityName:" & sChildPropertyName)
      Exit Sub
    End Try
    Dim sp As Object

    bParentEntity_Available = False
    gStrAssoc = objEDI.GetAssociation(TypeName(entParent), TypeName(entChild), sChildPropertyName, sJoinTable)

    nParentCount = entChild.ParentEntities(entParent).Count
    'If (gStrAssoc = "1-1") And (nParentCount <> 0) Then Exit Sub ' only one parent entry because it a 1-1 relationship.

    'If (gStrAssoc = "1-M") Then
    '  bParentEntity_Available = Not CheckAllParentsForChildPresence(entParent.mContainer, entChild, sChildPropertyName)
    'End If
    For Each sp In entParent.mContainer
      If Trim(sp.ID) = "" Then GoTo nextloop1

      If (gStrAssoc = "M-M") Or (gStrAssoc = "1-M") Then
        bParentEntity_Available = Not CheckThisParentForChildPresence(sp, entChild, sChildPropertyName)
      End If
      If (gStrAssoc = "M-1") Or (gStrAssoc = "1-1") Then
        bParentEntity_Available = Not CBool(sp.ChildEntities(sChildPropertyName).Count)   'The should be no child entities for the parent to show in the list
      End If
      If bParentEntity_Available = True Then
        Call entChild.AvailableParentEntities(entParent).Add(sp)

      End If
nextloop1:
    Next sp
  End Sub

  Public Shared Function FillChildEntities(ByRef objEDI As Object, ByRef entPr As Object, ByRef entChild As Object, ByVal sChildPropertyName As String) As Boolean
    'Fills in the Child entities in the Child entity list box
    Dim sJoinTable As String = String.Empty
    FillChildEntities = False
    Dim mObjectForJoinTable As Object
    Dim strParentPropertyName As String = ""
    Dim ent As Object
    Dim strAssoc As String
    strAssoc = objEDI.GetAssociation(TypeName(entPr), TypeName(entChild), sChildPropertyName, sJoinTable)
    If strAssoc = "M-M" And sJoinTable <> "CSR" Then
      entPr.ChildEntities(sChildPropertyName).RemoveAll()
      strParentPropertyName = objEDI.GetChildPropertyName(sJoinTable, TypeName(entPr))
      mObjectForJoinTable = objEDI.CreateObjectFromString(sJoinTable)
      Dim obj As Object
      For Each obj In mObjectForJoinTable.mContainer
        If entPr.ID = obj.ChildEntityString(strParentPropertyName) Then
          ent = objEDI.CreateObjectFromString(sChildPropertyName)
          ent = ent.mContainer.item(obj.ChildEntityString(sChildPropertyName))
          'Set ent = ent.mContainer.Item(strEnt)
          Call entPr.ChildEntities(sChildPropertyName).Add(ent)
        End If
      Next obj
      FillChildEntities = True
    End If

  End Function

  Public Shared Sub FillAvailableChildEntities(ByRef objEntityDataItems As Object, ByRef entChldContainer As Object, ByRef entPr As Object, ByRef entChild As Object, ByVal sChildPropertyName As String)
    'Fills in the available Child entities in the available Child entities list box. 
    'objEntityDataItems: Only used for getting the association.
    'entChldContainer: The container of child objects through which iteration is done whether a child has to be added to a parent's available children
    'entPr: Parent to which available children are added.
    'entChild: Child objects to be added to the parent's available child objects.
    'sChildPropertyName: Property name of the child through which association of the parent with child is determined (whether 1-1, 1-M or M-M)

    Dim sJoinTable As String = String.Empty
    Dim bChildEntity_Available As Boolean 'var that shows if a child entity is available to be added to the parent or not.
    Try
      Call entPr.AvailableChildEntities(entChild).RemoveAll() 'initially, remove all child entities from parent; each entity will be added to the parent after checks.
    Catch
      MsgBox("Error from FillAvailableChildEntities(): Parent:" & TypeName(entPr) & "; Child:" & TypeName(entChild) & "; ChildEntityName:" & sChildPropertyName)
      PASettings_Lib.WriteToErrorLog("Error from FillAvailableChildEntities(): Parent:" & TypeName(entPr) & "; Child:" & TypeName(entChild) & "; ChildEntityName:" & sChildPropertyName)
      Exit Sub
    End Try
    Dim sp As Object
    Dim nChildCount As Integer

    gStrAssoc = objEntityDataItems.GetAssociation(TypeName(entPr), TypeName(entChild), sChildPropertyName, sJoinTable)
    nChildCount = entPr.ChildEntities(sChildPropertyName).Count
    If (gStrAssoc = "M-1") And (nChildCount <> 0) Then Exit Sub ' only one child entry allowed in parent. Hence the available chld entities should not be there to be added
    If (gStrAssoc = "1-1") And (nChildCount <> 0) Then Exit Sub ' only one child entry allowed.

    For Each sp In entChldContainer 'iterate through the child container, a parameter passed on to this function. May be a subset of all children loaded not the complete c'Objects'
      If Trim(sp.ID) = "" Then GoTo nextloopchild
      Select Case gStrAssoc
        Case "1-1", "1-M" 'then check all parents to ensure none of the parents contains this child.
          bChildEntity_Available = Not CheckAllParentsForChildPresence(entPr.mContainer, sp, sChildPropertyName)
        Case "M-M", "M-1" 'M-1 has to be here otherwise bChildentityavailable is always false for m-1
          bChildEntity_Available = Not CheckThisParentForChildPresence(entPr, sp, sChildPropertyName)
      End Select
      If bChildEntity_Available = True Then
        Call entPr.AvailableChildEntities(entChild).add(sp)
      End If

nextloopchild:
    Next sp
  End Sub

  Public Shared Sub FillAvailableChildEntitiesForCBO(ByRef objEntityDataItems As Object, ByRef entPr As Object, ByRef entChild As Object, ByVal sChildPropertyName As String)
    'Fills in the available Child entities in the available Child entities combox box. Very similar to list box excpet the two lines are commentedComment/uncomment as shown for 1-m and m-m relationships
    Dim sJoinTable As String = String.Empty
    Dim bChildEntity_Available As Boolean
    Try
      Call entPr.AvailableChildEntities(entChild).RemoveAll()
    Catch
      MsgBox("Error from FillAvailableChildEntities(): Parent:" & TypeName(entPr) & "; Child:" & TypeName(entChild) & "; ChildEntityName:" & sChildPropertyName)
      PASettings_Lib.WriteToErrorLog("Error from FillAvailableChildEntities(): Parent:" & TypeName(entPr) & "; Child:" & TypeName(entChild) & "; ChildEntityName:" & sChildPropertyName)
      Exit Sub
    End Try
    Dim sp As Object
    Dim nChildCount As Integer

    gStrAssoc = objEntityDataItems.GetAssociation(TypeName(entPr), TypeName(entChild), sChildPropertyName, sJoinTable)
    nChildCount = entPr.ChildEntities(sChildPropertyName).Count
    'If (gStrAssoc = "M-1") And (nChildCount <> 0) Then Exit Sub ' only one child entry allowed. commented for cbo
    'If (gStrAssoc = "1-1") And (nChildCount <> 0) Then Exit Sub ' only one child entry allowed. commented for cbo
    For Each sp In entChild.mContainer
      If Trim(sp.ID) = "" Then GoTo nextloopchild
      Select Case gStrAssoc
        Case "1-1", "1-M"
          bChildEntity_Available = Not CheckAllParentsForChildPresence(entPr.mContainer, sp, sChildPropertyName)
        Case "M-M", "M-1" 'M-1 has to be here otherwise bChildentityavailable is always false for m-1
          bChildEntity_Available = Not CheckThisParentForChildPresence(entPr, sp, sChildPropertyName)
      End Select
      If bChildEntity_Available = True Then
        Call entPr.AvailableChildEntities(entChild).add(sp)
      End If
nextloopchild:
    Next sp
  End Sub

  Public Shared Function CheckAllParentsForChildPresence(ByRef objParents As Object, ByRef objChild As Object, ByVal sChildPropertyName As String) As Boolean
    'checks if a Child entity is still available in a 1-M relationship.
    Dim m As Object
    Dim nPresent As Integer
    nPresent = 0
    CheckAllParentsForChildPresence = False
    For Each m In objParents
      If CheckThisParentForChildPresence(m, objChild, sChildPropertyName) Then
        nPresent = nPresent + 1
      End If
    Next m
    If nPresent > 0 Then
      CheckAllParentsForChildPresence = True
    Else
      CheckAllParentsForChildPresence = False
    End If
  End Function

  Public Shared Function CheckThisParentForChildPresence(ByRef objParent As Object, ByRef objChild As Object, ByVal sChildPropertyName As String) As Boolean
    CheckThisParentForChildPresence = False
    Dim sp As Object
    For Each sp In objParent.ChildEntities(sChildPropertyName)
      If objChild.ID = sp.ID Then
        CheckThisParentForChildPresence = True
        Exit Function
      End If
    Next sp
    CheckThisParentForChildPresence = False
    'End Select
  End Function

  'to remove
  ' Public Shared Sub FillAvailableChildEntitiesOld(ByRef objEntityDataItems As Object, ByRef entPr As Object, ByRef entChild As Object, ByVal sChildPropertyName As String)
  '  'Fills in the available Child entities in the available Child entities list box. Comment/uncomment as shown for 1-m and m-m relationships
  '  Dim sJoinTable As String = String.Empty
  '  Dim bChildEntity_Available As Boolean
  '    Try
  '      Call entPr.AvailableChildEntities(entChild).RemoveAll()
  '    Catch
  '      MsgBox("Error from FillAvailableChildEntities(): Parent:" & TypeName(entPr) & "; Child:" & TypeName(entChild) & "; ChildEntityName:" & sChildPropertyName)
  '      PASettings_Lib.WriteToErrorLog("Error from FillAvailableChildEntities(): Parent:" & TypeName(entPr) & "; Child:" & TypeName(entChild) & "; ChildEntityName:" & sChildPropertyName)
  '      Exit Sub
  '    End Try
  '  Dim sp As Object
  '  Dim nChildCount As Integer

  '    gStrAssoc = objEntityDataItems.GetAssociation(TypeName(entPr), TypeName(entChild), sChildPropertyName, sJoinTable)
  '    nChildCount = entPr.ChildEntities(sChildPropertyName).Count
  '    If (gStrAssoc = "M-1") And (nChildCount <> 0) Then Exit Sub ' only one child entry allowed.
  '    If (gStrAssoc = "1-1") And (nChildCount <> 0) Then Exit Sub ' only one child entry allowed.

  '    For Each sp In entChild.mContainer
  '  'For Each sp In entPr.AvailableChildEntities(entChild)
  '      If Trim(sp.ID) = "" Then GoTo nextloopchild
  '      Select Case gStrAssoc
  '        Case "1-1", "1-M"
  '          bChildEntity_Available = Not CheckAllParentsForChildPresence(entPr.mContainer, sp, sChildPropertyName)
  '        Case "M-M", "M-1" 'M-1 has to be here otherwise bChildentityavailable is always false for m-1
  '          bChildEntity_Available = Not CheckThisParentForChildPresence(entPr, sp, sChildPropertyName)
  '      End Select
  '      If bChildEntity_Available = True Then
  '        Call entPr.AvailableChildEntities(entChild).add(sp)
  '      End If
  'nextloopchild:
  '    Next sp
  '  End Sub
End Class
