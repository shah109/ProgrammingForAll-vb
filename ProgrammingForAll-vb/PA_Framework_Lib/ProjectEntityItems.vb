Option Explicit On

Public Class ProjectEntityItems
  Inherits System.Collections.CollectionBase
  Dim nDecLoadOrder As Integer  'comment
  'Private collInt As Collection
  Public mContained As Object
  Private _sNoHashTable As New Hashtable

  Dim strChanges As String
  Public Event Changed()
  'Dim strSqlCall As String
  'Dim strConnFramework As String

  'Public Sub New()
  '  'collInt = New Collection
  '  mContained = New ProjectEntityItem
  'End Sub

  Public Sub Add(ByRef val As ProjectEntityItem)
    Me.List.Add(val)
    _sNoHashTable.Add(val.ID, val)
  End Sub

  Public ReadOnly Property item(ByVal index As Integer) As ProjectEntityItem
    Get
      Return Me.List.Item(index)
    End Get
  End Property

  Public Property Item(ByVal sid As String) As ProjectEntityItem
    Get
      Return _sNoHashTable.Item(sid)
    End Get
    Set(ByVal value As ProjectEntityItem)
      _sNoHashTable.Item(sid) = value
    End Set
  End Property

  Public Sub Remove(ByRef val As ProjectEntityItem)
    Me.List.Remove(val.ID)
  End Sub

  Public Function LoadSingleEntity(ByVal sid As String, ByVal sOperation As String) As Integer
    Dim ePrjEntItem_ As ProjectEntityItem
    DBStrings.gConnectionString = MGlobals.GetAppConnString
    gStrSqlCall = GetSQLQuery("LoadSingleEntity", sid)

    cnn.Open(MGlobals.GetAppConnString)
    rst.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    rst.Open(gStrSqlCall, cnn, _
              ADODB.CursorTypeEnum.adOpenStatic, 1)
    If rst.RecordCount <> 1 Then
      LoadSingleEntity = 0  '
      rst.Close()
      cnn.Close()
      Exit Function
    End If
    If sOperation = "N" Then
      ePrjEntItem_ = New ProjectEntityItem
    ElseIf sOperation = "U" Then
      ePrjEntItem_ = Me.item(sid)
    End If
    Call Me.LoadEntityItemsForThisEntity(ePrjEntItem_)

    ePrjEntItem_.Lastupdate = rst.Fields("LastUpdate").Value
    'ePrjEntItem_.Container = Me

    If sOperation = "N" Then
      ePrjEntItem_.LoadOrder = Me.Count - nDecLoadOrder
      Call Me.Add(ePrjEntItem_)
    End If
    rst.Close()
    cnn.Close()
    LoadSingleEntity = 1
  End Function

  Public Function Load()
    Dim ePrjEntItem_ As ProjectEntityItem
    Dim nLoadorder As Integer
    gStrSqlCall = GetSQLQuery("Load", "ID")
    
    cnn.Open(MGlobals.GetAppConnString)
    rst.CursorLocation = adUseClient
    rst.Open(gStrSqlCall, cnn, ADODB.CursorTypeEnum.adOpenStatic, 1)

    Me.Clear()
    rst.MoveFirst()
    nDecLoadOrder = 0
    'if this entity has a M-1 rel with any other entity, then a blank entity must be present for cbos.
    If cEntityDataItems.ParentHasM1Relationship("ProjectEntityItem") Then
      ePrjEntItem_ = New ProjectEntityItem   'this member is loaded for 0 index entries cbo's for M-1 relationships
      ePrjEntItem_.LoadOrder = 0
      ePrjEntItem_.ID = 0
      nDecLoadOrder = 1
      Call Me.Add(ePrjEntItem_)
    End If
    nLoadorder = 1
    Do While Not rst.EOF
      ePrjEntItem_ = New ProjectEntityItem

      Call Me.LoadEntityItemsForThisEntity(ePrjEntItem_)

      ePrjEntItem_.Lastupdate = rst.Fields("LastUpdate").Value
      ePrjEntItem_.LoadOrder = nLoadorder
      'ePrjEntItem_.mContainer = Me
      Call Me.Add(ePrjEntItem_)
      nLoadorder = nLoadorder + 1
      rst.MoveNext()
    Loop
    rst.Close()
    cnn.Close()
    'Build the domain model now
    'Call Me.BuildDomainModel
  End Function

  Sub BuildDomainModel()
    '  Dim eb_ As ProjectEntityItem
    '  Dim e1_ As Entity1
    '  'now add e2_s to e1_s
    '  For Each e1_ In cEntity1s.Items  'assumed cEntity1s have been loaded
    '    Set e1_.ChildEntity2s = e1_.BuildChildEntityObjects(e1_.ChildEntity2s, e1_.ChildEntityString(currePrjEntItem_))
    '    'For Each ePrjEntItem_ In e1_.ChildEntity2s.Items
    '    ' If Not ePrjEntItem_.ContainsParentEntity(e1_) Then  'BuildObjectModel is called from AddtoDB() too. hence to check for already existing.
    '    '  Call ePrjEntItem_.ParentEntity1s.Add(e1_)   'add parents to ProjectEntityItem
    '    'End If
    '    'Next ePrjEntItem_
    '  Next e1_
  End Sub

  Public Function UpdateDB(ByRef ent As ProjectEntityItem) As Integer
    'Updates the database table for the entity'.
    'Return 0 if the db record has been updated by other user.
    'Return the loadorder if successfull in update.
    Dim nCount As Integer
    gbChangesMade = False
    gStrSqlCall = GetSQLQuery("UpdateDB", ent.ID)
    If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
     cnn.Open(MGlobals.GetAppConnString)
    End If
    ' cnn.Open(strConnFramework)
    rst.Open(Source:=gStrSqlCall, ActiveConnection:=cnn, CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
    'DBStrings.cmd1.CommandText = gStrSqlCall
    'DBStrings.cmd1.CommandType = CommandType.Text
    'DBStrings.cmd1.Connection = DBStrings.cnn1
    'DBStrings.cnn1.Open()
    'DBStrings.dr1 = DBStrings.cmd1.ExecuteReader

    nCount = rst.RecordCount
    If nCount <> 1 Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    End If
    If rst.Fields("LastUpdate").Value <> ent.Lastupdate Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    Else
      rst.Fields("LastUpdate").Value = Now
      ent.Lastupdate = rst.Fields("LastUpdate").Value  'update the local last update so you can mame changes without refreshing.
    End If
    strChanges = ""
    ''PH:For Each Entity Item. Add to record changes int the DB (function UpdateDB()
    'Call BLFunctions.gRecordChanges(ePrjEntItem_, "sEntityItem_1", ePrjEntItem_.EntityItem_1, "fn:")
    ''PH: End
    'sPH_col_Update:
    Call Recordchanges("Update", ent)
    'Call BLFunctions.gRecordChanges("sEntityItem_1", ent.ItemChildName, "ei1:")
    'Call BLFunctions.gRecordChanges("sEntityItem_2", ent.ItemChildRelationship, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemDBName", ent.ItemDBName, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemDBNameType", ent.ItemDBNameType, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemChildName", ent.ItemChildName, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemChildRelationship", ent.ItemChildRelationship, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemRelationshipType", ent.ItemRelationshipType, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemSQLName", ent.ItemSQLName, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemInternalName", ent.ItemInternalName, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemInternalNameType", ent.ItemInternalNameType, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemPropertyName", ent.ItemPropertyName, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemPropertyNameType", ent.ItemPropertyNameType, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemULName", ent.ItemULName, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemSheetDisplayOrder", ent.ItemSheetDisplayOrder, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemFormControlName", ent.ItemFormControlName, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemChildListSetNeeded", ent.ItemChildListSetNeeded, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemParentListSetNeeded", ent.ItemParentListSetNeeded, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemControlReference", ent.ItemControlReference, "ei2:")
    'Call BLFunctions.gRecordChanges("sItemGenerateFlag", ent.ItemGenerateFlag, "ei2:")

    'sPH_col_Update:End

    ''PH: For Each Child Entity, add this line for recording the changes for child entities for this object.
    ' Call ePrjEntItem_.BuildChildEntityString(ePrjEntItem_.ChildEntities(e3_Par))
    ' Call BLFunctions.gRecordChanges(ePrjEntItem_, "sEntity3s", ePrjEntItem_.ChildEntityString(e3_Par), "e3_str:")
    ''PH: End
    'sPH_col_UpdateChild:
    'sPH_col_UpdateChild
    'sPH_col_UpdateChild:End

    rst.Update()
    rst.Close()
    MGlobals.currePrjEntItm_ = ent

    'No need to enter a record in update log if no change has been made
    If gbChangesMade = True Then
      AppSettings.LastUpdateID = gUpdateLogTable("ePrjEntItem_", "U", ent.ID)
    End If
    cnn.Close()
    Call DoChanged()
    UpdateDB = nCount
  End Function

  Public Function AddtoDB(ByRef ent As ProjectEntityItem) As Integer
    Dim strSQL As String
    Dim nCount As Integer
    strSQL = "select ID from ProjectEntityItems"
    If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      cnn.Open(MGlobals.GetAppConnString)
    End If
    rst.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
    nCount = rst.RecordCount
    If cEntityDataItems.ParentHasM1Relationship("ProjectEntityItem") Then nCount = nCount + 1
    If nCount <> cProjectEntityItems.Count - 1 Then  'other user has added a record
      rst.Close()
      cnn.Close()
      AddtoDB = 0
      Exit Function
    End If
    rst.Close()
    strSQL = "select max(ID) as MaxNo from ProjectEntityItems"
    Dim nMax As Integer
    rst.CursorLocation = adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, adLockOptimistic)
    nMax = rst.Fields("MaxNo").Value
    rst.Close()
    ent.ID = nMax + 1
    'now add the record
    rst.CursorType = adOpenKeySet
    rst.LockType = adLockOptimistic
    rst.Open("ProjectEntityItems", cnn, , , adCmdTable)
    rst.AddNew()

    ''PH:For Each Entity Item. Add the items to the db (function AddtoDB)
    'Call gAddChanges("EntityItem_2", ent.EntityItem_2, "ln:")
    ''PH: End
    'sPH_col_Add:
    Call gRecordChanges1("Add", "ID", ent.ID, "id:")
    Call Recordchanges("Add", ent)
    ' Call gAddChanges("EntityItem_2", ent.ItemChildRelationship, "ei2:")
    'sPH_col_Add:End

    ''PH: For Each Child Entity, add this line for saving the child string with e3_ is the child.
    ' Call gAddChanges("Entity_Cs", ent.ChildEntityString(e3_Par), "e2_str:")
    ''PH: End
    'sPH_col_AddChild:
    ''''sPH_col_AddChild

    'sPH_col_AddChild:End

    rst("LastUpdate").Value = Now
    ent.LastUpdate = rst("LastUpdate").Value
    'ent.mContainer = Me
    rst.Update()
    rst.Close()

    ent.LoadOrder = Me.Count - nDecLoadOrder
    'Call Me.Add(ent)

    Call Me.BuildDomainModel()
    MGlobals.currePrjEntItm_ = ent

    'now update the log
    AppSettings.LastUpdateID = gUpdateLogTable("ePrjEntItem_", "N", ent.ID)
    cnn.Close()
    Call DoChanged()
    AddtoDB = nCount + 1
  End Function

  Public Function DeleteFromDB(ByVal ent As ProjectEntityItem) As Integer
    Dim lAffected As Long

    gStrSqlCall = " Delete " & _
                  "FROM " & strDBO & "ProjectEntityItems " & _
                  " where " & strDBO & "ID LIKE '" & ent.ID & "'"
    cmd.ActiveConnection = cnn
    'Globals.gADOcmd.ActiveConnection = gStrConnGenEBA
    cmd.CommandText = gStrSqlCall
    cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
    cmd.Execute(lAffected, ADODB.ExecuteOptionEnum.adExecuteNoRecords)
    If lAffected = 1 Then
      Dim sEbID As String
      sEbID = ent.ID
      MGlobals.gstrChanges = "Deleted: " & "ID"
      Call cProjectEntityItems.Remove(ent)
      cnn.Open(gStrConnGenEBA)
      AppSettings.LastUpdateID = gUpdateLogTable("ePrjEntItem_", "D", sEbID)
      cnn.Close()
      Call DoChanged()
    End If
    cmd = Nothing
    DeleteFromDB = lAffected
  End Function

  Function GetSQLQuery(ByVal strFunc As String, ByVal strID As String) As String
    'This SQL query is used in 3 functions: LoadSingleEntity(). Load() and UpdateDB()
    Dim strTemp As String
    'PH: For Each Entity Item, add to the Select
    ' strDBO & "EntityItem_2 as sEntityItem_2, " & _
    'PH: End
    'sPH_col_Sql:
    strTemp = " Select " & _
              strDBO & "ID, " & _
              strDBO & "EntityItem_1, " & _
              strDBO & "EntityItem_2, " & _
              strDBO & "ItemDBName, " & _
              strDBO & "ItemDBNameType, " & _
              strDBO & "ItemChildName, " & _
              strDBO & "ItemChildRelationship, " & _
              strDBO & "ItemRelationshipType, " & _
              strDBO & "ItemSQLName, " & _
              strDBO & "ItemInternalName, " & _
              strDBO & "ItemInternalNameType, " & _
              strDBO & "ItemPropertyName, " & _
              strDBO & "ItemPropertyNameType, " & _
              strDBO & "ItemULName, " & _
              strDBO & "ItemSheetDisplayOrder, " & _
              strDBO & "ItemFormControlName, " & _
              strDBO & "ItemChildListSetNeeded, " & _
              strDBO & "ItemParentListSetNeeded, " & _
              strDBO & "ItemControlReference, " & _
              strDBO & "ItemGenerateFlag, " & _
              strDBO & "LastUpdate " & _
              "FROM " & strDBO & "ProjectEntityItems "
    'sPH_col_Sql:End

    If strFunc = "LoadSingleEntity" Or strFunc = "UpdateDB" Then
      strTemp = strTemp & " WHERE " & strDBO & "ID like " & strID
    ElseIf strFunc = "Load" Then
      strTemp = strTemp & " ORDER BY " & strDBO & strID
    End If
    GetSQLQuery = strTemp
  End Function

  Function LoadEntityItemsForThisEntity(ByVal ent As ProjectEntityItem)
    ''PH:For Each Entity Item. Add it for loading
    'eA_.EntityItem_1 = "" & rst.Fields("sEntityItem_1")
    ''PH End
    'sPH_col_Load:
    ent.ID = "" & rst.Fields("ID").Value
    ent.ItemDBName = "" & rst.Fields("ItemDBName").Value
    ent.ItemDBNameType = "" & rst.Fields("ItemDBNameType").Value
    ent.ItemChildName = "" & rst.Fields("ItemChildName").Value
    ent.ItemChildRelationship = "" & rst.Fields("ItemChildRelationship").Value
    ent.ItemRelationshipType = "" & rst.Fields("ItemRelationshipType").Value
    ent.ItemSQLName = "" & rst.Fields("ItemSQLName").Value
    ent.ItemInternalName = "" & rst.Fields("ItemInternalName").Value
    ent.ItemInternalNameType = "" & rst.Fields("ItemInternalNameType").Value
    ent.ItemPropertyName = "" & rst.Fields("ItemPropertyName").Value
    ent.ItemPropertyNameType = "" & rst.Fields("ItemPropertyNameType").Value
    ent.ItemULName = "" & rst.Fields("ItemULName").Value
    ent.ItemSheetDisplayOrder = "" & rst.Fields("ItemSheetDisplayOrder").Value
    ent.ItemFormControlName = "" & rst.Fields("ItemFormControlName").Value
    ent.ItemChildListSetNeeded = "" & rst.Fields("ItemChildListSetNeeded").Value
    ent.ItemParentListSetNeeded = "" & rst.Fields("ItemParentListSetNeeded").Value
    ent.ItemControlReference = "" & rst.Fields("ItemControlReference").Value
    ent.ItemGenerateFlag = "" & rst.Fields("ItemGenerateFlag").Value



    'ent.ID = "" & rst.Fields("sID").Value
    'ent.EntityItem_1 = "" & rst.Fields("sEntityItem_1").Value
    'ent.EntityItem_2 = "" & rst.Fields("sEntityItem_2").Value
    'sPH_col_Load:End

    ''PH: For Each Child Entity, add the two lines for building child entities for this object.
    ' Call ent.BuildChildEntityObjects(e3_Par, "" & rst.Fields("sEntity3s"))
    ''PH: End
    ' If cEntityDataItems.IsJoinTable(TypeName(ent)) = True Then
    'sPH_col_LoadChildEntitiesJT:
    'ent.ChildEntityString("JTOrders") = "" & rst.Fields("sOrders")
    'sPH_col_LoadChildEntitiesJT:End
    'Else
    'sPH_col_LoadChildEntitiesCS:
    ' Call ent.BuildChildEntityObjects("JTOrders", "" & rst.Fields("sOrders"))
    'sPH_col_LoadChildEntitiesCS:End
    'End If
  End Function

  Public Sub RaiseChanged()
    RaiseEvent Changed()
  End Sub

  Sub Recordchanges(ByVal sOperation As String, ByVal ent As ProjectEntityItem)
    Call BLFunctions.gRecordChanges1(sOperation, "EntityItem_1", ent.ItemChildName, "ei1:")
    Call BLFunctions.gRecordChanges1(sOperation, "EntityItem_2", ent.ItemChildRelationship, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemDBName", ent.ItemDBName, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemDBNameType", ent.ItemDBNameType, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemChildName", ent.ItemChildName, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemChildRelationship", ent.ItemChildRelationship, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemRelationshipType", ent.ItemRelationshipType, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemSQLName", ent.ItemSQLName, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemInternalName", ent.ItemInternalName, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemInternalNameType", ent.ItemInternalNameType, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemPropertyName", ent.ItemPropertyName, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemPropertyNameType", ent.ItemPropertyNameType, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemULName", ent.ItemULName, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemSheetDisplayOrder", ent.ItemSheetDisplayOrder, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemFormControlName", ent.ItemFormControlName, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemChildListSetNeeded", ent.ItemChildListSetNeeded, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemParentListSetNeeded", ent.ItemParentListSetNeeded, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemControlReference", ent.ItemControlReference, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "ItemGenerateFlag", ent.ItemGenerateFlag, "ei2:")

  End Sub

End Class
