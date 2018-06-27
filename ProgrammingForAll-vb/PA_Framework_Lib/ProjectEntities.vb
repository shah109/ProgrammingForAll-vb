Option Explicit On

Public Class ProjectEntities
  Inherits System.Collections.CollectionBase
  Private _sNoHashTable As New Hashtable
  'sPH_col_DateTime:
  'gfh
  'sPH_col_DateTime:End
  'Projects
  Dim nDecLoadOrder As Integer  'comment
  'Private collInt As Collection
  Public mContained As Object
  Dim strChanges As String
  Public Event Changed()
  Dim strSqlCall As String
  Dim strConnFramework As String

  Public Sub Add(ByRef val As ProjectEntity)
    Me.List.Add(val)
    _sNoHashTable.Add(val.ID, val)
  End Sub

  Public ReadOnly Property item(ByVal index As Integer) As ProjectEntity
    Get
      Return Me.List.Item(index)
    End Get
  End Property

  Public Property item(ByVal sid As String) As ProjectEntity
    Get
      Return _sNoHashTable.Item(sid)
    End Get
    Set(ByVal value As ProjectEntity)
      _sNoHashTable.Item(sid) = value
    End Set
  End Property

  Public Sub Remove(ByRef val As ProjectEntity)
    Me.List.Remove(val)
  End Sub

  Public Function LoadSingleEntity(ByVal sId As String, ByVal sOperation As String) As Integer
    Dim ePrjtEnt_ As ProjectEntity
    strConnFramework = MGlobals.GetAppConnString
    strSqlCall = GetSQLQuery("LoadSingleEntity", sId)
    cnn.Open(strConnFramework)
    rst.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    rst.Open(strSqlCall, cnn, _
              ADODB.CursorTypeEnum.adOpenStatic, 1)
    If rst.RecordCount <> 1 Then
      LoadSingleEntity = 0  '
      rst.Close()
      cnn.Close()
      Exit Function
    End If
    If sOperation = "N" Then
      ePrjtEnt_ = New ProjectEntity
    ElseIf sOperation = "U" Then
      ePrjtEnt_ = Me.item(sId)
    End If
    Call Me.LoadEntityItemsForThisEntity(ePrjtEnt_)

    ePrjtEnt_.LastUpdate = rst.Fields("LastUpdate").Value

    If sOperation = "N" Then
      ePrjtEnt_.LoadOrder = Me.Count - nDecLoadOrder
      Call Me.Add(ePrjtEnt_)
    End If
    rst.Close()
    cnn.Close()
    LoadSingleEntity = 1
  End Function

  Public Function Load()
    Dim ePrjtEnt_ As ProjectEntity
    Dim nLoadorder As Integer
    'strConnFramework = MGlobals.GetAppConnString
    gStrSqlCall = GetSQLQuery("Load", "ID")

    cnn.Open(MGlobals.GetAppConnString)
    rst.CursorLocation = adUseClient
    rst.Open(gStrSqlCall, cnn, ADODB.CursorTypeEnum.adOpenStatic, 1)
    Me.Clear()
    rst.MoveFirst()
    nDecLoadOrder = 0
    'if this entity has a M-1 rel with any other entity, then a blank entity must be present for cbos.
    If cEntityDataItems.ParentHasM1Relationship("ProjectEntity") Then
      ePrjtEnt_ = New ProjectEntity   'this member is loaded for 0 index entries cbo's for M-1 relationships
      ePrjtEnt_.LoadOrder = 0
      ePrjtEnt_.ID = 0
      nDecLoadOrder = 1
      Call Me.Add(ePrjtEnt_)
    End If
    nLoadorder = 1
    Do While Not rst.EOF
      ePrjtEnt_ = New ProjectEntity

      Call Me.LoadEntityItemsForThisEntity(ePrjtEnt_)

      ePrjtEnt_.LastUpdate = rst.Fields("LastUpdate").Value
      ePrjtEnt_.LoadOrder = nLoadorder
      Call Me.Add(ePrjtEnt_)
      nLoadorder = nLoadorder + 1
      rst.MoveNext()
    Loop
    rst.Close()
    cnn.Close()
    'Build the domain model now
    'Call Me.BuildDomainModel
  End Function

  Sub BuildDomainModel()
    '  Dim eb_ As Project
    '  Dim e1_ As Entity1
    '  'now add ePrjt_s to e1_s
    '  For Each e1_ In cEntity1s.Items  'assumed cEntity1s have been loaded
    '    Set e1_.ChildProjects = e1_.BuildChildEntityObjects(e1_.ChildProjects, e1_.ChildEntityString(CurrePrjt_))
    '    'For Each ePrjt_ In e1_.ChildProjects.Items
    '    ' If Not ePrjt_.ContainsParentEntity(e1_) Then  'BuildObjectModel is called from AddtoDB() too. hence to check for already existing.
    '    '  Call ePrjt_.ParentEntity1s.Add(e1_)   'add parents to Project
    '    'End If
    '    'Next ePrjt_
    '  Next e1_
  End Sub

  Public Function UpdateDB(ByRef ent As ProjectEntity) As Integer
    'Updates the database table for the entity'.
    'Return 0 if the db record has been updated by other user.
    'Return the loadorder if successfull in update.
    Dim nCount As Integer
    gbChangesMade = False
    gStrSqlCall = GetSQLQuery("UpdateDB", ent.ID)
    If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      cnn.Open(MGlobals.GetAppConnString)
    End If
    'cnn.Open(strConnFramework)
    rst.Open(Source:=gStrSqlCall, ActiveConnection:=cnn, _
             CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
    nCount = rst.RecordCount
    If nCount <> 1 Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    End If
    If rst.Fields("LastUpdate").Value <> ent.LastUpdate Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    Else
      rst.Fields("LastUpdate").Value = Now
      ent.LastUpdate = rst.Fields("LastUpdate").Value  'update the local last update so you can mame changes without refreshing.
    End If
    strChanges = ""
    ''PH:For Each Entity Item. Add to record changes int the DB (function UpdateDB()
    'Call BLFunctions.gRecordChanges(ePrjt_, "sEntityItem_1", ePrjt_.EntityItem_1, "fn:")
    ''PH: End
    'sPH_col_Update:
    Call RecordChanges("Update", ent)
    'Call BLFunctions.gRecordChanges("EntityName", ent.EntityName, "en:")
    'Call BLFunctions.gRecordChanges("EntityCollectionName", ent.EntityCollectionName, "ecn:")
    'sPH_col_Update:End

    ''PH: For Each Child Entity, add this line for recording the changes for child entities for this object.
    ' Call ePrjt_.BuildChildEntityString(ePrjt_.ChildEntities(e3_Par))
    ' Call BLFunctions.gRecordChanges(ePrjt_, "sEntity3s", ePrjt_.ChildEntityString(e3_Par), "e3_str:")
    ''PH: End
    'sPH_col_UpdateChild:
    'sPH_col_UpdateChild
    'sPH_col_UpdateChild:End

    rst.Update()
    rst.Close()
    MGlobals.currePrjEnt_ = ent

    'No need to enter a record in update log if no change has been made
    If gbChangesMade = True Then
      AppSettings.LastUpdateID = gUpdateLogTable("ePrjt_", "U", ent.ID)
    End If
    cnn.Close()
    Call DoChanged()
    UpdateDB = nCount
  End Function

  Public Function AddtoDB(ByRef ent As ProjectEntity) As Integer
    Dim strSQL As String
    Dim nCount As Integer
    strSQL = "select Id from ProjectEntities"
    If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      cnn.Open(MGlobals.GetAppConnString)
    End If
    'rst.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
    nCount = rst.RecordCount
    If cEntityDataItems.ParentHasM1Relationship("Project") Then nCount = nCount + 1
    If nCount <> cProjectEntities.Count - 1 Then  'other user has added a record
      rst.Close()
      cnn.Close()
      AddtoDB = 0
      Exit Function
    End If
    rst.Close()
    strSQL = "select max(ID) as MaxNo from ProjectEntities"
    Dim nMax As Integer
    rst.CursorLocation = adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, adLockOptimistic)
    nMax = rst.Fields("MaxNo").Value
    rst.Close()
    ent.ID = nMax + 1
    'now add the record
    rst.CursorType = adOpenKeySet
    rst.LockType = adLockOptimistic
    rst.Open("ProjectEntities", cnn, , , adCmdTable)
    rst.AddNew()

    ''PH:For Each Entity Item. Add the items to the db (function AddtoDB)
    'Call gAddChanges("EntityItem_2", ePrjt_.EntityItem_2, "ln:")
    ''PH: End
    'sPH_col_Add:

    Call BLFunctions.gRecordChanges1("Add", "ID", ent.ID, "id:")
    Call RecordChanges("Add", ent)
    'Call gAddChanges("ID", ent.ID, "id:")
    'Call gAddChanges("EntityName", ent.EntityName, "en:")
    'Call gAddChanges("EntityCollectionName", ent.EntityShortName, "ecn:")
    'sPH_col_Add:End

    ''PH: For Each Child Entity, add this line for saving the child string with e3_ is the child.
    ' Call gAddChanges("Entity_Cs", ePrjt_.ChildEntityString(e3_Par), "ePrjt_str:")
    ''PH: End
    'sPH_col_AddChild:
    ''''sPH_col_AddChild

    'sPH_col_AddChild:End

    rst("LastUpdate").Value = Now
    ent.LastUpdate = rst("LastUpdate").Value
    rst.Update()
    rst.Close()

    ent.LoadOrder = Me.Count - nDecLoadOrder
    'Call Me.Add(ent)

    Call Me.BuildDomainModel()
    MGlobals.currePrjEnt_ = ent

    'now update the log
    AppSettings.LastUpdateID = gUpdateLogTable("ePrjt_", "N", ent.ID)
    cnn.Close()
    Call DoChanged()
    AddtoDB = nCount + 1
  End Function

  Public Function DeleteFromDB(ByVal ent As ProjectEntity) As Integer
    Dim lAffected As Long

    gStrSqlCall = " Delete " & _
                  "FROM " & strDBO & "ProjectEntities " & _
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
      Call cProjectEntities.Remove(ent)
      cnn.Open(gStrConnGenEBA)
      AppSettings.LastUpdateID = gUpdateLogTable("ePrjEnt_", "D", sEbID)
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
              strDBO & "EntityName, " & _
              strDBO & "EntityCollectionName, " & _
              strDBO & "EntityDBTableName, " & _
              strDBO & "EntityShortName, " & _
              strDBO & "EntityItems, " & _
              strDBO & "GenerateFlag, " & _
              strDBO & "InputFile, " & _
              strDBO & "FilesToGenerate, " & _
              strDBO & "DateTimeCodeGenerated, " & _
              strDBO & "LastUpdate " & _
              "FROM " & strDBO & "ProjectEntities "

    If strFunc = "LoadSingleEntity" Or strFunc = "UpdateDB" Then
      strTemp = strTemp & " WHERE " & strDBO & "ID like " & strID
    ElseIf strFunc = "Load" Then
      strTemp = strTemp & " ORDER BY " & strDBO & strID
    End If
    GetSQLQuery = strTemp
  End Function

  Function LoadEntityItemsForThisEntity(ByVal ent As ProjectEntity)
    ''PH:For Each Entity Item. Add it for loading
    'eA_.EntityItem_1 = "" & rst.Fields("sEntityItem_1")
    ''PH End
    'sPH_col_Load:
    ent.ID = "" & rst.Fields("ID").Value
    ent.EntityName = "" & rst.Fields("EntityName").Value
    ent.EntityCollectionName = "" & rst.Fields("EntityCollectionName").Value
    ent.EntityDBTableName = "" & rst.Fields("EntityDBTableName").Value
    ent.EntityShortName = "" & rst.Fields("EntityShortName").Value
    ent.EntityItems = "" & rst.Fields("EntityItems").Value
    ent.GenerateFlag = "" & rst.Fields("GenerateFlag").Value
    ent.InputFile = "" & rst.Fields("InputFile").Value
    ent.FilesToGenerate = "" & rst.Fields("FilesToGenerate").Value
    ent.DateTimeCodeGenerated = rst.Fields("DateTimeCodeGenerated").Value

    'sPH_col_Load:End

    ''PH: For Each Child Entity, add the two lines for building child entities for this object.
    ' Call ent.BuildChildEntityObjects(e3_Par, "" & rst.Fields("sEntity3s"))
    ''PH: End
    ' If cEntityDataItems.IsJoinTable(TypeName(ent)) = True Then
    'sPH_col_LoadChildEntitiesJT:
    ' ent.SetChildEntityString("ProjectEntityItems", "" & rst.Fields("sOrders").Value)
    'sPH_col_LoadChildEntitiesJT:End
    'Else
    'sPH_col_LoadChildEntitiesCS:
    ' Call ent.BuildChildEntityObjects("JTOrders", "" & rst.Fields("sOrders"))
    'sPH_col_LoadChildEntitiesCS:End
    'End Ifs
  End Function

  Public Sub RaiseChanged()
    RaiseEvent Changed()
  End Sub

  Private Function RecordChanges(ByVal sOperation As String, ByVal ent As ProjectEntity)
    Call BLFunctions.gRecordChanges1(sOperation, "EntityName", ent.EntityName, "en:")
    Call BLFunctions.gRecordChanges1(sOperation, "EntityCollectionName", ent.EntityCollectionName, "en:")
    Call BLFunctions.gRecordChanges1(sOperation, "EntityShortName", ent.EntityShortName, "en:")
    Call BLFunctions.gRecordChanges1(sOperation, "EntityDBTableName", ent.EntityDBTableName, "en:")
    Call BLFunctions.gRecordChanges1(sOperation, "GenerateFlag", ent.GenerateFlag, "en:")
    Call BLFunctions.gRecordChanges1(sOperation, "InputFile", ent.InputFile, "en:")
    Call BLFunctions.gRecordChanges1(sOperation, "FilesToGenerate", ent.FilesToGenerate, "en:")
    Call BLFunctions.gRecordChanges1(sOperation, "DateTimeCodeGenerated", ent.DateTimeCodeGenerated, "en:")

  End Function

End Class
