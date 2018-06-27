Option Explicit On

Public Class Entity2

  'sPH_cls_DateTime:
  'gfhiii
  'sPH_cls_DateTime:End
  'sh-Entity2

  Public Loadorder As Integer  'Load order
  'PH:For Each Entity Item. Create fields with clear naming conventions and with the required data types
  'sPH_cls_Decl:
  Dim nid As Integer
  Dim sEntityItem_1 As String
  Dim sEntityItem_2 As String
  'sPH_cls_Decl:End
  'PH: End

  ''PH:For Each Child. Add two fields for each child entity this entity supports.
  'Replace all variable of Entity3 with your own <EntityName> variables
  'Dim mChildEntity3s As New Entity3s
  'Dim sChildEntity3sString As String
  ''PH: End
  'sPH_cls_ChildDecl:
  'sPH_cls_ChildDecl
  'sPH_cls_ChildDecl:End

  'Dim mChildEntityCs As New EntityCs
  'Dim sChildEntityCsString As String

  ''PH:For Each Parent, add a field for each of the parent entity this entity supports. Replace all variable of Entity1 with your own <EntityName> variables
  'Dim mParentEntity1s As New Entity1s
  ''PH: Ends
  'sPH_cls_ParentDecl:
  'sPH_cls_ParentDecl-placeholder for parent decls.
  'sPH_cls_ParentDecl:End

  Public mContainer As Object
  Dim sLastUpdate As Date
  'PH:For Each Entity Item. Create a Property. Use clear naming conventions and take care to match the data types

  'sPH_cls_Properties:
  Public Property ID() As Integer
    Get
      ID = nid
    End Get
    Set(ByVal value As Integer)
      nid = value
    End Set
  End Property
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
      'sPH_cls_ChildEntities
      'sPH_cls_ChildEntities:End
    End Select
  End Function

  Public Function GetChildEntityString(ByVal sEnt As String) As String
    Select Case sEnt
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3", "Entity3s":
      '    ChildEntityString = sChildEntity3sString
      '  'PH: End
      'sPH_cls_GetChildString:
      'sPH_cls_GetChildStringgdfgsx
      'sPH_cls_GetChildString:End
    End Select
  End Function

  Public Sub SetChildEntityString(ByVal sEnt As String, ByVal strEnt As String)
    Select Case sEnt
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3s", "Entity3":
      '    sChildEntity3sString = strEnt
      '  'PH: End
      'sPH_cls_LetChildString:
      'sPH_cls_LetChildStringytrd
      'sPH_cls_LetChildString:End
    End Select
  End Sub

  Public Sub BuildChildEntityObjects(ByVal strPar As String, ByVal strEnt As String)
    Call BLFunctions.gBuildChildEntityObjects(Me, strPar, strEnt)
  End Sub

  Public Sub BuildChildEntityString(ByRef enChlds As String)
    Call gBuildChildEntityString(Me, enChlds)
  End Sub

  ' Parent Entity Methods.
  Public Function ParentEntities(ByVal objPar As Object)
    Select Case TypeName(objPar)
      ''PH: For Each Parent entity, add a case statement
      '  Case "Entity1", "Entity1s":
      '    Set ParentEntities = mParentEntity1s
      ' 'PH: End
      'sPH_cls_ParentEntities:
      'pholder
      'sPH_cls_ParentEntities:End
    End Select
  End Function

  Public Sub New()
    mContainer = cEntity2s
  End Sub

  'Public Function GetItemProperty(ByVal strP As String) As String
  '  Select Case strP
  '    'sPH_cls_GetItemProperty:
  '    Case "EntityItem_1"
  '      GetItemProperty = Me.EntityItem_1
  '    Case "EntityItem_2"
  '      GetItemProperty = Me.EntityItem_2
  '      'sPH_cls_GetItemProperty:End
  '  End Select
  'End Function


End Class
Public Class Entity2s
  Inherits System.Collections.CollectionBase
  'sPH_col_DateTime:
  'gfh
  'sPH_col_DateTime:End
  'Entity2s
  Dim nDecLoadOrder As Integer  'comment
  'Private collInt As Collection
  Public mContained As Object
  Dim strChanges As String
  Public Event Changed()
  Dim strSqlCall As String
  'Dim strConnFramework As String

  Public Sub New()
    'collInt = New Collection
    mContained = New Entity2
  End Sub

  Public Sub Add(ByRef val As Entity2)
    Me.List.Add(val)
  End Sub


  Public Function GetCount() As Integer
    Return Me.List.Count
  End Function

  Public ReadOnly Property Items() As Collection
    Get
      Return Me.List
    End Get
  End Property

  Public Property item(ByVal index As Integer) As Entity2
    Get
      Return Me.List.Item(index)
    End Get
    Set(ByVal value As Entity2)
      Me.List.Item(index) = value
    End Set
  End Property

  Public Sub Remove(ByRef val As Entity2)
    Me.List.Remove(val.ID)
  End Sub

  Public Function LoadSingleEntity(ByVal sid As String, ByVal sOperation As String) As Integer
    Dim e2_ As Entity2
    DBStrings.gConnectionString = MGlobals.GetAppConnString
    strSqlCall = GetSQLQuery("LoadSingleEntity", sid)



    cnn.Open(DBStrings.gConnectionString)
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
      e2_ = New Entity2
    ElseIf sOperation = "U" Then
      e2_ = Me.item(sid)
    End If
    Call Me.LoadEntityItemsForThisEntity(e2_)

    e2_.Lastupdate = rst.Fields("sLastUpdate").Value
    e2_.mContainer = Me

    If sOperation = "N" Then
      e2_.Loadorder = Me.Count - nDecLoadOrder
      Call Me.Add(e2_)
    End If
    rst.Close()
    cnn.Close()
    LoadSingleEntity = 1
  End Function

    Public Function Load()
        Dim e2_ As Entity2
        Dim nLoadorder As Integer
        strSqlCall = GetSQLQuery("Load", "ID")
    'DBStrings.gConnectionString = MGlobals.GetAppConnString
        'DBStrings.cnn1.ConnectionString = DBStrings.gConnectionString
        DBStrings.cmd1.CommandText = strSqlCall
        DBStrings.cmd1.CommandType = CommandType.Text
        DBStrings.cmd1.Connection = DBStrings.cnn1
        DBStrings.cnn1.Open()
        DBStrings.dr1 = DBStrings.cmd1.ExecuteReader


        'cnn.Open(strConnFramework)
        'rst.CursorLocation = adUseClient
        'rst.Open(strSqlCall, cnn, _
        'ADODB.CursorTypeEnum.adOpenStatic, 1)
    Me.Clear()
        'rst.MoveFirst()
        nDecLoadOrder = 0
        'if this entity has a M-1 rel with any other entity, then a blank entity must be present for cbos.
        If cEntityDataItems.ParentHasM1Relationship("Entity2") Then
            e2_ = New Entity2   'this member is loaded for 0 index entries cbo's for M-1 relationships
            e2_.Loadorder = 0
            e2_.ID = 0
            nDecLoadOrder = 1
            Call Me.Add(e2_)
        End If
        nLoadorder = 1
        Do While DBStrings.dr1.Read
            'Do While Not rst.EOF
            e2_ = New Entity2

            Call Me.LoadEntityItemsForThisEntity(e2_)

            e2_.Lastupdate = DBStrings.dr1.Item("sLastUpdate")
            e2_.Loadorder = nLoadorder
            e2_.mContainer = Me
            Call Me.Add(e2_)
            nLoadorder = nLoadorder + 1
            'rst.MoveNext()
        Loop
        DBStrings.cmd1.Dispose()
        DBStrings.cnn1.Close()

        'Build the domain model now
        'Call Me.BuildDomainModel
    End Function

    Sub BuildDomainModel()
        '  Dim eb_ As Entity2
        '  Dim e1_ As Entity1
        '  'now add e2_s to e1_s
        '  For Each e1_ In cEntity1s.Items  'assumed cEntity1s have been loaded
        '    Set e1_.ChildEntity2s = e1_.BuildChildEntityObjects(e1_.ChildEntity2s, e1_.ChildEntityString(Curre2_))
        '    'For Each e2_ In e1_.ChildEntity2s.Items
        '    ' If Not e2_.ContainsParentEntity(e1_) Then  'BuildObjectModel is called from AddtoDB() too. hence to check for already existing.
        '    '  Call e2_.ParentEntity1s.Add(e1_)   'add parents to Entity2
        '    'End If
        '    'Next e2_
        '  Next e1_
    End Sub

    Public Function UpdateDB(ByRef e2_ As Entity2) As Integer
        'Updates the database table for the entity'.
        'Return 0 if the db record has been updated by other user.
        'Return the loadorder if successfull in update.
        Dim nCount As Integer
        gbChangesMade = False
        strSqlCall = GetSQLQuery("UpdateDB", e2_.ID)
        ' cnn.Open(strConnFramework)
        'rst.Open(Source:=gStrSqlCall, ActiveConnection:=cnn, _
        '        CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
        DBStrings.cmd1.CommandText = strSqlCall
        DBStrings.cmd1.CommandType = CommandType.Text
        DBStrings.cmd1.Connection = DBStrings.cnn1
        DBStrings.cnn1.Open()
        DBStrings.dr1 = DBStrings.cmd1.ExecuteReader

        nCount = DBStrings.dr1.HasRows
        If nCount <> 1 Then
            rst.Close()
            cnn.Close()
            UpdateDB = 0
            Exit Function
        End If
        If rst.Fields("sLastUpdate").Value <> e2_.Lastupdate Then
            rst.Close()
            cnn.Close()
            UpdateDB = 0
            Exit Function
        Else
            rst.Fields("sLastUpdate").Value = Now
            e2_.Lastupdate = rst.Fields("sLastUpdate").Value  'update the local last update so you can mame changes without refreshing.
        End If
        strChanges = ""
        ''PH:For Each Entity Item. Add to record changes int the DB (function UpdateDB()
        'Call BLFunctions.gRecordChanges(e2_, "sEntityItem_1", e2_.EntityItem_1, "fn:")
        ''PH: End
    'sPH_col_Update:
    Call RecordChanges("", e2_)
    'Call BLFunctions.gRecordChanges("sEntityItem_1", e2_.EntityItem_1, "ei1:")
    'Call BLFunctions.gRecordChanges("sEntityItem_2", e2_.EntityItem_2, "ei2:")
        'sPH_col_Update:End

        ''PH: For Each Child Entity, add this line for recording the changes for child entities for this object.
        ' Call e2_.BuildChildEntityString(e2_.ChildEntities(e3_Par))
        ' Call BLFunctions.gRecordChanges(e2_, "sEntity3s", e2_.ChildEntityString(e3_Par), "e3_str:")
        ''PH: End
        'sPH_col_UpdateChild:
        'sPH_col_UpdateChild
        'sPH_col_UpdateChild:End

        rst.Update()
        rst.Close()
    MGlobals.Curre2_ = e2_

    'No need to enter a record in update log if no change has been made
    If gbChangesMade = True Then
      AppSettings.LastUpdateID = gUpdateLogTable("e2_", "U", e2_.ID)
    End If
    cnn.Close()
    Call DoChanged()
    UpdateDB = nCount
  End Function

  Public Function AddtoDB(ByRef e2_ As Entity2) As Integer
    Dim strSQL As String
    Dim nCount As Integer
    strSQL = "select Id from Entity2s"
    'cnn.Open(gStrConnGenEBA)
    rst.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
    nCount = rst.RecordCount
    If cEntityDataItems.ParentHasM1Relationship("Entity2") Then nCount = nCount + 1
    If nCount <> cEntity2s.Count Then  'other user has added a record
      rst.Close()
      cnn.Close()
      AddtoDB = 0
      Exit Function
    End If
    rst.Close()
    strSQL = "select max(ID) as MaxNo from Entity2s"
    Dim nMax As Integer
    rst.CursorLocation = adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, adLockOptimistic)
    nMax = rst.Fields("MaxNo").Value
    rst.Close()
    e2_.ID = nMax + 1
    'now add the record
    rst.CursorType = adOpenKeySet
    rst.LockType = adLockOptimistic
    rst.Open("Entity2s", cnn, , , adCmdTable)
    rst.AddNew()

    ''PH:For Each Entity Item. Add the items to the db (function AddtoDB)
    'Call gAddChanges("EntityItem_2", e2_.EntityItem_2, "ln:")
    ''PH: End
    'sPH_col_Add:
    Call BLFunctions.gRecordChanges1("Add", "ID", e2_.ID, "id:")
    Call RecordChanges("Add", e2_)
    'Call gAddChanges("EntityItem_1", e2_.EntityItem_1, "ei1:")
    'Call gAddChanges("EntityItem_2", e2_.EntityItem_2, "ei2:")
    'sPH_col_Add:End

    ''PH: For Each Child Entity, add this line for saving the child string with e3_ is the child.
    ' Call gAddChanges("Entity_Cs", e2_.ChildEntityString(e3_Par), "e2_str:")
    ''PH: End
    'sPH_col_AddChild:
    ''''sPH_col_AddChild

    'sPH_col_AddChild:End

    rst("LastUpdate").Value = Now
    e2_.Lastupdate = rst("LastUpdate").Value
    e2_.mContainer = Me
    rst.Update()
    rst.Close()

    e2_.Loadorder = Me.Count - nDecLoadOrder
    Call Me.Add(e2_)

    Call Me.BuildDomainModel()
    MGlobals.Curre2_ = e2_

    'now update the log
    AppSettings.LastUpdateID = gUpdateLogTable("e2_", "N", e2_.ID)
    cnn.Close()
    Call DoChanged()
    AddtoDB = nCount + 1
  End Function

  Public Function DeleteFromDB(ByVal e2_ As Entity2) As Integer
    Dim lAffected As Long

    gStrSqlCall = " Delete " & _
                  "FROM " & strDBO & "Entity2s " & _
                  " where " & strDBO & "ID LIKE '" & e2_.ID & "'"
    cmd.ActiveConnection = cnn
    'Globals.gADOcmd.ActiveConnection = gStrConnGenEBA
    cmd.CommandText = gStrSqlCall
    cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
    cmd.Execute(lAffected, ADODB.ExecuteOptionEnum.adExecuteNoRecords)
    If lAffected = 1 Then
      Dim sEbID As String
      sEbID = e2_.ID
      MGlobals.gstrChanges = "Deleted: " & "ID"
      Call cEntity2s.Remove(e2_)
      cnn.Open(gStrConnGenEBA)
      AppSettings.LastUpdateID = gUpdateLogTable("e2_", "D", sEbID)
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
                  strDBO & "ID as sID, " & _
                  strDBO & "EntityItem_1 as sEntityItem_1, " & _
                  strDBO & "EntityItem_2 as sEntityItem_2, " & _
                  strDBO & "LastUpdate as sLastUpdate " & _
                  "FROM " & strDBO & "Entity2s "
        'sPH_col_Sql:End

        If strFunc = "LoadSingleEntity" Or strFunc = "UpdateDB" Then
            strTemp = strTemp & " WHERE " & strDBO & "ID like " & strID
        ElseIf strFunc = "Load" Then
            strTemp = strTemp & " ORDER BY " & strDBO & strID
        End If
        GetSQLQuery = strTemp
    End Function

    Function LoadEntityItemsForThisEntity(ByVal ent As Entity2)
        ''PH:For Each Entity Item. Add it for loading
        'eA_.EntityItem_1 = "" & rst.Fields("sEntityItem_1")
        ''PH End
        'sPH_col_Load:
        ent.ID = "" & DBStrings.dr1.Item("sID")
        ent.EntityItem_1 = "" & DBStrings.dr1.Item("sEntityItem_1")
        ent.EntityItem_2 = "" & DBStrings.dr1.Item("sEntityItem_2")
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

  Private Sub RecordChanges(ByVal sOperation As String, ByVal ent As Entity2)
    Call BLFunctions.gRecordChanges1(sOperation, "sEntityItem_1", ent.EntityItem_1, "ei1:")
    Call BLFunctions.gRecordChanges1(sOperation, "sEntityItem_2", ent.EntityItem_2, "ei2:")
  End Sub


End Class