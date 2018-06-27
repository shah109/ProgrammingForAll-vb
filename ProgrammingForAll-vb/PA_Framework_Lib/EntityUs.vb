Option Explicit On

Public Class EntityU
  Public Loadorder As Integer  'Load order
  Dim nID As Integer
  Dim sEntityItem_1 As String
  Dim sEntityItem_2 As String

  Dim sFirstName As String
  Dim sLastName As String
  Dim sMiddleName As String
  Dim sLoginID As String
  Dim sEmail As String
  Dim sAccessRights As String

  Public mContainer As Object
  Dim dLastUpdate As Date



  Public Property ID()
    Get
      ID = nID
    End Get
    Set(ByVal value)
      nID = value
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

  Public Property FirstName()
    Get
      FirstName = sFirstName
    End Get
    Set(ByVal value)
      sFirstName = value
    End Set
  End Property

  Public Property LastName()
    Get
      LastName = sLastName
    End Get
    Set(ByVal value)
      sLastName = value
    End Set
  End Property
  Public Function FullName() As String
    Dim sGap As String
    If sMiddleName = "" Then
      sGap = ""
    Else
      sGap = " "
    End If
    Return sFirstName & " " & sMiddleName & sGap & sLastName
  End Function

  Public Property AccessRights()
    Get
      AccessRights = sAccessRights
    End Get
    Set(ByVal value)
      sAccessRights = value
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
  Public Property Email()
    Get
      Email = sEmail
    End Get
    Set(ByVal value)
      sEmail = value
    End Set
  End Property


  Public Property LastUpdate()
    Get
      LastUpdate = dLastUpdate
    End Get
    Set(ByVal value)
      dLastUpdate = value
    End Set
  End Property
End Class

Public Class EntityUs
  Inherits System.Collections.CollectionBase
  Dim nDecLoadOrder As Integer  'comment
  'Private collInt As Collection
  Public mContained As Object
  Dim strChanges As String
  Public Event Changed()
  Dim strSqlCall As String
  Dim strConnFramework As String

  Public Sub New()
    'collInt = New Collection
    mContained = New EntityU
  End Sub

  Public Sub Add(ByRef val As EntityU)
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

  Public Property item(ByVal index As Integer) As EntityU
    Get
      Return Me.List.Item(index)
    End Get
    Set(ByVal value As EntityU)
      Me.List.Item(index) = value
    End Set
  End Property

  Public Sub Remove(ByRef val As EntityU)
    Me.List.Remove(val.ID)
  End Sub


  Public Function LoadSingleEntity(ByVal sid As String, ByVal sOperation As String) As Integer
    Dim eU_ As EntityU
    strConnFramework = MGlobals.GetAppConnString
    strSqlCall = GetSQLQuery("LoadSingleEntity", sid)
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
            eU_ = New EntityU
        ElseIf sOperation = "U" Then
            eU_ = Me.item(sid)
        End If
        Call Me.LoadEntityItemsForThisEntity(eU_)

    eU_.LastUpdate = rst.Fields("LastUpdate")
        eU_.mContainer = Me

        If sOperation = "N" Then
            eU_.Loadorder = Me.Count - nDecLoadOrder
            Call Me.Add(eU_)
        End If
        rst.Close()
        cnn.Close()
        LoadSingleEntity = 1
    End Function

    Public Function Load()
        Dim eU_ As EntityU
        Dim nLoadorder As Integer
    strConnFramework = MGlobals.GetAppConnString
        strSqlCall = GetSQLQuery("Load", "ID")

        DBStrings.cmd1.CommandText = strSqlCall
        DBStrings.cmd1.CommandType = CommandType.Text
        DBStrings.cmd1.Connection = DBStrings.cnn1
        DBStrings.cnn1.Open()
        DBStrings.dr1 = DBStrings.cmd1.ExecuteReader



    Me.Clear()
        nDecLoadOrder = 0
        'if this entity has a M-1 rel with any other entity, then a blank entity must be present for cbos.
        If cEntityDataItems.ParentHasM1Relationship("EntityU") Then
            eU_ = New EntityU   'this member is loaded for 0 index entries cbo's for M-1 relationships
            eU_.Loadorder = 0
            eU_.ID = 0
            nDecLoadOrder = 1
            Call Me.Add(eU_)
        End If
        nLoadorder = 1
        Do While DBStrings.dr1.Read
            'Do While Not rst.EOF
            eU_ = New EntityU

            Call Me.LoadEntityItemsForThisEntity(eU_)

            eU_.LastUpdate = DBStrings.dr1.Item("sLastUpdate")
            eU_.Loadorder = nLoadorder
            eU_.mContainer = Me
            Call Me.Add(eU_)
            nLoadorder = nLoadorder + 1
            'rst.MoveNext()
        Loop

        DBStrings.cmd1.Dispose()
        DBStrings.cnn1.Close()
        'Build the domain model now
        'Call Me.BuildDomainModel
    End Function

    Sub BuildDomainModel()
        '  Dim eb_ As EntityU
        '  Dim e1_ As Entity1
        '  'now add eU_s to e1_s
        '  For Each e1_ In cEntity1s.Items  'assumed cEntity1s have been loaded
        '    Set e1_.ChildEntityUs = e1_.BuildChildEntityObjects(e1_.ChildEntityUs, e1_.ChildEntityString(CurreU_))
        '    'For Each eU_ In e1_.ChildEntityUs.Items
        '    ' If Not eU_.ContainsParentEntity(e1_) Then  'BuildObjectModel is called from AddtoDB() too. hence to check for already existing.
        '    '  Call eU_.ParentEntity1s.Add(e1_)   'add parents to EntityU
        '    'End If
        '    'Next eU_
        '  Next e1_
    End Sub

    Public Function UpdateDB(ByRef eU_ As EntityU) As Integer
        'Updates the database table for the entity'.
        'Return 0 if the db record has been updated by other user.
        'Return the loadorder if successfull in update.
        Dim nCount As Integer
        gbChangesMade = False
        strSqlCall = GetSQLQuery("UpdateDB", eU_.ID)
        cnn.Open(strConnFramework)
        rst.Open(Source:=gStrSqlCall, ActiveConnection:=cnn, _
                 CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
        nCount = rst.RecordCount
        If nCount <> 1 Then
            rst.Close()
            cnn.Close()
            UpdateDB = 0
            Exit Function
        End If
        If rst.Fields("sLastUpdate").Value <> eU_.LastUpdate Then
            rst.Close()
            cnn.Close()
            UpdateDB = 0
            Exit Function
        Else
            rst.Fields("sLastUpdate").Value = Now
            eU_.LastUpdate = rst.Fields("sLastUpdate").Value  'update the local last update so you can mame changes without refreshing.
        End If
        strChanges = ""
        ''PH:For Each Entity Item. Add to record changes int the DB (function UpdateDB()
        'Call BLFunctions.gRecordChanges(eU_, "sEntityItem_1", eU_.EntityItem_1, "fn:")
        ''PH: End
    'sPH_col_Update:
    Call RecordChanges("", eU_)
    'Call BLFunctions.gRecordChanges("sEntityItem_1", eU_.EntityItem_1, "ei1:")
    'Call BLFunctions.gRecordChanges("sEntityItem_2", eU_.EntityItem_2, "ei2:")
        'sPH_col_Update:End

        ''PH: For Each Child Entity, add this line for recording the changes for child entities for this object.
        ' Call eU_.BuildChildEntityString(eU_.ChildEntities(e3_Par))
        ' Call BLFunctions.gRecordChanges(eU_, "sEntity3s", eU_.ChildEntityString(e3_Par), "e3_str:")
        ''PH: End
        'sPH_col_UpdateChild:
        'sPH_col_UpdateChild
        'sPH_col_UpdateChild:End

        rst.Update()
        rst.Close()
    MGlobals.CurreU_ = eU_

    'No need to enter a record in update log if no change has been made
    If gbChangesMade = True Then
      AppSettings.LastUpdateID = gUpdateLogTable("eU_", "U", eU_.ID)
    End If
    cnn.Close()
    Call DoChanged()
    UpdateDB = nCount
  End Function

  Public Function AddtoDB(ByRef eU_ As EntityU) As Integer
    Dim strSQL As String
    Dim nCount As Integer
    strSQL = "select Id from EntityUs"
    'cnn.Open(gStrConnGenEBA)
    rst.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
    nCount = rst.RecordCount
    If cEntityDataItems.ParentHasM1Relationship("EntityU") Then nCount = nCount + 1
    If nCount <> cEntityUs.Count Then  'other user has added a record
      rst.Close()
      cnn.Close()
      AddtoDB = 0
      Exit Function
    End If
    rst.Close()
    strSQL = "select max(ID) as MaxNo from EntityUs"
    Dim nMax As Integer
    rst.CursorLocation = adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, adLockOptimistic)
    nMax = rst.Fields("MaxNo").Value
    rst.Close()
    eU_.ID = nMax + 1
    'now add the record
    rst.CursorType = adOpenKeySet
    rst.LockType = adLockOptimistic
    rst.Open("EntityUs", cnn, , , adCmdTable)
    rst.AddNew()

    ''PH:For Each Entity Item. Add the items to the db (function AddtoDB)
    'Call gAddChanges("EntityItem_2", eU_.EntityItem_2, "ln:")
    ''PH: End
    'sPH_col_Add:
    Call BLFunctions.gRecordChanges1("Add", "EntityItem_1", eU_.EntityItem_1, "ei1:")
    Call RecordChanges("", eU_)
    ' Call gAddChanges("ID", eU_.ID, "id:")
    ' Call gAddChanges("EntityItem_1", eU_.EntityItem_1, "ei1:")
    ' Call gAddChanges("EntityItem_2", eU_.EntityItem_2, "ei2:")
    'sPH_col_Add:End

    ''PH: For Each Child Entity, add this line for saving the child string with e3_ is the child.
    ' Call gAddChanges("Entity_Cs", eU_.ChildEntityString(e3_Par), "eU_str:")
    ''PH: End
    'sPH_col_AddChild:
    ''''sPH_col_AddChild

    'sPH_col_AddChild:End

    rst("LastUpdate").Value = Now
    eU_.LastUpdate = rst("LastUpdate").Value
    eU_.mContainer = Me
    rst.Update()
    rst.Close()

    eU_.Loadorder = Me.Count - nDecLoadOrder
    Call Me.Add(eU_)

    Call Me.BuildDomainModel()
    MGlobals.CurreU_ = eU_

    'now update the log
    AppSettings.LastUpdateID = gUpdateLogTable("eU_", "N", eU_.ID)
    cnn.Close()
    Call DoChanged()
    AddtoDB = nCount + 1
  End Function

  Public Function DeleteFromDB(ByVal eU_ As EntityU) As Integer
    Dim lAffected As Long
    strConnFramework = MGlobals.GetAppConnString
    gStrSqlCall = " Delete " & _
                  "FROM " & strDBO & "EntityUs " & _
                  " where " & strDBO & "ID LIKE '" & eU_.ID & "'"
    cmd.ActiveConnection = cnn
    'Globals.gADOcmd.ActiveConnection = gStrConnGenEBA
    cmd.CommandText = gStrSqlCall
    cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
    cmd.Execute(lAffected, ADODB.ExecuteOptionEnum.adExecuteNoRecords)
    If lAffected = 1 Then
      Dim sEbID As String
      sEbID = eU_.ID
      MGlobals.gstrChanges = "Deleted: " & "ID"
      Call cEntityUs.Remove(eU_)
      cnn.Open(gStrConnGenEBA)
      AppSettings.LastUpdateID = gUpdateLogTable("eU_", "D", sEbID)
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
                  "FROM " & strDBO & "EntityUs "
        'sPH_col_Sql:End

        If strFunc = "LoadSingleEntity" Or strFunc = "UpdateDB" Then
            strTemp = strTemp & " WHERE " & strDBO & "ID like " & strID
        ElseIf strFunc = "Load" Then
            strTemp = strTemp & " ORDER BY " & strDBO & strID
        End If
        GetSQLQuery = strTemp
    End Function

    Function LoadEntityItemsForThisEntity(ByVal ent As EntityU)
        ''PH:For Each Entity Item. Add it for loading
        'eA_.EntityItem_1 = "" & rst.Fields("sEntityItem_1")
        ''PH End
        'sPH_col_Load:
        ent.ID = "" & DBStrings.dr1.Item("sID")
        ent.EntityItem_1 = "" & DBStrings.dr1.Item("sEntityItem_1")
        ent.EntityItem_2 = "" & DBStrings.dr1.Item("sEntityItem_2")
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
  Private Sub RecordChanges(ByVal sOperation As String, ByVal ent As EntityU)
    Call BLFunctions.gRecordChanges1(sOperation, "sEntityItem_1", ent.EntityItem_1, "ei1:")
    Call BLFunctions.gRecordChanges1(sOperation, "sEntityItem_2", ent.EntityItem_2, "ei2:")
  End Sub


End Class
