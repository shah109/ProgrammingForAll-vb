﻿Option Explicit On
Imports PA_Framework_Lib
Public Module Globals
    'sPH_mod_DateTime:
    'gfh
    'sPH_mod_DateTime:End
    '212
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'Globals Module Source -
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    'Types
    Public Structure PCTypeItem
        Dim DoneFlag As Boolean
        Dim eParent As String
        Dim eChild As String
        Dim eParentString As String
        Dim nParentCount As Integer
    End Structure
    'Public Type PCCollection
    '  item() As PCTypeItem
    'End Type

    'Declare Global variables
    'Declare ADODB variables
    'Public cnn As New OleDb.OleDbConnection
    'Public rst As New OleDb.OleDbDataReader
    Public cnn As New ADODB.Connection
    Public rst As New ADODB.Recordset
    Public cmd As New ADODB.Command
    'Public gADOcmd As New ADODB.Command
    Public gStrAssoc As String  'Association string with other child and parent entities
    Public gStrSqlCall As String  'sql call string
    Public gstrChanges As String  'string containing the changes made to entity during db update.
    Public gbChangesMade As Boolean  ' flag to indicate that changes were made to strchanges during update

    'Entity Data items variables
    Public gsJoinTable As String    'Join table entity string from Entity Data Items
    Public gsChildPropertyName As String  ' to get from Entity Data items
    Public gsChildEntityNameFromJT As String
    'PH: Application Name
    Public Const APPLongName = "The Framework"
    Public Const APPName = "FrmWrk"
    'PH: Application Name End

    Public cEntityDataItems As PA_Framework_Lib.EntityDataItems
    Public cUpdateLogItems As PA_Framework_Lib.UpdateLogItems

    'PH: For Each Entity. Global Class Collections
    'Public cRepairs As Repairs
    'PH: End

    'sPH_glbl_CollDecls:
    'Public cEntityUs As EntityUs
    Public cEntity2s As Entity2s
    Public cEntityUs As EntityUs
    Public cEntityAs As EntityAs
    Public cAppArrays As AppArrays
    Public cProjects As Projects
    Public cProjectEntities As ProjectEntities
    Public cProjectEntityItems As ProjectEntityItems

    'Public cEntityDs As EntityDs
    'Public cEntityCs As EntityCs
    'Public cEntityBs As EntityBs
    'Public cEntityAs As EntityAs
    'sPH_glbl_CollDecls:End

    'PH:For Each Entity. Currently Selected Entity
    'Public CurreRprs_ As New Repair
    'PH: End
    'sPH_glbl_CurrSelEntity:
    'Public CurreU_ As New EntityU
    Public Curre2_ As New Entity2
    Public CurreU_ As New EntityU
    Public currePrjt_ As New Project
    Public currePrjEnt_ As New ProjectEntity
    Public currePrjEntItem_ As New ProjectEntityItem
    Public curreA_ As New EntityA
    Public currRow As Integer  'current row of the selected sheet
    Public currObjCollection As Object  'obj collection of the currently selected sheet
    Public currForm As Object    'Entity form of the currently selected sheet.
    Public currEnt As Object  ' entity of the current sheet.

    'Database connection related constants
    'Public Const adOpenStatic = 3
    Public Const adLockOptimistic = 3
    Public Const adUseClient = 3
    Public Const adOpenKeySet = 1
    Public Const adCmdTable = 2
    Public Const adOpenStatic = 3

    Public gStrConnGenEBA As String  'Connection string for the GenEBA db.
    Public strDBO As String
    Public gDBLoaded As Boolean  ' a variable to determine if db is loaded. Variables lose their value if the app goes into runtime error

    Public Const NSTARTROW = 3
    Public bChangeColorOnEdit As Boolean  'used in control classes
    Public LoggedInUser As New PA_Framework_Lib.EntityU  'to store the logged in member object

    'DB Update related variables
    Public PrevUpdateID As Integer
    Public NewUpdateID As Integer

    'Miscellaneous enumerations

    Public Enum eNum_Column
        'the first 5 columns of sheets
        HypDetails = 2
        HypEdit = 3
        Loadorder = 4
        ID = 5
    End Enum

    Public Enum EntityUStatus
        Guest
        Dormant
        Valid
        Discontinued
    End Enum

    Public Enum FormEditMode
        browse = 0
        Edit = 1
        Add = 2
    End Enum

    Public Sub DBLoad()
        'Instantiates and loads all entity collections in the correct order. Child entities need to be loaded before their parents.
        'Dont call this function directly, preferrably call via CallLoadDBIfNeeded(). This way it does not get
        'called unnecessarily when the data is already up to date.
        'Framework Gen App:


        cEntityDataItems = New EntityDataItems
        cUpdateLogItems = New UpdateLogItems
        cAppArrays = New AppArrays
        'Create connectin string
        'cnn.ConnectionString = GetAppConnString()
        'Call GetAppConnString()
        DBStrings.gConnectionString = DBStrings.GetAppConnString
        DBStrings.cnn1.ConnectionString = DBStrings.gConnectionString

        cEntityDataItems.Load()
        cAppArrays.LoadAppArrays()


        'PH: For Each Entity. Instantiate the collections in DBLoad()
        'Set cEntity1s = New Entity1s
        'PH: End
        'sPH_glbl_CollInstantiation:
        cEntityUs = New EntityUs
        cEntity2s = New Entity2s
    cEntityAs = New EntityAs
    cProjects = New Projects
    cProjectEntities = New ProjectEntities
        'cEntityDs = New EntityDs
        'cEntityCs = New EntityCs
        'cEntityBs = New EntityBs
        'cEntityAs = New EntityAs
        'sPH_glbl_CollInstantiation:End

        'Start loading with entities that have no childs and then move up
        'PH: For Each Entity, Load the collection
        'cRepairs.Load  'has eRprs_ as a child entity.
        'PH: End
        'sPH_glbl_CollLoad:

        cEntityAs.Load()
        cEntityUs.Load()
        cEntity2s.Load()
    cProjects.Load()
    cProjectEntities.Load()
        'cEntityDs.Load()
        'cEntityCs.Load()
        'cEntityBs.Load()
        'cEntityAs.Load()
        'sPH_glbl_CollLoad:End

        LoggedInUser = Nothing
        AppSettings.LastUpdateID = cUpdateLogItems.GetMaxUpdateID

        gDBLoaded = True  ' Flag to verify that db is loaded ...Variables in
        'memory lose their value as soon as runtime error occurs.

        Call GetAccessRightForLoggedInUser()
    End Sub

    Sub GetAccessRightForLoggedInUser()
        'First check if member authorizations are disabled in AppSettings. This is the default
        'at first App Launch,or can be kept disabled if there is only a single user for the system.
        If AppSettings.EnableAccessRights = False Then
            AppSettings.AccessRights = 2
            If LoggedInUser Is Nothing Then LoggedInUser = cEntityUs(0) 'make the loggedinEntityU as the pseudo first user
            LoggedInUser.ID = 0
            Exit Sub
        End If
        Dim login As String
        Call GetUserDetails()  'writes user details (login name, computer name, domain name) to settings sheet
        login = LCase(AppSettings.sLoginName)
        Dim m As EntityU
        For Each m In cEntityUs.Items
            If login = LCase(m.LoginID) Then
                AppSettings.AccessRights = m.AccessRights  '0:Guest, 1:User, 2:Admin
                LoggedInUser = m
                Exit Sub
            End If
nextloop:
        Next m
        AppSettings.AccessRights = 0  'no matching login present in db ; the user is a guest
    End Sub

    Sub GetUserDetails()
        Dim objNet
        objNet = CreateObject("WScript.NetWork")
        If Err.Number <> 0 Then                 'If error occured then display notice
            MsgBox("WScript.NetWorkError from GetUserDetails")
        End If
        AppSettings.sLoginName = objNet.UserName
        AppSettings.sComputerName = objNet.ComputerName
        AppSettings.sDomainName = objNet.userdomain
    End Sub

    Function CallDBLoadIfNeeded() As Boolean
        'Call DBLoad only if state is lost. Else calls DBUpdateNeeded to selectively update only those entities that have
        'been updated or added.
        Dim nUpdate As Integer
        CallDBLoadIfNeeded = False
        If Globals.gDBLoaded <> True Then  'Is state lost ?
            Call Globals.DBLoad()
            CallDBLoadIfNeeded = True
            Exit Function    'when dbload is called, there is no need to call DBUpdateNeeded. All updates come down with DBLoad.
        End If
        nUpdate = DBUpdateNeeded()  'returns the number of db updates made since last called; future use
        '  If nUpdate <> 0 Then
        '  End If
        If nUpdate > 0 Then CallDBLoadIfNeeded = True
    End Function

    Function DBUpdateNeeded() As Integer
        'Determins how many updates have been made by others since last update.
        'Then calls DBLoadWithUpdateLogItems which selectively re-loads each of the updated (or added) entity from their respective table.
        DBUpdateNeeded = 0
        PrevUpdateID = AppSettings.LastUpdateID
        Call cUpdateLogItems.LoadFromLastUpdate(PrevUpdateID)    'Loads all updatelog items since last update.
        If cUpdateLogItems.Count > 0 Then
            Call DBLoadWithUpdateLogItems(cUpdateLogItems)
            DBUpdateNeeded = cUpdateLogItems.Count
            AppSettings.LastUpdateID = PrevUpdateID + DBUpdateNeeded  'updated after the db update is done
        End If
    End Function

    Function DBLoadWithUpdateLogItems(ByVal UpLIs As UpdateLogItems)
        'Input UpLIs contains the collection of all updatelog items since last update.
        Dim ul As UpdateLogItem
        For Each ul In UpLIs.Items
            Select Case ul.sTableID    'Now call LoadSingleEntity for each updated entity table row (since last update)
                'PH: For Each Entity. Add a case statement
                'Case "eRprs_":
                'Call cRepairs.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
                'PH: End
                'sPH_glbl_LoadSingleEntity:
                Case "ePrjt_"
                    Call cProjects.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
                Case "e2_"
                    Call cEntity2s.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
                Case "eU_"
                    Call cEntityUs.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)

                    'Case "eD_"
                    '  Call cEntityDs.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
                    'Case "eC_"
                    '  Call cEntityCs.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
                    'Case "eB_"
                    '  Call cEntityBs.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
                    'Case "eA_"
                    '  Call cEntityAs.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)


                    'sPH_glbl_LoadSingleEntity:End



            End Select
        Next ul
    End Function

    Function CreateObjectFromString(ByVal sEnt As String) As Object  'to add into design
        Select Case sEnt
            'sPH_glbl_CreateObjectFromString:
            Case "Entity2", "Entity2s"
                CreateObjectFromString = New Entity2
            Case "EntityU", "EntityUs"
                CreateObjectFromString = New EntityU
            Case "Project", "Projects"
                CreateObjectFromString = New Project


                'Case "EntityA", "EntityAs"
                '  CreateObjectFromString = New EntityA
                'Case "EntityB", "EntityBs"
                '  CreateObjectFromString = New EntityB
                'Case "EntityC", "EntityCs"
                '  CreateObjectFromString = New EntityC
                'Case "EntityD", "EntityDs"
                '  CreateObjectFromString = New EntityD
                'Case "EntityU", "EntityUs"
                '  CreateObjectFromString = New EntityU
                '  'sPH_glbl_CreateObjectFromString:End
        End Select
    End Function

    'Function GetContainerFromString(sEnt As String) As Object 'to remove
    '    Select Case sEnt
    '        'sPH_glbl_GetContainerFromString:
    '    Case "EntityA", "EntityAs":
    '        Set GetContainerFromString = New EntityA
    '    Case "EntityB", "EntityBs":
    '        Set GetContainerFromString = New EntityB
    '    Case "EntityC", "EntityCs":
    '        Set GetContainerFromString = New EntityC
    '    Case "EntityD", "EntityDs":
    '        Set GetContainerFromString = New EntityD
    '
    '    Case "EntityU", "EntityUs":
    '        Set GetContainerFromString = cEntityUs
    '        'sPH_glbl_GetContainerFromString:End
    '    End Select
    'End Function

    'Function GetObjectFromProperty(sEnt As String) As Object  'to add into design
    '    Select Case sEnt
    '        'sPH_glbl_GetContainerFromString:
    '    Case "EntityA", "EntityAs":
    '        Set GetObjectFromProperty = CurreA_
    '    Case "EntityB", "EntityBs":
    '        Set GetObjectFromProperty = CurreA_
    '    Case "EntityC", "EntityCs":
    '        Set GetObjectFromProperty = CurreA_
    '    Case "EntityD", "EntityDs":
    '        Set GetObjectFromProperty = CurreA_
    '    Case "EntityU", "EntityUs":
    '        Set GetObjectFromProperty = CurreA_
    '        'sPH_glbl_GetContainerFromString:End
    '    End Select
    'End Function

    Function gUpdateLogTable(ByVal sTableID As String, ByVal sOperation As String, ByVal recID As String) As Integer
        'updates the log table 'UpdateLog' for recording history of changes made  by users.
        'nTableID to be assigned to each entity
        'sOperation="U" for update of existing entity, "N" for creation of new entity.
        'recID= ID of the record in the table that was updated or created
        'In addition to the above info, it also logs the current time and the user ID of the user who made the change.
        Dim strSQL As String

        rst.CursorType = adOpenKeySet
        rst.LockType = adLockOptimistic
        rst.Open("UpdateLog", cnn, , , adCmdTable)
        rst.AddNew()
        rst("DateTime").Value = Now
        rst("LoginID").Value = Globals.LoggedInUser.ID    'frmSettings.LoginName  'strLoginID
        rst("TableID").Value = sTableID  'id of the table
        rst("KeyFieldNo").Value = recID
        rst("Changes").Value = gstrChanges
        rst("Operation").Value = sOperation  '
        rst.Update()
        gstrChanges = ""
        rst.Close()
        strSQL = "select max(ID) as MaxNo from Updatelog"
        rst.CursorLocation = adUseClient
        rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, adLockOptimistic)
        gUpdateLogTable = rst.Fields("MaxNo").Value
        rst.Close()
    End Function

    Public Function GetAppConnString() As String
        'Gets the connection string for the database based on AppSettings.
        Dim strTMUser As String
        Dim strTMpw As String

        strTMUser = "TMAS"
        strTMpw = AppSettings.sAcccessLoginPassword
        GetAppConnString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                         "Data Source= " & AppSettings.sAccessDBPath & ";" & _
                         "Jet OLEDB:Database Password=" & strTMpw
        strDBO = ""
        'nd If
    End Function

    Public Function GetCurrentLoginStatus() As String
        'returns a message to show the available access right to the user in different ui sheets.
        Dim sLoginInfo As String = ""
        Select Case AppSettings.AccessRights
            Case 0, 3
                sLoginInfo = LoggedInUser.LastName & "/" & "Guest"
            Case 1
                sLoginInfo = LoggedInUser.LastName & "/" & "Member"
            Case 2
                If LoggedInUser.FirstName = "" Then
                    sLoginInfo = "First User" & "/" & "Admin"
                Else
                    sLoginInfo = LoggedInUser.LastName & "/" & "Admin"
                End If
        End Select
        GetCurrentLoginStatus = sLoginInfo
    End Function




    Function DoChanged()
        'Whatever form is visible, its changed event is triggered and update is done.
        'Dim obj As Object
        'On Error Resume Next
        'For Each obj In VBA.UserForms
        '  If obj.Visible Then Call obj.EntToRaiseChanged.RaiseChanged()
        'Next obj
    End Function

    'Sub LoadAppData()
    '  'Load all the item data in the entity sheet
    '  Call Globals.DBLoad()
    '  On Error Resume Next
    '  Call ActiveSheet.LoadEntitiesInSheet()
    '  On Error GoTo 0
    'End Sub


End Module
