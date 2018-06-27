Option Explicit On
Imports PA_Framework_Lib
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class OMGlobals
  Public Event EntityUpdated(ByVal sUpdateType As String, ByRef oEntity As Object, ByVal sReturn As Integer)
  Public Shared myInstance As OMGlobals
  Public Delegate Function Compare(ByVal v1 As Object, ByVal v2 As Object) As Boolean
  Shared Sub New()
    If myInstance Is Nothing Then
      myInstance = New OMGlobals
      myInstance.DBLoad()
    End If
  End Sub

  Public Shared Function GetInstance() As OMGlobals
    If myInstance Is Nothing Then
      myInstance = New OMGlobals
      myInstance.DBLoad()
    End If
    Return myInstance
  End Function

  Public Shared PASettings As New AppSettings
  Public Shared MainForm As Object
  Public Shared gStrSqlCall As String  'sql call string
  Public Shared gstrChanges As String  'string containing the changes made to entity during db update.
  Public Shared gbChangesMade As Boolean  ' flag to indicate that changes were made to strchanges during update

  Public Shared APPLongName = "PA_InstituteXLS"
  Public Shared APPShortName = "PAI"

  Public Shared cUpdateLogItems As UpdateLogItems
  Public Shared cChangeHistorys As ChangeHistorys

  Public Shared strDBO As String
  Public Shared LoggedInUser As New Person  'to store the logged in member object

  Public Sub PassFormToOM(ByRef frm As Object)
    MainForm = New Object
    MainForm = frm
  End Sub

  Sub GetAccessRightForLoggedInUser()
    'First check if member authorizations are disabled in AppSettings. This is the default
    'at first App Launch,or can be kept disabled if there is only a single user for the system.
    Dim MultiUserAccess As String
    Dim sLoginName As String = AppSettings.GetSetting(AppSettings.LoginName)
    MultiUserAccess = AppSettings.GetSetting("MultiUserAccess")
    If MultiUserAccess = "False" Then
      AppSettings.SetSetting("AccessRights", 2)
      'If sLoginName = "" Then LoggedInUser = New Person 'make the loggedinEntityU as the pseudo first user
      'LoggedInUser.ID = 0
      'LoggedInUser.FirstName = "First User"
      Exit Sub
    End If
    'Dim login As String
    'Call GetUserDetails()  'writes user details (login name, computer name, domain name) to settings sheet
    sLoginName = LCase(sLoginName)
    Dim m As Person
    For Each m In cPersons
      If sLoginName = LCase(m.LoginID) Then
        AppSettings.SetSetting("AccessRights", m.AccessRight) '0:Guest, 1:User, 2:Admin
        LoggedInUser = m
        Exit Sub
      End If
nextloop:
    Next m
    AppSettings.SetSetting("AccessRights", 0)  'no matching login present in db ; the user is a guest

  End Sub

  '  Private Sub WriteToUserLog(ByVal logValue)
  '    'Input parameter: LogValue="Login" when called from wkbk_open, or "Logout" when called from wkbk_before3_lose.
  '    'Writes user log with details of who logged in at what time and from what machine. When called from wks_before3_lose, also logs the
  '    'logout time with its login number of the user. This can be used to determine duration of the system usage by users.
  '    'Another functionality is to control the versions which can be used by the users. You may provide only a message to user if they are
  '    'using an older version, or may not let the user to login if they are using a certain version.

  '    'If this function is deemed to be not necessary, do  not call from wkbk_open and wkbk_before3_lose'
  '    Dim bRecordAdded As String
  '    Dim nFlag As Integer
  '    Dim strMessage As String
  '    'Dim cnn1 As Object
  '    'Dim rstReport As Object
  '    'Set cnn1 = CreateObject("ADODB.Connection")
  '    'Set rstReport = CreateObject("ADODB.Recordset")
  '    Dim strCnn As String
  '    Dim strID As String
  '    Dim strFirstName As String
  '    Dim strLastName As String
  '    Dim strETASDBPath As String
  '    If logValue = "LogIn" Then frmSettings.LoginNumber = 0

  '    bRecordAdded = False
  '    Call Globals.GetAppConnString()
  '    cnn.Open(gStrConnGenEBA)
  '    rst.CursorType = adOpenKeySet
  '    rst.LockType = adLockOptimistic
  '    rst.Open("UserLogs", cnn, , , adCmdTable)
  '    'create new record
  '    rst.AddNew()
  '    rst!UserName = frmSettings.LoginName
  '    rst!ComputerName = frmSettings.ComputerName
  '    rst!DateTime = Now
  '    rst!LoginLogout = logValue
  '    rst!Version = Range("APPVersion")
  '    rst!Build = Range("APPBuild")
  '    rst!LoginNo = frmSettings.LoginNumber
  '    rst.Update()
  '    bRecordAdded = True
  '    rst.Close()
  '    cnn.Close()
  '    cnn.Open()
  '    strAppVerBld = Trim(Range("APPVersion") & Range("APPBuild"))
  '    Dim strSQL As String

  '    strSQL = "SELECT flag, Message FROM versioncontrol WHERE Version like '" & strAppVerBld & "'" & _
  '             " ORDER BY sNo ASC"
  '    rst.Open(strSQL, cnn, adOpenStatic, adLockOptimistic)
  '    If rst.RecordCount > 0 Then
  '      rst.MoveFirst()
  '      nFlag = rst.Fields("flag")
  '      strMessage = rst.Fields("Message")
  '    End If
  '    rst.Close()
  '    'In case the user is to be given a message only Flag = 1
  '    'If the user is to be given a message and not allowed to log, Flag is 2
  '    If logValue = "LogIn" Then
  '      rst.Open("SELECT MAX(SNo) AS MaxsNo FROM UserLogs  ", _
  '               cnn, adOpenStatic, 1)
  '      'If rst.RecordCount > 0 Then
  '      'rst.MoveFirst()
  '      'frmSettings.LoginNumber = rst.Fields("MaxsNo")
  '    End If
  '    rst.Close()
  '    cnn.Close()
  '    Select Case nFlag
  '      Case 1
  '        MsgBox(strMessage)
  '      Case 2    'Application close not working
  '        MsgBox(strMessage)
  '        'Me.Close(SaveChanges:=False)
  '        Exit Sub
  '    End Select
  '    Else
  '    cnn.Close()
  '    End If
  '    Exit Sub
  'Handler:
  '    If logValue = "LogIn" Then
  '      MsgBox("           Application is not able to access the specified database in 'Settings'" & vbCrLf & _
  '             " ", vbOKOnly, " & AppName & " & strAppVerBld)
  '    End If
  '  End Sub

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

  Public Shared Sub gRecordChanges(ByVal sOperation As String, ByVal rstField As String, ByRef objVar As Object, ByVal abbr As String)
    Select Case sOperation
      Case "Update"
        If CStr("" & AppSettings.rst.Fields(rstField).Value) <> CStr(objVar) Then
          AppSettings.rst.Fields(rstField).Value = objVar
          gstrChanges = gstrChanges & abbr & ":" & CStr(objVar) & ","
          gbChangesMade = True
        Else
          gstrChanges = gstrChanges & ","
        End If
      Case "Add"
        AppSettings.rst.Fields(rstField).Value = objVar
        gstrChanges = gstrChanges & abbr & CStr(objVar) & ","
    End Select
  End Sub

  Public Function DBUpdate(ByRef sEditMode As String, ByRef objcol As Object, ByRef ent As Object) As Integer
    Dim nUpdate As Integer
    Dim bPromptConfirm As Boolean
    bPromptConfirm = AppSettings.GetSetting("bAskForUpdateConfirmation")
    If bPromptConfirm = True Then
      Dim ret = MsgBox("Are you sure you want to execute the " & sEditMode & " operation for " & TypeName(ent) & " " & ent.ID, MsgBoxStyle.YesNo)
      If ret = MsgBoxResult.No Then
        objcol.loadSingleEntity(ent.ID, "U")
        DBUpdate = False
        Exit Function
      End If
    End If
    Select Case sEditMode
      Case "Delete"
        nUpdate = objcol.DeleteFromDB(ent)
        'If nUpdate = 0 Then
        'MsgBox("Deletion of " & TypeName(ent) & " " & ent.ID & " Failed")
        'DBUpdate = 0
        'Exit Function
        'ElseIf nUpdate = 1 Then
        ' MsgBox("Successully deleted " & TypeName(ent) & " " & ent.ID)
        'End If
      Case "Add"
        nUpdate = objcol.AddtoDB(ent)
        'If nUpdate = 0 Then
        'MsgBox("A record has been added to the DB since you last refreshed. Please Load Data again")
        ' DBUpdate = 0
        ' Exit Function
        'End If
      Case "Update"
        nUpdate = objcol.UpdateDB(ent)
        'If nUpdate = 0 Then
        'MsgBox("This record (ID:" & ent.ID & ") has been updated since you last refreshed. Please load data again and then update.")
        'DBUpdate = 0
        'Exit Function   'db not updated because the record has been updated after last refresh.
        'ElseIf nUpdate = 1 Then

        'MsgBox(TypeName(ent) & " " & ent.id & " has been updated")
        'ElseIf nUpdate = -1 Then
        'MsgBox("Nothing to Update")
        'End If
      Case "Load"  'this loads and refreshes the entity from the db to the collection
        objcol.loadSingleEntity(ent.ID, "U")
        'DBUpdate = 1
    End Select
    RaiseEvent EntityUpdated(sEditMode, ent, nUpdate)
    DBUpdate = nUpdate
  End Function

  Public Sub DBLoad()
    'Instantiates and loads all entity collections in the correct order. Child entities need to be loaded before their parents.
    'Dont call this function directly, preferrably call via CallLoadDBIfNeeded(). This way it does not get
    'called unnecessarily when the data is already up to date.
    'Framework Gen App:
    'OMGlobalsM.InitApp()
    cUpdateLogItems = New UpdateLogItems
    MGlobals.cUpdateLogItems = cUpdateLogItems
    If AppSettings.cnn.State = ADODB.ObjectStateEnum.adStateOpen Then
      AppSettings.cnn.Close()
    End If

    Call LoadEntities()

    LoggedInUser = Nothing
    AppSettings.SetSetting("LastUpdateID", cUpdateLogItems.GetLastUpdateID(PASettings))
    'AppSettings.SetSetting("CurrPath", AppSettings.GetDLLPath)
    Call GetAccessRightForLoggedInUser()
  End Sub

  Public Function FillChildEntities(ByRef objEDI As Object, ByRef entPr As Object, ByRef entChild As Object, ByVal sChildPropertyName As String) As Boolean
    Return UIFunctions.FillChildEntities(objEDI, entPr, entChild, sChildPropertyName)
  End Function

  Public Sub FillAvailableChildEntities(ByRef objEntityDataItems As Object, ByRef entChldContainer As Object, ByRef entPr As Object, ByRef entChild As Object, ByVal sChildPropertyName As String)
    Call UIFunctions.FillAvailableChildEntities(objEntityDataItems, entChldContainer, entPr, entChild, sChildPropertyName)
  End Sub

  Public Sub FillAvailableChildEntitiesForCBO(ByRef objEntityDataItems As Object, ByRef entPr As Object, ByRef entChild As Object, ByVal sChildPropertyName As String)
    Call UIFunctions.FillAvailableChildEntitiesForCBO(objEntityDataItems, entPr, entChild, sChildPropertyName)
  End Sub

  Public Sub FillParentEntities(ByRef objEDI As Object, ByRef entParent As Object, ByRef entChild As Object, ByVal sChildPropertyName As String)
    Call UIFunctions.FillParentEntities(objEDI, entParent, entChild, sChildPropertyName)
  End Sub

  Public Sub FillAvailableParentEntities(ByRef objEDI As Object, ByRef entParent As Object, ByRef entChild As Object, ByVal sChildPropertyName As String)
    Call UIFunctions.FillAvailableParentEntities(objEDI, entParent, entChild, sChildPropertyName)
  End Sub

  Public Sub FillAvailableParentEntitiesForCBO(ByRef objEntityDataItems As Object, ByRef entPr As Object, ByRef entChild As Object, ByVal sChildPropertyName As String)
    Call UIFunctions.FillAvailableParentEntitiesForCBO(objEntityDataItems, entPr, entChild, sChildPropertyName)
  End Sub


  Public Function AddChildEntity(ByRef objEDI As Object, ByRef enParent As Object, ByRef enChld As Object, ByVal sChildPropertyName As String) As Integer
    Return BLFunctions.AddChildEntity(objEDI, enParent, enChld, sChildPropertyName)
  End Function

  Public Function RemoveChildEntity(ByRef objEDI As Object, ByRef enParent As Object, ByRef enChld As Object, ByVal sChildPropertyName As String) As Integer
    Return BLFunctions.RemoveChildEntity(objEDI, enParent, enChld, sChildPropertyName)
  End Function

  Public Function ReOrderChildEntities(ByRef enParent As Object, ByRef enChlds As Object, ByVal strChild As String) As Integer
    Return BLFunctions.ReOrderChildEntities(enParent, enChlds, strChild)
  End Function

  Public Function GetDBUpdates(ByRef objEDI As Object, ByVal nPrevUpdateID As Integer) As Integer
    Return MGlobals.GetDBUpdates(objEDI, nPrevUpdateID)
  End Function

  Public Function GetSetting(ByVal sKey As String) As String
    Return AppSettings.GetSetting(sKey)
  End Function
  Public Sub SetSetting(ByVal skey As String, ByVal sValue As String)
    Call AppSettings.SetSetting(skey, sValue)
  End Sub

  Public Sub LaunchSettingsDlg()

  End Sub

  Public Sub GenerateObjectModelCode(ByVal sPAProjectName As String)
    'Code for generating code
    '1. Create Entitys class collection
    PACode_EntityCollectionClass.GenerateEntityCollectionClass(sPAProjectName)
    '2. Create Entity base class
    PACode_EntityBaseClass.GenerateEntityBaseClass(sPAProjectName)
    '3. Create Entity class 
    PACode_EntityClass.GenerateEntityClass(sPAProjectName)
    '4 Create Global custom code
    PACode_Global.GenerateOMGlobalCustom(sPAProjectName)
    '5 Create Metadata custom code.
    PACode_Metadata.GenerateMetaDataCustom(sPAProjectName)
  End Sub


  Public Sub GenerateUIForExcel(ByVal sPAProjectName As String)
    Call PACode_ExcelUI.GenerateExcelUI(sPAProjectName)
  End Sub

  Public Sub CallDoSortOfAttendances(ByRef objCal As Calendars)
    Call DoSort1(objCal, AddressOf Calendars.CompareDate)
  End Sub

 
End Class
