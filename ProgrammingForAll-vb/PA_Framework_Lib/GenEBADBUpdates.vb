'Imports Microsoft.Office.Interop
Imports Microsoft.Win32

<ComClass(GenEBADBUpdates.ClassId, GenEBADBUpdates.InterfaceId, GenEBADBUpdates.EventsId)> _
Public Class GenEBADBUpdates

#Region "COM GUIDs"
  ' These  GUIDs provide the COM identity for this class 
  ' and its COM interfaces. If you change them, existing 
  ' clients will no longer be able to access the class.
  Public Const ClassId As String = "dc864e46-e104-41e6-ba7a-0c92a02ba1a9"
  Public Const InterfaceId As String = "21a6079c-7bda-4a43-8605-202d2176a9aa"
  Public Const EventsId As String = "4c8d8c01-95b7-4396-85ff-caddad3ffd67"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    'Public Shared ExApp As Excel.Application
    'Public Shared cnn As New ADODB.Connection
    'Public Shared rst As New ADODB.Recordset
    Private strConnTM As String, strDBO As String
  Private PrevUpdateID As Integer
  Private NewUpdateID As Integer
  Private Shared pwii As String = "TThhaabbttii"
  Private gDBLoaded As Boolean
  Private Shared bVal As Boolean

  Public Shared adOpenStatic = 3
  Public Shared adLockOptimistic = 3
  Public Shared adUseClient = 3
  Public Shared adCmdTable = 2
  Public Shared adOpenKeySet = 1
  Public frmIntroduction As frmIntro
  Friend Shared strLicenseKey As String
  Private bDontContinue As Boolean
  Private strLicencedTo As String
  Private sLicensee As String
  Public Sub New()
    MyBase.New()

  End Sub
  Public Sub Test()
    Dim ddatevalid1
    ddatevalid1 = DecodeDate("")

    'frmIntroduction.ShowDialog()
    'MsgBox(strLicenseKey)
    'MsgBox("geneba1226")
    'par1 = par1 + 20
  End Sub
  Private Function GetRemainingDays(ByVal sval As String)
    Dim dDateValid As Date
    Dim nDateDiff As Integer
    sval = sval.Substring(6)
    dDateValid = CDate(sval)
    nDateDiff = DateDiff(DateInterval.Day, dDateValid, Today)
    Return (31 - nDateDiff)

  End Function

  Public Function Initialize(ByRef ex As Object) As Boolean
    Dim sVal As String = ""
    'Dim dDateValid As Date
    Dim strFirstEval As String
    'Dim nDateDiff As Integer
    Dim nRemainingEvalDays As Integer
    Dim frmIntroduction As New frmIntro
    Initialize = False
        'ExApp = ex
        'strFirstEval = ExApp.Run("GetExcelPar", "FirstEval")
        GetFromReg("lic", sVal)
    If sVal <> "NewInstall" Then sVal = Decode(sVal)

        'If (sVal = "NewInstall" And strFirstEval = "Eval") Then
        '  Call ExApp.Run("SetExcelPar", "FirstEval", "")
        '  sVal = "Trial:" & Today
        '  Dim wrapper As New Simple3Des(pwii)
        '  Dim cipherText As String = wrapper.EncryptData(sVal)
        '  SaveinReg("lic", cipherText)
        '  sLicensee = "Trial User"
        'End If
        If sVal.Contains("Trial:") Then
      nRemainingEvalDays = GetRemainingDays(sVal)
      If nRemainingEvalDays >= 0 Then
        frmIntroduction.txtLicensedTo.Text = "Trial User"
        frmIntroduction.txtLicenseStatus.ForeColor = Drawing.Color.Black
        frmIntroduction.txtLicenseStatus.Text = CStr(nRemainingEvalDays) & " Days remaining"
      Else
        frmIntroduction.txtLicensedTo.Text = "NA"
        frmIntroduction.txtLicenseStatus.ForeColor = Drawing.Color.Red
        frmIntroduction.txtLicenseStatus.Text = "Trial Period Expired"
        sLicensee = "Invalid"
      End If
    ElseIf sVal = "Error" Then
      sLicensee = "Invalid"
      frmIntroduction.txtLicensedTo.Text = sLicensee
      frmIntroduction.txtLicenseStatus.ForeColor = Drawing.Color.Red
      frmIntroduction.txtLicenseStatus.Text = "Invalid"
    ElseIf sVal.Contains("Valid:") Then
      sLicensee = "Valid"
      frmIntroduction.txtLicensedTo.Text = sVal.Substring(6)
      frmIntroduction.txtLicenseStatus.ForeColor = Drawing.Color.Blue
      frmIntroduction.txtLicenseStatus.Text = "Valid"
    End If
    'Dim sval1 As String = ""
    'GetFromReg("lic", sval1)
    'sLicensee = Decode(sval1)
    'If sLicensee.Contains("Trial:") Then  'if trial
    'Else  'if other than trial
    '    'sLicensee = Decode(sval1)
    '    If sLicensee = "Error" Then
    '        frmIntroduction.txtLicensedTo.Text = "Invalid"
    '        frmIntroduction.txtLicenseStatus.ForeColor = Drawing.Color.Red
    '        frmIntroduction.txtLicenseStatus.Text = "Invalid"
    '    Else
    '        frmIntroduction.txtLicensedTo.Text = sLicensee
    '        frmIntroduction.txtLicenseStatus.ForeColor = Drawing.Color.Blue
    '        frmIntroduction.txtLicenseStatus.Text = "Valid"
    '    End If
    'End If
    frmIntroduction.Show()
  End Function

  Public Sub Message()
    Dim ab As String
        'ab = ExApp.Range("LoginNumber").ToString
        'MsgBox(ab)
    End Sub
  'Private Function EncodeDate(ByVal dDate As Date) As String
  '    Dim sD As Integer
  '    Dim sM As Integer
  '    Dim sY As Integer
  '    sM = Month(dDate) * 9
  '    sD = Day(dDate) * 55
  '    sY = Year(dDate) * 8
  '    EncodeDate = sM.ToString & "." & sD.ToString & "." & sY.ToString
  'End Function
  Private Function DecodeDate(ByVal sCode As String) As Date
    Dim sArr() As String
    Dim nDay As Integer, nMonth As Integer, nYear As Integer
    sArr = Split(sCode, ".")
    Try
      nMonth = CInt(sArr(0)) / 9
      nDay = CInt(sArr(1)) / 55
      nYear = CInt(sArr(2)) / 8
      DecodeDate = nMonth.ToString & "/" & nDay.ToString & "/" & nYear.ToString
      If IsDate(DecodeDate) Then
        'MsgBox("decodedate is date")
        Return DecodeDate
      Else
        'MsgBox("decodedate is not a date")
        Return Today.AddMonths(-1)
      End If
    Catch e As InvalidCastException
      Return Today.AddMonths(-2)
    End Try
  End Function

  Public Sub SaveinReg(ByVal sKey As String, ByRef sValue As String)
    Dim sReg As RegistryKey
    If (Registry.CurrentUser.OpenSubKey("SOFTWARE\\GenEBA\\", False)) Is Nothing Then
      sReg = Registry.CurrentUser.CreateSubKey("SOFTWARE\\GenEBA\\Options")
    End If
    sReg = Registry.CurrentUser.OpenSubKey("SOFTWARE\\GenEBA\\Options", True)
    sReg.SetValue(sKey, sValue)
  End Sub
  Public Sub GetFromReg(ByVal sKey As String, ByRef sValue As String)
    Dim sReg As RegistryKey
    If (Registry.CurrentUser.OpenSubKey("SOFTWARE\\GenEBA\\", False)) Is Nothing Then
      sReg = Registry.CurrentUser.CreateSubKey("SOFTWARE\\GenEBA\\Options")
    End If

    sReg = Registry.CurrentUser.OpenSubKey("SOFTWARE\\GenEBA\\Options", True)
    sValue = sReg.GetValue(sKey, "NewInstall")
  End Sub

  'Public Shared Sub ValidateGeneba(ByVal strLicenseKey As String)
  '    Call ExApp.Run("SetExcelPar", "ValidateApp", strLicenseKey)  'in case user has entered a key
  '    bVal = ValidateApp()

  'End Sub
  Public Function Encode(ByVal strText As String) As String

    Dim wrapper As New Simple3Des(pwii)
    Dim cipherText As String = wrapper.EncryptData(strText)

    'MsgBox("The cipher text is: " & cipherText)
    'SaveinReg("lic", cipherText)
    Return cipherText
  End Function

  Public Shared Function Decode(ByVal CipherText As String) As String
    'Dim cipherText As String = GetFromReg("lic", ciphertext)
    'Dim password As String = InputBox("Enter the password:")
    Dim wrapper As New Simple3Des(pwii)

    ' DecryptData throws if the wrong password is used.
    Try
      Dim plainText As String = wrapper.DecryptData(CipherText)
      'MsgBox("The plain text is: " & plainText)
      Return plainText
    Catch ex As System.Security.Cryptography.CryptographicException
      Return "Error"
    End Try
  End Function
End Class


'Public Function Initialize1(ByRef ex As Excel.Application) As Boolean
'    Dim sDate As String
'    Dim sCodedDate As String
'    Dim dDateValid As Date
'    Dim nDateDiff As Integer
'    Dim strFirstEval As String
'    'Dim strFirstEvalfromReg As String
'    Dim sCodedDateFromReg As String = ""
'    Dim nRemainingEvalDays As Integer
'    'Dim bDate As Boolean
'    'bDate = ValidateApp()
'    'sCodedDate = EncodeDate("8/2/2010")
'    'dDateValid = DecodeDate(sCodedDate)
'    Initialize1 = True

'    ExApp = ex
'    frmIntroduction = New frmIntro
'    bVal = ValidateApp()
'    strFirstEval = ExApp.Run("GetExcelPar", "FirstEval")
'    'Me.GetFromReg("Eval", strFirstEvalfromReg)

'    'MsgBox(strFirstEval)
'    If strFirstEval = "Eval" Then
'        sDate = EncodeDate(Today)
'        Call ExApp.Run("SetExcelPar", "FirstEval", "")

'        Me.SaveinReg("Eval", "E")
'        Call ExApp.Run("SetExcelPar", "DateValid", sDate)
'        Me.SaveinReg("Date", sDate)
'    End If

'    sCodedDate = ExApp.Run("GetExcelPar", "DateValid")
'    Me.GetFromReg("Date", sCodedDateFromReg)
'    If sCodedDate = sCodedDateFromReg Then
'        dDateValid = DecodeDate(sCodedDate)
'    End If
'    'MsgBox(dDateValid)

'    'bDate = IsDate(dDateValid)
'    nDateDiff = DateDiff(DateInterval.Day, dDateValid, Today)
'    If bVal = True Then
'        frmIntroduction.txtLicenseStatus.Text = "Valid"
'        frmIntroduction.txtLicensedTo.Text = strLicencedTo
'        frmIntroduction.txtLicenseFile.Enabled = True
'    Else
'        nRemainingEvalDays = (31 - nDateDiff)
'        If nRemainingEvalDays >= 0 Then
'            frmIntroduction.txtLicenseStatus.Text = CStr(nRemainingEvalDays) & " Days remaining"
'        Else
'            frmIntroduction.txtLicensedTo.Text = "NA"
'            frmIntroduction.txtLicenseStatus.ForeColor = Drawing.Color.Red
'            frmIntroduction.txtLicenseStatus.Text = "Expired"
'        End If
'        frmIntroduction.txtLicenseFile.Enabled = True
'    End If
'    strLicenseKey = ExApp.Run("GetExcelPar", "ValidateApp")
'    frmIntroduction.Show()
'    strLicenseKey = frmIntroduction.txtLicenseFile.Text
'    If bVal = True Then
'        Exit Function
'    End If
'    'If strLicenseKey <> "" Then
'    '    Call ExApp.Run("SetExcelPar", "ValidateApp", strLicenseKey)
'    '    MsgBox("Please Launch GenEBA again")
'    '    Initialize = False
'    '    Exit Function
'    'End If
'    bDontContinue = nDateDiff > 31 And bVal = False
'    If bDontContinue Then
'        'MsgBox("Evaluation Period Expired" & vbCrLf & "GenEBA will now close")
'        Initialize1 = False
'        'ExApp.Quit()
'        'Exit Function
'    End If
'End Function
'Public Shared Function ValidateApp() As Boolean
'    Dim sRet As String, sCode As Integer
'    Dim sVali() As String
'    sRet = ExApp.Run("GetExcelPar", "ValidateApp")

'    sVali = Split(sRet, "-")
'    If sVali.Length <> 4 Then
'        ValidateApp = False
'        Exit Function
'    End If
'    sCode = CInt(sVali(0)) + CInt(sVali(1)) + CInt(sVali(2)) + CInt(sVali(3))
'    If sCode = 10890 Then
'        ValidateApp = True
'    Else
'        ValidateApp = False
'    End If
'End Function
'Public Function GetTMConnString(ByRef strConnTM, ByRef strDBO) As Boolean
'    'Gets the connection string for the database based on settings.
'    Dim sTMSvr As String
'    Dim sTMDB As String
'    Dim strTMUser As String
'    Dim strTMpw As String

'    strTMUser = "TMAS"
'    strTMpw = "TMAS"
'    Dim bUseMDBDatabase As Boolean
'    Dim SQLServerName As String
'    Dim SQLServerDatabase As String
'    Dim MDBDatabasePath As String
'    bUseMDBDatabase = ExApp.Run("GetExcelPar", "bUseMDBDatabase")
'    SQLServerName = ExApp.Run("GetExcelPar", "SQLServerName")
'    MDBDatabasePath = ExApp.Run("GetExcelPar", "MDBDatabasePath")
'    SQLServerDatabase = ExApp.Run("GetExcelPar", "SQLServerDatabase")

'    If bUseMDBDatabase = False Then   'use SQL Server database
'        sTMSvr = SQLServerName
'        sTMDB = SQLServerDatabase

'        strConnTM = "Provider=sqloledb;" & _
'               "Data Source= " & sTMSvr & ";" & _
'               "Initial Catalog= " & sTMDB & ";" & _
'               "User Id=" & strTMUser & ";" & _
'               "Password=" & strTMpw
'        strDBO = ""
'    Else    'use mdb database
'        strConnTM = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'               "Data Source= " & MDBDatabasePath & ";" & _
'               "Jet OLEDB:Database Password=tmas"
'        strDBO = ""
'    End If
'    GetTMConnString = True
'End Function

'Public Function DBUpdateNeeded() As Integer
'    'Call GenCom.Initialize(ThisWorkbook.Application)
'    'Determines if update from DB is needed based on the following.
'    '1. Get the last update ID from Updatelog table in DB and assign it to 'NewUpdateID'
'    '2. Saves frmSettings.LastUpdateID as 'PrevUpdateID'
'    '3. save ID from (1) as frmSettings.LastDBUpdateID
'    '4. Returns the number of updates done to db since DBUpdateNeeded was last called.
'    ' this is essentially NewUpdateId - PrevUpdateID
'    Dim strSQL As String
'    'Set cnn = CreateObject("ADODB.Connection")
'    'Set rst = CreateObject("ADODB.Recordset")
'    Call GetTMConnString(strConnTM, strDBO)
'    strSQL = "select max(ID) as MaxNo from updatelog"
'    cnn.Open(strConnTM)
'    rst.CursorLocation = adUseClient
'    rst.Open(strSQL, cnn, adOpenStatic, adLockOptimistic)
'    'Now set the two global variables
'    NewUpdateID = rst.Fields("MaxNo").Value
'    PrevUpdateID = ExApp.Run("GetExcelPar", "LastUpdateID")
'    'PrevUpdateID = frmSettings.LastUpdateID
'    rst.Close()
'    Call ExApp.Run("SetExcelPar", "LastUpdateID", NewUpdateID)
'    'frmSettings.LastUpdateID = NewUpdateID
'    'save it in settings
'    DBUpdateNeeded = NewUpdateID - PrevUpdateID
'    cnn.Close()
'End Function

'Public Function CallDBLoadIfNeeded(ByVal ex As Excel.Application) As Boolean
'    'Returns TRUE if DBLoad is called by this function, else returns FALASE
'    Dim nUpdate As Integer
'    ExApp = ex
'    CallDBLoadIfNeeded = False
'    gDBLoaded = ExApp.Run("GetExcelPar", "gDBLoaded")
'    nUpdate = DBUpdateNeeded() 'returns the number of db updates made since last called
'    If gDBLoaded <> True Or nUpdate <> 0 Then
'        'Call GenCom.Initialize(ThisWorkbook.Application)
'        Call ExApp.Run("DBLoad")
'        'Call Globals.DBLoad()
'        CallDBLoadIfNeeded = True
'    End If
'    Dim randValue As Integer = CInt(Int((6 * Rnd()) + 1))
'    If sLicensee.ToLower = "invalid" And randValue > 5 Then
'        Call Me.Initialize(ex)
'        'MsgBox("No valid license license Available")
'        'ExApp.Quit()
'    End If
'End Function

'Public Sub UpdateLogTable(ByVal cnn As Object, ByVal rst As ADODB.Recordset, ByVal nTableID As Integer, ByVal sOperation As String, ByVal recID As String, ByVal strChanges As String)
'    'updates the log table 'UpdateLog' for recording history of changes made  by users.
'    'nTableID=1 for members, 2 for meetings, 3 for speeches
'    'sOperation="U" for update of existing entity, "N" for creation of new entity.
'    'recID= ID of the record in the table that was updated or created
'    'In addition to the above info, it also logs the current time and the user ID of the user who made the change.
'    rst.Open("UpdateLog", cnn, , , adCmdTable)
'    rst.AddNew()
'    rst("DateTime").Value = Now
'    rst("LoginID").Value = ExApp.Run("GetExcelPar", "LoggedInMember")  'Globals.LoggedInMember.MemID    'frmSettings.LoginName  'strLoginID
'    rst("TableID").Value = nTableID  'id of the table  1:members, 2:meetings, 3: Speeches
'    rst("KeyFieldNo").Value = recID
'    rst("Changes").Value = strChanges
'    rst("Operation").Value = sOperation  '
'    rst.Update()
'    strChanges = ""
'    rst.Close()

'    Dim randValue As Integer = CInt(Int((6 * Rnd()) + 1))
'    If sLicensee = "InValid" And randValue > 3 Then
'        Call Me.Initialize(ExApp)
'        'MsgBox("No valid license license Available")
'        'ExApp.Quit()
'    End If
'End Sub

'    'Private Sub WriteToUserLog(ByVal logValue)  'cleanupthe code. shah
'    'Input parameter: LogValue="Login" when called from wkbk_open, or "Logout" when called from wkbk_beforeClose.
'    'Writes user log with details of who logged in at what time and from what machine. When called from wks_beforeClose, also logs the
'    'logout time with its login number of the user. This can be used to determine duration of the system usage by users.
'    'Another functionality is to control the versions which can be used by the users. You may provide only a message to user if they are
'    'using an older version, or may not let the user to login if they are using a certain version.

'    'If this function is deemed to be not necessary, do  not call from wkbk_open and wkbk_beforeClose'
'    Dim bRecordAdded As String
'    Dim nFlag As Integer
'    Dim strMessage As String
'    Dim cnn1 As Object
'    Dim rstReport As Object
'        cnn1 = CreateObject("ADODB.Connection")
'        rstReport = CreateObject("ADODB.Recordset")
'    Dim strCnn As String
'    Dim strID As String
'    Dim strFirstName As String
'    Dim strLastName As String
'    Dim strETASDBPath As String

'    'ExApp.Run("GetExcelPar", "LastUpdateID")
'    'ExApp.Run("SetExcelPar", "LastUpdateID", NewUpdateID)

'        On Error GoTo Handler
'        If logValue = "LogIn" Then Call ExApp.Run("SetExcelPar", "LoginNumber", 0) 'frmSettings.LoginNumber = 0

'    'Get TMAS connection string
'        Call GenCom.GetTMConnString(strConnTMAS, strDBO)
'        bRecordAdded = False
'        cnn1.Open(strConnTMAS)
'        rstReport.CursorType = adOpenKeySet
'        rstReport.LockType = adLockOptimistic
'        rstReport.Open("UserLogs", cnn1, , , adCmdTable)
'    'create new record
'        rstReport.AddNew()
'        rstReport!UserName = frmSettings.LoginName
'        rstReport!ComputerName = frmSettings.ComputerName()
'        rstReport!DateTime = Now
'        rstReport!LoginLogout = logValue
'        rstReport!Version = Range("APPVersion")
'        rstReport!Build = Range("APPBuild")
'        rstReport!LoginNo = frmSettings.LoginNumber
'        rstReport.Update()
'        bRecordAdded = True
'        rstReport.Close()
'        cnn1.Close()
'        cnn1.Open()
'        strAppVerBld = Trim(Range("APPVersion") & Range("APPBuild"))
'    Dim strSQL As String

'        strSQL = "SELECT flag, Message FROM versioncontrol WHERE Version like '" & strAppVerBld & "'" & _
'                 " ORDER BY sNo ASC"
'        rstReport.Open(strSQL, cnn1, adOpenStatic, adLockOptimistic)
'        If rstReport.RecordCount > 0 Then
'            rstReport.MoveFirst()
'            nFlag = rstReport.Fields("flag")
'            strMessage = rstReport.Fields("Message")
'        End If
'        rstReport.Close()
'    'In case the user is to be given a message only Flag = 1
'    'If the user is to be given a message and not allowed to log, Flag is 2
'        If logValue = "LogIn" Then
'            rstReport.Open("SELECT MAX(SNo) AS MaxsNo FROM UserLogs  ", _
'                cnn1, adOpenStatic, 1)
'            If rstReport.RecordCount > 0 Then
'                rstReport.MoveFirst()
'                frmSettings.LoginNumber = rstReport.Fields("MaxsNo")
'            End If
'            rstReport.Close()
'            cnn1.Close()
'            Select Case nFlag
'                Case 1
'                    MsgBox(strMessage)
'                Case 2    'Application close not working
'                    MsgBox(strMessage)
'                    Me.Close(SaveChanges:=False)
'                    Exit Sub
'            End Select
'        Else
'            cnn1.Close()
'        End If
'        Exit Sub
'Handler:
'        If logValue = "LogIn" Then
'            MsgBox("           Application is not able to access the specified database in 'Settings'" & vbCrLf & _
'                   " ", vbOKOnly, " & AppName & " & strAppVerBld)
'        End If
'    End Sub
