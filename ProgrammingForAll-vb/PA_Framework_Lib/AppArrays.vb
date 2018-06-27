Public Class AppArray
  Public nID As Integer
  Public sAppArrayName As String
  Public sAppArrayString As String  '
  Public sAppArrayItems() As String

 
End Class

Public Class AppArrays
  Inherits System.Collections.CollectionBase

  Public Sub Add(ByRef val As AppArray)
    Me.List.Add(val)
  End Sub


  Public Sub LoadAppArrays()
    'Dim cnn As New OleDb.OleDbConnection
    'Dim cmd As New OleDb.OleDbCommand
    Dim eAppArr_ As AppArray
    Dim strSQLCall As String
    strSQLCall = " Select " & _
                  DBStrings.strDBO & "ID as sNo, " & _
                  DBStrings.strDBO & "ArrayName as sArrayName, " & _
                  DBStrings.strDBO & "ArrayString as sArrayString " & _
                  "FROM " & DBStrings.strDBO & "AppArrays " & _
                  " ORDER BY " & DBStrings.strDBO & "ArrayName asc"
    'DBStrings.gConnectionString = MGlobals.GetAppConnString
    'DBStrings.cnn1.ConnectionString = DBStrings.gConnectionString
    DBStrings.cmd1.CommandText = strSQLCall
    DBStrings.cmd1.CommandType = CommandType.Text
    DBStrings.cmd1.Connection = DBStrings.cnn1
    DBStrings.cnn1.Open()
    DBStrings.dr1 = DBStrings.cmd1.ExecuteReader

    Me.Clear()
    'Call GetAppConnString


    If Not DBStrings.dr1.HasRows Then Exit Sub
        'rst.MoveFirst()
        Dim n As Integer
        Dim m As Integer
        n = 0
        m = 0
        'Dim log As HistoryLog
        Do While DBStrings.dr1.Read
            'ReDim Preserve sApparrayItems(n + 1)
            eAppArr_ = New AppArray
            eAppArr_.nID = DBStrings.dr1.Item("sNo")
            eAppArr_.sAppArrayName = DBStrings.dr1.Item("sArrayName")
            eAppArr_.sAppArrayString = DBStrings.dr1.Item("sArrayString")

            'ReDim Preserve AppItems(n).ArrayItems(m + 1)
            eAppArr_.sAppArrayItems = Split(eAppArr_.sAppArrayString, ";")
            'sArray = Split(AppItems(n).ArrayString, ";")
            'm = UBound(sArray)
            'For i = 1 To m - 1
            '  ReDim Preserve AppItems(n).ArrayItems(i)
            '  AppItems(n).ArrayItems(i) = sArray(i)
            'Next i
            Me.Add(eAppArr_)
            'n = n + 1
            'rst.MoveNext()
        Loop
    DBStrings.cmd1.Dispose()
    DBStrings.cnn1.Close()
  End Sub

  Function GetArrayItems(ByVal sName As String) As String()
    Dim oAppArr As AppArray
    For Each oAppArr In Me.List
      If oAppArr.sAppArrayName = sName Then
        GetArrayItems = oAppArr.sAppArrayItems
        Exit Function
      End If
    Next oAppArr
    GetArrayItems = Nothing
  End Function

End Class
