Imports PA_Framework_Lib
Imports System.Text
Imports System.Data
Imports PA_Framework_OM.OMGlobals
Public Class PA_DBCreation
  'Public Shared nPrevUpdateID As Integer
  Public Shared objMyArrayList As New ArrayList
  Structure stTable
    Dim name As String
  End Structure

  Public Shared Function ValidateTablesAndColumns(ByVal sProjectName As String) As String
    Dim oPAProject As New PAProject
    'Description: 
    'for each entity item of each entity in sProjectName, checks whether a column is present in the related table for the entity. 
    'If present, it labels the output line as 'OK', otherwise it labels it as 'Not OK'.
    'Calls function ColumnExists to check each entity item against against the column for the entity.
    Dim sMissingColumns As String = String.Empty
    Dim sTotalMissingCols As String = String.Empty
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Then
        oPAProject = pap  'get the project for which validation is needed
      End If
      Exit For
    Next
    For Each pae As ProjectEntity In oPAProject.ChildprojectEntities
      For Each paei As ProjectEntityItem In pae.mChildProjectEntityItems
        sMissingColumns = ColumnExists(pae.EntityDBTableName, paei.DBColName, paei.DBColNameType, False)
        If sMissingColumns.Contains("Not OK") Then
        End If
        sTotalMissingCols = sTotalMissingCols & sMissingColumns
      Next paei
      's1.WriteLine("  Public {0:C} As {1:C}", pae.EntityShortName, pae.EntityName)
    Next pae
    Return sTotalMissingCols
  End Function

  Public Shared Function TableExists(ByVal sWhichTable As String) As Boolean
    TableExists = False
    Dim sMissingFields As String = String.Empty
    Dim s1 As String
    Dim s2 As String
    If AppSettings.cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      AppSettings.cnn.Open(AppSettings.GetAppConnString)
    End If

    AppSettings.rst = AppSettings.cnn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)
    AppSettings.rst.MoveFirst()
    Do Until AppSettings.rst.EOF
      s1 = LCase(AppSettings.rst.Fields("Table_Name").Value)
      s2 = LCase(sWhichTable)
      If s1 = s2 Then
        TableExists = True
        Exit Function
      End If '
      AppSettings.rst.MoveNext()
    Loop
  End Function

  Public Shared Function ColumnExists(ByVal sWhichTable As String, ByVal sWhichColumn As String, ByVal sColumnType As String, ByVal bModify As Boolean) As String
    Dim sMissingFields As New StringBuilder
    Dim sT1 As String, sC1 As String
    Dim sT2 As String, sC2 As String
    Dim tbl As New stTable
    objMyArrayList.Clear()
    If AppSettings.cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      AppSettings.cnn.Open(AppSettings.GetAppConnString)
    End If
    AppSettings.rst = AppSettings.cnn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)
    AppSettings.rst.MoveFirst()
    Do Until AppSettings.rst.EOF
      sT1 = LCase(AppSettings.rst.Fields("Table_Name").Value)
      sT2 = LCase(sWhichTable)
      sC1 = LCase(AppSettings.rst.Fields("Column_Name").Value)
      sC2 = LCase(sWhichColumn)
      tbl.name = sT1
      objMyArrayList.Add(tbl.name)

      If sT1 = sT2 And sC1 = sC2 Then

        Return sMissingFields.AppendFormat("OK,{0:C} , {1:C} , {2:C}" & vbCrLf, sWhichTable, sWhichColumn, sColumnType).ToString() 'column exists hence OK

        Exit Function
      End If

      AppSettings.rst.MoveNext()
    Loop
    sMissingFields.AppendFormat("NOT OK, {0:C} , {1:C} , {2:C}" & vbCrLf, sWhichTable, sWhichColumn, sColumnType)  'a column does not exist in the table for the entity.

    Return sMissingFields.ToString
  End Function

  Public Sub UpdateDBField()
    If AppSettings.cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      AppSettings.cnn.Open(AppSettings.GetAppConnString)
    End If
  End Sub

  Public Shared Sub CreateEntityItemsFromDatabase(ByRef nPrevUpdateID)
    Dim omg As New OMGlobals
    Dim cnn1 As New ADODB.Connection
    Dim rst1 As New ADODB.Recordset
    Dim rst2 As New ADODB.Recordset
    Dim strColumnName As String, strTableName As String
    Dim objEnt As ProjectEntity
    Dim objEntItm As ProjectEntityItem
    'Dim objEntCol As New ProjectEntities
    'Dim objEntItmCol As New ProjectEntityItems
    If cnn1.State = ADODB.ObjectStateEnum.adStateClosed Then
      cnn1.Open(AppSettings.GetAppConnString)
    End If

    'Dim n As Integer = 1
    'If cnn1.State = ADODB.ObjectStateEnum.adStateClosed Then
    '  cnn1.Open(AppSettings.GetAppConnString)
    'End If
    rst1 = cnn1.OpenSchema(ADODB.SchemaEnum.adSchemaTables)
    rst2 = cnn1.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)
    Dim oEnt As ProjectEntity
    Do Until rst1.EOF
      strTableName = rst1.Fields("TABLE_NAME").Value
      For Each oEnt In cProjectEntities
        If LCase(oEnt.EntityDBTableName) = LCase(strTableName) Then GoTo NextProjectEntity 'do not create duplicate entitynames in the db.
      Next
      objEnt = New ProjectEntity
      objEnt.EntityDBTableName = strTableName
      omg.DBUpdate("Add", cProjectEntities, objEnt)
      rst2.MoveFirst()
      Do Until rst2.EOF
        objEntItm = New ProjectEntityItem

        strColumnName = rst2.Fields("Column_Name").Value
        strTableName = rst2.Fields("Table_Name").Value

        If strTableName = objEnt.EntityDBTableName Then
          objEntItm.DBColName = strColumnName
          objEntItm.DBColNameType = rst2.Fields("DATA_TYPE").Value
          omg.DBUpdate("Add", cProjectEntityItems, objEntItm)
          omg.AddChildEntity(cPAProjects, objEnt, objEntItm, "ProjectEntityItems")
          nPrevUpdateID = omg.GetSetting("LastUpdateID")
        End If
        rst2.MoveNext()
      Loop


      'n = n + 1
NextProjectEntity:

      rst1.MoveNext()
    Loop

  End Sub

End Class
