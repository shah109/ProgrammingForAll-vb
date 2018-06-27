Option Explicit On
Imports System.IO
Imports System.Text
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class PACode_EntityCollectionClass
  Public Shared Function GenerateEntityCollectionClass(ByVal sProjectName As String) As String

    'Dim pae As ProjectEntity
    Dim oPAProject As New PAProject
    Dim oPAEntity As New ProjectEntity
    Dim oPAEntityItem As ProjectEntityItem = Nothing
    GenerateEntityCollectionClass = String.Empty
    Dim sCode As String = sProjectName


    'get the project name
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Then
        oPAProject = pap
        Exit For
      End If

    Next
    'Dim fs1 As New FileStream
    'get the entity name 
    For Each pae As ProjectEntity In oPAProject.ChildprojectEntities
      'If pae.EntityCollectionName = "Customers" Then
      oPAEntity = pae
      'End If
      'Exit For


      Dim fs1 As FileStream = New FileStream("c:\Framework\" & oPAEntity.EntityCollectionName & ".vb", FileMode.Create, FileAccess.Write)
      Dim s1 As StreamWriter = New StreamWriter(fs1)
      s1.WriteLine("'PA_Code Generator Code")
      s1.WriteLine("'" & Now & ":, " & sCode)
      s1.WriteLine()

      s1.WriteLine("Option Explicit On")
      s1.WriteLine("Imports PA_Framework_Lib")
      s1.WriteLine("Imports PA_Framework_OM.OMGlobals")
      s1.WriteLine("Imports System.Runtime.InteropServices")

      s1.WriteLine()
      s1.WriteLine("<ClassInterface(ClassInterfaceType.AutoDual)> _")
      s1.WriteLine("Public Class " & oPAEntity.EntityCollectionName)
      s1.WriteLine("  Inherits PAEnts")
      s1.WriteLine()


      s1.WriteLine("  Overrides Sub LoadEntityItemsForThisEntity(ByRef ent As Object)")
      For Each oPAEntityItem In oPAEntity.mChildProjectEntityItems
        If oPAEntityItem.ChildName = String.Empty Then
          '    ent.ID = "" & AppSettings.rst.Fields("ID").Value
          s1.WriteLine("    ent.{0:C} = """" & AppSettings.rst.Fields(""{1:C}"").Value", _
                       oPAEntityItem.PropertyName, oPAEntityItem.DBColName)
        End If
      Next
      s1.WriteLine("")
      For Each oPAEntityItem In oPAEntity.mChildProjectEntityItems
        If oPAEntityItem.ChildName <> String.Empty And oPAEntityItem.RelationshipType = "CSR" Then
          s1.WriteLine("    Call BLFunctions.BuildChildEntityObjects(cPAProjects, ent, ""{0:C}"", """" & AppSettings.rst.Fields(""{1:C}"").Value)", oPAEntityItem.PropertyName, oPAEntityItem.DBColName)
        End If
      Next
      s1.WriteLine("  End Sub")
      s1.WriteLine("")

      s1.WriteLine("  Overrides Sub RecordChanges(ByVal sOperation As String, ByRef ent As Object)")
      'For Each oPAEntity In oPAProject.ChildprojectEntities
      For Each oPAEntityItem In oPAEntity.mChildProjectEntityItems
        If oPAEntityItem.ChildName = String.Empty And Trim(oPAEntityItem.PropertyName) <> "ID" Then
          s1.WriteLine("    Call gRecordChanges(sOperation, ""{0}"", ent.{1}, ""{2}"")", _
                       oPAEntityItem.DBColName, oPAEntityItem.PropertyName, oPAEntityItem.ULName)
        End If
      Next
      s1.WriteLine("")
      For Each oPAEntityItem In oPAEntity.mChildProjectEntityItems
        If oPAEntityItem.ChildName <> String.Empty And Trim(oPAEntityItem.RelationshipType) = "CSR" Then
          s1.WriteLine("    Call gRecordChanges(sOperation, ""{0:C}"", ent.ChildEntityString(""{1:C}""), ""{2:C}"")", _
                       oPAEntityItem.DBColName, oPAEntityItem.PropertyName, oPAEntityItem.ULName)
        End If
      Next
      s1.WriteLine("  End Sub")
      s1.WriteLine("")



      s1.WriteLine("  Overrides Function GetSelectDBItems() As String")
      s1.WriteLine("    GetSelectDBItems = _ ")
      Dim strSB As New StringBuilder
      'Dim sLen As Integer

      For Each oPAEntityItem In oPAEntity.mChildProjectEntityItems
        strSB.AppendFormat("      strDBO & "", {0:C} "" & _" & vbCrLf, oPAEntityItem.DBColName)
      Next
      strSB.AppendFormat("      strDBO & "", LastUpdate """)
      strSB.Remove(16, 1)  '16th chr removes the (') on the first ID line.
      'sLen = strSB.Length
      'strSB.Remove(sLen - 5, 5)   ' remove the last '& _'
      s1.Write(strSB)
      s1.WriteLine()
      s1.WriteLine("  End Function")

      s1.WriteLine("")
      s1.WriteLine("")
      s1.WriteLine()
      s1.WriteLine("  Overrides Function GetSelectFromTable() as String")
      'GetSelectFromTable = " FROM " & strDBO & "Courses "
      s1.WriteLine("    GetSelectFromTable = "" FROM "" & strDBO & "" {0:C} """, oPAEntity.EntityCollectionName)
      s1.WriteLine("  End Function")
      s1.WriteLine()
      s1.WriteLine("  Overrides Function CreateNewEntity() As PAEnt")
      s1.WriteLine("    CreateNewEntity = New {0:C} ", oPAEntity.EntityName)
      s1.WriteLine("  End Function")
      s1.WriteLine()
      s1.WriteLine("  Overrides Sub SetCurrentEntity(ByRef ent As PAEnt)")
      s1.WriteLine("    curr{0:C} = ent", oPAEntity.EntityShortName)
      s1.WriteLine("  End Sub")
      s1.WriteLine()
      s1.WriteLine("  Overrides Sub BuildDomainModel()")
      s1.WriteLine("  End Sub")
      s1.WriteLine(" End Class")

      s1.Close()
      fs1.Close()
      s1 = Nothing
      fs1 = Nothing
    Next

  End Function

  Function GenerateEntityClass()
    GenerateEntityClass = Nothing

  End Function
  Function GenerateMetaDataFile() As String
    GenerateMetaDataFile = String.Empty

  End Function

  Function GenerateOMGlobalsCustom() As String
    GenerateOMGlobalsCustom = String.Empty

  End Function


End Class
'Solution for tlb registration
'1)     Open a regular command prompt (NOT elevated)
'2)     Locate regcap.exe (it’s typically located in \program files\microsoft visual studio 9.0\common7\tools\deployment\ - replace “program files” with “program files (x86)” on 64-bit systems), and add this directory to your path (for example, run “set PATH=%PATH%;c:\program files\microsoft visual studio 9.0\common7\tools\deployment\”).
'3)     Run the following command: “regcap /I /O outputfile.reg inputfile.tlb”, replacing “outputfile.reg” with a .reg filename in a location you have write access to, and “inputfile.tlb” with the TLB you’re trying to create registration information for
'4)     Locate the .reg file you created in step #3, and import it into your setup project through the registry editor (to find the option to do this, right-click on the root node “registry on the target machine” and choose “Import…”)
'5)     In the “detected dependencies” folder, select your TLB, and hit F4 to see the properties grid.
'6)     In the properties grid, change the “Register” field to be “vsdrfDoNotRegister” (this will eliminate the build warning you were seeing previously) 