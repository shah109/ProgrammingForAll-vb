Option Explicit On
Imports System.IO
Imports System.Text
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class PACode_Metadata
  Public Shared Function GenerateMetaDataCustom(ByVal sProjectName As String) As String
    'Creates all global declarations

    Dim oPAProject As New PAProject
    Dim oPAEntity As New ProjectEntity
    Dim oPAEntityItem As ProjectEntityItem = Nothing
    GenerateMetaDataCustom = String.Empty
    Dim sCode As String = sProjectName

    'get the project name
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Then
        oPAProject = pap
        Exit For
      End If

    Next
    Dim fs1 As FileStream = New FileStream("c:\Framework\MetaData_Custom.vb", FileMode.Create, FileAccess.Write)
    Dim s1 As StreamWriter = New StreamWriter(fs1)
    s1.WriteLine("  'PA_Code Generator Code")
    s1.WriteLine("  'Global declarations")
    s1.WriteLine("  '" & Now & ":, " & sCode)
    s1.WriteLine()

    s1.WriteLine("  Option Explicit On")
    s1.WriteLine("  Imports PA_Framework_Lib")
    s1.WriteLine("  Partial Public Class PAProjects")

    s1.WriteLine("  Public Function CreateObjectFromString(ByVal sPropertyName As String) As Object")
    s1.WriteLine("    'Creates object from child property name ")
    s1.WriteLine("    CreateObjectFromString = Nothing")
    s1.WriteLine("    Select sPropertyName")


    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Or pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("    Case ""{0:C}"", ""{1:C}""", pae.EntityName, pae.EntityCollectionName)
          s1.WriteLine("      CreateObjectFromString = New {0:C}", pae.EntityName)
        Next
      End If
    Next
    s1.WriteLine("")

    s1.WriteLine("    End Select")
    s1.WriteLine("    If CreateObjectFromString Is Nothing Then")
    s1.WriteLine("'     MsgBox(""Error: CreateObjectFromString does not seem to contain the case for "" & sPropertyName)")
    s1.WriteLine("     Call AppSettings.WriteToErrorLog(""Error: CreateObjectFromString does not seem to contain the case for "" & sPropertyName)")
    s1.WriteLine("    End If")
    s1.WriteLine("  End Function")

    s1.WriteLine("")
    s1.WriteLine("")


    s1.WriteLine("  Public Sub DBLoadWithUpdateLogItems(ByVal UpLIs As UpdateLogItems)")
    s1.WriteLine("    'Input UpLIs contains the collection of all updatelog items since last update.")
    s1.WriteLine("    Dim ul As UpdateLogItem")
    s1.WriteLine("     For Each ul In UpLIs")
    s1.WriteLine("     Select ul.sTableID")

    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Or pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("    Case ""{0:C}""", pae.EntityShortName)
          s1.WriteLine("      Call c{0:C}.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)", pae.EntityCollectionName)
        Next
      End If
    Next
    s1.WriteLine("    Case Else")

    s1.WriteLine("      MsgBox(""Error in DBLoadWithUpdateLogItems(); UpdatelogItem '"" & ul.sTableID & ""' not present"")")
    s1.WriteLine("      Call AppSettings.WriteToErrorLog(""Error in DBLoadWithUpdateLogItems(); UpdatelogItem '"" & ul.sTableID & ""' not present"")")
    s1.WriteLine("    End Select")
    s1.WriteLine("  Next ul")
    s1.WriteLine("  End Sub")
    s1.WriteLine("End Class")



    s1.Close()
    fs1.Close()
    s1 = Nothing
    fs1 = Nothing

  End Function
End Class
'Solution for tlb registration
'1)     Open a regular command prompt (NOT elevated)
'2)     Locate regcap.exe (it’s typically located in \program files\microsoft visual studio 9.0\common7\tools\deployment\ - replace “program files” with “program files (x86)” on 64-bit systems), and add this directory to your path (for example, run “set PATH=%PATH%;c:\program files\microsoft visual studio 9.0\common7\tools\deployment\”).
'3)     Run the following command: “regcap /I /O outputfile.reg inputfile.tlb”, replacing “outputfile.reg” with a .reg filename in a location you have write access to, and “inputfile.tlb” with the TLB you’re trying to create registration information for
'4)     Locate the .reg file you created in step #3, and import it into your setup project through the registry editor (to find the option to do this, right-click on the root node “registry on the target machine” and choose “Import…”)
'5)     In the “detected dependencies” folder, select your TLB, and hit F4 to see the properties grid.
'6)     In the properties grid, change the “Register” field to be “vsdrfDoNotRegister” (this will eliminate the build warning you were seeing previously) 