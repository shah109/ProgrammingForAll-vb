Option Explicit On
Imports System.IO
Imports System.Text
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class PACode_Global
  Public Shared Function GenerateOMGlobalCustom(ByVal sProjectName As String) As String
    'Creates all global declarations

    Dim oPAProject As New PAProject
    Dim oPAEntity As New ProjectEntity
    Dim oPAEntityItem As ProjectEntityItem = Nothing
    GenerateOMGlobalCustom = String.Empty
    Dim sCode As String = sProjectName

    'get the project name
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Then
        oPAProject = pap
        Exit For
      End If

    Next
    Dim fs1 As FileStream = New FileStream("c:\Framework\OMGlobals_Custom.vb", FileMode.Create, FileAccess.Write)
    Dim s1 As StreamWriter = New StreamWriter(fs1)
    s1.WriteLine("  'PA_Code Generator Code")
    s1.WriteLine("  'Global declarations")
    s1.WriteLine("  '" & Now & ":, " & sCode)
    s1.WriteLine()

    s1.WriteLine("  Option Explicit On")
    s1.WriteLine("  Imports PA_Framework_Lib")
    s1.WriteLine("  Public Module OMGlobals_Custom")
    s1.WriteLine("  'SECTION 0")
    s1.WriteLine("  'Short Entity Names")

    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Or pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("  Public {0:C} As {1:C}", pae.EntityShortName, pae.EntityName)
        Next
      End If
    Next
    s1.WriteLine("")

    s1.WriteLine("  'SECTION 1")
    s1.WriteLine("  'Define Collections for each entity in the solution")

    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Or pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("  Public c{0:C} As {0:C}", pae.EntityCollectionName)
        Next
      End If
    Next

    s1.WriteLine("")
    s1.WriteLine("  'SECTION 2")
    s1.WriteLine("  'Define the current entity for each entity in the solution. Used to track the current selected entity on the user interface")

    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Or pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("  Public curr{0:C} As New {1:C}", pae.EntityShortName, pae.EntityName)
        Next
      End If
    Next

    s1.WriteLine("")
    s1.WriteLine("  Public Sub LoadEntities()")
    s1.WriteLine("  'SECTION 3")
    s1.WriteLine("  'For Each Entity Instantiate a new collection.")

    'two for each loops because paproject entities need to be loaded first.
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("    c{0:C} = New {0:C}", pae.EntityCollectionName)
        Next
      End If
    Next
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("    c{0:C} = New {0:C}", pae.EntityCollectionName)
        Next
      End If
    Next

    s1.WriteLine("")

    s1.WriteLine("  'SECTION 4")
    s1.WriteLine("    'Load each entity starting with  entities with no childs and moving up")

    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("    c{0:C}.Load({1:C})", pae.EntityCollectionName, pae.EntityShortName)
        Next
      End If
    Next
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("    c{0:C}.Load({1:C})", pae.EntityCollectionName, pae.EntityShortName)
        Next
      End If
    Next

    s1.WriteLine("")
    s1.WriteLine("    End Sub")
    s1.WriteLine("  End Module")

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