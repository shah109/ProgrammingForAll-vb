Option Explicit On
Imports System.IO
Imports System.Text
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class PACode_EntityBaseClass
  Public Shared Function GenerateEntityBaseClass(ByVal sProjectName As String) As String
    'Creates Entity base class 'PAEntity' with all common functions

    Dim oPAProject As New PAProject
    Dim oPAEntity As New ProjectEntity
    Dim oPAEntityItem As ProjectEntityItem = Nothing
    GenerateEntityBaseClass = String.Empty
    Dim sCode As String = sProjectName

    'get the project name
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Then
        oPAProject = pap
        Exit For
      End If

    Next
    Dim fs1 As FileStream = New FileStream("c:\Framework\PAEnt.vb", FileMode.Create, FileAccess.Write)
    Dim s1 As StreamWriter = New StreamWriter(fs1)
    s1.WriteLine("'PA_Code Generator Code")
    s1.WriteLine("'PAEntity Class")
    s1.WriteLine("'" & Now & ":, " & sCode)
    s1.WriteLine()

    s1.WriteLine("Option Explicit On")
    s1.WriteLine("Imports PA_Framework_Lib")
    s1.WriteLine("Public Class PAEnt")
    s1.WriteLine("  Public Loadorder As Integer")
    s1.WriteLine("  Dim sID As String")
    s1.WriteLine("  Public IsDirty As Boolean")
    s1.WriteLine("  Public mContainer As PAEnts")
    s1.WriteLine("  Protected sLastUpdate As Date")

    s1.WriteLine("")
    s1.WriteLine("  'SECTION 1: Add three declarations for each child entities of your entity as follows:")

    'get the entity name 
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Or pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities

          s1.WriteLine("  '{0:C}", pae.EntityCollectionName)
          s1.WriteLine("  Public mChild{0:C} As New {0:C}", pae.EntityCollectionName)
          s1.WriteLine("  Protected sChild{0:C}String As String", pae.EntityCollectionName)
          s1.WriteLine("  Protected mAvChild{0:C} As New {0:C}", pae.EntityCollectionName)
          s1.WriteLine("")
        Next
      End If
    Next
    s1.WriteLine("")
    'SECTION 2: Add two declarations for each Parent entities of your entity as follows:

    s1.WriteLine("  'SECTION 2: Add two declarations for each Parent entities of your entity as follows:")
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Or pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("  '{0:C}", pae.EntityCollectionName)
          s1.WriteLine("  Protected mParent{0:C} As New {0:C}", pae.EntityCollectionName)
          s1.WriteLine("  Protected mAvParent{0:C} As New {0:C}", pae.EntityCollectionName)
          s1.WriteLine("")
        Next pae
      End If
    Next pap
    s1.WriteLine("  Public Property ID() As String")
    s1.WriteLine("    Get")
    s1.WriteLine("      ID = sid")
    s1.WriteLine("    End Get")
    s1.WriteLine("      Set(ByVal value As String)")
    s1.WriteLine("      sid = value")
    s1.WriteLine("    End Set")
    s1.WriteLine("  End Property")
    s1.WriteLine()
    s1.WriteLine("  Public Property Lastupdate() As Date")
    s1.WriteLine("    Get")
    s1.WriteLine("      Lastupdate = sLastUpdate")
    s1.WriteLine("    End Get")
    s1.WriteLine("      Set(ByVal value As Date)")
    s1.WriteLine("      sLastUpdate = value")
    s1.WriteLine("    End Set")
    s1.WriteLine("  End Property")
    s1.WriteLine()

    s1.WriteLine("  Public Overridable Function ChildEntityString(ByVal ent As String) As String")
    s1.WriteLine("    ChildEntityString = """"")
    s1.WriteLine("    Select ent")
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Or pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("    Case ""{0:C}"", ""{1:C}""", pae.EntityName, pae.EntityCollectionName)
          s1.WriteLine("      ChildEntityString = sChild{0:C}String", pae.EntityCollectionName)
        Next pae
      End If
    Next pap
    s1.WriteLine()
    s1.WriteLine("    End Select")
    s1.WriteLine("  End Function")
    s1.WriteLine("")

    s1.WriteLine("  Public Overridable Sub ChildEntityString(ByVal ent As String, ByVal strEnt As String)")
    s1.WriteLine("    Select ent")
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Or pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("    Case ""{0:C}"", ""{1:C}""", pae.EntityName, pae.EntityCollectionName)
          s1.WriteLine("      sChild{0:C}String = strEnt", pae.EntityCollectionName)
        Next pae
      End If
    Next pap
    s1.WriteLine()
    s1.WriteLine("    End Select")
    s1.WriteLine("  End Sub")
    s1.WriteLine()


    s1.WriteLine("  Public Overridable Function ChildEntities(ByVal Ent As String) As PAEnts")
    s1.WriteLine("    ChildEntities = Nothing")
    s1.WriteLine("    Select Ent")
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Or pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities

          s1.WriteLine("    Case ""{0:C}"", ""{1:C}""", pae.EntityName, pae.EntityCollectionName)
          s1.WriteLine("      ChildEntities = mChild{0:C}", pae.EntityCollectionName)
          's1.WriteLine("    Case ""{0:C}"", ""{1:C}""", pae.EntityName, pae.EntityCollectionName)
          's1.WriteLine("      ChildEntityString = sChild{0:C}String", pae.EntityCollectionName)
        Next pae
      End If
    Next pap
    s1.WriteLine()
    s1.WriteLine("    End Select")
    s1.WriteLine("  End Function")
    s1.WriteLine()

    s1.WriteLine("  Public Overridable Function AvailableChildEntities(ByRef objChld As PAEnt) As PAEnts")
    s1.WriteLine("    AvailableChildEntities = Nothing")
    s1.WriteLine("    Select TypeName(objChld)")
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Or pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("    Case ""{0:C}"", ""{1:C}""", pae.EntityName, pae.EntityCollectionName)
          s1.WriteLine("      AvailableChildEntities = mAvChild{0:C}", pae.EntityCollectionName)
        Next pae
      End If
    Next pap
    s1.WriteLine()
    s1.WriteLine("    End Select")
    s1.WriteLine("  End Function")
    s1.WriteLine()

    s1.WriteLine("  Public Overridable Function ParentEntities(ByRef objPar As PAEnt) As PAEnts")
    s1.WriteLine("    ParentEntities = Nothing")
    s1.WriteLine("    Select TypeName(objPar)")
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Or pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("    Case ""{0:C}"", ""{1:C}""", pae.EntityName, pae.EntityCollectionName)
          s1.WriteLine("      ParentEntities = mParent{0:C}", pae.EntityCollectionName)
        Next pae
      End If
    Next pap
    s1.WriteLine()
    s1.WriteLine("    End Select")
    s1.WriteLine("  End Function")
    s1.WriteLine()


    s1.WriteLine("  Public Overridable Function AvailableParentEntities(ByRef objPar As PAEnt) As PAEnts")
    s1.WriteLine("    AvailableParentEntities = Nothing")
    s1.WriteLine("    Select TypeName(objPar)")
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Or pap.ProjectName = "PA Framework" Then
        For Each pae As ProjectEntity In pap.ChildprojectEntities
          s1.WriteLine("    Case ""{0:C}"", ""{1:C}""", pae.EntityName, pae.EntityCollectionName)
          s1.WriteLine("      AvailableParentEntities = mAvParent{0:C}", pae.EntityCollectionName)
        Next pae
      End If
    Next pap
    s1.WriteLine()
    s1.WriteLine("    End Select")
    s1.WriteLine("  End Function")
    s1.WriteLine("End Class")

    s1.WriteLine("")
    s1.WriteLine("")
    s1.WriteLine("")
    s1.WriteLine("")
    s1.WriteLine("")


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