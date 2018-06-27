Option Explicit On
Imports System.IO
Imports System.Text
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class PACode_EntityClass
  Public Shared Function GenerateEntityClass(ByVal sProjectName As String) As String

    'Dim pae As ProjectEntity
    Dim oPAProject As New PAProject
    Dim oPAEntity As New ProjectEntity
    Dim oPAEntityItem As ProjectEntityItem = Nothing
    GenerateEntityClass = String.Empty
    Dim sCode As String = sProjectName


    'get the project name
    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Then
        oPAProject = pap
        Exit For
      End If

    Next

    For Each pae As ProjectEntity In oPAProject.ChildprojectEntities
      oPAEntity = pae

      Dim fs1 As FileStream = New FileStream("c:\Framework\" & oPAEntity.EntityName & ".vb", FileMode.Create, FileAccess.Write)
      Dim s1 As StreamWriter = New StreamWriter(fs1)
      s1.WriteLine("  'PA_Code Generator Code")
      s1.WriteLine("  '" & Now & ":, " & sCode)
      s1.WriteLine()

      s1.WriteLine("Option Explicit On")
      s1.WriteLine("Imports PA_Framework_Lib")
      s1.WriteLine("Imports PA_Framework_OM.OMGlobals")
      s1.WriteLine("Imports System.Runtime.InteropServices")

      s1.WriteLine("<ClassInterface(ClassInterfaceType.AutoDual)> _")
      s1.WriteLine("Public Class " & oPAEntity.EntityName)
      s1.WriteLine("  Inherits PAEnt")
      s1.WriteLine()

      For Each paei As ProjectEntityItem In oPAEntity.mChildProjectEntityItems
        If paei.ChildName <> "" Then GoTo NextForEachPaei
        If paei.PropertyName = "ID" Then GoTo NextForEachPaei
        s1.WriteLine("  Dim {0:C} As {1:C}", paei.InternalName, paei.InternalNameType)
NextForEachPaei:
      Next
      s1.WriteLine()

      For Each paei As ProjectEntityItem In oPAEntity.mChildProjectEntityItems
        If paei.ChildName <> "" Then GoTo NextForEachPaei1
        If paei.PropertyName = "ID" Then GoTo NextForEachPaei1
        s1.WriteLine("  Public Property {0:C}() As {1:C}", paei.PropertyName, paei.PropertyNameType)
        s1.WriteLine("    Get")
        s1.WriteLine("      {0:C} = {1:C}", paei.PropertyName, paei.InternalName)
        s1.WriteLine("    End Get")
        s1.WriteLine("    Set(ByVal value As {0:C})", paei.PropertyNameType)
        s1.WriteLine("      {0:C} = value", paei.InternalName)
        s1.WriteLine("    End Set")
        s1.WriteLine("  End Property")
        s1.WriteLine("")
NextForEachPaei1:
      Next

      s1.WriteLine("")
      s1.WriteLine("    Public Sub New()")
      s1.WriteLine("      mContainer = c{0:C}", oPAEntity.EntityCollectionName)
      s1.WriteLine("    End Sub")
      s1.WriteLine("  End Class")

      s1.Close()
      fs1.Close()
      s1 = Nothing
      fs1 = Nothing
    Next
  End Function

  'Function GenerateEntityClass()
  '  GenerateEntityClass = Nothing
  'End Function
  'Function GenerateMetaDataFile() As String
  '  GenerateMetaDataFile = String.Empty
  'End Function

  'Function GenerateOMGlobalsCustom() As String
  'GenerateOMGlobalsCustom = String.Empty
  'End Function

End Class
'Solution for tlb registration
'1)     Open a regular command prompt (NOT elevated)
'2)     Locate regcap.exe (it’s typically located in \program files\microsoft visual studio 9.0\common7\tools\deployment\ - replace “program files” with “program files (x86)” on 64-bit systems), and add this directory to your path (for example, run “set PATH=%PATH%;c:\program files\microsoft visual studio 9.0\common7\tools\deployment\”).
'3)     Run the following command: “regcap /I /O outputfile.reg inputfile.tlb”, replacing “outputfile.reg” with a .reg filename in a location you have write access to, and “inputfile.tlb” with the TLB you’re trying to create registration information for
'4)     Locate the .reg file you created in step #3, and import it into your setup project through the registry editor (to find the option to do this, right-click on the root node “registry on the target machine” and choose “Import…”)
'5)     In the “detected dependencies” folder, select your TLB, and hit F4 to see the properties grid.
'6)     In the properties grid, change the “Register” field to be “vsdrfDoNotRegister” (this will eliminate the build warning you were seeing previously) 