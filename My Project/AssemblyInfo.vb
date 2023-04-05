Imports System.Reflection
Imports System.Resources
Imports System.Runtime.InteropServices

' Les informations générales relatives à un assembly dépendent de 
' l'ensemble d'attributs suivant. Changez les valeurs de ces attributs pour modifier les informations
' associées à un assembly.

' Vérifiez les valeurs des attributs de l'assembly

<Assembly: AssemblyTitle("SifacToEdc")>
<Assembly: AssemblyDescription("Utilise des extractions SIFAC pour générer un EDC")>
<Assembly: AssemblyCompany("David Navarre")>
<Assembly: AssemblyProduct("SIFAC To EDC")>
<Assembly: AssemblyCopyright("Copyright ©  2023")>
<Assembly: AssemblyTrademark("")>

' L'affectation de la valeur false à ComVisible rend les types invisibles dans cet assembly 
' aux composants COM.  Si vous devez accéder à un type dans cet assembly à partir de 
' COM, affectez la valeur true à l'attribut ComVisible sur ce type.
<Assembly: ComVisible(False)>

'Le GUID suivant est pour l'ID de la typelib si ce projet est exposé à COM
<Assembly: Guid("277167e2-ea55-43bb-9e16-3b17e6d3cd8d")>

' Les informations de version pour un assembly se composent des quatre valeurs suivantes :
'
'      Version principale
'      Version secondaire 
'      Numéro de build
'      Révision
'
' Vous pouvez spécifier toutes les valeurs ou utiliser par défaut les numéros de build et de révision 
' en utilisant '*', comme indiqué ci-dessous :
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.0")>
<Assembly: AssemblyFileVersion("1.0.0.0")>
<Assembly: NeutralResourcesLanguage("fr")>
Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module
