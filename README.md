Excel-DNA Registration Helper
=============================

This library implements helper functions to assist and modify the Excel-DNA function registration, by applying various transformations before the functions are registered.

The following transformations have been implemented:

Generation of wrapper functions for:

- Functions returning Task<T> or IObservable<T> as asynchronous or RTD-based functions (including F# Async<T> functions)
- Optional parameters (with default values), 'params' parameters and Nullable<T> parameters
- Range parameters in Visual Basic functions

Examples of general function transformations:

- Logging / Caching / Timing handlers
- Suppress in Function Arguments dialog

_If you've previously used the CustomRegistration library, note that I've renamed and rearranged the project source, and renamed the output assembly from ExcelDna.CustomRegistration to ExcelDna.Registration. The last state of the project before the large-scale rearrangement is marked by the git tag **CustomRegistration_Before_Rename**, and can be retrieved from the release tab on GitHub._

### Getting Started
To make a simple add-in that uses the Excel-DNA Registration extension to dynamically update the HelpTopic information for function registrations:

1. Create a new C# Class Library (.NET Framework) project, e.g. called RegistrationHelpUpdate.
2. Open the Package Manager Console.
3. `PM> Install-Package ExcelDna.AddIn`
4. `PM> Install-Package ExcelDna.Registration`
5. Edit the RegistrationHelpUpdate-AddIn.dna file to add the ExceplicitRegistration flag:
```xml
<DnaLibrary Name="RegistrationHelpUpdate Add-In" RuntimeVersion="v4.0">
  <ExternalLibrary Path="RegistrationHelpUpdate.dll" ExplicitExports="false" ExplicitRegistration="true" LoadFromBytes="true" Pack="true" />
</DnaLibrary>
```
6. Insert the following code:
```cs
using System.Linq;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace RegistrationHelpUpdate
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            RegisterFunctions();
        }

        public void AutoClose()
        {
        }

        public void RegisterFunctions()
        {
            ExcelRegistration.GetExcelFunctions()
                             .Select(UpdateHelpTopic)
                             .RegisterFunctions(); 
        }

        public ExcelFunctionRegistration UpdateHelpTopic(ExcelFunctionRegistration funcReg)
        {
            funcReg.FunctionAttribute.HelpTopic = "http://www.bing.com";
            return funcReg;
        }
    }

    public class Functions
    {
        [ExcelFunction(HelpTopic ="http://www.google.com")]
        public static object SayHello()
        {
            return "Hello!!!";
        }
    }
}
```
7. Press F5 to compile and start in Excel.
8. Start typing `=SayHello(` in a cell and press the Fx button to open the function wizard. Check that  the HelpTopic has been updated during registration to open Bing instead of Google.

See the add-ins in the Samples directory to see various registration update extensions.

### Step-by-step for Visual Basic

Once you have a basic Visual Basic add-in working.

1. From the NuGet Package Manager Console, (or the Manage NuGet Packages dialog): 
    PM>  Install-Package ExcelDna.Registration.VisualBasic 

2. Fix up your .dna file by changing to ExplicitRegistration for your library registration, and packing the extra ExcelDna.Registration libraries: 

```xml
        <DnaLibrary Name="MyVisualBasic Add-In" RuntimeVersion="v4.0" > 
          <ExternalLibrary Path="MyVisualBasic.dll" ExplicitRegistration="true" LoadFromBytes="true" Pack="true" /> 
          <Reference Path="ExcelDna.Registration.dll" Pack="true" /> 
          <Reference Path="ExcelDna.Registration.VisualBasic.dll" Pack="true" /> 
        </DnaLibrary> 
```

3.Perform the explicit registration in your AutoOpen by calling ExcelDna.Registration.VisualBasic.PerformDefaultRegistration(): 

```vb
        Imports ExcelDna.Integration 
        Imports ExcelDna.Registration.VisualBasic 

        Public Class MyAddIn 
                Implements IExcelAddIn 

                Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen 
                        ' Code here will run eery time the add-in is loaded 
                        PerformDefaultRegistration() 
                End Sub 

                Public Sub AutoClose() Implements IExcelAddIn.AutoClose 
                        ' Code in here will run when the add-in is removed in the Add-Ins dialog, 
                        ' but not when Excel closes normally 
                End Sub 
        End Class 


    Public Function dnaTestParams(date1 As Date, ParamArray s() As String) As String 
        Return s.Length 
    End Function 
    
    Public Function dnaTestOptional(date1 As Date, Optional head As Boolean = True) As String 
        Return head.ToString() 
    End Function 
```

### _Registration [Error] Repeated function name..._
_If you receive this error when opening your Excel addin, you need to add `ExplicitRegistration="true"` to the `<ExternalLibrary Path="MyAddin.dll"...` command in your .dna file_.
