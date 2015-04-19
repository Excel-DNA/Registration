
' Performs function lookup and registration selected for VBA compatibility
' * Does not require function to be marked by ExcelFunction
' * Enables optional parameters and default values
' * Enables Params parameter arrays
' * Does ReferenceToRange conversions (including setting IsMacroType=True for such functions)

Imports System.Reflection
Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration
Imports ExcelDna.Registration

Public Module VbaCompatibleRegistration

    Sub PerformDefaultRegistration()

        Dim conversionConfig As ParameterConversionConfiguration
        conversionConfig = New ParameterConversionConfiguration() _
                                .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing:=False)) _
                                .AddParameterConversion(AddressOf RangeParameterConversion.ParameterConversion, Nothing)

        GetAllPublicSharedFunctions() _
            .ProcessParamsRegistrations() _
            .UpdateRegistrationsForRangeParameters() _
            .ProcessParameterConversions(conversionConfig) _
            .RegisterFunctions()

        GetAllPublicSharedSubs().RegisterCommands()
    End Sub

    ' Gets the Public Shared methods that don't return Void
    Private Function GetAllPublicSharedFunctions() As IEnumerable(Of ExcelFunctionRegistration)
        Return From ass In ExcelIntegration.GetExportedAssemblies()
               From typ In ass.GetTypes()
               Where Not typ.FullName.Contains(".My.")
               From mi In typ.GetMethods(BindingFlags.Public Or BindingFlags.Static)
               Where Not mi.ReturnType = GetType(Void)
               Select New ExcelFunctionRegistration(mi)
    End Function

    ' Gets the Public Shared methods that return Void
    Private Function GetAllPublicSharedSubs() As IEnumerable(Of ExcelCommandRegistration)
        Return From ass In ExcelIntegration.GetExportedAssemblies()
               From typ In ass.GetTypes()
               Where Not typ.FullName.Contains(".My.")
               From mi In typ.GetMethods(BindingFlags.Public Or BindingFlags.Static)
               Where mi.ReturnType = GetType(Void)
               Select New ExcelCommandRegistration(mi)
    End Function

End Module
