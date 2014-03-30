
' Performs function lookup and registration selected for VBA compatibility
' * Does not require function to be marked by ExcelFunction
' * Enables optional parameters and default values
' * Enables Params parameter arrays
' * Does ReferenceToRange conversions (including setting IsMacroType=True for such functions)

Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration
Imports System.Reflection

Public Module VbaCompatibleRegistration

    Sub PerformDefaultRegistration()
        Dim conversionConfig As ParameterConversionConfiguration
        
        conversionConfig = New ParameterConversionConfiguration() _
                                .AddParameterConversion(AddressOf ParameterConversions.OptionalConversion) _
                                .AddParameterConversionFunc(Of ExcelReference, Range)(AddressOf ReferenceToRange)

        GetAllPublicStaticFunctions() _
            .UpdateRegistrationsForRangeParameters() _
            .ProcessParamsRegistrations() _
            .ProcessParameterConversions(conversionConfig) _
            .RegisterFunctions()
    End Sub

    Private Function GetAllPublicStaticFunctions() As IEnumerable(Of ExcelFunctionRegistration)
        Return from ass in ExcelIntegration.GetExportedAssemblies()
               from typ in ass.GetTypes()
               from mi in typ.GetMethods(BindingFlags.Public Or BindingFlags.Static)
               select new ExcelFunctionRegistration(mi)
    End Function

End Module
