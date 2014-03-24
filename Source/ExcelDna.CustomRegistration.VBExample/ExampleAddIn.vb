Imports ExcelDna.Integration

Public Class ExampleAddIn
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ExcelIntegration.RegisterUnhandledExceptionHandler(Function(ex) "!!! ERROR: " + ex.ToString())

        Dim conversionConfig = New ParameterConversionConfiguration()
        conversionConfig.AddParameterConversion(AddressOf ParameterConversions.OptionalConversion)
        
        Registration.GetExcelFunctions() _
                    .ProcessParameterConversions(conversionConfig) _
                    .ProcessParamsRegistrations() _
                    .RegisterFunctions()

        ' Could add Async too...
    End Sub
    
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose

    End Sub

End Class
