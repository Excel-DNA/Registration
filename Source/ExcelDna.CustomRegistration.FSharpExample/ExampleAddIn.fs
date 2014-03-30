namespace ExcelDna.CustomRegistration.FSharpExample

open ExcelDna.Integration
open ExcelDna.CustomRegistration
open ExcelDna.CustomRegistration.FSharp

type FsAsyncAddIn () =
    interface IExcelAddIn with
        member this.AutoOpen ()  = 
            let paramConvertConfig = ParameterConversionConfiguration()
                                        .AddParameterConversion(FsParameterConversions.FsOptionalParameterConversion)

            Registration.GetExcelFunctions ()
            |> fun fns -> ParameterConversionRegistration.ProcessParameterConversions (fns, paramConvertConfig)
            |> FsAsyncRegistration.ProcessFsAsyncRegistrations
            |> AsyncRegistration.ProcessAsyncRegistrations
            |> Registration.RegisterFunctions
        
        member this.AutoClose () = ()
