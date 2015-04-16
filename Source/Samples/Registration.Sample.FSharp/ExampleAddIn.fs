namespace Registration.Samples.FSharp

open ExcelDna.Integration
open ExcelDna.Registration
open ExcelDna.Registration.FSharp

type FsAsyncAddIn () =
    interface IExcelAddIn with
        member this.AutoOpen ()  = 
            let paramConvertConfig = ParameterConversionConfiguration()
                                        .AddParameterConversion(FsParameterConversions.FsOptionalParameterConversion)

            ExcelRegistration.GetExcelFunctions ()
            |> fun fns -> ParameterConversionRegistration.ProcessParameterConversions (fns, paramConvertConfig)
            |> FsAsyncRegistration.ProcessFsAsyncRegistrations
            |> AsyncRegistration.ProcessAsyncRegistrations
            |> MapArrayFunctionRegistration.ProcessMapArrayFunctions
            |> ExcelRegistration.RegisterFunctions
        
        member this.AutoClose () = ()
