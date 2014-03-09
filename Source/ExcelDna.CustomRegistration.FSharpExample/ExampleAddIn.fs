namespace ExcelDna.CustomRegistration.FSharpExample

open ExcelDna.Integration
open ExcelDna.CustomRegistration
open ExcelDna.CustomRegistration.FSharp

type FsAsyncAddIn () =
    interface IExcelAddIn with
        member this.AutoOpen ()  = 
            Registration.GetExcelFunctions ()
            |> FsAsyncRegistration.ProcessFsAsyncRegistrations
            |> AsyncRegistration.ProcessAsyncRegistrations
            |> Registration.RegisterFunctions
        
        member this.AutoClose () = ()
    