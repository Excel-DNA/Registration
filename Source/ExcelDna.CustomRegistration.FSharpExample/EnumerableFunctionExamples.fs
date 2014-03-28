namespace ExcelDna.CustomRegistration.FSharpExample

open System
open System.Threading
open System.Net
open Microsoft.FSharp.Control.WebExtensions
open ExcelDna.Integration
open ExcelDna.CustomRegistration.FSharp

module EnumerableFunctionExamples =

    type Input = {
        Date : System.DateTime;
        Bid : double;
        Ask : double;
    }

    type Output = {
        Date : System.DateTime;
        Mid : double;
    }

    [<ExcelFunction>]
    let CalculateMids (input:seq<Input>) :seq<Output> = 
        let CalculateMid (input:Input) :Output = 
            { Date = input.Date; Mid = (input.Bid + input.Ask)/2.}
        Seq.map CalculateMid input

