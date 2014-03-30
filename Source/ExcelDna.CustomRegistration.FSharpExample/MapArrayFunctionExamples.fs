namespace ExcelDna.CustomRegistration.FSharpExample

open System
open System.Threading
open System.Net
open Microsoft.FSharp.Control.WebExtensions
open ExcelDna.Integration
open ExcelDna.CustomRegistration
open ExcelDna.CustomRegistration.FSharp

module MapArrayFunctionExamples =

    /// In Excel, use an Array Formula, e.g.
    ///       | A                 B       C         D                                    E       
    ///     --+---------------------------------------------------------------------------- 
    ///     1 | Date             Bid     Ask        {=MyFunc(A1:B3)} -> Date            Mid 
    ///     2 | 31 March 2014    1.99     2.01      {=MyFunc(A1:B3)} -> 31 March 2014    2.0
    ///     3 | 1 April 2014     2.01     2.05      {=MyFunc(A1:B3)} -> 1 April 2014     2.03

    type Input = {
        Date : System.DateTime;
        Bid : double;
        Ask : double;
    }

    type Output = {
        Date : System.DateTime;
        Mid : double;
    }

    [<ExcelMapArrayFunction>]
    let dnaFsCalculateMids (input:seq<Input>) :seq<Output> = 
        let CalculateMid (input:Input) :Output = 
            { Date = input.Date; Mid = (input.Bid + input.Ask)/2.}
        Seq.map CalculateMid input

