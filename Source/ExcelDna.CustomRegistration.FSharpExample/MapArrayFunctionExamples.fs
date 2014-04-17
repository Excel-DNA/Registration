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

    type DateBidAsk = {
        Date : System.DateTime;
        Bid : double;
        Ask : double;
    }

    type DateMid = {
        Date : System.DateTime;
        Mid : double;
    }

    [<ExcelMapArrayFunction>]
    let dnaFsCalculateMids (input:seq<DateBidAsk>) :seq<DateMid> = 
        let CalculateMid (input:DateBidAsk) :DateMid = 
            { Date = input.Date; Mid = (input.Bid + input.Ask)/2.}
        Seq.map CalculateMid input

    // combines two blocks of Date Bid Ask, creating a single block ordered by Date
    [<ExcelMapArrayFunction>]
    let dnaFsCombineByDate (input1:seq<DateBidAsk>) (input2:seq<DateBidAsk>) :seq<DateBidAsk> = 
        let calculateAverage date (input:seq<DateBidAsk>) =
            let (bid,ask,count) = input |> Seq.fold (fun (bid,ask,count) item -> (bid+item.Bid,ask+item.Ask,count+1)) (0.,0.,0)
            if count > 0 then
                { Date=date; Bid=bid/float(count); Ask=ask/float(count) }
            else
                { Date=date; Bid=0.; Ask=0. }
        input1 |> Seq.append input2 |> Seq.groupBy (fun item -> item.Date) |> Seq.map (fun (date,items) -> 
            items |> calculateAverage date) |> Seq.sortBy (fun item -> item.Date)
