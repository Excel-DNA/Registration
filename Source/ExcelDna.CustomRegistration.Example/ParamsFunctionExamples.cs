using ExcelDna.Integration;

namespace ExcelDna.CustomRegistration.Example
{
    public static class ParamsFunctionExamples
    {

        // This function has its final argument marked with 'params' 
        // Via the CustomRegistration helper will be registered in Excel as a function with 29 or 125 arguments,
        // and the wrapper will automatically remove 'ExcelMissing' values.
        //
        // If ExplicitRegistration="true" was _not_ in the .dna file, then
        // this function would normally be registed automatically by Excel-DNA.
        // (without the params processing) before being registered again here with the params expansion.
        //
        // We prevent that by adding the ExplicitRegistration=true falg.
        // 
        // Check how the parameters and their descriptions appear in the Function Arguments dialog...
        [ExcelFunction(ExplicitRegistration = true)]
        public static string dnaParamsFunc(
            [ExcelArgument(Name = "first.Input", Description = "is a useful start")]
            object input,
            [ExcelArgument(Description = "is another param start")]
            string QtherInpEt,
            [ExcelArgument(Name = "Value", Description = "gives the Rest")]
            params object[] args)
        {
            return input + "," + QtherInpEt + ", : " + args.Length;
        }

        [ExcelFunction(ExplicitRegistration = true)]
        public static string dnaParamsFunc2(
            [ExcelArgument(Name = "first.Input", Description = "is a useful start")]
            object input,
            [ExcelArgument(Name = "second.Input", Description = "is some more stuff")]
            string input2, 
            [ExcelArgument(Description = "is another param ")]
            string QtherInpEt,
            [ExcelArgument(Name = "Value", Description = "gives the Rest")]
            params object[] args)
        {
            return input + "," + QtherInpEt + ", : " + args.Length;
        }
    }
}
