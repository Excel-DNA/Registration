using System;
using System.Collections.Generic;

namespace ExcelDna.CustomRegistration
{
    public class FunctionExecutionConfiguration
    {
        internal List<Func<ExcelFunctionRegistration, FunctionExecutionHandler>> FunctionHandlerSelectors { get; private set; }

        public FunctionExecutionConfiguration()
        {
            FunctionHandlerSelectors = new List<Func<ExcelFunctionRegistration, FunctionExecutionHandler>>();
        }

        public void AddFunctionExecutionHandler(Func<ExcelFunctionRegistration, FunctionExecutionHandler> functionHandlerSelector)
        {
            FunctionHandlerSelectors.Add(functionHandlerSelector);
        }
    }

}
