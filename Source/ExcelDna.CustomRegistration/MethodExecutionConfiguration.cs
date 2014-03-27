using System;
using System.Collections.Generic;

namespace ExcelDna.CustomRegistration
{
    public class MethodExecutionConfiguration
    {
        internal List<Func<ExcelFunctionRegistration, MethodExecutionHandler>> MethodHandlerSelectors { get; private set; }

        public MethodExecutionConfiguration()
        {
            MethodHandlerSelectors = new List<Func<ExcelFunctionRegistration, MethodExecutionHandler>>();
        }

        public void AddMethodHandler(Func<ExcelFunctionRegistration, MethodExecutionHandler> methodHandlerSelector)
        {
            MethodHandlerSelectors.Add(methodHandlerSelector);
        }
    }

}
