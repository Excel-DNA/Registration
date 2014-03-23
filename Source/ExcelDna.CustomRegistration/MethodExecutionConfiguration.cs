using System;
using System.Collections.Generic;

namespace ExcelDna.CustomRegistration
{
    public class MethodExecutionConfiguration
    {
        internal List<Func<ExcelFunctionRegistration, MethodExecutionHandler>> MethodHandlers { get; private set; }

        public MethodExecutionConfiguration()
        {
            MethodHandlers = new List<Func<ExcelFunctionRegistration, MethodExecutionHandler>>();
        }

        public void AddMethodHandler(Func<ExcelFunctionRegistration, MethodExecutionHandler> methodHandlerSelector)
        {
            MethodHandlers.Add(methodHandlerSelector);
        }
    }

}
