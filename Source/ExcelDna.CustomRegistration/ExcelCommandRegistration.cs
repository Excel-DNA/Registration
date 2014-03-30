using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using ExcelDna.Integration;

namespace ExcelDna.CustomRegistration
{
    // Explicit support here for ExcelCommands is to encourage ExplicitRegistration=true
    // for all add-ins that use CustomRegistration.
    // But to support this we need to take care of ExcelCommands explicitly too.

    // Maybe one day we'll do Command/Function unification
    // For now we mirror core Excel-DNA approach
    // Note that Excel-DNA does support ExcelCommands that take parameters and return values.
    // However, these are not available as worksheet functions, and are unusual - 
    // so the ExcelCommandRegistration here doesn't support attributes on such parameters or return values.
    public class ExcelCommandRegistration
    {
        // These are used for registration
        public LambdaExpression CommandLambda { get; set; }
        public ExcelCommandAttribute CommandAttribute { get; set; }        // May not be null

        // These are used only for the CustomRegistration processing
        public List<object> CustomAttributes { get; set; }                 // List may not be null

        // TODO: Constructors, Registration etc. etc.
    }
}
