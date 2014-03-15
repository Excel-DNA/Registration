using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelDna.Integration;

namespace ExcelDna.CustomRegistration
{
    public class ExcelParameterRegistration
    {
        // Used for the final Excel-DNA registration
        public ExcelArgumentAttribute ArgumentAttribute { get; set; }

        // Used only for the CustomRegistration processing
        public List<object> CustomAttributes { get; set; } // Should not be null, and elements should not be null

        public ExcelParameterRegistration(ExcelArgumentAttribute argumentAttribute)
        {
            if (argumentAttribute == null) throw new ArgumentNullException("argumentAttribute");
            ArgumentAttribute = argumentAttribute;

            CustomAttributes = new List<object>();
        }

        // Checks that the property invariants are met, particularly regarding the attributes lists.
        internal bool IsValid()
        {
            return ArgumentAttribute != null && CustomAttributes != null && CustomAttributes.All(att => att != null);
        }
    }

    // CONSIDER: Improve safety here... make invalid data unrepresentable.
    // CONSIDER: Should ExcelCommands also be handled here...? For the moment not...
    public class ExcelFunctionRegistration
    {
        // These are used for registration
        public LambdaExpression FunctionLambda { get; set; }                        
        public ExcelFunctionAttribute FunctionAttribute { get; set; }                   // May not be null
        public List<ExcelParameterRegistration> ParameterRegistrations { get; set; }    // A list of ExcelParameterRegistrations with length equal to the number of parameters in Delegate

        // These are used only for the CustomRegistration processing
        public List<object> CustomAttributes { get; set; }                 // List may not be null
        public List<object> ReturnCustomAttributes { get; set; }                 // List may not be null

        // Checks that the property invariants are met, particularly regarding the attributes lists.
        internal bool IsValid()
        {
            return FunctionLambda != null &&
                   FunctionAttribute != null &&
                   ParameterRegistrations != null &&
                   ParameterRegistrations.Count == FunctionLambda.Parameters.Count &&
                   CustomAttributes != null &&
                   CustomAttributes.All(att => att != null) &&
                   ReturnCustomAttributes != null &&
                   ReturnCustomAttributes.All(att => att != null) &&
                   ParameterRegistrations.All(pr => pr.IsValid());
        }

        /// <summary>
        /// Creates a new ExcelFunctionRegistration with the given LambdaExpression.
        /// Uses the passes in attributes for registration.
        /// 
        /// The number of ExcelParameterRegistrations passed in must match the number of parameters in the LambdaExpression.
        /// </summary>
        /// <param name="functionLambda"></param>
        /// <param name="functionAttribute"></param>
        /// <param name="parameterRegistrations"></param>
        public ExcelFunctionRegistration(LambdaExpression functionLambda, ExcelFunctionAttribute functionAttribute, IEnumerable<ExcelParameterRegistration> parameterRegistrations = null)
        {
            if (functionLambda == null) throw new ArgumentNullException("functionLambda");
            if (functionAttribute == null) throw new ArgumentNullException("functionLambda");

            FunctionLambda = functionLambda;
            FunctionAttribute = functionAttribute;
            if (parameterRegistrations == null)
            {
                if (functionLambda.Parameters.Count != 0) throw new ArgumentOutOfRangeException("parameterRegistrations", "No parameter registrations provided, but function has parameters.");
                ParameterRegistrations = new List<ExcelParameterRegistration>();
            }
            else
            {
                ParameterRegistrations = new List<ExcelParameterRegistration>(parameterRegistrations);
                if (functionLambda.Parameters.Count != ParameterRegistrations.Count) throw new ArgumentOutOfRangeException("parameterRegistrations", "Mismatched number of ParameterRegistrations provided.");
            } 

            // Create the lists - hope the rest is filled in right...?
            CustomAttributes = new List<object>();
            ReturnCustomAttributes = new List<object>();
        }

        /// <summary>
        /// Creates a new ExcelFunctionRegistration from a LambdaExpression.
        /// Uses the Name and Parameter Names to fill in the default attributes.
        /// </summary>
        /// <param name="functionLambda"></param>
        public ExcelFunctionRegistration(LambdaExpression functionLambda)
        {
            if (functionLambda == null) throw new ArgumentNullException("functionLambda");

            FunctionLambda = functionLambda;
            FunctionAttribute = new ExcelFunctionAttribute { Name = functionLambda.Name };
            ParameterRegistrations = functionLambda.Parameters
                                     .Select( p => new ExcelParameterRegistration(new ExcelArgumentAttribute { Name = p.Name }))
                                     .ToList();

            CustomAttributes = new List<object>();
            ReturnCustomAttributes = new List<object>();
        }

        // NOTE: 16 parameter max for Expression.GetDelegateType
        // Copies all the (non Excel...) attributes from the method into the CustomAttribute lists.
        /// <summary>
        /// Creates a new ExcelFunctionRegistration from a MethodInfo, with a LambdaExpression that represents a call to the method.
        /// Uses the Name and Parameter Names from the MethodInfo to fill in the default attributes.
        /// All CustomAttributes on the method and parameters are copies to the respective collections in the ExcelFunctionRegistration.
        /// </summary>
        /// <param name="methodInfo"></param>
        public ExcelFunctionRegistration(MethodInfo methodInfo)
        {
            CustomAttributes = new List<object>();
            ReturnCustomAttributes = new List<object>();
            ParameterRegistrations = new List<ExcelParameterRegistration>();

            var paramExprs = methodInfo.GetParameters()
                             .Select(pi => Expression.Parameter(pi.ParameterType, pi.Name))
                             .ToList();
            FunctionLambda = Expression.Lambda(Expression.Call(methodInfo, paramExprs), methodInfo.Name, paramExprs);

            var allMethodAttributes = methodInfo.GetCustomAttributes(true);
            foreach (var att in allMethodAttributes)
            {
                var funcAtt = att as ExcelFunctionAttribute;
                if (funcAtt != null)
                {
                    FunctionAttribute = funcAtt;
                    // At least ensure that name is set - from the method if need be.
                    if (string.IsNullOrEmpty(FunctionAttribute.Name))
                        FunctionAttribute.Name = methodInfo.Name;
                }
                else
                {
                    CustomAttributes.Add(att);
                }
            }
            // Check that ExcelFunctionAttribute has been set
            if (FunctionAttribute == null)
            {
                FunctionAttribute = new ExcelFunctionAttribute { Name = methodInfo.Name };
            }

            foreach (var pi in methodInfo.GetParameters())
            {
                ExcelArgumentAttribute argumentAttribute = null;
                var paramCustomAttributes = new List<object>();

                var allParameterAttributes = pi.GetCustomAttributes(true);
                foreach (var att in allParameterAttributes)
                {
                    var argAtt = att as ExcelArgumentAttribute;
                    if (argAtt != null)
                    {
                        argumentAttribute = argAtt;
                        if (string.IsNullOrEmpty(argumentAttribute.Name))
                            argumentAttribute.Name = pi.Name;
                    }
                    else
                    {
                        paramCustomAttributes.Add(att);
                    }
                }

                // Check that the ExcelArgumentAttribute has been set
                if (argumentAttribute == null)
                {
                    argumentAttribute = new ExcelArgumentAttribute { Name = pi.Name };
                }

                var paramReg = new ExcelParameterRegistration(argumentAttribute);
                paramReg.CustomAttributes.AddRange(paramCustomAttributes);

                ParameterRegistrations.Add(paramReg);
            }

            ReturnCustomAttributes.AddRange(methodInfo.ReturnParameter.GetCustomAttributes(true));

            // Check that we haven't made a mistake
            Debug.Assert(IsValid());
        }
    }
}
