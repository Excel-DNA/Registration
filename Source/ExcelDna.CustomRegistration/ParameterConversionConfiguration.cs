using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelDna.CustomRegistration
{

    // TODO: Maybe need to turn these into objects with type and name so that we can trace and debug....
    public delegate LambdaExpression ParameterConversion(Type parameterType, ExcelParameterRegistration parameterRegistration);
    public delegate LambdaExpression ReturnConversion(Type returnType, List<object> returnCustomAttributes);

    public class ParameterConversionConfiguration
    {
        // UGLY: We use Void as a special value to indicate the conversions to be processed for all types
        //       I try to hide that as an implementation, to the external functions use null to indicate the universal case.
        // Some concerns: What about native async function, they return 'void' type?
        // (Might interfere with our abuse of void in the Dictionary)

        internal Dictionary<Type, List<ParameterConversion>> ParameterConversions { get; private set; }
        internal Dictionary<Type, List<ReturnConversion>> ReturnConversions {get; private set; }

        public ParameterConversionConfiguration()
        {
            ParameterConversions = new Dictionary<Type, List<ParameterConversion>>();
            ReturnConversions = new Dictionary<Type, List<ReturnConversion>>();

            // Add room for the special 'global' conversions, applied for all types.
            ParameterConversions.Add(typeof(void), new List<ParameterConversion>());
            ReturnConversions.Add(typeof(void), new List<ReturnConversion>());
        }

        #region Various overloads for adding conversions

        // Most general case - called by the overloads below
        /// <summary>
        /// Converts a parameter from an Excel-friendly type (e.g. object, or string) to an add-in friendly type, e.g. double? or InternalType.
        /// Will only be considered for those parameters that have a 'to' type that matches targetTypeOrNull,
        ///  or for all types if null is passes for the first parameter.
        /// </summary>
        /// <param name="targetTypeOrNull"></param>
        /// <param name="parameterConversion"></param>
        public void AddParameterConversion(Type targetTypeOrNull, ParameterConversion parameterConversion)
        {
            var targetTypeOrVoid = targetTypeOrNull ?? typeof(void);

            List<ParameterConversion> typeConversions;
            if (ParameterConversions.TryGetValue(targetTypeOrVoid, out typeConversions))
            {
                typeConversions.Add(parameterConversion);
            }
            else
            {
                ParameterConversions[targetTypeOrVoid] = new List<ParameterConversion> { parameterConversion };
            }
        }

        public void AddParameterConversion(ParameterConversion parameterConversion)
        {
            AddParameterConversion(null, parameterConversion);
        }

        public void AddParameterConversion<TTo>(ParameterConversion parameterConversion)
        {
            AddParameterConversion(typeof(TTo), parameterConversion);
        }

        public void AddParameterConversion<TFrom, TTo>(Expression<Func<TFrom, TTo>> convert)
        {
            AddParameterConversion<TTo>((unusedParamType, unusedParamReg) => convert);
        }

        // This is a nice signature for registering conversions, but is worse than Expression<...> when applying
        public void AddParameterConversionFunc<TFrom, TTo>(Func<TFrom, TTo> convert)
        {
            AddParameterConversion<TTo>(
                (unusedParamType, unusedParamReg) =>
                    (Expression<Func<TFrom, TTo>>)(value => convert(value)));
        }

        public void AddParameterConversion<TFrom, TTo>(Func<List<object>, TFrom, TTo> convertWithAttributes)
        {
            // CONSIDER: We really don't want our the CustomRegistration compilation to build out a closure object here...
            AddParameterConversion<TTo>(
                (unusedParamType, paramReg) =>
                    (Expression<Func<TFrom, TTo>>)(value => convertWithAttributes(paramReg.CustomAttributes, value)));
        }

        // Most general case - called by the overloads below
        public void AddReturnConversion(Type targetTypeOrNull, ReturnConversion returnConversion)
        {
            var targetTypeOrVoid = targetTypeOrNull ?? typeof(void);

            List<ReturnConversion> typeConversions;
            if (ReturnConversions.TryGetValue(targetTypeOrVoid, out typeConversions))
            {
                typeConversions.Add(returnConversion);
            }
            else
            {
                ReturnConversions[targetTypeOrVoid] = new List<ReturnConversion> { returnConversion };
            }
        }

        public void AddReturnConversion<TFrom>(ReturnConversion returnConversion)
        {
            AddReturnConversion(typeof(TFrom), returnConversion);
        }

        public void AddReturnConversion<TFrom, TTo>(Func<TFrom, TTo> convert)
        {
            AddReturnConversion<TFrom>(
                (unusedReturnType, unusedAttributes) =>
                    (Expression<Func<TFrom, TTo>>)(value => convert(value)));
        }

        public void AddReturnConversion<TFrom, TTo>(Func<List<object>, TFrom, TTo> convertWithAttributes)
        {
            AddReturnConversion<TFrom>(
                (unusedReturnType, returnAttributes) =>
                    (Expression<Func<TFrom, TTo>>)(value => convertWithAttributes(returnAttributes, value)));
        }
        #endregion
    }


}
