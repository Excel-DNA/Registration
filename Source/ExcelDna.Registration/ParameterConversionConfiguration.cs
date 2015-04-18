using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace ExcelDna.Registration
{
    // TODO: We need to turn these into objects with type and name so that we can trace and debug....
    public delegate LambdaExpression ParameterConversion(Type parameterType, ExcelParameterRegistration parameterRegistration);
    public delegate LambdaExpression ReturnConversion(Type returnType, List<object> returnCustomAttributes);

    // CONSIDER: Do we need to consider Co-/Contravariance and allow processing of sub-/super-types?
    // What about native async function, they return 'void' type?
    public class ParameterConversionConfiguration
    {
        // Token type used to indicate the conversions applied to all types in the Dictionaries.
        class GlobalConversionToken { }
        static internal Type GlobalConversionType = typeof(GlobalConversionToken);

        Dictionary<Type, List<ParameterConversion>> _parameterConversions;
        Dictionary<Type, List<ReturnConversion>> _returnConversions;

        // NOTE: Special extension of the type interpretation here, mainly to cater for the Range COM type equivalence
        public List<ParameterConversion> GetParameterConversions(Type paramType) 
        {
            return _parameterConversions.Where(kv => paramType == kv.Key || paramType.IsEquivalentTo(kv.Key))
                                       .SelectMany(kv => kv.Value)
                                       .Union(_parameterConversions[GlobalConversionType])
                                       .ToList();
        } 
    
        public List<ReturnConversion> GetReturnConversions(Type returnType) 
        {
            return _returnConversions.Where(kv => returnType == kv.Key || returnType.IsEquivalentTo(kv.Key))
                                    .SelectMany(kv => kv.Value)
                                    .Union(_returnConversions[GlobalConversionType])
                                    .ToList();
        }

        public ParameterConversionConfiguration()
        {
            _parameterConversions = new Dictionary<Type, List<ParameterConversion>>();
            _returnConversions = new Dictionary<Type, List<ReturnConversion>>();

            // Add room for the special 'global' conversions, applied for all types.
            _parameterConversions.Add(GlobalConversionType, new List<ParameterConversion>());
            _returnConversions.Add(GlobalConversionType, new List<ReturnConversion>());
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
        public ParameterConversionConfiguration AddParameterConversion(Type targetTypeOrNull, ParameterConversion parameterConversion)
        {
            var targetTypeOrGlobal = targetTypeOrNull ?? GlobalConversionType;

            List<ParameterConversion> typeConversions;
            if (_parameterConversions.TryGetValue(targetTypeOrGlobal, out typeConversions))
            {
                typeConversions.Add(parameterConversion);
            }
            else
            {
                _parameterConversions[targetTypeOrGlobal] = new List<ParameterConversion> { parameterConversion };
            }
            return this;
        }

        public ParameterConversionConfiguration AddParameterConversion(ParameterConversion parameterConversion)
        {
            AddParameterConversion(GlobalConversionType, parameterConversion);
            return this;
        }

        public ParameterConversionConfiguration AddParameterConversion<TTo>(ParameterConversion parameterConversion)
        {
            AddParameterConversion(typeof(TTo), parameterConversion);
            return this;
        }

        public ParameterConversionConfiguration AddParameterConversion<TFrom, TTo>(Expression<Func<TFrom, TTo>> convert)
        {
            AddParameterConversion<TTo>((unusedParamType, unusedParamReg) => convert);
            return this;
        }

        // This is a nice signature for registering conversions, but is worse than Expression<...> when applying
        public ParameterConversionConfiguration AddParameterConversionFunc<TFrom, TTo>(Func<TFrom, TTo> convert)
        {
            AddParameterConversion<TTo>(
                (unusedParamType, unusedParamReg) =>
                    (Expression<Func<TFrom, TTo>>)(value => convert(value)));
            return this;
        }

        public ParameterConversionConfiguration AddParameterConversion<TFrom, TTo>(Func<List<object>, TFrom, TTo> convertWithAttributes)
        {
            // CONSIDER: We really don't want our Registration processing to build out a closure object here...
            AddParameterConversion<TTo>(
                (unusedParamType, paramReg) =>
                    (Expression<Func<TFrom, TTo>>)(value => convertWithAttributes(paramReg.CustomAttributes, value)));
            return this;
        }

        // Most general case - called by the overloads below
        public ParameterConversionConfiguration AddReturnConversion(Type targetTypeOrNull, ReturnConversion returnConversion)
        {
            var targetTypeOrVoid = targetTypeOrNull ?? GlobalConversionType;

            List<ReturnConversion> typeConversions;
            if (_returnConversions.TryGetValue(targetTypeOrVoid, out typeConversions))
            {
                typeConversions.Add(returnConversion);
            }
            else
            {
                _returnConversions[targetTypeOrVoid] = new List<ReturnConversion> { returnConversion };
            }
            return this;
        }

        public ParameterConversionConfiguration AddReturnConversion<TFrom>(ReturnConversion returnConversion)
        {
            AddReturnConversion(typeof(TFrom), returnConversion);
            return this;
        }

        public ParameterConversionConfiguration AddReturnConversion<TFrom, TTo>(Func<TFrom, TTo> convert)
        {
            AddReturnConversion<TFrom>(
                (unusedReturnType, unusedAttributes) =>
                    (Expression<Func<TFrom, TTo>>)(value => convert(value)));
            return this;
        }

        public ParameterConversionConfiguration AddReturnConversion<TFrom, TTo>(Func<List<object>, TFrom, TTo> convertWithAttributes)
        {
            AddReturnConversion<TFrom>(
                (unusedReturnType, returnAttributes) =>
                    (Expression<Func<TFrom, TTo>>)(value => convertWithAttributes(returnAttributes, value)));
            return this;
        }
        #endregion
    }


}
