using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace Registration.Sample
{
    [ExcelMarshalByRef]
    public interface ISampleMarshalByRefInterface
    {
        double Property { get; }
    }

    [ExcelMarshalByRef]
    public class SampleClass1 : ISampleMarshalByRefInterface
    {
        public double Property { private set; get; }
        public SampleClass1(double p)
        {
            Property = p;
        }
    }

    [ExcelMarshalByRef]
    public class SampleClass2 : ISampleMarshalByRefInterface
    {
        public double Property { private set; get; }
        public SampleClass2(double p)
        {
            Property = p*p;
        }
    }

    [ExcelMarshalByRef]
    public class Compound : ISampleMarshalByRefInterface
    {
        public double Property { private set; get; }
        public Compound(ISampleMarshalByRefInterface c1, ISampleMarshalByRefInterface c2)
        {
            Property = c1.Property + c2.Property;
        }
    }

    public enum SomeEnum
    {
        One,
        Two,
        Three
    }

    public static class MarshalByRefExamples
    {
        [ExcelMapArrayFunction]
        public static IEnumerable<ISampleMarshalByRefInterface> dnaFactoryMultiple(IEnumerable<SomeEnum> enumValues,
            IEnumerable<double> doubleValues)
        {
            var enumsIter = enumValues.GetEnumerator();
            var valuesIter = doubleValues.GetEnumerator();
            while(enumsIter.MoveNext() && valuesIter.MoveNext())
            {
                yield return dnaFactorySingle(enumsIter.Current, valuesIter.Current);
            }
        }

        [ExcelFunction]
        public static ISampleMarshalByRefInterface dnaFactorySingle(SomeEnum enumValue, double doubleValue)
        {
            ISampleMarshalByRefInterface item;
            switch (enumValue)
            {
                case SomeEnum.One:
                    item = new SampleClass1(doubleValue);
                    break;
                case SomeEnum.Two:
                    item = new SampleClass2(doubleValue);
                    break;
                default:
                    throw new ArgumentException($"Don't know how to create an object of type {enumValue}.");
            }
            return item;
        }

        [ExcelMapArrayFunction]
        public static double dnaMarshalByRef(IEnumerable<ISampleMarshalByRefInterface> objects)
        {
            return objects.Sum(x => x.Property);
        }

        [ExcelFunction]
        public static ISampleMarshalByRefInterface dnaFactoryCompound(ISampleMarshalByRefInterface c1, ISampleMarshalByRefInterface c2)
        {
            return new Compound(c1, c2);
        }
    }
}
