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

    public enum SomeEnum
    {
        One,
        Two,
        Three
    }

    public static class MarshalByRefExamples
    {
        [ExcelMapArrayFunction]
        public static IEnumerable<ISampleMarshalByRefInterface> dnaFactory(IEnumerable<SomeEnum> enumValues,
            IEnumerable<double> doubleValues)
        {
            var enumsIter = enumValues.GetEnumerator();
            var valuesIter = doubleValues.GetEnumerator();
            while(enumsIter.MoveNext() && valuesIter.MoveNext())
            {
                ISampleMarshalByRefInterface item;
                switch (enumsIter.Current)
                {
                    case SomeEnum.One:
                        item = new SampleClass1(valuesIter.Current);
                        break;
                    case SomeEnum.Two:
                        item = new SampleClass2(valuesIter.Current);
                        break;
                    default:
                        throw new ArgumentException($"Don't know how to create an object of type {enumsIter.Current}.");
                }
                yield return item;
            }
        }

        [ExcelMapArrayFunction]
        public static double dnaMarshalByRef(ISampleMarshalByRefInterface[] objects)
        {
            return objects.Sum(x => x.Property);
        }
    }
}
