using NUnit.Framework;
using PTO;
using System;

namespace PTOtests
{
    public class Tests
    {
        [TestCase(0.01, ExpectedResult = 6.982)]
        [TestCase(7.00, ExpectedResult = 164.96)]
        [TestCase(221.0, ExpectedResult = 374.06)]
        [TestCase(221.15, ExpectedResult = 374.15)]
        public double “Õ¿—€Ÿ≈Õ»ﬂ_test(double value)
        {
            return PTOfunctionCalls.“Õ¿—€Ÿ≈Õ»ﬂ(value - 1.01325, 0);

        }

        [TestCase (222.0)]
        public void “Õ¿—€Ÿ≈Õ»ﬂ_testEx(double value)
        {
            Assert.Throws(Is.TypeOf<ArgumentOutOfRangeException>().Or.Null, () => PTOfunctionCalls.“Õ¿—€Ÿ≈Õ»ﬂ(value - 1.01325, 0.0));
        }

        [TestCase(0.01, ExpectedResult = 29.33)]
        [TestCase(7.00, ExpectedResult = 697.1)]
        [TestCase(221.0, ExpectedResult = 2045.0)]
        [TestCase(221.15, ExpectedResult = 2095.2)]
        public double ›Õ“¿À‹œ»ﬂ¬_test(double value)
        {
            return PTOfunctionCalls.›Õ“¿À‹œ»ﬂ¬(value - 1.01325, 0.0);
        }

        [TestCase(0.01, ExpectedResult = 2513.8)]
        [TestCase(7.00, ExpectedResult = 2762.9)]
        [TestCase(221.0, ExpectedResult = 2147.6)]
        [TestCase(221.15, ExpectedResult = 2095.2)]
        public double ›Õ“¿À‹œ»ﬂœ_test(double value)
        {
            return PTOfunctionCalls.›Õ“¿À‹œ»ﬂœ(value - 1.01325, 0.0);
        }

    }
}