using NUnit.Framework;
using PTO;
using System;
using System.Threading.Tasks;

namespace PTOtests
{
    internal class ЭНТАЛЬПИЯtests
    {

        [TestCase(0.0, 0.01, ExpectedResult = 0.0)]
        [TestCase(0.0, 70.0, ExpectedResult = 7.1)]
        [TestCase(0.0, 220.0, ExpectedResult = 22.1)]
        [TestCase(0.0, 950.0, ExpectedResult = 91.3)]

        [TestCase(10.0, 0.01, ExpectedResult = 2519.5)]
        //[TestCase(10.0, 70.0, ExpectedResult = 48.8)]
        [TestCase(10.0, 220.0, ExpectedResult = 63.2)]
        [TestCase(10.0, 950.0, ExpectedResult = 130.1)]
        public double ЭНТАЛЬПИЯ_test(double T, double P)
        {
            return PTOfunctionCalls.ЭНТАЛЬПИЯ(T, P - 1.01325, 0.0);
        }

        [Test]
        public void ЭНТАЛЬПИЯExceptionTest()
        {
            double P;

            for (P = 0.0; P <= 1000.0; P += 0.01)
            {
                Parallel.For(0, 8000, (int t) =>
                {
                    Assert.DoesNotThrow(() => PTOfunctionCalls.ЭНТАЛЬПИЯ(t / 10.0, P - 1.01325, 0.0));
                });
                TestContext.Progress.Write($"Running tests for P={P}");
                TestContext.Progress.Write("\r");
            }
        }
    }
}
