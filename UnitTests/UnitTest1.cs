using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            DateTime date = new DateTime(2017,4,1);
            int quart = (date.Month - 1) / 3 + 1;
            var from = new DateTime(date.Year, 1, 1).AddMonths(3 * (quart - 1));
        }
    }
}
