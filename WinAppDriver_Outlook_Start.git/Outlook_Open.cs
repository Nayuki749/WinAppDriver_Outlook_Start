using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace WinAppDriver_Outlook_Start
{
    [TestClass]
    public class OutlookStart: OutlookSession
    {
        [TestMethod]
        public void TestMethod1()
        {
            
        }

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            Setup(context);
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            TearDown();
        }
    }
}
