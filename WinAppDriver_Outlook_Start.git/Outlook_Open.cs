using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace WinAppDriver_Outlook_Start
{
    [TestClass]
    public class OutlookStart: OutlookSession
    {
        [TestMethod]
        public void Outlook_Start()
        {
            Assert.AreEqual("受信トレイ - nayuki@live.jp - Microsoft Outlook", session.Title, "Expected:受信トレイ - nayuki@live.jp - Microsoft Outlook" + "actual:" + session.Title);
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
