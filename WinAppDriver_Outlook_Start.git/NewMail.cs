using System.Collections.Generic;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium;
using System;
using System.Diagnostics;

namespace WinAppDriver_Outlook_Start
{
    [TestClass]
    public class NewMail : OutlookSession
    {
        [TestMethod]
        public void NewMail_Open_and_Close_Test()
        {
            session.FindElementByName("新しい電子メール").Click();
            // Wait for a new window to appear
            WaitForNewWindow(session);
            // Change Window session
            session.SwitchTo().Window(session.WindowHandles[0]);

            Assert.AreEqual("無題 - メッセージ (HTML 形式) ", session.Title);

            // Click to Close button
            session.FindElementByName("閉じる").Click();
            // Change Window session
            session.SwitchTo().Window(session.WindowHandles[0]);
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
