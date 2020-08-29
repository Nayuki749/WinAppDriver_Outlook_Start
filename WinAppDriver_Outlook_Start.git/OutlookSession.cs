using System.Collections.Generic;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium;
using System;

namespace WinAppDriver_Outlook_Start
{
    public class OutlookSession
    {
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        private const string NotepadAppId = @"C:\Program Files\Microsoft Office\Office14\OUTLOOK.EXE";

        protected static WindowsDriver<WindowsElement> session;
        protected static WindowsElement editBox;

        public static void Setup(TestContext context)
        {
            // Launch a new instance of Notepad application
            if (session == null)
            {
                // Create a new session to launch Notepad application
                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                appCapabilities.SetCapability("app", NotepadAppId);
                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
                Assert.IsNotNull(session);
                Assert.IsNotNull(session.SessionId);

                // Wait for a new window to appear
                WaitForNewWindow(session);
                // Change Window session
                session.SwitchTo().Window(session.WindowHandles[0]);

                // Verify that Notepad is started with untitled new file
                Assert.AreEqual("受信トレイ - nayuki@live.jp - Microsoft Outlook", session.Title);

                // Set implicit timeout to 1.5 seconds to make element search to retry every 500 ms for at most three times
                session.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1.5);
            }
        }

        public static void WaitForNewWindow(IWebDriver session)
        {
            // The current window handle of the running window.
            var windowHandle = session.CurrentWindowHandle;

            // retry coujnt. 
            const int retries = 120;

            for (var i = 0; i < retries; i++)
            {
                //wait 500ms Change it if you want
                Thread.Sleep(500);
                // if not found new window handle
                if (session.WindowHandles.Count <= 0) continue;

                //One or more new window handles are found.
                if (windowHandle != session.WindowHandles[0])
                    return;
            }
        }

        public static void TearDown()
        {
            // Close the application and delete the session
            if (session != null)
            {
                session.Close();

                try
                {

                }
                catch { }

                session.Quit();
                session = null;
            }
        }

        [TestInitialize]
        public void TestInitialize()
        {
        }
    }
}
