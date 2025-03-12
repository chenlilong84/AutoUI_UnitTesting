using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;

namespace PP5AutoUITests.Session
{
    public class MainPanelDriver : DriverBase, ISessionDriver
    {
        public MainPanelDriver()
        {
            SessionName = PowerPro5Config.MainPanelProcessName;
            WindowName = PowerPro5Config.MainPanelWindowName;
            sessionType = SessionType.MainPanel;
            //targetAppPath = @"C:\Users\adam.chen\Desktop\Debug\bin\Chroma.MainPanel.exe";
            //targetAppWorkingDir = @"C:\Users\adam.chen\Desktop\Debug\bin\";
            WaitForAppLaunch = 10;
            //targetAppPath = @"C:\Program Files (x86)\Chroma\PowerPro5\Bin\Chroma.MainPanel.exe";
            //targetAppWorkingDir = @"C:\Program Files (x86)\Chroma\PowerPro5\Bin\";
            targetAppPath = string.Format(PowerPro5Config.SubPathPattern, PowerPro5Config.ReleaseBinFolder, PowerPro5Config.MainPanelProcessName + ".exe");
            //targetAppWorkingDir = PowerPro5Config.ReleaseBinFolder;
        }

        public PP5Driver CreateNewDriver()
        {
            OpenQA.Selenium.Appium.AppiumOptions appCapabilities = new OpenQA.Selenium.Appium.AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", targetAppPath);
            appCapabilities.AddAdditionalCapability("appWorkingDir", targetAppWorkingDir);
            appCapabilities.AddAdditionalCapability("deviceName", DeviceName);
            appCapabilities.AddAdditionalCapability("ms:waitForAppLaunch", WaitForAppLaunch);
            currentDriver = new PP5Driver(new Uri(WindowsApplicationDriverUrl), appCapabilities, TimeSpan.FromSeconds(180), sessionType);

            currentDriver.ShouldNotBeNull($"PP5Driver: {sessionType.ToString()} created failed!");

            // Set implicit timeout to 1.5 seconds to make element search to retry every 500 ms for at most three times
            //currentDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10.0);
            return currentDriver;
        }

        public override PP5Driver CreateDriver()
        {
            processes = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(targetAppPath));
            if (processes.Length == 0)
            {
                CreateNewDriver();
            }
            else
            {
                AttachExistingDriver();
            }
            return currentDriver;
        }

        //public new void AttachExistingDriver()
        //{
        //    base.AttachExistingDriver();
        //}
    }
}
