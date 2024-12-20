using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;

namespace PP5AutoUITests
{
    internal static class LogHelper
    {
        internal static void LogFindElementFailedByXPath(string xpath) 
        {
            Logger.LogMessage($"Failed to find element using xpath: {xpath}");
            Assert.Fail();
        }

        internal static void LogFindElementFailedByAutomationID(string automationID)
        {
            Logger.LogMessage($"Failed to find element using automationID: {automationID}");
            Assert.Fail();
        }

        internal static void LogFindElementFailedByName(string name)
        {
            Logger.LogMessage($"Failed to find element using name: {name}");
            Assert.Fail();
        }

        internal static void LogSendKeys(string inputText) 
        {
            Logger.LogMessage($"KeyboardInput: \"{inputText}\"");
        }
    }
}
