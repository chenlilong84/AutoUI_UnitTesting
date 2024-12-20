using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;
using OpenQA.Selenium;

namespace PP5AutoUITests
{
    public class DriverLogger
    {
        private ILogs driverLogs;  
        private Executor executor;
        private IOptions driverOps => executor.driverOps;

        public DriverLogger(Executor _executor) 
        {
            executor = _executor;
        }

        public void SetDriverLog()
        {
            driverLogs = driverOps.Logs;
        }

        public ReadOnlyCollection<string> GetAvailableLogTypes()
        {
            SetDriverLog();
            return driverLogs.AvailableLogTypes;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="driverLogType">Log type: Client, Driver, Browser, Server, Profiler, Performance</param>
        /// <returns></returns>
        public ReadOnlyCollection<LogEntry> GetLogEntries(string driverLogType)
        {
            if (!GetAvailableLogTypes().Contains(driverLogType))
                throw new ArgumentException("Input LogType is not supported!!");

            return driverLogs.GetLog(driverLogType);
        }

        public void DumpDriverLog(string driverLogType)
        {
            foreach(LogEntry log in GetLogEntries(driverLogType))
            {
                Logger.LogMessage(log.ToString(), ",", " LogType: ", driverLogType);
            }
        }

        public void DumpAllTypesDriverLog()
        {
            foreach (string logType in GetAvailableLogTypes())
            {
                foreach (LogEntry log in GetLogEntries(logType))
                {
                    Logger.LogMessage(log.ToString(), ",", " LogType: ", logType);
                }
            }
        }
    }
}
