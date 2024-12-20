// <copyright file="TargetLocator.cs" company="WebDriver Committers">
// Licensed to the Software Freedom Conservancy (SFC) under one
// or more contributor license agreements. See the NOTICE file
// distributed with this work for additional information
// regarding copyright ownership. The SFC licenses this file
// to you under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
// </copyright>

using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Internal;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PP5AutoUITests
{
    /// <summary>
    /// Provides a mechanism for finding elements on the page with locators.
    /// </summary>
    internal class MyTargetLocator : OpenQA.Selenium.ITargetLocator
    {
        private IDriver<IElement> driver;
        private ICommandExecutor executor;
        private SessionId sessionId;

        /// <summary>
        /// Initializes a new instance of the <see cref="TargetLocator"/> class
        /// </summary>
        /// <param name="driver">The driver that is currently in use</param>
        public MyTargetLocator(IDriver<IElement> driver)
        {
            this.driver = driver;
        }

        /// <summary>
        /// Change to the Window by passing in the name
        /// </summary>
        /// <param name="windowHandleOrName">Window handle or name of the window that you wish to move to</param>
        /// <returns>A WebDriver instance that is currently in use</returns>
        public IDriver<IElement> Window(string windowHandleOrName) => base.(windowHandleOrName);

        /// <summary>
        /// Finds the active element on the page and returns it
        /// </summary>
        /// <returns>Element that is active</returns>
        public IElement ActiveElement()
        {
            Response response = this.driver.InternalExecute(DriverCommand.GetActiveElement, null);
            return this.driver.GetElementFromResponse(response);
        }

        /// <summary>
        /// Executes commands with the driver
        /// </summary>
        /// <param name="driverCommandToExecute">Command that needs executing</param>
        /// <param name="parameters">Parameters needed for the command</param>
        /// <returns>WebDriver Response</returns>
        internal Response InternalExecute(string driverCommandToExecute, Dictionary<string, object> parameters)
        {
            return Task.Run(() => this.InternalExecuteAsync(driverCommandToExecute, parameters)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Executes commands with the driver asynchronously
        /// </summary>
        /// <param name="driverCommandToExecute">Command that needs executing</param>
        /// <param name="parameters">Parameters needed for the command</param>
        /// <returns>A task object representing the asynchronous operation</returns>
        internal Task<Response> InternalExecuteAsync(string driverCommandToExecute,
            Dictionary<string, object> parameters)
        {
            return this.ExecuteAsync(driverCommandToExecute, parameters);
        }

        /// <summary>
        /// Executes a command with this driver.
        /// </summary>
        /// <param name="driverCommandToExecute">A <see cref="DriverCommand"/> value representing the command to execute.</param>
        /// <param name="parameters">A <see cref="Dictionary{K, V}"/> containing the names and values of the parameters of the command.</param>
        /// <returns>A <see cref="Response"/> containing information about the success or failure of the command and any data returned by the command.</returns>
        protected virtual async Task<Response> ExecuteAsync(string driverCommandToExecute, Dictionary<string, object> parameters)
        {
            Command commandToExecute = new Command(this.sessionId, driverCommandToExecute, parameters);

            Response commandResponse;

            try
            {
                commandResponse = await this.executor.ExecuteAsync(commandToExecute).ConfigureAwait(false);
            }
            catch (System.Net.Http.HttpRequestException e)
            {
                commandResponse = new Response
                {
                    Status = WebDriverResult.UnhandledError,
                    Value = e
                };
            }

            if (commandResponse.Status != WebDriverResult.Success)
            {
                UnpackAndThrowOnError(commandResponse, driverCommandToExecute);
            }

            return commandResponse;
        }
    }
}
