// <copyright file="RetryingElementLocator.cs" company="WebDriver Committers">
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

#if !NETSTANDARD2_0
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;
using OpenQA.Selenium;
using SeleniumExtras.PageObjects;

namespace PP5AutoUITests
{
    /// <summary>
    /// A locator for elements for use with the <see cref="PageFactory"/> that retries locating
    /// the element up to a timeout if the element is not found.
    /// </summary>
    public class RetryingElementLocator
    {
        private static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(5);
        private static readonly TimeSpan DefaultPollingInterval = TimeSpan.FromMilliseconds(500);

        //private ISearchContext searchContext;
        private IElement element;
        private PP5Driver driver;

        private TimeSpan timeout;
        private TimeSpan pollingInterval;

        /// <summary>
        /// Initializes a new instance of the <see cref="RetryingElementLocator"/> class.
        /// </summary>
        /// <param name="element">The <see cref="IElement"/> object that the
        /// locator uses for locating elements.</param>
        public RetryingElementLocator(IElement element)
            : this(element, DefaultTimeout)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RetryingElementLocator"/> class.
        /// </summary>
        /// <param name="element">The <see cref="IElement"/> object that the
        /// locator uses for locating elements.</param>
        /// <param name="timeout">The <see cref="TimeSpan"/> indicating how long the locator should
        /// retry before timing out.</param>
        public RetryingElementLocator(IElement element, TimeSpan timeout)
            : this(element, timeout, DefaultPollingInterval)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RetryingElementLocator"/> class.
        /// </summary>
        /// <param name="element">The <see cref="IElement"/> object that the
        /// locator uses for locating elements.</param>
        /// <param name="timeout">The <see cref="TimeSpan"/> indicating how long the locator should
        /// retry before timing out.</param>
        /// <param name="pollingInterval">The <see cref="TimeSpan"/> indicating how often to poll
        /// for the existence of the element.</param>
        public RetryingElementLocator(IElement element, TimeSpan timeout, TimeSpan pollingInterval)
        {
            this.driver = null;
            this.element = element;
            this.timeout = timeout;
            this.pollingInterval = pollingInterval;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RetryingElementLocator"/> class.
        /// </summary>
        /// <param name="driver">The <see cref="PP5Driver"/> object that the
        /// locator uses for locating elements.</param>
        public RetryingElementLocator(PP5Driver driver)
            : this(driver, DefaultTimeout)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RetryingElementLocator"/> class.
        /// </summary>
        /// <param name="driver">The <see cref="PP5Driver"/> object that the
        /// locator uses for locating elements.</param>
        /// <param name="timeout">The <see cref="TimeSpan"/> indicating how long the locator should
        /// retry before timing out.</param>
        public RetryingElementLocator(PP5Driver driver, TimeSpan timeout)
            : this(driver, timeout, DefaultPollingInterval)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RetryingElementLocator"/> class.
        /// </summary>
        /// <param name="driver">The <see cref="PP5Driver"/> object that the
        /// locator uses for locating elements.</param>
        /// <param name="timeout">The <see cref="TimeSpan"/> indicating how long the locator should
        /// retry before timing out.</param>
        /// <param name="pollingInterval">The <see cref="TimeSpan"/> indicating how often to poll
        /// for the existence of the element.</param>
        public RetryingElementLocator(PP5Driver driver, TimeSpan timeout, TimeSpan pollingInterval)
        {
            this.element = null;
            this.driver = driver;
            this.timeout = timeout;
            this.pollingInterval = pollingInterval;
        }

        /// <summary>
        /// Gets the <see cref="IElement"/> to be used in locating elements.
        /// </summary>
        public IElement Element
        {
            get { return this.element; }
        }

        /// <summary>
        /// Gets the <see cref="PP5Driver"/> to be used in locating elements.
        /// </summary>
        public PP5Driver Driver
        {
            get { return this.driver; }
        }

        /// <summary>
        /// Locates an element using the given list of <see cref="By"/> criteria.
        /// </summary>
        /// <param name="bys">The list of methods by which to search for the element.</param>
        /// <returns>An <see cref="IElement"/> which is the first match under the desired criteria.</returns>
        public IElement LocateElement(IEnumerable<By> bys)
        {
            if (bys == null)
            {
                throw new ArgumentNullException("bys", "List of criteria may not be null");
            }

            string errorString = null;
            DateTime endTime = DateTime.Now.Add(this.timeout);
            bool timeoutReached = DateTime.Now > endTime;
            while (!timeoutReached)
            {
                foreach (var by in bys)
                {
                    try
                    {
                        if (element == null)
                            return this.driver.FindElement(by);
                        else if (driver == null)
                            return this.element.FindElement(by);
                        //return this.element.FindElement(by);
                    }
                    catch (NoSuchElementException)
                    {
                        errorString = (errorString == null ? "Could not find element by: " : errorString + ", or: ") + by;
                    }
                }

                timeoutReached = DateTime.Now > endTime;
                if (!timeoutReached)
                {
                    Thread.Sleep(this.pollingInterval);
                }
            }

            throw new NoSuchElementException(errorString);
        }

        /// <summary>
        /// Locates a list of elements using the given list of <see cref="By"/> criteria.
        /// </summary>
        /// <param name="bys">The list of methods by which to search for the elements.</param>
        /// <returns>A list of all elements which match the desired criteria.</returns>
        public ReadOnlyCollection<IElement> LocateElements(IEnumerable<By> bys)
        {
            if (bys == null)
            {
                throw new ArgumentNullException("bys", "List of criteria may not be null");
            }

            List<IElement> collection = new List<IElement>();
            DateTime endTime = DateTime.Now.Add(this.timeout);
            bool timeoutReached = DateTime.Now > endTime;
            while (!timeoutReached)
            {
                foreach (var by in bys)
                {
                    ReadOnlyCollection<IElement> list = null;
                    if (element == null)
                        list = this.driver.FindElements(by);
                    else if (driver == null)
                        list = this.element.FindElements(by);
                    //ReadOnlyCollection<IElement> list = this.element.FindElements(by);
                    collection.AddRange(list);
                }

                timeoutReached = collection.Count != 0 || DateTime.Now > endTime;
                if (!timeoutReached)
                {
                    Thread.Sleep(this.pollingInterval);
                }
            }

            return collection.AsReadOnly();
        }
    }
}
#endif
