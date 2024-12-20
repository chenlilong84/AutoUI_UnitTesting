// <copyright file="DefaultElementLocator.cs" company="WebDriver Committers">
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
using OpenQA.Selenium;
//using SeleniumExtras.PageObjects;

namespace PP5AutoUITests
{
    /// <summary>
    /// A default locator for elements for use with the <see cref="PageFactory"/>. This locator
    /// implements no retry logic for elements not being found, nor for elements being stale.
    /// </summary>
    public class DefaultElementLocator
    {
        private IElement element;
        private PP5Driver driver;

        /// <summary>
        /// Initializes a new instance of the <see cref="DefaultElementLocator"/> class.
        /// </summary>
        /// <param name="element">The <see cref="IElement"/> used by this locator
        /// to locate elements.</param>
        public DefaultElementLocator(IElement element)
        {
            this.element = element;
            this.driver = null;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DefaultElementLocator"/> class.
        /// </summary>
        /// <param name="driver">The <see cref="PP5Driver"/> used by this locator
        /// to locate elements.</param>
        public DefaultElementLocator(PP5Driver driver)
        {
            this.driver = driver;
            this.element = null;
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
            foreach (var by in bys)
            {
                try
                {
                    //return this.Element.FindElement(by);
                    if (element == null)
                        return this.driver.FindElement(by);
                    else if (driver == null)
                        return this.element.FindElement(by);
                }
                catch (NoSuchElementException)
                {
                    errorString = (errorString == null ? "Could not find element by: " : errorString + ", or: ") + by;
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

            return collection.AsReadOnly();
        }
    }
}
#endif
