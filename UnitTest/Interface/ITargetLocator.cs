// <copyright file="ITargetLocator.cs" company="WebDriver Committers">
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

namespace PP5AutoUITests
{
    /// <summary>
    /// Defines the interface through which the user can locate a given frame or window.
    /// </summary>
    public interface ITargetLocator : OpenQA.Selenium.ITargetLocator
    {
        /// <summary>
        /// Switches the focus of future commands for this driver to the window with the given name.
        /// </summary>
        /// <param name="windowName">The name of the window to select.</param>
        /// <returns>An <see cref="IDriver<IElement>"/> instance focused on the given window.</returns>
        /// <exception cref="NoSuchWindowException">If the window cannot be found.</exception>
        IDriver<IElement> Window(string windowName);

        /// <summary>
        /// Switches to the element that currently has the focus, or the body element
        /// if no element with focus can be detected.
        /// </summary>
        /// <returns>An <see cref="IElement"/> instance representing the element
        /// with the focus, or the body element if no element with focus can be detected.</returns>
        IElement ActiveElement();
    }
}
