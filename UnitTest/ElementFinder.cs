#define DEBUG
using System;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Automation;
//using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Castle.Core.Internal;
using Chroma.UnitTest.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Html5;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
//using SeleniumExtras.PageObjects;
using static OpenQA.Selenium.Support.UI.ExpectedConditions;
using static PP5AutoUITests.AutoUIActionHelper;

namespace PP5AutoUITests
{
    public static class ElementFinder
    {
        #region Main Methods

        #region Legacy Find Element methods
        //public static IElement FindElementByAbsoluteXPath(this PP5Driver driver, string xPath, int nTryCount = 3)
        //{
        //    IElement uiTarget = null;

        //    while (nTryCount-- > 0)
        //    {
        //        try
        //        {
        //            uiTarget = driver.FindElementByXPath(xPath);
        //        }
        //        catch
        //        {
        //        }

        //        if (uiTarget != null)
        //        {
        //            break;
        //        }
        //        else
        //        {
        //            System.Threading.Thread.Sleep(10);
        //        }
        //    }

        //    return uiTarget;
        //}

        //public static IElement FindElementByAutomationId(this PP5Driver driver, string AutomationId, int nTryCount = 3)
        //{
        //    IElement uiTarget = null;

        //    while (nTryCount-- > 0)
        //    {
        //        try
        //        {
        //            uiTarget = driver.FindElementByAccessibilityId(AutomationId);
        //        }
        //        catch
        //        {
        //        }

        //        if (uiTarget != null)
        //        {
        //            break;
        //        }
        //        else
        //        {
        //            System.Threading.Thread.Sleep(10);
        //        }
        //    }

        //    return uiTarget;
        //}

        //public static ReadOnlyCollection<IElement> FindElementsByAutomationId(this PP5Driver driver, string AutomationId, int nTryCount = 3)
        //{
        //    ReadOnlyCollection<IElement> uiTargets = null;

        //    while (nTryCount-- > 0)
        //    {
        //        try
        //        {
        //            uiTargets = (ReadOnlyCollection<IElement>)driver.FindElementsByAccessibilityId(AutomationId);
        //        }
        //        catch
        //        {
        //        }

        //        if (uiTargets != null)
        //        {
        //            break;
        //        }
        //        else
        //        {
        //            System.Threading.Thread.Sleep(10);
        //        }
        //    }

        //    return uiTargets;
        //}

        //public static IElement FindElementByName(this PP5Driver driver, string Name, int nTryCount = 3)
        //{
        //    IElement uiTarget = null;

        //    while (nTryCount-- > 0)
        //    {
        //        try
        //        {
        //            uiTarget = driver.FindElementByName(Name);
        //        }
        //        catch
        //        {
        //        }

        //        if (uiTarget != null)
        //        {
        //            break;
        //        }
        //        else
        //        {
        //            System.Threading.Thread.Sleep(10);
        //        }
        //    }

        //    return uiTarget;
        //}

        //public static ReadOnlyCollection<IElement> FindElementsByName(this PP5Driver driver, string Name, int nTryCount = 3)
        //{
        //    ReadOnlyCollection<IElement> uiTargets = null;

        //    while (nTryCount-- > 0)
        //    {
        //        try
        //        {
        //            uiTargets = driver.FindElementsByName(Name);
        //        }
        //        catch
        //        {
        //        }

        //        if (uiTargets != null)
        //        {
        //            break;
        //        }
        //        else
        //        {
        //            System.Threading.Thread.Sleep(10);
        //        }
        //    }

        //    return uiTargets;
        //}

        //public static IElement GetElementById(this PP5Driver driver, string automationId, string propertyName, int timeOut = 10000)
        //{
        //    IElement element = null;
        //    var wait = new DefaultWait<WindowsDriver<WindowsElement>>(driver)
        //    {
        //        Timeout = TimeSpan.FromMilliseconds(timeOut),
        //        Message = $"Element with automationId \"{automationId}\" not found."
        //    };

        //    wait.IgnoreExceptionTypes(typeof(WebDriverException));

        //    try
        //    {
        //        wait.Until(Driver =>
        //        {
        //            element = Driver.FindElementByAccessibilityId(automationId);

        //            return element != null;
        //        });
        //    }
        //    catch (WebDriverTimeoutException ex)
        //    {
        //        Logger.LogMessage($"{ex}, {automationId}, {propertyName}");
        //        Assert.Fail(ex.Message);
        //    }

        //    return element;
        //}

        //public static IElement GetElementByName(this PP5Driver driver, string name, string propertyName, int timeOut = 10000)
        //{
        //    IElement element = null;
        //    var wait = new DefaultWait<WindowsDriver<WindowsElement>>(driver)
        //    {
        //        Timeout = TimeSpan.FromMilliseconds(timeOut),
        //        Message = $"Element with name \"{name}\" not found."
        //    };

        //    wait.IgnoreExceptionTypes(typeof(WebDriverException));

        //    try
        //    {
        //        wait.Until(Driver =>
        //        {
        //            element = Driver.FindElementByName(name);

        //            return element != null;
        //        });
        //    }
        //    catch (WebDriverTimeoutException ex)
        //    {
        //        Logger.LogMessage($"{ex}, {name}, {propertyName}");
        //        Assert.Fail(ex.Message);
        //    }

        //    return element;
        //}

        //public static ReadOnlyCollection<IElement> GetElementsById(this PP5Driver driver, string automationId, string propertyName, int timeOut = 10000)
        //{
        //    ReadOnlyCollection<IElement> elements = null;
        //    var wait = new DefaultWait<WindowsDriver<WindowsElement>>(driver)
        //    {
        //        Timeout = TimeSpan.FromMilliseconds(timeOut),
        //        Message = $"Element with automationId \"{automationId}\" not found."
        //    };

        //    wait.IgnoreExceptionTypes(typeof(WebDriverException));

        //    try
        //    {
        //        wait.Until(Driver =>
        //        {
        //            elements = (ReadOnlyCollection<IElement>)Driver.FindElementsByAccessibilityId(automationId);

        //            return elements != null;
        //        });
        //    }
        //    catch (WebDriverTimeoutException ex)
        //    {
        //        Logger.LogMessage($"{ex}, {automationId}, {propertyName}");
        //        Assert.Fail(ex.Message);
        //    }

        //    return elements;
        //}

        //public static IReadOnlyCollection<IElement> GetElementsByName(this PP5Driver driver, string name, string propertyName, int timeOut = 10000)
        //{
        //    IReadOnlyCollection<IElement> elements = null;
        //    var wait = new DefaultWait<WindowsDriver<WindowsElement>>(driver)
        //    {
        //        Timeout = TimeSpan.FromMilliseconds(timeOut),
        //        Message = $"Element with name \"{name}\" not found."
        //    };

        //    wait.IgnoreExceptionTypes(typeof(WebDriverException));

        //    try
        //    {
        //        wait.Until(Driver =>
        //        {
        //            elements = Driver.FindElementsByName(name);

        //            return elements != null;
        //        });
        //    }
        //    catch (WebDriverTimeoutException ex)
        //    {
        //        Logger.LogMessage($"{ex}, {name}, {propertyName}");
        //        Assert.Fail(ex.Message);
        //    }

        //    return elements;
        //}
        #endregion

        #region Find Element methods

        #region Base method to FindElement with retry
        // FindElement from driver, given By locator, retry
        public static IWebElement FindElement(this IWebDriver driver, By by, int nTryCount = 3)
        {
            IWebElement uiTarget = null;

            while (nTryCount-- > 0)
            {
                try
                {
                    uiTarget = driver.FindElement(by);
                }
                catch
                {
                }

                if (uiTarget != null)
                {
                    break;
                }
                else
                {
                    System.Threading.Thread.Sleep(10);
                }
            }

            return uiTarget;
        }

        // FindElements from driver, given By locator, retry
        public static ReadOnlyCollection<IWebElement> FindElements(this IWebDriver driver, By by, int nTryCount = 3)
        {
            ReadOnlyCollection<IWebElement> uiTargets = null;

            while (nTryCount-- > 0)
            {
                try
                {
                    uiTargets = driver.FindElements(by);
                }
                catch
                {
                }

                if (uiTargets != null)
                {
                    break;
                }
                else
                {
                    System.Threading.Thread.Sleep(10);
                }
            }

            return uiTargets;
        }

        // FindElement from element, given By locator, retry
        public static IWebElement FindElement(this IWebElement element, By by, int nTryCount = 3)
        {
            IWebElement uiTarget = null;

            while (nTryCount-- > 0)
            {
                try
                {
                    uiTarget = element.FindElement(by);
                }
                catch
                {
                }

                if (uiTarget != null)
                {
                    break;
                }
                else
                {
                    System.Threading.Thread.Sleep(10);
                }
            }

            return uiTarget;
        }

        // FindElements from element, given By locator, retry
        public static ReadOnlyCollection<IWebElement> FindElements(this IWebElement element, By by, int nTryCount = 3)
        {
            ReadOnlyCollection<IWebElement> uiTargets = null;

            while (nTryCount-- > 0)
            {
                try
                {
                    uiTargets = element.FindElements(by);
                }
                catch
                {
                }

                if (uiTargets != null)
                {
                    break;
                }
                else
                {
                    System.Threading.Thread.Sleep(10);
                }
            }

            return uiTargets;
        }
        #endregion

        #region GetElement No retry

        // Generic method to find element/elements by IWebdriver/PP5Driver/IWebElement/IElement
        public static TOutput GetElement<TInput, TOutput>(this TInput driverOrElement, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            TOutput elementResult;

            string inputTypeName = typeof(TInput).Name;
            string outputTypeName = typeof(TOutput).Name;

            if (inputTypeName is nameof(IWebElement) && outputTypeName is nameof(IWebElement))
                elementResult = (TOutput)(driverOrElement as IWebElement).GetWebElementFromWebElement(findType, timeOut);
            else if (inputTypeName is nameof(IWebElement) && outputTypeName is nameof(IElement))
                elementResult = (TOutput)(driverOrElement as IWebElement).GetElementFromWebElement(findType, timeOut);
            else if (inputTypeName is nameof(IElement) && outputTypeName is nameof(IWebElement))
                elementResult = (TOutput)(driverOrElement as IElement).GetWebElementFromPP5Element(findType, timeOut);
            else if (inputTypeName is nameof(IElement) && outputTypeName is nameof(IElement))
                elementResult = (TOutput)(driverOrElement as IElement).GetElementFromPP5Element(findType, timeOut);
            else if (inputTypeName is nameof(PP5Driver) && outputTypeName is nameof(IElement))
                elementResult = (TOutput)(driverOrElement as PP5Driver).GetElementFromPP5Driver(findType, timeOut);
            else
                throw new NotSupportedException("Not supported input or output type in GetElement method!");

            return elementResult;
        }

        public static TOutput GetElement<TInput, TOutput>(this TInput driverOrElement, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            TOutput elementResult;

            string inputTypeName = typeof(TInput).Name;
            string outputTypeName = typeof(TOutput).Name;

            if (inputTypeName is nameof(IWebElement) && outputTypeName is nameof(IWebElement))
                elementResult = (TOutput)(driverOrElement as IWebElement).GetWebElementFromWebElement(timeOut, findTypes);
            else if (inputTypeName is nameof(IWebElement) && outputTypeName is nameof(IElement))
                elementResult = (TOutput)(driverOrElement as IWebElement).GetElementFromWebElement(timeOut, findTypes);
            else if (inputTypeName is nameof(IElement) && outputTypeName is nameof(IWebElement))
                elementResult = (TOutput)(driverOrElement as IElement).GetWebElementFromPP5Element(timeOut, findTypes);
            else if (inputTypeName is nameof(IElement) && outputTypeName is nameof(IElement))
                elementResult = (TOutput)(driverOrElement as IElement).GetElementFromPP5Element(timeOut, findTypes);
            else if (inputTypeName is nameof(PP5Driver) && outputTypeName is nameof(IElement))
                elementResult = (TOutput)(driverOrElement as PP5Driver).GetElementFromPP5Driver(timeOut, findTypes);
            else
                throw new NotSupportedException("Not supported input or output type in GetElement method!");

            return elementResult;
        }

        // Generic method to find element/elements by IWebdriver/PP5Driver/IWebElement/IElement
        public static ReadOnlyCollection<TOutput> GetElements<TInput, TOutput>(this TInput driverOrElement, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            ReadOnlyCollection<TOutput> elementsResult;

            string inputTypeName = typeof(TInput).Name;
            string outputTypeName = typeof(TOutput).Name;

            if (inputTypeName is nameof(IWebElement) && outputTypeName is nameof(IWebElement))
                elementsResult = (ReadOnlyCollection<TOutput>)(driverOrElement as IWebElement).GetWebElementsFromWebElement(findType, timeOut).Cast<TOutput>();
            else if (inputTypeName is nameof(IElement) && outputTypeName is nameof(IElement))
                elementsResult = (ReadOnlyCollection<TOutput>)(driverOrElement as IElement).GetElementsFromPP5Element(findType, timeOut).Cast<TOutput>();
            else if (inputTypeName is nameof(PP5Driver) && outputTypeName is nameof(IElement))
                elementsResult = (ReadOnlyCollection<TOutput>)(driverOrElement as PP5Driver).GetElementsFromPP5Driver(findType, timeOut).Cast<TOutput>();
            else
                throw new NotSupportedException("Not supported input or output type in GetElements method!");

            return elementsResult;
        }

        public static ReadOnlyCollection<TOutput> GetElements<TInput, TOutput>(this TInput driverOrElement, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            ReadOnlyCollection<TOutput> elementsResult;

            string inputTypeName = typeof(TInput).Name;
            string outputTypeName = typeof(TOutput).Name;

            if (inputTypeName is nameof(IWebElement) && outputTypeName == nameof(IWebElement))
                elementsResult = (ReadOnlyCollection<TOutput>)(driverOrElement as IWebElement).GetWebElementsFromWebElement(timeOut, findTypes).Cast<TOutput>();
            else if (inputTypeName is nameof(IElement) && outputTypeName == nameof(IElement))
                elementsResult = (ReadOnlyCollection<TOutput>)(driverOrElement as IElement).GetElementsFromPP5Element(timeOut, findTypes).Cast<TOutput>();
            else if (inputTypeName is nameof(PP5Driver) && outputTypeName == nameof(IElement))
                elementsResult = (ReadOnlyCollection<TOutput>)(driverOrElement as PP5Driver).GetElementsFromPP5Driver(timeOut, findTypes).Cast<TOutput>();
            else
                throw new NotSupportedException("Not supported input or output type in GetElements method!");

            return elementsResult;
        }

        public static TOutput GetElementWithRetry<TInput, TOutput>(this TInput driverOrElement, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nRetry = 3)
        {
            TOutput elementResult;
            string inputTypeName = typeof(TInput).Name;
            string outputTypeName = typeof(TOutput).Name;

            if (inputTypeName is nameof(IWebElement) && outputTypeName == nameof(IWebElement))
                elementResult = (TOutput) (driverOrElement as IWebElement).GetWebElementFromWebElementRetry(findType, timeOut, nRetry);
            else if (inputTypeName is nameof(IWebElement) && outputTypeName == nameof(IElement))
                elementResult = (TOutput) (driverOrElement as IElement).GetElementFromWebElementRetry(findType, timeOut, nRetry);
            else if (inputTypeName is nameof(IElement) && outputTypeName == nameof(IElement))
                elementResult = (TOutput) (driverOrElement as IElement).GetElementFromPP5ElementRetry(findType, timeOut, nRetry);
            else if (inputTypeName is nameof(PP5Driver) && outputTypeName is nameof(IElement))
                elementResult = (TOutput)(driverOrElement as PP5Driver).GetElementFromPP5DriverRetry(findType, timeOut, nRetry);
            else
                throw new NotSupportedException("Not supported input or output type in GetElementWithRetry method!");

            return elementResult;
        }

        public static TOutput GetElementWithRetry<TInput, TOutput>(this TInput driverOrElement, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            TOutput elementResult = default(TOutput);

            //if (typeof(TInput) is IWebElement && typeof(TOutput) is IWebElement)
            //    elementResult = (TOutput)(driverOrElement as IWebElement).GetWebElementFromWebElementRetry(timeOut, findTypes);
            //else if (typeof(TInput) is IWebElement && typeof(TOutput) is IElement)
            //    elementResult = (TOutput)(driverOrElement as IElement).GetElementFromWebElementRetry(timeOut, findTypes);
            string inputTypeName = typeof(TInput).Name;
            string outputTypeName = typeof(TOutput).Name;

            if (inputTypeName is nameof(IElement) && outputTypeName == nameof(IElement))
                elementResult = (TOutput)(driverOrElement as IElement).GetElementFromPP5ElementRetry(timeOut, findTypes);
            else if (inputTypeName is nameof(PP5Driver) && outputTypeName == nameof(IElement))
                elementResult = (TOutput)(driverOrElement as PP5Driver).GetElementFromPP5DriverRetry(timeOut, findTypes);
            else
                throw new NotSupportedException("Not supported input or output type in GetElementWithRetry method!");

            return elementResult;
        }

        public static ReadOnlyCollection<TOutput> GetElementsWithRetry<TInput, TOutput>(this TInput driverOrElement, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nRetry = 3)
        {
            ReadOnlyCollection<TOutput> elementsResult;

            string inputTypeName = typeof(TInput).Name;
            string outputTypeName = typeof(TOutput).Name;

            //if (typeof(TInput) is IWebElement && typeof(TOutput) is IWebElement)
            //    elementsResult = (ReadOnlyCollection<TOutput>)(driverOrElement as IWebElement).GetWebElementsFromWebElementRetry(findType, timeOut).Cast<TOutput>();
            if (inputTypeName is nameof(IElement) && outputTypeName is nameof(IElement))
                elementsResult = (ReadOnlyCollection<TOutput>)(driverOrElement as IElement).GetElementsFromPP5ElementRetry(findType, timeOut, nRetry).Cast<TOutput>();
            else if (inputTypeName is nameof(PP5Driver) && outputTypeName is nameof(IElement))
                elementsResult = (ReadOnlyCollection<TOutput>)(driverOrElement as PP5Driver).GetElementsFromPP5DriverRetry(findType, timeOut, nRetry).Cast<TOutput>();
            else
                throw new NotSupportedException("Not supported input or output type in GetElementsWithRetry method!");

            return elementsResult;
        }

        public static ReadOnlyCollection<TOutput> GetElementsWithRetry<TInput, TOutput>(this TInput driverOrElement, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            ReadOnlyCollection<TOutput> elementsResult;

            string inputTypeName = typeof(TInput).Name;
            string outputTypeName = typeof(TOutput).Name;

            //if (typeof(TInput) is IWebElement && typeof(TOutput) is IWebElement)
            //    elementsResult = (ReadOnlyCollection<TOutput>)(driverOrElement as IWebElement).GetWebElementsFromWebElementRetry(timeOut, findTypes);
            //else if (typeof(TInput) is IWebElement && typeof(TOutput) is IElement)
            //    elementsResult = (ReadOnlyCollection<TOutput>)(driverOrElement as IElement).GetElementsFromWebElementRetry(timeOut, findTypes);
            if (inputTypeName is nameof(IElement) && outputTypeName is nameof(IElement))
                elementsResult = (ReadOnlyCollection<TOutput>)(driverOrElement as IElement).GetElementsFromPP5ElementRetry(timeOut, findTypes).Cast<TOutput>();
            else if (inputTypeName is nameof(PP5Driver) && outputTypeName is nameof(IElement))
                elementsResult = (ReadOnlyCollection<TOutput>)(driverOrElement as PP5Driver).GetElementsFromPP5DriverRetry(timeOut, findTypes).Cast<TOutput>();
            else
                throw new NotSupportedException("Not supported input or output type in GetElementsWithRetry method!");

            return elementsResult;
        }

        //GetElementFromWebElement from driver, given By locator, timeOut


        public static IWebElement GetWebElementFromWebDriver(this IWebDriver driver, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            IWebElement element = null;
            //bool isElementClickable = false;
            DefaultWait<IWebDriver> waitDriver;
            var webdriverwait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(timeOut));
            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            //var wait = new DefaultWait<WindowsDriver<WindowsElement>>(driver)
            //{
            //    Timeout = TimeSpan.FromMilliseconds(timeOut),
            //    Message = Message,
            //};

            //wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                //wait.Until(Driver =>
                //{
                //    element = Driver.FindElement(findType);
                //    //if (!element.Displayed || !element.Enabled)
                //    //    return null;

                //    //return element;
                //    return element != null;
                //});

                waitDriver = new WebDriverWait(driver, TimeSpan.FromMilliseconds(timeOut));
                waitDriver.IgnoreExceptionTypes(typeof(WebDriverException));
                element = waitDriver.Until(ElementToBeClickable(findType));
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
                //Assert.Fail(ex.Message);
            }

            return element;
        }

        // GetElementFromWebElement from driver, given By[] locators, timeOut, retry
        public static IWebElement GetWebElementFromWebDriver(this IWebDriver driver, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            IWebElement element = null;
            int CurrFindElementDepth = 1;
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                if (CurrFindElementDepth == 1)
                    element = driver.GetWebElementFromWebDriver(type, timeOut);
                else
                {
                    element = element.GetWebElementFromWebElement(type, timeOut);
                }
                CurrFindElementDepth++;
            }
            return element;
        }

        public static IElement GetElementFromPP5Driver(this PP5Driver driver, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            IElement element = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<PP5Driver>(driver, TimeSpan.FromMilliseconds(timeOut), 1);

            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    element = ElementSrc.FindElement(findType);
                    return element;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return element;
        }

        public static IElement GetElementFromWebElement(this IWebElement elementSrc, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            IElement element = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<IWebElement>(elementSrc, TimeSpan.FromMilliseconds(timeOut), 1);

            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    element = ElementSrc.FindElement(findType).ConvertToElement();
                    return element;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return element;
        }

        // GetElementFromWebElement from element, given By locator, timeOut
        public static IElement GetElementFromWebElement(this IWebElement elementSrc, ByAutomationIdOrName findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            IElement element = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<IWebElement>(elementSrc, TimeSpan.FromMilliseconds(timeOut), 1);

            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    element = ElementSrc.FindElement(findType).ConvertToElement();
                    return element;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return element;
        }

        public static IElement GetElementFromPP5Element(this IElement elementSrc, ByAutomationIdOrName findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            IElement element = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<IElement>(elementSrc, TimeSpan.FromMilliseconds(timeOut), 1);

            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    element = ElementSrc.FindElement(findType);
                    return element;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return element;
        }

        // GetElementFromWebElement from element, given By locator, timeOut
        public static IWebElement GetWebElementFromWebElement(this IWebElement elementSrc, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            IWebElement element = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<IWebElement>(elementSrc, TimeSpan.FromMilliseconds(timeOut), 1);
            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    element = ElementSrc.FindElement(findType);
                    return element;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return element;
        }

        public static IElement GetElementFromPP5Element(this IElement elementSrc, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            IElement element = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<IElement>(elementSrc, TimeSpan.FromMilliseconds(timeOut), 1);

            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    element = ElementSrc.FindElement(findType);
                    return element;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return element;
        }

        public static IWebElement GetWebElementFromPP5Element(this IElement elementSrc, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            IWebElement element = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<IElement>(elementSrc, TimeSpan.FromMilliseconds(timeOut), 1);

            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    element = ElementSrc.FindElement(findType);
                    return element;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return element;
        }

        public static IElement GetElementFromWebElement(this IWebElement elementSrc, int timeOut = SharedSetting.NORMAL_TIMEOUT, params ByAutomationIdOrName[] findTypes)
        {
            IElement element = elementSrc.ConvertToElement();
            foreach (ByAutomationIdOrName type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                element = element.GetElementFromPP5Element(type, timeOut);
            }
            return element;
        }

        // GetElementFromWebElement from element, given By[] locators, timeOut, retry
        public static IWebElement GetWebElementFromWebElement(this IWebElement elementSrc, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            IWebElement element = elementSrc;
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                element = element.GetWebElementFromWebElement(type, timeOut);
            }
            return element;
        }

        public static IWebElement GetWebElementFromPP5Element(this IElement elementSrc, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            IWebElement element = (IWebElement)elementSrc;
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                element = element.GetWebElementFromWebElement(type, timeOut);
            }
            return element;
        }

        public static IElement GetElementFromWebElement(this IWebElement elementSrc, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            IElement element = elementSrc.ConvertToElement();
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                element = element.GetElementFromPP5Element(type, timeOut);
            }
            return element;
        }

        public static IElement GetElementFromPP5Element(this IElement elementSrc, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            IElement element = elementSrc;
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                element = element.GetElementFromPP5Element(type, timeOut);
            }
            return element;
        }

        public static IElement GetElementFromPP5Driver(this PP5Driver driver, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            IElement element = null;
            int CurrFindElementDepth = 1;
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                if (CurrFindElementDepth == 1)
                    element = driver.GetElementFromPP5Driver(type, timeOut);
                else
                {
                    element = element.GetElementFromPP5Element(type, timeOut);
                }
                CurrFindElementDepth++;
            }
            return element;
        }

        public static ReadOnlyCollection<IElement> GetElementsFromPP5Driver(this PP5Driver driver, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            ReadOnlyCollection<IElement> elements = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<PP5Driver>(driver, TimeSpan.FromMilliseconds(timeOut), 1);

            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    elements = ElementSrc.FindElements(findType);
                    return elements != null;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return elements;
        }

        // GetElements from driver, given By locator, timeOut, retry
        public static ReadOnlyCollection<IWebElement> GetWebElementsFromWebDriver(this IWebDriver driver, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            ReadOnlyCollection<IWebElement> elements = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new DefaultWait<IWebDriver>(driver)
            {
                Timeout = TimeSpan.FromMilliseconds(timeOut),

                Message = Message
            };

            try
            {
                wait.Until(Driver =>
                {
                    elements = Driver.FindElements(findType);
                    return elements != null;
                });
            }

            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return elements;
        }

        // GetElements from element, given By locator, timeOut, retry
        public static ReadOnlyCollection<IWebElement> GetWebElementsFromWebElement(this IWebElement elementSrc, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            ReadOnlyCollection<IWebElement> elements = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new DefaultWait<IWebElement>(elementSrc)
            {
                Timeout = TimeSpan.FromMilliseconds(timeOut),

                Message = Message
            };

            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    elements = ElementSrc.FindElements(findType);
                    return elements != null;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return elements;
        }

        //public static ReadOnlyCollection<IWebElement> GetElements(this IWebElement elementSrc, ByAutomationIdOrName findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        //{
        //    ReadOnlyCollection<IWebElement> elements = null;

        //    // Set error message if element not found
        //    GetNotFoundMessageAndFindingText(findType, out string Message);

        //    var wait = new DefaultWait<IWebElement>(elementSrc)
        //    {
        //        Timeout = TimeSpan.FromMilliseconds(timeOut),

        //        Message = Message
        //    };

        //    wait.IgnoreExceptionTypes(typeof(WebDriverException));

        //    try
        //    {
        //        wait.Until(ElementSrc =>
        //        {
        //            elements = ElementSrc.FindElements(findType);
        //            return elements != null;
        //        });
        //    }
        //    catch (WebDriverTimeoutException ex)
        //    {
        //        Logger.LogMessage($"{ex.Message}, {Message}");
        //        return null;
        //    }
        //    return elements;
        //}

        // Integrate retry and timeout in getelement

        public static ReadOnlyCollection<IElement> GetElementsFromPP5Element(this IElement elementSrc, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            ReadOnlyCollection<IElement> elements = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<IElement>(elementSrc, TimeSpan.FromMilliseconds(timeOut), 1);

            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    elements = ElementSrc.FindElements(findType);
                    return elements != null;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return elements;
        }

        public static ReadOnlyCollection<IElement> GetElementsFromPP5Element(this IElement elementSrc, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            IElement eleTmp = elementSrc;
            ReadOnlyCollection<IElement> elements = null;
            int iter = 1;
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                if (iter < findTypes.Count())
                    eleTmp = eleTmp.GetElementFromPP5Element(type, timeOut);
                else
                    elements = eleTmp.GetElementsFromPP5Element(type, timeOut);
                iter++;
            }
            return elements;
        }

        public static ReadOnlyCollection<IElement> GetElementsFromPP5Driver(this PP5Driver driver, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            IElement element = null;
            ReadOnlyCollection<IElement> elements = null;
            int iter = 1;
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                if (iter == 1)
                    element = driver.GetElementFromPP5Driver(type, timeOut);
                else if (iter < findTypes.Count())
                    element = element.GetElementFromPP5Element(type, timeOut);
                else
                    elements = element.GetElementsFromPP5Element(type, timeOut);
                iter++;
            }
            return elements;
        }

        public static ReadOnlyCollection<IWebElement> GetWebElementsFromWebElement(this IWebElement elementSrc, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            IWebElement eleTmp = elementSrc;
            ReadOnlyCollection<IWebElement> elements = null;
            int iter = 1;
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                if (iter < findTypes.Count())
                    eleTmp = eleTmp.GetWebElementFromWebElement(type, timeOut);
                else
                    elements = eleTmp.GetWebElementsFromWebElement(type, timeOut);
                iter++;
            }
            return elements;
        }

        public static IWebElement GetWebElementFromWebElementRetry(this IWebElement elementSrc, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nRetry = 3)
        {
            IWebElement element = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<IWebElement>(elementSrc, TimeSpan.FromMilliseconds(timeOut), nRetry);
            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    element = ElementSrc.FindElement(findType);
                    return element;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return element;
        }

        public static IElement GetElementFromWebElementRetry(this IWebElement elementSrc, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nRetry = 3)
        {
            IElement element = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<IWebElement>(elementSrc, TimeSpan.FromMilliseconds(timeOut), nRetry);
            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    element = ElementSrc.FindElement(findType).ConvertToElement();
                    return element;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return element;
        }

        public static IElement GetElementFromPP5DriverRetry(this PP5Driver driver, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            IElement element = null;
            int CurrFindElementDepth = 1;
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                if (CurrFindElementDepth == 1)
                    element = driver.GetElementFromPP5DriverRetry(type, timeOut);
                else
                {
                    element = element.GetElementFromPP5ElementRetry(type, timeOut);
                }
                CurrFindElementDepth++;
            }
            return element;
        }

        public static IElement GetElementFromPP5DriverRetry(this PP5Driver driver, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nRetry = 3)
        {
            IElement element = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<PP5Driver>(driver, TimeSpan.FromMilliseconds(timeOut), nRetry);

            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    element = ElementSrc.FindElement(findType);
                    return element;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return element;
        }

        public static IElement GetElementFromPP5ElementRetry(this IElement elementSrc, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            IElement element = elementSrc;
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                element = element.GetElementFromPP5ElementRetry(type, timeOut);
            }
            return element;
        }

        public static IElement GetElementFromPP5ElementRetry(this IElement elementSrc, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nRetry = 3)
        {
            IElement element = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<IElement>(elementSrc, TimeSpan.FromMilliseconds(timeOut), nRetry);

            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    element = ElementSrc.FindElement(findType);
                    return element;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return element;
        }

        public static IElement GetElementFromPP5ElementRetry(this IElement elementSrc, ByAutomationIdOrName findType, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nRetry = 3)
        {
            IElement element = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<IElement>(elementSrc, TimeSpan.FromMilliseconds(timeOut), nRetry);

            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    element = ElementSrc.FindElement(findType);
                    return element;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return element;
        }

        public static IElement GetElementFromPP5ElementRetry(this IElement elementSrc, int timeOut = SharedSetting.NORMAL_TIMEOUT, params ByAutomationIdOrName[] findTypes)
        {
            IElement element = elementSrc;
            foreach (ByAutomationIdOrName type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                element = element.GetElementFromPP5ElementRetry(type, timeOut);
            }
            return element;
        }

        public static ReadOnlyCollection<IElement> GetElementsFromPP5DriverRetry(this PP5Driver driver, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            IElement element = null;
            ReadOnlyCollection<IElement> elements = null;
            int iter = 1;
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                if (iter == 1)
                    element = driver.GetElementFromPP5DriverRetry(type, timeOut);
                else if (iter < findTypes.Count())
                    element = element.GetElementFromPP5ElementRetry(type, timeOut);
                else
                    elements = element.GetElementsFromPP5ElementRetry(type, timeOut);
                iter++;
            }
            return elements;
        }

        public static ReadOnlyCollection<IElement> GetElementsFromPP5DriverRetry(this PP5Driver driver, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nRetry = 3)
        {
            ReadOnlyCollection<IElement> elements = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<PP5Driver>(driver, TimeSpan.FromMilliseconds(timeOut), nRetry);

            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    elements = ElementSrc.FindElements(findType);
                    return elements != null;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return elements;
        }

        public static ReadOnlyCollection<IElement> GetElementsFromPP5ElementRetry(this IElement elementSrc, int timeOut = SharedSetting.NORMAL_TIMEOUT, params By[] findTypes)
        {
            IElement eleTmp = elementSrc;
            ReadOnlyCollection<IElement> elements = null;
            int iter = 1;
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                if (iter < findTypes.Count())
                    eleTmp = eleTmp.GetElementFromPP5ElementRetry(type, timeOut);
                else
                    elements = eleTmp.GetElementsFromPP5ElementRetry(type, timeOut);
                iter++;
            }
            return elements;
        }

        public static ReadOnlyCollection<IElement> GetElementsFromPP5ElementRetry(this IElement elementSrc, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nRetry = 3)
        {
            ReadOnlyCollection<IElement> elements = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new PP5DefaultWait<IElement>(elementSrc, TimeSpan.FromMilliseconds(timeOut), nRetry);

            wait.Message = Message;
            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    elements = ElementSrc.FindElements(findType);
                    if (!elements.IsNullOrEmpty())
                    {
                        if (elements.FirstOrDefault().IsSameElement(elementSrc) && elementSrc.Name.Contains(elementSrc.ModuleName))
                        {
                            var elementsTmp = elements.ToList();
                            elementsTmp.RemoveAt(0);
                            elements = elementsTmp.AsReadOnly();
                        }
                    }
                    return elements != null;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }
            return elements;
        }



        public static IElement GetExtendedElementBySingle(this IElement elementSrc, By findList)
        {
            DefaultElementLocator locator = new DefaultElementLocator(elementSrc);
            return locator.LocateElement(new[] { findList });
        }

        public static IElement GetExtendedElementBySingle(this PP5Driver driver, By findList)
        {
            DefaultElementLocator locator = new DefaultElementLocator(driver);
            return locator.LocateElement(new[] { findList });
        }

        public static IElement GetExtendedElementByChained(this IElement elementSrc, IEnumerable<By> findList)
        {
            DefaultElementLocator locator = new DefaultElementLocator(elementSrc);
            return locator.LocateElement(findList);
        }

        public static IElement GetExtendedElementByChained(this PP5Driver driver, IEnumerable<By> findList)
        {
            DefaultElementLocator locator = new DefaultElementLocator(driver);
            return locator.LocateElement(findList);
        }

        public static IElement GetExtendedElementBySingleWithRetry(this IElement elementSrc, By findList, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            RetryingElementLocator locator = new RetryingElementLocator(elementSrc, TimeSpan.FromMilliseconds(timeOut));
            return locator.LocateElement(new[] { findList });
        }

        public static IElement GetExtendedElementBySingleWithRetry(this PP5Driver driver, By findList, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            RetryingElementLocator locator = new RetryingElementLocator(driver, TimeSpan.FromMilliseconds(timeOut));
            return locator.LocateElement(new[] { findList });
        }

        public static IElement GetExtendedElementByChainedWithRetry(this IElement elementSrc, IEnumerable<By> findList, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            RetryingElementLocator locator = new RetryingElementLocator(elementSrc, TimeSpan.FromMilliseconds(timeOut));
            return locator.LocateElement(findList);
        }

        public static IElement GetExtendedElementByChainedWithRetry(this PP5Driver driver, IEnumerable<By> findList, int timeOut = SharedSetting.NORMAL_TIMEOUT)
        {
            RetryingElementLocator locator = new RetryingElementLocator(driver, TimeSpan.FromMilliseconds(timeOut));
            return locator.LocateElement(findList);
        }


        public static IElement GetExtendedElement(this IElement elementSrc, By findType)
        {
            return elementSrc.FindElement(findType);
        }

        public static IElement GetExtendedElement(this PP5Driver driver, By findType)
        {
            return driver.FindElement(findType);
        }

        public static ReadOnlyCollection<IElement> GetExtendedElements(this IElement elementSrc, By findType)
        {
            return elementSrc.FindElements(findType);
        }

        public static ReadOnlyCollection<IElement> GetExtendedElements(this PP5Driver driver, By findType)
        {
            return driver.FindElements(findType);
        }
        #endregion

        #region GetElement with retry
        // GetElementFromWebElement from element, given By locator, timeOut, retry
        public static IWebElement GetElementWithRetry(this IWebElement elementSrc, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nTryCount = 3)
        {
            IWebElement element = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new DefaultWait<IWebElement>(elementSrc)
            {
                Timeout = TimeSpan.FromMilliseconds(timeOut),

                Message = Message
            };

            wait.IgnoreExceptionTypes(typeof(WebDriverException));
            
            try
            {
                wait.Until(ElementSrc =>
                {
                    element = ElementSrc.FindElement(findType, nTryCount);      // 20240830, Adam, add retry when getting element
                    //if (!element.Displayed || !element.Enabled)
                    //    return null;

                    return element;

                    //return element != null;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
                //Assert.Fail(ex.Message);
            }

            return element;
        }

        // GetElementFromWebElement from driver, given By[] locators, timeOut, retry
        public static IWebElement GetElementWithRetry(this IWebDriver driver, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nTryCount = 3, params By[] findTypes)
        {
            IWebElement element = null;
            //IWebElement elementSrcTemp = null;
            //DefaultWait<IWebElement> waitElement;
            //DefaultWait<IWebDriver> waitDriver;

            int CurrFindElementDepth = 1;
            foreach (By type in findTypes)
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                if (CurrFindElementDepth == 1)
                    element = driver.GetWebElementFromWebDriver(type, timeOut);
                else
                {
                    element = element.GetElementWithRetry(type, timeOut, nTryCount);
                }
                CurrFindElementDepth++;

                /*
                //// Set error message if element not found
                //GetNotFoundMessageAndFindingText(type, out string Message);

                //try
                //{
                //    // Find the element by element info
                //    if (CurrFindElementDepth == 1)
                //    {
                //        waitDriver = new DefaultWait<IWebDriver>(driver)
                //        {
                //            Timeout = TimeSpan.FromMilliseconds(timeOut),

                //            Message = Message
                //        };
                //        waitDriver.IgnoreExceptionTypes(typeof(WebDriverException));

                //        //ExpectedConditions.AlertIsPresent
                //        //waitDriver.Until(Driver =>
                //        //{
                //        //    element = (IWebElement)Driver.FindElement(type);
                //        //    //if (!element.Displayed || !element.Enabled)
                //        //    //    return null;

                //        //    return element;
                //        //    //return element != null;
                //        //});
                //        element = waitDriver.Until(ElementToBeClickable(type));

                //        //var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeOut));
                //        //element = (WindowsElement)wait.Until(drv => drv.FindElement(type));

                //        //return wait.Until(ctx => {
                //        //    var elem = ctx.FindElement(by);
                //        //    if (!elem.Displayed)
                //        //        return null;

                //        //    return elem;
                //        //});
                //    }
                //    else
                //    {
                //        waitElement = new DefaultWait<IWebElement>(elementSrcTemp)
                //        {
                //            Timeout = TimeSpan.FromMilliseconds(timeOut),

                //            Message = Message
                //        };
                //        waitElement.IgnoreExceptionTypes(typeof(WebDriverException));

                //        waitElement.Until(elementSrc =>
                //        {
                //            element = elementSrc.FindElement(type, nTryCount);
                //            //if (!element.Displayed || !element.Enabled)
                //            //    return null;

                //            return element;
                //            //return element != null;
                //        });
                //    }

                //    // Update current element source for next element finding
                //    elementSrcTemp = element;

                //    CurrFindElementDepth++;
                //}
                //catch (WebDriverTimeoutException ex)
                //{
                //    Logger.LogMessage($"{ex.Message}, {Message}");
                //    return null;
                //    //Assert.Fail(ex.Message);
                //}
                */
            }

            return element;
        }

        // GetElementFromWebElement from element, given By[] locators, timeOut, retry
        public static IWebElement GetElementWithRetry(this IWebElement elementSrc, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nTryCount = 3, params By[] findTypes)
        {
            IWebElement element = elementSrc;
            //IWebElement elementSrcTemp = elementSrc;

            foreach (By type in findTypes) 
            {
                // 20240830, Adam, modify the method to call get element method (single By)
                element = element.GetElementWithRetry(type, timeOut, nTryCount);

                //// Set error message if element not found
                //GetNotFoundMessageAndFindingText(type, out string Message);

                //// Find the element by element info
                //var wait = new DefaultWait<IWebElement>(elementSrc)
                //{
                //    Timeout = TimeSpan.FromMilliseconds(timeOut),

                //    Message = Message
                //};

                //wait.IgnoreExceptionTypes(typeof(WebDriverException));

                //try
                //{
                //    wait.Until(ElementSrc =>
                //    {
                //        element = ElementSrc.FindElement(type, nTryCount);
                //        //if (!element.Displayed || !element.Enabled)
                //        //    return null;

                //        return element;

                //        //return element != null;
                //    });

                //    // Update current element source for next element finding
                //    elementSrcTemp = element;
                //}
                //catch (WebDriverTimeoutException ex)
                //{
                //    Logger.LogMessage($"{ex.Message}, {Message}");
                //    return null;
                //    //Assert.Fail(ex.Message);
                //}
            }

            return element;
        }

        // GetElements from driver, given By locator, timeOut, retry
        public static ReadOnlyCollection<IWebElement> GetElementsWithRetry(this IWebDriver driver, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nTryCount = 3)
        {
            ReadOnlyCollection<IWebElement> elements = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new DefaultWait<IWebDriver>(driver)
            {
                Timeout = TimeSpan.FromMilliseconds(timeOut),

                Message = Message
            };

            //wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(Driver =>
                {
                    elements = Driver.FindElements(findType, nTryCount);        // 20240830, Adam, add retry when getting element
                    //if (!elements.All(e => e.Displayed))
                    //    return null;

                    //return elements;
                    return elements != null;
                });

                //var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeOut));
                //elements = wait.Until(drv => drv.FindElements(findType));
            }

            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
                //Assert.Fail(ex.Message);
            }

            return elements;
        }

        // GetElements from element, given By locator, timeOut, retry
        public static ReadOnlyCollection<IWebElement> GetElementsWithRetry(this IWebElement elementSrc, By findType, int timeOut = SharedSetting.NORMAL_TIMEOUT, int nTryCount = 3)
        {
            ReadOnlyCollection<IWebElement> elements = null;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(findType, out string Message);

            var wait = new DefaultWait<IWebElement>(elementSrc)
            {
                Timeout = TimeSpan.FromMilliseconds(timeOut),

                Message = Message
            };

            wait.IgnoreExceptionTypes(typeof(WebDriverException));

            try
            {
                wait.Until(ElementSrc =>
                {
                    elements = ElementSrc.FindElements(findType, nTryCount);      // 20240830, Adam, add retry when getting element
                    //if (!elements.All(e => e.Displayed))
                    //    return null;

                    //return elements;
                    return elements != null;
                });
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
                //Assert.Fail(ex.Message);
            }

            return elements;
        }
        #endregion

        #endregion

        #region Element related methods
        public static bool CheckElementExisted(this IWebDriver driver, By ByElementInfo, int timeOut = SharedSetting.SUPERSHORT_TIMEOUT, int sleepingInterval = 500)
        {
            IWebElement element = null;
            DefaultWait<IWebDriver> waitDriver;

            // Set error message if element not found
            GetNotFoundMessageAndFindingText(ByElementInfo, out string Message);

            try
            {
                // Find the element by element info
                waitDriver = new WebDriverWait(new SystemClock(), driver, TimeSpan.FromMilliseconds(timeOut), TimeSpan.FromMilliseconds(sleepingInterval));
                waitDriver.IgnoreExceptionTypes(typeof(WebDriverException));
                element = (IWebElement)waitDriver.Until(ElementIsVisible(ByElementInfo));
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return false;
                //Assert.Fail(ex.Message);
            }

            return element != null;
        }

        public static bool CheckElementExistedNoTimeout(this IWebDriver driver, By findType)
        {
            try
            {
                return driver.FindElement(findType) != null;
            }
            catch (Exception ex)
            {
                Logger.LogMessage(ex.Message);
                return false;
            }
        }

        public static bool CheckWindowTitle(this IWebDriver driver, string WindowTitle, int timeOut = SharedSetting.SHORT_TIMEOUT)
        {
            bool isWindowOpened = false;
            DefaultWait<IWebDriver> waitDriver;

            // Set error message if element not found
            //GetNotFoundMessageAndFindingText(ByElementInfo, out string Message);
            string Message = $"Element with Name: \"{WindowTitle}\" not found.";

            try
            {
                // Find the element by element info
                waitDriver = new WebDriverWait(driver, TimeSpan.FromMilliseconds(timeOut));
                waitDriver.IgnoreExceptionTypes(typeof(WebDriverException));
                isWindowOpened = waitDriver.Until(TitleIs(WindowTitle));
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return false;
                //Assert.Fail(ex.Message);
            }

            return isWindowOpened;
        }

        public static bool CheckElementSelected(this IWebDriver driver, IWebElement element, int timeOut = SharedSetting.EXTREMESHORT_TIMEOUT)
        {
            bool isElementSelected = false;
            DefaultWait<IWebDriver> waitDriver;

            // Set error message if element not found
            //GetNotFoundMessageAndFindingText(ByElementInfo, out string Message);
            string Message = $"Element with Name: \"{element.Text}\" not found.";

            try
            {
                // Find the element by element info
                waitDriver = new WebDriverWait(driver, TimeSpan.FromMilliseconds(timeOut));
                waitDriver.IgnoreExceptionTypes(typeof(WebDriverException));
                isElementSelected = waitDriver.Until(ElementToBeSelected(element));
            }
            catch (WebDriverTimeoutException ex)
            {
                Logger.LogMessage($"{ex.Message}, {Message}");
                return false;
                //Assert.Fail(ex.Message);
            }

            return isElementSelected;
        }

        public static bool isElementChecked(this IElement element)
        {
            //bool isElementChecked = element.GetAttribute("Value.Value") == "Checked" ? 
            //                         element.GetAttribute("Value.Value") == "Unchecked" ? 
            //                         true : false : false;

            //if (element.GetAttribute("Value.Value") == "Checked") { return true; }
            //else if (element.GetAttribute("Value.Value") == "Unchecked") { return false; }
            //else return false;
            //return element.GetAttribute("Value.Value") == "Checked";

#if DEBUG
            string toggleState = element.GetAttribute("Toggle.ToggleState");
#endif
            if (element.IsCheckBox)
                return element.GetAttribute("Toggle.ToggleState") == "1";   // Toggle.ToggleState: On (1)
            else if (element.ControlType == ElementControlType.Custom)
                return element.GetAttribute("Value.Value") == "Checked";
            else
                throw new Exception("This element is not supporting element checked property");
        }

        public static bool isElementCollapsed(this IElement element)
        {
            return element?.GetAttribute("ExpandCollapse.ExpandCollapseState") == ExpandCollapseState.Collapsed.ToString();
        }

        public static bool isElementExpanded(this IElement element)
        {
            return element?.GetAttribute("ExpandCollapse.ExpandCollapseState") == ExpandCollapseState.Expanded.ToString();
        }

        public static bool isElementAtLeafNode(this IElement element)
        {
            return element?.GetAttribute("ExpandCollapse.ExpandCollapseState") == ExpandCollapseState.LeafNode.ToString();
        }

        public static bool isElementVisible(this IElement element)
        {
            return element.hasAttribute("BoundingRectangle");
        }

        public static string GetElementName(this IElement element)
        {
            return element?.GetAttribute("Name");
        }

        public static string GetElementBoundingRectangle(this IElement element)
        {
            return element?.GetAttribute("BoundingRectangle");
        }

        public static int GetElementWidth(this IElement element)
        {
            string br = element.GetElementBoundingRectangle();
            var width = int.Parse(br.Split(' ')[2].Split(':')[1]);        // left position - right position
            return width;
        }

        public static int GetElementHeight(this IElement element)
        {
            string br = element.GetElementBoundingRectangle();
            var height = int.Parse(br.Split(' ')[3].Split(':')[1]);       // bottom position - top position
            return height;
        }
        #endregion

        //public static IAlert CheckAlertWindowOpened(this IWebDriver driver, int timeOut = 100)
        //{
        //    IAlert AlertWindow = null;
        //    DefaultWait<IWebDriver> waitDriver;

        //    // Set error message if element not found
        //    //GetNotFoundMessageAndFindingText(ByElementInfo, out string Message);
        //    //string Message = $"Element with Name: \"{element.Text}\" not found.";

        //    try
        //    {
        //        // Find the element by element info
        //        waitDriver = new WebDriverWait(driver, TimeSpan.FromMilliseconds(timeOut));
        //        waitDriver.IgnoreExceptionTypes(typeof(WebDriverException));
        //        AlertWindow = waitDriver.Until(AlertIsPresent());
        //    }
        //    catch (WebDriverTimeoutException ex)
        //    {
        //        Console.WriteLine($"{ex.Message}");
        //        //Assert.Fail(ex.Message);
        //    }

        //    return AlertWindow;
        //}

        #region Find Element with given element type methods

        #region Public Methods
        //public static string GetEditContent(this IWebElement elementSrc, string elementName) => GetChildEditElementContent(elementSrc, elementName);
        //public static string GetEditContent(this IWebElement elementSrc, int index) => GetChildEditElementContent(elementSrc, index);

        //public static string GetTextContent(this IWebElement elementSrc, string elementName) => GetChildTextElementContent(elementSrc, elementName);
        //public static string GetTextContent(this IWebElement elementSrc, int index) => GetChildTextElementContent(elementSrc, index);

        //public static string GetPaneContent(this IWebElement elementSrc, string elementName) => GetChildPaneElementContent(elementSrc, elementName);

        //public static string GetRdoBtnContent(this IWebElement elementSrc, string elementName) => GetChildRdoBtnElementContent(elementSrc, elementName);
        //public static string GetRdoBtnContent(this IWebElement elementSrc, int index) => GetChildRdoBtnElementContent(elementSrc, index);

        //public static string GetBtnContent(this IWebElement elementSrc, string elementName) => GetChildBtnElementContent(elementSrc, elementName);
        //public static string GetBtnContent(this IWebElement elementSrc, int index) => GetChildBtnElementContent(elementSrc, index);

        //public static IWebElement GetEditElement(this IWebElement elementSrc, string elementName) => GetChildEditElement(elementSrc, elementName);
        //public static IWebElement GetEditElement(this IWebElement elementSrc, int index) => GetChildEditElement(elementSrc, index);

        //public static IWebElement GetTextElement(this IWebElement elementSrc, string elementName) => GetChildTextElement(elementSrc, elementName);
        //public static IWebElement GetTextElement(this IWebElement elementSrc, int index) => GetChildTextElement(elementSrc, index);

        //public static IWebElement GetPaneElement(this IWebElement elementSrc, string elementName) => GetChildPaneElement(elementSrc, elementName);

        //public static IWebElement GetRdoBtnElement(this IWebElement elementSrc, string elementName) => GetChildRdoBtnElement(elementSrc, elementName);
        //public static IWebElement GetRdoBtnElement(this IWebElement elementSrc, int index) => GetChildRdoBtnElement(elementSrc, index);

        //public static IWebElement GetBtnElement(this IWebElement elementSrc, string elementName) => GetChildBtnElement(elementSrc, elementName);
        //public static IWebElement GetBtnElement(this IWebElement elementSrc, int index) => GetChildBtnElement(elementSrc, index);

        //public static IWebElement GetDataGridElement(this IWebElement elementSrc, string elementName) => GetChildDataGridElement(elementSrc, elementName);



        //public static string GetFirstEditContent(this IWebElement elementSrc) => GetFirstChildEditElementContent(elementSrc);

        //public static string GetFirstTextContent(this IWebElement elementSrc) => GetFirstChildTextElementContent(elementSrc);

        //public static string GetFirstPaneContent(this IWebElement elementSrc) => GetFirstChildPaneElementContent(elementSrc);

        //public static string GetFirstRdoBtnContent(this IWebElement elementSrc) => GetFirstChildRdoBtnElementContent(elementSrc);

        //public static IWebElement GetFirstEditElement(this IWebElement elementSrc) => GetFirstChildEditElement(elementSrc);

        //public static IWebElement GetFirstTextElement(this IWebElement elementSrc) => GetFirstChildTextElement(elementSrc);

        //public static IWebElement GetFirstPaneElement(this IWebElement elementSrc) => GetFirstChildPaneElement(elementSrc);

        //public static IWebElement GetFirstRdoBtnElement(this IWebElement elementSrc) => GetFirstChildRdoBtnElement(elementSrc);

        //public static IWebElement GetFirstDataGridElement(this IWebElement elementSrc) => GetFirstChildDataGridElement(elementSrc);

        //public static IWebElement GetFirstTreeViewElement(this IWebElement elementSrc) => GetFirstChildTreeViewElement(elementSrc);

        //public static IWebElement GetFirstComboBoxElement(this IWebElement elementSrc) => GetFirstChildComboBoxElement(elementSrc);
        /// <summary>
        /// Get first matched specific kind of control element or element's content
        /// </summary>
        /// <param name="elementSrc"></param>
        /// <returns></returns>
        public static IElement GetFirstTextElement(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildOfControlType(ElementControlType.TextBlock, searchType);
        }

        public static string GetFirstTextContent(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildContentOfControlType(ElementControlType.TextBlock, searchType);
        }

        public static IElement GetFirstEditElement(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildOfControlType(ElementControlType.TextBox, searchType);
        }

        public static string GetFirstEditContent(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildContentOfControlType(ElementControlType.TextBox, searchType);
        }

        public static IElement GetFirstPaneElement(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildOfControlType(ElementControlType.ScrollViewer, searchType);
        }

        public static string GetFirstPaneContent(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildContentOfControlType(ElementControlType.ScrollViewer, searchType);
        }

        public static IElement GetFirstRdoBtnElement(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildOfControlType(ElementControlType.RadioButton, searchType);
        }

        public static string GetFirstRdoBtnContent(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildContentOfControlType(ElementControlType.RadioButton, searchType);
        }

        public static IElement GetFirstDataGridElement(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildOfControlType(ElementControlType.DataGrid, searchType);
        }

        public static IElement GetFirstTreeViewElement(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildOfControlType(ElementControlType.TreeView, searchType);
        }

        public static IElement GetFirstComboBoxElement(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildOfControlType(ElementControlType.ComboBox, searchType);
        }

        public static IElement GetFirstTabControlElement(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildOfControlType(ElementControlType.TabControl, searchType);
        }

        public static IElement GetFirstCustomElement(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildOfControlType(ElementControlType.Custom, searchType);
        }

        public static IElement GetFirstListBoxItemElement(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildOfControlType(ElementControlType.ListBoxItem, searchType);
        }

        public static IElement GetFirstCheckBoxElement(this IElement elementSrc, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetFirstChildOfControlType(ElementControlType.CheckBox, searchType);
        }

        /// <summary>
        /// Get specific kind of control element or element's content by name
        /// </summary>
        /// <param name="elementSrc"></param>
        /// <param name="elementName"></param>
        /// <returns></returns>

        public static IElement GetTextElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.TextBlock, elementName, searchType);
        }

        public static string GetTextContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.TextBlock, elementName, searchType);
        }

        public static IElement GetEditElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.TextBox, elementName, searchType);
        }

        public static string GetEditContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.TextBox, elementName, searchType);
        }

        public static IElement GetPaneElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.ScrollViewer, elementName, searchType);
        }

        public static string GetPaneContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.ScrollViewer, elementName, searchType);
        }

        public static IElement GetRdoBtnElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.RadioButton, elementName, searchType);
        }

        public static string GetRdoBtnContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.RadioButton, elementName, searchType);
        }

        public static IElement GetBtnElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.Button, elementName, searchType);
        }

        public static string GetBtnContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.Button, elementName, searchType);
        }

        public static IElement GetDataGridElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.DataGrid, elementName, searchType);
        }

        public static string GetDataGridContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.DataGrid, elementName, searchType);
        }

        public static IElement GetComboBoxElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.ComboBox, elementName, searchType);
        }

        public static string GetComboBoxContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.ComboBox, elementName, searchType);
        }

        public static IElement GetCustomElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.Custom, elementName, searchType);
        }

        public static string GetCustomContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.Custom, elementName, searchType);
        }

        public static IElement GetWindowElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.Window, elementName, searchType);
        }

        public static string GetWindowContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.Window, elementName, searchType);
        }

        public static IElement GetCheckBoxElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.CheckBox, elementName, searchType);
        }

        public static string GetCheckBoxContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.CheckBox, elementName, searchType);
        }

        public static IElement GetTreeViewElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.TreeView, elementName, searchType);
        }

        public static string GetTreeViewContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.TreeView, elementName, searchType);
        }

        public static IElement GetTreeViewItemElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.TreeViewItem, elementName, searchType);
        }

        public static string GetTreeViewItemContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.TreeViewItem, elementName, searchType);
        }

        public static IElement GetTabItemElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.TabItem, elementName, searchType);
        }

        public static string GetTabItemContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.TabItem, elementName, searchType);
        }

        public static IElement GetTabControlElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.TabControl, elementName, searchType);
        }

        public static string GetListBoxItemContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.ListBoxItem, elementName, searchType);
        }

        public static IElement GetListBoxItemElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.ListBoxItem, elementName, searchType);
        }

        public static string GetListContent(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.ListBox, elementName, searchType);
        }

        public static IElement GetListElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.ListBox, elementName, searchType);
        }

        public static IElement GetToolbarElement(this IElement elementSrc, string elementName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.ToolBar, elementName, searchType);
        }


        /// <summary>
        /// Get specific kind of control element or element's content by index
        /// </summary>
        /// <param name="elementSrc"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public static IElement GetTextElement(this IElement elementSrc, int index, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.TextBlock, index, searchType);
        }

        public static string GetTextContent(this IElement elementSrc, int index, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.TextBlock, index, searchType);
        }

        public static IElement GetEditElement(this IElement elementSrc, int index, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.TextBox, index, searchType);
        }

        public static string GetEditContent(this IElement elementSrc, int index, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.TextBox, index, searchType);
        }

        public static IElement GetRdoBtnElement(this IElement elementSrc, int index, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.RadioButton, index, searchType);
        }

        public static string GetRdoBtnContent(this IElement elementSrc, int index, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.RadioButton, index, searchType);
        }

        public static IElement GetBtnElement(this IElement elementSrc, int index, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.Button, index, searchType);
        }

        public static string GetBtnContent(this IElement elementSrc, int index, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.Button, index, searchType);
        }

        public static IElement GetCustomElement(this IElement elementSrc, int index, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.Custom, index, searchType);
        }

        public static string GetCustomContent(this IElement elementSrc, int index, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.Custom, index, searchType);
        }

        public static IElement GetTabControlElement(this IElement elementSrc, int index, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.TabControl, index, searchType);
        }

        public static string GetTabControlContent(this IElement elementSrc, int index, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(ElementControlType.TabControl, index, searchType);
        }

        #endregion

        /// <summary>
        /// Get specific kind of control element by condition
        /// </summary>
        /// <param name="elementSrc"></param>
        /// <param name="condition"></param>
        /// <returns></returns>
        public static IElement GetToolbarElement(this IElement elementSrc, Func<IElement, bool> condition)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.ToolBar, condition);
        }

        public static IElement GetCustomElement(this IElement elementSrc, Func<IElement, bool> condition)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.Custom, condition);
        }

        public static IElement GetCustomElement(this IElement elementSrc, string childName, Func<IElement, bool> condition)
        {
            return elementSrc.GetSpecificChildOfControlType(ElementControlType.Custom, childName, condition);
        }

        /// <summary>
        /// Get specific kind of control element collection
        /// </summary>
        /// <param name="elementSrc"></param>
        /// <returns></returns>
        public static ReadOnlyCollection<IElement> GetDataItems(this IElement elementSrc)
        {
            return elementSrc.GetSpecificChildrenOfControlTypeBFS(ElementControlType.DataItem);
        }

        public static ReadOnlyCollection<IElement> GetTabItems(this IElement elementSrc)
        {
            return elementSrc.GetSpecificChildrenOfControlType(ElementControlType.TabItem);
        }

        public static ReadOnlyCollection<IElement> GetMenuItems(this IElement elementSrc)
        {
            return elementSrc.GetSpecificChildrenOfControlType(ElementControlType.MenuItem);
        }

        public static ReadOnlyCollection<IElement> GetTreeViewItems(this IElement elementSrc)
        {
            return elementSrc.GetSpecificChildrenOfControlType(ElementControlType.TreeViewItem);
        }

        #region Base Methods
        /// <summary>
        /// Base methods for finding different kinds of control element/elements
        /// </summary>
        /// <param name="elementSrc"></param>
        /// <returns></returns>
        public static ReadOnlyCollection<IElement> GetChildElements(this IElement elementSrc)
        {
            return elementSrc.GetElementsFromPP5Element(PP5By.XPath("*/*"));
            //return elementSrc.GetElements(By.XPath(".//child::*"));
        }

        public static int GetChildElementsCount(this IElement elementSrc)
        {
            return elementSrc.GetChildElements().Count;
        }

        // Legacy method
        public static ReadOnlyCollection<IElement> GetChildElementsOfControlType(this IElement elementSrc, ElementControlType elementType)
        {
            return elementSrc.GetChildElements()
                             .Where(e => e.ControlType == elementType)
                             .ToList().AsReadOnly();
        }

        // Legacy method
        public static List<string> GetChildElementsContentOfControlType(this IElement elementSrc, ElementControlType elementType)
        {
            //var ChildElements = elementSrc.GetChildElements();
            //var ChildElementsOfControlType = ChildElements.Where(e => e.TagName == elementType.GetDescription());
            return elementSrc.GetChildElements()
                             .Where(e => e.ControlType == elementType)
                             .Select(e => e.Text)
                             .ToList();
        }

        public static IElement GetParentElement(this PP5Driver driver, IElement elementSrc)
        {
            // Get the immediate parent:
            string xPathParent = $"//*[@RuntimeId='{((WindowsElement)elementSrc).Id}']/parent::node()";
            return driver.GetElementFromPP5Driver(PP5By.XPath(xPathParent));
        }

        public static IElement GetGrandParentElement(this PP5Driver driver, IElement elementSrc)
        {
            // Get the immediate parent:
            string xPathParent = $"//*[@RuntimeId='{((WindowsElement)elementSrc).Id}']/parent::node()/parent::node()";
            return driver.GetElementFromPP5Driver(PP5By.XPath(xPathParent));
        }

        public static string GetSpecificChildContentOfControlType(this IElement elementSrc, ElementControlType elementType, int childIndex = 0, ElementSearchType searchType = ElementSearchType.BFS)
        {
            IElement elementFound = elementSrc.GetSpecificChildOfControlType(elementType, childIndex, searchType);
            string elementContent = elementFound != null ? elementFound.Text : null;
            return elementContent;

            //if (elementSrcChildrenContents.Count == 0)
            //{
            //    return null;
            //}

            //if (childIndex >= elementSrcChildrenContents.Count || childIndex < -1)
            //{
            //    throw new IndexOutOfRangeException(childIndex.ToString());
            //}

            //return childIndex == -1 ? elementSrcChildrenContents.Last() : elementSrcChildrenContents[childIndex];
        }

        public static string GetSpecificChildContentOfControlType(this IElement elementSrc, ElementControlType elementType, string childName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            //List<string> elementSrcChildrenContents = elementSrc.GetChildElementsContentOfControlType(elementType);
            //return elementSrc.GetChildElements()
            //                 .Where(e => e.TagName == elementType.GetDescription() && e.CheckElementHasNameOrId(childName))
            //                 .Select(e => e.Text)
            //                 .FirstOrDefault();

            IElement elementFound = elementSrc.GetSpecificChildOfControlType(elementType, childName, searchType);
            string elementContent = elementFound != null ? elementFound.Text : null;
            return elementContent;
        }

        public static string GetFirstChildContentOfControlType(this IElement elementSrc, ElementControlType elementType, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(elementType, 0, searchType);
        }

        public static string GetLastChildContentOfControlType(this IElement elementSrc, ElementControlType elementType, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildContentOfControlType(elementType, -1, searchType);
        }

        public static IElement GetSpecificChildOfControlType(this IElement elementSrc, ElementControlType elementType, int childIndex = 0, ElementSearchType searchType = ElementSearchType.BFS)
        {
            if (searchType == ElementSearchType.BFS)
                return elementSrc.GetSpecificChildOfControlTypeByBFS(elementType, childIndex);
            else
                return elementSrc.GetSpecificChildOfControlTypeByDFS(elementType, childIndex);
        }

        public static IElement GetSpecificChildOfControlType(this IElement elementSrc, ElementControlType elementType, string childName, ElementSearchType searchType = ElementSearchType.BFS)
        {
            if (searchType == ElementSearchType.BFS)
                return elementSrc.GetSpecificChildOfControlTypeByBFS(elementType, childName);
            else
                return elementSrc.GetSpecificChildOfControlTypeByDFS(elementType, childName);
        }

        public static IElement GetSpecificChildOfControlType(this IElement elementSrc, ElementControlType elementType, Func<IElement, bool> condition, ElementSearchType searchType = ElementSearchType.BFS)
        {
            if (searchType == ElementSearchType.BFS)
                return elementSrc.GetSpecificChildOfControlTypeByBFS(elementType, condition);
            else
                return elementSrc.GetSpecificChildOfControlTypeByDFS(elementType, condition);
        }

        public static IElement GetSpecificChildOfControlType(this IElement elementSrc, ElementControlType elementType, string childName, Func<IElement, bool> condition, ElementSearchType searchType = ElementSearchType.BFS)
        {
            if (searchType == ElementSearchType.BFS)
                return elementSrc.GetSpecificChildOfControlTypeByBFS(elementType, childName, condition);
            else
                return elementSrc.GetSpecificChildOfControlTypeByDFS(elementType, childName, condition);
        }

        // Get Specific Child by ControlType and child index (DFS)
        public static IElement GetSpecificChildOfControlTypeByDFS(this IElement elementSrc, ElementControlType elementType, int childIndex = 0)
        {
            //ReadOnlyCollection<IElement> children = elementSrc.GetChildElements();
            IElement elementFound;
            ReadOnlyCollection<IElement> elementsMatched = elementSrc.GetElementsWithRetry<IElement, IElement>(GetLocatorByElementType(elementType));

            if (childIndex == 0)
                elementFound = elementsMatched.FirstOrDefault();
            else if (childIndex == -1)
                elementFound = elementsMatched.LastOrDefault();
            else
                elementFound = elementsMatched[childIndex];

            if (elementFound != null) { return elementFound; }

            foreach (var child in elementSrc.GetChildElements())
            {
                // Recursive call to search in child elements
                IElement childElement = GetSpecificChildOfControlTypeByDFS(child, elementType);

                // If the element is found in the child hierarchy, return it
                if (childElement != null)
                {
                    return childElement;
                }
            }

            return elementFound;

            //ReadOnlyCollection<IElement> elementSrcChildren = elementSrc.GetChildElementsOfControlType(elementType);

            //if (elementSrcChildren.Count == 0)
            //{
            //    return null;
            //}

            //if (childIndex >= elementSrcChildren.Count || childIndex < -1)
            //{
            //    throw new IndexOutOfRangeException(childIndex.ToString());
            //}

            //return childIndex == -1 ? elementSrcChildren.Last() : elementSrcChildren[childIndex];
        }

        // Get Specific Child by ControlType and element name/Id (DFS)
        public static IElement GetSpecificChildOfControlTypeByDFS(this IElement elementSrc, ElementControlType elementType, string childName)
        {
            //var elementFound = children.Where(e => e.TagName == elementType.GetDescription() && e.GetAttribute("Name") == childName)
            //                            .SingleOrDefault();

            //var elementFound = children.FirstOrDefault(e => e.TagName == elementType.GetDescription() && e.GetAttribute("Name") == childName);

            //var children = elementSrc.GetChildElements();
            //var elementFound = children.FirstOrDefault(e => e.TagName == elementType.GetDescription() && e.CheckElementHasNameOrId(childName));

            //var children = elementSrc.GetElements(By.ClassName(elementType.ToString()));
            //Console.WriteLine($"elementType: \"{elementType.ToString()}\"");

            //ReadOnlyCollection<IElement> elementsMatched = elementSrc.GetElements(By.TagName(elementType.GetDescription()));

            // Get element with Element locator (By/ByIdOrName)
            ReadOnlyCollection<IElement> elementsMatched = null;
            elementsMatched = elementSrc.GetElementsWithRetry<IElement, IElement>(GetLocatorByElementType(elementType, childName));
            //if (GetLocatorByElementType(elementType, childName) is ByAutomationIdOrName byIdOrNameLocator)
            //{
            //    elementsMatched = elementSrc.GetExtendedElements(byIdOrNameLocator);
            //}
            //else if (GetLocatorByElementType(elementType, childName) is By byLocator)
            //{
            //    elementsMatched = elementSrc.GetExtendedElements(byLocator);
            //}

            IElement elementFound = elementsMatched.FirstOrDefault(e => e.CheckElementHasNameOrId(childName));
            if (elementFound != null) { return elementFound; }
#if DEBUG
                    //Console.WriteLine($"elementSrc: {elementSrc.Text}");
                    //Console.WriteLine($"ChildElementsCount: {elementSrc.GetChildElementsCount()}");
#endif
            foreach (var child in elementSrc.GetChildElements())
            {
                // Recursive call to search in child elements
                IElement childElement = GetSpecificChildOfControlTypeByDFS(child, elementType, childName);

                // If the element is found in the child hierarchy, return it
                if (childElement != null)
                {
                    return childElement;
                }
            }

            return elementFound;

            //ReadOnlyCollection<IElement> elementSrcChildren = elementSrc.GetChildElementsOfControlType(elementType);
            //return elementSrc.GetChildElements()
            //                 .Where(e => e.TagName == elementType.GetDescription() && e.GetAttribute("Name") == childName)
            //                 .SingleOrDefault();
        }

        // Get Specific Child by ControlType and child index (BFS)
        public static IElement GetSpecificChildOfControlTypeByBFS(this IElement elementSrc, ElementControlType elementType, int childIndex = 0)
        {
            // Create a queue for BFS and enqueue the starting element
            Queue<IElement> queue = new Queue<IElement>();
            queue.Enqueue(elementSrc);

            while (queue.Count > 0)
            {
                // Dequeue the front element
                IElement currentElement = queue.Dequeue();

                // Check if the current element matches the criteria
                IElement elementFound;
                ReadOnlyCollection<IElement> elementsMatched = currentElement.GetElementsWithRetry<IElement, IElement>(GetLocatorByElementType(elementType));

                if (childIndex == 0)
                    elementFound = elementsMatched.FirstOrDefault();
                else if (childIndex == -1)
                    elementFound = elementsMatched.LastOrDefault();
                else
                    elementFound = elementsMatched[childIndex];

                if (elementFound != null) { return elementFound; }

                //#if DEBUG
                //Console.WriteLine($"currentElement: {currentElement.Text}({currentElement.TagName})");
                //Console.WriteLine($"ChildElementsCount: {currentElement.GetChildElementsCount()}");
                //#endif

                // Enqueue all child elements
                foreach (var child in currentElement.GetChildElements())
                {
                    queue.Enqueue(child);
                }
            }

            // Return null if no matching element was found
            return null;
        }

        // Assuming necessary namespaces are already included
        // Get Specific Child by ControlType and element name/Id (BFS)
        public static IElement GetSpecificChildOfControlTypeByBFS(this IElement elementSrc, ElementControlType elementType, string childName)
        {
            // Create a queue for BFS and enqueue the starting element
            Queue<IElement> queue = new Queue<IElement>();
            queue.Enqueue(elementSrc);

            while (queue.Count > 0)
            {
                // Dequeue the front element
                IElement currentElement = queue.Dequeue();

                // Check if the current element matches the criteria
                //Console.WriteLine($"TagName to query: {elementType.GetDescription()}");

                // Get element with Element locator (By/ByIdOrName)
                ReadOnlyCollection<IElement> elementsMatched = null;
                elementsMatched = currentElement.GetElementsWithRetry<IElement, IElement>(GetLocatorByElementType(elementType, childName));
                //if (GetLocatorByElementType(elementType, childName) is ByAutomationIdOrName byIdOrNameLocator)
                //{
                //    elementsMatched = currentElement.GetElementsWithTimeout(byIdOrNameLocator);
                //}
                //else if (GetLocatorByElementType(elementType, childName) is By byLocator)
                //{
                //    elementsMatched = currentElement.GetElementsWithTimeout(byLocator);
                //}

                //ReadOnlyCollection<IElement> elementsMatched = currentElement.GetElements(By.TagName(elementType.GetDescription()));

                //Logger.LogMessage($"elementsMatched size: {elementsMatched.Count}");

                IElement elementFound = elementsMatched.FirstOrDefault(e => e.CheckElementHasNameOrId(childName));
                if (elementFound != null)
                {
                    return elementFound;
                }
//#if DEBUG
                //Console.WriteLine($"currentElement: {currentElement.Text}({currentElement.TagName})");
                //Console.WriteLine($"ChildElementsCount: {currentElement.GetChildElementsCount()}");
//#endif
                // Enqueue all child elements
                foreach (var child in currentElement.GetChildElements())
                {
                    queue.Enqueue(child);
                }
            }

            // Return null if no matching element was found
            return null;
        }

        // Get Child by ControlType and condition (DFS)
        public static IElement GetSpecificChildOfControlTypeByDFS(this IElement elementSrc, ElementControlType elementType, Func<IElement, bool> condition)
        {
            //ReadOnlyCollection<IElement> children = elementSrc.GetChildElements();
            //elementFound = children.FirstOrDefault(e => e.TagName == elementType.GetDescription() && condition(e));

            // Element locator
            IElement elementFound = elementSrc.GetElementsWithRetry<IElement, IElement>(GetLocatorByElementType(elementType))
                                              .FirstOrDefault(e => condition(e));

            if (elementFound != null) { return elementFound; }

            foreach (var child in elementSrc.GetChildElements())
            {
                // Recursive call to search in child elements
                IElement childElement = child.GetSpecificChildOfControlTypeByDFS(elementType, condition);

                // If the element is found in the child hierarchy, return it
                if (childElement != null)
                {
                    return childElement;
                }
            }

            return elementFound;
        }

        // Get Child by ControlType and condition (BFS)
        public static IElement GetSpecificChildOfControlTypeByBFS(this IElement elementSrc, ElementControlType elementType, Func<IElement, bool> condition)
        {
            //ReadOnlyCollection<IElement> children = elementSrc.GetChildElements();
            //elementFound = children.FirstOrDefault(e => e.TagName == elementType.GetDescription() && condition(e));

            // Create a queue for BFS and enqueue the starting element
            Queue<IElement> queue = new Queue<IElement>();
            queue.Enqueue(elementSrc);

            while (queue.Count > 0)
            {
                // Dequeue the front element
                IElement currentElement = queue.Dequeue();
                IElement elementFound = currentElement.GetElementsWithRetry<IElement, IElement>(GetLocatorByElementType(elementType))
                                                      .FirstOrDefault(e => condition(e));

                if (elementFound != null) { return elementFound; }

                // Enqueue all child elements
                foreach (var child in currentElement.GetChildElements())
                {
                    queue.Enqueue(child);
                }
            }

            // Return an empty collection if no matching element was found
            return null;
        }

        // Get Child by ControlType, childName and condition (DFS)
        public static IElement GetSpecificChildOfControlTypeByDFS(this IElement elementSrc, ElementControlType elementType, string childName, Func<IElement, bool> condition)
        {
            //ReadOnlyCollection<IElement> children = elementSrc.GetChildElements();
            //elementFound = children.FirstOrDefault(e => e.TagName == elementType.GetDescription() && condition(e));

            // Get element with Element locator (By/ByAutomationIdOrName)
            IElement elementFound = null;
            elementFound = elementSrc.GetElementsWithRetry<IElement, IElement>(GetLocatorByElementType(elementType, childName))
                                     .FirstOrDefault(e => condition(e));

            //if (GetLocatorByElementType(elementType, childName) is ByAutomationIdOrName byIdOrNameLocator)
            //{
            //    elementFound = elementSrc.GetExtendedElements(byIdOrNameLocator)
            //                             .FirstOrDefault(e => condition(e));
            //}
            //else if(GetLocatorByElementType(elementType, childName) is By byLocator)
            //{
            //    elementFound = elementSrc.GetExtendedElements(byLocator)
            //                             .FirstOrDefault(e => condition(e));
            //}

            if (elementFound != null) { return elementFound; }

            foreach (var child in elementSrc.GetChildElements())
            {
                // Recursive call to search in child elements
                IElement childElement = child.GetSpecificChildOfControlTypeByDFS(elementType, childName, condition);

                // If the element is found in the child hierarchy, return it
                if (childElement != null)
                {
                    return childElement;
                }
            }

            return elementFound;
        }

        // Get Child by ControlType, childName and condition (BFS)
        public static IElement GetSpecificChildOfControlTypeByBFS(this IElement elementSrc, ElementControlType elementType, string childName, Func<IElement, bool> condition)
        {
            //ReadOnlyCollection<IElement> children = elementSrc.GetChildElements();
            //elementFound = children.FirstOrDefault(e => e.TagName == elementType.GetDescription() && condition(e));

            // Create a queue for BFS and enqueue the starting element
            Queue<IElement> queue = new Queue<IElement>();
            queue.Enqueue(elementSrc);

            while (queue.Count > 0)
            {
                // Dequeue the front element
                IElement currentElement = queue.Dequeue();

                // Get element with Element locator (By/ByAutomationIdOrName)
                IElement elementFound = null;
                elementFound = currentElement.GetElementsWithRetry<IElement, IElement>(GetLocatorByElementType(elementType, childName))
                                             .FirstOrDefault(e => condition(e));
                //if (GetLocatorByElementType(elementType, childName) is ByAutomationIdOrName byIdOrNameLocator)
                //{
                //    elementFound = currentElement.GetExtendedElements(byIdOrNameLocator)
                //                                 .FirstOrDefault(e => condition(e));
                //}
                //else if (GetLocatorByElementType(elementType, childName) is By byLocator)
                //{
                //    elementFound = currentElement.GetExtendedElements(byLocator)
                //                                 .FirstOrDefault(e => condition(e));
                //}

                if (elementFound != null) { return elementFound; }

                // Enqueue all child elements
                foreach (var child in currentElement.GetChildElements())
                {
                    queue.Enqueue(child);
                }
            }

            // Return an empty collection if no matching element was found
            return null;
        }

        // Get Child by ControlType and condition (BFS)
        public static ReadOnlyCollection<IElement> GetSpecificChildrenOfControlType(this IElement elementSrc, ElementControlType elementType, Func<IElement, bool> condition)
        {
            // Create a queue for BFS and enqueue the starting element
            Queue<IElement> queue = new Queue<IElement>();
            queue.Enqueue(elementSrc);

            while (queue.Count > 0)
            {
                // Dequeue the front element
                IElement currentElement = queue.Dequeue();

                // elements Found
                List<IElement> elementsFound = currentElement.GetElementsWithRetry<IElement, IElement>(GetLocatorByElementType(elementType))
                                                             .Where(condition)
                                                             .ToList();
                if (elementsFound.Count > 0)
                {
                    return elementsFound.AsReadOnly();
                }

                // Enqueue all child elements
                foreach (var child in currentElement.GetChildElements())
                {
                    queue.Enqueue(child);
                }
            }

            // Return an empty collection if no matching element was found
            return new List<IElement>().AsReadOnly();
        }

        public static IEnumerable<string> GetSpecificChildrenContentOfControlType(this IElement elementSrc, ElementControlType elementType, Func<IElement, bool> condition)
        {
            ReadOnlyCollection<IElement> elementsFound = elementSrc.GetSpecificChildrenOfControlType(elementType, condition);
            IEnumerable<string> elementContents = elementsFound != null ? elementsFound.Select(x => x.Text) : null;
            return elementContents;
        }

        // Get Children by ControlType (DFS)
        public static ReadOnlyCollection<IElement> GetSpecificChildrenOfControlType(this IElement elementSrc, ElementControlType elementType)
        {
            //var elementFound = children.Where(e => e.TagName == elementType.GetDescription() && e.GetAttribute("Name") == childName)
            //                            .SingleOrDefault();

            //var elementFound = children.FirstOrDefault(e => e.TagName == elementType.GetDescription() && e.GetAttribute("Name") == childName);

            //var children = elementSrc.GetChildElements();
            //var elementsFound = children.ToList().FindAll(e => e.TagName == elementType.GetDescription());

            try
            {
#if DEBUG
                //Logger.LogMessage($"elementSrc.TagName:{elementSrc.TagName}");
#endif
                ReadOnlyCollection<IElement> elementsFound = elementSrc.GetElementsWithRetry<IElement, IElement>(GetLocatorByElementType(elementType));

                if (elementsFound.Count > 0) { return elementsFound; }

                foreach (var child in elementSrc.GetChildElements())
                {
                    // Recursive call to search in child elements
                    ReadOnlyCollection<IElement> grandChildren = GetSpecificChildrenOfControlType(child, elementType);

                    // If the elements is found in the child hierarchy, return it
                    if (grandChildren.Count > 0)
                    {
                        return grandChildren;
                    }
                }

                //return elementsFound;
            }
            catch (ArgumentNullException ex)
            {
                string Message = $"element is: \"{elementType}\"";
                Logger.LogMessage($"{ex.Message}, {Message}");
                return null;
            }

            // Return an empty collection if no matching element was found
            return new List<IElement>().AsReadOnly();
        }

        // Get Children by ControlType (BFS)
        public static ReadOnlyCollection<IElement> GetSpecificChildrenOfControlTypeBFS(this IElement elementSrc, ElementControlType elementType)
        {
            try
            {
                // Create a queue for BFS and enqueue the starting element
                Queue<IElement> queue = new Queue<IElement>();
                queue.Enqueue(elementSrc);

                while (queue.Count > 0)
                {
                    // Dequeue the front element
                    IElement currentElement = queue.Dequeue();

                    // Get element with Element locator (By/ByAutomationIdOrName)
                    //IElement elementFound = null;
                    ReadOnlyCollection<IElement> elementsFound = currentElement.GetElementsWithRetry<IElement, IElement>(GetLocatorByElementType(elementType));

                    if (elementsFound.Count > 0) { return elementsFound; }

                    // Enqueue all child elements
                    foreach (var child in currentElement.GetChildElements())
                    {
                        queue.Enqueue(child);
                    }
                }
            }
            catch (ArgumentNullException ex)
            {
                string Message = $"element type is: \"{elementType.ToString()}\"";
                Logger.LogMessage($"{ex.Message}, {Message}");
            }

            // Return an empty collection if no matching element was found
            return new List<IElement>().AsReadOnly();
        }

        public static IElement GetFirstChildOfControlType(this IElement elementSrc, ElementControlType elementType, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(elementType, 0, searchType);
        }

        public static IElement GetLastChildOfControlType(this IElement elementSrc, ElementControlType elementType, ElementSearchType searchType = ElementSearchType.BFS)
        {
            return elementSrc.GetSpecificChildOfControlType(elementType, -1, searchType);
        }
#endregion

#endregion

        #endregion

        #region DataTable Methods

        public static ReadOnlyCollection<IElement> GetDataTableRowElements(this PP5Driver driver, string DataGridAutomationID)
        {
            return driver.GetDataTableElement(DataGridAutomationID)
                         /*.GetElements(By.XPath(".//DataItem"));*/
                         .GetDataTableRowElements();
        }

        public static IEnumerable<string> GetDataTableHeaders(this PP5Driver driver, string DataGridAutomationID)
        {
            //return driver.GetDataTableElement(DataGridAutomationID)
            //             .GetElements(By.XPath(".//HeaderItem"))
            //             .Select(tg => tg.GetElementFromWebElement(By.XPath(".//Text")).Text);
            return driver.GetDataTableElement(DataGridAutomationID)
                         .GetDataTableHeaders();
        }

        public static ReadOnlyCollection<IElement> GetDataTableRowElements(this IElement element, string DataGridAutomationID)
        {
            return element.GetDataTableElement(DataGridAutomationID)
                          /*.GetElements(By.XPath(".//DataItem"));*/
                          //.GetDataItems().AsReadOnly();
                          .GetDataTableRowElements();
        }

        public static IEnumerable<string> GetDataTableHeaders(this IElement element, string DataGridAutomationID)
        {
            //return element.GetDataTableElement(DataGridAutomationID)
            //              .GetElements(By.XPath(".//HeaderItem"))
            //              .Select(tg => tg.GetElementFromWebElement(By.XPath(".//Text")).Text);
            return element.GetDataTableElement(DataGridAutomationID)
                          .GetDataTableHeaders();
        }

        public static IElement GetDataTableElement(this IElement element, string DataGridAutomationID)
        {
            return element.GetExtendedElement(PP5By.Id(DataGridAutomationID));
        }

        public static IElement GetDataTableElement(this PP5Driver driver, string DataGridAutomationID)
        {
            return driver.GetExtendedElement(PP5By.Id(DataGridAutomationID));
        }

        public static IEnumerable<IElement> GetSelectedRowCellElements(this PP5Driver driver, string DataGridAutomationID)
        {
            return driver.GetDataTableRowElements(DataGridAutomationID).GetSelectedRowCellElements();
        }

        public static IEnumerable<IElement> GetSelectedRowCellElements(this IElement element, string DataGridAutomationID)
        {
            return element.GetDataTableRowElements(DataGridAutomationID).GetSelectedRowCellElements();
        }

        public static ReadOnlyCollection<IElement> GetDataTableRowElements(this IElement DataGridElement)
        {
            //return DataGridElement.GetElementsFromPP5Element(PP5By.XPath(".//DataItem"));
            return DataGridElement.GetDataItems();
        }

        public static IEnumerable<string> GetDataTableHeaders(this IElement DataGridElement)
        {
            //return DataGridElement.GetWebElementsFromWebElement(By.XPath(".//HeaderItem//Text"))
            //                      .Select(e => e.Text);
            return DataGridElement.GetDataTableHeaderElements()
                                  .Select(e => e.GetText());
        }

        public static IEnumerable<IElement> GetDataTableHeaderElements(this IElement DataGridElement)
        {
            //return DataGridElement.GetWebElementsFromWebElement(By.XPath(".//HeaderItem//Text"))
            //                      .Select(e => e.Text);
            return DataGridElement.GetElementsFromPP5Element(PP5By.XPath(".//HeaderItem"));
        }


        public static IEnumerable<IElement> GetSelectedRowCellElements(this IEnumerable<IElement> rowElements)
        {
            return rowElements.FirstOrDefault(r => r.Selected)
                              .GetCellElementsOfRow();
        }

        public static IEnumerable<IElement> GetCellElementsOfRow(this IElement row)
        {
            return row.GetChildElements()
                      .Where(r => r.GetAttribute("IsGridItemPatternAvailable") == bool.TrueString);
        }

        public static IElement GetCellElementByColumnIndex(this PP5Driver driver, string DataGridAutomationID, int ColumnIndex)
        {
            return driver.GetSelectedRowCellElements(DataGridAutomationID)
                         .GetCellElementByColumnIndex(ColumnIndex);
        }

        public static IElement GetCellElementByColumnIndex(this IElement element, string DataGridAutomationID, int ColumnIndex)
        {
            return element.GetSelectedRowCellElements(DataGridAutomationID)
                          .GetCellElementByColumnIndex(ColumnIndex);
        }

        public static IElement GetCellElementByColumnIndex(this IEnumerable<IElement> cellElements, int ColumnIndex)
        {
            return cellElements.FirstOrDefault(c => c.GetAttribute("GridItem.Column") == ColumnIndex.ToString());
        }

        public static int GetColumnIndexOfCellElement(this IElement cellElement)
        {
            return int.Parse(cellElement.GetAttribute("GridItem.Column"));
        }

        public static int GetRowIndexOfCellElement(this IElement cellElement)
        {
            return int.Parse(cellElement.GetAttribute("GridItem.Row"));
        }

        // Original datagrid method 
        public static IElement GetCellBy(this IElement dataGridElement, int rowNo, string colName)
        {
            //List<string> headers = dataGridElement.GetDataTableHeaders().ToList();
            List<string> headers = dataGridElement.GetDataGridHeaders();
            // Retrieve all row elements once
            ReadOnlyCollection<IElement> rows = dataGridElement.GetDataTableRowElements();

            return rows[rowNo - 1].GetCellElementsOfRow().GetCellElementByColumnIndex(headers.IndexOf(colName));
        }

        public static SelectedCellInfo GetSelectedCellBy(this PP5DataGrid dataGridElement, int rowNo, string colName)
        {
            return new SelectedCellInfo(() => GetCellBy(dataGridElement, rowNo, colName), dataGridElement);
        }

        public static PP5Element GetSelectedCell(this PP5DataGrid dataGridElement)
        {
            return new SelectedCellInfo(() => (PP5Element)dataGridElement.GetSelectedRow()?.GetCellElementsOfRow().FirstOrDefault(c => c.Selected), dataGridElement).SelectedCell;
        }

        // Adam 2024/05/20, modified datagrid method 
        public static IElement GetCellBy(this IElement dataGridElement, int rowNo, int colNo)
        {
            //List<string> headers = dataGridElement.GetDataTableHeaders().ToList();
            // Retrieve all row elements once
            ReadOnlyCollection<IElement> rows = dataGridElement.GetDataTableRowElements();

            return rows[rowNo - 1].GetCellElementsOfRow().GetCellElementByColumnIndex(colNo - 1);
        }

        // row get cell
        public static IElement GetCellBy(this IElement rowElement, int colNo)
        {
            //List<string> headers = dataGridElement.GetDataTableHeaders().ToList();
            // Retrieve all row elements once

            return rowElement.GetCellElementsOfRow().GetCellElementByColumnIndex(colNo - 1);
        }

        public static IElement GetCellByName(this IElement dataGridElement, int colNo, string cellName)
        {
            // Retrieve all row elements once
            ReadOnlyCollection<IElement> rows = dataGridElement.GetDataTableRowElements();

            //// Use LINQ to find the first matching cell
            //IElement cellFound = rows.Where(row => row.GetCellElementsOfRow().GetCellElementByColumnIndex(colNo - 1))
            //                            .FirstOrDefault(e => e.GetCellValue() == cellName);
            
            IElement cellFound = (from row in rows
                                     from cell in row.GetCellElementsOfRow()
                                     where cell.GetColumnIndexOfCellElement() == (colNo - 1)
                                     select cell)
                                     .FirstOrDefault(c => c.GetCellValue() == cellName);

            return cellFound;
        }


        public static string GetCellValue(this IElement dataGridElement, int rowNo, string colName)
        {
            // Get cell value By rowNo & colName
            IElement cell = dataGridElement.GetCellBy(rowNo, colName);
            return cell.GetCellValue();
        }

        public static string GetCellValue(this IElement dataGridElement, int rowNo, int colNo)
        {
            // Get cell value By rowNo & colNo
            IElement cell = dataGridElement.GetCellBy(rowNo, colNo);
            return cell.GetCellValue();
        }

        public static string GetCellValue(this IElement rowElement, int colNo)
        {
            // Get cell value By colNo
            IElement cell = rowElement.GetCellBy(colNo);
            return cell.GetCellValue();
        }

        public static string GetCellValueByName(this IElement dataGridElement, int colNo, string cellName)
        {
            // Get cell value By cellname
            IElement cell = dataGridElement.GetCellByName(colNo, cellName);
            return cell.GetCellValue();
        }

        public static string GetCellValue(this IElement dataGridElement, string DataGridAutomationID, int ColumnIndex)
        {
            // Get cell value By SelectedRow And ColumnIndex
            IElement cell = dataGridElement.GetSelectedRowCellElements(DataGridAutomationID)
                                              .GetCellElementByColumnIndex(ColumnIndex);
            return cell.GetCellValue();
        }

        public static string GetCellValue(this IElement cellElement)
        {
            return (cellElement as PP5Element).GetText();
            //if (cellElement == null)
            //    throw new ArgumentNullException(cellElement.ToString(), "The element passed in is null!");

            //PP5Element pp5Element = cellElement as PP5Element;
            //if (pp5Element.Value != null)
            //    return pp5Element.Value;
            //else if (pp5Element.HasTextPattern)
            //    return pp5Element.Text;
            //else if (pp5Element.HasRangeValuePattern)
            //    return pp5Element.GetAttribute("RangeValue.Value");
            //else if (!pp5Element.Name.IsEmpty())
            //    return pp5Element.Name;
            //else if (!pp5Element.AutomationId.IsEmpty())
            //    return pp5Element.AutomationId;
            //else if (!pp5Element.ClassName.IsEmpty())
            //    return pp5Element.ClassName;
            //return pp5Element.ToString();

            //string showName = !this.GetText().IsEmpty() ? !this.GetAttribute("AutomationId").IsEmpty()
            //                    ? this.GetText()
            //                    : this.GetAttribute("AutomationId")
            //                    : this.GetAttribute("ClassName");
        }

        public static IEnumerable<IElement> GetColumnBy(this IElement dataGridElement, int colNo)
        {
            // Retrieve all row elements once
            ReadOnlyCollection<IElement> rows = dataGridElement.GetDataTableRowElements();

            IEnumerable<IElement> ColumnCells = (from row in rows
                                                    from cell in row.GetCellElementsOfRow()
                                                    where cell.GetColumnIndexOfCellElement() == (colNo - 1)
                                                    select cell);

            return ColumnCells;
        }

        public static IElement GetRowByName(this IElement dataGridElement, int colNo, string cellName)
        {
            // Retrieve all row elements once and return the row with cellName matched
            ReadOnlyCollection<IElement> rows = dataGridElement.GetDataTableRowElements();  
            return rows.FirstOrDefault(row => row.GetCellValue(colNo) == cellName);
        }

        public static IElement GetRowBy(this IElement dataGridElement, int rowNo)
        {
            // Retrieve all row elements once
            return dataGridElement.GetDataTableRowElements()[rowNo - 1];
        }

        public static IElement GetSelectedRow(this IElement dataGridElement)
        {
            return dataGridElement.GetDataTableRowElements().FirstOrDefault(r => r.Selected);
        }

        public static IElement GetSelectedRow(this IElement element, string DataGridAutomationID)
        {
            return element.GetDataTableElement(DataGridAutomationID).GetSelectedRow();
        }

        public static int GetRowCount(this IElement dataGridElement)
        {
            return int.Parse(dataGridElement.GetAttribute("Grid.RowCount"));
        }

        public static int GetDisplayedRowCount(this IElement dataGridElement)
        {
            return dataGridElement.GetDataTableRowElements().Count;
        }

        public static int GetColumnCount(this IElement dataGridElement)
        {
            return dataGridElement.GetDataTableRowElements().FirstOrDefault()
                                  .GetCellElementsOfRow().Count();
        }

        public static int GetDisplayedColumnCount(this IElement dataGridElement)
        {
            return dataGridElement.GetDataTableRowElements().FirstOrDefault()
                                  .GetCellElementsOfRow().Where(c => c.IsWithinScreen).Count();
        }

        #endregion

        #region ComboBox/ListBox Methods
        // 先暫時分成兩個方法(給/不給comboBoxID)
        // 不給comboBoxID作法: combobox後再做動作
        public static void ComboBoxSelectByName(this IElement element, string name)
        {
            switch (element.ControlType)
            {
                // Scenario A: dataGrid/Cell中，點選後會開啟的列表框(ListView)
                case ElementControlType.Custom:
                    if (element.ModuleName == WindowType.Management.GetDescription())
                    {
                        element.DoubleClickWithDelay(10);                                       // Open the combobox / listbox
                        if (element.GetFirstComboBoxElement() != null)                          // Case A-1: gridcell觸發Combobox子元素 (management內適用)
                            element.GetFirstComboBoxElement().SelectComboBoxItemByName2(name);
                    }
                    else if (element.ModuleName == WindowType.TIEditor.GetDescription() ||
                             element.ModuleName == WindowType.TPEditor.GetDescription())
                    {
                        element.SelectComboBoxItemByName2(name);                            // Case A-2: gridcell觸發Listbox顯示於popup視窗，在控件中直接輸入Name選擇item (TI/TP內適用)
                        Press(Keys.Enter);
                    }
                    break;

                // Scenario B: 原本即有combobox控件的情形
                //  1.management各種多選控件
                //  2.TI / TP Editor open/ load 視窗的group多選
                //  3.Report中Category, Test Items多選控件      
                //    > 行為同下拉式組合框
                case ElementControlType.ComboBox:
                    //element.SelectComboBoxItemByName2(name);
                    element.LeftClick();
                    element.SendText(name);
                    //Press(Keys.Enter);
                    break;

                // Scenario C: 複合式視窗，但包含列表選項(ListItem)的情形:
                //  1.management的TI/TestCommand color設定，Font&Background按鈕開啟的colorPicker列表框(ListView)
                //  2.GUI Editor中，colorPicker視窗
                case ElementControlType.ListBox:
                    if (element.CheckComboBoxHasItemByName(name, out IElement cmbItem))
                        cmbItem.LeftClick();
                    break;

                // Scenario D: FileDialog filePath toolbar
                case ElementControlType.ToolBar:
                    //var br = element.GetAttribute("BoundingRectangle");
                    //Logger.LogMessage($"ToolBar tagName: {element.TagName}");
                    //var width = int.Parse(br.Split(' ')[2].Split(':')[1]);        // left position - right position
                    //var height = int.Parse(br.Split(' ')[3].Split(':')[1]);       // bottom position - top position
                    var width = element.GetElementWidth();
                    element.MoveToElementAndDoubleClickWithDelay(width / 2 + 10, 0, OpenQA.Selenium.Interactions.MoveToElementOffsetOrigin.Center, 50);           // Open the ListBox

                    Press(name);
                    Press(Keys.Enter);
                    break;
                    //if (element.GetFirstComboBoxElement() != null)                          // Case D-1: gridcell觸發Combobox子元素 (management內適用)
                    //    element.GetFirstComboBoxElement().SelectComboBoxItemByName2(name);
                    //else
                    //    element.SelectComboBoxItemByName2(name);                            // Case D-2: gridcell觸發Listbox顯示於popup視窗，在控件中直接輸入Name選擇item (TI/TP內適用)
                    //break;

                default:
                    throw new NotImplementedException();
            }
            //element.DoubleClickWithDelay(10);       // Open the combobox / listbox

            //if (element.GetAttribute("Value.IsReadOnly") == bool.FalseString)
            //{
            //    cmb = PP5IDEWindow.GetWindowElement("Popup").GetElementFromWebElement(MobileBy.AccessibilityId("PART_Content"));
            //}

            ////if (element.GetAttribute("IsValuePatternAvailable") == bool.FalseString)
            ////{

            ////}

            //ComboBoxSelectByName(name);
            //else
            //    return;
        }

        public static void ComboBoxSelectByIndex(this IElement element, int index)
        {
            switch (element.ControlType)
            {
                // Scenario A: dataGrid/Cell中，點選後會開啟的列表框(ListView)
                case ElementControlType.Custom:
                    if (element.ModuleName == WindowType.Management.GetDescription())
                    {
                        element.DoubleClickWithDelay(10);                                       // Open the combobox / listbox
                        if (element.GetFirstComboBoxElement() != null)                          // Case A-1: gridcell觸發Combobox子元素 (management內適用)
                            element.GetFirstComboBoxElement().SelectComboBoxItemByIndex2(index);
                    }
                    else if (element.ModuleName == WindowType.TIEditor.GetDescription() ||
                             element.ModuleName == WindowType.TPEditor.GetDescription())
                    {
                        //if (element.Selected)
                        //    Press(Keys.Up);
                        //else
                            element.LeftClick();
                        element.SelectComboBoxItemByIndex2(index);                            // Case A-2: gridcell觸發Listbox顯示於popup視窗，在控件中直接輸入Name選擇item (TI/TP內適用)
                    }
                    break;

                // Scenario B: 原本即有combobox控件的情形
                //  1.management各種多選控件
                //  2.TI / TP Editor open / load 視窗的group多選
                //  3.Report中Category, Test Items多選控件      
                //    > 行為同下拉式組合框
                case ElementControlType.ComboBox:
                    element.SelectComboBoxItemByIndex2(index);
                    break;

                // Scenario C: 複合式視窗，但包含列表選項(ListItem)的情形:
                //  1.management的TI/TestCommand color設定，Font&Background按鈕開啟的colorPicker列表框(ListView)
                //  2.GUI Editor中，colorPicker視窗
                case ElementControlType.ListBox:
                    element.GetComboBoxItems(out ReadOnlyCollection<IElement> cmbItems);
                    cmbItems[index].LeftClick();
                    break;

                default:
                    throw new NotImplementedException();
            }
        }
        #endregion

        #region CheckBox Methods
        //public static void TickCheckBox(this IElement dataTableCell)
        //{
        //    if (!checkBox.isElementChecked())
        //        checkBox.LeftClick();
        //}

        //public static void UnTickCheckBox(this IElement checkBox)
        //{
        //    if (checkBox.isElementChecked())
        //        checkBox.LeftClick();
        //}
        #endregion

        #region MISC

        private static void GetNotFoundMessageAndFindingText(By findType, out string Message)
        {
            Message = string.Empty;

            // Set default values
            string toFind;
            if (findType.ToString().StartsWith("By.Id"))
            {
                toFind = findType.ToString().Replace("By.Id: ", "");
                Message = $"Element with automationId: \"{toFind}\" not found.";
            }
            if (findType.ToString().StartsWith("ByAccessibilityId"))
            {
                toFind = findType.ToString().Replace("ByAccessibilityId", "").TrimStart('(').TrimEnd(')');
                Message = $"Element with automationId: \"{toFind}\" not found.";
            }
            else if (findType.ToString().StartsWith("By.Name"))
            {
                toFind = findType.ToString().Replace("By.Name: ", "");
                Message = $"Element with Name: \"{toFind}\" not found.";
            }
            else if (findType.ToString().StartsWith("By.XPath"))
            {
                toFind = findType.ToString().Replace("By.XPath: ", "");
                Message = $"Element with XPath: \"{toFind}\" not found.";
            }
            else if (findType.ToString().StartsWith("By.ClassName[Contains]"))
            {
                toFind = findType.ToString().Replace("By.ClassName[Contains]: ", "");
                Message = $"Element with ClassName: \"{toFind}\" not found.";
            }
            else if (findType.ToString().StartsWith("By.TagName"))
            {
                toFind = findType.ToString().Replace("By.TagName: ", "");
                Message = $"Element with TagName: \"{toFind}\" not found.";
            }
            else if (findType.ToString().StartsWith("ByAutomationIdOrName"))
            {
                toFind = findType.ToString().Replace("ByAutomationIdOrName", "");
                Message = $"Element with Id Or Name: \"{toFind}\" not found.";
            }
        }

        public static By GetLocatorByElementType(ElementControlType controlType, string locatorValue = "")
        {
            switch (controlType) 
            {
                case ElementControlType.Button:
                case ElementControlType.CheckBox:
                case ElementControlType.ComboBox:
                case ElementControlType.TextBox:
                case ElementControlType.Image:
                case ElementControlType.Menu:
                case ElementControlType.MenuItem:
                case ElementControlType.ProgressBar:
                case ElementControlType.RadioButton:
                case ElementControlType.ScrollBar:
                case ElementControlType.TabControl:
                case ElementControlType.TabItem:
                case ElementControlType.TextBlock:
                case ElementControlType.ToolBar:
                case ElementControlType.ToolTip:
                case ElementControlType.TreeView:
                case ElementControlType.TreeViewItem:
                case ElementControlType.Window:
                case ElementControlType.ScrollViewer:
                case ElementControlType.DatePicker:
                case ElementControlType.Calendar:
                case ElementControlType.CalendarDayButton:
                case ElementControlType.CalendarButton:
                case ElementControlType.Separator:
                    return PP5By.ClassName(controlType.ToString());

                case ElementControlType.ListBox:
                case ElementControlType.ListBoxItem:
                case ElementControlType.Header:
                case ElementControlType.HeaderItem:
                case ElementControlType.DataGrid:
                case ElementControlType.DataItem:
                case ElementControlType.Thumb:
                case ElementControlType.Group:
                case ElementControlType.Custom:
                    if (!locatorValue.IsNullOrEmpty())
                        return PP5By.IdOrName(locatorValue);
                    else
#if DEBUG
                        //Logger.LogMessage($"controlType.GetDescription():{controlType.GetDescription()}");
#endif
                        return PP5By.TagName(controlType.GetDescription());
                default:
                    return PP5By.ClassName(controlType.ToString());
            }
        }

        public static bool CheckElementHasNameOrId(this IElement element, string NameOrAutomationID)
        {
            if (NameOrAutomationID.IsNullOrEmpty())
                return false;

            if (element == null) return false;

            List<string> identifiers = new List<string> 
            {
                element.GetAttribute("Name"),
                element.GetAttribute("AutomationId")
            };

            if (identifiers.Contains(NameOrAutomationID))
                return true;
            else
                return false;
        }

        public static bool hasAttribute(this IElement element, string attributeName)
        {
            return element.GetAttribute(attributeName) != null ? true : false;
        }
#endregion
    }
}
