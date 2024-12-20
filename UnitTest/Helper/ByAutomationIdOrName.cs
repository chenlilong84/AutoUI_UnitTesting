using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;

namespace PP5AutoUITests.SeleniumSupport
{
    //
    // 摘要:
    //     Finds element when the id or the name attribute has the specified value.
    public class ByAutomationIdOrName : By
    {
        private string elementIdentifier = string.Empty;

        private By automationIdFinder;

        private By nameFinder;

        //
        // 摘要:
        //     Initializes a new instance of the OpenQA.Selenium.Support.PageObjects.ByAutomationIdOrName
        //     class.
        //
        // 參數:
        //   elementIdentifier:
        //     The AutomationId or Name to use in finding the element.
        public ByAutomationIdOrName(string elementIdentifier)
        {
            if (string.IsNullOrEmpty(elementIdentifier))
            {
                throw new ArgumentException("element identifier cannot be null or the empty string", "elementIdentifier");
            }

            this.elementIdentifier = elementIdentifier;
            automationIdFinder = PP5By.Id(this.elementIdentifier);
            nameFinder = PP5By.Name(this.elementIdentifier);
        }

        //
        // 摘要:
        //     Find a single element.
        //
        // 參數:
        //   context:
        //     Context used to find the element.
        //
        // 傳回:
        //     The element that matches
        public new IElement FindElement(ISearchContext context)
        {
            try
            {
                return ((WindowsElement)automationIdFinder.FindElement(context)).ConvertToElement();
            }
            catch (WebDriverException)
            {
                return ((WindowsElement)nameFinder.FindElement(context)).ConvertToElement();
            }
            //catch (NoSuchElementException)
            //{
            //    return nameFinder.FindElement(context);
            //}

            //if (automationIdFinder.FindElement(context) != null)
            //    return automationIdFinder.FindElement(context);
            //else
            //    return nameFinder.FindElement(context);
        }

        //
        // 摘要:
        //     Finds many elements
        //
        // 參數:
        //   context:
        //     Context used to find the element.
        //
        // 傳回:
        //     A readonly collection of elements that match.
        public new ReadOnlyCollection<IElement> FindElements(ISearchContext context)
        {
            List<IElement> list = new List<IElement>();
            list.AddRange(automationIdFinder.FindElements(context).Cast<WindowsElement>().ToList().AsReadOnly().ConvertToElements());
            list.AddRange(nameFinder.FindElements(context).Cast<WindowsElement>().ToList().AsReadOnly().ConvertToElements());
            return list.AsReadOnly();
        }

        //
        // 摘要:
        //     Writes out a description of this By object.
        //
        // 傳回:
        //     Converts the value of this instance to a System.String
        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture, "ByAutomationIdOrName([{0}])", elementIdentifier);
        }
    }
}
