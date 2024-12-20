using System.Collections.ObjectModel;
using System.Globalization;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Enums;
using OpenQA.Selenium.Appium.Interfaces;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.PageObjects;
//using OpenQA.Selenium.Support.PageObjects;
//using static PP5AutoUITests.ElementTypeConvertor;

namespace PP5AutoUITests
{
    public abstract class PP5By : By
    {
        protected readonly string selector;
        private readonly string _searchingCriteriaName;
        protected readonly By[] bys;
        protected static By[] byss;

        internal PP5By(string selector, string searchingCriteriaName)
        {
            if (string.IsNullOrEmpty(selector))
            {
                throw new ArgumentException("Selector identifier cannot be null or the empty string", nameof(selector));
            }

            this.selector = selector;
            _searchingCriteriaName = searchingCriteriaName;
        }
        
        internal PP5By(params By[] bys)
        {
            this.bys = bys;
        }

        public static new By Id(string selector) => MobileBy.AccessibilityId(selector);

        public static new By Name(string selector) => By.Name(selector);

        public static new By ClassName(string selector) => By.ClassName(selector);

        public static new By TagName(string selector) => By.TagName(selector);

        public static new By XPath(string selector) => By.XPath(selector);

        public static ByAutomationIdOrName IdOrName(string selector) => new ByAutomationIdOrName(selector);

        //public static ByChained Chained(params By[] bys) => new ByChained(bys);

        //public static ByAll All(params By[] bys) => new ByAll(bys);

        // Override FindElement to return IElement instead of IWebElement
        public new IElement FindElement(ISearchContext context)
        {
            // Using the base class find method but cast to IElement.
            //WindowsElement webElement = (WindowsElement)base.FindElement(context);
            //return webElement.ConvertToElement(); // Assuming appropriate casting.

            if (context is IFindsByFluentSelector<IWebElement> finder)
                return ((WindowsElement)finder.FindElement(_searchingCriteriaName, selector)).ConvertToElement();
            throw new InvalidCastException($"Unable to cast {context.GetType().FullName} " +
                                           $"to {nameof(IFindsByFluentSelector<IWebElement>)}");
        }

        // Override FindElements to return ReadOnlyCollection<IElement> instead of IWebElement
        public new ReadOnlyCollection<IElement> FindElements(ISearchContext context)
        {
            if (context is IFindsByFluentSelector<IWebElement> finder)
                return finder.FindElements(_searchingCriteriaName, selector).Cast<WindowsElement>().ToList().AsReadOnly().ConvertToElements();
            throw new InvalidCastException($"Unable to cast {context.GetType().FullName} " +
                                           $"to {nameof(IFindsByFluentSelector<IWebElement>)}");

            //// Using the base class find method but cast to ReadOnlyCollection<IElement>.
            //ReadOnlyCollection<WindowsElement> webElements = base.FindElements(context).Cast<WindowsElement>().ToList().AsReadOnly();

            //// Convert ReadOnlyCollection<IWebElement> to ReadOnlyCollection<IElement>
            //return webElements.ConvertToElements();
        }
    }

    public class ById : PP5By
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ById"/> class.
        /// </summary>
        /// <param name="selector">The selector to use in finding the element.</param>
        public ById(string selector) : base(selector, MobileSelector.Accessibility)
        {
        }

        public new IElement FindElement(ISearchContext context)
        {
            if (context is IFindByAccessibilityId<IWebElement> finder)
                return ((WindowsElement)finder.FindElementByAccessibilityId(selector)).ConvertToElement();
            return base.FindElement(context);
        }

        public new ReadOnlyCollection<IElement> FindElements(ISearchContext context)
        {
            if (context is IFindByAccessibilityId<IWebElement> finder)
                return finder.FindElementsByAccessibilityId(selector).Cast<WindowsElement>().ToList().AsReadOnly().ConvertToElements();
            return base.FindElements(context);
        }

        public override string ToString() =>
            $"ByAutomationId({selector})";
    }

    public class ByName : PP5By
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ByName"/> class.
        /// </summary>
        /// <param name="selector">The selector to use in finding the element.</param>
        public ByName(string selector) : base(selector, MobileSelector.Name)
        {
        }

        public new IElement FindElement(ISearchContext context)
        {
            if (context is IFindsByName<IWebElement> finder)
                return ((WindowsElement)finder.FindElementByName(selector)).ConvertToElement();
            return base.FindElement(context);
        }

        public new ReadOnlyCollection<IElement> FindElements(ISearchContext context)
        {
            if (context is IFindsByName<IWebElement> finder)
                return finder.FindElementsByName(selector).Cast<WindowsElement>().ToList().AsReadOnly().ConvertToElements();
            return base.FindElements(context);
        }

        public override string ToString() =>
            $"ByName({selector})";
    }

    public class ByClassName : PP5By
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ByClassName"/> class.
        /// </summary>
        /// <param name="selector">The selector to use in finding the element.</param>
        public ByClassName(string selector) : base(selector, MobileSelector.ClassName)
        {
        }

        public new IElement FindElement(ISearchContext context)
        {
            return base.FindElement(context);
        }

        public new ReadOnlyCollection<IElement> FindElements(ISearchContext context)
        {
            return base.FindElements(context);
        }

        public override string ToString() =>
            $"ByClassName({selector})";
    }

    public class ByTagName : PP5By
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ByTagName"/> class.
        /// </summary>
        /// <param name="selector">The selector to use in finding the element.</param>
        public ByTagName(string selector) : base(selector, MobileSelector.TagName)
        {
        }

        public new IElement FindElement(ISearchContext context)
        {
            return base.FindElement(context);
        }

        public new ReadOnlyCollection<IElement> FindElements(ISearchContext context)
        {
            return base.FindElements(context);
        }

        public override string ToString() =>
            $"ByTagName({selector})";
    }

    public class ByXPath : PP5By
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ByXPath"/> class.
        /// </summary>
        /// <param name="selector">The selector to use in finding the element.</param>
        public ByXPath(string selector) : base(selector, "xpath")
        {
        }

        public new IElement FindElement(ISearchContext context)
        {
            return base.FindElement(context);
        }

        public new ReadOnlyCollection<IElement> FindElements(ISearchContext context)
        {
            return base.FindElements(context);
        }

        public override string ToString() =>
            $"ByXPath({selector})";
    }

    /// <summary>
    /// Mechanism used to locate elements within a document using a series of other lookups. <br/>
    /// This class will find all DOM elements that matches each of the locators in sequence
    /// <para> The following code will will find all elements that match by2 and appear under an element that matches by1.</para>
    /// <example>
    /// <code>
    /// driver.findElements(new ByChained(by1, by2))
    /// </code>
    /// </example>
    /// </summary>
    public class ByChained : PP5By
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ByChained"/> class with one or more <see cref="By"/> objects.
        /// </summary>
        /// <param name="bys">One or more <see cref="By"/> references</param>
        public ByChained(params By[] bys) : base(bys)
        {

        }

        /// <summary>
        /// Find a single element.
        /// </summary>
        /// <param name="context">Context used to find the element.</param>
        /// <returns>The element that matches</returns>
        public new IElement FindElement(ISearchContext context)
        {
            ReadOnlyCollection<IElement> elements = this.FindElements(context);
            if (elements.Count == 0)
            {
                throw new NoSuchElementException("Cannot locate an element using " + this.ToString());
            }

            return elements[0];
        }

        /// <summary>
        /// Finds many elements
        /// </summary>
        /// <param name="context">Context used to find the element.</param>
        /// <returns>A readonly collection of elements that match.</returns>
        public new ReadOnlyCollection<IElement> FindElements(ISearchContext context)
        {
            if (this.bys.Length == 0)
            {
                return new List<IElement>().AsReadOnly();
            }

            List<IElement> elems = null;
            foreach (By by in this.bys)
            {
                List<IElement> newElems = new List<IElement>();

                if (elems == null)
                {
                    newElems.AddRange(by.FindElements(context).Cast<WindowsElement>().ToList().AsReadOnly().ConvertToElements());
                }
                else
                {
                    foreach (IElement elem in elems)
                    {
                        newElems.AddRange(elem.FindElements(by));
                    }
                }

                elems = newElems;
            }

            return elems.AsReadOnly();
        }

        /// <summary>
        /// Writes out a comma separated list of the <see cref="By"/> objects used in the chain.
        /// </summary>
        /// <returns>Converts the value of this instance to a <see cref="string"/></returns>
        public override string ToString()
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (By by in this.bys)
            {
                if (stringBuilder.Length > 0)
                {
                    stringBuilder.Append(",");
                }

                stringBuilder.Append(by);
            }

            return string.Format(CultureInfo.InvariantCulture, "By.Chained([{0}])", stringBuilder.ToString());
        }
    }

    /// <summary>
    /// Mechanism used to locate elements within a document using a series of lookups. <br/>
    /// This class will find all DOM elements that matches all of the locators in sequence, e.g.
    /// <para>The following code will find all elements that match by1 and then all elements that also match by2.</para>
    /// <example>
    /// <code>
    /// driver.findElements(new ByAll(by1, by2))
    /// </code>
    /// This means that the list of elements returned may not be in document order.
    /// </example>
    /// </summary>
    public class ByAll : PP5By
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ByAll"/> class with one or more <see cref="By"/> objects.
        /// </summary>
        /// <param name="bys">One or more <see cref="By"/> references</param>
        public ByAll(params By[] bys) : base(bys)
        {

        }

        /// <summary>
        /// Find a single element.
        /// </summary>
        /// <param name="context">Context used to find the element.</param>
        /// <returns>The element that matches</returns>
        public new IElement FindElement(ISearchContext context)
        {
            var elements = this.FindElements(context);
            if (elements.Count == 0)
            {
                throw new NoSuchElementException("Cannot locate an element using " + this.ToString());
            }

            return elements[0];
        }

        /// <summary>
        /// Finds many elements
        /// </summary>
        /// <param name="context">Context used to find the element.</param>
        /// <returns>A readonly collection of elements that match.</returns>
        public new ReadOnlyCollection<IElement> FindElements(ISearchContext context)
        {
            if (this.bys.Length == 0)
            {
                return new List<IElement>().AsReadOnly();
            }

            IEnumerable<IElement> elements = null;
            foreach (PP5By by in this.bys)
            {
                ReadOnlyCollection<IElement> foundElements = by.FindElements(context);
                if (foundElements.Count == 0)
                {
                    // Optimization: If at any time a find returns no elements, the
                    // only possible result for find-all is an empty collection.
                    return new List<IElement>().AsReadOnly();
                }

                if (elements == null)
                {
                    elements = foundElements;
                }
                else
                {
                    elements = elements.Intersect(by.FindElements(context));
                }
            }

            return elements.ToList().AsReadOnly();
        }

        /// <summary>
        /// Writes out a comma separated list of the <see cref="By"/> objects used in the chain.
        /// </summary>
        /// <returns>Converts the value of this instance to a <see cref="string"/></returns>
        public override string ToString()
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (By by in this.bys)
            {
                if (stringBuilder.Length > 0)
                {
                    stringBuilder.Append(",");
                }

                stringBuilder.Append(by);
            }

            return string.Format(CultureInfo.InvariantCulture, "By.All([{0}])", stringBuilder.ToString());
        }
    }

    /// <summary>
    /// 根據指定的值查找元素，當 id 或 name 屬性符合時。
    /// </summary>
    public class ByAutomationIdOrName : By
    {
        private string elementIdentifier = string.Empty;

        private By automationIdFinder;
        private By nameFinder;

        /// <summary>
        /// 初始化 PP5AutoUITests.ByAutomationIdOrName 類的新實例。
        /// </summary>
        /// <param name="elementIdentifier">用於查找元素的 AutomationId 或 Name。</param>
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

        /// <summary>
        /// 查找單個元素。
        /// </summary>
        /// <param name="context">用於查找元素的上下文。</param>
        /// <returns>匹配的元素。</returns>
        public override IWebElement FindElement(ISearchContext context)
        {
            try
            {
                return automationIdFinder.FindElement(context);
            }
            catch (WebDriverException)
            {
                return nameFinder.FindElement(context);
            }
        }

        /// <summary>
        /// 查找多個元素。
        /// </summary>
        /// <param name="context">用於查找元素的上下文。</param>
        /// <returns>匹配元素的只讀集合。</returns>
        public override ReadOnlyCollection<IWebElement> FindElements(ISearchContext context)
        {
            List<IWebElement> list = new List<IWebElement>();
            list.AddRange(automationIdFinder.FindElements(context));
            list.AddRange(nameFinder.FindElements(context));
            return list.AsReadOnly();
        }

        /// <summary>
        /// 輸出此 By 物件的描述。
        /// </summary>
        /// <returns>將此實例的值轉換為 System.String。</returns>
        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture, "ByAutomationIdOrName([{0}])", elementIdentifier);
        }
    }

}