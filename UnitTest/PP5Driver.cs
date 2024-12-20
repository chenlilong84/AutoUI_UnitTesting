using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Appium.Service;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium;
using System.Globalization;
using Chroma.UnitTest.Common;

namespace PP5AutoUITests
{
    public class PP5Driver<W> : WindowsDriver<W>, IDriver<W> where W : IElement
    {
        /// <summary>
        /// Initializes a new instance of the PP5Driver class using Appium options
        /// </summary>
        /// <param name="AppiumOptions">An <see cref="AppiumOptions"/> object containing the Appium options of the browser.</param>
        public PP5Driver(AppiumOptions AppiumOptions)
            : base(AppiumOptions)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using Appium options and command timeout
        /// </summary>
        /// <param name="AppiumOptions">An <see cref="ICapabilities"/> object containing the Appium options.</param>
        /// <param name="commandTimeout">The maximum amount of time to wait for each command.</param>
        public PP5Driver(AppiumOptions AppiumOptions, TimeSpan commandTimeout)
            : base(AppiumOptions, commandTimeout)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using the AppiumServiceBuilder instance and Appium options
        /// </summary>
        /// <param name="builder"> object containing settings of the Appium local service which is going to be started</param>
        /// <param name="AppiumOptions">An <see cref="ICapabilities"/> object containing the Appium options.</param>
        public PP5Driver(AppiumServiceBuilder builder, AppiumOptions AppiumOptions)
            : base(builder, AppiumOptions)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using the AppiumServiceBuilder instance, Appium options and command timeout
        /// </summary>
        /// <param name="builder"> object containing settings of the Appium local service which is going to be started</param>
        /// <param name="AppiumOptions">An <see cref="ICapabilities"/> object containing the Appium options.</param>
        /// <param name="commandTimeout">The maximum amount of time to wait for each command.</param>
        public PP5Driver(AppiumServiceBuilder builder, AppiumOptions AppiumOptions, TimeSpan commandTimeout)
            : base(builder, AppiumOptions, commandTimeout)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using the specified remote address and Appium options
        /// </summary>
        /// <param name="remoteAddress">URI containing the address of the WebDriver remote server (e.g. http://127.0.0.1:4723/wd/hub).</param>
        /// <param name="AppiumOptions">An <see cref="AppiumOptions"/> object containing the Appium options.</param>
        public PP5Driver(Uri remoteAddress, AppiumOptions AppiumOptions)
            : base(remoteAddress, AppiumOptions)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using the specified Appium local service and Appium options
        /// </summary>
        /// <param name="service">the specified Appium local service</param>
        /// <param name="AppiumOptions">An <see cref="ICapabilities"/> object containing the Appium options of the browser.</param>
        public PP5Driver(AppiumLocalService service, AppiumOptions AppiumOptions)
            : base(service, AppiumOptions)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using the specified remote address, Appium options, and command timeout.
        /// </summary>
        /// <param name="remoteAddress">URI containing the address of the WebDriver remote server (e.g. http://127.0.0.1:4723/wd/hub).</param>
        /// <param name="AppiumOptions">An <see cref="AppiumOptions"/> object containing the Appium options.</param>
        /// <param name="commandTimeout">The maximum amount of time to wait for each command.</param>
        public PP5Driver(Uri remoteAddress, AppiumOptions AppiumOptions, TimeSpan commandTimeout)
            : base(remoteAddress, AppiumOptions, commandTimeout)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using the specified Appium local service, Appium options, and command timeout.
        /// </summary>
        /// <param name="service">the specified Appium local service</param>
        /// <param name="AppiumOptions">An <see cref="ICapabilities"/> object containing the Appium options.</param>
        /// <param name="commandTimeout">The maximum amount of time to wait for each command.</param>
        public PP5Driver(AppiumLocalService service, AppiumOptions AppiumOptions, TimeSpan commandTimeout)
            : base(service, AppiumOptions, commandTimeout)
        {
        }

        public new IElement FindElement(By by)
        {
            var windowsElement = base.FindElement(by) as WindowsElement;
            return windowsElement.ConvertToElement();
        }

        public new ReadOnlyCollection<IElement> FindElements(By by)
        {
            var windowsElements = base.FindElements(by);
            List<IElement> elements = new List<IElement>();
            foreach (var element in windowsElements)
            {
                elements.Add((element as WindowsElement).ConvertToElement()); // Assuming appropriate casting.
            }
            return new ReadOnlyCollection<IElement>(elements);
        }

        //private IElement ConvertToElement(WindowsElement windowsElement)
        //{
        //    return new Element(windowsElement);
        //}
    }

    public class PP5Driver : WindowsDriver<WindowsElement>, IDriver<IElement>
    {
        #region Properties
        public SessionType sessionType { get; set; }
        #endregion

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using Appium options
        /// </summary>
        /// <param name="AppiumOptions">An <see cref="AppiumOptions"/> object containing the Appium options of the browser.</param>
        public PP5Driver(AppiumOptions AppiumOptions)
            : base(AppiumOptions)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using Appium options and command timeout
        /// </summary>
        /// <param name="AppiumOptions">An <see cref="ICapabilities"/> object containing the Appium options.</param>
        /// <param name="commandTimeout">The maximum amount of time to wait for each command.</param>
        public PP5Driver(AppiumOptions AppiumOptions, TimeSpan commandTimeout)
            : base(AppiumOptions, commandTimeout)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using the AppiumServiceBuilder instance and Appium options
        /// </summary>
        /// <param name="builder"> object containing settings of the Appium local service which is going to be started</param>
        /// <param name="AppiumOptions">An <see cref="ICapabilities"/> object containing the Appium options.</param>
        public PP5Driver(AppiumServiceBuilder builder, AppiumOptions AppiumOptions)
            : base(builder, AppiumOptions)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using the AppiumServiceBuilder instance, Appium options and command timeout
        /// </summary>
        /// <param name="builder"> object containing settings of the Appium local service which is going to be started</param>
        /// <param name="AppiumOptions">An <see cref="ICapabilities"/> object containing the Appium options.</param>
        /// <param name="commandTimeout">The maximum amount of time to wait for each command.</param>
        public PP5Driver(AppiumServiceBuilder builder, AppiumOptions AppiumOptions, TimeSpan commandTimeout)
            : base(builder, AppiumOptions, commandTimeout)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using the specified remote address and Appium options
        /// </summary>
        /// <param name="remoteAddress">URI containing the address of the WebDriver remote server (e.g. http://127.0.0.1:4723/wd/hub).</param>
        /// <param name="AppiumOptions">An <see cref="AppiumOptions"/> object containing the Appium options.</param>
        public PP5Driver(Uri remoteAddress, AppiumOptions AppiumOptions, SessionType _sessionType)
            : base(remoteAddress, AppiumOptions)
        {
            sessionType = _sessionType;
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using the specified Appium local service and Appium options
        /// </summary>
        /// <param name="service">the specified Appium local service</param>
        /// <param name="AppiumOptions">An <see cref="ICapabilities"/> object containing the Appium options of the browser.</param>
        public PP5Driver(AppiumLocalService service, AppiumOptions AppiumOptions)
            : base(service, AppiumOptions)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using the specified remote address, Appium options, and command timeout.
        /// </summary>
        /// <param name="remoteAddress">URI containing the address of the WebDriver remote server (e.g. http://127.0.0.1:4723/wd/hub).</param>
        /// <param name="AppiumOptions">An <see cref="AppiumOptions"/> object containing the Appium options.</param>
        /// <param name="commandTimeout">The maximum amount of time to wait for each command.</param>
        public PP5Driver(Uri remoteAddress, AppiumOptions AppiumOptions, TimeSpan commandTimeout, SessionType _sessionType)
            : base(remoteAddress, AppiumOptions, commandTimeout)
        {
            sessionType = _sessionType;
        }

        /// <summary>
        /// Initializes a new instance of the PP5Driver class using the specified Appium local service, Appium options, and command timeout.
        /// </summary>
        /// <param name="service">the specified Appium local service</param>
        /// <param name="AppiumOptions">An <see cref="ICapabilities"/> object containing the Appium options.</param>
        /// <param name="commandTimeout">The maximum amount of time to wait for each command.</param>
        public PP5Driver(AppiumLocalService service, AppiumOptions AppiumOptions, TimeSpan commandTimeout)
            : base(service, AppiumOptions, commandTimeout)
        {
        }

        public new IElement FindElement(By by)
        {
            var windowsElement = base.FindElement(by);
            return windowsElement.ConvertToElement();
        }

        public new ReadOnlyCollection<IElement> FindElements(By by)
        {
            var windowsElements = base.FindElements(by);
            List<IElement> elements = new List<IElement>();
            foreach (var element in windowsElements)
            {
                elements.Add((element).ConvertToElement()); // Assuming appropriate casting.
            }
            return new ReadOnlyCollection<IElement>(elements);
        }

        public string GetToolTipMessage(PP5Element element)
        {
            if (this == null)
                throw new ArgumentNullException(this.GetType().ToString(), "The element is null!");
            
            if (this.sessionType != SessionType.Desktop)
                throw new Exception("Wrong session type of this driver! Only SessionType.Desktop is allowed.");

            MoveToElementOffsetStartingPoint offsetStartingPoint = MoveToElementOffsetStartingPoint.Center;
            if (element.IsTextBox)
                offsetStartingPoint = MoveToElementOffsetStartingPoint.OuterCenterLeft;
            else if (element.IsGridCell)
                offsetStartingPoint = MoveToElementOffsetStartingPoint.InnerCenterLeft;
            element.MoveToElement(offsetStartingPoint);
            return this.PerformGetElement("/ByClass#Retry[Popup]/ToolTip[0]").GetText();
        }

        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture, "Driver ({0})", this.sessionType);
        }

        //private IElement ConvertToElement(WindowsElement windowsElement)
        //{
        //    return new Element(windowsElement);
        //}
    }
}
