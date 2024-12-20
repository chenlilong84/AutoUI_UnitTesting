using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using static System.Net.Mime.MediaTypeNames;

namespace PP5AutoUITests.SeleniumSupport
{
    /// <summary>
    /// Provides the ability to wait for an arbitrary condition during test execution.
    /// Reference By: OpenQA.Selenium.Support.UI.WebDriverWait
    /// </summary>
    /// <example>
    /// <code>
    /// IWait wait = new WebElementWait(element, TimeSpan.FromSeconds(3))
    /// IWebElement element = wait.Until(element => element.FindElement(By.Name("q")));
    /// </code>
    /// </example>
    public class PP5DefaultWait<TInput> : DefaultWait<TInput>
    {
        private IClock clock;
        private TInput input;
        private List<Type> ignoredExceptions = new List<Type>();
        private int nTryCount;

        /// <summary>
        /// Initializes a new instance of the <see cref="WebElementWait"/> class.
        /// </summary>
        /// <param name="element">The WebElement instance used to wait.</param>
        /// <param name="timeout">The timeout value indicating how long to wait for the condition.</param>
        public PP5DefaultWait(TInput input, TimeSpan timeout)
            : this(new SystemClock(), input, timeout, DefaultSleepTimeout, DefaultRetryCount)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WebElementWait"/> class.
        /// </summary>
        /// <param name="element">The WebElement instance used to wait.</param>
        /// <param name="timeout">The timeout value indicating how long to wait for the condition.</param>
        /// <param name="_nTryCount">The retry count indicating how many times to retry for the condition.</param>
        public PP5DefaultWait(TInput input, TimeSpan timeout, int _nTryCount)
            : this(new SystemClock(), input, timeout, DefaultSleepTimeout, _nTryCount)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WebElementWait"/> class.
        /// </summary>
        /// <param name="clock">An object implementing the <see cref="IClock"/> interface used to determine when time has passed.</param>
        /// <param name="element">The WebElement instance used to wait.</param>
        /// <param name="timeout">The timeout value indicating how long to wait for the condition.</param>
        /// <param name="sleepInterval">A <see cref="TimeSpan"/> value indicating how often to check for the condition to be true.</param>
        /// <param name="_nTryCount">The retry count indicating how many times to retry for the condition.</param>
        public PP5DefaultWait(IClock clock, TInput input, TimeSpan timeout, TimeSpan sleepInterval, int _nTryCount)
            : base(input)
        {
            this.clock = clock;
            this.input = input;
            this.Timeout = timeout;
            this.PollingInterval = sleepInterval;
            this.nTryCount = _nTryCount;
        }

        private static int DefaultRetryCount
        {
            get { return 3; }
        }

        private static TimeSpan DefaultSleepTimeout
        {
            get { return TimeSpan.FromMilliseconds(500); }
        }

        private bool IsIgnoredException(Exception exception)
        {
            return ignoredExceptions.Any((Type type) => type.IsAssignableFrom(exception.GetType()));
        }

        protected override void ThrowTimeoutException(string exceptionMessage, Exception lastException)
        {
            throw new WebDriverTimeoutException(exceptionMessage, lastException);
        }

        public new void IgnoreExceptionTypes(params Type[] exceptionTypes)
        {
            if (exceptionTypes == null)
            {
                throw new ArgumentNullException("exceptionTypes", "exceptionTypes cannot be null");
            }

            foreach (Type c in exceptionTypes)
            {
                if (!typeof(Exception).IsAssignableFrom(c))
                {
                    throw new ArgumentException("All types to be ignored must derive from System.Exception", "exceptionTypes");
                }
            }

            this.ignoredExceptions.AddRange(exceptionTypes);
        }



        /// <summary>
        /// Repeatedly applies this instance's input value to the given function until one of the following
        /// occurs:
        /// <para>
        /// <list type="bullet">
        /// <item>the function returns neither null nor false</item>
        /// <item>the function throws an exception that is not in the list of ignored exception types</item>
        /// <item>the timeout expires</item>
        /// <item>the retry count reached</item>
        /// </list>
        /// </para>
        /// </summary>
        /// <typeparam name="TResult">The delegate's expected return type.</typeparam>
        /// <param name="condition">A delegate taking an object of type IWebElement as its parameter, and returning a TResult.</param>
        /// <returns>The delegate's return value.</returns>
        public new TOutput Until<TOutput>(Func<TInput, TOutput> condition)
        {
            if (condition == null)
            {
                throw new ArgumentNullException("condition", "condition cannot be null");
            }

            Type typeFromHandle = typeof(TOutput);
            if ((typeFromHandle.IsValueType && typeFromHandle != typeof(bool)) || !typeof(object).IsAssignableFrom(typeFromHandle))
            {
                throw new ArgumentException("Can only wait on an object or boolean response, tried to use type: " + typeFromHandle.ToString(), "condition");
            }

            Exception lastException = null;
            DateTime otherDateTime = this.clock.LaterBy(base.Timeout);
            // TResultOutput is a class or interface type, default(TResult) is the null reference.
            int nRetryCounter = 0;
            while (true)
            {
                try
                {
                    TOutput val = condition(this.input);
                    if (typeFromHandle == typeof(bool))
                    {
                        bool? flag = val as bool?;
                        if (flag.HasValue && flag.Value)
                        {
                            //Logger.LogMessage($"val: {val}");
                            return val;
                        }
                    }
                    else if (val != null)
                    {
                        //Logger.LogMessage($"val: {val}");
                        return val;
                    }
                }
                catch (Exception ex)
                {
                    if (!IsIgnoredException(ex))
                    {
                        //Logger.LogMessage($"IsIgnoredException(ex) is False");
                        //Logger.LogMessage(ex.Message);
                        throw;
                    }

                    lastException = ex;
                }

                if (!this.clock.IsNowBefore(otherDateTime))
                {
                    string text = string.Format(CultureInfo.InvariantCulture, "Timed out after {0} seconds, retry count {1}", base.Timeout.TotalSeconds, nRetryCounter++);
                    if (!string.IsNullOrEmpty(base.Message))
                    {
                        text = text + ": " + base.Message;
                    }

                    Logger.LogMessage(text, lastException);
                    if (nRetryCounter != nTryCount)
                        continue;

                    ThrowTimeoutException(text, lastException);
                }

                Thread.Sleep(base.PollingInterval);
            }
        }
    }
}
