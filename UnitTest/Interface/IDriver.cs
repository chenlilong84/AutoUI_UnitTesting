using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace PP5AutoUITests
{
    public interface IDriver<T> where T : IElement
    {
        T FindElement(By by);

        ReadOnlyCollection<T> FindElements(By by);
    }
}
