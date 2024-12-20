using System.Collections.ObjectModel;
using System.Diagnostics;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;

namespace PP5AutoUITests
{
    public interface ISessionDriver
    {
        PP5Driver CreateDriver();

        PP5Driver AttachExistingDriver();

        PP5Driver CreateNewDriver();
    }
}