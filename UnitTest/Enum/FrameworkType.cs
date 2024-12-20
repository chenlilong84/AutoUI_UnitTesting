using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PP5AutoUITests
{
    /// <summary>
    /// The framework type of the element
    /// </summary>
    public enum FrameworkType
    {
        /// <summary>
        /// WPF framework (mainly use this)
        /// </summary>
        WPF = 0,
        /// <summary>
        /// Winform framework
        /// </summary>
        WinForm,
        /// <summary>
        /// Win32 aplication (FileDialog)
        /// </summary>
        Win32
    }
}
