using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PP5AutoUITests
{
    /// <summary>
    /// 定義要啟用/關閉的 Device 種類
    /// </summary>
    public enum DeviceType : int
    {
        /// <summary>
        /// USB: KeyPro
        /// </summary>
        [System.ComponentModel.Description("Sentinel USB Key")]
        KeyPro = 0,
        /// <summary>
        /// Mouse
        /// </summary>
        [System.ComponentModel.Description("HID-compliant mouse")]
        Mouse,
        /// <summary>
        /// Keyboard
        /// </summary>
        [System.ComponentModel.Description("HID Keyboard Device")]
        Keyboard,
    }
}
