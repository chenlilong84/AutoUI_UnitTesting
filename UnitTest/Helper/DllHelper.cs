using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using MyCursor;
using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;
using RGiesecke.DllExport;
using System.Reflection.Emit;
using System.Diagnostics;
using Chroma.UnitTest.Common;
using System.ComponentModel;
using static PP5AutoUITests.DllHelper;

namespace PP5AutoUITests
{
    static class Constants
    {
        public const int CCH_RM_MAX_APP_NAME = 255;
        public const int CCH_RM_MAX_SVC_NAME = 63;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RM_UNIQUE_PROCESS
    {
        public int dwProcessId;
        public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct RM_PROCESS_INFO
    {
        public RM_UNIQUE_PROCESS Process;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = Constants.CCH_RM_MAX_APP_NAME + 1)]
        public string strAppName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = Constants.CCH_RM_MAX_SVC_NAME + 1)]
        public string strServiceShortName;
        public RM_APP_TYPE ApplicationType;
        public uint AppStatus;
        public uint TSSessionId;
        [MarshalAs(UnmanagedType.Bool)]
        public bool bRestartable;
    }

    public enum RM_APP_TYPE
    {
        RmUnknownApp = 0,
        RmMainWindow = 1,
        RmOtherWindow = 2,
        RmService = 3,
        RmExplorer = 4,
        RmConsole = 5,
        RmCritical = 1000
    }

    public class DllHelper
    {
        //internal const int SW_SHOWNORMAL = 1;
        private const int WM_SYSCOMMAND = 0x0112;
        private const int SC_CLOSE = 0xF060;

        const int RmRebootReasonNone = 0;

        public static void CloseWindow(IntPtr hwnd)
        {
            NativeMethods.SendMessage(hwnd, WM_SYSCOMMAND, (IntPtr)SC_CLOSE, IntPtr.Zero);
        }

        internal static void DisableMemoryMonitorWindow()
        {
            int registerValue = 0;

            const string RegisterAddress = "Software\\Chroma\\PowerPro5";
            const string RegisterKey = "MemoryCheckEnabled";

            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegisterAddress, writable: true))
                {
                    if (key != null)
                    {
                        Object o = key.GetValue(RegisterKey);
                        if (o != null && (int)o == registerValue) { }
                        else
                            key.SetValue(RegisterKey, registerValue);
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        /// <summary>
        /// Enable or disable the mouse by device Type
        /// </summary>
        /// <param name="enable"></param>
        /// <param name="deviceType">The device type</param>
        internal static bool ActivateDevice(bool enable, DeviceType deviceType)
        {
            Guid deviceGuid = new Guid();                           // The hard-coded GUID of device
            string deviceDesc = deviceType.GetDescription();        // The device description property, Gets from the properties dialog box of the device in Device Manager
            switch (deviceType)
            {
                case DeviceType.KeyPro: 
                    deviceGuid = new Guid("{36fc9e60-c465-11cf-8056-444553540000}");    // GUID of USB (USB Bus Devices (hubs and host controllers))
                    break;
                case DeviceType.Mouse:
                    deviceGuid = new Guid("{4d36e96f-e325-11ce-bfc1-08002be10318}");    // GUID of Mouse
                    break;
                case DeviceType.Keyboard:
                    deviceGuid = new Guid("{4d36e96b-e325-11ce-bfc1-08002be10318}");    // GUID of Keyboard
                    break;
            }
            return DeviceHelper.SetDeviceEnabledByName(deviceGuid, deviceDesc, enable); 
        }

        /// <summary>
        /// Enable or disable the mouse
        /// </summary>
        /// <param name="enable">To enable/disable the mouse</param>
        internal static void EnableMouse(bool enable)
        {
            // every type of device has a hard-coded GUID, this is the one for mice
            Guid mouseGuid = new Guid("{4d36e96f-e325-11ce-bfc1-08002be10318}");

            // get this from the properties dialog box of this device in Device Manager
            string instancePath = @"HID\VID_093A&PID_2510\6&3399AC26&0&0000";

            DeviceHelper.SetDeviceEnabled(mouseGuid, instancePath, enable);

            if (enable)
                Logger.LogMessage("Mouse is enabled.");
            else
                Logger.LogMessage("Mouse is disabled.");
        }

        internal static bool EnableMouseByName(bool enable)
        {
            bool result = ActivateDevice(enable, DeviceType.Mouse);

            if (result)
            {
                if (enable)
                    Logger.LogMessage("Mouse is enabled.");
                else
                    Logger.LogMessage("Mouse is disabled.");
            }
            return result;
        }

        internal class DummyDllGenerator
        {
            internal static void GenerateDummyDll(string dllName = "DummyLibrary", string dllClassName = "DummyClass", string dllFunctionName = "ExportedFunction")
            {
                // Setup the AssemblyName
                AssemblyName assemblyName = new AssemblyName(dllName);

                // Define the dynamic assembly and module
                AssemblyBuilder assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Save);
                ModuleBuilder moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule", $"{dllName}.dll");

                // Define a public class named "DummyClass" in the assembly
                TypeBuilder typeBuilder = moduleBuilder.DefineType($"{dllName}.{dllClassName}", TypeAttributes.Public | TypeAttributes.Class);

                // Add a method to be exported
                MethodBuilder methodBuilder = typeBuilder.DefineMethod(dllFunctionName, MethodAttributes.Public | MethodAttributes.Static, typeof(int), Type.EmptyTypes);
                ILGenerator ilGenerator = methodBuilder.GetILGenerator();
                ilGenerator.Emit(OpCodes.Ldc_I4, 42);  // Return an integer (42)
                ilGenerator.Emit(OpCodes.Ret);

                // Mark the method with DllExport attribute (equivalent)
                CustomAttributeBuilder dllExportAttr = new CustomAttributeBuilder(
                    typeof(DllExportAttribute).GetConstructor(new Type[] { typeof(string), typeof(CallingConvention) }),
                    new object[] { dllFunctionName, CallingConvention.StdCall }
                );
                methodBuilder.SetCustomAttribute(dllExportAttr);

                // Create the type in the assembly
                typeBuilder.CreateType();

                // Save the assembly as a DLL
                assemblyBuilder.Save($"{dllName}.dll");
            }

            [AttributeUsage(AttributeTargets.Method, Inherited = false)]
            internal class DllExportAttribute : Attribute
            {
                public string ExportName { get; }
                public CallingConvention CallingConvention { get; }

                public DllExportAttribute(string exportName, CallingConvention callingConvention)
                {
                    ExportName = exportName;
                    CallingConvention = callingConvention;
                }
            }
        }

        internal static Cursor GetCursor()
        {
            return CursorHelper.GetCursor();
        }

        public static bool IsFileLocked(string filepath)
        {
            uint handle;
            string key = Guid.NewGuid().ToString();
            string[] resources = new string[] { filepath };

            try
            {
                int result = NativeMethods.RmStartSession(out handle, 0, key);
                if (result != 0) throw new Win32Exception(result);

                try
                {
                    uint pnProcInfoNeeded = 0,
                            pnProcInfo = 0,
                            lpdwRebootReasons = RmRebootReasonNone;

                    result = NativeMethods.RmRegisterResources(handle, (uint)resources.Length, resources, 0, null, 0, null);
                    if (result != 0) throw new Win32Exception(result);

                    result = NativeMethods.RmGetList(handle, out pnProcInfoNeeded, ref pnProcInfo, null, ref lpdwRebootReasons);
                    if (result != 0) throw new Win32Exception(result);

                    return pnProcInfoNeeded > 0;
                }
                finally
                {
                    NativeMethods.RmEndSession(handle);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking file lock: {ex.Message}");
                return false;
            }
        }
    }
}
