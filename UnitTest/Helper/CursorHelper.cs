using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.ComponentModel;
using Microsoft.Win32.SafeHandles;
using System.Security;
using System.Runtime.ConstrainedExecution;
using System.Management;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;

namespace MyCursor
{
    public enum Cursor
    {
        UNKOWNCURSOR,
        AppStarting,
        Arrow,
        Cross,
        Default,
        IBeam,
        No,
        SizeAll,
        SizeNESW,
        SizeNS,
        SizeNWSE,
        SizeWE,
        UpArrow,
        WaitCursor,
        Help,
        HSplit,
        VSplit,
        NoMove2D,
        NoMoveHoriz,
        NoMoveVert,
        PanEast,
        PanNE,
        PanNorth,
        PanNW,
        PanSE,
        PanSouth,
        PanSW,
        PanWest,
        Hand
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct POINT
    {
        public Int32 x;
        public Int32 y;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct CURSORINFO
    {
        public Int32 cbSize;        // Specifies the size, in bytes, of the structure. 
                                    // The caller must set this to Marshal.SizeOf(typeof(CURSORINFO)).
        public Int32 flags;         // Specifies the cursor state. This parameter can be one of the following values:
                                    //    0             The cursor is hidden.
                                    //    CURSOR_SHOWING    The cursor is showing.
        public IntPtr hCursor;          // Handle to the cursor. 
        public POINT ptScreenPos;       // A POINT structure that receives the screen coordinates of the cursor. 
    }


    internal class NativeMethods
    {
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool GetCursorInfo(out CURSORINFO pci);
    }


    public sealed class CursorHelper
    {
        private CursorHelper() { }

        public static Cursor GetCursor()
        {
            CURSORINFO pci;
            pci.cbSize = Marshal.SizeOf(typeof(CURSORINFO));
            if (!NativeMethods.GetCursorInfo(out pci))
                throw new Win32Exception(Marshal.GetLastWin32Error());

            //Logger.LogMessage("Get Cursor info success, CursorsType: ");

            var h = pci.hCursor;
            if (h == Cursors.AppStarting.Handle) return Cursor.AppStarting;
            if (h == Cursors.Arrow.Handle) return Cursor.Arrow;
            if (h == Cursors.Cross.Handle) return Cursor.Cross;
            if (h == Cursors.Default.Handle) return Cursor.Default;
            if (h == Cursors.IBeam.Handle) return Cursor.IBeam;
            if (h == Cursors.No.Handle) return Cursor.No;
            if (h == Cursors.SizeAll.Handle) return Cursor.SizeAll;
            if (h == Cursors.SizeNESW.Handle) return Cursor.SizeNESW;
            if (h == Cursors.SizeNS.Handle) return Cursor.SizeNS;
            if (h == Cursors.SizeNWSE.Handle) return Cursor.SizeNWSE;
            if (h == Cursors.SizeWE.Handle) return Cursor.SizeWE;
            if (h == Cursors.UpArrow.Handle) return Cursor.UpArrow;
            if (h == Cursors.WaitCursor.Handle) return Cursor.WaitCursor;
            if (h == Cursors.Help.Handle) return Cursor.Help;
            if (h == Cursors.HSplit.Handle) return Cursor.HSplit;
            if (h == Cursors.VSplit.Handle) return Cursor.VSplit;
            if (h == Cursors.NoMove2D.Handle) return Cursor.NoMove2D;
            if (h == Cursors.NoMoveHoriz.Handle) return Cursor.NoMoveHoriz;
            if (h == Cursors.NoMoveVert.Handle) return Cursor.NoMoveVert;
            if (h == Cursors.PanEast.Handle) return Cursor.PanEast;
            if (h == Cursors.PanNE.Handle) return Cursor.PanNE;
            if (h == Cursors.PanNorth.Handle) return Cursor.PanNorth;
            if (h == Cursors.PanNW.Handle) return Cursor.PanNW;
            if (h == Cursors.PanSE.Handle) return Cursor.PanSE;
            if (h == Cursors.PanSouth.Handle) return Cursor.PanSouth;
            if (h == Cursors.PanSW.Handle) return Cursor.PanSW;
            if (h == Cursors.PanWest.Handle) return Cursor.PanWest;
            if (h == Cursors.Hand.Handle) return Cursor.Hand;
            return Cursor.UNKOWNCURSOR;
        }
    }
}
