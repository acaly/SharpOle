using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle
{
    public static class OleUtilities
    {
        [DllImport("ole32.dll")]
        private static extern int OleInitialize(IntPtr pvReserved);

        [DllImport("ole32.dll")]
        private static extern void OleUninitialize();

        public static void OleInit()
        {
            OleInitialize(IntPtr.Zero);
        }

        public static void OleUninit()
        {
            OleUninitialize();
        }
    }
}
