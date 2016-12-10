using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [StructLayout(LayoutKind.Sequential, Pack = 2)]
    public struct LogPalette
    {
        public ushort palVersion;
        public ushort palNumEntries;
        public IntPtr palPalEntry;
    }

}
