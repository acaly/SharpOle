using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct FormatEtc
    {
        public ushort cfFormat;
        public IntPtr ptd;
        public uint dwAspect;
        public int lindex;
        public uint tymed;
    }

}
