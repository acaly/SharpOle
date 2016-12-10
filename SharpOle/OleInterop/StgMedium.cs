using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct StgMedium
    {
        public uint tymed;
        public IntPtr unionmember;
        [MarshalAs(UnmanagedType.IUnknown)]
        public object pUnkForRelease;
    }

}
