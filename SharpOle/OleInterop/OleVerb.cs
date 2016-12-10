using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct OleVerb
    {
        public int lVerb;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string lpszVerbName;
        public uint fuFlags;
        public uint grfAttribs;
    }

}
