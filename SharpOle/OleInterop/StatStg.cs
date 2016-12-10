using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [StructLayout(LayoutKind.Sequential, Pack = 8)]
    public struct StatStg
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string pwcsName;
        public uint type;
        public ulong cbSize;
        public FileTime mtime;
        public FileTime ctime;
        public FileTime atime;
        public uint grfMode;
        public uint grfLocksSupported;
        public Guid clsid;
        public uint grfStateBits;
        public uint reserved;
    }

}
