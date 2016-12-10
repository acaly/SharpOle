using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [Guid("0000011B-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IOleContainer // : IParseDisplayName
    {
        void ParseDisplayName(
            [MarshalAs(UnmanagedType.Interface)] [In] IBindCtx pbc,
            [MarshalAs(UnmanagedType.LPWStr)] [In] string pszDisplayName,
            out uint pchEaten,
            [MarshalAs(UnmanagedType.Interface)] out IMoniker ppmkOut);
        void EnumObjects(
            [In] uint grfFlags,
            [MarshalAs(UnmanagedType.Interface)] out IEnumUnknown ppEnum);
        void LockContainer(
            [In] int fLock);
    }
}
