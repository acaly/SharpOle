using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [Guid("00000118-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IOleClientSite
    {
        void SaveObject();
        void GetMoniker(
            [In] uint dwAssign,
            [In] uint dwWhichMoniker,
            [MarshalAs(UnmanagedType.Interface)] out IMoniker ppmk);
        void GetContainer(
            [MarshalAs(UnmanagedType.Interface)] out IOleContainer ppContainer);
        void ShowObject();
        void OnShowWindow(
            [In] int fShow);
        void RequestNewObjectLayout();
    }
}
