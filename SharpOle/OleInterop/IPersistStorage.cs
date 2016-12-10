using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [Guid("0000010A-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IPersistStorage //we don't need the base interface IPersist
    {
        [PreserveSig]
        void GetClassID(
            out Guid pClassID);
        [PreserveSig]
        int IsDirty();
        void InitNew(
            [MarshalAs(UnmanagedType.Interface)] [In] IStorage pstg);
        void Load(
            [MarshalAs(UnmanagedType.Interface)] [In] IStorage pstg);
        void Save(
            [MarshalAs(UnmanagedType.Interface)] [In] IStorage pStgSave,
            [In] int fSameAsLoad);
        void SaveCompleted(
            [MarshalAs(UnmanagedType.Interface)] [In] IStorage pStgNew);
        void HandsOffStorage();
    }
}
