using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [Guid("0000000F-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IMoniker // we don't need IPersistStream
    {
        [PreserveSig]
        int GetClassID(
            out Guid pClassID);
        [PreserveSig]
        int IsDirty();
        void Load(
            [MarshalAs(28)] [In] IStream pstm);
        void Save(
            [MarshalAs(28)] [In] IStream pstm,
            [In] int fClearDirty);
        void GetSizeMax(
            [MarshalAs(42)] [Out] ulong[] pcbSize);
        void BindToObject(
            [MarshalAs(28)] [In] IBindCtx pbc,
            [MarshalAs(28)] [In] IMoniker pmkToLeft,
            [In] ref Guid riidResult,
            [MarshalAs(25)] out object ppvResult);
        void BindToStorage(
            [MarshalAs(28)] [In] IBindCtx pbc,
            [MarshalAs(28)] [In] IMoniker pmkToLeft,
            [In] ref Guid riid,
            [MarshalAs(25)] out object ppvObj);
        void Reduce(
            [MarshalAs(28)] [In] IBindCtx pbc,
            [In] uint dwReduceHowFar,
            [MarshalAs(28)] [In] [Out] ref IMoniker ppmkToLeft,
            [MarshalAs(28)] out IMoniker ppmkReduced);
        void ComposeWith(
            [MarshalAs(28)] [In] IMoniker pmkRight,
            [In] int fOnlyIfNotGeneric,
            [MarshalAs(28)] out IMoniker ppmkComposite);
        void Enum(
            [In] int fForward,
            [MarshalAs(28)] out IEnumMoniker ppenumMoniker);
        void IsEqual(
            [MarshalAs(28)] [In] IMoniker pmkOtherMoniker);
        void Hash(
            out uint pdwHash);
        [PreserveSig]
        int IsRunning(
            [MarshalAs(28)] [In] IBindCtx pbc,
            [MarshalAs(28)] [In] IMoniker pmkToLeft,
            [MarshalAs(28)] [In] IMoniker pmkNewlyRunning);
        void GetTimeOfLastChange(
            [MarshalAs(28)] [In] IBindCtx pbc,
            [MarshalAs(28)] [In] IMoniker pmkToLeft,
            [MarshalAs(42)] [Out] FileTime[] pFileTime);
        void Inverse(
            [MarshalAs(28)] out IMoniker ppmk);
        void CommonPrefixWith(
            [MarshalAs(28)] [In] IMoniker pmkOther,
            [MarshalAs(28)] out IMoniker ppmkPrefix);
        void RelativePathTo(
            [MarshalAs(28)] [In] IMoniker pmkOther,
            [MarshalAs(28)] out IMoniker ppmkRelPath);
        void GetDisplayName(
            [MarshalAs(28)] [In] IBindCtx pbc,
            [MarshalAs(28)] [In] IMoniker pmkToLeft,
            [MarshalAs(21)] out string ppszDisplayName);
        void ParseDisplayName(
            [MarshalAs(28)] [In] IBindCtx pbc,
            [MarshalAs(28)] [In] IMoniker pmkToLeft,
            [MarshalAs(21)] [In] string pszDisplayName,
            out uint pchEaten,
            [MarshalAs(28)] out IMoniker ppmkOut);
        void IsSystemMoniker(
            out uint pdwMksys);
    }
}
