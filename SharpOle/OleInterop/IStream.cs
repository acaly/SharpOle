using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [Guid("0000000C-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IStream //we don't need ISequentialStream
    {
        void Read(
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] [Out] byte[] pv,
            [In] uint cb,
            out uint pcbRead);
        void Write(
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] [In] byte[] pv,
            [In] uint cb,
            out uint pcbWritten);
        void Seek(
            [In] long dlibMove,
            [In] uint dwOrigin,
            [MarshalAs(UnmanagedType.LPArray)] [Out] ulong[] plibNewPosition);
        void SetSize(
            [In] ulong libNewSize);
        void CopyTo(
            [MarshalAs(UnmanagedType.Interface)] [In] IStream pstm,
            [In] ulong cb,
            [MarshalAs(UnmanagedType.LPArray)] [Out] ulong[] pcbRead,
            [MarshalAs(UnmanagedType.LPArray)] [Out] ulong[] pcbWritten);
        void Commit(
            [In] uint grfCommitFlags);
        void Revert();
        void LockRegion(
            [In] ulong libOffset,
            [In] ulong cb,
            [In] uint dwLockType);
        void UnlockRegion(
            [In] ulong libOffset,
            [In] ulong cb,
            [In] uint dwLockType);
        void Stat(
            [MarshalAs(UnmanagedType.LPArray)] [Out] StatStg[] pstatstg,
            [In] uint grfStatFlag);
        void Clone(
            [MarshalAs(UnmanagedType.Interface)] out IStream ppstm);
    }
}
