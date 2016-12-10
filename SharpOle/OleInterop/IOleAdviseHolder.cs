using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [Guid("00000111-0000-0000-C000-000000000046")]
    [InterfaceType(1)]
    public interface IOleAdviseHolder
    {
        void Advise(IAdviseSink pAdvise, out uint pdwConnection);
        void EnumAdvise(out IEnumStatData ppenumAdvise);
        void SendOnClose();
        void SendOnRename(IMoniker pmk);
        void SendOnSave(IMoniker pmk);
        void Unadvise(uint dwConnection);
    }
}
