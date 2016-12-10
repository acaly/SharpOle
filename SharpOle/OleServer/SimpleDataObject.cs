using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleServer
{
    [ComVisible(true)]
    [Guid("E4380DA4-757A-473C-9829-C0E0CB7138C5")]
    [ProgId("SharpOleLib.SimpleDataObject")]
    [ClassInterface(ClassInterfaceType.None)]
    class SimpleDataObject : AbstractDataObject
    {
        public SimpleDataObject(AbstractOleObject oleObj)
        {
            this.OleObject = oleObj;
        }

        public override Guid GetGuid()
        {
            return OleObject.GetGuid();
        }
    }
}
