using Microsoft.VisualStudio.OLE.Interop;
using SharpOle;
using SharpOle.OleServer;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SharpOLETest
{
    [ComVisible(true)]
    [Guid("e8b6e3e0-05dc-4f19-a349-5ac070ae8402")]
    [ProgId("SharpOleTest.COLEObjectTest")]
    [ClassInterface(ClassInterfaceType.None)]
    public class COLEObjectTest : AbstractOleObject
    {
        private static Font _Font = new Font(SystemFonts.DefaultFont.FontFamily, 50.0f);

        public override void Draw(Graphics g)
        {
            g.DrawRectangle(System.Drawing.Pens.Black, 0, 0, 300, 300);
            g.DrawString(DateTime.Now.ToShortTimeString(), _Font, Brushes.Green, new RectangleF(0, 0, 300, 300));
        }
    }

    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            OleUtilities.OleInit();

            var ole = new COLEObjectTest();
            ole.CopyToClipboard();

            Application.Run(new Form1());

            OleUtilities.OleUninit();
        }
    }
}
