using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace bore_hole1
{
    public class Class1
    {
        [CommandMethod("BlockModel")]
        public void BHL()
        {
            var BHL = new ClassLibrary2.Form1();
            BHL.Show();
        }
    }
}