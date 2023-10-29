using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ClassLibrary1
{
    [Guid("4CF8DC42-0A3C-4D48-A9EC-5D234BE8B640"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    //internal interface lsqypj_Interface
    public interface lsqypj_Interface
    {
        [DispId(1)]
        void ShowMessage();
        [DispId(2)]
        int add(int a, int b);
    }
    [Guid("ECC86196-93CF-4328-B49E-C39F34B3756F"), ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(typeof(lsqypj_Interface))]
    public class lsqypj_function : lsqypj_Interface 
    {
        public void ShowMessage()
        {
            MessageBox.Show("ahahha");
        }

        public int add(int a, int b)
        {
            return a + b;
        }
    }
}
