using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using Extensibility;
using System.Runtime.InteropServices; //为添加 GUID 和 ProgID
using System.Reflection;
using Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Microsoft.Office.Interop.PowerPoint;

namespace UpdateLinks
{
    [Guid("38A6AEE0-5131-42B0-B74F-8FD8CEDE3740")]
    [ProgId("UpdateLinks.Connect")]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private Microsoft.Vbe.Interop.VBE instance;
        private Excel.Application excelApp;
        private PowerPoint.Application powerpointApp;
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            System.Windows.Forms.MessageBox.Show("C#.NET ComAddin 你好 ");
            COMAddIn addin = AddInInst as COMAddIn;
            MessageBox.Show("My add-in ProgID is " + addin.ProgId);
            instance = Application as VBE;
            //Excel.Application xlapp =  (Microsoft.Office.Interop.Excel.Application) System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            //Excel.Application instance = null;
            try
            {
                //instance = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                Excel.Application instance = (Excel.Application)Application;
                MessageBox.Show(string.Format("The application this loaded into is called {0}.", instance.Name));
            }
            //catch (System.Runtime.InteropServices.COMException ex)
            catch (Exception ex)
            {
                //  instance = new Excel.ApplicationClass();
                MessageBox.Show(string.Format("Instance Not Excel.Application!--{0}", ex.Message));
            }

            //PowerPoint.Application app = Application as PowerPoint.Application;            
            //MessageBox.Show(string.Format("The application this loaded into is called {0}.", app.Name));
            //MessageBox.Show(string.Format("The application this loaded into is called {0} --- {1}.",Application.GetType().Name, app.GetType().Name));
            MessageBox.Show(string.Format("Load mode was {0}.", ConnectMode.ToString()));
            //https://www.office-forums.com/threads/how-to-get-the-application-object-type.2164001/
            //https://www.ka-net.org/blog/?p=5552#google_vignette
            try
            {
                if (Microsoft.VisualBasic.Information.TypeName(Application) == "Application")
                {
                    String appName = Application.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, Application, null) as String;

                    if (appName == "Microsoft PowerPoint")
                    {
                        powerpointApp = Application as PowerPoint.Application;
                        if (powerpointApp != null)
                        {
                            MessageBox.Show("application is PowerPoint app!");
                        }
                        else
                        {
                            MessageBox.Show("powerpointApp is null");
                        }
                    }
                    else if (appName == "Microsoft Excel")
                    {
                        excelApp = Application as Excel.Application;
                        if (excelApp != null)
                        {
                            MessageBox.Show("application is Excel app");
                        }
                        else
                        {
                            MessageBox.Show("excelApp is null");
                        }
                    }
                    else
                    {
                        MessageBox.Show("application is neither PowerPoint nor Excel");
                    }
                }
                else
                {
                    MessageBox.Show("application is not an Office app");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            MessageBox.Show("OnDisconnection");
            MessageBox.Show(string.Format("Disconnect mode was {0}.", RemoveMode.ToString()));
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            MessageBox.Show("OnAddinsUpdate");
        }

        public void OnStartupComplete(ref Array custom)
        {
            MessageBox.Show("OnStarUpComplete");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            MessageBox.Show("OnBeginShutDown");
        }

        public string GetCustomUI(string RibbonID)
        {
            return updateLinks.Properties.Resources.ribbon;
        }
        public void updateLink()
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.InitialDirectory = powerpointApp.ActivePresentation.Path;
            openfile.Title = "请选择需要导入的Excel文件";
            openfile.Multiselect = false;
            openfile.Filter = "Excel文件(*.xlsx;*.xls)|*.xlsx;*.xls";
            openfile.RestoreDirectory = true;
            if (openfile.ShowDialog() == DialogResult.OK)
            {
                string path = openfile.FileName;
                foreach (Slide slide in powerpointApp.ActivePresentation.Slides)
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.HasChart == MsoTriState.msoTrue && shape.Chart.ChartData.IsLinked && (shape.LinkFormat.SourceFullName != null))
                        //if (shape.LinkFormat.SourceFullName != null)
                        {
                            shape.LinkFormat.SourceFullName = path;
                            shape.LinkFormat.Update();
                            shape.Chart.Refresh();
                        }

                    }
                }
            }

        }
        public void showHello(IRibbonControl control)
        {
            if (powerpointApp != null)
            {
                MessageBox.Show("我是PowerPoint!");
                updateLink();
            }
            else
            {
                MessageBox.Show("你点击了我哈哈!");
            }
        }
    }
}
