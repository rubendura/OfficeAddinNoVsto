using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Extensibility;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using System.IO;

namespace Office_Addin
{

    public interface IMyAddIn
    {
        void ShowMessageBase();
        Office.IRibbonUI Ribbon { get; }
        void ActivateTab(string id);
        void ActivateTabMso(string id);
        void ActivateTabQ(string id, string ns);
    }

    public class AddIn : IMyAddIn
    {
        public void ShowMessageBase()
        {
            MessageBox.Show("Hello World!");
        }

        public Office.IRibbonUI Ribbon { get; set; }

        public void ActivateTab(string id)
        {
            Ribbon.ActivateTab(id);
        }

        public void ActivateTabQ(string id, string ns)
        {
            Ribbon.ActivateTabQ(id, ns);
        }

        public void ActivateTabMso(string id)
        {
            Ribbon.ActivateTabMso(id);
        }
    }

    [Guid("9C5533D9-3372-4F9C-81F2-E18F9AEEF2B8")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("TestAddin.Connect")]
    public class Connect : IDTExtensibility2, IMyAddIn, Office.IRibbonExtensibility
    {

        private AddIn _addIn;
        private AddIn AddIn { get { return _addIn; } }
        private Office.IRibbonUI _ribbon;
        public Office.IRibbonUI Ribbon { get { return _ribbon; } }
        
        public void OnAddInsUpdate(ref Array custom)
        {
            //throw new NotImplementedException();
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _addIn = new AddIn();
            ((Office.COMAddIn)AddInInst).Object = AddIn;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            //throw new NotImplementedException();
        }

        public void OnStartupComplete(ref Array custom)
        {
            //MessageBox.Show("Hello World!");
            //throw new NotImplementedException();
        }


        public void OnBeginShutdown(ref Array custom)
        {
            //throw new NotImplementedException();
        }

        public void ShowMessage(Office.IRibbonControl control)
        {
            //ShowMessageBase();
            //Ribbon.ActivateTabMso("TabHome");
            Ribbon.ActivateTabQ("MyTab","testnamespace");
        }

        public void ActivateTab(string id)
        {
            AddIn.ActivateTab(id);
        }

        public void ActivateTabQ(string id, string ns)
        {
            AddIn.ActivateTabQ(id, ns);
        }

        public void ActivateTabMso(string id)
        {
            AddIn.ActivateTabMso(id);
        }

        //public void ShowMessage(Office.IRibbonControl control, ref bool CancelDefault)
        //{
        //    ShowMessageBase();
        //}

        public void ShowMessageBase()
        {
            MessageBox.Show("Hello World!");
        }

        public string GetCustomUI(string RibbonID)
        {
            return File.ReadAllText(@"C:\Users\ruben\Documents\Visual Studio 2013\Projects\Office Addin\Office Addin\bin\Debug\Ribbon.xml");
        }

        public void OnLoad(Office.IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            AddIn.Ribbon = Ribbon;
        }
    }
}
