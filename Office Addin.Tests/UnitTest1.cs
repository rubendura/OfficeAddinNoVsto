using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Office_Addin;
using System.Windows.Forms;

namespace Office_Addin.Tests
{
    [TestClass]
    public class UnitTest1
    {
        public UnitTest1()
        {
            try
            {
                _wordApp = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
            }
            catch
            {
                System.Diagnostics.Process.Start(@"C:\Program Files\Microsoft Office 15\root\office15\WINWORD.EXE");

                while (true)
                {
                    try
                    {
                        _wordApp = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                        break;
                    }
                    catch { }
                }
            }

            _comAddin = WordApp.COMAddIns.Item("TestAddin.Connect");
            _myAddin = (IMyAddIn) ComAddin.Object;
        }

        private readonly Word.Application _wordApp;
        private Word.Application WordApp { get { return _wordApp; } }
        private readonly Office.COMAddIn _comAddin;
        private Office.COMAddIn ComAddin { get { return _comAddin; } }
        private readonly IMyAddIn _myAddin;
        private IMyAddIn MyAddin { get { return _myAddin; } }

        [TestMethod]
        public void ShowMessage()
        {
            MyAddin.ShowMessageBase();
        }

        [TestMethod]
        public void ChangeTab()
        {
            //MyAddin.Ribbon.ActivateTabQ("MyTab", "testnamespace");
            MyAddin.ActivateTabQ("MyTab", "testnamespace");
        }

        [TestMethod]
        public void ChangeTabMso()
        {
            //MyAddin.Ribbon.ActivateTabMso("TabReviewWord");
            MyAddin.Ribbon.ActivateTabMso("TabHome");
        }

        [TestMethod]
        public void InvalidateRibbon()
        {
            MyAddin.Ribbon.Invalidate();
        }

    }

}
