using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

using System.Windows.Forms;

using IronPython.Hosting;
using Microsoft.Scripting.Hosting;

using Microsoft.Office.Tools.Ribbon;



namespace PyExcelAddIn
{

    public interface IPyRibbon : Office.IRibbonExtensibility
    {
        void Ribbon_Load(Office.IRibbonUI ribbonUI);
        void OnButtonClick(Office.IRibbonControl control);
    }



    public partial class ThisAddIn
    {

        public static ScriptRuntime PyRunTime = Python.CreateRuntime();
        public static dynamic BootMod=null;
        //public static string ScripPath = @"E:\VS2010\Projects\PyExcelAddIn\PyExcelAddIn\Scripts";
        public static string ScripPath = @"H:\Visual Studio 2010\Projects\PyExcelAddIn\PyExcelAddIn\Scripts";
        public static string LibsFile = ScripPath + @"\Libs.zip";
        public static string BootFile = ScripPath + @"\boot.py";

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {

            try
            {
                if (BootMod == null)
                {
                    BootMod = PyRunTime.ExecuteFile(BootFile);
                    BootMod.init("PyExcelAddIn", ScripPath, LibsFile);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                //dynamic gui = PyRunTime.ImportModule("xmlgui");
                return BootMod.getRibbon();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }


        }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                if (BootMod == null)
                {
                    BootMod = PyRunTime.ExecuteFile(BootFile);
                    BootMod.init("PyExcelAddIn", ScripPath, LibsFile);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
