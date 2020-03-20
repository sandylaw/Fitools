using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace FiTools
{
    public partial class ThisAddIn
    {
        UserControl_FaPiao uc_fapiao;
        Microsoft.Office.Tools.CustomTaskPane ctp1;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            uc_fapiao=new UserControl_FaPiao();
            ctp1 = Globals.ThisAddIn.CustomTaskPanes.Add(uc_fapiao, "发票拆票");
            
            ctp1.VisibleChanged +=
                new EventHandler(ctp1_VisibleChanged);

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        private void ctp1_VisibleChanged(object sender, System.EventArgs e)
        {

            Globals.Ribbons.Ribbon1.toggleButton_fapiao.Checked =
                ctp1.Visible;

        }
        public Microsoft.Office.Tools.CustomTaskPane ctp
        {
            get
            {
                return ctp1;
            }
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
