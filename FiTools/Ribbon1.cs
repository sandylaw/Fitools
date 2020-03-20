using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace FiTools
{

    public partial class Ribbon1
    {
        public Excel.Application ExcelApp;
        public int jizhun;// 乘除基准
        public string lie;//定义列
        public string zhi;//定义单元格特征值        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            ExcelApp = Globals.ThisAddIn.Application;
        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            //默认一万
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            jizhun = Int32.Parse(editBox1.Text);
            foreach (Excel.Range rg in ExcelApp.Selection)
            { //乘以指定数，默认乘以一万
                if (ExcelApp.WorksheetFunction.IsNumber(rg))
                { rg.Value = rg.Value * jizhun; }
                else
                { rg.Value = 0; }


            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            jizhun = Int32.Parse(editBox1.Text);
            foreach (Excel.Range rg in ExcelApp.Selection)
            { //除以指定数，默认除以一万
                if (ExcelApp.WorksheetFunction.IsNumber(rg))
                { rg.Value = rg.Value / jizhun; }
                else
                { rg.Value = 0; }


            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            //带公式乘以数，默认一万
            jizhun = Int32.Parse(editBox1.Text);
            foreach (Excel.Range rg in ExcelApp.Selection)
            {
                if (ExcelApp.WorksheetFunction.IsNumber(rg.Value))
                {
                    if (rg.HasFormula) //判断是否含有公式
                    { 
                        rg.Formula = "=(" + rg.Formula.Substring(1) + ")" + "*" + jizhun; //substring用来去除原来公式的=号
                    }
                    else
                    {
                        rg.Formula = "=(" + rg.Formula + ")" + "*" + jizhun; 
                    }
                }
                

                else
                { }


            }
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            //带公式除以数，默认一万
            jizhun = Int32.Parse(editBox1.Text);
            foreach (Excel.Range rg in ExcelApp.Selection)
            {
                if (ExcelApp.WorksheetFunction.IsNumber(rg.Value))
                {
                    if (rg.HasFormula)//判断是否含有公式
                    { 
                        rg.Formula = "=(" + rg.Formula.Substring(1) + ")" + "/" + jizhun;
                    }
                    else
                    {
                        rg.Formula = "=(" + rg.Formula + ")" + "/" + jizhun;
                    }
                }
                

                else
                { }


            }
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            // 将不等于0的数值及其他字符转为NA
            foreach (Excel.Range rg in ExcelApp.Selection)
            {
                if (ExcelApp.WorksheetFunction.IsNumber(rg) && rg.Value != 0)
                {
                    continue;
                }
                else
                { rg.Formula = "=Na()"; }


            }

        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            //把#NA变为0，方便数据处理
            foreach (Excel.Range rg in ExcelApp.Selection)
            {
                if (ExcelApp.WorksheetFunction.IsNumber(rg))
                { continue; }
                else
                {
                    rg.Value = 0;
                }


            }

        }

        private void editBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {
            //默认B列
        }

        private void editBox3_TextChanged(object sender, RibbonControlEventArgs e)
        {
            //默认空值
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            //删除指定列指定单元格数值的所在行
            lie = editBox2.Text.ToString();//列取自编辑框2
            zhi = editBox3.Text.ToString();//值取自编辑框3

            ExcelApp.ScreenUpdating = false;
            ExcelApp.Cells.Select();
            ExcelApp.Selection.Copy();
            ExcelApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteValues);
            int lastrow = ExcelApp.ActiveSheet.UsedRange.Rows.count + ExcelApp.ActiveSheet.UsedRange.Rows(1).Row - 1;
            int x;
            for (x = lastrow; x >= 1; x--)
            {
                string y = x.ToString();
                object z;
                if (ExcelApp.Cells[y, lie].Value == null)//判断指定单元格的值是否为空
                {
                    if (zhi == "")//判断指定值是否为空
                    {
                        ExcelApp.Rows[y].Delete();//删除行
                    }
                    else
                    {

                    }
                }

                else
                {
                    z = ExcelApp.Cells[y, lie].Value;   //如果指定单元格值不为空，赋值为object              
                    string zz = z.ToString();//转为字符型
                    if (zz == zhi)//如果与指定值相同
                    { ExcelApp.Rows[y].Delete(); }//删除行
                }
            }
            ExcelApp.ScreenUpdating = true;

        }

        private void zhengfu_Click(object sender, RibbonControlEventArgs e)
        {//改变所选数值的正负号
            foreach (Excel.Range rg in ExcelApp.Selection)
            {
                if (ExcelApp.WorksheetFunction.IsNumber(rg))
                { rg.Value = -rg.Value; }
                else
                {  }


            }
        }

        private void xiaoshuwei_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void R_num_Click(object sender, RibbonControlEventArgs e)
        {//按照指定的小数位Round一下数值
            int xsw = Int32.Parse(xiaoshuwei.Text);//获取小数位
            foreach (Excel.Range rg in ExcelApp.Selection)
            {
                if (ExcelApp.WorksheetFunction.IsNumber(rg))
                { rg.Value = ExcelApp.WorksheetFunction.Round(rg.Value, xsw); }
                else
                { }
            }

        }

        private void R_Set_Click(object sender, RibbonControlEventArgs e)
        {//按照指定的小数位加Round公式
            int xsw = Int32.Parse(xiaoshuwei.Text);//获取小数位
            foreach (Excel.Range rg in ExcelApp.Selection)
            {
                if (ExcelApp.WorksheetFunction.IsNumber(rg))
                {
                    if (rg.HasFormula)//判断是否含有公式
                    { rg.Formula = "=round(" + rg.Formula.Substring(1) + ","+xsw+")" ; }
                    else
                    {
                        rg.Formula = "=round(" + rg.Formula + "," + xsw + ")"; 
                    }
                }
                

                else
                { }
            }
        }

        private void cal_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.ProcessStartInfo Info = new System.Diagnostics.ProcessStartInfo();
            Info.FileName = "calc.exe ";//"calc.exe"为计算器，"notepad.exe"为记事本
            System.Diagnostics.Process Proc = System.Diagnostics.Process.Start(Info);

        }

        private void ntbook_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.ProcessStartInfo Info = new System.Diagnostics.ProcessStartInfo();
            Info.FileName = "notepad.exe ";//"calc.exe"为计算器，"notepad.exe"为记事本
            System.Diagnostics.Process Proc = System.Diagnostics.Process.Start(Info);

        }

        private void about0_Click(object sender, RibbonControlEventArgs e)
        {
            AboutBox1 box = new AboutBox1();
            box.ShowDialog();
        }

        private void Rup_num_Click(object sender, RibbonControlEventArgs e)
        {    //向上round
            int xsw = Int32.Parse(xiaoshuwei.Text);//获取小数位
            foreach (Excel.Range rg in ExcelApp.Selection)
            {
                if (ExcelApp.WorksheetFunction.IsNumber(rg))
                { rg.Value = ExcelApp.WorksheetFunction.RoundUp(rg.Value, xsw); }
                else
                { }
            }
        }

        private void Rup_set_Click(object sender, RibbonControlEventArgs e)
        {
            //按照指定的小数位加Roundup公式
            int xsw = Int32.Parse(xiaoshuwei.Text);//获取小数位
            foreach (Excel.Range rg in ExcelApp.Selection)
            {
                if (ExcelApp.WorksheetFunction.IsNumber(rg))
                {
                    if (rg.HasFormula)//判断是否含有公式
                    { rg.Formula = "=roundup(" + rg.Formula.Substring(1) + "," + xsw + ")"; }
                    else
                    {
                        rg.Formula = "=roundup(" + rg.Formula + "," + xsw + ")";
                    }
                }


                else
                { }
            }
        }

        private void Rdown_num_Click(object sender, RibbonControlEventArgs e)
        {
            //向下round
            int xsw = Int32.Parse(xiaoshuwei.Text);//获取小数位
            foreach (Excel.Range rg in ExcelApp.Selection)
            {
                if (ExcelApp.WorksheetFunction.IsNumber(rg))
                { rg.Value = ExcelApp.WorksheetFunction.RoundDown(rg.Value, xsw); }
                else
                { }
            }
        }

        private void Rdown_set_Click(object sender, RibbonControlEventArgs e)
        {//按照指定的小数位加Rounddown公式
            int xsw = Int32.Parse(xiaoshuwei.Text);//获取小数位
            foreach (Excel.Range rg in ExcelApp.Selection)
            {
                if (ExcelApp.WorksheetFunction.IsNumber(rg))
                {
                    if (rg.HasFormula)//判断是否含有公式
                    { rg.Formula = "=rounddown(" + rg.Formula.Substring(1) + "," + xsw + ")"; }
                    else
                    {
                        rg.Formula = "=rounddown(" + rg.Formula + "," + xsw + ")";
                    }
                }


                else
                { }
            }
        }

        private void toggleButton_fapiao_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ctp.Visible = ((RibbonToggleButton)sender).Checked;
        }

       

        

       
    }
}