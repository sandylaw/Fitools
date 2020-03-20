using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace FiTools
{
    public partial class UserControl_FaPiao : UserControl
    {
        Excel.Application ExcelApp;
        int JinEFuDu;//定义金额调节幅度数值类型；
        public UserControl_FaPiao()
        {
            InitializeComponent();
        }
        public bool IsNumberic(string oText)
        {
            try
            {
                int var1 = Convert.ToInt32(oText);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void UserControl_FaPiao_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 2;//默认“百”
            textBox4.Text = "1.输入开票金额、税率、单张发票限额；"
                + Environment.NewLine + "2.鼠标定位到某单元格,并单击；"
                + Environment.NewLine + "3.单击拆票按钮。";


        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void button1_Click(object sender, EventArgs e)
        {

            ExcelApp = Globals.ThisAddIn.Application;

            Excel.Range rg;
            rg = (Excel.Range)this.ExcelApp.ActiveCell;
            object cellvalue = rg.Value;

            //定义变量
            while (true)
            {
                if (IsNumberic(textBox1.Text) && IsNumberic(textBox2.Text) && IsNumberic(textBox3.Text))
                {
                    double KaiPiaoJinE = float.Parse(textBox1.Text);//开票总金额，含税。
                    double ShuiLv = Math.Round(double.Parse(textBox2.Text) / 100, 2, MidpointRounding.AwayFromZero);//税率
                    int FaPiaoXianE = int.Parse(textBox3.Text);//发票限额
                    string JinEBuZi = comboBox1.Text;//获取金额调节幅度选择项

                    if (Math.Round(double.Parse(textBox2.Text) / 100, 2, MidpointRounding.AwayFromZero) > 0 && Math.Round(double.Parse(textBox2.Text) / 100, 2, MidpointRounding.AwayFromZero) < 0.25 && float.Parse(textBox1.Text) > 0 && int.Parse(textBox3.Text) > 0)
                    {


                        //写入标题
                        rg[1, 2] = "需开票情况";
                        rg[2, 2] = "总含税金额";
                        rg[2, 3] = "不含税金额";
                        rg[2, 4] = "税率%";
                        rg[2, 5] = "税额";

                        ShuiECal ZCal = new ShuiECal();//实例化发票总额的税额计算类；
                        double[] Zarray = ZCal.ShuiCal(KaiPiaoJinE, ShuiLv);//计算结果放入数组Zarray
                        rg[3, 2] = Zarray[0];
                        rg[3, 3] = Zarray[1];
                        rg[3, 4] = Zarray[2] * 100;
                        rg[3, 5] = Zarray[3];
                        //写入发票明细表标题
                        rg[4, 2] = "拆分发票明细";
                        rg[4, 4] = "张发票";
                        rg[5, 2] = "含税金额";
                        rg[5, 3] = "不含税金额";
                        rg[5, 4] = "税率%";
                        rg[5, 5] = "税额";

                        //根据金额调节步子不同，转行为数字
                        if (JinEBuZi == "个")
                        { JinEFuDu = 1; }
                        else if (JinEBuZi == "十")
                        { JinEFuDu = 10; }
                        else if (JinEBuZi == "百")
                        { JinEFuDu = 100; }
                        else if (JinEBuZi == "千")
                        { JinEFuDu = 1000; }
                        else if (JinEBuZi == "万")
                        { JinEFuDu = 10000; }
                        else if (JinEBuZi == "十万")
                        { JinEFuDu = 100000; }
                        else if (JinEBuZi == "百万")
                        { JinEFuDu = 1000000; }
                        else if (JinEBuZi == "千万")
                        { JinEFuDu = 10000000; }

                        int zhangshu;//定义开票的相同金额的张数
                        ShuiECal XCal = new ShuiECal();//实例化单张发票的税额计算类；
                        double[] Xarray;//定义相同金额的发票的数组，存储含税金额、不含税金额、税率、税额
                        double[] Yarray;//定义最后一张发票的数组，存储含税金额、不含税金额、税率、税额

                        for (int x = 0; x <= 1000; x++)//x为循环尝试次数
                        {


                            zhangshu = (int)(Math.Floor(KaiPiaoJinE / (FaPiaoXianE - JinEFuDu * x)));//计算重复发票张数，向下取整

                            Xarray = XCal.ShuiCal(FaPiaoXianE - JinEFuDu * x, ShuiLv);//存储相同金额的发票的含税金额、不含税金额、税率、税额

                            //double singlejine=Math.Round(Xarray[3],ShuiEJingDu,MidpointRounding.AwayFromZero)

                            Yarray = XCal.ShuiCal(KaiPiaoJinE - zhangshu * Xarray[0], ShuiLv);//存储最后一张发票的含税金额、不含税金额、税率、税额


                            if (Yarray[0] > Xarray[0])//最后一张发票金额大于相同发票的金额，需要调节金额步子
                            { MessageBox.Show("溢出，请调整金额调节步子！"); }
                            else
                            {
                                if (Math.Round(Xarray[3] * zhangshu + Yarray[3] - Zarray[3], 2, MidpointRounding.AwayFromZero) == 0)//判断税额是否凑齐
                                {
                                    //循环写入发票数据
                                    for (int row = 6; row < 6 + zhangshu; row++)
                                    {
                                        rg[row, 1] = row - 5;
                                        rg[row, 2] = Xarray[0];
                                        rg[row, 3] = Xarray[1];
                                        rg[row, 4] = Xarray[2] * 100;
                                        rg[row, 5] = Xarray[3];
                                    }
                                    rg[zhangshu + 6, 1] = zhangshu + 1;
                                    rg[zhangshu + 6, 2] = Yarray[0];
                                    rg[zhangshu + 6, 3] = Yarray[1];
                                    rg[zhangshu + 6, 4] = Yarray[2] * 100;
                                    rg[zhangshu + 6, 5] = Yarray[3];
                                    rg[4, 3] = zhangshu + 1;
                                    string FapiaoSum = Convert.ToString(zhangshu + 1);

                                    MessageBox.Show("拆票成功,一共需开票" + FapiaoSum + "张");

                                    return;
                                }
                                else
                                { continue; }

                            }



                        }
                    }
                    else
                    {
                        MessageBox.Show("请检查开票参数");
                        //Application.Exit();
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("请检查开票参数");
                    //  this.Dispose(); 
                    // Application.Exit();
                    return;
                }
            }


        }


    }
}
