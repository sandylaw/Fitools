using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FiTools
{
    class ShuiECal
    {//自定义类，传入函数金额和税率，返回一个数组，依次为函数金额、不含税金额、税率、税额。

        double buhanshui;
        double shuie;

        public double[] ShuiCal(double hanshui, double sl)//税额计算函数
        {

            double[] myarray = new double[4];

            buhanshui = Math.Round(hanshui / (1 + sl), 2, MidpointRounding.AwayFromZero);
            shuie = hanshui - buhanshui;
            myarray[0] = hanshui;
            myarray[1] = buhanshui;
            myarray[2] = sl;
            myarray[3] = shuie;
            return myarray;
        }
    }
}
