using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;

namespace FA.DB
{
    public class clsDatabaseinfo
    {
    }
    public class clszichanfuzaibiaoinfo
    {
        public string Order_id { get; set; }
        public string xiangmu { get; set; }
        public string hangci { get; set; }
        public string qimojine { get; set; }
        public string shangniantongqishu { get; set; }

        public string nianchujine { get; set; }
        public string xiangmuF { get; set; }
        public string hangciG { get; set; }
        public string qimojineH { get; set; }
        public string shangniantongqishuI { get; set; }
        public string nianchujineJ { get; set; }
        public string Input_Date { get; set; }

        public string bianzhidanwei { get; set; }
        public string riqi { get; set; }
        public string danwei { get; set; }


    }
    public class clszhuyaojingyingzhibiaowanchengqingkuanginfo
    {
        public string Order_id { get; set; }
        public string xuhao1 { get; set; }//序号1
        public string zhibiaomingcheng { get; set; }//指标名称
        public string nianchuzhibiaozhihuoqichushu { get; set; }//年初指标值或期初数
        public string benyuewancheng { get; set; }//本月完成


        public string huanbizengjian { get; set; }//环比增减

        public string leijiwanchenghuoqimoshu { get; set; }//累计完成或期末数

        public string wanchengbili { get; set; }//完成比例


        public string shangniantongqileijiwancheng { get; set; }//上年同期累计完成

        public string tongbizengzhang { get; set; }//同比增减

    
        public string Input_Date { get; set; }

   
        public string danwei { get; set; }


    }

}
