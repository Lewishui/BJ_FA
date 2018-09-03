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
    public class clsLirunjilirunfenpeibiao_info
    {
        public string _id { get; set; }
        public string xiangmu { get; set; }
        public string hangci { get; set; }
        public string benyueshu  { get; set; }
        public string bennianleijishu  { get; set; }
        public string shangniantongqishu { get; set; }
        public string yingyezongshouru { get; set; }//营业总收入    万元


        public string bensan { get; set; }//本月三项费用总和
        public string leijisanxiang { get; set; }//累计三项费用总和
        public string shangniantongqi { get; set; }//上年同期三项费用总和


        public string Input_Date { get; set; }
        public string bianzhidanwei { get; set; }
        public string riqi { get; set; }
        public string danwei { get; set; }

        public string rowindex { get; set; }

    }

    public class clsXianjinliu_info
    {
        public string _id { get; set; }
        public string xiangmu { get; set; }
        public string hangci { get; set; }
        public string bennianjine { get; set; }
        public string shangnianjine { get; set; }

        public string Input_Date { get; set; }
        public string bianzhidanwei { get; set; }
        public string riqi { get; set; }
        public string danwei { get; set; }

        public string rowindex { get; set; }

        

    }
    public class cls8xiangfeiyongzhichu_info
    {
 
        public string Order_id { get; set; }
        public string xiangmu { get; set; }//项目 
        public string hangci { get; set; }//行次
        public string shangnianquannianfasheng { get; set; }// 上年全年发生 
        public string nianduyusuan { get; set; }//年度预算
        public string heji_benyueshu { get; set; }//合计 本月数
        public string heji_bennianleiji { get; set; }// 合计 本年累计数
        public string heji_shangniantongqishu { get; set; }// 合计 上年同期数
        public string zaijian_benyueshu { get; set; }//在建工程（不含前期）本月数
        public string zaijian_bennianleijishu { get; set; }//在建工程（不含前期）本年累计数
        public string zaijian_shangniantongqishu { get; set; }//在建工程（不含前期） 上年同期数
        public string xiangmuqian_benyueshu { get; set; }//项目前期费 本月数
        public string xiangmuqian_bennianleijishu { get; set; }//项目前期费 本年累计数
        public string xiangmuqian_shangniantongqishu { get; set; }//项目前期费   上年同期数
        public string gongchengshigong_benyueshu { get; set; }//工程施工 本月数
        public string gongchengshigong_bennianleijishu { get; set; }//工程施工 本年累计数
        public string gongchengshigong_shangniantongqishu { get; set; }//工程施工  上年同期数
        public string shengchancheng_benyueshu { get; set; }//生产成本 本月数
        public string shengchancheng_bennianleijishu { get; set; }// 生产成本 本年累计数
        public string shengchancheng_shangniantongqishu { get; set; }//生产成本 上年同期数 
        public string guanlifei_benyueshu { get; set; }//管理费用 本月数
        public string guanlifei_bennianleijishu { get; set; }//管理费用 本年累计数
        public string guanlifei_shangniantongqishu { get; set; }//管理费用 上年同期数 
        public string xiaoshoufei_benyueshu { get; set; }//销售费用 本月数
        public string xiaoshoufei_bennianleijishu { get; set; }//销售费用 本年累计数
        public string xiaoshoufei_shangniantongqishu { get; set; }//销售费用 上年同期数
        public string qita_benyueshu { get; set; }//其他 本月数
        public string qita_bennianleijishu { get; set; }//其他 本年累计数
        public string qita_shangniantongqishu { get; set; }//其他 上年同期数
        public string Input_Date { get; set; }


        public string bianzhidanwei { get; set; }
        public string riqi { get; set; }
        public string danwei { get; set; }



        public string rowindex { get; set; }


    }
    public class clsQijianfeiyong_info
    {
        //项目		本月合计		环比增减	本年累计	上年同期		同比增减


        public string _id { get; set; }
        public string xiangmu { get; set; }
        public string benyueheji { get; set; }
        public string huanbizengjian { get; set; }
        public string bennianleiji { get; set; }
        public string shangniantongqi { get; set; }
        public string tongbizengjian { get; set; }
        public string Input_Date { get; set; }

    }
    public class clsmaolilv_info
    {
        //项目		本月合计		环比增减	本年累计	上年同期		同比增减


        public string _id { get; set; }
        public string xiangmu { get; set; }
        public string benyueheji { get; set; }
        public string huanbizengjian { get; set; }
        public string bennianleiji { get; set; }
        public string shangniantongqi { get; set; }
        public string tongbizengjian { get; set; }
        public string Input_Date { get; set; }

    }
    public class clscunhuo_info
    {
        //项目		本月合计		环比增减	本年累计	上年同期		同比增减
       // 项目	本月新增	环比增减	本年累计	上年同期	同比增减


        public string _id { get; set; }
        public string xiangmu { get; set; }
        public string benyuexinzheng { get; set; }
        public string huanbizengjian { get; set; }
        public string bennianleiji { get; set; }
        public string shangniantongqi { get; set; }
        public string tongbizengjian { get; set; }
        public string Input_Date { get; set; }

    }
    public class clsxianjinliu_info
    {
   // 项目(元)	本年金额	上年金额	同比变动

        public string _id { get; set; }
        public string xiangmu { get; set; }
        public string bennianjine { get; set; }
        public string shangnianjine { get; set; }
        public string tongbibiandong { get; set; }
       
        public string Input_Date { get; set; }

    }
}
