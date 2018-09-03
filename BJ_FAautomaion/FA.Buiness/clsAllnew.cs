using FA.Common;
using FA.DB;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FA.Buiness
{

    public class clsAllnew
    {
        private BackgroundWorker bgWorker1;
        //private object missing = System.Reflection.Missing.Value;
        public ToolStripProgressBar pbStatus { get; set; }
        public ToolStripStatusLabel tsStatusLabel1 { get; set; }
        public log4net.ILog ProcessLogger { get; set; }
        public log4net.ILog ExceptionLogger { get; set; }
        //财务 资产负债表
        public List<clszichanfuzaibiaoinfo> zichanfuzaibiao_Result;
        //主要经营指标完成情况
        public List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> zhuyao_Result;

        //财务 利润及利润分配表
        public List<clsLirunjilirunfenpeibiao_info> Lirunjilirunfenpei_Result;
        //财务 现金流量表
        public List<clsXianjinliu_info> Xianjinliu_Result;

        //八项费用支出表
        public List<cls8xiangfeiyongzhichu_info> baxiangfeiyong_Result;

        //期间费用情况
        public List<clsQijianfeiyong_info> qijianfeiyong_Result;

        //毛利率情况
        public List<clsmaolilv_info> maolilv_Result;

        //存货情况
        public List<clscunhuo_info> cunhuo_Result;

        //现金流净额
        public List<clsxianjinliu_info> xianjinliu_Result;



        public List<clszichanfuzaibiaoinfo> ReadDatasources(ref BackgroundWorker bgWorker, string filename)
        {
            zichanfuzaibiao_Result = new List<clszichanfuzaibiaoinfo>();
            zhuyao_Result = new List<clszhuyaojingyingzhibiaowanchengqingkuanginfo>();
            Lirunjilirunfenpei_Result = new List<clsLirunjilirunfenpeibiao_info>();
            baxiangfeiyong_Result = new List<cls8xiangfeiyongzhichu_info>();
            Xianjinliu_Result = new List<clsXianjinliu_info>();
            qijianfeiyong_Result = new List<clsQijianfeiyong_info>();
            cunhuo_Result = new List<clscunhuo_info>();
            xianjinliu_Result = new List<clsxianjinliu_info>();
            maolilv_Result = new List<clsmaolilv_info>();

            string path = AppDomain.CurrentDomain.BaseDirectory + "Resources";
            List<string> Alist = GetBy_CategoryReportFileName(path);

           

                for (int i = 0; i < Alist.Count; i++)
                {
                    GetKEYnfo(path + "\\" + Alist[i]);
                }


            return zichanfuzaibiao_Result;


        }
        //获取文件路径方法‘
        public List<string> GetBy_CategoryReportFileName(string dirPath)
        {

            List<string> FileNameList = new List<string>();
            ArrayList list = new ArrayList();

            if (Directory.Exists(dirPath))
            {
                list.AddRange(Directory.GetFiles(dirPath));
            }
            if (list.Count > 0)
            {
                foreach (object item in list)
                {
                    if (!item.ToString().Contains("~$"))
                        FileNameList.Add(item.ToString().Replace(dirPath + "\\", ""));
                }
            }

            return FileNameList;
        }
        //读取关键字
        public List<clszichanfuzaibiaoinfo> GetKEYnfo(string Alist)
        {



            try
            {
                List<clszichanfuzaibiaoinfo> WANGYINResult = new List<clszichanfuzaibiaoinfo>();
                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Microsoft.Office.Interop.Excel.Application excelApp;
                {
                    string path = Alist;
                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing,
                        "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["财务 资产负债表"];
                    Microsoft.Office.Interop.Excel.Range rng;
                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    int rowCount = WS.UsedRange.Rows.Count - 1;
                    object[,] o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    int wscount = analyWK.Worksheets.Count;

                    for (int i = 5; i <= rowCount; i++)
                    {
                        clszichanfuzaibiaoinfo temp = new clszichanfuzaibiaoinfo();

                        #region 基础信息

                        temp.xiangmu = "";
                        if (o[i, 1] != null)
                            temp.xiangmu = o[i, 1].ToString().Trim();

                        temp.hangci = "";
                        if (o[i, 2] != null)
                            temp.hangci = o[i, 2].ToString().Trim();
                        if (temp.hangci == null || temp.hangci == "")
                            continue;


                        temp.qimojine = "";
                        if (o[i, 3] != null)
                            temp.qimojine = o[i, 3].ToString().Trim();


                        temp.shangniantongqishu = "";
                        if (o[i, 4] != null)
                            temp.shangniantongqishu = o[i, 4].ToString().Trim();

                        temp.nianchujine = "";
                        if (o[i, 5] != null)
                            temp.nianchujine = o[i, 5].ToString().Trim();

                        temp.xiangmuF = "";
                        if (o[i, 6] != null)
                            temp.xiangmuF = o[i, 6].ToString().Trim();
                        temp.hangciG = "";
                        if (o[i, 7] != null)
                            temp.hangciG = o[i, 7].ToString().Trim();

                        temp.qimojineH = "";
                        if (o[i, 8] != null)
                            temp.qimojineH = o[i, 8].ToString().Trim();


                        temp.shangniantongqishuI = "";
                        if (o[i, 9] != null)
                            temp.shangniantongqishuI = o[i, 9].ToString().Trim();

                        temp.nianchujineJ = "";
                        if (o[i, 10] != null)
                            temp.nianchujineJ = o[i, 10].ToString().Trim();
                        //

                        temp.bianzhidanwei = "";
                        if (o[3, 1] != null)
                            temp.bianzhidanwei = o[3, 1].ToString().Trim();

                        temp.riqi = "";
                        if (o[3, 5] != null)
                            temp.riqi = o[3, 5].ToString().Trim();

                        temp.danwei = "";
                        if (o[3, 10] != null)
                            temp.danwei = o[3, 10].ToString().Trim();

                        temp.Input_Date = DateTime.Now.ToString("yyyy/MM/dd");

                        #endregion
                        zichanfuzaibiao_Result.Add(temp);
                    }

                    #region MyRegion
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["主要经营指标完成情况"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count - 1;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 3; i <= rowCount; i++)
                    {
                        clszhuyaojingyingzhibiaowanchengqingkuanginfo temp = new clszhuyaojingyingzhibiaowanchengqingkuanginfo();

                        #region 基础信息

                        temp.xuhao1 = "";
                        if (o[i, 1] != null)
                            temp.xuhao1 = o[i, 1].ToString().Trim();

                        temp.zhibiaomingcheng = "";
                        if (o[i, 2] != null)
                            temp.zhibiaomingcheng = o[i, 2].ToString().Trim();
                        if (temp.zhibiaomingcheng == null || temp.zhibiaomingcheng == "")
                            continue;


                        temp.nianchuzhibiaozhihuoqichushu = "";
                        if (o[i, 3] != null)
                            temp.nianchuzhibiaozhihuoqichushu = o[i, 3].ToString().Trim();


                        temp.benyuewancheng = "";
                        if (o[i, 4] != null)
                            temp.benyuewancheng = o[i, 4].ToString().Trim();

                        temp.huanbizengjian = "";
                        if (o[i, 5] != null)
                            temp.huanbizengjian = o[i, 5].ToString().Trim();

                        temp.leijiwanchenghuoqimoshu = "";
                        if (o[i, 6] != null)
                            temp.leijiwanchenghuoqimoshu = o[i, 6].ToString().Trim();
                        temp.wanchengbili = "";
                        if (o[i, 7] != null)
                            temp.wanchengbili = o[i, 7].ToString().Trim();

                        temp.shangniantongqileijiwancheng = "";
                        if (o[i, 8] != null)
                            temp.shangniantongqileijiwancheng = o[i, 8].ToString().Trim();


                        temp.tongbizengzhang = "";
                        if (o[i, 9] != null)
                            temp.tongbizengzhang = o[i, 9].ToString().Trim();


                        temp.danwei = "";
                        if (o[1, 1] != null)
                            temp.danwei = o[1, 1].ToString().Trim();

                        temp.Input_Date = DateTime.Now.ToString("yyyy/MM/dd");

                        #endregion
                        zhuyao_Result.Add(temp);
                    }
                    #endregion


                    #region 财务 利润及利润分配表
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["财务 利润及利润分配表"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count - 1;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 5; i <= rowCount; i++)
                    {
                        clsLirunjilirunfenpeibiao_info temp = new clsLirunjilirunfenpeibiao_info();

                        #region 基础信息

                        temp.xiangmu = "";
                        if (o[i, 1] != null)
                            temp.xiangmu = o[i, 1].ToString().Trim();

                        temp.hangci = "";
                        if (o[i, 2] != null)
                            temp.hangci = o[i, 2].ToString().Trim();
                        if (temp.hangci == null || temp.hangci == "")
                            continue;
                        temp.benyueshu = "";
                        if (o[i, 3] != null)
                            temp.benyueshu = o[i, 3].ToString().Trim();


                        temp.bennianleijishu = "";
                        if (o[i, 4] != null)
                            temp.bennianleijishu = o[i, 4].ToString().Trim();

                        temp.shangniantongqishu = "";
                        if (o[i, 5] != null)
                            temp.shangniantongqishu = o[i, 5].ToString().Trim();


                        temp.bianzhidanwei = "";
                        if (o[3, 1] != null)
                            temp.bianzhidanwei = o[3, 1].ToString().Trim();

                        temp.riqi = "";
                        if (o[3, 3] != null)
                            temp.riqi = o[3, 3].ToString().Trim();

                        temp.danwei = "";
                        if (o[5, 5] != null)
                            temp.danwei = o[5, 3].ToString().Trim();

                        temp.Input_Date = DateTime.Now.ToString("yyyy/MM/dd");



                        temp.rowindex = i.ToString().Trim();

                        #endregion
                        Lirunjilirunfenpei_Result.Add(temp);
                    }

                    #endregion

                    #region 财务 现金流量表
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["财务 现金流量表"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count - 1;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 5; i <= rowCount; i++)
                    {
                        clsXianjinliu_info temp = new clsXianjinliu_info();

                        #region 基础信息

                        temp.xiangmu = "";
                        if (o[i, 1] != null)
                            temp.xiangmu = o[i, 1].ToString().Trim();

                        temp.hangci = "";
                        if (o[i, 2] != null)
                            temp.hangci = o[i, 2].ToString().Trim();
                        if (temp.hangci == null || temp.hangci == "")
                            continue;
                        temp.bennianjine = "";
                        if (o[i, 3] != null)
                            temp.bennianjine = o[i, 3].ToString().Trim();


                        temp.shangnianjine = "";
                        if (o[i, 4] != null)
                            temp.shangnianjine = o[i, 4].ToString().Trim();


                        temp.bianzhidanwei = "";
                        if (o[3, 1] != null)
                            temp.bianzhidanwei = o[3, 1].ToString().Trim();

                        temp.riqi = "";
                        if (o[3, 3] != null)
                            temp.riqi = o[3, 3].ToString().Trim();

                        temp.danwei = "";
                        if (o[5, 5] != null)
                            temp.danwei = o[5, 3].ToString().Trim();

                        temp.Input_Date = DateTime.Now.ToString("yyyy/MM/dd");


                        temp.rowindex = i.ToString().Trim();


                        #endregion
                        Xianjinliu_Result.Add(temp);
                    }

                    #endregion

                    #region 八项费用支出表
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["八项费用支出表"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 50]];
                    rowCount = WS.UsedRange.Rows.Count - 1;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 6; i <= rowCount; i++)
                    {
                        cls8xiangfeiyongzhichu_info temp = new cls8xiangfeiyongzhichu_info();

                        #region 基础信息

                        temp.xiangmu = "";
                        if (o[i, 1] != null)
                            temp.xiangmu = o[i, 1].ToString().Trim();

                        temp.hangci = "";
                        if (o[i, 2] != null)
                            temp.hangci = o[i, 2].ToString().Trim();
                        if (temp.hangci == null || temp.hangci == "")
                            continue;

                        temp.shangnianquannianfasheng = "";
                        if (o[i, 3] != null)
                            temp.shangnianquannianfasheng = o[i, 3].ToString().Trim();

                        temp.nianduyusuan = "";
                        if (o[i, 4] != null)
                            temp.nianduyusuan = o[i, 4].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 4]);
                        temp.heji_benyueshu = "";
                        if (o[i, 5] != null)
                            temp.heji_benyueshu = o[i, 5].ToString().Trim();

                        temp.heji_bennianleiji = "";
                        if (o[i, 6] != null)
                            temp.heji_bennianleiji = o[i, 6].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 6]);

                        temp.heji_shangniantongqishu = "";
                        if (o[i, 7] != null)
                            temp.heji_shangniantongqishu = o[i, 7].ToString().Trim();
                        temp.zaijian_benyueshu = "";
                        if (o[i, 8] != null)
                            temp.zaijian_benyueshu = o[i, 8].ToString().Trim();

                        temp.zaijian_bennianleijishu = "";
                        if (o[i, 9] != null)
                            temp.zaijian_bennianleijishu = o[i, 9].ToString().Trim();

                        temp.zaijian_shangniantongqishu = "";
                        if (o[i, 10] != null)
                            temp.zaijian_shangniantongqishu = o[i, 10].ToString().Trim();

                        temp.xiangmuqian_benyueshu = "";
                        if (o[i, 11] != null)
                            temp.xiangmuqian_benyueshu = o[i, 11].ToString().Trim();
                        temp.xiangmuqian_bennianleijishu = "";
                        if (o[i, 12] != null)
                            temp.xiangmuqian_bennianleijishu = o[i, 12].ToString().Trim();

                        temp.xiangmuqian_shangniantongqishu = "";
                        if (o[i, 13] != null)
                            temp.xiangmuqian_shangniantongqishu = o[i, 13].ToString().Trim().ToUpper();


                        temp.gongchengshigong_benyueshu = "";
                        if (o[i, 14] != null)
                            temp.gongchengshigong_benyueshu = o[i, 14].ToString().Trim();


                        temp.gongchengshigong_bennianleijishu = "";
                        if (o[i, 15] != null)
                            temp.gongchengshigong_bennianleijishu = o[i, 15].ToString().Trim();


                        temp.gongchengshigong_shangniantongqishu = "";
                        if (o[i, 16] != null)
                            temp.gongchengshigong_shangniantongqishu = o[i, 16].ToString().Trim();


                        temp.shengchancheng_benyueshu = "";
                        if (o[i, 17] != null)
                            temp.shengchancheng_benyueshu = o[i, 17].ToString().Trim();


                        temp.shengchancheng_bennianleijishu = "";
                        if (o[i, 18] != null)
                            temp.shengchancheng_bennianleijishu = o[i, 18].ToString().Trim();


                        temp.shengchancheng_shangniantongqishu = "";
                        if (o[i, 19] != null)
                            temp.shengchancheng_shangniantongqishu = o[i, 19].ToString().Trim();


                        temp.guanlifei_benyueshu = "";
                        if (o[i, 20] != null)
                            temp.guanlifei_benyueshu = o[i, 20].ToString().Trim();


                        temp.guanlifei_bennianleijishu = "";
                        if (o[i, 21] != null)
                            temp.guanlifei_bennianleijishu = o[i, 21].ToString().Trim();

                        temp.guanlifei_shangniantongqishu = "";
                        if (o[i, 22] != null)
                            temp.guanlifei_shangniantongqishu = o[i, 22].ToString().Trim();

                        temp.xiaoshoufei_benyueshu = "";
                        if (o[i, 23] != null)
                            temp.xiaoshoufei_benyueshu = o[i, 23].ToString().Trim();

                        temp.xiaoshoufei_bennianleijishu = "";
                        if (o[i, 24] != null)
                            temp.xiaoshoufei_bennianleijishu = o[i, 24].ToString().Trim();

                        temp.xiaoshoufei_shangniantongqishu = "";
                        if (o[i, 25] != null)
                            temp.xiaoshoufei_shangniantongqishu = o[i, 25].ToString().Trim();

                        temp.qita_benyueshu = "";
                        if (o[i, 26] != null)
                            temp.qita_benyueshu = o[i, 26].ToString().Trim();

                        temp.qita_bennianleijishu = "";
                        if (o[i, 27] != null)
                            temp.qita_bennianleijishu = o[i, 27].ToString().Trim();

                        temp.qita_shangniantongqishu = "";
                        if (o[i, 28] != null)
                            temp.qita_shangniantongqishu = o[i, 28].ToString().Trim();

                        temp.Input_Date = DateTime.Now.ToString("yyyy/MM/dd");

                        temp.bianzhidanwei = "";
                        if (o[3, 1] != null)
                            temp.bianzhidanwei = o[3, 1].ToString().Trim();

                        temp.riqi = "";
                        if (o[3, 9] != null)
                            temp.riqi = o[3, 9].ToString().Trim();

                        temp.danwei = "";
                        if (o[3, 28] != null)
                            temp.danwei = o[3, 28].ToString().Trim();

                        temp.Input_Date = DateTime.Now.ToString("yyyy/MM/dd");


                        temp.rowindex = i.ToString().Trim();




                        #endregion
                        baxiangfeiyong_Result.Add(temp);
                    }

                    #endregion


                    #region 期间费用情况
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["期间费用情况"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 4; i <= rowCount; i++)
                    {
                        clsQijianfeiyong_info temp = new clsQijianfeiyong_info();

                        #region 基础信息

                        temp.xiangmu = "";
                        if (o[i, 1] != null)
                            temp.xiangmu = o[i, 1].ToString().Trim();

                        temp.benyueheji = "";
                        if (o[i, 2] != null)
                            temp.benyueheji = o[i, 2].ToString().Trim();
                        if (temp.xiangmu == null || temp.xiangmu == "")
                            continue;
                        temp.huanbizengjian = "";
                        if (o[i, 3] != null && o[i, 3].ToString() != "-2146826281")
                            temp.huanbizengjian = o[i, 3].ToString().Trim();


                        temp.bennianleiji = "";
                        if (o[i, 4] != null && o[i, 4].ToString() != "-2146826281")
                            temp.bennianleiji = o[i, 4].ToString().Trim();


                        temp.shangniantongqi = "";
                        if (o[i, 5] != null && o[i, 5].ToString() != "-2146826281")
                            temp.shangniantongqi = o[i, 5].ToString().Trim();

                        temp.bennianleiji = "";
                        if (o[i, 6] != null && o[i, 6].ToString() != "-2146826281")
                            temp.bennianleiji = o[i, 6].ToString().Trim();


                        temp.tongbizengjian = "";
                        if (o[i, 7] != null && o[i, 7].ToString() != "-2146826281")
                            temp.tongbizengjian = o[i, 7].ToString().Trim();


                        temp.Input_Date = DateTime.Now.ToString("yyyy/MM/dd");

                        #endregion
                        qijianfeiyong_Result.Add(temp);
                    }

                    #endregion
                    #region 毛利率情况
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["毛利率情况"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 4; i <= rowCount; i++)
                    {
                        clsmaolilv_info temp = new clsmaolilv_info();

                        #region 基础信息

                        temp.xiangmu = "";
                        if (o[i, 1] != null)
                            temp.xiangmu = o[i, 1].ToString().Trim();

                        temp.benyueheji = "";
                        if (o[i, 2] != null)
                            temp.benyueheji = o[i, 2].ToString().Trim();
                        if (temp.benyueheji == null || temp.benyueheji == "")
                            continue;
                        temp.huanbizengjian = "";
                        if (o[i, 3] != null)
                            temp.huanbizengjian = o[i, 3].ToString().Trim();


                        temp.bennianleiji = "";
                        if (o[i, 4] != null)
                            temp.bennianleiji = o[i, 4].ToString().Trim();


                        temp.shangniantongqi = "";
                        if (o[i, 5] != null)
                            temp.shangniantongqi = o[i, 5].ToString().Trim();
                        if (temp.shangniantongqi == "-2146826281")
                            temp.shangniantongqi = "0";

                        temp.bennianleiji = "";
                        if (o[i, 6] != null)
                            temp.bennianleiji = o[i, 6].ToString().Trim();

                        temp.tongbizengjian = "";
                        if (o[i, 7] != null)
                            temp.tongbizengjian = o[i, 7].ToString().Trim();


                        temp.Input_Date = DateTime.Now.ToString("yyyy/MM/dd");

                        #endregion
                        maolilv_Result.Add(temp);
                    }
                    #endregion
                    #region 存货情况
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["存货情况"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 4; i <= rowCount; i++)
                    {
                        clscunhuo_info temp = new clscunhuo_info();

                        #region 基础信息

                        temp.xiangmu = "";
                        if (o[i, 1] != null)
                            temp.xiangmu = o[i, 1].ToString().Trim();

                        temp.benyuexinzheng = "";
                        if (o[i, 2] != null)
                            temp.benyuexinzheng = o[i, 2].ToString().Trim();
                        if (temp.xiangmu == null || temp.xiangmu == "")
                            continue;
                        temp.huanbizengjian = "";
                        if (o[i, 3] != null)
                            temp.huanbizengjian = o[i, 3].ToString().Trim();


                        temp.bennianleiji = "";
                        if (o[i, 4] != null)
                            temp.bennianleiji = o[i, 4].ToString().Trim();


                        temp.shangniantongqi = "";
                        if (o[i, 5] != null)
                            temp.shangniantongqi = o[i, 5].ToString().Trim();

                        temp.bennianleiji = "";
                        if (o[i, 6] != null)
                            temp.bennianleiji = o[i, 6].ToString().Trim();

                        temp.tongbizengjian = "";
                        if (o[i, 7] != null)
                            temp.tongbizengjian = o[i, 7].ToString().Trim();
                        temp.Input_Date = DateTime.Now.ToString("yyyy/MM/dd");

                        #endregion

                        cunhuo_Result.Add(temp);
                    }
                    #endregion

                    #region 现金流净额
                    WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["现金流净额"];

                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    rowCount = WS.UsedRange.Rows.Count;
                    o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    wscount = analyWK.Worksheets.Count;

                    for (int i = 3; i <= rowCount; i++)
                    {
                        clsxianjinliu_info temp = new clsxianjinliu_info();

                        #region 基础信息

                        temp.xiangmu = "";
                        if (o[i, 1] != null)
                            temp.xiangmu = o[i, 1].ToString().Trim();

                        temp.bennianjine = "";
                        if (o[i, 2] != null)
                            temp.bennianjine = o[i, 2].ToString().Trim();
                        if (temp.bennianjine == null || temp.bennianjine == "")
                            continue;
                        temp.shangnianjine = "";
                        if (o[i, 3] != null)
                            temp.shangnianjine = o[i, 3].ToString().Trim();
                        temp.tongbibiandong = "";
                        if (o[i, 4] != null)
                            temp.tongbibiandong = o[i, 4].ToString().Trim();
                        temp.Input_Date = DateTime.Now.ToString("yyyy/MM/dd");

                        #endregion

                        xianjinliu_Result.Add(temp);
                    }
                    #endregion

                    clsCommHelp.CloseExcel(excelApp, analyWK);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: 01032" + ex);
                return null;

                throw;
            }
            return zichanfuzaibiao_Result;

        }



        public void InitializeDataSource(List<clszichanfuzaibiaoinfo> TBB1, List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> PDF_ChildType1, List<clsLirunjilirunfenpeibiao_info> Lirunjilirunfenpei_Result1, List<clsXianjinliu_info> Xianjinliu_Result1, List<cls8xiangfeiyongzhichu_info> baxiangfeiyong_Result1, List<clsQijianfeiyong_info> qijianfeiyong_Result1, List<clsmaolilv_info> maolilv_Result1, List<clscunhuo_info> cunhuo_Result1, List<clsxianjinliu_info> xianjinliuJINGE_Result1)
        {
            //excel
            zichanfuzaibiao_Result = new List<clszichanfuzaibiaoinfo>();
            zhuyao_Result = new List<clszhuyaojingyingzhibiaowanchengqingkuanginfo>();


            zichanfuzaibiao_Result = TBB1;
            zhuyao_Result = PDF_ChildType1;


            Lirunjilirunfenpei_Result = new List<clsLirunjilirunfenpeibiao_info>();

            Lirunjilirunfenpei_Result = Lirunjilirunfenpei_Result1;
            Xianjinliu_Result = new List<clsXianjinliu_info>();
            Xianjinliu_Result = Xianjinliu_Result1;

            baxiangfeiyong_Result = new List<cls8xiangfeiyongzhichu_info>();

            baxiangfeiyong_Result = baxiangfeiyong_Result1;

            qijianfeiyong_Result = new List<clsQijianfeiyong_info>();

            qijianfeiyong_Result = qijianfeiyong_Result1;

            maolilv_Result = new List<clsmaolilv_info>();
            maolilv_Result = maolilv_Result1;

            cunhuo_Result = new List<clscunhuo_info>();
            cunhuo_Result = cunhuo_Result1;

            xianjinliu_Result = new List<clsxianjinliu_info>();

            xianjinliu_Result = xianjinliuJINGE_Result1;



        }
        public void DownLoadExcel(ref BackgroundWorker bgWorker, string pathname)
        {
            bgWorker1 = bgWorker;

            #region 获取模板路径
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            string fullPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources\\"), "DATA.xlsx");
            SaveFileDialog sfdDownFile = new SaveFileDialog();
            sfdDownFile.OverwritePrompt = false;
            string DesktopPath = Convert.ToString(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            sfdDownFile.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";
            string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results\\");
            string[] temp1 = System.Text.RegularExpressions.Regex.Split(pathname, " ");
            sfdDownFile.FileName = pathname;

            string strExcelFileName = string.Empty;
            #endregion

            #region 导出前校验模板信息
            if (string.IsNullOrEmpty(sfdDownFile.FileName))
            {
                MessageBox.Show("File name can't be empty, please confirm, thanks!");
                return;
            }
            if (!File.Exists(fullPath))
            {
                MessageBox.Show("Template file does not exist, please confirm, thanks!");
                return;
            }
            else
            {
                strExcelFileName = sfdDownFile.FileName;
                strExcelFileName = pathname;

            }
            #endregion
            #region Excel 初始化

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            System.Reflection.Missing missingValue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel._Workbook ExcelBook =
            ExcelApp.Workbooks.Open(fullPath, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue);
            #endregion
            #region Sheet 初始化
            try
            {
                Microsoft.Office.Interop.Excel._Worksheet ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets["财务 资产负债表"];
                //打开时是否显示Excel
                //ExcelApp.Visible = true;
                //ExcelApp.ScreenUpdating = true;
            #endregion

                #region 填充数据
                int doing = 0;

                if (zichanfuzaibiao_Result != null)
                //foreach (clszichanfuzaibiaoinfo item in zichanfuzaibiao_Result)
                {

                    int RowIndex = 4;
                    doing++;

                    bgWorker1.ReportProgress(0, "导出进度  :  " + RowIndex.ToString() + "/" + zichanfuzaibiao_Result.Count.ToString());



                    //ExcelApp.Visible = true;
                    //ExcelApp.ScreenUpdating = true;
                    List<string> namelist = new List<string>();
                    Add_zichanfuzaibiao_leftRowName(namelist);

                    #region 财务 资产负债表
                    for (int i = 0; i < namelist.Count; i++)
                    {
                        List<clszichanfuzaibiaoinfo> zcfzb = zichanfuzaibiao_Result.FindAll(sQ => (sQ.xiangmu != null && sQ.xiangmu == namelist[i]));
                        RowIndex++;
                        if (zcfzb.Count > 0)
                        {

                            double nullableQty = (from s in zcfzb
                                                  where s.qimojine != null && s.qimojine != "" && s.qimojine != "－"
                                                  select Convert.ToDouble(s.qimojine)).Sum();

                            ExcelSheet.Cells[RowIndex, 3] = nullableQty.ToString();// zcfzb[0].qimojine;//流动资产：

                            nullableQty = (from s in zcfzb
                                           where s.shangniantongqishu != null && s.shangniantongqishu != "" && s.shangniantongqishu != "－"
                                           select Convert.ToDouble(s.shangniantongqishu)).Sum();

                            ExcelSheet.Cells[RowIndex, 4] = nullableQty.ToString(); //zcfzb[0].shangniantongqishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.nianchujine != null && s.nianchujine != "" && s.nianchujine != "－"
                                           select Convert.ToDouble(s.nianchujine)).Sum();

                            ExcelSheet.Cells[RowIndex, 5] = nullableQty.ToString();// zcfzb[0].nianchujine;// ：
                        }

                    }
                    namelist = new List<string>();
                    Add_zichanfuzaibiao_RightRowName(namelist);

                    RowIndex = 4;
                    for (int i = 0; i < namelist.Count; i++)
                    {
                        List<clszichanfuzaibiaoinfo> zcfzb = zichanfuzaibiao_Result.FindAll(sQ => (sQ.xiangmuF != null && sQ.xiangmuF == namelist[i]));
                        RowIndex++;
                        if (zcfzb.Count > 0)
                        {
                            double nullableQty = (from s in zcfzb
                                                  where s.qimojine != null && s.qimojine != "" && s.qimojine != "－"
                                                  select Convert.ToDouble(s.qimojine)).Sum();


                            ExcelSheet.Cells[RowIndex, 8] = nullableQty.ToString();// zcfzb[0].qimojine;//流动资产：

                            nullableQty = (from s in zcfzb
                                           where s.shangniantongqishu != null && s.shangniantongqishu != "" && s.shangniantongqishu != "－"
                                           select Convert.ToDouble(s.shangniantongqishu)).Sum();

                            ExcelSheet.Cells[RowIndex, 9] = nullableQty.ToString();// zcfzb[0].shangniantongqishu;// ：

                            nullableQty = (from s in zcfzb
                                           where s.nianchujine != null && s.nianchujine != "" && s.nianchujine != "－"
                                           select Convert.ToDouble(s.nianchujine)).Sum();

                            ExcelSheet.Cells[RowIndex, 10] = nullableQty.ToString();// zcfzb[0].nianchujine;// ：
                        }

                    }
                    #endregion

                    #region 主要经营指标完成情况
                    ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets["主要经营指标完成情况"];

                    RowIndex = 2;
                    namelist = new List<string>();
                    Add_zhuyao_RowName(namelist);
                    for (int i = 0; i < namelist.Count; i++)
                    {
                        List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> zcfzb = zhuyao_Result.FindAll(sQ => (sQ.zhibiaomingcheng != null && sQ.zhibiaomingcheng == namelist[i]));
                        RowIndex++;
                        if (zcfzb.Count > 0)
                        {
                            double nullableQty = (from s in zcfzb
                                                  where s.benyuewancheng != null && s.benyuewancheng != "" && s.benyuewancheng != "－"
                                                  select Convert.ToDouble(s.benyuewancheng)).Sum();


                            ExcelSheet.Cells[RowIndex, 4] = nullableQty.ToString(); //zcfzb[0].benyuewancheng;//流动资产：
                            nullableQty = (from s in zcfzb
                                           where s.huanbizengjian != null && s.huanbizengjian != "" && !s.huanbizengjian.Contains("-")
                                           select Convert.ToDouble(s.huanbizengjian)).Sum();

                            ExcelSheet.Cells[RowIndex, 5] = nullableQty.ToString();// zcfzb[0].huanbizengjian;// ：
                            nullableQty = (from s in zcfzb
                                           where s.leijiwanchenghuoqimoshu != null && s.leijiwanchenghuoqimoshu != "" && s.leijiwanchenghuoqimoshu != "－"
                                           select Convert.ToDouble(s.leijiwanchenghuoqimoshu)).Sum();

                            ExcelSheet.Cells[RowIndex, 6] = nullableQty.ToString(); //zcfzb[0].leijiwanchenghuoqimoshu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.wanchengbili != null && s.wanchengbili != "" && !s.wanchengbili.Contains("-")
                                           select Convert.ToDouble(s.wanchengbili)).Sum();

                            ExcelSheet.Cells[RowIndex, 7] = nullableQty.ToString();// zcfzb[0].wanchengbili;// ：
                            nullableQty = (from s in zcfzb
                                           where s.shangniantongqileijiwancheng != null && s.shangniantongqileijiwancheng != "" && s.shangniantongqileijiwancheng != "－"
                                           select Convert.ToDouble(s.shangniantongqileijiwancheng)).Sum();

                            ExcelSheet.Cells[RowIndex, 8] = nullableQty.ToString();// zcfzb[0].shangniantongqileijiwancheng;// ：
                            nullableQty = (from s in zcfzb
                                           where s.tongbizengzhang != null && s.tongbizengzhang != "" && s.tongbizengzhang != "－"
                                           select Convert.ToDouble(s.tongbizengzhang)).Sum();

                            ExcelSheet.Cells[RowIndex, 9] = nullableQty.ToString();//zcfzb[0].tongbizengzhang;// ：
                        }
                    }
                    #endregion

                    #region 期间费用情况
                    ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets["期间费用情况"];

                    RowIndex = 3;
                    namelist = new List<string>();
                    Add_qijianfeiyong_RowName(namelist);
                    for (int i = 0; i < namelist.Count; i++)
                    {

                        List<clsQijianfeiyong_info> zcfzb = qijianfeiyong_Result.FindAll(sQ => (sQ.xiangmu != null && sQ.xiangmu == namelist[i]));
                        RowIndex++;
                        if (zcfzb.Count > 0)
                        {
                            double nullableQty = (from s in zcfzb
                                                  where s.benyueheji != null && s.benyueheji != "" && s.benyueheji != "－"
                                                  select Convert.ToDouble(s.benyueheji)).Sum();

                            ExcelSheet.Cells[RowIndex, 4] = nullableQty.ToString(); //zcfzb[0].benyueheji;//流动资产：

                            nullableQty = (from s in zcfzb
                                           where s.huanbizengjian != null && s.huanbizengjian != "" && s.huanbizengjian != "－"
                                           select Convert.ToDouble(s.huanbizengjian)).Sum();


                            ExcelSheet.Cells[RowIndex, 5] = nullableQty.ToString();// zcfzb[0].huanbizengjian;// ：
                            nullableQty = (from s in zcfzb
                                           where s.bennianleiji != null && s.bennianleiji != "" && s.bennianleiji != "－"
                                           select Convert.ToDouble(s.bennianleiji)).Sum();


                            ExcelSheet.Cells[RowIndex, 6] = nullableQty.ToString(); //zcfzb[0].bennianleiji;// ：
                            nullableQty = (from s in zcfzb
                                           where s.shangniantongqi != null && s.shangniantongqi != "" && s.shangniantongqi != "－"
                                           select Convert.ToDouble(s.shangniantongqi)).Sum();

                            ExcelSheet.Cells[RowIndex, 7] = nullableQty.ToString();// zcfzb[0].shangniantongqi;// ：

                            nullableQty = (from s in zcfzb
                                           where s.tongbizengjian != null && s.tongbizengjian != "" && s.tongbizengjian != "－"
                                           select Convert.ToDouble(s.tongbizengjian)).Sum();



                            ExcelSheet.Cells[RowIndex, 9] = nullableQty.ToString(); //zcfzb[0].tongbizengjian;// ：
                        }
                    }
                    #endregion

                    #region 毛利率情况
                    ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets["毛利率情况"];

                    RowIndex = 3;
                    namelist = new List<string>();
                    Add_maolilv_RowName(namelist);
                    for (int i = 0; i < namelist.Count; i++)
                    {

                        List<clsmaolilv_info> zcfzb = maolilv_Result.FindAll(sQ => (sQ.xiangmu != null && sQ.xiangmu == namelist[i]));
                        RowIndex++;
                        if (zcfzb.Count > 0)
                        {
                            double nullableQty = (from s in zcfzb
                                                  where s.benyueheji != null && s.benyueheji != "" && s.benyueheji != "－"
                                                  select Convert.ToDouble(s.benyueheji)).Sum();

                            ExcelSheet.Cells[RowIndex, 2] = nullableQty.ToString(); //zcfzb[0].benyueheji;//流动资产：
                            nullableQty = (from s in zcfzb
                                           where s.huanbizengjian != null && s.huanbizengjian != "" && s.huanbizengjian != "－"
                                           select Convert.ToDouble(s.huanbizengjian)).Sum();


                            ExcelSheet.Cells[RowIndex, 3] = nullableQty.ToString();// zcfzb[0].huanbizengjian;// ：
                            nullableQty = (from s in zcfzb
                                           where s.bennianleiji != null && s.bennianleiji != "" && s.bennianleiji != "－"
                                           select Convert.ToDouble(s.bennianleiji)).Sum();

                            ExcelSheet.Cells[RowIndex, 4] = nullableQty.ToString(); //zcfzb[0].bennianleiji;// ：
                            nullableQty = (from s in zcfzb
                                           where s.shangniantongqi != null && s.shangniantongqi != "" && s.shangniantongqi != "－"
                                           select Convert.ToDouble(s.shangniantongqi)).Sum();

                            ExcelSheet.Cells[RowIndex, 5] = nullableQty.ToString();// zcfzb[0].shangniantongqi;// ：
                            nullableQty = (from s in zcfzb
                                           where s.tongbizengjian != null && s.tongbizengjian != "" && s.tongbizengjian != "－"
                                           select Convert.ToDouble(s.tongbizengjian)).Sum();

                            ExcelSheet.Cells[RowIndex, 6] = nullableQty.ToString();// zcfzb[0].tongbizengjian;// ：
                        }
                    }
                    #endregion

                    #region 存货情况
                    ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets["存货情况"];

                    RowIndex = 3;
                    namelist = new List<string>();
                    Add_cunhuo_RowName(namelist);
                    for (int i = 0; i < namelist.Count; i++)
                    {

                        List<clscunhuo_info> zcfzb = cunhuo_Result.FindAll(sQ => (sQ.xiangmu != null && sQ.xiangmu == namelist[i]));
                        RowIndex++;
                        if (zcfzb.Count > 0)
                        {
                            double nullableQty = (from s in zcfzb
                                                  where s.benyuexinzheng != null && s.benyuexinzheng != "" && s.benyuexinzheng != "－"
                                                  select Convert.ToDouble(s.benyuexinzheng)).Sum();

                            ExcelSheet.Cells[RowIndex, 2] = nullableQty.ToString();// zcfzb[0].benyuexinzheng;//流动资产：
                            nullableQty = (from s in zcfzb
                                           where s.huanbizengjian != null && s.huanbizengjian != "" && s.huanbizengjian != "－"
                                           select Convert.ToDouble(s.huanbizengjian)).Sum();


                            ExcelSheet.Cells[RowIndex, 3] = nullableQty.ToString();// zcfzb[0].huanbizengjian;// ：
                            nullableQty = (from s in zcfzb
                                           where s.bennianleiji != null && s.bennianleiji != "" && s.bennianleiji != "－"
                                           select Convert.ToDouble(s.bennianleiji)).Sum();

                            ExcelSheet.Cells[RowIndex, 4] = nullableQty.ToString();// zcfzb[0].bennianleiji;// ：
                            nullableQty = (from s in zcfzb
                                           where s.shangniantongqi != null && s.shangniantongqi != "" && s.shangniantongqi != "－"
                                           select Convert.ToDouble(s.shangniantongqi)).Sum();

                            ExcelSheet.Cells[RowIndex, 5] = nullableQty.ToString();// zcfzb[0].shangniantongqi;// ：
                            nullableQty = (from s in zcfzb
                                           where s.tongbizengjian != null && s.tongbizengjian != "" && s.tongbizengjian != "－"
                                           select Convert.ToDouble(s.tongbizengjian)).Sum();


                            ExcelSheet.Cells[RowIndex, 6] = nullableQty.ToString(); //zcfzb[0].tongbizengjian;// ：
                        }
                    }
                    #endregion
                    #region 现金流净额
                    ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets["现金流净额"];

                    RowIndex = 2;
                    namelist = new List<string>();
                    Add_Add_xianjinliuJINGE_RowName(namelist);
                    for (int i = 0; i < namelist.Count; i++)
                    {

                        List<clsxianjinliu_info> zcfzb = xianjinliu_Result.FindAll(sQ => (sQ.xiangmu != null && sQ.xiangmu == namelist[i]));
                        RowIndex++;
                        if (zcfzb.Count > 0)
                        {
                            double nullableQty = (from s in zcfzb
                                                  where s.bennianjine != null && s.bennianjine != "" && s.bennianjine != "－"
                                                  select Convert.ToDouble(s.bennianjine)).Sum();



                            ExcelSheet.Cells[RowIndex, 2] = nullableQty.ToString(); //zcfzb[0].bennianjine;//流动资产：

                            nullableQty = (from s in zcfzb
                                           where s.shangnianjine != null && s.shangnianjine != "" && s.shangnianjine != "－"
                                           select Convert.ToDouble(s.shangnianjine)).Sum();

                            ExcelSheet.Cells[RowIndex, 3] = nullableQty.ToString();// zcfzb[0].shangnianjine;// ：

                            nullableQty = (from s in zcfzb
                                           where s.tongbibiandong != null && s.tongbibiandong != "" && s.tongbibiandong != "－"
                                           select Convert.ToDouble(s.tongbibiandong)).Sum();

                            ExcelSheet.Cells[RowIndex, 4] = nullableQty.ToString();// zcfzb[0].tongbibiandong;// ：
                        }
                    }
                    #endregion


                    #region 财务 利润及利润分配表
                    ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets["财务 利润及利润分配表"];

                    RowIndex = 3;
                    namelist = new List<string>();
                    //Add_Lirunjilirunfenpei_RowName(namelist);

                    namelist = (from v in Lirunjilirunfenpei_Result select v.xiangmu).Distinct().ToList();

                    for (int i = 0; i < namelist.Count; i++)
                    {
                        List<clsLirunjilirunfenpeibiao_info> zcfzb = Lirunjilirunfenpei_Result.FindAll(sQ => (sQ.xiangmu != null && sQ.xiangmu == namelist[i]));
                        RowIndex++;
                        if (zcfzb.Count > 0)
                        {
                            ExcelApp.Visible = true;
                            ExcelApp.ScreenUpdating = true;

                            double nullableQty = (from s in zcfzb
                                                  where s.hangci != null && s.hangci != "" && s.hangci != "－"
                                                  select Convert.ToDouble(s.hangci)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 2] = nullableQty.ToString(); //zcfzb[0].hangci;// ：
                            nullableQty = (from s in zcfzb
                                           where s.benyueshu != null && s.benyueshu != "" && s.benyueshu != "－"
                                           select Convert.ToDouble(s.benyueshu)).Sum();


                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 3] = nullableQty.ToString();// zcfzb[0].benyueshu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.bennianleijishu != null && s.bennianleijishu != "" && s.bennianleijishu != "－"
                                           select Convert.ToDouble(s.bennianleijishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 4] = nullableQty.ToString(); //zcfzb[0].bennianleijishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.shangniantongqi != null && s.shangniantongqi != "" && s.shangniantongqi != "－"
                                           select Convert.ToDouble(s.shangniantongqi)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 5] = nullableQty.ToString();// zcfzb[0].shangniantongqi;// ：                           
                        }
                    }
                    #endregion

                    #region 财务 现金流量表
                    ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets["财务 现金流量表"];

                    RowIndex = 3;
                    namelist = new List<string>();

                    namelist = (from v in Xianjinliu_Result select v.xiangmu).Distinct().ToList();

                    for (int i = 0; i < namelist.Count; i++)
                    {
                        List<clsXianjinliu_info> zcfzb = Xianjinliu_Result.FindAll(sQ => (sQ.xiangmu != null && sQ.xiangmu == namelist[i]));
                        RowIndex++;
                        if (zcfzb.Count > 0)
                        {
                            double nullableQty = (from s in zcfzb
                                                  where s.hangci != null && s.hangci != "" && s.hangci != "－"
                                                  select Convert.ToDouble(s.hangci)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 2] = nullableQty.ToString();// zcfzb[0].hangci;// ：
                            nullableQty = (from s in zcfzb
                                           where s.bennianjine != null && s.bennianjine != "" && s.bennianjine != "－"
                                           select Convert.ToDouble(s.bennianjine)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 3] = nullableQty.ToString();// zcfzb[0].bennianjine;// ：
                            nullableQty = (from s in zcfzb
                                           where s.shangnianjine != null && s.shangnianjine != "" && s.shangnianjine != "－"
                                           select Convert.ToDouble(s.shangnianjine)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 4] = nullableQty.ToString();// zcfzb[0].shangnianjine;// ：
                        }
                    }
                    #endregion

                    #region 八项费用支出表
                    ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets["八项费用支出表"];

                    RowIndex = 3;
                    namelist = new List<string>();

                    namelist = (from v in baxiangfeiyong_Result select v.xiangmu).Distinct().ToList();

                    for (int i = 0; i < namelist.Count; i++)
                    {

                        List<cls8xiangfeiyongzhichu_info> zcfzb = baxiangfeiyong_Result.FindAll(sQ => (sQ.xiangmu != null && sQ.xiangmu == namelist[i]));
                        RowIndex++;
                        if (zcfzb.Count > 0)
                        {
                            #region MyRegion
                            double nullableQty = (from s in zcfzb
                                                  where s.hangci != null && s.hangci != "" && s.hangci != "-"
                                                  select Convert.ToDouble(s.hangci)).Sum();


                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 2] = nullableQty.ToString(); //zcfzb[0].hangci;// ：
                            nullableQty = (from s in zcfzb
                                           where s.shangnianquannianfasheng != null && s.shangnianquannianfasheng != "" && s.shangnianquannianfasheng != "－"
                                           select Convert.ToDouble(s.shangnianquannianfasheng)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 3] = nullableQty.ToString(); //zcfzb[0].shangnianquannianfasheng;// ：
                            nullableQty = (from s in zcfzb
                                           where s.nianduyusuan != null && s.nianduyusuan != "" && s.nianduyusuan != "－"
                                           select Convert.ToDouble(s.nianduyusuan)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 4] = nullableQty.ToString();// zcfzb[0].nianduyusuan;// ：
                            nullableQty = (from s in zcfzb
                                           where s.heji_benyueshu != null && s.heji_benyueshu != "" && s.heji_benyueshu != "－"
                                           select Convert.ToDouble(s.heji_benyueshu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 5] = nullableQty.ToString(); //zcfzb[0].heji_benyueshu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.heji_bennianleiji != null && s.heji_bennianleiji != "" && s.heji_bennianleiji != "－"
                                           select Convert.ToDouble(s.heji_bennianleiji)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 6] = nullableQty.ToString();// zcfzb[0].heji_bennianleiji;// ：
                            nullableQty = (from s in zcfzb
                                           where s.heji_shangniantongqishu != null && s.heji_shangniantongqishu != "" && s.heji_shangniantongqishu != "－"
                                           select Convert.ToDouble(s.heji_shangniantongqishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 7] = nullableQty.ToString(); //zcfzb[0].heji_shangniantongqishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.zaijian_benyueshu != null && s.zaijian_benyueshu != "" && s.zaijian_benyueshu != "－"
                                           select Convert.ToDouble(s.zaijian_benyueshu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 8] = nullableQty.ToString();// zcfzb[0].zaijian_benyueshu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.zaijian_bennianleijishu != null && s.zaijian_bennianleijishu != "" && s.zaijian_bennianleijishu != "－"
                                           select Convert.ToDouble(s.zaijian_bennianleijishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 9] = nullableQty.ToString();// zcfzb[0].zaijian_bennianleijishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.zaijian_shangniantongqishu != null && s.zaijian_shangniantongqishu != "" && s.zaijian_shangniantongqishu != "－"
                                           select Convert.ToDouble(s.zaijian_shangniantongqishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 10] = nullableQty.ToString(); //zcfzb[0].zaijian_shangniantongqishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.xiangmuqian_benyueshu != null && s.xiangmuqian_benyueshu != "" && s.xiangmuqian_benyueshu != "－"
                                           select Convert.ToDouble(s.xiangmuqian_benyueshu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 11] = nullableQty.ToString(); //zcfzb[0].xiangmuqian_benyueshu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.xiangmuqian_bennianleijishu != null && s.xiangmuqian_bennianleijishu != "" && s.xiangmuqian_bennianleijishu != "－"
                                           select Convert.ToDouble(s.xiangmuqian_bennianleijishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 12] = nullableQty.ToString(); //zcfzb[0].xiangmuqian_bennianleijishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.xiangmuqian_shangniantongqishu != null && s.xiangmuqian_shangniantongqishu != "" && s.xiangmuqian_shangniantongqishu != "－"
                                           select Convert.ToDouble(s.xiangmuqian_shangniantongqishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 13] = nullableQty.ToString();// zcfzb[0].xiangmuqian_shangniantongqishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.gongchengshigong_benyueshu != null && s.gongchengshigong_benyueshu != "" && s.gongchengshigong_benyueshu != "－"
                                           select Convert.ToDouble(s.gongchengshigong_benyueshu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 14] = nullableQty.ToString();// zcfzb[0].gongchengshigong_benyueshu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.gongchengshigong_bennianleijishu != null && s.gongchengshigong_bennianleijishu != "" && s.gongchengshigong_bennianleijishu != "－"
                                           select Convert.ToDouble(s.gongchengshigong_bennianleijishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 15] = nullableQty.ToString();// zcfzb[0].gongchengshigong_bennianleijishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.gongchengshigong_shangniantongqishu != null && s.gongchengshigong_shangniantongqishu != "" && s.gongchengshigong_shangniantongqishu != "－"
                                           select Convert.ToDouble(s.gongchengshigong_shangniantongqishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 16] = nullableQty.ToString();// zcfzb[0].gongchengshigong_shangniantongqishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.shengchancheng_benyueshu != null && s.shengchancheng_benyueshu != "" && s.shengchancheng_benyueshu != "－"
                                           select Convert.ToDouble(s.shengchancheng_benyueshu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 17] = nullableQty.ToString();// zcfzb[0].shengchancheng_benyueshu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.shengchancheng_bennianleijishu != null && s.shengchancheng_bennianleijishu != "" && s.shengchancheng_bennianleijishu != "－"
                                           select Convert.ToDouble(s.shengchancheng_bennianleijishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 18] = nullableQty.ToString();// zcfzb[0].shengchancheng_bennianleijishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.shengchancheng_shangniantongqishu != null && s.shengchancheng_shangniantongqishu != "" && s.shengchancheng_shangniantongqishu != "－"
                                           select Convert.ToDouble(s.shengchancheng_shangniantongqishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 19] = nullableQty.ToString(); //zcfzb[0].shengchancheng_shangniantongqishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.guanlifei_benyueshu != null && s.guanlifei_benyueshu != "" && s.guanlifei_benyueshu != "－"
                                           select Convert.ToDouble(s.guanlifei_benyueshu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 20] = nullableQty.ToString();// zcfzb[0].guanlifei_benyueshu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.guanlifei_bennianleijishu != null && s.guanlifei_bennianleijishu != "" && s.guanlifei_bennianleijishu != "－"
                                           select Convert.ToDouble(s.guanlifei_bennianleijishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 21] = nullableQty.ToString();// zcfzb[0].guanlifei_bennianleijishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.guanlifei_shangniantongqishu != null && s.guanlifei_shangniantongqishu != "" && s.guanlifei_shangniantongqishu != "－"
                                           select Convert.ToDouble(s.guanlifei_shangniantongqishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 22] = nullableQty.ToString(); //zcfzb[0].guanlifei_shangniantongqishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.xiaoshoufei_benyueshu != null && s.xiaoshoufei_benyueshu != "" && s.xiaoshoufei_benyueshu != "－"
                                           select Convert.ToDouble(s.xiaoshoufei_benyueshu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 23] = nullableQty.ToString();// zcfzb[0].xiaoshoufei_benyueshu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.xiaoshoufei_bennianleijishu != null && s.xiaoshoufei_bennianleijishu != "" && s.xiaoshoufei_bennianleijishu != "－"
                                           select Convert.ToDouble(s.xiaoshoufei_bennianleijishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 24] = nullableQty.ToString();// zcfzb[0].xiaoshoufei_bennianleijishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.xiaoshoufei_shangniantongqishu != null && s.xiaoshoufei_shangniantongqishu != "" && s.xiaoshoufei_shangniantongqishu != "－"
                                           select Convert.ToDouble(s.xiaoshoufei_shangniantongqishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 25] = nullableQty.ToString();// zcfzb[0].xiaoshoufei_shangniantongqishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.qita_benyueshu != null && s.qita_benyueshu != "" && s.qita_benyueshu != "－"
                                           select Convert.ToDouble(s.qita_benyueshu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 26] = nullableQty.ToString();// zcfzb[0].qita_benyueshu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.qita_bennianleijishu != null && s.qita_bennianleijishu != "" && s.qita_bennianleijishu != "－"
                                           select Convert.ToDouble(s.qita_bennianleijishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 27] = nullableQty.ToString();// zcfzb[0].qita_bennianleijishu;// ：
                            nullableQty = (from s in zcfzb
                                           where s.qita_shangniantongqishu != null && s.qita_shangniantongqishu != "" && s.qita_shangniantongqishu != "－"
                                           select Convert.ToDouble(s.qita_shangniantongqishu)).Sum();

                            ExcelSheet.Cells[Convert.ToInt32(zcfzb[0].rowindex), 28] = nullableQty.ToString();// zcfzb[0].qita_shangniantongqishu;// ：

                            #endregion
                        }
                    }
                    #endregion

                }
                ExcelBook.RefreshAll();
                #region 写入文件
                ExcelApp.ScreenUpdating = true;
                if (doing != 0)
                    ExcelBook.SaveAs(strExcelFileName, missingValue, missingValue, missingValue, missingValue, missingValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingValue, missingValue, missingValue, missingValue, missingValue);
                ExcelApp.DisplayAlerts = false;

                #endregion
            }

            #region 异常处理
            catch (Exception ex)
            {
                ExcelApp.DisplayAlerts = false;
                ExcelApp.Quit();
                ExcelBook = null;
                ExcelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                throw ex;
            }
            #endregion

            #region Finally垃圾回收
            finally
            {
                ExcelBook.Close(false, missingValue, missingValue);
                ExcelBook = null;
                ExcelApp.DisplayAlerts = true;
                ExcelApp.Quit();
                clsKeyMyExcelProcess.Kill(ExcelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            #endregion

                #endregion
        }

        private static void Add_zichanfuzaibiao_leftRowName(List<string> namelist)
        {
            #region 左侧列
            namelist.Add("流动资产：");
            namelist.Add("货币资金");
            namelist.Add("△结算备付金");
            namelist.Add("△拆出资金");
            namelist.Add("以公允价值计量且其变动计入当期损益的金融资产");
            namelist.Add("衍生金融资产");
            namelist.Add("应收票据");
            namelist.Add("应收账款");
            namelist.Add("其中:应收电费");
            namelist.Add("应收热费");
            namelist.Add("应收煤款");
            namelist.Add("减：坏账准备");
            namelist.Add("应收账款净额");
            namelist.Add("预付款项");
            namelist.Add("△应收保费");
            namelist.Add("△应收分保账款");
            namelist.Add("△应收分保准备金");
            namelist.Add("应收利息");
            namelist.Add("应收股利");
            namelist.Add("其他应收款");
            namelist.Add("减：坏账准备");
            namelist.Add("其他应收账款净额");
            namelist.Add("△买入返售金融资产");
            namelist.Add("存货");
            namelist.Add("其中:原材料");
            namelist.Add("其中：燃料");
            namelist.Add("库存商品(产成品)");
            namelist.Add("其中：煤炭");
            namelist.Add("工程施工");
            namelist.Add("划分为持有待售的资产");
            namelist.Add("一年内到期的非流动资产");
            namelist.Add("其他流动资产");
            namelist.Add("流动资产合计");
            namelist.Add("非流动资产");
            namelist.Add("△发放贷款及垫款");
            namelist.Add("可供出售金融资产");
            namelist.Add("持有至到期投资");
            namelist.Add("长期应收款");
            namelist.Add("长期股权投资");
            namelist.Add("拨付所属资金");
            namelist.Add("投资性房地产");
            namelist.Add("固定资产原价");
            namelist.Add("减：累计折旧");
            namelist.Add("固定资产净值");
            namelist.Add("减：固定资产减值准备");
            namelist.Add("固定资产净额");
            namelist.Add("固定资产净额");
            namelist.Add("工程物资");
            namelist.Add("工程物资");
            namelist.Add("生产性生物资产");
            namelist.Add("油气资产");
            namelist.Add("油气资产");
            namelist.Add("其中：土地使用权");
            namelist.Add("开发支出");
            namelist.Add("商誉");
            namelist.Add("长期待摊费用");
            namelist.Add("长期待摊费用");
            namelist.Add("其他非流动资产");
            namelist.Add("非流动资产合计");
            namelist.Add("－");
            namelist.Add("－");
            namelist.Add("－");
            namelist.Add("－");
            namelist.Add("－");
            namelist.Add("资  产  总  计");
            #endregion
        }
        private static void Add_zichanfuzaibiao_RightRowName(List<string> namelist)
        {
            #region 左侧列
            namelist.Add("流动负债：");
            namelist.Add("短期借款");
            namelist.Add("△向中央银行借款");
            namelist.Add("△吸收存款及同业存放");
            namelist.Add("△拆入资金");
            namelist.Add("以公允价值计量且其变动计入当期损益的金融负债");
            namelist.Add("衍生金融负债");
            namelist.Add("应付票据");
            namelist.Add("应付账款");
            namelist.Add("预收款项");
            namelist.Add("△卖出回购金融资产款");
            namelist.Add("△应付手续费及佣金");
            namelist.Add("应付职工薪酬");
            namelist.Add("应交税费");
            namelist.Add("其中：应交税金");
            namelist.Add("应付利息");
            namelist.Add("应付股利");
            namelist.Add("其他应付款");
            namelist.Add("△应付分保账款");
            namelist.Add("△保险合同准备金");
            namelist.Add("△代理买卖证券款");
            namelist.Add("△代理承销证券款");
            namelist.Add("内部往来");
            namelist.Add("划分为持有待售的负债");
            namelist.Add("一年内到期的非流动负债");
            namelist.Add("其他流动负债");
            namelist.Add("流动负债合计");
            namelist.Add("非流动负债：");
            namelist.Add("长期借款");
            namelist.Add("应付债券");
            namelist.Add("长期应付款");
            namelist.Add("长期应付职工薪酬");
            namelist.Add("专项应付款");
            namelist.Add("递延收益");
            namelist.Add("预计负债");
            namelist.Add("递延所得税负债");
            namelist.Add("其他非流动负债");
            namelist.Add("非流动负债合计");
            namelist.Add("负 债 合 计");
            namelist.Add("上级拨入资金");
            namelist.Add("所有者权益（或股东权益）：");
            namelist.Add("实收资本（股本）");
            namelist.Add("国有资本");
            namelist.Add("其中：国有法人资本");
            namelist.Add("集体资本");
            namelist.Add("民营资本");
            namelist.Add("其中： 个人资本");
            namelist.Add("外商资本");
            namelist.Add("#减：已归还投资");
            namelist.Add("实收资本（或股本）净额");
            namelist.Add("其他权益工具");
            namelist.Add("其中:优先股");
            namelist.Add("永续债");
            namelist.Add("资本公积");
            namelist.Add("减：库存股");
            namelist.Add("其他综合收益");
            namelist.Add("其中：外币报表折算差额");
            namelist.Add("专项储备");
            namelist.Add("盈余公积");
            namelist.Add("△一般风险准备");
            namelist.Add("未分配利润");
            namelist.Add("归属于母公司所有者权益合计");
            namelist.Add("*少数股东权益");
            namelist.Add("所有者权益合计");
            namelist.Add("负债和所有者权益总计");

            #endregion
        }
        private static void Add_zhuyao_RowName(List<string> namelist)
        {
            #region 左侧列
            namelist.Add("资产总额");
            namelist.Add("负债总额");
            namelist.Add("资产负债率");
            namelist.Add("营业收入");
            namelist.Add("利润总额");
            namelist.Add("期间费用");
            namelist.Add("主营业务毛利率");
            namelist.Add("应收账款");
            namelist.Add("存货");
            namelist.Add("科研投入占比");
            namelist.Add("净资产");
            namelist.Add("净资产收益率");
            namelist.Add("三项费用占收入比");
            namelist.Add("全口径应收账款");

            #endregion
        }

        private static void Add_qijianfeiyong_RowName(List<string> namelist)
        {
            #region 左侧列
            namelist.Add("公司合并（抵消后）");
            namelist.Add("销售费用");
            namelist.Add("管理费用");
            namelist.Add("财务费用");
            namelist.Add("三项费用合计");
            namelist.Add("费用收入比");

            #endregion
        }
        private static void Add_maolilv_RowName(List<string> namelist)
        {
            #region 左侧列
            namelist.Add("公司合并（抵消后）");
            namelist.Add("营业收入");
            namelist.Add("营业成本（变动");
            namelist.Add("边际利润");
            namelist.Add("营业成本（固定");
            namelist.Add("毛利");
            namelist.Add("毛利率（");
            #endregion
        }
        private static void Add_cunhuo_RowName(List<string> namelist)
        {
            #region 左侧列.
            namelist.Add("公司合并（抵消后）");
            namelist.Add("存货");

            #endregion
        }

        private static void Add_Add_xianjinliuJINGE_RowName(List<string> namelist)
        {
            #region 左侧列
            namelist.Add("经营活动产生的现金流量净额");
            namelist.Add("投资活动产生的现金流量净额");
            namelist.Add("筹资活动产生的现金流量净额");
            namelist.Add("现金流量净额");

            #endregion
        }
        private static void Add_Lirunjilirunfenpei_RowName(List<string> namelist)
        {
            #region 左侧列
            namelist.Add("公司合并（抵消后）");
            namelist.Add("营业收入");
            namelist.Add("营业成本（变动");
            namelist.Add("边际利润");
            namelist.Add("营业成本（固定");
            namelist.Add("毛利");
            namelist.Add("毛利率（");
            #endregion
        }

        public void DownLoadPDF(ref BackgroundWorker bgWorker, string pathname)
        {
            bgWorker1 = bgWorker;

            if (XLSConvertToPDF(pathname, pathname.Replace("xlsx", "pdf")))
            {
                var dir = System.IO.Path.GetDirectoryName(pathname);
                string namesave = System.IO.Path.GetFileName(pathname);
                //File.Copy(pathname.Replace("xlsx", "pdf"), dir + "\\" + namesave.Replace("xlsx", "pdf"));

                //File.Delete(pathname);
                //File.Delete(pathname.Replace("xlsx", "pdf"));
            }



        }
        private bool XLSConvertToPDF(string sourcePath, string targetPath)
        {
            bool result = false;
            Microsoft.Office.Interop.Excel.XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
            object missing = Type.Missing;
            Microsoft.Office.Interop.Excel.Application ExcelApp = null;
            Microsoft.Office.Interop.Excel._Workbook ExcelBook = null;
            try
            {

                object target = targetPath;
                object type = targetType;

                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                System.Reflection.Missing missingValue = System.Reflection.Missing.Value;
                ExcelBook = ExcelApp.Workbooks.Open(sourcePath, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue);

                ExcelBook.ExportAsFixedFormat(targetType, target, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);
                result = true;


            }
            catch
            {
                result = false;
            }
            finally
            {
                if (ExcelBook != null)
                {
                    ExcelBook.Close(true, missing, missing);
                    ExcelBook = null;
                }
                if (ExcelApp != null)
                {
                    ExcelApp.Quit();
                    ExcelApp = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }


        public List<clszichanfuzaibiaoinfo> ReadEv_Datasources(ref BackgroundWorker bgWorker, string filename)
        {
            zichanfuzaibiao_Result = new List<clszichanfuzaibiaoinfo>();
            zhuyao_Result = new List<clszhuyaojingyingzhibiaowanchengqingkuanginfo>();
            Lirunjilirunfenpei_Result = new List<clsLirunjilirunfenpeibiao_info>();
            baxiangfeiyong_Result = new List<cls8xiangfeiyongzhichu_info>();
            Xianjinliu_Result = new List<clsXianjinliu_info>();
            qijianfeiyong_Result = new List<clsQijianfeiyong_info>();
            cunhuo_Result = new List<clscunhuo_info>();
            xianjinliu_Result = new List<clsxianjinliu_info>();
            maolilv_Result = new List<clsmaolilv_info>();

            string path = AppDomain.CurrentDomain.BaseDirectory + "FileList";
            List<string> Alist = GetBy_CategoryReportFileName(path);

            if (Alist.Count > 1)
                for (int i = 0; i < Alist.Count; i++)
                {
                    GetKEYnfo(path + "\\" + Alist[i]);
                }


            return zichanfuzaibiao_Result;


        }
    }
}
