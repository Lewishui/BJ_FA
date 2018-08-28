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

        public List<clszichanfuzaibiaoinfo> zichanfuzaibiao_Result;

        public List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> zhuyao_Result;

        public List<clszichanfuzaibiaoinfo> ReadDatasources(ref BackgroundWorker bgWorker, string filename)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "Resources";
            List<string> Alist = GetBy_CategoryReportFileName(path);
            for (int i = 0; i < Alist.Count; i++)
            {
                GetKEYnfo(path + "\\" + Alist[i]);
            }


            return zichanfuzaibiao_Result;


        }
        //获取文件路径方法‘
        private List<string> GetBy_CategoryReportFileName(string dirPath)
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

            zichanfuzaibiao_Result = new List<clszichanfuzaibiaoinfo>();
            zhuyao_Result = new List<clszhuyaojingyingzhibiaowanchengqingkuanginfo>();

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
                        if (o[3, 6] != null)
                            temp.riqi = o[3, 6].ToString().Trim();

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




    }
}
