using FA.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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


        public List<clszichanfuzaibiaoinfo> ReadDatasources(ref BackgroundWorker bgWorker)
        {




            return null;


        }

        //读取关键字
        public List<clszichanfuzaibiaoinfo> GetKEYnfo(string Alist)
        {

            List<clszichanfuzaibiaoinfo> MAPPINGResult = new List<clszichanfuzaibiaoinfo>();
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

                    for (int i = 2; i <= rowCount; i++)
                    {
                        clszichanfuzaibiaoinfo temp = new clszichanfuzaibiaoinfo();

                        #region 基础信息

                        temp.xiangmu = "";
                        if (o[i, 1] != null)
                            temp.xiangmu = o[i, 1].ToString().Trim();

                        temp.hangci = "";
                        if (o[i, 2] != null)
                            temp.hangci = o[i, 2].ToString().Trim();


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

                        temp.Input_Date = DateTime.Now.ToString("yyyy/MM/dd");

                        #endregion
                        MAPPINGResult.Add(temp);
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: 01032" + ex);
                return null;

                throw;
            }
            return MAPPINGResult;

        }




    }
}
