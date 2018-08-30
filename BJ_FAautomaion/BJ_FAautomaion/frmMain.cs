using DCTS.CustomComponents;
using FA.Buiness;
using FA.Common;
using FA.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using WeifenLuo.WinFormsUI.Docking;

namespace BJ_FAautomaion
{
    public partial class frmMain : DockContent
    {

        // 后台执行控件
        private BackgroundWorker bgWorker;
        // 消息显示窗体
        private frmMessageShow frmMessageShow;
        // 后台操作是否正常完成
        private bool blnBackGroundWorkIsOK = false;
        //后加的后台属性显
        private bool backGroundRunResult;

        List<clszichanfuzaibiaoinfo> ClaimReport;
        List<clszichanfuzaibiaoinfo> zichanfuzaibiao_Result;

        List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> zhuyao_Result;


        //财务 利润及利润分配表
        List<clsLirunjilirunfenpeibiao_info> Lirunjilirunfenpei_Result;
        //财务 现金流量表
        List<clsXianjinliu_info> Xianjinliu_Result;

        //八项费用支出表
        List<cls8xiangfeiyongzhichu_info> baxiangfeiyong_Result;

        //期间费用情况
        List<clsQijianfeiyong_info> qijianfeiyong_Result;

        //毛利率情况
        List<clsmaolilv_info> maolilv_Result;

        //存货情况
        List<clscunhuo_info> cunhuo_Result;

        //现金流净额
        List<clsxianjinliu_info> xianjinliu_Result;



        private SortableBindingList<clszichanfuzaibiaoinfo> sortablezichanfuzaibiaoList;
        string strFileName;


        public frmMain()
        {
            InitializeComponent();
        }

        private void InitialBackGroundWorker()
        {
            bgWorker = new BackgroundWorker();
            bgWorker.WorkerReportsProgress = true;
            bgWorker.WorkerSupportsCancellation = true;
            bgWorker.RunWorkerCompleted +=
                new RunWorkerCompletedEventHandler(bgWorker_RunWorkerCompleted);
            bgWorker.ProgressChanged +=
                new ProgressChangedEventHandler(bgWorker_ProgressChanged);
        }

        private void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                blnBackGroundWorkIsOK = false;
            }
            else if (e.Cancelled)
            {
                blnBackGroundWorkIsOK = true;
            }
            else
            {
                blnBackGroundWorkIsOK = true;
            }
        }

        private void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (frmMessageShow != null && frmMessageShow.Visible == true)
            {
                //设置显示的消息
                frmMessageShow.setMessage(e.UserState.ToString());
                //设置显示的按钮文字
                if (e.ProgressPercentage == clsConstant.Thread_Progress_OK)
                {
                    frmMessageShow.setStatus(clsConstant.Dialog_Status_Enable);
                }
            }
        }

        private void 读取_Click(object sender, EventArgs e)
        {

            try
            {
                hide_label_pbStatus();
                InitialBackGroundWorker();
                bgWorker.DoWork += new DoWorkEventHandler(ReadclaimreportfromServer);

                bgWorker.RunWorkerAsync();

                // 启动消息显示画面
                frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();

                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {
                    InitializeOrderData(zichanfuzaibiao_Result, zhuyao_Result);

                }
            }
            catch (Exception ex)
            {
                return;
                throw ex;
            }
        }

        private void Show_label_pbStatus(int co, int SelectedIndex)
        {
            this.pbStatus.Visible = false;
            this.toolStripLabel1.Text = "共计  : " + co.ToString();
            this.tabControl1.SelectedIndex = SelectedIndex;
        }
        private void hide_label_pbStatus()
        {
            this.pbStatus.Visible = true;
            this.toolStripLabel1.Text = "正在运行... ";

        }
        private void ReadclaimreportfromServer(object sender, DoWorkEventArgs e)
        {
            ClaimReport = new List<clszichanfuzaibiaoinfo>();
            zichanfuzaibiao_Result = new List<clszichanfuzaibiaoinfo>();
            zhuyao_Result = new List<clszhuyaojingyingzhibiaowanchengqingkuanginfo>();
            Lirunjilirunfenpei_Result = new List<clsLirunjilirunfenpeibiao_info>();
            baxiangfeiyong_Result = new List<cls8xiangfeiyongzhichu_info>();
            Xianjinliu_Result = new List<clsXianjinliu_info>();
            qijianfeiyong_Result = new List<clsQijianfeiyong_info>();
            cunhuo_Result = new List<clscunhuo_info>();
            xianjinliu_Result = new List<clsxianjinliu_info>();
            maolilv_Result = new List<clsmaolilv_info>();
            //初始化信息
            clsAllnew BusinessHelp = new clsAllnew();
            //导入程序集
            DateTime oldDate = DateTime.Now;

            BusinessHelp.ReadDatasources(ref this.bgWorker, "");

            zhuyao_Result = BusinessHelp.zhuyao_Result;
            zichanfuzaibiao_Result = BusinessHelp.zichanfuzaibiao_Result;
            Lirunjilirunfenpei_Result = BusinessHelp.Lirunjilirunfenpei_Result;
            baxiangfeiyong_Result = BusinessHelp.baxiangfeiyong_Result;
            Xianjinliu_Result = BusinessHelp.Xianjinliu_Result;
            qijianfeiyong_Result = BusinessHelp.qijianfeiyong_Result;
            cunhuo_Result = BusinessHelp.cunhuo_Result;
            xianjinliu_Result = BusinessHelp.xianjinliu_Result;
            maolilv_Result = BusinessHelp.maolilv_Result;

            Data_maintain();

            DateTime FinishTime = DateTime.Now;
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);


        }

        private void Data_maintain()
        {

            #region 资产总额
            List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> cloumnlistSQ = zhuyao_Result.FindAll(sQ => sQ.zhibiaomingcheng != null && sQ.zhibiaomingcheng.Contains("资产总额"));
            if (cloumnlistSQ.Count != 0 && cloumnlistSQ.Count == 1)
            {

                double dd = 0;
                double f3 = 0;
                double H3 = 0;
                //资产总额--本月完成
                if (zichanfuzaibiao_Result[zichanfuzaibiao_Result.Count - 1].qimojine != "")
                    dd = Convert.ToDouble(zichanfuzaibiao_Result[zichanfuzaibiao_Result.Count - 1].qimojine) / 10000;

                cloumnlistSQ[0].benyuewancheng = dd.ToString();
                //资产总额--累计完成或期末数
                if (zichanfuzaibiao_Result[zichanfuzaibiao_Result.Count - 1].qimojineH != "")
                    f3 = Convert.ToDouble(zichanfuzaibiao_Result[zichanfuzaibiao_Result.Count - 1].qimojineH) / 10000;

                cloumnlistSQ[0].leijiwanchenghuoqimoshu = f3.ToString();

                //资产总额--上年同期累计完成
                if (zichanfuzaibiao_Result[zichanfuzaibiao_Result.Count - 1].shangniantongqishuI != "")
                    H3 = Convert.ToDouble(zichanfuzaibiao_Result[zichanfuzaibiao_Result.Count - 1].shangniantongqishuI) / 10000;

                cloumnlistSQ[0].shangniantongqileijiwancheng = H3.ToString();
                //同比增减
                double I3 = f3 - H3;

                cloumnlistSQ[0].tongbizengzhang = I3.ToString();

            }
            #endregion

            #region 负债总额
            List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> fz = zhuyao_Result.FindAll(sQ => sQ.zhibiaomingcheng != null && sQ.zhibiaomingcheng.Contains("负债总额"));
            if (fz.Count != 0 && fz.Count == 1)
            {

                double dd = 0;
                double f4 = 0;
                double H4 = 0;
                //资产总额--本月完成
                List<clszichanfuzaibiaoinfo> zcfzb = zichanfuzaibiao_Result.FindAll(sQ => sQ.xiangmuF != null && sQ.xiangmuF.Contains("负 债 合 计"));

                if (zcfzb.Count == 1 && zcfzb[0].qimojineH != "")
                    dd = Convert.ToDouble(zcfzb[0].qimojineH) / 10000;

                fz[0].benyuewancheng = dd.ToString();
                //资产总额--累计完成或期末数
                if (zcfzb.Count == 1 && zcfzb[0].qimojineH != "")
                    f4 = Convert.ToDouble(zcfzb[0].qimojineH) / 10000;

                fz[0].leijiwanchenghuoqimoshu = f4.ToString();

                //资产总额--上年同期累计完成
                if (zcfzb[0].shangniantongqishuI != "")
                    H4 = Convert.ToDouble(zcfzb[0].shangniantongqishuI) / 10000;

                fz[0].shangniantongqileijiwancheng = H4.ToString();
                //同比增减
                double I4 = f4 - H4;

                cloumnlistSQ[0].tongbizengzhang = I4.ToString();

            }

            #endregion

            #region 资产负债率
            double d5 = Convert.ToDouble(fz[0].benyuewancheng) / Convert.ToDouble(cloumnlistSQ[0].benyuewancheng);
            List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> zcfzl = zhuyao_Result.FindAll(sQ => sQ.zhibiaomingcheng != null && sQ.zhibiaomingcheng.Contains("资产负债率"));
            zcfzl[0].benyuewancheng = d5.ToString();

            double f5 = Convert.ToDouble(fz[0].leijiwanchenghuoqimoshu) / Convert.ToDouble(cloumnlistSQ[0].leijiwanchenghuoqimoshu);
            zcfzl[0].leijiwanchenghuoqimoshu = f5.ToString();

            double h5 = Convert.ToDouble(fz[0].shangniantongqileijiwancheng) / Convert.ToDouble(cloumnlistSQ[0].shangniantongqileijiwancheng);
            zcfzl[0].shangniantongqileijiwancheng = h5.ToString();

            double I5 = f5 - h5;

            zcfzl[0].tongbizengzhang = I5.ToString();
            #endregion

            #region 营业收入
            List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> yysr = zhuyao_Result.FindAll(sQ => sQ.zhibiaomingcheng != null && sQ.zhibiaomingcheng.Contains("营业收入"));
            if (yysr.Count != 0 && yysr.Count == 1)
            {
                double d6 = 0;
                double f6 = 0;
                double H6 = 0;
                List<clsLirunjilirunfenpeibiao_info> lr = Lirunjilirunfenpei_Result.FindAll(sQ => sQ.xiangmu != null && sQ.xiangmu.Contains("一、营业总收入"));
                if (lr.Count == 1 && lr[0].benyueshu != "")
                    d6 = Convert.ToDouble(lr[0].benyueshu) / 10000;
                yysr[0].benyuewancheng = d6.ToString();


                if (lr.Count == 1 && lr[0].bennianleijishu != "")
                    f6 = Convert.ToDouble(lr[0].bennianleijishu) / 10000;
                yysr[0].leijiwanchenghuoqimoshu = f6.ToString();
                if (lr.Count == 1 && lr[0].shangniantongqi != "")
                    H6 = Convert.ToDouble(lr[0].shangniantongqi) / 10000;
                yysr[0].shangniantongqileijiwancheng = H6.ToString();

                double I6 = f6 - H6;
                yysr[0].tongbizengzhang = I6.ToString();
            }


            #endregion

            #region 利润总额
            List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> lrze = zhuyao_Result.FindAll(sQ => sQ.zhibiaomingcheng != null && sQ.zhibiaomingcheng.Contains("利润总额"));
            if (lrze.Count != 0 && lrze.Count == 1)
            {
                double d7 = 0;
                double f7 = 0;
                double H7 = 0;
                List<clsLirunjilirunfenpeibiao_info> lr = Lirunjilirunfenpei_Result.FindAll(sQ => sQ.xiangmu != null && sQ.xiangmu.Contains("润总额（亏损总额以“－”号填列"));
                if (lr.Count == 1 && lr[0].benyueshu != "")
                    d7 = Convert.ToDouble(lr[0].benyueshu) / 10000;
                lrze[0].benyuewancheng = d7.ToString();


                if (lr.Count == 1 && lr[0].bennianleijishu != "")
                    f7 = Convert.ToDouble(lr[0].bennianleijishu) / 10000;
                lrze[0].leijiwanchenghuoqimoshu = f7.ToString();
                if (lr.Count == 1 && lr[0].shangniantongqi != "")
                    H7 = Convert.ToDouble(lr[0].shangniantongqi) / 10000;
                lrze[0].shangniantongqileijiwancheng = H7.ToString();

                double I7 = f7 - H7;
                lrze[0].tongbizengzhang = I7.ToString();
            }

            #endregion
            #region 期间费用
            List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> qjfy = zhuyao_Result.FindAll(sQ => sQ.zhibiaomingcheng != null && sQ.zhibiaomingcheng.Contains("期间费用"));
            if (qjfy.Count != 0 && qjfy.Count == 1)
            {
                double d8 = 0;
                double f8 = 0;
                double H8 = 0;
                List<clsLirunjilirunfenpeibiao_info> xsfy = Lirunjilirunfenpei_Result.FindAll(sQ => sQ.xiangmu != null && sQ.xiangmu.Contains("销售费用"));
                List<clsLirunjilirunfenpeibiao_info> glfy = Lirunjilirunfenpei_Result.FindAll(sQ => sQ.xiangmu != null && sQ.xiangmu.Contains("管理费用"));
                List<clsLirunjilirunfenpeibiao_info> cwfy = Lirunjilirunfenpei_Result.FindAll(sQ => sQ.xiangmu != null && sQ.xiangmu.Contains("财务费用"));
                double c24 = 0;
                double c25 = 0;
                double c27 = 0;

                if (xsfy.Count == 1 && xsfy[0].benyueshu != "")
                    c24 = Convert.ToDouble(xsfy[0].benyueshu);
                if (glfy.Count == 1 && glfy[0].benyueshu != "")
                    c25 = Convert.ToDouble(glfy[0].benyueshu);
                if (cwfy.Count == 1 && cwfy[0].benyueshu != "")
                    c27 = Convert.ToDouble(cwfy[0].benyueshu);

                double total = c24 / 10000 + c25 / 10000 + c27 / 10000;

                d8 = total / 10000;
                qjfy[0].benyuewancheng = d8.ToString();
                //f8
                double d24 = 0;
                double d25 = 0;
                double d27 = 0;

                if (xsfy.Count == 1 && xsfy[0].bennianleijishu != "")
                    d24 = Convert.ToDouble(xsfy[0].bennianleijishu);
                if (glfy.Count == 1 && glfy[0].bennianleijishu != "")
                    d25 = Convert.ToDouble(glfy[0].bennianleijishu);
                if (cwfy.Count == 1 && cwfy[0].bennianleijishu != "")
                    d27 = Convert.ToDouble(cwfy[0].bennianleijishu);

                total = d24 + d25 + d27;
                f8 = total / 10000;
                qjfy[0].leijiwanchenghuoqimoshu = f8.ToString();

                //h8

                double e24 = 0;
                double e25 = 0;
                double e27 = 0;

                if (xsfy.Count == 1 && xsfy[0].shangniantongqishu != "")
                    e24 = Convert.ToDouble(xsfy[0].shangniantongqishu);
                if (glfy.Count == 1 && glfy[0].shangniantongqishu != "")
                    e25 = Convert.ToDouble(glfy[0].shangniantongqishu);
                if (cwfy.Count == 1 && cwfy[0].shangniantongqishu != "")
                    e27 = Convert.ToDouble(cwfy[0].shangniantongqishu);

                total = e24 + e25 + e27;
                H8 = total / 10000;
                qjfy[0].shangniantongqileijiwancheng = H8.ToString();

                double I8 = f8 - H8;
                lrze[0].tongbizengzhang = I8.ToString();
            }
            #endregion

            #region 主营业务毛利率

            List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> zyywmll = zhuyao_Result.FindAll(sQ => sQ.zhibiaomingcheng != null && sQ.zhibiaomingcheng.Contains("主营业务毛利率"));
            if (zyywmll.Count != 0 && zyywmll.Count == 1)
            {
                double d9 = 0;
                double f9 = 0;
                double H9 = 0;
                List<clsLirunjilirunfenpeibiao_info> c39 = Lirunjilirunfenpei_Result.FindAll(sQ => sQ.xiangmu != null && sQ.xiangmu.Contains("营业利润（亏损以“－”号填列"));
                List<clsLirunjilirunfenpeibiao_info> c5 = Lirunjilirunfenpei_Result.FindAll(sQ => sQ.xiangmu != null && sQ.xiangmu.Contains("一、营业总收入"));

                double c39n = 0;
                double c5n = 0;
                double c27 = 0;

                if (c39.Count == 1 && c39[0].benyueshu != "")
                    c39n = Convert.ToDouble(c39[0].benyueshu);
                if (c5.Count == 1 && c5[0].benyueshu != "")
                    c5n = Convert.ToDouble(c5[0].benyueshu);


                double total = (c39n / 10000) / (c5n / 10000);

                d9 = total;
                zyywmll[0].benyuewancheng = d9.ToString();
                //f9
                double d39 = 0;
                double d5n = 0;


                if (c39.Count == 1 && c39[0].bennianleijishu != "")
                    d39 = Convert.ToDouble(c39[0].bennianleijishu);
                if (c5.Count == 1 && c5[0].bennianleijishu != "")
                    d5n = Convert.ToDouble(c5[0].bennianleijishu);


                total = (d39 / 10000) / (d5n / 10000);
                f9 = total;
                zyywmll[0].leijiwanchenghuoqimoshu = f9.ToString();

                //h9

                double e39 = 0;
                double e4n = 0;


                if (c39.Count == 1 && c39[0].shangniantongqishu != "")
                    e39 = Convert.ToDouble(c39[0].shangniantongqishu);
                if (c5.Count == 1 && c5[0].shangniantongqishu != "")
                    e4n = Convert.ToDouble(c5[0].shangniantongqishu);


                total = (e39 / 10000) / (e4n / 10000);
                H9 = total;
                zyywmll[0].leijiwanchenghuoqimoshu = H9.ToString();


                double I9 = f9 - H9;
                zyywmll[0].tongbizengzhang = I9.ToString();
            }



            #endregion


            #region 应收账款
            List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> yszk = zhuyao_Result.FindAll(sQ => sQ.zhibiaomingcheng != null && sQ.zhibiaomingcheng.Contains("应收账款"));
            if (yszk.Count != 0 && yszk.Count == 1)
            {
                double d10 = 0;
                double f10 = 0;
                double H10 = 0;
                //资产总额--本月完成
                List<clszichanfuzaibiaoinfo> lr = zichanfuzaibiao_Result.FindAll(sQ => sQ.xiangmu != null && sQ.xiangmu=="应收账款");
                if (lr.Count == 1 && lr[0].qimojine != "")
                {
                    d10 = Convert.ToDouble(lr[0].qimojine) / 10000;

                    H10 = Convert.ToDouble(lr[0].shangniantongqishu) / 10000;                    
                }
                yszk[0].benyuewancheng = d10.ToString();

                yszk[0].leijiwanchenghuoqimoshu = d10.ToString();

                yszk[0].shangniantongqileijiwancheng = H10.ToString();

                //同比增减
                double I10 = d10 - H10;

                yszk[0].tongbizengzhang = I10.ToString();                                
            }

            #endregion
            #region 存货
            List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> ch = zhuyao_Result.FindAll(sQ => sQ.zhibiaomingcheng != null && sQ.zhibiaomingcheng.Contains("存货"));
            if (ch.Count != 0 && ch.Count == 1)
            {
                double d10 = 0;
                double f10 = 0;
                double H10 = 0;
                //资产总额--本月完成
                List<clszichanfuzaibiaoinfo> lr = zichanfuzaibiao_Result.FindAll(sQ => sQ.xiangmu != null && sQ.xiangmu == "应收账款");
                if (lr.Count == 1 && lr[0].qimojine != "")
                {
                    d10 = Convert.ToDouble(lr[0].qimojine) / 10000;

                    H10 = Convert.ToDouble(lr[0].shangniantongqishu) / 10000;
                }
                yszk[0].benyuewancheng = d10.ToString();

                yszk[0].leijiwanchenghuoqimoshu = d10.ToString();

                yszk[0].shangniantongqileijiwancheng = H10.ToString();

                //同比增减
                double I10 = d10 - H10;

                yszk[0].tongbizengzhang = I10.ToString();
            }

            #endregion

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int s = this.tabControl1.SelectedIndex;
            if (s == 0)
            {

                Show_label_pbStatus(dataGridView.RowCount, 0);
            }
            if (s == 1)
            {

                Show_label_pbStatus(dataGridView1.RowCount, 1);
            }
        }

        private void InitializeOrderData(List<clszichanfuzaibiaoinfo> zichanfuzaibiao_Result, List<clszhuyaojingyingzhibiaowanchengqingkuanginfo> zhuyao_Result)
        {
            Data_maintain();

            this.dataGridView.DataSource = null;
            this.dataGridView.AutoGenerateColumns = false;

            if (zichanfuzaibiao_Result.Count != 0)
            {
                sortablezichanfuzaibiaoList = new SortableBindingList<clszichanfuzaibiaoinfo>(zichanfuzaibiao_Result);
                this.bindingSource1.DataSource = this.sortablezichanfuzaibiaoList;


                this.dataGridView.DataSource = this.bindingSource1;
                Show_label_pbStatus(zichanfuzaibiao_Result.Count, 0);

                List<clszichanfuzaibiaoinfo> zcfzb = zichanfuzaibiao_Result.FindAll(sQ => sQ.xiangmu != null && sQ.xiangmu.Contains("资  产  总  计"));

                label21.Text = zcfzb[0].qimojine;
                List<clszichanfuzaibiaoinfo> zcfzb1 = zichanfuzaibiao_Result.FindAll(sQ => sQ.xiangmuF != null && sQ.xiangmuF.Contains("负债和所有者权益总计"));

                label22.Text = zcfzb1[0].qimojineH;
            }

            this.dataGridView1.DataSource = null;
            this.dataGridView1.AutoGenerateColumns = false;

            if (zhuyao_Result.Count != 0)
            {

                this.dataGridView1.DataSource = zhuyao_Result;
                //Show_label_pbStatus(zhuyao_Result.Count, 0);
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            InitializeOrderData(zichanfuzaibiao_Result, zhuyao_Result);

        }


        private void downcsv(DataGridView dataGridView)
        {

            if (dataGridView.Rows.Count == 0)
            {
                MessageBox.Show("Sorry , No Data Output !", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".csv";
            saveFileDialog.Filter = "csv|*.csv";
            string strFileName = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            saveFileDialog.FileName = strFileName;
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                strFileName = saveFileDialog.FileName.ToString();
            }
            else
            {
                return;
            }
            FileStream fa = new FileStream(strFileName, FileMode.Create);
            StreamWriter sw = new StreamWriter(fa, Encoding.Unicode);
            string delimiter = "\t";
            string strHeader = "";
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                strHeader += dataGridView.Columns[i].HeaderText + delimiter;
            }
            sw.WriteLine(strHeader);

            //output rows data
            for (int j = 0; j < dataGridView.Rows.Count; j++)
            {
                string strRowValue = "";

                for (int k = 0; k < dataGridView.Columns.Count; k++)
                {
                    if (dataGridView.Rows[j].Cells[k].Value != null)
                    {
                        strRowValue += dataGridView.Rows[j].Cells[k].Value.ToString().Replace("\r\n", " ").Replace("\n", "") + delimiter;
                        if (dataGridView.Rows[j].Cells[k].Value.ToString() == "LIP201507-35")
                        {

                        }

                    }
                    else
                    {
                        strRowValue += dataGridView.Rows[j].Cells[k].Value + delimiter;
                    }
                }
                sw.WriteLine(strRowValue);
            }
            sw.Close();
            fa.Close();
            MessageBox.Show("下载完成 ！", "System", MessageBoxButtons.OK, MessageBoxIcon.Information);


        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            int s = this.tabControl1.SelectedIndex;
            if (s == 0)
            {
                downcsv(dataGridView);
            }
            else if (s == 1)
            {
                downcsv(dataGridView1);
            }
        }

        private void filterButton_Click(object sender, EventArgs e)
        {

            //List<clszichanfuzaibiaoinfo> zcfzb = zichanfuzaibiao_Result.FindAll(sQ => ((sQ.xiangmu != null && sQ.xiangmu.Contains(textBox6.Text)) || (sQ.xiangmuF != null && sQ.xiangmuF.Contains(textBox6.Text))) && Convert.ToDateTime(sQ.riqi) > Convert.ToDateTime(stockOutDateTimePicker.Text) && Convert.ToDateTime(sQ.riqi) < Convert.ToDateTime(stockInDateTimePicker1.Text));
            List<clszichanfuzaibiaoinfo> zcfzb = zichanfuzaibiao_Result.FindAll(sQ => (sQ.xiangmu != null && sQ.xiangmu.Contains(textBox6.Text)) || (sQ.xiangmuF != null && sQ.xiangmuF.Contains(textBox6.Text)));
            List<clszichanfuzaibiaoinfo> zcfzb3 = zcfzb.FindAll(sQ => Convert.ToDateTime(sQ.riqi) > Convert.ToDateTime(stockOutDateTimePicker.Text) && Convert.ToDateTime(sQ.riqi) < Convert.ToDateTime(stockInDateTimePicker1.Text));

            InitializeOrderData(zcfzb3, zhuyao_Result);

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".xlsx";
            saveFileDialog.Filter = "Excel Files(*.xls,*.xlsx,*.xlsm,*.xlsb)|*.xls;*.xlsx;*.xlsm;*.xlsb";
            strFileName = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            saveFileDialog.FileName = strFileName;
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                strFileName = saveFileDialog.FileName.ToString();
            }
            else
            {
                return;
            }
            try
            {
                InitialBackGroundWorker();
                bgWorker.DoWork += new DoWorkEventHandler(downreport);
                bgWorker.RunWorkerAsync();
                // 启动消息显示画面
                frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();
                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {
                    //string ZFCEPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results"), "");
                    //System.Diagnostics.Process.Start("explorer.exe", strFileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ex" + ex);
                return;
                throw ex;
            }
        }
        private void downreport(object sender, DoWorkEventArgs e)
        {
            DateTime oldDate = DateTime.Now;

            //初始化信息
            clsAllnew BusinessHelp = new clsAllnew();

            BusinessHelp.InitializeDataSource(zichanfuzaibiao_Result, zhuyao_Result);

            BusinessHelp.pbStatus = pbStatus;
            BusinessHelp.tsStatusLabel1 = toolStripLabel1;
            BusinessHelp.DownLoadExcel(ref this.bgWorker, strFileName);

            //暂停
            BusinessHelp.DownLoadPDF(ref this.bgWorker, strFileName);

            DateTime FinishTime = DateTime.Now;
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_015 + "\r\n" + Showtime);


        }


    }
}
