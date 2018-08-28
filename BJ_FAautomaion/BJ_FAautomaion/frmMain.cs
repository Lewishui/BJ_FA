using FA.Buiness;
using FA.Common;
using FA.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
                    this.dataGridView.DataSource = null;
                    this.dataGridView.AutoGenerateColumns = false;

                    if (zichanfuzaibiao_Result.Count != 0)
                    {
                        this.dataGridView.DataSource = zichanfuzaibiao_Result;
                        Show_label_pbStatus(zichanfuzaibiao_Result.Count, 0);

                    }

                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.AutoGenerateColumns = false;

                    if (zhuyao_Result.Count != 0)
                    {

                        this.dataGridView1.DataSource = zhuyao_Result;
                        //Show_label_pbStatus(zhuyao_Result.Count, 0);

                    }
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

            //初始化信息
            clsAllnew BusinessHelp = new clsAllnew();
            //导入程序集
            DateTime oldDate = DateTime.Now;

            BusinessHelp.ReadDatasources(ref this.bgWorker, "");

            zhuyao_Result = BusinessHelp.zhuyao_Result;
            zichanfuzaibiao_Result = BusinessHelp.zichanfuzaibiao_Result;

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
                double H4= 0;
                //资产总额--本月完成
                List<clszichanfuzaibiaoinfo> zcfzb = zichanfuzaibiao_Result.FindAll(sQ => sQ.xiangmuF != null && sQ.xiangmuF.Contains("负 债 合 计"));

                if (zcfzb.Count==1&&zcfzb[0].qimojineH != "")
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

            DateTime FinishTime = DateTime.Now;
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);


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
    }
}
