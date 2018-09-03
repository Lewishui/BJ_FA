using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace BJ_FAautomaion
{
    public partial class frmLogin : Form
    {
        public log4net.ILog ProcessLogger;
        public log4net.ILog ExceptionLogger;
        private TextBox txtSAPPassword;
        private CheckBox chkSaveInfo;
        Sunisoft.IrisSkin.SkinEngine se = null;
        frmAboutBox aboutbox;
        private frmMain frmMain;
        //存放要显示的信息
        List<string> messages;
        //要显示信息的下标索引
        int index = 0;

        string usrlin = "";

        public frmLogin()
        {
            InitializeComponent();
            aboutbox = new frmAboutBox();
            InitialPassword();

            this.txtSAPUserId.Text = "admin";
            this.txtSAPPassword.Text = "000000";

        }

        private void 关于系统ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            aboutbox.ShowDialog();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            if (frmMain == null)
            {
                frmMain = new frmMain(usrlin);
                frmMain.FormClosed += new FormClosedEventHandler(FrmOMS_FormClosed);
            }
            if (frmMain == null)
            {
                frmMain = new frmMain(usrlin);
            }
            frmMain.Show(this.dockPanel2);
        }
        void FrmOMS_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (sender is frmMain)
            {
                frmMain = null;
            }
        }

        private void btmain_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            if (frmMain == null)
            {
                frmMain = new frmMain(usrlin);
                frmMain.FormClosed += new FormClosedEventHandler(FrmOMS_FormClosed);
            }
            if (frmMain == null)
            {
                frmMain = new frmMain(usrlin);
            }
            frmMain.Show(this.dockPanel2);
        }
        private void InitialPassword()
        {
            try
            {
                txtSAPPassword = new TextBox();
                txtSAPPassword.PasswordChar = '*';
                ToolStripControlHost t = new ToolStripControlHost(txtSAPPassword);
                t.Width = 100;
                t.AutoSize = false;
                t.Alignment = ToolStripItemAlignment.Right;
                this.toolStrip1.Items.Insert(this.toolStrip1.Items.Count - 4, t);

            }
            catch (Exception ex)
            {
                //clsLogPrint.WriteLog("<frmMain> InitialPassword:" + ex.Message);
                throw ex;
            }
        }
        private void tsbLogin_Click(object sender, EventArgs e)
        {
            if (this.txtSAPUserId.Text == "admin" || this.txtSAPUserId.Text == "user")
            {
                if (this.txtSAPPassword.Text == "000000")
                {

                    toolStripDropDownButton1.Enabled = true;
                    toolStripDropDownButton2.Enabled = true;

                    usrlin = "admin";
                }
                else if (this.txtSAPUserId.Text == "user" && this.txtSAPPassword.Text == "123")
                {
                    usrlin = "user";
                    toolStripDropDownButton1.Enabled = false;
                    toolStripDropDownButton2.Enabled = true;
                }
                MessageBox.Show("登录成功！");
            }
            else
            {
                toolStripDropDownButton2.Enabled = false;
                toolStripDropDownButton1.Enabled = false;
            }
        }
    }
}
