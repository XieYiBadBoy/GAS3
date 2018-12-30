using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 天然气市场需求分析软件_求你不死机版_
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }
        public Windows7 EnergyConsumptionAnalys;
        public Windows1 CityGas;
        public Windows3 Chemical;
        public Windows5 Traffic;
        public Windows10 SmartAnalysCityGas;
        public Windows6 SavingEnergy;
        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 关闭ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 关闭所有ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 交通ToolStripMenuItem_Click(object sender, EventArgs e)
        {  
            Traffic = new Windows5();
            Traffic.MdiParent = this;
            Traffic.Show();
        }

        private void 能耗系数ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EnergyConsumptionAnalys = new Windows7();
            EnergyConsumptionAnalys.MdiParent = this;
            EnergyConsumptionAnalys.Show();
        }

        private void 复制ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SendKeys.Send("^{C}");

        }

        private void 城市燃气ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CityGas = new Windows1();
            CityGas.MdiParent = this;
            CityGas.Show();
        }



        private void 压气站布置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Chemical = new Windows3();
            Chemical.MdiParent = this;
            Chemical.Show();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.toolStripStatusLabel3.Text = "系统当前时间：" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            this.toolStripStatusLabel3.Text = "系统当前时间：" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");

            this.timer1.Interval = 1000;

            this.timer1.Start();
        }

        private void 智能分析城市燃气ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SmartAnalysCityGas = new Windows10();
            SmartAnalysCityGas.MdiParent = this;
            SmartAnalysCityGas.Show();
        }

        private void 节能减排ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SavingEnergy = new Windows6();
            SavingEnergy.MdiParent = this;
            SavingEnergy.Show();
        }

        private void 状态栏ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem12.Checked = !ToolStripMenuItem12.Checked;
            statusStrip1.Visible = !statusStrip1.Visible;
        }

        private void ToolStripMenuItem22_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem22.Checked = !ToolStripMenuItem22.Checked;
            statusStrip1.Visible = !statusStrip1.Visible;
        }
    }
}
