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
    public partial class Windows6 : Form
    {
        public Windows6()
        {
            InitializeComponent();
        }
        private void Add1()
        {
            try
            {
                double Input1 = Convert.ToDouble(txtInput1.Text);  //输入量
                txtOutput1.Text  = (Input1 * 20.2).ToString("0.0000");
                txtOutput2.Text = (Input1 * 7.6).ToString("0.0000");
                txtOutput3.Text = (Input1 *37.3).ToString("0.0000");
                txtOutput4.Text = (Input1 * 0.2561).ToString("0.0000");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void txtInput1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                Add1();
            }
            else
            {
                MessageBox.Show("请选择计算类型");
            }
        }
        private void Clear()
        {
            txtInput1.Text = "";
            txtOutput1.Text = "";
            txtOutput2.Text = "";
            txtOutput3.Text = "";
            txtOutput4.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
