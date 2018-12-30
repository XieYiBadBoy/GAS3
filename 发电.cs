using System;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace 天然气市场需求分析软件_求你不死机版_
{
    public partial class Windows3 : Form
    {

        public Windows3()
        {
            InitializeComponent();
        }
        public int MiddleVar1=3;
        public double MiddleVar2=390;
        public double MiddleVar3=52;
        public double MiddleVar4=15100;
        public double MiddleVar5=17100;
        public double MiddleVar6=89;
       
        private void ChangeLabel(int i)
        {
            if (i == 1)
            {
                label2.Text = "机组台数：";
                label5.Text = "单台容量";
                label7.Text = "机组效率";
                
                label3.Text = "台";
                label4.Text = "MW";
                label6.Text = "%";
            }
            if (i == 2)
            {
                label2.Text = "年发电量：";
                label5.Text = "年供热量";
                label7.Text = "年制冷量";

                label3.Text = "MWh";
                label4.Text = "GJ";
                label6.Text = "GJ";
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                ChangeLabel(1);
                button1.Enabled = true;
                button2.Enabled = false;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                ChangeLabel(2);
                button1.Enabled = false;
                button2.Enabled = true;
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked == true)
            {
                ChangeLabel(2);
                button1.Enabled = false;
                button2.Enabled = true;
            }
        }


        private void  InstalledCapacity(int a,double b,double c, double x, double y, double z)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "工作簿166667.xlsx";
            FileStream fileProcessStream1 = new FileStream(fileProcess, FileMode.Open, FileAccess.Read);
            if (fileProcess.IndexOf(".xlsx") > 0) // 2007版本
            {
                workbook1 = new XSSFWorkbook(fileProcessStream1);  //xlsx数据读入workbook
                fileProcessStream1.Close();
            }
            else if (fileProcess.IndexOf(".xls") > 0) // 2003版本
            {
                workbook1 = new HSSFWorkbook(fileProcessStream1);  //xls数据读入workbook
                fileProcessStream1.Close();
            }

            ISheet sheet1 = workbook1.GetSheetAt(0);
            sheet1.GetRow(9).GetCell(4).SetCellValue(a);
            sheet1.GetRow(10).GetCell(4).SetCellValue(b);
            sheet1.GetRow(11).GetCell(4).SetCellValue(c);
            sheet1.GetRow(9).GetCell(7).SetCellValue(x);
            sheet1.GetRow(10).GetCell(7).SetCellValue(y);
            sheet1.GetRow(11).GetCell(7).SetCellValue(z);



            for (int i = 18; i < 23; i++)
            {
                IRow row = sheet1.GetRow(i);
                for (int j = 3; j < 10; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (cell.CellType == CellType.Formula)
                    {
                        IFormulaEvaluator m = null;
                        if (fileProcess.IndexOf(".xlsx") > 0) // 2007版本
                        {
                            m = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                        }
                        else if (fileProcess.IndexOf(".xls") > 0) // 2003版本
                        {
                            m = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        }
                        m.EvaluateInCell(cell);
                    }
                }
            }

            FileStream file = new FileStream("C:\\Users\\Administrator\\Desktop\\需求预测-发电.xlsx", FileMode.OpenOrCreate);
            workbook1.Write(file);
            file.Close();
            workbook1.Close();
        }

        private void ProductOutput(double x,double y,double z)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "工作簿166667.xlsx";
            FileStream fileProcessStream1 = new FileStream(fileProcess, FileMode.Open, FileAccess.Read);
            if (fileProcess.IndexOf(".xlsx") > 0) // 2007版本
            {
                workbook1 = new XSSFWorkbook(fileProcessStream1);  //xlsx数据读入workbook
                fileProcessStream1.Close();
            }
            else if (fileProcess.IndexOf(".xls") > 0) // 2003版本
            {
                workbook1 = new HSSFWorkbook(fileProcessStream1);  //xls数据读入workbook
                fileProcessStream1.Close();
            }

            ISheet sheet1 = workbook1.GetSheetAt(0);
            sheet1.GetRow(9).GetCell(7).SetCellValue(x);
            sheet1.GetRow(10).GetCell(7).SetCellValue(y);
            sheet1.GetRow(11).GetCell(7).SetCellValue(z);



            for (int i = 18; i < 23; i++)
            {
                IRow row = sheet1.GetRow(i);
                for (int j = 3; j < 10; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (cell.CellType == CellType.Formula)
                    {
                        IFormulaEvaluator m = null;
                        if (fileProcess.IndexOf(".xlsx") > 0) // 2007版本
                        {
                            m = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                        }
                        else if (fileProcess.IndexOf(".xls") > 0) // 2003版本
                        {
                            m = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        }
                        m.EvaluateInCell(cell);
                    }
                }
            }

            FileStream file = new FileStream("C:\\Users\\Administrator\\Desktop\\需求预测-发电.xlsx", FileMode.OpenOrCreate);
            workbook1.Write(file);
            file.Close();
            workbook1.Close();

        }
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked == true)
            {
                ChangeLabel(1);
                button1.Enabled = true;
                button2.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            MiddleVar1 = Convert.ToInt32(textBox1.Text);
            MiddleVar2 = Convert.ToDouble(textBox2.Text);
            MiddleVar3 = Convert.ToDouble(textBox3.Text);

            InstalledCapacity(MiddleVar1,MiddleVar2,MiddleVar3, MiddleVar4, MiddleVar5, MiddleVar6);
            //ProductOutput(MiddleVar4,MiddleVar5,MiddleVar6);
          
        }

        private void button2_Click(object sender, EventArgs e)
        {
            


            MiddleVar4 = Convert.ToDouble(textBox1.Text);
            MiddleVar5 = Convert.ToDouble(textBox2.Text);
            MiddleVar6 = Convert.ToDouble(textBox3.Text);

            InstalledCapacity(MiddleVar1, MiddleVar2, MiddleVar3, MiddleVar4, MiddleVar5, MiddleVar6);
            //ProductOutput(MiddleVar4, MiddleVar5, MiddleVar6);



        }
    }
}
