using System;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Diagnostics;

namespace 天然气市场需求分析软件_求你不死机版_
{
    public partial class Windows10 : Form
    {
        public Windows10()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "城市燃气智能化分析.xlsx";
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

            sheet1.GetRow(3).GetCell(1).SetCellValue(comboBox1.Text);
            sheet1.GetRow(3).GetCell(3).SetCellValue(textBox1.Text);
            sheet1.GetRow(3).GetCell(4).SetCellValue(textBox2.Text);
            sheet1.GetRow(3).GetCell(5).SetCellValue(textBox3.Text);
            if (checkBox1.Checked == true)
            {
                sheet1.GetRow(3).GetCell(2).SetCellValue(1);
            }
            else
            {
                sheet1.GetRow(3).GetCell(2).SetCellValue(0);
            }

            IRow r = sheet1.GetRow(3);
            ICell c = r.GetCell(6);
            if (c.CellType == CellType.Formula)
            {
                IFormulaEvaluator m = null;
                if (fileProcess.IndexOf(".xlsx") > 0) // 2007版本
                {
                    m = new XSSFFormulaEvaluator(c.Sheet.Workbook);
                }
                else if (fileProcess.IndexOf(".xls") > 0) // 2003版本
                {
                    m = new HSSFFormulaEvaluator(c.Sheet.Workbook);
                }
                m.EvaluateInCell(c);
            }

            for (int i = 6; i < 10; i++)
            {
                IRow row = sheet1.GetRow(i);
                for (int j = 2; j < 16; j++)
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

            FileStream file = new FileStream("C:\\Users\\Administrator\\Desktop\\需求预测-城市燃气1111.xlsx", FileMode.OpenOrCreate);
            workbook1.Write(file);
            file.Close();
            workbook1.Close();
        }
    }
}
