using System;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Diagnostics;

namespace 天然气市场需求分析软件_求你不死机版_
{
    public partial class Windows7 : Form
    {
        public Windows7()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "天然气需求预测能耗指数法.xlsx";
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
            sheet1.GetRow(6).GetCell(2).SetCellValue(textBox1.Text);
            sheet1.GetRow(7).GetCell(2).SetCellValue(textBox2.Text);


            for (int i = 11; i < 15; i++)
            {
                IRow row = sheet1.GetRow(i);
                for (int j = 4; j < 12; j++)
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


            for (int i = 22; i < 26; i++)
            {
                IRow row = sheet1.GetRow(i);
                for (int j = 2; j < 13; j++)
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

             FileStream file = new FileStream("C:\\Users\\Administrator\\Desktop\\需求预测-能耗系数法.xlsx", FileMode.OpenOrCreate);
            workbook1.Write(file);
            file.Close();
            workbook1.Close();
        }
    }
}
