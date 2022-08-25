using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelApp.Model
{
    class ExcelGenerator
    {
        ExcelModel _excelModel;

        public ExcelGenerator(ExcelModel excelModel)
        {
            _excelModel = excelModel;
        }

        public void Generate()
        {
            Excel.Application application = new Excel.Application();
            Excel.Workbook workbook = application.Workbooks.Add();
            Excel.Worksheet worksheet = application.Worksheets[1];
            try
            {
                worksheet.Name = _excelModel.ListName;
                worksheet.Cells[1, 1] = _excelModel.TableName;
                worksheet.Columns.AutoFit();
                Random random = new Random();
                var min = _excelModel.RandomMin;
                var max = _excelModel.RandomMax;
                for (int i = 2; i <_excelModel.CellsCount+1; i++)
                {
                    worksheet.Cells[i, 1] = random.Next(min, max);
                }

                workbook.SaveAs("c:\\1\\1.xlsx");
                workbook.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                workbook.Close();
            }
        }
    }
}
