using ClosedXML.Excel;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace My.Activities.TipeCelByPattern
{
    public class TypeCellByPattern : CodeActivity
    {

        private IXLWorkbook book;
        private IXLWorksheet worksheet;

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> PathExcel { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> SheetName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> Cell { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public string Type { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public XLDataType DataType { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            string path = PathExcel.Get(context);
            string sheetName = SheetName.Get(context);
            string cell = Cell.Get(context);
            SetTip(path, sheetName, cell, Type, DataType);
        }

        public void SetTip(string path, string sheetName, string cell, string format, XLDataType dataType)
        {
            if (File.Exists(path))
            {
                book = new XLWorkbook(path);

                try
                {
                    worksheet = book.Worksheet(sheetName);
                }
                catch
                {
                    worksheet = book.AddWorksheet(sheetName);
                }
            }
            else
            {
                book = new XLWorkbook();
                worksheet = book.AddWorksheet(sheetName);
            }

            if (!cell.Contains(":"))
                TipCell(cell, format, dataType);
            else
                TipRange(cell, format, dataType);
            book.Save();
            //wb.SaveAs(filePath);
        }

        private void TipCell(string target, string format,XLDataType dataType)
        {
            worksheet.Cell(target).SetDataType(dataType);
            worksheet.Cell(target).Style.NumberFormat.Format = format;
        }
        private void TipRange(string target, string format, XLDataType dataType)
        {
            string[] range = target.Split(':');

            IXLRange rangeXL;
            if (string.IsNullOrWhiteSpace(range[1]))
            {
                rangeXL = worksheet.Range(range[0].ToUpper(), GetAlfb(worksheet.RangeUsed().FirstRowUsed().CellCount() - 1) + (worksheet.RangeUsed().RowCount()));
            }
            else
            {
                rangeXL = worksheet.Range(range[0].ToUpper(), range[1].ToUpper());
            }
            worksheet.Cell(target).SetDataType(dataType);
            rangeXL.Style.NumberFormat.Format = format;
        }
        private string GetAlfb(int num)
        {
            return (065 + num) > 90 ? ((char)Math.Floor(64 + (64.0 + num) / 90)).ToString() + ((char)(num % 90)).ToString() : ((char)(065 + num)).ToString();
        }
    }
}
