using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace CvsReader
{
    class ExcelHelper
    {
        string _path { get; set; }
        Application _excel = new Application();
        Workbook _wb;
        Worksheet _ws;

        public ExcelHelper(string path, int sheetNum)
        {
            _path = path;
            _wb = _excel.Workbooks.Open(path);
            _ws = _wb.Worksheets[sheetNum];
        }

        public string ReadCell(int row, int column)
        {
            if (_ws.Cells[row, column].Value2 != null)
            {
                return _ws.Cells[row, column].Value2;
            }
            else
            {
                return "";
            }
        }

        public int FindRow(PaymentResult result)
        {
            string currentStr = "";
            for (int i = 12; i < 47; i++)
            {
                currentStr = ReadCell(i, 1).Trim();
                if (result.Law.Trim() == currentStr)
                    return i;
            }

            return -1;
        }

        public void Write(PaymentResult result, int row, LawEnum law)
        {
            if (law == LawEnum.Chaes)
            {
                Range range = (Range)_ws.Range[_ws.Cells[row, 2], _ws.Cells[row, 4]];
                range.Value2 = PaymentResultExtension.ToArrayPaymentResult(result);
            }
            else if (law == LawEnum.Maiak)
            {
                Range range = (Range)_ws.Range[_ws.Cells[row, 5], _ws.Cells[row, 7]];
                range.Value2 = PaymentResultExtension.ToArrayPaymentResult(result);
            }
            else if (law == LawEnum.Semipalat)
            {
                Range range = (Range)_ws.Range[_ws.Cells[row, 8], _ws.Cells[row, 10]];
                range.Value2 = PaymentResultExtension.ToArrayPaymentResult(result);
            }
        }

        public void WritePersons(List<Person> persons)
        {
            int row = 3;
            foreach (var person in persons)
            {
                Range range = (Range)_ws.Range[_ws.Cells[row, 1], _ws.Cells[row, 7]];
                range.Value2 = person.ToArrayPerson();
                row++;
            }
        }

        public void SetWidth()
        {
            _ws.Columns.AutoFit();
        }

        public void WriteRajonHeader(string district)
        {
            _ws.Cells[4, 1].Value2 = district;
        }

        public void WriteRajonHeaderForReport(string district)
        {
            _ws.Cells[1, 1].Value2 = district;
        }

        public void SetZeroValue()
        {
            for (int i = 12; i < 47; i++)
            {
                for (int j = 1; j < 11; j++)
                {
                    if (_ws.Cells[i, j].Value2 == null)
                    {
                        _ws.Cells[i, j].Value2 = "0,0";
                    }
                }
            }
        }

        public void Save()
        {
            _wb.Save();
        }

        public void SaveAs(string path)
        {
            _wb.SaveAs(path);
        }

        public void Close()
        {
            _wb.Close();
        }

    }
}
