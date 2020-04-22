using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;

namespace ExcelToJsonConverter2
{
    public class ExcelReader
    {
        public Excel.Application Application;
        public Excel.Workbooks Workbooks;
        public Excel.Workbook Workbook;
        public Excel.Sheets Sheets;
        public Excel.Worksheet Sheet;
        public Excel.Range Cells;
        public Excel.Range Range;
        

        public ExcelReader()
        {
            this.Application = new Excel.Application();
            this.Workbooks = Application.Workbooks;
            this.Workbook = Application.Workbooks.Add(FIleManager.FilePath);
            this.Sheets = this.Workbook.Worksheets;
            this.Sheet = Workbook.Worksheets.Item[1];
            this.Cells = (Excel.Range)Sheet.Cells;
            this.Range = this.Sheet.UsedRange;
        }

        public bool CheckHasType()
        {
            string[] typeArray = { "int", "float", "double", "string",
                "short", "long", "char", "bool", "uint", "byte"};

            foreach (string type in typeArray)
            {
                if (this.Range.Cells[2, 1].Value.ToString() == type)
                    return true;
            }
            return false;
        }

        public JArray GetJsonArray()
        {
            JArray rtnArray = new JArray();

            //데이터 이름 뽑아내기.
            List<string> nameList = new List<string>();
            for(int i = 1; i <= Range.Columns.Count; i++)
            {
                nameList.Add(Range.Cells[1, i].Value.ToString());
            }

            //모든 컬럼을 돌며 JArray로 변환.
            int startRow = (this.CheckHasType())? 3 : 2;
            for (int row = startRow; row <= Range.Rows.Count; row++)
            {
                var jObject = new JObject();

                for (int col = 1; col <= Range.Columns.Count; col++)
                {
                    jObject.Add(nameList[col - 1], Range.Cells[row, col].Value);
                }
                Console.WriteLine(jObject.ToString());
                rtnArray.Add(jObject);
            }

            return rtnArray;
        }

        public void Free()
        {
            //저장할지 물어보는거 취소.
            this.Application.DisplayAlerts = false;
            this.Application.Quit();

            Marshal.ReleaseComObject(this.Range);
            Marshal.ReleaseComObject(this.Cells);
            Marshal.ReleaseComObject(this.Sheet);
            Marshal.ReleaseComObject(this.Sheets);
            Marshal.ReleaseComObject(this.Workbook);
            Marshal.ReleaseComObject(this.Workbooks);
            Marshal.ReleaseComObject(this.Application);
        }
    }
}
