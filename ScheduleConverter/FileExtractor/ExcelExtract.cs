using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace ScheduleConverter.FileExtractor
{
    public static class ExcelExtract
    {
        public static List<DataTable> DataReader(string fileToRead,string sheetName,string col)
        {
            var dts = new List<DataTable>();
            using (FileStream file = new FileStream(fileToRead, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = new StreamReader(file))
                {
                    using (var package = new ExcelPackage(file))
                    {
                       
                        for (int i = 1; i <= package.Workbook.Worksheets.Count; i++)
                        {
                            var worksheet = package.Workbook.Worksheets[i];
                            DataTable table = new DataTable();
                            if (string.IsNullOrEmpty(sheetName) || worksheet.Name.Equals(sheetName))
                            {
                                string sheet = worksheet.Name;
                                int colCount = worksheet.Dimension.End.Column;
                                for(int colinx=1; colinx<= colCount; colinx++)
                                {
                                    table.Columns.Add("Col" + colinx);
                                }
                            }
                            var rowCount = worksheet.Dimension.Rows;
                            for(var rowNumber=1;rowNumber<= rowCount; rowNumber++)
                            {
                                var cellText = worksheet.Cells[col + rowNumber + ":" + col + rowNumber].Text;
                                if (!string.IsNullOrEmpty(cellText))
                                {
                                    var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                                    var data = ((object[,])row.Value).Cast<dynamic>().ToArray();
                                    table.Rows.Add(data);
                                }

                            }
                            dts.Add(table);
                        }

                    }
                }
            }
            return dts;
        }
    }
}
