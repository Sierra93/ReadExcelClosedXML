using ClosedXML.Excel;
using System;
using System.Data;
using System.Diagnostics;

namespace ReadExcelClosedXML {
    class Program {
        static void Main() {
            var datatable = new DataTable();
            var workbook = new XLWorkbook(filePath);
            var xlWorksheet = workbook.Worksheet(1);
            var range = xlWorksheet.Range(xlWorksheet.FirstCellUsed(), xlWorksheet.LastCellUsed());
            int col = range.ColumnCount();
            int row = range.RowCount();

            datatable.Clear();

            for (int i = 1; i <= col; i++) {
                var column = xlWorksheet.Cell(1, i);
                datatable.Columns.Add(column.Value.ToString());
            }

            int firstHeadRow = 0;

            foreach (var item in range.Rows()) {
                if (firstHeadRow != 0) {
                    var array = new object[col];

                    for (int y = 1; y <= col; y++) {
                        array[y - 1] = item.Cell(y).Value;
                    }

                    datatable.Rows.Add(array);
                }

                firstHeadRow++;
            }

            foreach (DataRow oRow in datatable.Rows) {
                //oRow.ItemArray
                Debugger.Break();
            }
        }
    }
}
