//using OfficeOpenXml;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace covert2.Class
//{
//    internal class Class3
//    {

//    }
//    class Program
//    {
//        static void Main()
//        {
//            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

//            Console.OutputEncoding = Encoding.UTF8;
//            string filePath = "D:\\Downloads\\Hướng dẫn sử dụng Bmate.xlsx";
//            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(filePath)))
//            {
//                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Lấy worksheet đầu tiên

//                // Bắt đầu DFS từ ô (1,1)
//                DfsExcel(worksheet, 1, 1);
//            }
//        }

//        static void DfsExcel(ExcelWorksheet worksheet, int row, int col)
//        {
//            // Lấy giá trị của ô tại hàng `row` và cột `col`
//            object value = worksheet.Cells[row, col].Value;

//            // In giá trị của ô
//            Console.WriteLine(value);

//            // Kiểm tra xem ô hiện tại có merge với ô khác không
//            ExcelRangeBase cell = worksheet.Cells[row, col];
//            if (cell.Merge)
//            {
//                // Lấy vùng merge
//                var mergedCells = worksheet.MergedCells;
//                // Tìm vùng merge chứa ô hiện tại
//                foreach (var range in mergedCells)
//                {
//                    if (range.Start.Row <= row && range.End.Row >= row &&
//                        range.Start.Column <= col && range.End.Column >= col)
//                    {
//                        // Lấy ô cha
//                        int parentRow = range.Start.Row;
//                        int parentCol = range.Start.Column;
//                        // Gọi đệ quy DFS với ô cha
//                        DfsExcel(worksheet, parentRow, parentCol);
//                    }
//                }
//            }

//            // Tìm ô con của ô hiện tại (nếu có)
//            int nextRow = row + 1;
//            while (worksheet.Cells[nextRow, col].Value != null)
//            {
//                DfsExcel(worksheet, nextRow, col);
//                nextRow++;
//            }
//        }
//    }
//}