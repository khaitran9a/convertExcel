using System;
using System.IO;
using System.Text;
using ConverExcelApp.Class;
using OfficeOpenXml;
using System.Text.Json;


class Program
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Thay đổi đường dẫn đến file Excel của bạn
        string excelFilePath = "D:\\Downloads\\Hướng dẫn sử dụng Bmate.xlsx";
        Console.OutputEncoding = Encoding.UTF8;

        // Gọi hàm để đọc và in ra tên các sheet
        ReadSheetNames(excelFilePath);
    }

    static int FindLastRowWithData(ExcelWorksheet worksheet)
    {
        int lastRow = worksheet.Dimension.End.Row;

        for (int row = lastRow; row >= 1; row--)
        {
            // Kiểm tra dữ liệu trong một số cột quan trọng, chẳng hạn cột A
            object valueInColumnA = worksheet.Cells[row, 3].Value;

            // Kiểm tra nếu dữ liệu trong cột A của hàng này có tồn tại hay không
            if (valueInColumnA != null && !string.IsNullOrWhiteSpace(valueInColumnA.ToString()))
            {
                // Trả về số hàng đầu tiên mà có dữ liệu
                return row;
            }
        }

        // Nếu không tìm thấy dữ liệu nào, trả về 0
        return 0;
    }

    static void WriteToFile(string filePath, string content)
    {
        using (StreamWriter writer = new StreamWriter(filePath))
        {
            writer.Write(content);
        }
    }
    static string EscapeSpecialCharacters(string input)
    {
        input = input.Replace("\\", "\\\\"); 
        input = input.Replace("\"", "\\\"");
        input = input.Replace("\n", "\\n"); 

        return input;
    }
    static void ReadSheetNames(string filePath)
    {
        FileInfo fileInfo = new FileInfo(filePath);
        var file = new ImportFileModel();
        file.FileName = fileInfo.Name;
        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            // Lặp qua các sheet trong file Excel
            int chucNangChinhColumnIndex = 1;
            int chiTietChucNangColumnIndex = 2;
            int cacBuocColumnIndex = 3;
            int huongDanColumnIndex = 4;
            int ghiChuColumnIndex = 5;
            StringBuilder sb = new StringBuilder();
            sb.Append("{\"FileExcel\": [");
            foreach (var sheet in package.Workbook.Worksheets)
            {
                //if (sheet.Index != 6)
                //{
                //    continue;
                //}

                sb.Append("{\"Sheet\": ");
                sb.Append("\"" + sheet.Name + "\"");
                sb.Append(",");
                sb.Append("\"Chức năng\": [");
                


                // Lấy chỉ dòng và cột có dữ liệu
                int startRow = sheet.Dimension.Start.Row + 1;
                int endRow = FindLastRowWithData(sheet);
           

                // Lặp qua từng cột từ phải qua trái
                for (int row = startRow; row <= endRow;)
                {
                    sb.Append("{\"Tên chức năng\": ");
                    var cell = sheet.Cells[row, chucNangChinhColumnIndex];
                    string chucnangchinh = cell.Text;
                    chucnangchinh = EscapeSpecialCharacters(chucnangchinh);
                    sb.Append("\"" + chucnangchinh + "\",");
                    var mergedRows = 1;
                    sb.Append("\"Chi tiết chức năng\": [");
                    if (cell.Merge)
                    {
                        int mergedColumns = 0;

                        // Lặp qua từng cell trong vùng merge để đếm số dòng và số cột
                        foreach (var mergedCell in sheet.MergedCells)
                        {
                            var mergedRange = sheet.Cells[mergedCell];
                            if (mergedRange.Start.Row == cell.Start.Row && mergedRange.Start.Column == cell.Start.Column)
                            {
                                mergedRows = mergedRange.Rows;
                                mergedColumns = mergedRange.Columns;
                                
                            }
                        }
                    }
                    for (int row2 = row; row2 < row + mergedRows;)
                    {
                        sb.Append("{\"Nội dung\": ");
                        var cell2 = sheet.Cells[row2, chiTietChucNangColumnIndex];
                        string chitiet = cell2.Text;
                        chitiet = EscapeSpecialCharacters(chitiet);

                        sb.Append("\"" + chitiet + "\",");
                        var mergedRows2 = 1;
                        sb.Append("\"Các bước thực hiện\": [");
                        if (cell2.Merge)
                        {
                            // Lặp qua từng cell trong vùng merge để đếm số dòng và số cột
                            foreach (var mergedCell in sheet.MergedCells)
                            {
                                var mergedRange = sheet.Cells[mergedCell];
                                if (mergedRange.Start.Row == cell2.Start.Row && mergedRange.Start.Column == cell2.Start.Column)
                                {
                                    mergedRows2 = mergedRange.Rows;

                                }
                            }
                        }

                        for (int row3 = row2; row3 < row2 + mergedRows2;)
                        {
                            var cell3 = sheet.Cells[row3, cacBuocColumnIndex];
                            string buoc = cell3.Text;
                            buoc = EscapeSpecialCharacters(buoc);

                            sb.Append("{\""+buoc+"\": ");
                            var cell4 = sheet.Cells[row3, huongDanColumnIndex];
                            string huongdan = cell4.Text;
                            huongdan = EscapeSpecialCharacters(huongdan);

                            sb.Append("\"" + huongdan + "\",");
                            var cell5 = sheet.Cells[row3, ghiChuColumnIndex];
                            string ghichu = cell5.Text;
                            sb.Append("\"Ghi chú\": ");
                            ghichu = EscapeSpecialCharacters(ghichu);
                            sb.Append("\"" + ghichu + "\"");
                            sb.Append("}");
                            if (row3 < row2 + mergedRows2-1)
                            {
                                sb.Append(",");
                            }
                            row3++;
                        }
                        sb.Append("]}");
                        if (row2 < row + mergedRows - 1)
                        {
                            sb.Append(",");
                        }
                        row2 += mergedRows2;
                    }
                    sb.Append("]}");
                    if (row < endRow-mergedRows)
                    {
                        sb.Append(",");
                    }
                    row += mergedRows;
                }
                sb.Append("]}");
                if (sheet.Index < package.Workbook.Worksheets.Count -1)
                {
                    sb.Append(",");
                }
            }
            sb.Append("]}");

            var json = sb.ToString();
            DateTime now = DateTime.Now;
            string fileName = $"fileConvert_{now.Hour}_{now.Minute}_{now.Second}_{now.Year}_{now.Month}_{now.Day}.txt";

            string filePath2 = Path.Combine("D:\\Downloads", fileName);
            WriteToFile(filePath2, json);

        }
    }
}
