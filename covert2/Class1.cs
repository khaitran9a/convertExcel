//using OfficeOpenXml;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Text.Json;

//public class NoiDungSheet
//{
//    public string TenSheet { get; set; }
//    public List<ChucNang> ChucNangs { get; set; }
//}

//public class ChucNang
//{
//    public string ChucNangChinh { get; set; }
//    public List<ChiTiet> ChiTiets { get; set; }
//}

//public class ChiTiet
//{
//    public string NoiDung { get; set; }
//    public List<CacBuoc> CacBuocs { get; set; }
//}

//public class CacBuoc
//{
//    public string TenBuoc { get; set; }
//    public string HuongDan { get; set; }
//    public string GhiChu { get; set; }
//}

//public class ExcelReader
//{
//    public NoiDungSheet ReadExcel(string filePath)
//    {
//        NoiDungSheet noiDungSheet = new NoiDungSheet();
//        noiDungSheet.ChucNangs = new List<ChucNang>();

//        FileInfo fileInfo = new FileInfo(filePath);

//        using (ExcelPackage package = new ExcelPackage(fileInfo))
//        {
//            ExcelWorksheet worksheet = package.Workbook.Worksheets[1]; // Lấy sheet thứ nhất

//            // Đọc tiêu đề của sheet
//            noiDungSheet.TenSheet = worksheet.Name;

//            // Đọc dữ liệu từ cột CHỨC NĂNG CHÍNH
//            var chucNangs = worksheet.Cells[2, 1, worksheet.Dimension.End.Row, 1]
//                .GroupBy(cell => cell.Text)
//                .Select(group => group.First())
//                .ToList();

//            foreach (var chucNangCell in chucNangs)
//            {
//                ChucNang chucNang = new ChucNang();
//                chucNang.ChucNangChinh = chucNangCell.Text;
//                chucNang.ChiTiets = new List<ChiTiet>();

//                // Đọc dữ liệu từ cột Chi tiết chức năng
//                var chiTiets = worksheet.Cells[chucNangCell.Start.Row, 2, chucNangCell.Start.Row, worksheet.Dimension.End.Column]
//                    .GroupBy(cell => cell.Text)
//                    .Select(group => group.First())
//                    .ToList();

//                foreach (var chiTietCell in chiTiets)
//                {
//                    ChiTiet chiTiet = new ChiTiet();
//                    chiTiet.NoiDung = chiTietCell.Text;
//                    chiTiet.CacBuocs = new List<CacBuoc>();

//                    // Đọc dữ liệu từ cột Các bước
//                    var cacBuocs = worksheet.Cells[chiTietCell.Start.Row, 3, chiTietCell.Start.Row, worksheet.Dimension.End.Column]
//                        .ToList();

//                    for (int i = 0; i < cacBuocs.Count; i += 3)
//                    {
//                        CacBuoc cacBuoc = new CacBuoc();
//                        cacBuoc.TenBuoc = cacBuocs[i].Text;
//                        cacBuoc.HuongDan = cacBuocs[i + 1].Text;
//                        cacBuoc.GhiChu = cacBuocs[i + 2].Text;

//                        chiTiet.CacBuocs.Add(cacBuoc);
//                    }

//                    chucNang.ChiTiets.Add(chiTiet);
//                }

//                noiDungSheet.ChucNangs.Add(chucNang);
//            }
//        }

//        return noiDungSheet;
//    }
//}

//class Class1
//{
//    static void Main()
//    {
//        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
//        string excelFilePath = "D:\\Downloads\\Hướng dẫn sử dụng Bmate.xlsx";
//        ExcelReader excelReader = new ExcelReader();
//        NoiDungSheet result = excelReader.ReadExcel(excelFilePath);
//        Console.WriteLine(JsonSerializer.Serialize(result));
//        // Sử dụng đối tượng result theo nhu cầu của bạn
//    }
//}
