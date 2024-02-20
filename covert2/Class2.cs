//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Text;
//using System.Text.Json;
//using Newtonsoft.Json.Linq;
//using Newtonsoft.Json;
//using OfficeOpenXml;

//public class CacBuoc
//{
//    public string TenBuoc { get; set; }
//    //public string HuongDan { get; set; }
//    public string GhiChu { get; set; }
//}

//public class ChiTiet
//{
//    public string NoiDung { get; set; }
//    //public List<T> CacBuocs { get; set; }
//}

//public class ChucNang
//{
//    public string ChucNangChinh { get; set; }
//    public List<ChiTiet> ChiTiets { get; set; }
//}

//public class NoiDungSheet
//{
//    public string TenSheet { get; set; }
//    public List<ChucNang> ChucNangs { get; set; }
//}

//class Program
//{
//    static void Main()
//    {
//        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

//        Console.OutputEncoding = Encoding.UTF8;
//        string filePath = "D:\\Downloads\\Hướng dẫn sử dụng Bmate.xlsx";

//        using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
//        {

//            var worksheet = package.Workbook.Worksheets[1];


//            var noiDungSheet = new NoiDungSheet
//            {
//                TenSheet = worksheet.Name,
//                ChucNangs = new List<ChucNang>()
//            };

//            int rowCount = worksheet.Dimension.Rows;
//            int colCount = worksheet.Dimension.Columns;

//            int chucNangChinhColumnIndex = 1;
//            int chiTietChucNangColumnIndex = 2;
//            int cacBuocColumnIndex = 3;
//            int huongDanColumnIndex = 4;
//            int ghiChuColumnIndex = 5;

//            ChucNang currentChucNang = null;
//            ChiTiet currentChiTiet = null;

//            for (int row = 2; row <= rowCount; row++)
//            {
//                string chucNangChinh = worksheet.Cells[row, chucNangChinhColumnIndex].Text;
//                string chiTietChucNang = worksheet.Cells[row, chiTietChucNangColumnIndex].Text;
//                string cacBuoc = worksheet.Cells[row, cacBuocColumnIndex].Text;
//                string huongDan = worksheet.Cells[row, huongDanColumnIndex].Text;
//                string ghiChu = worksheet.Cells[row, ghiChuColumnIndex].Text;

//                if (!string.IsNullOrEmpty(chucNangChinh))
//                {
//                    currentChucNang = new ChucNang
//                    {
//                        ChucNangChinh = chucNangChinh,
//                        ChiTiets = new List<ChiTiet>()
//                    };

//                    noiDungSheet.ChucNangs.Add(currentChucNang);
//                }

//                if (!string.IsNullOrEmpty(chiTietChucNang))
//                {
//                    currentChiTiet = new ChiTiet
//                    {
//                        NoiDung = chiTietChucNang,
//                        //CacBuocs = new List<T>()
//                    };

//                    currentChucNang.ChiTiets.Add(currentChiTiet);
//                }

//                if (!string.IsNullOrEmpty(cacBuoc))
//                {
//                    //cacBuoc = cacBuoc + "\"";
//                    //string newStr = cacBuoc.Remove(6, 7);
//                    //huongDan = '"' + huongDan;
//                    //string newStr2 = huongDan.Remove(0, 1);
//                    string json = JsonConvert.SerializeObject(new
//                    {
//                        cacBuoc. = huongDan
//                    });
//                    Console.WriteLine(json);
//                    //currentChiTiet.CacBuocs.Add(new{cacBuoc = huongDan, });
//                }
//            }

//            Console.WriteLine($"TenSheet: {noiDungSheet.TenSheet}");

//            var string2 = System.Text.Json.JsonSerializer.Serialize(noiDungSheet);
//            string2 = string2.Replace("TenSheet", "Sheet");
//            string2 = string2.Replace("ChucNangChinh", "Chức năng chính");
//            string2 = string2.Replace("ChiTiets", "Chi tiết chức năng");
//            string2 = string2.Replace("NoiDung", "Nội dung");

//            string2 = string2.Replace("HuongDan", "Hướng dẫn");

//            string2 = string2.Replace("CacBuocs", "Các bước thực hiện");
//            string2 = string2.Replace("GhiChu", "Ghi chú");

//            //string2 = string2.Replace("\"TenBuoc\":", " ");
//            DateTime now = DateTime.Now;
//            string fileName = $"fileConvert_{now.Hour}_{now.Minute}_{now.Second}_{now.Year}_{now.Month}_{now.Day}.txt";

//            string filePath2 = Path.Combine("D:\\Downloads", fileName);
//            WriteToFile(filePath2, string2);
//        }
//        Console.WriteLine("Chuỗi đã được viết vào tệp.");
//    }

//    static void WriteToFile(string filePath, string content)
//    {
//        // Sử dụng using để đảm bảo đóng StreamWriter sau khi sử dụng
//        using (StreamWriter writer = new StreamWriter(filePath))
//        {
//            // Ghi chuỗi vào tệp
//            writer.Write(content);
//        }
//    }
//}
