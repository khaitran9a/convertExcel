using System.Collections.Generic;
using System;

namespace ConverExcelApp.Class
{
    public class ImportFileModel
    {
        public ImportFileModel()
        {
            items = new List<ImportFileDataModel>();


        }
        /// <summary>
        /// Là 1 phiên giao dịch nào đó, hoặc là giá trị số hóa đơn chứng từ lúc import
        /// </summary>
        public string FileName { get; set; }
        public List<ImportFileDataModel> items { get; set; }
       
    }

    public class ImportFileDataModel
    {
        public ImportFileDataModel()
        {
            itemModels = new List<FileItemExcel>();
        }
        //public String SessionId { get; set; }
        //public string ImportFileGuid { get; set; }
        public string ImportSheetName { get; set; }
        //public bool IsApplyOrders { get; set; }
        //public DateTime DateTransaction { get; set; }
        //public DateTime DateTransactionDen { get; set; }
        public IList<FileItemExcel> itemModels { get; set; }
    }

    public class FileItemExcel
    {
        public FileItemExcel()
        {
            ChiTiets = new List<ChiTiet>() ;
        }
      
        public string ChucNangChinh { get; set; }
        public List<ChiTiet> ChiTiets { get; set; }
    }
    public class ChiTiet
    {
        public ChiTiet()
        {
            CacBuocs = new List<CacBuoc>();
        }
        public string Name { get; set; }
        public List<CacBuoc> CacBuocs { get; set; }
    }
    public class CacBuoc
    {
        public string TenBuoc { get; set; }
        public string HuongDan { get; set; }
        public string GhiChu { get; set; }
    }
}
