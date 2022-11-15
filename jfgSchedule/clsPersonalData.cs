using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Excel = Microsoft.Office.Interop.Excel;

namespace jfgSchedule
{
    /// ---------------------------------------------------
    /// <summary>
    ///     個人情報クラス </summary>
    /// ---------------------------------------------------
    class clsPersonalData
    {
        public double cNo { get; set; }
        public string furi { get; set; }
        public int eYear { get; set; }
        public int eMonth { get; set; }
        public DateTime sDt { get; set; }
        //public Excel.Worksheet exl { get; set; }
        public int colidx { get; set; }
        public string memo { get; set; }
        public int kadouDays { get; set; }
        public int jikoShinDays { get; set;}
        public ClosedXML.Excel.IXLWorksheet sheet { get; set; }     // 2018/02/22
    }
}
