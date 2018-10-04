using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace scheTask
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
        public Excel.Worksheet exl { get; set; }
        public int colidx { get; set; }
        public string memo { get; set; }
    }
}
