using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace scheTask
{
    class Program
    {
        static void Main(string[] args)
        {
            // ログ書き出し先ファイルがあるか？なければ作成する
            string logFile = Properties.Settings.Default.xlsxPath + Properties.Settings.Default.logFileName;
            if (!System.IO.File.Exists(logFile)) System.IO.File.Create(logFile);

            // エクセル予定申告シートより会員稼働予定テーブルを更新する
            DateTime dt = DateTime.Parse("2015/02/20 13:00");
            clsXls xls = new clsXls();
            xls.xlsSelect(Properties.Settings.Default.xlsxPath, dt);

            // アサイン担当用稼働表エクセルシートを作成する
            clsWorks cw = new clsWorks();
        }
    }
}
