using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace jfgSchedule
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            // GitHub masterブランチ作成：2018/10/04
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            // ログメッセージ
            string logText = string.Empty;

            // ログ書き出し先ファイルがあるか？なければ作成する
            string logFile = Properties.Settings.Default.xlsxPath + Properties.Settings.Default.logFileName;
            if (!System.IO.File.Exists(logFile))
            {
                System.IO.File.Create(logFile);
            }

            // 開始ログ出力
            logText = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " 処理を開始しました。" + Environment.NewLine;
            System.IO.File.AppendAllText(logFile, logText, System.Text.Encoding.GetEncoding(932));
            
            // 前回更新日時を取得
            //DateTime dt = DateTime.Parse("2015/03/03 13:00");
            DateTime dt = getUpdateDate();

            // エクセル予定申告シートより会員稼働予定テーブルを更新する
            clsXls xls = new clsXls();
            int uCnt = xls.xlsSelect(Properties.Settings.Default.xlsxPath, dt);

            // 前回更新日時フィールドに現在の日時を書き込む
            setUpdateDate();

            // 更新された予定申告データがあったとき
            if (uCnt > 0)
            {
                // アサイン担当用稼働表エクセルシートを作成する
                clsWorks cw = new clsWorks();

                // ログ出力
                logText = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " アサイン担当用稼働表を更新しました。" + Environment.NewLine;
                System.IO.File.AppendAllText(logFile, logText, System.Text.Encoding.GetEncoding(932));            
            }

            // 終了ログ出力
            logText = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " 処理を終了しました。" + Environment.NewLine;
            System.IO.File.AppendAllText(logFile, logText, System.Text.Encoding.GetEncoding(932));
            
            // 終了
            Environment.Exit(0);
        }

        /// ----------------------------------------------------
        /// <summary>
        ///     前回更新日時を取得する </summary>
        /// <returns>
        ///     前回更新日時 </returns>
        /// ----------------------------------------------------
        private DateTime getUpdateDate()
        {
            DateTime dt = DateTime.Parse("1900/01/01 00:00:00");

            jfgDataClassDataContext db = new jfgDataClassDataContext();
            var s = db.稼働設定.Single(a => a.ID == Utility.configKey);

            if (s.前回更新日時 != null)
            {
                dt = (DateTime)s.前回更新日時;
            }

            return dt;
        }

        /// ----------------------------------------------------
        /// <summary>
        ///     前回更新日時を更新する </summary>
        /// ----------------------------------------------------
        private void setUpdateDate()
        {
            jfgDataClassDataContext db = new jfgDataClassDataContext();
            var s = db.稼働設定.Single(a => a.ID == Utility.configKey);

            s.前回更新日時 = DateTime.Now;

            db.SubmitChanges();
        }
    }
}
