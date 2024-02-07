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
            // ログ書き出し先ファイルがあるか？なければ作成する
            string logFile = Properties.Settings.Default.xlsxPath + Properties.Settings.Default.logFileName;
            if (!System.IO.File.Exists(logFile))
            {
                System.IO.File.Create(logFile);
            }

            // 開始ログ出力
            System.IO.File.AppendAllText(logFile, GetNowTime(" 処理を開始しました。"), System.Text.Encoding.GetEncoding(932));
            
            // 前回更新日時を取得
            DateTime dt = GetUpdateDate();

            // データベース更新エラーで未更新の予定申告データを2023/11/1に遡って更新：ただし、2024/02/07に限って実行：2024/02/07
            if (DateTime.Today == new DateTime(2024, 02, 07))
            {
                dt = new DateTime(2023, 11, 01);
            }

            // エクセル予定申告シートより会員稼働予定テーブルを更新する
            clsXls xls = new clsXls();
            _ = xls.xlsSelect(Properties.Settings.Default.xlsxPath, dt);

            // 前回更新日時フィールドに現在の日時を書き込む
            SetUpdateDate();

            // コメント化：本稼働開始後は「更新された予定申告データがあったとき」に戻す：2023/03/28
            //// 更新された予定申告データがあったとき
            //if (uCnt > 0)
            //{
            //    // アサイン担当用稼働表エクセルシートとホテル向けガイド稼働表を作成する：2022/11/08
            //    _ = new clsWorks(logFile);

            //    // 過去の予定表データを削除する：2023/02/17
            //    if (DeletePastData(logFile))
            //    {
            //        // ログ出力
            //        System.IO.File.AppendAllText(logFile, GetNowTime(" 前月までの予定表データを削除しました"), System.Text.Encoding.GetEncoding(932));
            //    }
            //}
            //else
            //{
            //    // ログ出力
            //    System.IO.File.AppendAllText(logFile, GetNowTime(" 更新された予定申告データはありませんでした。"), System.Text.Encoding.GetEncoding(932));
            //}

            // 旧ホテル向けと並行稼働のため予定申告データの有無にかかわらず稼働予定表を作成：2023/03/28
            // アサイン担当用稼働表エクセルシートとホテル向けガイド稼働表を作成する
            _ = new clsWorks(logFile);

            // 過去の予定表データを削除する：2023/02/17
            if (DeletePastData(logFile))
            {
                // ログ出力
                System.IO.File.AppendAllText(logFile, GetNowTime(" 前月までの予定表データを削除しました"), System.Text.Encoding.GetEncoding(932));
            }

            // 終了ログ出力
            System.IO.File.AppendAllText(logFile, GetNowTime(" 処理を終了しました。"), System.Text.Encoding.GetEncoding(932));
            
            // 終了
            Environment.Exit(0);
        }

        public static string GetNowTime(string msg)
        {
            return DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + ":" + DateTime.Now.Millisecond.ToString().PadLeft(3, '0') + msg + Environment.NewLine;
        }

        /// ----------------------------------------------------
        /// <summary>
        ///     前回更新日時を取得する </summary>
        /// <returns>
        ///     前回更新日時 </returns>
        /// ----------------------------------------------------
        private DateTime GetUpdateDate()
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
        private void SetUpdateDate()
        {
            jfgDataClassDataContext db = new jfgDataClassDataContext();
            var s = db.稼働設定.Single(a => a.ID == Utility.configKey);

            s.前回更新日時 = DateTime.Now;

            db.SubmitChanges();
        }

        /// <summary>
        ///     前月までの予定表データを削除する </summary>
        /// <param name="logFile">
        ///     ログファイルパス</param>
        private bool DeletePastData(string logFile)
        {
            try
            {
                var yymm = DateTime.Today.Year * 100 + DateTime.Today.Month;

                jfgDataClassDataContext db = new jfgDataClassDataContext();
                db.会員稼働予定.DeleteAllOnSubmit(db.会員稼働予定.Where(a => (a.年 * 100 + a.月) < yymm));
                db.SubmitChanges();

                return true;
            }
            catch (Exception ex)
            {
                // ログ出力
                System.IO.File.AppendAllText(logFile, GetNowTime(ex.Message), System.Text.Encoding.GetEncoding(932));
                return false;
            }
        }
    }
}
