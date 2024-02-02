using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace jfgSchedule
{
    class Utility
    {
        public const int configKey = 1;     // 稼働設定テーブル：レコードキー

        public class DBConnect
        {
            OleDbConnection cn = new OleDbConnection();

            public OleDbConnection Cn
            {
                get
                {
                    return cn;
                }
            }

            private string sServerName;
            private string sLogin;
            private string sPass;
            private string sDatabase;

            public DBConnect()
            {
                try
                {
                    // MySeting項目の取得
                    // サーバ名
                    sServerName = Properties.Settings.Default.ServerName;

                    // ログイン名
                    sLogin = Properties.Settings.Default.Login;

                    // パスワード
                    sPass = Properties.Settings.Default.Pass;

                    // データベース名
                    sDatabase = Properties.Settings.Default.Database;

                    // データベース接続文字列
                    cn.ConnectionString = "";
                    cn.ConnectionString += "Provider=SQLOLEDB;";
                    cn.ConnectionString += "SERVER=" + sServerName + ";";
                    cn.ConnectionString += "DataBase=" + sDatabase + ";";
                    cn.ConnectionString += "UID=" + sLogin + ";";
                    cn.ConnectionString += "PWD=" + sPass + ";";
                    //cn.ConnectionString += "WSID=";

                    cn.Open();

                }

                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        /// <summary>
        /// 文字列の値が数字かチェックする
        /// </summary>
        /// <param name="tempStr">検証する文字列</param>
        /// <returns>数字:true,数字でない:false</returns>
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     カード番号検証 （double型変換可能な値か検証）</summary>
        /// <param name="cel">
        ///     セルの値</param>
        /// <returns>
        ///     double型のとき値を返す、double型変換エラーのとき-1を返す</returns>
        ///-------------------------------------------------------------------------
        public static double cNumberCheck(string cel)
        {
            double rtn = -1;
            double cNo;
            if (double.TryParse(cel, out cNo))
            {
                rtn = cNo;
            }

            return rtn;
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     オブジェクトをstring型に変換して返す　</summary>
        /// <param name="obj">
        ///     対象のオブジェクト</param>
        /// <returns>
        ///     string型の戻り値</returns>
        /// -------------------------------------------------------------------------
        public static string nulltoString(object obj)
        {
            string sVal = string.Empty;

            if (obj == null)
            {
                sVal = string.Empty;
            }
            else
            {
                sVal = obj.ToString();
            }

            return sVal;
        }
        ///-------------------------------------------------------------------------
        /// <summary>
        ///     Excelファイルをパスワード付きでオープン・クローズする </summary>
        /// <param name="sPath">
        ///     Excelファイルパス</param>
        /// <param name="rPw">
        ///     読み込みパスワード</param>
        /// <param name="wPw">
        ///     書き込みパスワード</param>
        /// <param name="logFile">
        ///     ログファイルパス</param>
        /// <returns>
        ///     成功：true, 失敗：false</returns>
        ///-------------------------------------------------------------------------
        public static bool PwdXlsFile(string sPath, string rPw, string wPw, string logFile)
        {
            if (rPw == string.Empty)
            {
                return true;
            }

            // ログ出力
            System.IO.File.AppendAllText(logFile, Form1.GetNowTime(" Excelを起動しています..."), Encoding.GetEncoding(932));

            System.Threading.Thread.Sleep(100);
            Application.DoEvents();

            // エクセルオブジェクト
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = null;

            try
            {
                //if (wPw != string.Empty)
                //{
                //    lblMsg.Text = sPath + " のパスワードを解除しています...";
                //}
                //else
                //{
                //    lblMsg.Text = sPath + " を開いています...";
                //}

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                // Excelファイルを開く
                oXlsBook = (oXls.Workbooks.Open(sPath, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, wPw, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                    //if (wPw != string.Empty)
                    //{
                    //    lblMsg.Text = sPath + " のパスワードが解除されました...";
                    //}
                    //else
                    //{
                    //    lblMsg.Text = sPath + " を開きました...";
                    //}

                    System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                oXls.DisplayAlerts = false;

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                //// Excelファイル書き込み
                //oXlsBook.SaveAs(sPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, rPw,
                //                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                //                Type.Missing, Type.Missing);

                // Excelファイル書き込み
                oXlsBook.SaveAs(sPath, Type.Missing, rPw, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);

                //lblMsg.Text = sPath + " を保存しました...";

                // ログ出力
                if (rPw != string.Empty)
                {
                    System.IO.File.AppendAllText(logFile, Form1.GetNowTime(" " + sPath + " をパスワード付きで保存しました..."), Encoding.GetEncoding(932));
                }

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                //lblMsg.Text = "Excelを終了しました...";
                System.IO.File.AppendAllText(logFile, Form1.GetNowTime(" Excelを終了しました..."), Encoding.GetEncoding(932));

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {
                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                if (oXlsBook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;

                GC.Collect();
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     文字列を指定文字数をＭＡＸとして返します</summary>
        /// <param name="s">
        ///     文字列</param>
        /// <param name="n">
        ///     文字数</param>
        /// <returns>
        ///     文字数範囲内の文字列</returns>
        /// --------------------------------------------------------------------
        public static string GetStringSubMax(string s, int n)
        {
            string val;

            // 文字間のスペースを除去 2015/03/10
            s = s.Replace(" ", "");

            if (s.Length > n) val = s.Substring(0, n);
            else val = s;

            return val;
        }
    }
}
