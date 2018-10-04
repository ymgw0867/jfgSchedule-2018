using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;

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
    }
}
