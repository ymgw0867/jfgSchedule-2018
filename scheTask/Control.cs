using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace scheTask
{
    class Control
    {
        /// <summary>
        /// DataControlクラスの基本クラス
        /// </summary>
        public class BaseControl
        {
            private Utility.DBConnect DBConnect;

            //BaseControlのコンストラクタ。DBConnectクラスのインスタンスを作成します。
            public BaseControl()
            {
                DBConnect = new Utility.DBConnect();
            }

            public OleDbConnection GetConnection()
            {
                return DBConnect.Cn;
            }
        }

        public class DataControl : BaseControl
        {
            public OleDbConnection Cn = new OleDbConnection();

            //データコントロールクラスのコンストラクタ
            public DataControl()
            {
            }

            /// <summary>
            /// データベース接続解除
            /// </summary>
            public void Close()
            {
                if (Cn.State == ConnectionState.Open)
                {
                    Cn.Close();
                }
            }

            /// -------------------------------------------------------------
            /// <summary>
            ///     データリーダーを取得する </summary>
            /// <param name="tempSQL">
            ///     SQL文    </param>
            /// <returns>
            ///     データリーダー</returns>
            /// -------------------------------------------------------------
            public OleDbDataReader FreeReader(string tempSQL)
            {
                Cn = GetConnection();
                OleDbCommand sCom = new OleDbCommand();
                sCom.Connection = Cn;
                sCom.CommandText = tempSQL;
                return sCom.ExecuteReader();
            }
        }
    }
}
