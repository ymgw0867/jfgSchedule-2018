using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;

namespace scheTask
{
    class clsWorks
    {
        int sCnt = 2;       // 東西繰り返し回数
        int sheetNum = 10;  // シート数

        string [,] gengo;   // 言語配列

        string [,] sheetYYMM = new string [6, 2];       // 年月と開始列
        int sheetStRow = 4;                             // エクセルシート明細開始行
        //const int S_colSMAX = 194;                    // 稼働表Temp列Max
        int[] S_colSMAX = { 195, 196 };                 // 稼働表Temp列Max
        string[] sheetName = { "東", "西" };            // シート名見出し

        public clsWorks()
        {
            // 言語配列読み込み
            readLang();

            // 稼働表作成
            worksOutput();
        }

        ///----------------------------------------------------
        /// <summary>
        ///     稼働表作成 </summary>
        ///----------------------------------------------------
        public void worksOutput()
        {
            // Excelテンプレートシート開く
            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsKadouPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Workbook oXlsWork = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsTempPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = null;           // 入力シート
            Excel.Worksheet oxlsWorkSheet = null;       // 作業用ワークシート

            Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];
            Excel.Range rngDay;
            Excel.Range[] rngs = new Microsoft.Office.Interop.Excel.Range[31];

            DateTime stDate;
            DateTime edDate;
                        
            try
            {
                // 稼働予定開始年月日
                stDate = DateTime.Parse(DateTime.Today.Year.ToString() + "/" + DateTime.Today.Month.ToString() + "/01");
            
                // 稼働予定終了年月日
                edDate = stDate.AddMonths(6).AddDays(-1);

                int ew = 0;

                while (ew < sCnt)
                {
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[ew + 1];    // 東西テンプレートシート

                    // 言語別シートを作成
                    for (int i = 1; i <= gengo.GetLength(0); i++)
                    {
                        int sNum = i + (10 * ew);

                        // ページシートを追加する
                        oxlsSheet.Copy(Type.Missing, oXlsWork.Sheets[sNum]);

                        // カレントシート
                        oxlsWorkSheet = (Excel.Worksheet)oXlsWork.Sheets[sNum + 1];
                        oxlsWorkSheet.Name = sheetName[ew] + "・" + gengo[i - 1, 1];

                        int xCol = 0;   // 日列初期値

                        // 稼働予定期間のカレンダーをセット
                        for (int mon = 0; mon < 6; mon++)
                        {
                            // 該当月
                            DateTime wDt = stDate.AddMonths(mon);
                            xCol = 31 * mon + ew + 9;
                            oxlsWorkSheet.Cells[1, xCol] = wDt.ToShortDateString();　// 9,40,71,102,・・・

                            // 年月と開始列の配列にセット
                            sheetYYMM[mon, 0] = wDt.Year.ToString() + wDt.Month.ToString().PadLeft(2, '0');
                            sheetYYMM[mon, 1] = xCol.ToString();

                            DateTime dDay;

                            // 該当月の暦
                            int dy = 0;
                            while (dy < 31)
                            {
                                if (DateTime.TryParse(wDt.Year.ToString() + "/" + wDt.Month.ToString() + "/" + (dy + 1).ToString(), out dDay))
                                {
                                    oxlsWorkSheet.Cells[2, xCol + dy] = (dy + 1).ToString();    // 日
                                    oxlsWorkSheet.Cells[3, xCol + dy] = dDay.ToString("ddd");   // 曜日
                                }
                                else
                                {
                                    oxlsWorkSheet.Cells[2, xCol + dy] = string.Empty;
                                    oxlsWorkSheet.Cells[3, xCol + dy] = string.Empty;
                                }

                                dy++;
                            }
                        }

                        // 組合員予定申告データを取得
                        string cardNum = string.Empty;
                        string gCode = gengo[i - 1, 0];
                        int sRow = sheetStRow;

                        Control.DataControl con = new Control.DataControl();

                        StringBuilder sb = new StringBuilder();

                        if (ew == 0)
                        {
                            sb.Append("select * from 会員稼働予定 inner join ");
                            sb.Append("(select カード番号 as cardno,氏名, 携帯電話番号, JFG加入年 from 会員情報 ");
                            sb.Append("where (言語1 = " + gCode);
                            sb.Append(" or 言語2 = " + gCode);
                            sb.Append(" or 言語3 = " + gCode);
                            sb.Append(" or 言語4 = " + gCode);
                            sb.Append(" or 言語5 = " + gCode);
                            sb.Append(") and 東西 = 1) as a ");
                            sb.Append("on 会員稼働予定.カード番号 = a.cardno ");
                            sb.Append("order by 会員稼働予定.フリガナ, 会員稼働予定.カード番号, 会員稼働予定.年,会員稼働予定.月");
                        }
                        else if (ew == 1)
                        {
                            sb.Append("select 地域コード, 地域名, 会員情報.カード番号 as cardno,氏名, ");
                            sb.Append("携帯電話番号, JFG加入年, a.* ");
                            sb.Append("from 会員情報 inner join ");
                            sb.Append("(select * from 会員稼働予定) as a ");
                            sb.Append("on 会員情報.カード番号 = a.カード番号 ");
                            sb.Append("where (言語1 = " + gCode);
                            sb.Append("or 言語2 = " + gCode);
                            sb.Append("or 言語3 = " + gCode);
                            sb.Append("or 言語4 = " + gCode);
                            sb.Append("or 言語5 = " + gCode);
                            sb.Append(") and 東西 = 2 ");
                            sb.Append("order by 会員情報.地域コード,a.フリガナ, 会員情報.カード番号, a.年,a.月");                    
                        }                        

                        OleDbDataReader dR = con.FreeReader(sb.ToString());

                        while (dR.Read())
                        {
                            // 該当期間のデータか検証
                            string yymm = dR["年"].ToString() + dR["月"].ToString().PadLeft(2, '0');

                            bool yymmOn = false;
                            string col = string.Empty;

                            for (int iX = 0; iX < 6; iX++)
                            {
                                if (sheetYYMM[iX, 0] == yymm)
                                {
                                    col = sheetYYMM[iX, 1];
                                    yymmOn = true;
                                    break;
                                }
                            }

                            if (!yymmOn) continue; // 非該当期間のため読み飛ばし

                            // 組合員が変わったら行番号を加算する
                            if (cardNum != string.Empty && cardNum != dR["カード番号"].ToString())
                            {
                                //セル下部へ実線ヨコ罫線を引く
                                rng[0] = (Excel.Range)oxlsWorkSheet.Cells[sRow, 1];
                                rng[1] = (Excel.Range)oxlsWorkSheet.Cells[sRow, S_colSMAX[ew]];
                                oxlsWorkSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot;

                                sRow++;
                            }

                            cardNum = dR["カード番号"].ToString();    // カード番号

                            // アサイン担当者か検証する
                            jfgDataClassDataContext db = new jfgDataClassDataContext();
                            var s = db.アサイン担当者.Where(a => a.カード番号 == double.Parse(cardNum));
                            if (s.Count() > 0)
                            {
                                
                            }


                            // 作業用シートにデータ貼り付け
                            if (ew == 0)
                            { 
                                oxlsWorkSheet.Cells[sRow, 1] = dR["氏名"].ToString();
                                oxlsWorkSheet.Cells[sRow, 2] = dR["フリガナ"].ToString();
                                oxlsWorkSheet.Cells[sRow, 3] = dR["JFG加入年"].ToString();
                                oxlsWorkSheet.Cells[sRow, 4] = dR["携帯電話番号"].ToString();
                                oxlsWorkSheet.Cells[sRow, 5] = string.Empty;
                                oxlsWorkSheet.Cells[sRow, 6] = string.Empty;
                                oxlsWorkSheet.Cells[sRow, 7] = dR["備考"].ToString();
                                oxlsWorkSheet.Cells[sRow, 8] = Utility.nulltoString(dR["申告年月日"]);
                            }
                            else if (ew == 1)
                            {
                                oxlsWorkSheet.Cells[sRow, 1] = dR["地域名"].ToString();
                                oxlsWorkSheet.Cells[sRow, 2] = dR["氏名"].ToString();
                                oxlsWorkSheet.Cells[sRow, 3] = dR["フリガナ"].ToString();
                                oxlsWorkSheet.Cells[sRow, 4] = dR["JFG加入年"].ToString();
                                oxlsWorkSheet.Cells[sRow, 5] = dR["携帯電話番号"].ToString();
                                oxlsWorkSheet.Cells[sRow, 6] = string.Empty;
                                oxlsWorkSheet.Cells[sRow, 7] = string.Empty;
                                oxlsWorkSheet.Cells[sRow, 8] = Utility.nulltoString(dR["申告年月日"]);
                            }

                            // 予定申告内容
                            for (int j = 0; j < 31; j++)
                            {
                                oxlsWorkSheet.Cells[sRow, int.Parse(col) + j] = dR["d" + (j + 1).ToString()].ToString();
                            }
                        }

                        dR.Close();
                        con.Close();

                        // カレンダーにない日の列削除
                        bool colDelStatus = true;
                        while (colDelStatus)
                        {
                            for (int cl = (ew + 9); cl <= oxlsWorkSheet.UsedRange.Columns.Count; cl++)
                            {
                                rngDay = (Excel.Range)oxlsWorkSheet.Cells[2, cl];
                                if (!Utility.NumericCheck(rngDay.Text.Trim()))
                                {
                                    oxlsWorkSheet.Columns[cl].Delete();
                                    colDelStatus = true;
                                    break;
                                }
                                else
                                {
                                    colDelStatus = false;
                                }
                            }
                        }

                        // 列数を再取得
                        int colsMaxFin = oxlsWorkSheet.UsedRange.Columns.Count;

                        //セル下部へ実線ヨコ罫線を引く
                        rng[0] = (Excel.Range)oxlsWorkSheet.Cells[oxlsWorkSheet.UsedRange.Rows.Count, 1];
                        rng[1] = (Excel.Range)oxlsWorkSheet.Cells[oxlsWorkSheet.UsedRange.Rows.Count, colsMaxFin];
                        oxlsWorkSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        //表全体に実線縦罫線を引く
                        rng[0] = (Excel.Range)oxlsWorkSheet.Cells[1, 1];
                        rng[1] = (Excel.Range)oxlsWorkSheet.Cells[oxlsWorkSheet.UsedRange.Rows.Count, colsMaxFin];
                        oxlsWorkSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        //表全体の左端縦罫線
                        rng[0] = (Excel.Range)oxlsWorkSheet.Cells[1, 1];
                        rng[1] = (Excel.Range)oxlsWorkSheet.Cells[oxlsWorkSheet.UsedRange.Rows.Count, 1];
                        oxlsWorkSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        //表全体の右端縦罫線
                        rng[0] = (Excel.Range)oxlsWorkSheet.Cells[1, colsMaxFin];
                        rng[1] = (Excel.Range)oxlsWorkSheet.Cells[oxlsWorkSheet.UsedRange.Rows.Count, colsMaxFin];
                        oxlsWorkSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        // 出力シートコンソール表示
                        Console.WriteLine(oxlsWorkSheet.Name);                    
                    }

                    ew++;
                }
                
                // 作業用BOOKの1番目のシートは削除する
                ((Excel.Worksheet)oXlsWork.Sheets[1]).Delete();

                // カレントシート
                oxlsWorkSheet = (Excel.Worksheet)oXlsWork.Sheets[1];

                // ウィンドウを非表示にする
                oXls.Visible = false;

                //保存処理
                oXls.DisplayAlerts = false;

                oXlsWork.SaveAs(Properties.Settings.Default.xlsWorksPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
            }
            finally
            {
                // ウィンドウを非表示にする
                oXls.Visible = false;

                //保存処理
                oXls.DisplayAlerts = false;

                // ExcelBookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);
                oXlsWork.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excel終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsWorkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsWork);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);
            }
        }

        ///----------------------------------------------------
        /// <summary>
        ///     言語配列作成　</summary>
        ///----------------------------------------------------
        private void readLang()
        {
            jfgDataClassDataContext db = new jfgDataClassDataContext();
            var s = db.言語.Where(a => a.言語名1 != "J").OrderBy(a => a.言語番号);

            gengo = new string [s.Count(), 2];

            int i = 0;
            foreach (var item in s)
            {
                gengo[i, 0] = item.言語番号.ToString();
                gengo[i, 1] = item.言語名2;
                i++;
            }
        }
    }
}
