using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Data.OleDb;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;

namespace jfgSchedule
{
    class clsWorks
    {
        int sCnt = 2;                                   // 東西繰り返し回数
        string[,] gengo;                               // 言語配列
        string[,] sheetYYMM = new string[6, 2];       // 年月と開始列
        int sheetStRow = 4;                             // エクセルシート明細開始行
        int[] S_colSMAX = { 195, 196 };                 // 稼働表Temp列Max
        string[] sheetName = { "東", "西" };            // シート名見出し
        const int cEAST = 0;                            // 東定数
        const int cWEST = 1;                            // 西定数

        public clsWorks()
        {
            // 言語配列読み込み
            readLang();

            // 稼働表作成
            worksOutputXML();

            // ホテル向けガイドリスト(英語) 稼働表作成
            worksOutputXML_FromExcel();
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

                        jfgDataClassDataContext db = new jfgDataClassDataContext();

                        // 東・LINQ
                        var linqEast = db.会員情報.Where(a => (a.言語1 == int.Parse(gCode) || a.言語2 == int.Parse(gCode) ||
                                                       a.言語3 == int.Parse(gCode) || a.言語4 == int.Parse(gCode) ||
                                                       a.言語5 == int.Parse(gCode)) && a.東西 == 1)
                                             .OrderBy(a => a.会員稼働予定.フリガナ).ThenBy(a => a.会員稼働予定.カード番号).ThenBy(a => a.会員稼働予定.年).ThenBy(a => a.会員稼働予定.月)
                                             .Select(a => new
                                             {
                                                 cardno = a.カード番号,
                                                 氏名 = a.氏名,
                                                 携帯電話番号 = a.携帯電話番号,
                                                 JFG加入年 = a.JFG加入年,
                                                 a.会員稼働予定
                                             });

                        // 西・LINQ
                        var linqWest = db.会員情報.Where(a => (a.言語1 == int.Parse(gCode) || a.言語2 == int.Parse(gCode) ||
                                                       a.言語3 == int.Parse(gCode) || a.言語4 == int.Parse(gCode) ||
                                                       a.言語5 == int.Parse(gCode)) && a.東西 == 2)
                                             .OrderBy(a => a.地域コード).ThenBy(a => a.会員稼働予定.フリガナ).ThenBy(a => a.カード番号).ThenBy(a => a.会員稼働予定.年).ThenBy(a => a.会員稼働予定.月)
                                             .Select(a => new
                                             {
                                                 地域コード = a.地域コード,
                                                 地域名 = a.地域名,
                                                 cardno = a.カード番号,
                                                 氏名 = a.氏名,
                                                 携帯電話番号 = a.携帯電話番号,
                                                 JFG加入年 = a.JFG加入年,
                                                 a.会員稼働予定
                                             });

                        if (ew == cEAST)    // 東
                        {
                            // 組合員予定申告データクラスのインスタンス生成
                            clsWorksTbl w = new clsWorksTbl();
                            w.cardNumBox = string.Empty;
                            w.sRow = sheetStRow;
                            w.ew = ew;

                            foreach (var t in linqEast)
                            {
                                w.cardNo = t.cardno;
                                w.氏名 = t.氏名;
                                w.携帯電話番号 = t.携帯電話番号;
                                w.JFG加入年 = (short)t.JFG加入年;
                                w.会員稼働予定 = t.会員稼働予定;

                                // エクセル稼働表作成
                                if (!xlsCellsSet(w, ref oxlsWorkSheet)) continue;
                            }
                        }
                        else if (ew == cWEST)   // 西
                        {
                            // 組合員予定申告データクラスのインスタンス生成
                            clsWorksTbl w = new clsWorksTbl();
                            w.cardNumBox = string.Empty;
                            w.sRow = sheetStRow;
                            w.ew = ew;

                            foreach (var t in linqWest)
                            {
                                w.地域コード = (int)t.地域コード;
                                w.地域名 = t.地域名;
                                w.cardNo = t.cardno;
                                w.氏名 = t.氏名;
                                w.携帯電話番号 = t.携帯電話番号;
                                w.JFG加入年 = (short)t.JFG加入年;
                                w.会員稼働予定 = t.会員稼働予定;

                                // エクセル稼働表作成
                                if (!xlsCellsSet(w, ref oxlsWorkSheet)) continue;
                            }
                        }

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

                        //// 出力シートコンソール表示
                        //Console.WriteLine(oxlsWorkSheet.Name);                    
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
        ///     稼働表作成 : closedXML版 2018/02/22</summary>
        ///----------------------------------------------------
        public void worksOutputXML()
        {
            DateTime stDate;
            DateTime edDate;

            //IXLWorksheet tmpSheet = null;

            try
            {
                using (var book = new XLWorkbook(Properties.Settings.Default.xlsKadouPath, XLEventTracking.Disabled))
                {
                    // 稼働予定開始年月日
                    stDate = DateTime.Parse(DateTime.Today.Year.ToString() + "/" + DateTime.Today.Month.ToString() + "/01");

                    // 稼働予定終了年月日
                    edDate = stDate.AddMonths(6).AddDays(-1);

                    int ew = 0;

                    while (ew < sCnt)
                    {
                        // 言語別シートを作成
                        for (int i = 1; i <= gengo.GetLength(0); i++)
                        {
                            int sNum = i + (10 * ew);

                            // シートを追加する 2018/02/22
                            if (ew == cEAST)
                            {
                                book.Worksheet("東").CopyTo(book, sheetName[ew] + "・" + gengo[i - 1, 1], sNum);
                            }
                            else if (ew == cWEST)
                            {
                                book.Worksheet("西").CopyTo(book, sheetName[ew] + "・" + gengo[i - 1, 1], sNum);
                            }

                            // カレントシート
                            IXLWorksheet tmpSheet = book.Worksheet(sNum);

                            int xCol = 0;   // 日列初期値

                            // 稼働予定期間のカレンダーをセット
                            for (int mon = 0; mon < 6; mon++)
                            {
                                // 該当月
                                DateTime wDt = stDate.AddMonths(mon);
                                xCol = 31 * mon + ew + 9;
                                //tmpSheet.Cell(1, xCol).Value = wDt.ToShortDateString(); // 9,40,71,102,・・・ 2018/02/23
                                tmpSheet.Cell(1, xCol).SetValue(wDt.Year + "年" + wDt.Month + "月"); // 9,40,71,102,・・・ 2018/02/28

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
                                        // 2018/02/23 
                                        tmpSheet.Cell(2, xCol + dy).SetValue((dy + 1).ToString());    // 日
                                        tmpSheet.Cell(3, xCol + dy).SetValue(dDay.ToString("ddd"));   // 曜日
                                    }
                                    else
                                    {
                                        // 2018/02/23
                                        tmpSheet.Cell(2, xCol + dy).SetValue(string.Empty);
                                        tmpSheet.Cell(3, xCol + dy).SetValue(string.Empty);
                                    }

                                    dy++;
                                }
                            }

                            // 組合員予定申告データを取得
                            string cardNum = string.Empty;
                            string gCode = gengo[i - 1, 0];
                            int sRow = sheetStRow;

                            jfgDataClassDataContext db = new jfgDataClassDataContext();

                            // 東・LINQ
                            var linqEast = db.会員情報.Where(a => (a.言語1 == int.Parse(gCode) || a.言語2 == int.Parse(gCode) ||
                                                           a.言語3 == int.Parse(gCode) || a.言語4 == int.Parse(gCode) ||
                                                           a.言語5 == int.Parse(gCode)) && a.東西 == 1)
                                                 .OrderBy(a => a.会員稼働予定.フリガナ).ThenBy(a => a.会員稼働予定.カード番号).ThenBy(a => a.会員稼働予定.年).ThenBy(a => a.会員稼働予定.月)
                                                 .Select(a => new
                                                 {
                                                     cardno = a.カード番号,
                                                     氏名 = a.氏名,
                                                     携帯電話番号 = a.携帯電話番号,
                                                     JFG加入年 = a.JFG加入年,
                                                     a.会員稼働予定
                                                 });

                            // 西・LINQ
                            var linqWest = db.会員情報.Where(a => (a.言語1 == int.Parse(gCode) || a.言語2 == int.Parse(gCode) ||
                                                           a.言語3 == int.Parse(gCode) || a.言語4 == int.Parse(gCode) ||
                                                           a.言語5 == int.Parse(gCode)) && a.東西 == 2)
                                                 .OrderBy(a => a.地域コード).ThenBy(a => a.会員稼働予定.フリガナ).ThenBy(a => a.カード番号).ThenBy(a => a.会員稼働予定.年).ThenBy(a => a.会員稼働予定.月)
                                                 .Select(a => new
                                                 {
                                                     地域コード = a.地域コード,
                                                     地域名 = a.地域名,
                                                     cardno = a.カード番号,
                                                     氏名 = a.氏名,
                                                     携帯電話番号 = a.携帯電話番号,
                                                     JFG加入年 = a.JFG加入年,
                                                     a.会員稼働予定
                                                 });

                            if (ew == cEAST)    // 東
                            {
                                // 組合員予定申告データクラスのインスタンス生成
                                clsWorksTbl w = new clsWorksTbl();
                                w.cardNumBox = string.Empty;
                                w.sRow = sheetStRow;
                                w.ew = ew;

                                foreach (var t in linqEast)
                                {
                                    w.cardNo = t.cardno;
                                    w.氏名 = t.氏名;
                                    w.携帯電話番号 = t.携帯電話番号;
                                    w.JFG加入年 = (short)t.JFG加入年;
                                    w.会員稼働予定 = t.会員稼働予定;

                                    // エクセル稼働表作成 2018/02/26
                                    if (!xlsCellsSetXML(w, tmpSheet))
                                    {
                                        continue;
                                    }

                                }
                            }
                            else if (ew == cWEST)   // 西
                            {
                                // 組合員予定申告データクラスのインスタンス生成
                                clsWorksTbl w = new clsWorksTbl();
                                w.cardNumBox = string.Empty;
                                w.sRow = sheetStRow;
                                w.ew = ew;

                                foreach (var t in linqWest)
                                {
                                    w.地域コード = (int)t.地域コード;
                                    w.地域名 = t.地域名;
                                    w.cardNo = t.cardno;
                                    w.氏名 = t.氏名;
                                    w.携帯電話番号 = t.携帯電話番号;
                                    w.JFG加入年 = (short)t.JFG加入年;
                                    w.会員稼働予定 = t.会員稼働予定;

                                    // エクセル稼働表作成 2018/02/26
                                    if (!xlsCellsSetXML(w, tmpSheet))
                                    {
                                        continue;
                                    }
                                }
                            }

                            // カレンダーにない日の列削除
                            bool colDelStatus = true;

                            // 2018/02/26
                            while (colDelStatus)
                            {
                                for (int cl = (ew + 9); cl <= tmpSheet.RangeUsed().RangeAddress.LastAddress.ColumnNumber; cl++)
                                {
                                    if (!Utility.NumericCheck(Utility.nulltoString(tmpSheet.Cell(2, cl).Value).Trim()))
                                    {
                                        tmpSheet.Column(cl).Delete();  // 2018/02/26
                                        colDelStatus = true;
                                        break;
                                    }
                                    else
                                    {
                                        colDelStatus = false;
                                    }
                                }
                            }

                            // 年月を表すセルを結合する 2018/02/28
                            int stCell = 0;
                            int edCell = 0;

                            tmpSheet.Range(tmpSheet.Cell(1, ew + 9).Address,
                                           tmpSheet.Cell(1, tmpSheet.LastCellUsed().Address.ColumnNumber).Address).Style
                                           .Border.BottomBorder = XLBorderStyleValues.Thin;
                            for (int cl = (ew + 9); cl <= tmpSheet.LastCellUsed().Address.ColumnNumber; cl++)
                            {
                                if (Utility.nulltoString(tmpSheet.Cell(2, cl).Value).Trim() == "1")
                                {
                                    if (stCell == 0)
                                    {
                                        stCell = cl;
                                    }
                                    else
                                    {
                                        // セル結合
                                        tmpSheet.Range(tmpSheet.Cell(1, stCell).Address, tmpSheet.Cell(1, edCell).Address).Merge(false);

                                        // IsMerge()パフォ劣化回避のためのStyle変更
                                        for (int cc = stCell; cc <= edCell; cc++)
                                        {
                                            tmpSheet.Cell(1, cc).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                        }

                                        stCell = cl;
                                    }
                                }
                                else
                                {
                                    edCell = cl;
                                }
                            }

                            if (stCell != 0)
                            {
                                // セル結合
                                tmpSheet.Range(tmpSheet.Cell(1, stCell).Address, tmpSheet.Cell(1, edCell).Address).Merge(false);

                                // IsMerge()パフォ劣化回避のためのStyle変更
                                for (int cc = stCell; cc <= edCell; cc++)
                                {
                                    tmpSheet.Cell(1, cc).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                }
                            }

                            // 表の外枠罫線を引く 2018/02/26
                            var range = tmpSheet.Range(tmpSheet.Cell("A1").Address, tmpSheet.LastCellUsed().Address);
                            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                            // 年月セル下部に罫線を引く 2018/02/28
                            tmpSheet.Range(tmpSheet.Cell(1, ew + 9).Address,
                                           tmpSheet.Cell(1, tmpSheet.LastCellUsed().Address.ColumnNumber).Address).Style
                                           .Border.BottomBorder = XLBorderStyleValues.Thin;

                            tmpSheet.Range(tmpSheet.Cell(2, ew + 9).Address,
                                           tmpSheet.Cell(2, tmpSheet.LastCellUsed().Address.ColumnNumber).Address).Style
                                           .Border.BottomBorder = XLBorderStyleValues.Dotted;

                            // 明細最上部に罫線を引く 2018/02/27
                            tmpSheet.Range(tmpSheet.Cell("A4").Address,
                                           tmpSheet.Cell(4, tmpSheet.LastCellUsed().Address.ColumnNumber).Address).Style
                                           .Border.TopBorder = XLBorderStyleValues.Thin;

                            // 表の外枠左罫線を引く 2018/02/27
                            tmpSheet.Range(tmpSheet.Cell("A1").Address, tmpSheet.LastCellUsed().Address).Style
                                .Border.LeftBorder = XLBorderStyleValues.Thin;

                            // 見出しの背景色 2018/02/28
                            tmpSheet.Range(tmpSheet.Cell("A1").Address, tmpSheet.Cell(3, tmpSheet.LastCellUsed().Address.ColumnNumber).Address)
                                .Style.Fill.BackgroundColor = XLColor.WhiteSmoke;

                            // 日曜日の背景色
                            range = tmpSheet.Range(tmpSheet.Cell(3, ew + 9).Address, tmpSheet.Cell(3, tmpSheet.LastCellUsed().Address.ColumnNumber).Address);
                            range.AddConditionalFormat()
                                 .WhenEquals("日")
                                 .Fill.SetBackgroundColor(XLColor.MistyRose);

                            var range2 = tmpSheet.Range(tmpSheet.Cell(2, ew + 9).Address, tmpSheet.Cell(2, tmpSheet.LastCellUsed().Address.ColumnNumber).Address);

                            if (ew == cEAST)
                            {
                                // 日曜日の日付の背景色
                                range2.AddConditionalFormat()
                                      .WhenIsTrue("=I3=" + @"""日""")
                                      .Fill.BackgroundColor = XLColor.MistyRose;

                                // ウィンドウ枠の固定
                                tmpSheet.SheetView.Freeze(3, 2);

                                // 見出し
                                tmpSheet.Cell("A2").SetValue("氏名").Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                tmpSheet.Cell("B2").SetValue("フリガナ").Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                tmpSheet.Cell("C2").SetValue("入会年度").Style.Font.SetBold(true);
                                tmpSheet.Cell("D2").SetValue("携帯電話").Style.Font.SetBold(true);
                                tmpSheet.Cell("E2").SetValue("稼働日数").Style.Font.SetBold(true);
                                tmpSheet.Cell("F2").SetValue("自己申告").Style.Font.SetBold(true);
                                tmpSheet.Cell("F3").SetValue("日数").Style.Font.SetBold(true);
                                tmpSheet.Cell("G2").SetValue("備考").Style.Font.SetBold(true);
                                tmpSheet.Cell("H2").SetValue("更新日").Style.Font.SetBold(true);

                                // 見出しはBold 2018/02/28
                                tmpSheet.Range(tmpSheet.Cell("I1").Address, tmpSheet.Cell(3, tmpSheet.LastCellUsed().Address.ColumnNumber).Address)
                                    .Style.Font.SetBold(true);
                            }
                            else if (ew == cWEST)
                            {
                                // 日曜日の日付の背景色
                                range2.AddConditionalFormat()
                                      .WhenIsTrue("=J3=" + @"""日""")
                                      .Fill.BackgroundColor = XLColor.MistyRose;

                                // ウィンドウ枠の固定
                                tmpSheet.SheetView.Freeze(3, 3);

                                // 見出し
                                tmpSheet.Cell("A2").SetValue("地域").Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                tmpSheet.Cell("B2").SetValue("氏名").Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                tmpSheet.Cell("C2").SetValue("フリガナ").Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                tmpSheet.Cell("D2").SetValue("入会年度").Style.Font.SetBold(true);
                                tmpSheet.Cell("E2").SetValue("携帯電話").Style.Font.SetBold(true);
                                tmpSheet.Cell("F2").SetValue("稼働日数").Style.Font.SetBold(true);
                                tmpSheet.Cell("G2").SetValue("自己申告").Style.Font.SetBold(true);
                                tmpSheet.Cell("G3").SetValue("日数").Style.Font.SetBold(true);
                                tmpSheet.Cell("H2").SetValue("備考").Style.Font.SetBold(true);
                                tmpSheet.Cell("I2").SetValue("更新日").Style.Font.SetBold(true);

                                // 見出しはBold 2018/02/28
                                tmpSheet.Range(tmpSheet.Cell("J1").Address, tmpSheet.Cell(3, tmpSheet.LastCellUsed().Address.ColumnNumber).Address)
                                    .Style.Font.SetBold(true);
                            }
                        }

                        ew++;
                    }

                    // テンプレートシートは削除する 2018/02/26
                    book.Worksheet("東").Delete();
                    book.Worksheet("西").Delete();

                    // カレントシート 2018/02/26
                    //tmpSheet = book.Worksheet(1);

                    //保存処理 2018/02/26
                    book.SaveAs(Properties.Settings.Default.xlsWorksPath);
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
            }
            finally
            {

            }
        }


        /// <summary>
        /// Excelシート名簿稼働表作成 : closedXML版 2022/11/07
        /// </summary>
        public void worksOutputXML_FromExcel()
        {
            DateTime stDate;
            DateTime edDate;

            try
            {
                // Excelガイドリストをテーブルに読み込む
                IXLTable tbl;
                using (var selectBook = new XLWorkbook(Properties.Settings.Default.xlsHotelGuideListPath))
                using (var selSheet = selectBook.Worksheet(1))
                {
                    // カード番号開始セル
                    var cell1 = selSheet.Cell("A4");
                    // 最終行を取得
                    var lastRow = selSheet.LastRowUsed().RowNumber();
                    // カード番号最終セル
                    var cell2 = selSheet.Cell(lastRow, 1);
                    // カード番号をテーブルで取得
                    tbl = selSheet.Range(cell1, cell2).AsTable();
                }

                // テーブル有効行がないときは終わる
                if (tbl.RowCount() < 1)
                {
                    return;
                }

                using (var book = new XLWorkbook(Properties.Settings.Default.xlsKadouPath, XLEventTracking.Disabled))
                {
                    // 稼働予定開始年月日
                    stDate = DateTime.Parse(DateTime.Today.Year.ToString() + "/" + DateTime.Today.Month.ToString() + "/01");

                    // 稼働予定終了年月日
                    edDate = stDate.AddMonths(6).AddDays(-1);

                    int ew = 0;

                    // シートを追加する
                    book.Worksheet("東").CopyTo(book, sheetName[ew] + "・" + gengo[0, 1], 1);

                    // カレントシート
                    IXLWorksheet tmpSheet = book.Worksheet(1);

                    int xCol = 0;   // 日列初期値

                    // 稼働予定期間のカレンダーをセット
                    for (int mon = 0; mon < 6; mon++)
                    {
                        // 該当月
                        DateTime wDt = stDate.AddMonths(mon);
                        xCol = 31 * mon + ew + 9;
                        tmpSheet.Cell(1, xCol).SetValue(wDt.Year + "年" + wDt.Month + "月"); // 9,40,71,102,・・・ 

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
                                tmpSheet.Cell(2, xCol + dy).SetValue((dy + 1).ToString());    // 日
                                tmpSheet.Cell(3, xCol + dy).SetValue(dDay.ToString("ddd"));   // 曜日
                            }
                            else
                            {
                                tmpSheet.Cell(2, xCol + dy).SetValue(string.Empty);
                                tmpSheet.Cell(3, xCol + dy).SetValue(string.Empty);
                            }

                            dy++;
                        }
                    }

                    // 組合員予定申告データを取得
                    string cardNum = string.Empty;
                    string gCode   = gengo[0, 0];
                    int    sRow    = sheetStRow;

                    jfgDataClassDataContext db = new jfgDataClassDataContext();

                    // 東・LINQ
                    var linqEast = db.会員情報.Where(a => (a.言語1 == int.Parse(gCode) || a.言語2 == int.Parse(gCode) ||
                                                          a.言語3 == int.Parse(gCode) || a.言語4 == int.Parse(gCode) ||
                                                          a.言語5 == int.Parse(gCode)) && a.東西 == 1)
                                         .OrderBy(a => a.会員稼働予定.フリガナ).ThenBy(a => a.会員稼働予定.カード番号).ThenBy(a => a.会員稼働予定.年).ThenBy(a => a.会員稼働予定.月)
                                         .Select(a => new
                                         {
                                             cardno = a.カード番号,
                                             氏名 = a.氏名,
                                             携帯電話番号 = a.携帯電話番号,
                                             JFG加入年 = a.JFG加入年,
                                             a.会員稼働予定
                                         });

                    // 組合員予定申告データクラスのインスタンス生成
                    clsWorksTbl w = new clsWorksTbl();
                    w.cardNumBox  = string.Empty;
                    w.sRow        = sheetStRow;
                    w.ew          = ew;

                    foreach (var t in linqEast)
                    {
                        bool listMember = false;

                        // ホテル向けガイドリスト(英語)を参照
                        foreach (var row in tbl.Rows())
                        {
                            var card = row.Cell(1).Value;
                            if (string.IsNullOrEmpty(card.ToString()))
                            {
                                continue;
                            }

                            if (card.ToString() == t.cardno.ToString())
                            {
                                listMember = true;
                                break;
                            }
                        }

                        // ホテル向けガイドリスト(英語)対象以外はネグる
                        if (!listMember)
                        {
                            continue;
                        }

                        w.cardNo = t.cardno;
                        w.氏名 = t.氏名;
                        w.携帯電話番号 = t.携帯電話番号;
                        w.JFG加入年 = (short)t.JFG加入年;
                        w.会員稼働予定 = t.会員稼働予定;

                        // エクセル稼働表作成
                        if (!xlsCellsSetXML(w, tmpSheet))
                        {
                            continue;
                        }
                    }

                    // カレンダーにない日の列削除
                    bool colDelStatus = true;

                    while (colDelStatus)
                    {
                        for (int cl = (ew + 9); cl <= tmpSheet.RangeUsed().RangeAddress.LastAddress.ColumnNumber; cl++)
                        {
                            if (!Utility.NumericCheck(Utility.nulltoString(tmpSheet.Cell(2, cl).Value).Trim()))
                            {
                                tmpSheet.Column(cl).Delete();
                                colDelStatus = true;
                                break;
                            }
                            else
                            {
                                colDelStatus = false;
                            }
                        }
                    }

                    // 年月を表すセルを結合する
                    int stCell = 0;
                    int edCell = 0;

                    tmpSheet.Range(tmpSheet.Cell(1, ew + 9).Address,
                                   tmpSheet.Cell(1, tmpSheet.LastCellUsed().Address.ColumnNumber).Address).Style
                                   .Border.BottomBorder = XLBorderStyleValues.Thin;
                    for (int cl = (ew + 9); cl <= tmpSheet.LastCellUsed().Address.ColumnNumber; cl++)
                    {
                        if (Utility.nulltoString(tmpSheet.Cell(2, cl).Value).Trim() == "1")
                        {
                            if (stCell == 0)
                            {
                                stCell = cl;
                            }
                            else
                            {
                                // セル結合
                                tmpSheet.Range(tmpSheet.Cell(1, stCell).Address, tmpSheet.Cell(1, edCell).Address).Merge(false);

                                // IsMerge()パフォ劣化回避のためのStyle変更
                                for (int cc = stCell; cc <= edCell; cc++)
                                {
                                    tmpSheet.Cell(1, cc).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                }

                                stCell = cl;
                            }
                        }
                        else
                        {
                            edCell = cl;
                        }
                    }

                    if (stCell != 0)
                    {
                        // セル結合
                        tmpSheet.Range(tmpSheet.Cell(1, stCell).Address, tmpSheet.Cell(1, edCell).Address).Merge(false);

                        // IsMerge()パフォ劣化回避のためのStyle変更
                        for (int cc = stCell; cc <= edCell; cc++)
                        {
                            tmpSheet.Cell(1, cc).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        }
                    }

                    // 表の外枠罫線を引く
                    var range = tmpSheet.Range(tmpSheet.Cell("A1").Address, tmpSheet.LastCellUsed().Address);
                    range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    // 年月セル下部に罫線を引く
                    tmpSheet.Range(tmpSheet.Cell(1, ew + 9).Address,
                                   tmpSheet.Cell(1, tmpSheet.LastCellUsed().Address.ColumnNumber).Address).Style
                                   .Border.BottomBorder = XLBorderStyleValues.Thin;

                    tmpSheet.Range(tmpSheet.Cell(2, ew + 9).Address,
                                   tmpSheet.Cell(2, tmpSheet.LastCellUsed().Address.ColumnNumber).Address).Style
                                   .Border.BottomBorder = XLBorderStyleValues.Dotted;

                    // 明細最上部に罫線を引く
                    tmpSheet.Range(tmpSheet.Cell("A4").Address,
                                   tmpSheet.Cell(4, tmpSheet.LastCellUsed().Address.ColumnNumber).Address).Style
                                   .Border.TopBorder = XLBorderStyleValues.Thin;

                    // 表の外枠左罫線を引く
                    tmpSheet.Range(tmpSheet.Cell("A1").Address, tmpSheet.LastCellUsed().Address).Style
                        .Border.LeftBorder = XLBorderStyleValues.Thin;

                    // 見出しの背景色 
                    tmpSheet.Range(tmpSheet.Cell("A1").Address, tmpSheet.Cell(3, tmpSheet.LastCellUsed().Address.ColumnNumber).Address)
                        .Style.Fill.BackgroundColor = XLColor.WhiteSmoke;

                    // 日曜日の背景色
                    range = tmpSheet.Range(tmpSheet.Cell(3, ew + 9).Address, tmpSheet.Cell(3, tmpSheet.LastCellUsed().Address.ColumnNumber).Address);
                    range.AddConditionalFormat()
                         .WhenEquals("日")
                         .Fill.SetBackgroundColor(XLColor.MistyRose);

                    var range2 = tmpSheet.Range(tmpSheet.Cell(2, ew + 9).Address, tmpSheet.Cell(2, tmpSheet.LastCellUsed().Address.ColumnNumber).Address);

                    if (ew == cEAST)
                    {
                        // 日曜日の日付の背景色
                        range2.AddConditionalFormat()
                              .WhenIsTrue("=I3=" + @"""日""")
                              .Fill.BackgroundColor = XLColor.MistyRose;

                        // ウィンドウ枠の固定
                        tmpSheet.SheetView.Freeze(3, 2);

                        // 見出し
                        tmpSheet.Cell("A2").SetValue("氏名").Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        tmpSheet.Cell("B2").SetValue("フリガナ").Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        tmpSheet.Cell("C2").SetValue("入会年度").Style.Font.SetBold(true);
                        tmpSheet.Cell("D2").SetValue("携帯電話").Style.Font.SetBold(true);
                        tmpSheet.Cell("E2").SetValue("稼働日数").Style.Font.SetBold(true);
                        tmpSheet.Cell("F2").SetValue("自己申告").Style.Font.SetBold(true);
                        tmpSheet.Cell("F3").SetValue("日数").Style.Font.SetBold(true);
                        tmpSheet.Cell("G2").SetValue("備考").Style.Font.SetBold(true);
                        tmpSheet.Cell("H2").SetValue("更新日").Style.Font.SetBold(true);

                        // 見出しはBold
                        tmpSheet.Range(tmpSheet.Cell("I1").Address, tmpSheet.Cell(3, tmpSheet.LastCellUsed().Address.ColumnNumber).Address)
                            .Style.Font.SetBold(true);
                    }
                    else if (ew == cWEST)
                    {
                        // 日曜日の日付の背景色
                        range2.AddConditionalFormat()
                              .WhenIsTrue("=J3=" + @"""日""")
                              .Fill.BackgroundColor = XLColor.MistyRose;

                        // ウィンドウ枠の固定
                        tmpSheet.SheetView.Freeze(3, 3);

                        // 見出し
                        tmpSheet.Cell("A2").SetValue("地域").Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        tmpSheet.Cell("B2").SetValue("氏名").Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        tmpSheet.Cell("C2").SetValue("フリガナ").Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        tmpSheet.Cell("D2").SetValue("入会年度").Style.Font.SetBold(true);
                        tmpSheet.Cell("E2").SetValue("携帯電話").Style.Font.SetBold(true);
                        tmpSheet.Cell("F2").SetValue("稼働日数").Style.Font.SetBold(true);
                        tmpSheet.Cell("G2").SetValue("自己申告").Style.Font.SetBold(true);
                        tmpSheet.Cell("G3").SetValue("日数").Style.Font.SetBold(true);
                        tmpSheet.Cell("H2").SetValue("備考").Style.Font.SetBold(true);
                        tmpSheet.Cell("I2").SetValue("更新日").Style.Font.SetBold(true);

                        // 見出しはBold
                        tmpSheet.Range(tmpSheet.Cell("J1").Address, tmpSheet.Cell(3, tmpSheet.LastCellUsed().Address.ColumnNumber).Address)
                            .Style.Font.SetBold(true);
                    }

                    // テンプレートシートは削除する
                    book.Worksheet("東").Delete();

                    //保存処理
                    book.SaveAs(Properties.Settings.Default.xlsHotelsWorksPath);
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
            }
            finally
            {

            }
        }

        /// --------------------------------------------------------------------------
        /// <summary>
        ///     アサイン担当用稼働表エクセルシート作成 </summary>
        /// <param name="t">
        ///     組合員予定申告データクラス </param>
        /// <param name="oxlsWorkSheet">
        ///     アサイン担当用稼働表エクセルシート </param>
        /// <returns>
        ///     作成：true, 未作成：false</returns>
        /// --------------------------------------------------------------------------
        private bool xlsCellsSet(clsWorksTbl t, ref Excel.Worksheet oxlsWorkSheet)
        {
            Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

            // 該当期間のデータか検証
            string yymm = t.会員稼働予定.年.ToString() + t.会員稼働予定.月.ToString().PadLeft(2, '0');

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

            if (!yymmOn) return false; // 非該当期間のため読み飛ばし

            // 組合員が変わったら行番号を加算する
            if (t.cardNumBox != string.Empty && t.cardNumBox != t.cardNo.ToString())
            {
                //セル下部へ実線ヨコ罫線を引く
                rng[0] = (Excel.Range)oxlsWorkSheet.Cells[t.sRow, 1];
                rng[1] = (Excel.Range)oxlsWorkSheet.Cells[t.sRow, S_colSMAX[t.ew]];
                oxlsWorkSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot;

                t.sRow++;
            }

            t.cardNumBox = t.cardNo.ToString();    // カード番号

            // 作業用シートにデータ貼り付け
            if (t.ew == cEAST)
            {
                oxlsWorkSheet.Cells[t.sRow, 1] = t.氏名.ToString();
                oxlsWorkSheet.Cells[t.sRow, 2] = t.会員稼働予定.フリガナ.ToString();
                oxlsWorkSheet.Cells[t.sRow, 3] = t.JFG加入年.ToString();
                oxlsWorkSheet.Cells[t.sRow, 4] = t.携帯電話番号.ToString();
                oxlsWorkSheet.Cells[t.sRow, 5] = t.会員稼働予定.稼働日数.ToString();
                oxlsWorkSheet.Cells[t.sRow, 6] = t.会員稼働予定.自己申告日数.ToString();
                oxlsWorkSheet.Cells[t.sRow, 7] = t.会員稼働予定.備考.ToString();
                oxlsWorkSheet.Cells[t.sRow, 8] = Utility.nulltoString(t.会員稼働予定.申告年月日);
            }
            else if (t.ew == cWEST)
            {
                oxlsWorkSheet.Cells[t.sRow, 1] = t.地域名;
                oxlsWorkSheet.Cells[t.sRow, 2] = t.氏名.ToString();
                oxlsWorkSheet.Cells[t.sRow, 3] = t.会員稼働予定.フリガナ.ToString();
                oxlsWorkSheet.Cells[t.sRow, 4] = t.JFG加入年.ToString();
                oxlsWorkSheet.Cells[t.sRow, 5] = t.携帯電話番号.ToString();
                oxlsWorkSheet.Cells[t.sRow, 6] = t.会員稼働予定.稼働日数.ToString();
                oxlsWorkSheet.Cells[t.sRow, 7] = t.会員稼働予定.自己申告日数.ToString();
                oxlsWorkSheet.Cells[t.sRow, 8] = t.会員稼働予定.備考.ToString();
                oxlsWorkSheet.Cells[t.sRow, 9] = Utility.nulltoString(t.会員稼働予定.申告年月日);
            }

            // アサイン担当者か検証する
            jfgDataClassDataContext db = new jfgDataClassDataContext();
            var s = db.アサイン担当者.Where(a => a.カード番号 == double.Parse(t.cardNumBox));
            if (s.Count() > 0)
            {
                // 氏名セルのBackColorを黄色にする
                if (t.ew == cEAST) // 東
                {
                    rng[0] = (Excel.Range)oxlsWorkSheet.Cells[t.sRow, 1];
                    rng[1] = (Excel.Range)oxlsWorkSheet.Cells[t.sRow, 2];
                }
                else if (t.ew == cWEST)　// 西
                {
                    rng[0] = (Excel.Range)oxlsWorkSheet.Cells[t.sRow, 1];
                    rng[1] = (Excel.Range)oxlsWorkSheet.Cells[t.sRow, 3];
                }

                oxlsWorkSheet.get_Range(rng[0], rng[1]).Interior.Color = Color.Yellow;
            }

            // 予定申告内容をセルに貼り付ける
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col)] = t.会員稼働予定.d1;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 1] = t.会員稼働予定.d2;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 2] = t.会員稼働予定.d3;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 3] = t.会員稼働予定.d4;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 4] = t.会員稼働予定.d5;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 5] = t.会員稼働予定.d6;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 6] = t.会員稼働予定.d7;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 7] = t.会員稼働予定.d8;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 8] = t.会員稼働予定.d9;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 9] = t.会員稼働予定.d10;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 10] = t.会員稼働予定.d11;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 11] = t.会員稼働予定.d12;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 12] = t.会員稼働予定.d13;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 13] = t.会員稼働予定.d14;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 14] = t.会員稼働予定.d15;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 15] = t.会員稼働予定.d16;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 16] = t.会員稼働予定.d17;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 17] = t.会員稼働予定.d18;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 18] = t.会員稼働予定.d19;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 19] = t.会員稼働予定.d20;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 20] = t.会員稼働予定.d21;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 21] = t.会員稼働予定.d22;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 22] = t.会員稼働予定.d23;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 23] = t.会員稼働予定.d24;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 24] = t.会員稼働予定.d25;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 25] = t.会員稼働予定.d26;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 26] = t.会員稼働予定.d27;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 27] = t.会員稼働予定.d28;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 28] = t.会員稼働予定.d29;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 29] = t.会員稼働予定.d30;
            oxlsWorkSheet.Cells[t.sRow, int.Parse(col) + 30] = t.会員稼働予定.d31;

            return true;
        }


        /// --------------------------------------------------------------------------
        /// <summary>
        ///     アサイン担当用稼働表エクセルシート作成 
        ///     : closedXML版 2018/02/22</summary>
        /// <param name="t">
        ///     組合員予定申告データクラス </param>
        /// <param name="oxlsWorkSheet">
        ///     アサイン担当用稼働表エクセルシート </param>
        /// <returns>
        ///     作成：true, 未作成：false</returns>
        /// --------------------------------------------------------------------------
        private bool xlsCellsSetXML(clsWorksTbl t, ClosedXML.Excel.IXLWorksheet sheet)
        {
            //Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

            // 該当期間のデータか検証
            string yymm = t.会員稼働予定.年.ToString() + t.会員稼働予定.月.ToString().PadLeft(2, '0');

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

            if (!yymmOn)
            {
                return false; // 非該当期間のため読み飛ばし
            }

            // 組合員が変わったら行番号を加算する
            if (t.cardNumBox != string.Empty && t.cardNumBox != t.cardNo.ToString())
            {
                //セル下部へ点線ヨコ罫線を引く 2018/02/28
                sheet.Range(sheet.Cell(t.sRow, 1).Address,
                            sheet.Cell(t.sRow, sheet.LastCellUsed().Address.ColumnNumber).Address).Style.Border.BottomBorder = XLBorderStyleValues.Dotted;

                //// 表の外枠右罫線を引く 2018/02/27
                //sheet.Range(sheet.Cell("A1"), sheet.Cell("E10")).Style.Border.RightBorder = XLBorderStyleValues.Thin;

                t.sRow++;
            }

            t.cardNumBox = t.cardNo.ToString();    // カード番号

            // 作業用シートにデータ貼り付け
            if (t.ew == cEAST)
            {
                sheet.Cell(t.sRow, 1).SetValue(t.氏名);
                sheet.Cell(t.sRow, 2).SetValue(t.会員稼働予定.フリガナ);
                sheet.Cell(t.sRow, 3).SetValue(t.JFG加入年.ToString());
                sheet.Cell(t.sRow, 4).SetValue(t.携帯電話番号);
                sheet.Cell(t.sRow, 5).SetValue(t.会員稼働予定.稼働日数.ToString());
                sheet.Cell(t.sRow, 6).SetValue(t.会員稼働予定.自己申告日数.ToString());
                sheet.Cell(t.sRow, 7).SetValue(t.会員稼働予定.備考);
                sheet.Cell(t.sRow, 8).SetValue(Utility.nulltoString(t.会員稼働予定.申告年月日));
            }
            else if (t.ew == cWEST)
            {
                sheet.Cell(t.sRow, 1).SetValue(t.地域名);
                sheet.Cell(t.sRow, 2).SetValue(t.氏名);
                sheet.Cell(t.sRow, 3).SetValue(t.会員稼働予定.フリガナ);
                sheet.Cell(t.sRow, 4).SetValue(t.JFG加入年.ToString());
                sheet.Cell(t.sRow, 5).SetValue(t.携帯電話番号);
                sheet.Cell(t.sRow, 6).SetValue(t.会員稼働予定.稼働日数.ToString());
                sheet.Cell(t.sRow, 7).SetValue(t.会員稼働予定.自己申告日数.ToString());
                sheet.Cell(t.sRow, 8).SetValue(t.会員稼働予定.備考);
                sheet.Cell(t.sRow, 9).SetValue(Utility.nulltoString(t.会員稼働予定.申告年月日));
            }

            // アサイン担当者か検証する
            jfgDataClassDataContext db = new jfgDataClassDataContext();
            var s = db.アサイン担当者.Where(a => a.カード番号 == double.Parse(t.cardNumBox));
            if (s.Count() > 0)
            {
                // 氏名セルのBackColorを黄色にする
                if (t.ew == cEAST) // 東
                {
                    // 2018/02/22
                    sheet.Cell(t.sRow, 1).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                    sheet.Cell(t.sRow, 2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                }
                else if (t.ew == cWEST)　// 西
                {
                    // 2018/02/22
                    sheet.Cell(t.sRow, 1).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                    sheet.Cell(t.sRow, 2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                    sheet.Cell(t.sRow, 3).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                }
            }

            // 予定申告内容をセルに貼り付ける
            sheet.Cell(t.sRow, int.Parse(col)).SetValue(t.会員稼働予定.d1);
            sheet.Cell(t.sRow, int.Parse(col) + 1).SetValue(t.会員稼働予定.d2);
            sheet.Cell(t.sRow, int.Parse(col) + 2).SetValue(t.会員稼働予定.d3);
            sheet.Cell(t.sRow, int.Parse(col) + 3).SetValue(t.会員稼働予定.d4);
            sheet.Cell(t.sRow, int.Parse(col) + 4).SetValue(t.会員稼働予定.d5);
            sheet.Cell(t.sRow, int.Parse(col) + 5).SetValue(t.会員稼働予定.d6);
            sheet.Cell(t.sRow, int.Parse(col) + 6).SetValue(t.会員稼働予定.d7);
            sheet.Cell(t.sRow, int.Parse(col) + 7).SetValue(t.会員稼働予定.d8);
            sheet.Cell(t.sRow, int.Parse(col) + 8).SetValue(t.会員稼働予定.d9);
            sheet.Cell(t.sRow, int.Parse(col) + 9).SetValue(t.会員稼働予定.d10);
            sheet.Cell(t.sRow, int.Parse(col) + 10).SetValue(t.会員稼働予定.d11);
            sheet.Cell(t.sRow, int.Parse(col) + 11).SetValue(t.会員稼働予定.d12);
            sheet.Cell(t.sRow, int.Parse(col) + 12).SetValue(t.会員稼働予定.d13);
            sheet.Cell(t.sRow, int.Parse(col) + 13).SetValue(t.会員稼働予定.d14);
            sheet.Cell(t.sRow, int.Parse(col) + 14).SetValue(t.会員稼働予定.d15);
            sheet.Cell(t.sRow, int.Parse(col) + 15).SetValue(t.会員稼働予定.d16);
            sheet.Cell(t.sRow, int.Parse(col) + 16).SetValue(t.会員稼働予定.d17);
            sheet.Cell(t.sRow, int.Parse(col) + 17).SetValue(t.会員稼働予定.d18);
            sheet.Cell(t.sRow, int.Parse(col) + 18).SetValue(t.会員稼働予定.d19);
            sheet.Cell(t.sRow, int.Parse(col) + 19).SetValue(t.会員稼働予定.d20);
            sheet.Cell(t.sRow, int.Parse(col) + 20).SetValue(t.会員稼働予定.d21);
            sheet.Cell(t.sRow, int.Parse(col) + 21).SetValue(t.会員稼働予定.d22);
            sheet.Cell(t.sRow, int.Parse(col) + 22).SetValue(t.会員稼働予定.d23);
            sheet.Cell(t.sRow, int.Parse(col) + 23).SetValue(t.会員稼働予定.d24);
            sheet.Cell(t.sRow, int.Parse(col) + 24).SetValue(t.会員稼働予定.d25);
            sheet.Cell(t.sRow, int.Parse(col) + 25).SetValue(t.会員稼働予定.d26);
            sheet.Cell(t.sRow, int.Parse(col) + 26).SetValue(t.会員稼働予定.d27);
            sheet.Cell(t.sRow, int.Parse(col) + 27).SetValue(t.会員稼働予定.d28);
            sheet.Cell(t.sRow, int.Parse(col) + 28).SetValue(t.会員稼働予定.d29);
            sheet.Cell(t.sRow, int.Parse(col) + 29).SetValue(t.会員稼働予定.d30);
            sheet.Cell(t.sRow, int.Parse(col) + 30).SetValue(t.会員稼働予定.d31);

            return true;
        }

        ///----------------------------------------------------
        /// <summary>
        ///     言語配列作成　</summary>
        ///----------------------------------------------------
        private void readLang()
        {
            jfgDataClassDataContext db = new jfgDataClassDataContext();
            var s = db.言語.Where(a => a.言語名1 != "J").OrderBy(a => a.言語番号);

            gengo = new string[s.Count(), 2];

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
