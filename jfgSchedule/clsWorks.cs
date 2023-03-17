using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Data.OleDb;
using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using System.Reflection;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Wordprocessing;

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
        const int xCol = 21;                           // 日列初期値
        readonly XLColor HeaderBackColor = XLColor.FromArgb(79, 129, 189);  // 見出し行背景色
        readonly XLColor LineBackColor = XLColor.FromArgb(220, 230, 241); // 奇数明細行背景色
        readonly string SheetName = "新ホテル向けガイド稼働表";

        public clsWorks(string logFile)
        {
            // 言語配列読み込み
            ReadLang();

            // 稼働表作成
            WorksOutputXML(logFile);

            // 新ホテル向けガイド稼働表作成
            //worksOutputXML_FromExcel(logFile);
            WorksOutputXML_FromExcel_BySheet(logFile);  // 新ホテル向けガイドリスト仕様 2023/02/17
        }

        /// <summary>
        /// 稼働表作成 : closedXML版 2018/02/22
        /// </summary>
        /// <param name="logFile">
        /// ログ出力パス
        /// </param>
        public void WorksOutputXML(string logFile)
        {
            DateTime stDate;
            DateTime edDate;

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

                    //保存処理 2018/02/26
                    book.SaveAs(Properties.Settings.Default.xlsWorksPath);
                }

                // ログ出力
                System.IO.File.AppendAllText(logFile, Form1.GetNowTime(" アサイン担当用稼働表を更新しました。"), Encoding.GetEncoding(932));
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
        /// <param name="logFile">
        /// ログ出力パス
        /// </param>
        public void WorksOutputXML_FromExcel(string logFile)
        {
            // ホテル向けガイドリストExcelファイルの存在確認
            if (!System.IO.File.Exists(Properties.Settings.Default.xlsHotelGuideListPath))
            {
                // ログ出力
                System.IO.File.AppendAllText(logFile, Form1.GetNowTime(" ホテル向けガイドリストExcelファイル（" + Properties.Settings.Default.xlsHotelGuideListPath + "）が見つかりませんでした。"), Encoding.GetEncoding(932));
                return;
            }

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

                // ガイドリストテーブル有効行がないときは終わる
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

                        // 該当月の暦
                        int dy = 0;
                        while (dy < 31)
                        {
                            if (DateTime.TryParse(wDt.Year.ToString() + "/" + wDt.Month.ToString() + "/" + (dy + 1).ToString(), out DateTime dDay))
                            {
                                if (dDay >= DateTime.Today)
                                {
                                    tmpSheet.Cell(1, xCol + dy).SetValue(wDt.Year + "年" + wDt.Month + "月");  // 年月：2023/01/26
                                    tmpSheet.Cell(2, xCol + dy).SetValue((dy + 1).ToString());    // 日
                                    tmpSheet.Cell(3, xCol + dy).SetValue(dDay.ToString("ddd"));   // 曜日
                                }
                                else
                                {
                                    // 作成前日以前はセルを空白とする：2023/01/25
                                    tmpSheet.Cell(2, xCol + dy).SetValue(string.Empty);
                                    tmpSheet.Cell(3, xCol + dy).SetValue(string.Empty);
                                }
                            }
                            else
                            {
                                // 存在しない日付はセルを空白とする
                                tmpSheet.Cell(2, xCol + dy).SetValue(string.Empty);
                                tmpSheet.Cell(3, xCol + dy).SetValue(string.Empty);
                            }

                            dy++;
                        }
                    }

                    // 組合員予定申告データを取得
                    string cardNum = string.Empty;
                    string gCode = gengo[0, 0];
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
                        // 2023/01/26 : 作成日以前の列が削除されたため2日以降で始まるケースあり
                        if (stCell == 0)
                        {
                            stCell = cl;
                        }

                        if (Utility.nulltoString(tmpSheet.Cell(2, cl).Value).Trim() == "1")
                        {
                            // 2023/01/26 コメント化
                            //if (stCell == 0)
                            //{
                            //    stCell = cl;
                            //}
                            //else
                            //{
                            //    // セル結合
                            //    tmpSheet.Range(tmpSheet.Cell(1, stCell).Address, tmpSheet.Cell(1, edCell).Address).Merge(false);

                            //    // IsMerge()パフォ劣化回避のためのStyle変更
                            //    for (int cc = stCell; cc <= edCell; cc++)
                            //    {
                            //        tmpSheet.Cell(1, cc).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            //    }

                            //    stCell = cl;
                            //}

                            // セル結合
                            tmpSheet.Range(tmpSheet.Cell(1, stCell).Address, tmpSheet.Cell(1, edCell).Address).Merge(false);

                            // IsMerge()パフォ劣化回避のためのStyle変更
                            for (int cc = stCell; cc <= edCell; cc++)
                            {
                                tmpSheet.Cell(1, cc).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }

                            stCell = cl;
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

                    // フィルタの設定：2023/1/25
                    tmpSheet.Row(3).SetAutoFilter();

                    // テンプレートシートは削除する
                    book.Worksheet("東").Delete();

                    //保存処理
                    book.SaveAs(Properties.Settings.Default.xlsHotelsWorksPath);
                }

                // ログ出力
                System.IO.File.AppendAllText(logFile, Form1.GetNowTime(" ホテル向けガイド稼働表を更新しました。"), Encoding.GetEncoding(932));
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
        /// 仕様簿稼働表作成： 2023/01/30
        ///     ・新ホテル向けガイドリストを基に作成
        ///     ・月をシート毎に分ける
        /// </summary>
        /// <param name="logFile">
        /// ログ出力パス
        /// </param>
        public void WorksOutputXML_FromExcel_BySheet(string logFile)
        {
            // ホテル向けガイドリストExcelファイルの存在確認
            if (!System.IO.File.Exists(Properties.Settings.Default.xlsHotelGuideListPath))
            {
                // ログ出力
                System.IO.File.AppendAllText(logFile, Form1.GetNowTime(" ホテル向けガイドリストExcelファイル（" + Properties.Settings.Default.xlsHotelGuideListPath + "）が見つかりませんでした。"), Encoding.GetEncoding(932));
                return;
            }

            DateTime stDate;
            DateTime edDate;
            string[] headerArray = new string[17];

            try
            {
                // Excelシート：新ホテル向けガイドリストをテーブルに読み込む
                IXLTable Hoteltbl;
                using (var selectBook = new XLWorkbook(Properties.Settings.Default.xlsNewHotelGuideListPath))
                {
                    using (var selSheet = selectBook.Worksheet(1))
                    {
                        // カード番号開始セル
                        var cell1 = selSheet.Cell("A2");
                        // 最終行を取得
                        var lastRow = selSheet.LastRowUsed().RowNumber();
                        // カード番号最終セル
                        var cell2 = selSheet.Cell(lastRow, 20);
                        // カード番号をテーブルで取得
                        Hoteltbl = selSheet.Range(cell1, cell2).AsTable();

                        // 列見出し文言を取得
                        headerArray[0]  = selSheet.Cell("D1").Value.ToString();
                        headerArray[1]  = selSheet.Cell("E1").Value.ToString();
                        headerArray[2]  = selSheet.Cell("F1").Value.ToString();
                        headerArray[3]  = selSheet.Cell("G1").Value.ToString();
                        headerArray[4]  = selSheet.Cell("H1").Value.ToString();
                        headerArray[5]  = selSheet.Cell("J1").Value.ToString();
                        headerArray[6]  = selSheet.Cell("K1").Value.ToString();
                        headerArray[7]  = selSheet.Cell("L1").Value.ToString();
                        headerArray[8]  = selSheet.Cell("M1").Value.ToString();
                        headerArray[9]  = selSheet.Cell("N1").Value.ToString().Replace(" ", "").Replace("　", "");
                        headerArray[10] = selSheet.Cell("O1").Value.ToString();
                        headerArray[11] = selSheet.Cell("P1").Value.ToString();
                        headerArray[12] = selSheet.Cell("Q1").Value.ToString();
                        headerArray[13] = selSheet.Cell("R1").Value.ToString();
                        headerArray[14] = selSheet.Cell("S1").Value.ToString();
                    }
                }

                // ガイドリストテーブル有効行がないときは終わる
                if (Hoteltbl.RowCount() < 1)
                {
                    // ログ出力
                    System.IO.File.AppendAllText(logFile, Form1.GetNowTime(" ホテル向けガイドリストExcelシートに有効行がありませんでした。"), Encoding.GetEncoding(932));
                    return;
                }

                // 稼働表ブック作成
                using (var book = new XLWorkbook(XLEventTracking.Disabled))
                {
                    // 稼働予定開始年月日
                    stDate = DateTime.Parse(DateTime.Today.Year.ToString() + "/" + DateTime.Today.Month.ToString() + "/01");

                    // 稼働予定終了年月日
                    edDate = stDate.AddMonths(6).AddDays(-1);

                    // シート作成
                    book.AddWorksheet(SheetName);
                    var tmpSheet = book.Worksheet(SheetName);

                    // 見出し 2023/02/08
                    tmpSheet.Cell("A1").SetValue("カード番号");
                    tmpSheet.Cell("B1").SetValue("氏名");
                    tmpSheet.Cell("C1").SetValue("フリガナ");
                    tmpSheet.Cell("D1").SetValue(headerArray[0]);
                    tmpSheet.Cell("E1").SetValue(headerArray[1]);
                    tmpSheet.Cell("F1").SetValue(headerArray[2]);
                    tmpSheet.Cell("G1").SetValue(headerArray[3]);
                    tmpSheet.Cell("H1").SetValue(headerArray[4]);
                    tmpSheet.Cell("I1").SetValue(headerArray[5]);
                    tmpSheet.Cell("J1").SetValue(headerArray[6]);
                    tmpSheet.Cell("K1").SetValue(headerArray[7]);
                    tmpSheet.Cell("L1").SetValue(headerArray[8]);
                    tmpSheet.Cell("M1").SetValue(headerArray[9]);
                    tmpSheet.Cell("N1").SetValue(headerArray[10]);
                    tmpSheet.Cell("O1").SetValue(headerArray[11]);
                    tmpSheet.Cell("P1").SetValue(headerArray[12]);
                    tmpSheet.Cell("Q1").SetValue(headerArray[13]);
                    tmpSheet.Cell("R1").SetValue(headerArray[14]);
                    tmpSheet.Cell("S1").SetValue("稼働日数");
                    tmpSheet.Cell("T1").SetValue("更新日");

                    // 稼働予定期間のカレンダーをセット
                    for (int mon = 0; mon < 6; mon++)
                    {
                        // 該当月
                        DateTime wDt = stDate.AddMonths(mon);
                        var xCol = 31 * mon + 21;
                        tmpSheet.Cell(1, xCol).SetValue(wDt.Year + "年" + wDt.Month + "月"); // 21,52,83,114,・・・ 

                        // 年月と開始列の配列にセット
                        sheetYYMM[mon, 0] = wDt.Year.ToString() + wDt.Month.ToString().PadLeft(2, '0');
                        sheetYYMM[mon, 1] = xCol.ToString();

                        // 該当月の暦
                        int dy = 0;
                        while (dy < 31)
                        {
                            if (DateTime.TryParse(wDt.Year.ToString() + "/" + wDt.Month.ToString() + "/" + (dy + 1).ToString(), out DateTime dDay))
                            {
                                if (dDay >= DateTime.Today)
                                {
                                    tmpSheet.Cell(1, xCol + dy).SetValue(wDt.Year + "年" + wDt.Month + "月");  // 年月：2023/01/26
                                    tmpSheet.Cell(2, xCol + dy).SetValue((dy + 1).ToString());    // 日
                                    tmpSheet.Cell(3, xCol + dy).SetValue(dDay.ToString("ddd"));   // 曜日
                                }
                                else
                                {
                                    // 作成前日以前はセルを空白とする：2023/01/25
                                    tmpSheet.Cell(2, xCol + dy).SetValue(string.Empty);
                                    tmpSheet.Cell(3, xCol + dy).SetValue(string.Empty);
                                }
                            }
                            else
                            {
                                // 存在しない日付はセルを空白とする
                                tmpSheet.Cell(2, xCol + dy).SetValue(string.Empty);
                                tmpSheet.Cell(3, xCol + dy).SetValue(string.Empty);
                            }

                            dy++;
                        }
                    }

                    // 組合員予定申告データを取得
                    string cardNum = string.Empty;
                    string gCode   = gengo[0, 0];

                    jfgDataClassDataContext db = new jfgDataClassDataContext();

                    // 東・LINQ
                    var linqEast = db.会員情報.Where(a => (a.言語1 == int.Parse(gCode) || a.言語2 == int.Parse(gCode) ||
                                                          a.言語3 == int.Parse(gCode) || a.言語4 == int.Parse(gCode) ||
                                                          a.言語5 == int.Parse(gCode)) && a.東西 == 1)
                                         .OrderBy(a => a.会員稼働予定.フリガナ).ThenBy(a => a.会員稼働予定.カード番号).ThenBy(a => a.会員稼働予定.年).ThenBy(a => a.会員稼働予定.月)
                                         .Select(a => new
                                         {
                                             a.カード番号,
                                             a.氏名,
                                             a.会員稼働予定
                                         });

                    ClsHotelScheduleXls clsHotel    = null;
                    ClsScheduleDays[]   clsSchedule = new ClsScheduleDays[31];

                    int col;

                    foreach (var t in linqEast)
                    {
                        col = 0;

                        // 該当期間のデータか検証
                        if (!IsTargetPeriod(t.会員稼働予定.年.ToString() + t.会員稼働予定.月.ToString().PadLeft(2, '0'), out col))
                        {
                            // 非該当期間のとき読み飛ばし
                            continue;
                        }

                        ClsEastEng eastEng = new ClsEastEng
                        {
                            カード番号 = t.カード番号,
                            氏名 = t.氏名,
                            会員稼働予定 = t.会員稼働予定
                        };

                        // ホテル向けガイドリスト(英語)を参照
                        if (!IsHotelListMember(Hoteltbl, out clsHotel, out clsSchedule, eastEng))
                        {
                            // ホテル向けガイドリスト(英語)未掲載はネグる
                            continue;
                        }

                        // 稼働予定を含む組合員情報を稼働表エクセルシートに貼付
                        if (!XlsCellsSetXML_BySheet(clsHotel, clsSchedule, tmpSheet, sheetStRow, col, cardNum, logFile))
                        {
                            continue;
                        }

                        // カード番号
                        cardNum = clsHotel.カード番号;
                    }

                    // 表のフォーマットを整える（罫線、列結合）
                    SheetFormat(tmpSheet, xCol, logFile);

                    //保存処理
                    book.SaveAs(Properties.Settings.Default.xlsHotelsWorksPath);
                }

                // ログ出力
                System.IO.File.AppendAllText(logFile, Form1.GetNowTime(" ホテル向けガイド稼働表を更新しました。"), Encoding.GetEncoding(932));

                // パスワード付きで再度書き換え：2023/03/17
                _ = Utility.PwdXlsFile(Properties.Settings.Default.xlsHotelsWorksPath, Properties.Settings.Default.xlsPassword, "", logFile);
            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
                // ログ出力
                System.IO.File.AppendAllText(logFile, Form1.GetNowTime(ex.ToString()), Encoding.GetEncoding(932));
            }
            finally
            {

            }
        }

        /// <summary>
        ///     対象期間内の予定か検証 </summary>
        /// <param name="yymm">
        ///     年月文字列</param>
        /// <param name="col">
        ///     開始列</param>
        /// <returns>
        ///     true:期間内、false:期間外</returns>
        private bool IsTargetPeriod(string yymm, out int col)
        {
            col = 0;
            bool yymmOn = false;

            for (int iX = 0; iX < 6; iX++)
            {
                if (sheetYYMM[iX, 0] == yymm)
                {
                    col = int.TryParse(sheetYYMM[iX, 1], out int x) ? x : 0;
                    yymmOn = true;
                    break;
                }
            }
            return yymmOn;
        }

        /// <summary>
        ///     新ホテル向けガイドリストに掲載されているか検証　</summary>
        /// <param name="Hoteltbl">
        ///     新ホテル向けガイドリストテーブル</param>
        /// <param name="clsHotel">
        ///     新ホテル向けガイド稼働表・組合員情報クラス</param>
        /// <param name="clsSchedule">
        ///     新ホテル向けガイド稼働表・予定表クラス</param>
        /// <param name="t">
        ///     会員稼働予定情報（東、英語）</param>
        /// <returns>
        ///     掲載：true, 非掲載：false</returns>
        private bool IsHotelListMember(IXLTable Hoteltbl, out ClsHotelScheduleXls clsHotel, out ClsScheduleDays[] clsSchedule, ClsEastEng t)
        {
            bool rtn = false;

            clsHotel = null;
            clsSchedule = new ClsScheduleDays[31];

            // ホテル向けガイドリスト(英語)を参照
            foreach (var row in Hoteltbl.Rows())
            {
                var card = row.Cell(1).Value;
                if (string.IsNullOrEmpty(card.ToString()))
                {
                    continue;
                }

                if (card.ToString() == t.カード番号.ToString())
                {
                    clsHotel = new ClsHotelScheduleXls
                    {
                        Year = t.会員稼働予定.年,
                        Month = t.会員稼働予定.月,
                        カード番号 = t.カード番号.ToString(),
                        氏名 = t.氏名,
                        フリガナ = t.会員稼働予定.フリガナ,
                        携帯電話 = GetNewHotelXCellValue(row.Cell(4).Value),
                        分類 = GetNewHotelXCellValue(row.Cell(5).Value),
                        アサイン2019 = GetNewHotelXCellValue(row.Cell(6).Value),
                        アサイン2020 = GetNewHotelXCellValue(row.Cell(7).Value),
                        クレーム履歴 = GetNewHotelXCellValue(row.Cell(8).Value),
                        プレゼン面談年月 = GetMeetingDate(GetNewHotelXCellValue(row.Cell(10).Value)),
                        得意分野 = GetNewHotelXCellValue(row.Cell(11).Value),
                        保険加入 = GetNewHotelXCellValue(row.Cell(12).Value),
                        都道府県 = GetNewHotelXCellValue(row.Cell(13).Value),
                        市区町村 = GetNewHotelXCellValue(row.Cell(14).Value),
                        メールアドレス = GetNewHotelXCellValue(row.Cell(15).Value),
                        他言語 = GetNewHotelXCellValue(row.Cell(16).Value),
                        FIT = GetNewHotelXCellValue(row.Cell(17).Value),
                        マンダリン = GetNewHotelXCellValue(row.Cell(18).Value),
                        ペニンシュラ = GetNewHotelXCellValue(row.Cell(19).Value),
                        稼働日数 = t.会員稼働予定.稼働日数.ToString(),
                        更新日 = t.会員稼働予定.更新日.ToString()
                    };

                    for (int i = 0; i < 31; i++)
                    {
                        clsSchedule[i] = new ClsScheduleDays();
                        if (i ==  0) clsSchedule[i].予定 = t.会員稼働予定.d1;
                        if (i ==  1) clsSchedule[i].予定 = t.会員稼働予定.d2;
                        if (i ==  2) clsSchedule[i].予定 = t.会員稼働予定.d3;
                        if (i ==  3) clsSchedule[i].予定 = t.会員稼働予定.d4;
                        if (i ==  4) clsSchedule[i].予定 = t.会員稼働予定.d5;
                        if (i ==  5) clsSchedule[i].予定 = t.会員稼働予定.d6;
                        if (i ==  6) clsSchedule[i].予定 = t.会員稼働予定.d7;
                        if (i ==  7) clsSchedule[i].予定 = t.会員稼働予定.d8;
                        if (i ==  8) clsSchedule[i].予定 = t.会員稼働予定.d9;
                        if (i ==  9) clsSchedule[i].予定 = t.会員稼働予定.d10;
                        if (i == 10) clsSchedule[i].予定 = t.会員稼働予定.d11;
                        if (i == 11) clsSchedule[i].予定 = t.会員稼働予定.d12;
                        if (i == 12) clsSchedule[i].予定 = t.会員稼働予定.d13;
                        if (i == 13) clsSchedule[i].予定 = t.会員稼働予定.d14;
                        if (i == 14) clsSchedule[i].予定 = t.会員稼働予定.d15;
                        if (i == 15) clsSchedule[i].予定 = t.会員稼働予定.d16;
                        if (i == 16) clsSchedule[i].予定 = t.会員稼働予定.d17;
                        if (i == 17) clsSchedule[i].予定 = t.会員稼働予定.d18;
                        if (i == 18) clsSchedule[i].予定 = t.会員稼働予定.d19;
                        if (i == 19) clsSchedule[i].予定 = t.会員稼働予定.d20;
                        if (i == 20) clsSchedule[i].予定 = t.会員稼働予定.d21;
                        if (i == 21) clsSchedule[i].予定 = t.会員稼働予定.d22;
                        if (i == 22) clsSchedule[i].予定 = t.会員稼働予定.d23;
                        if (i == 23) clsSchedule[i].予定 = t.会員稼働予定.d24;
                        if (i == 24) clsSchedule[i].予定 = t.会員稼働予定.d25;
                        if (i == 25) clsSchedule[i].予定 = t.会員稼働予定.d26;
                        if (i == 26) clsSchedule[i].予定 = t.会員稼働予定.d27;
                        if (i == 27) clsSchedule[i].予定 = t.会員稼働予定.d28;
                        if (i == 28) clsSchedule[i].予定 = t.会員稼働予定.d29;
                        if (i == 29) clsSchedule[i].予定 = t.会員稼働予定.d30;
                        if (i == 30) clsSchedule[i].予定 = t.会員稼働予定.d31;
                    }

                    rtn = true;
                    break;
                };
            }

            return rtn;
        }

        /// <summary>
        ///     シートの書式を設定する </summary>
        /// <param name="tmpSheet">
        ///     カレントシート</param>
        /// <param name="sCol">
        ///     予定開始列</param>
        /// <param name="logFile">
        ///     ログファイルパス</param>
        private void SheetFormat(IXLWorksheet tmpSheet, int sCol, string logFile)
        {
            tmpSheet.Style.Font.SetFontName("ＭＳ Ｐゴシック");

            SetExcelSheetProperty<ClsHotelScheduleXls>(tmpSheet);

            SetExcelScheduledSheetProperty<ClsScheduleDays>(tmpSheet, sCol, tmpSheet.LastCellUsed().Address.ColumnNumber);

            // ヘッダ行書式設定
            for (int i = 1; i < 4; i++)
            {
                // 行の高さ
                tmpSheet.Row(i).Height = 45;

                // ヘッダ書式設定（縦横位置、折り返して全体を表示）
                tmpSheet.Row(i).Style.Font.SetBold(true)
                                     .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                                     .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                                     .Alignment.SetWrapText(true);
            }

            // カレンダーにない日の列削除
            bool colDelStatus = true;

            while (colDelStatus)
            {
                for (int cl = sCol; cl <= tmpSheet.RangeUsed().RangeAddress.LastAddress.ColumnNumber; cl++)
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

            // 稼働予定年月を表すセルを結合する
            int stCell = 0;
            int edCell = 0;

            for (int cl = xCol; cl <= tmpSheet.LastCellUsed().Address.ColumnNumber; cl++)
            {
                // 2023/01/26 : 作成日以前の列が削除されたため2日以降で始まるケースあり
                if (stCell == 0)
                {
                    stCell = cl;
                }

                if (Utility.nulltoString(tmpSheet.Cell(2, cl).Value).Trim() == "1")
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
            var range = tmpSheet.Range(tmpSheet.Cell(1, 1).Address, tmpSheet.LastCellUsed().Address);
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            // 年月セル下部に罫線を引く
            range = tmpSheet.Range(tmpSheet.Cell(1, sCol).Address, tmpSheet.Cell(1, tmpSheet.LastCellUsed().Address.ColumnNumber).Address);
            range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

            range = tmpSheet.Range(tmpSheet.Cell(2, sCol).Address, tmpSheet.Cell(2, tmpSheet.LastCellUsed().Address.ColumnNumber).Address);
            range.Style.Border.BottomBorder = XLBorderStyleValues.Dotted;

            // 明細最上部に罫線を引く
            range = tmpSheet.Range(tmpSheet.Cell(4, 1).Address, tmpSheet.Cell(4, tmpSheet.LastCellUsed().Address.ColumnNumber).Address);
            range.Style.Border.TopBorder = XLBorderStyleValues.Thin;

            // 表の外枠左罫線を引く
            range = tmpSheet.Range(tmpSheet.Cell(1, 1).Address, tmpSheet.LastCellUsed().Address);
            range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;

            // 見出しの背景色 
            range = tmpSheet.Range(tmpSheet.Cell(1, 1).Address, tmpSheet.Cell(3, tmpSheet.LastCellUsed().Address.ColumnNumber).Address);
            range.Style.Fill.SetBackgroundColor(HeaderBackColor).Font.SetFontColor(XLColor.White);

            // 日曜日の背景色
            range = tmpSheet.Range(tmpSheet.Cell(3, sCol).Address, tmpSheet.Cell(3, tmpSheet.LastCellUsed().Address.ColumnNumber).Address);
            range.AddConditionalFormat().WhenEquals("日").Fill.SetBackgroundColor(XLColor.MistyRose).Font.SetFontColor(XLColor.Black);

            // 日曜日の日付の背景色
            var range2 = tmpSheet.Range(tmpSheet.Cell(2, sCol).Address, tmpSheet.Cell(2, tmpSheet.LastCellUsed().Address.ColumnNumber).Address);
            range2.AddConditionalFormat().WhenIsTrue("=U3=" + @"""日""").Fill.SetBackgroundColor(XLColor.MistyRose).Font.SetFontColor(XLColor.Black);
            
            // ウィンドウ枠の固定
            tmpSheet.SheetView.Freeze(3, 5);

            // フィルタの設定：2023/1/25
            tmpSheet.Row(2).SetAutoFilter();
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
                else if (t.ew == cWEST) // 西
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

        /// --------------------------------------------------------------------------
        /// <summary>
        ///     新ホテル向けガイド稼働表エクセルシート作成 2023/02/14 </summary>
        /// <param name="clsHotel">
        ///     新ホテル向けガイド稼働表・組合員情報クラス</param>
        /// <param name="clsSchedule">
        ///     新ホテル向けガイド稼働表・予定表クラス</param>
        /// <param name="sheet">
        ///     アサイン担当用稼働表エクセルシート </param>
        /// <param name="sRow">
        ///     明細開始行</param>
        /// <param name="sCol">
        ///     該当月予定開始列</param>
        /// <param name="cardNum">
        ///     カード番号</param>
        /// --------------------------------------------------------------------------
        private bool XlsCellsSetXML_BySheet(ClsHotelScheduleXls clsHotel, ClsScheduleDays[] clsSchedule, ClosedXML.Excel.IXLWorksheet sheet, int stRow, int sCol, string cardNum, string logFile)
        {
            var SRow = stRow;
            if (SRow < sheet.LastRowUsed().RowNumber())
            {
                SRow = sheet.LastRowUsed().RowNumber();
            }

            // 組合員が変わったとき
            if (cardNum != string.Empty && cardNum != clsHotel.カード番号)
            {
                // 行下部へヨコ罫線を引く
                sheet.Range(sheet.Cell(SRow, 1).Address,
                sheet.Cell(SRow, sheet.LastCellUsed().Address.ColumnNumber).Address).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                SRow++;
            }

            if (cardNum != clsHotel.カード番号)
            {
                sheet.Row(SRow).Height = 18;

                // 組合員情報クラスからデータ貼り付け
                foreach (PropertyInfo propertyInfo in typeof(ClsHotelScheduleXls).GetProperties())
                {
                    //属性が定義されたプロパティだけを参照するため、fixedAttrがnullなら処理の対象外
                    if (Attribute.GetCustomAttribute(propertyInfo, typeof(ColumnNameAttribute)) is ColumnNameAttribute attribute)
                    {
                        // 値取得
                        var val = propertyInfo.GetValue(clsHotel) == null ? "" : propertyInfo.GetValue(clsHotel).ToString();

                        // セルに貼り付け
                        sheet.Cell(SRow, attribute.ColumnName).SetValue(val);
                    }
                }

                // 奇数行なら背景色
                if (SRow % 2 == 0)
                {
                    var range = sheet.Range(sheet.Cell(SRow, 1).Address, sheet.Cell(SRow, sheet.LastCellUsed().Address.ColumnNumber).Address);
                    range.Style.Fill.SetBackgroundColor(LineBackColor);
                }
            }

            // 予定申告内容をセルに貼り付ける
            for (int i = 0; i < clsSchedule.Length; i++)
            {
                sheet.Cell(SRow, sCol + i).SetValue(clsSchedule[i].予定);
            }

            // アサイン担当者か検証する
            jfgDataClassDataContext db = new jfgDataClassDataContext();
            var s = db.アサイン担当者.Where(a => a.カード番号 == double.Parse(clsHotel.カード番号));
            if (s.Count() > 0)
            {
                // カード番号、氏名、フリガナ各セルのBackColorを黄色にする
                for (int i = 1; i < 4; i++)
                {
                    sheet.Cell(SRow, i).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                }
            }

            return true;
        }

        /// <summary>
        /// 新ホテル向けガイドリストセル値取得
        /// </summary>
        /// <param name="obj">セル値</param>
        /// <returns>string</returns>
        private string GetNewHotelXCellValue(object obj)
        {
            return obj == null ? "" : obj.ToString();
        }

        /// <summary>
        ///新ホテル向けガイドリストセル値取得・プレゼン用面談年月
        /// </summary>
        /// <param name="obj">セル値</param>
        /// <returns>面談年月文字列</returns>
        private string GetMeetingDate(string yymmdd)
        {
            string yymm = "";
            if (DateTime.TryParse(yymmdd, out DateTime dt))
            {
                yymm = dt.Year + "年" + dt.Month + "月";
            }
            return yymm;
        }


        ///----------------------------------------------------
        /// <summary>
        ///     言語配列作成　</summary>
        ///----------------------------------------------------
        private void ReadLang()
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


        ///------------------------------------------------------------------------------
        /// <summary>
        ///     クラスやプロパティに設定された属性の値を取得する　</summary>
        /// <typeparam name="T">
        ///     属性が定義されたクラス</typeparam>
        ///------------------------------------------------------------------------------
        public void SetExcelSheetProperty<T>(IXLWorksheet tmpSheet)
        {
            foreach (PropertyInfo propertyInfo in typeof(T).GetProperties())
            {
                //属性が定義されたプロパティだけを参照するため、fixedAttrがnullなら処理の対象外
                if (Attribute.GetCustomAttribute(propertyInfo, typeof(ColumnNameAttribute)) is ColumnNameAttribute attribute)
                {
                    // 列幅
                    tmpSheet.Column(attribute.ColumnName).Width = attribute.Width;
                    tmpSheet.Column(attribute.ColumnName).Style.Alignment.SetHorizontal(attribute.AlignHorizon)
                                                               .Alignment.SetVertical(attribute.AlignVertial);

                    // フォントサイズ
                    var rng = tmpSheet.Range(tmpSheet.Cell(1, attribute.ColumnIndex).Address, tmpSheet.Cell(3, attribute.ColumnIndex).Address);
                    rng.Style.Font.FontSize = attribute.HeaderFontSize;

                    //セル結合
                    rng.Merge(false);
                }
            }
        }

        ///------------------------------------------------------------------------------
        /// <summary>
        ///     クラスやプロパティに設定された属性の値を取得する　</summary>
        /// <typeparam name="T">
        ///     属性が定義されたクラス</typeparam>
        ///------------------------------------------------------------------------------
        public void SetExcelScheduledSheetProperty<T>(IXLWorksheet tmpSheet, int sCol, int Columns)
        {
            foreach (PropertyInfo propertyInfo in typeof(T).GetProperties())
            {
                //属性が定義されたプロパティだけを参照するため、fixedAttrがnullなら処理の対象外
                if (Attribute.GetCustomAttribute(propertyInfo, typeof(ColumnNameAttribute)) is ColumnNameAttribute attribute)
                {
                    for (int i = sCol; i <= Columns; i++)
                    {
                        tmpSheet.Column(i).Width = attribute.Width;
                        tmpSheet.Column(i).Style.Alignment.SetHorizontal(attribute.AlignHorizon)
                                                .Alignment.SetVertical(attribute.AlignVertial);
                    }
                }
            }
        }
    }
}
