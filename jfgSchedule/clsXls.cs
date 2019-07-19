using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;

namespace jfgSchedule
{
    class clsXls
    {
        string xlsPath = string.Empty;

        jfgDataClassDataContext db = new jfgDataClassDataContext();

        // ログ配列
        string[] log = new string[1];

        // ログ件数
        int logCnt = 0;

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     メインルーチン </summary>
        /// <param name="sPath">
        ///     予定申告エクセルファイルのパス</param>
        /// <param name="sTime">
        ///     前回更新日時</param>
        /// <returns>
        ///     組合員稼働予定テーブル更新データ数</returns>
        ///-----------------------------------------------------------------------
        public int xlsSelect(string sPath, DateTime sTime)
        {
            int sCnt = 0;

            // 指定フォルダのエクセル予定申告シートを取得します
            xlsPath = sPath;
            foreach (string file in System.IO.Directory.GetFiles(xlsPath, "*.xlsx"))
            {
                // 前回の実行時間より更新日が新しいエクセル予定申告シートを判定します
                DateTime dt = System.IO.File.GetLastWriteTime(file);
                if (dt.CompareTo(sTime) >= 0)
                {
                    // 組合員稼働予定テーブルを更新します
                    //dataOutput(file, dt);
                    dataOutputXML(file, dt);    // 2018/02/22
                    sCnt++;
                }
            }

            //Console.WriteLine(sCnt.ToString()+ "枚の予定申告書シートから会員稼働予定テーブルが更新されました");

            // ログ出力
            if (!(log.Length == 1 && log[0] == null))
            {
                writeLog(Properties.Settings.Default.xlsxPath + Properties.Settings.Default.logFileName, log);
            }

            // 更新されたエクセル予定申告シートがあった場合
            if (sCnt > 0)
            {
                // 会員稼働予定テーブル：稼働日数、自己申告日数更新
                updateDays();
            }

            // 更新データ数を返す
            return sCnt;
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     会員稼働予定テーブル更新メインルーチン </summary>
        /// <param name="sFile">
        ///     予定申告エクセルファイルパス</param>
        ///------------------------------------------------------------------
        private void dataOutput(string sFile, DateTime sDt)
        {
            // Excelテンプレートシート開く
            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

            Excel.Range rng;
            Excel.Range rngYear;
            Excel.Range rngMonth;
            Excel.Range[] rngs = new Microsoft.Office.Interop.Excel.Range[31];

            string lg = string.Empty;
            
            jfgDataClassDataContext db = new jfgDataClassDataContext();

            // 個人データクラス
            clsPersonalData cPData = new clsPersonalData();
            cPData.sDt = sDt;

            try
            {
                // 予定申告書シートよりカード番号取得
                rng = (Excel.Range)oxlsSheet.Cells[2, 11];
                cPData.cNo = Utility.cNumberCheck(rng.Text.Trim());

                if (cPData.cNo == -1)
                {
                    // ログ文字列作成
                    lg = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " " + System.IO.Path.GetFileName(sFile) + " " +
                        "カード番号が正しい数字情報ではありません : 更新から除外されました。";

                    // ログ配列に出力
                    arrayLog(lg);

                    // 2017/11/30 以下、コメント化
                    //// ExcelBookをクローズ
                    //oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //// Excel終了
                    //oXls.Quit();

                    // 戻る
                    return;
                }
                
                // 会員情報よりフリガナを取得
                var s = db.会員情報.Where(a => a.カード番号 == cPData.cNo);
                if (s.Count() == 0)
                {
                    // ログ文字列作成
                    lg = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " " + System.IO.Path.GetFileName(sFile) + " " +
                        "カード番号に該当する会員が存在しません [" + cPData.cNo.ToString() + "] : 更新から除外されました。";

                    // ログ配列に出力
                    arrayLog(lg);

                    // 2017/11/30 以下、コメント化
                    //// ExcelBookをクローズ
                    //oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //// Excel終了
                    //oXls.Quit();

                    // 戻る
                    return;
                }
                else
                {
                    foreach (var t in s)
                    {
                        cPData.furi = t.フリガナ;
                    }
                }

                // 各月の予定を会員稼働予定テーブルに書き込む
                for (int i = 0; i < 6; i++)
                {
                    // 年を取得
                    rngYear = (Excel.Range)oxlsSheet.Cells[5, i * 4 + 2];
                    cPData.eYear = cIntCheck(rangeNullToString(rngYear.Text).Trim());

                    if (cPData.eYear == -1)
                    {
                        // ログ文字列作成
                        lg = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " " + System.IO.Path.GetFileName(sFile) + " " +
                            (i + 1).ToString() + "番目の年が正しくありません : 更新から除外されました。";

                        // ログ配列に出力
                        arrayLog(lg);

                        // 次ぎへ
                        continue;
                    }

                    // 月を取得
                    rngMonth = (Excel.Range)oxlsSheet.Cells[5, i * 4 + 3];
                    cPData.eMonth = cIntCheck(rangeNullToString(rngMonth.Text).Trim());

                    if (cPData.eMonth == -1)
                    {
                        // ログ文字列作成
                        lg = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " " + System.IO.Path.GetFileName(sFile) + " " +
                            (i + 1).ToString() + "番目の月が正しくありません : 更新から除外されました。";

                        // ログ配列に出力
                        arrayLog(lg);

                        // 次ぎへ
                        continue;
                    }

                    // 過去の予定は更新しない
                    int nowYYMM = DateTime.Today.Year * 100 + DateTime.Today.Month;
                    int xlsYYMM = cPData.eYear * 100 + cPData.eMonth;
                    if (xlsYYMM < nowYYMM) continue; // 次へ

                    // 連絡事項を取得
                    rng = (Excel.Range)oxlsSheet.Cells[2, 15];
                    cPData.memo = rng.Text.Trim();

                    // 該当列を取得
                    cPData.colidx = i * 4 + 3;

                    // エクセルシートオブジェクトを取得
                    cPData.exl = oxlsSheet;

                    // データ書き込み
                    if (saveData(cPData))
                    {
                        // ログ文字列作成
                        lg = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " " + System.IO.Path.GetFileName(sFile) + " " +
                            cPData.eYear.ToString() + "年" + cPData.eMonth.ToString() + "月の予定が登録されました。";
                    }
                    else
                    {
                        // ログ文字列作成
                        lg = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " " + System.IO.Path.GetFileName(sFile) + " " +
                            cPData.eYear.ToString() + "年" + cPData.eMonth.ToString() + "月の予定の登録に失敗しました。";                        
                    }

                    // ログ配列に出力
                    arrayLog(lg);
                }

                //Console.WriteLine(cPData.cNo + " " + cPData.furi);
            }
            catch(Exception e)
            {

            }
            finally
            {
                // ExcelBookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excel終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);
            }
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     会員稼働予定テーブル更新メインルーチン 
        ///     : closedXML版 2018/02/22</summary>
        /// <param name="sFile">
        ///     予定申告エクセルファイルパス</param>
        ///------------------------------------------------------------------
        private void dataOutputXML(string sFile, DateTime sDt)
        {
            string lg = string.Empty;

            jfgDataClassDataContext db = new jfgDataClassDataContext();

            // 個人データクラス
            clsPersonalData cPData = new clsPersonalData();
            cPData.sDt = sDt;

            try
            {
                using (var book = new XLWorkbook(sFile, XLEventTracking.Disabled))
                {
                    // ワークシートを取得する
                    IXLWorksheet sheet = book.Worksheet(1);

                    // 予定申告書シートよりカード番号取得
                    cPData.cNo = Utility.cNumberCheck(Utility.nulltoString(sheet.Cell(2, 11).Value).Trim());

                    if (cPData.cNo == -1)
                    {
                        // ログ文字列作成
                        lg = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " " + System.IO.Path.GetFileName(sFile) + " " +
                            "カード番号が正しい数字情報ではありません : 更新から除外されました。";

                        // ログ配列に出力
                        arrayLog(lg);

                        // 戻る
                        return;
                    }

                    // 会員情報よりフリガナを取得
                    if (!db.会員情報.Any(a => a.カード番号 == cPData.cNo))
                    {
                        // ログ文字列作成
                        lg = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " " + System.IO.Path.GetFileName(sFile) + " " +
                            "カード番号に該当する会員が存在しません [" + cPData.cNo.ToString() + "] : 更新から除外されました。";

                        // ログ配列に出力
                        arrayLog(lg);

                        // 戻る
                        return;
                    }
                    else
                    {
                        foreach (var t in db.会員情報.Where(a => a.カード番号 == cPData.cNo))
                        {
                            cPData.furi = t.フリガナ;
                        }
                    }

                    // 各月の予定を会員稼働予定テーブルに書き込む
                    for (int i = 0; i < 6; i++)
                    {
                        // 年を取得
                        cPData.eYear = cIntCheck(rangeNullToString(Utility.nulltoString(sheet.Cell(5, i * 4 + 2).Value).Trim()));

                        if (cPData.eYear == -1)
                        {
                            // ログ文字列作成
                            lg = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " " + System.IO.Path.GetFileName(sFile) + " " +
                                (i + 1).ToString() + "番目の年が正しくありません : 更新から除外されました。";

                            // ログ配列に出力
                            arrayLog(lg);

                            // 次ぎへ
                            continue;
                        }

                        // 月を取得
                        cPData.eMonth = cIntCheck(Utility.nulltoString(sheet.Cell(5, i * 4 + 3).Value).Trim());

                        if (cPData.eMonth == -1)
                        {
                            // ログ文字列作成
                            lg = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " " + System.IO.Path.GetFileName(sFile) + " " +
                                (i + 1).ToString() + "番目の月が正しくありません : 更新から除外されました。";

                            // ログ配列に出力
                            arrayLog(lg);

                            // 次ぎへ
                            continue;
                        }

                        // 過去の予定は更新しない
                        int nowYYMM = DateTime.Today.Year * 100 + DateTime.Today.Month;
                        int xlsYYMM = cPData.eYear * 100 + cPData.eMonth;
                        if (xlsYYMM < nowYYMM)
                        {
                            continue; // 次へ
                        }

                        // 連絡事項を取得
                        cPData.memo = Utility.nulltoString(sheet.Cell(2, 15).Value).Trim();

                        // 該当列を取得
                        cPData.colidx = i * 4 + 3;

                        // エクセルシートオブジェクトを取得
                        cPData.sheet = sheet;

                        // データ書き込み
                        if (saveData(cPData))
                        {
                            // ログ文字列作成
                            lg = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " " + System.IO.Path.GetFileName(sFile) + " " +
                                cPData.eYear.ToString() + "年" + cPData.eMonth.ToString() + "月の予定が登録されました。";
                        }
                        else
                        {
                            // ログ文字列作成
                            lg = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " " + System.IO.Path.GetFileName(sFile) + " " +
                                cPData.eYear.ToString() + "年" + cPData.eMonth.ToString() + "月の予定の登録に失敗しました。";
                        }

                        // ログ配列に出力
                        arrayLog(lg);
                    }

                    // シート後片付け
                    sheet.Dispose();
                }
            }
            catch (Exception e)
            {

            }
            finally
            {
                
            }
        }

        /// ------------------------------------------------------------------------------------
        /// <summary>
        ///     会員稼働予定データ登録 </summary>
        /// <param name="cPData">
        ///     個人情報クラス</param>
        /// ------------------------------------------------------------------------------------
        private bool saveData(clsPersonalData cPData)
        {
            bool rtn = false;

            try
            {
                //Excel.Range[] rngs = new Microsoft.Office.Interop.Excel.Range[31];
                string lg = string.Empty;

                // 月末日を取得する（翌月1日の1日前）
                DateTime dt = new DateTime(cPData.eYear, cPData.eMonth, 1).AddMonths(1).AddDays(-1);

                // データが登録済みか確認します
                if (!getScheRec(cPData))
                {
                    // 新規登録
                    if (insertData(cPData, dt.Day))
                    {
                        rtn = true;
                    }
                }
                else
                {
                    // 更新
                    if (updateData(cPData, dt.Day))
                    {
                        rtn = true;
                    }
                }

                return rtn;
            }
            catch (Exception e)
            {
                return false;
            }
        }
        
        /// ------------------------------------------------------------------------------------
        /// <summary>
        ///     会員稼働予定データ新規登録 </summary>
        /// <param name="cPData">
        ///     個人情報クラス</param>
        /// <param name="dDay">
        ///     該当月の日数</param>
        /// ------------------------------------------------------------------------------------
        private bool insertData(clsPersonalData cPData, int dDay)
        {
            try
            {
                // 会員稼働予定テーブルのインスタンスを新規に作成
                会員稼働予定 tbl = new 会員稼働予定();
                
                // エンティティセット
                entitySet(tbl, cPData, dDay);

                // レコード挿入
                db.会員稼働予定.InsertOnSubmit(tbl);

                // 登録を確定させる
                db.SubmitChanges();

                return true;
            }
            catch(Exception e)
            {
                return false;
            }
        }

        /// ------------------------------------------------------------------------------------
        /// <summary>
        ///     会員稼働予定データ更新 </summary>
        /// <param name="cPData">
        ///     個人情報クラス</param>
        /// <param name="dDay">
        ///     該当月の日数</param>
        /// ------------------------------------------------------------------------------------
        private bool updateData(clsPersonalData cPData, int dDay)
        {
            try
            {
                // 対象データを取得
                会員稼働予定 s = (会員稼働予定)db.会員稼働予定.Single(a => a.カード番号 == cPData.cNo && a.年 == cPData.eYear && a.月 == cPData.eMonth);

                // エンティティセット
                entitySet(s, cPData, dDay);

                // 更新を確定させる
                db.SubmitChanges();

                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        /// ------------------------------------------------------------------------------------
        /// <summary>
        ///     会員稼働予定データエンティティセット </summary>
        /// <param name="tbl">
        ///     会員稼働予定インスタンス</param>
        /// <param name="cPData">
        ///     個人情報クラス</param>
        /// <param name="dDay">
        ///     該当月の日数</param>
        /// ------------------------------------------------------------------------------------
        private void entitySet(会員稼働予定 tbl, clsPersonalData cPData, int dDay)
        {
            const int sRow = 8;

            tbl.カード番号 = cPData.cNo;
            tbl.年 = cPData.eYear;
            tbl.月 = cPData.eMonth;
            tbl.フリガナ = cPData.furi;
            tbl.申告年月日 = (DateTime)cPData.sDt;
            tbl.アサイン区分 = 0;

            for (int i = 0; i < 31; i++)
            {
                // コメント化 2019/07/18
                //string dVal = Utility.nulltoString(cPData.sheet.Cell(sRow + i, cPData.colidx).Value);

                // Trim()を追加 2019/07/18
                string dVal = Utility.nulltoString(cPData.sheet.Cell(sRow + i, cPData.colidx).Value).Trim();

                if (i == 0) tbl.d1 = dVal;
                if (i == 1) tbl.d2 = dVal;
                if (i == 2) tbl.d3 = dVal;
                if (i == 3) tbl.d4 = dVal;
                if (i == 4) tbl.d5 = dVal;
                if (i == 5) tbl.d6 = dVal;
                if (i == 6) tbl.d7 = dVal;
                if (i == 7) tbl.d8 = dVal;
                if (i == 8) tbl.d9 = dVal;
                if (i == 9) tbl.d10 = dVal;
                if (i == 10) tbl.d11 = dVal;
                if (i == 11) tbl.d12 = dVal;
                if (i == 12) tbl.d13 = dVal;
                if (i == 13) tbl.d14 = dVal;
                if (i == 14) tbl.d15 = dVal;
                if (i == 15) tbl.d16 = dVal;
                if (i == 16) tbl.d17 = dVal;
                if (i == 17) tbl.d18 = dVal;
                if (i == 18) tbl.d19 = dVal;
                if (i == 19) tbl.d20 = dVal;
                if (i == 20) tbl.d21 = dVal;
                if (i == 21) tbl.d22 = dVal;
                if (i == 22) tbl.d23 = dVal;
                if (i == 23) tbl.d24 = dVal;
                if (i == 24) tbl.d25 = dVal;
                if (i == 25) tbl.d26 = dVal;
                if (i == 26) tbl.d27 = dVal;
                if (i == 27) tbl.d28 = dVal;

                if (i == 28)
                {
                    if (dDay < 29)
                    {
                        tbl.d29 = string.Empty;
                    }
                    else
                    {
                        tbl.d29 = dVal;
                    }
                }

                if (i == 29)
                {
                    if (dDay < 30)
                    {
                        tbl.d30 = string.Empty;
                    }
                    else
                    {
                        tbl.d30 = dVal;
                    }
                }
                    
                if (i == 30)
                {
                    if (dDay < 31)
                    {
                        tbl.d31 = string.Empty;
                    }
                    else
                    {
                        tbl.d31 = dVal;
                    }
                }

                tbl.備考 = cPData.memo;
                tbl.更新日 = DateTime.Now;
            }

        }
        
        ///-------------------------------------------------------------------------
        /// <summary>
        ///     int型変換可能な値か検証</summary>
        /// <param name="cel">
        ///     セルの値</param>
        /// <returns>
        ///     int型変換のとき値を返す、int型変換エラーのとき-1を返す</returns>
        ///-------------------------------------------------------------------------
        private int cIntCheck(string cel)
        {
            int rtn = -1;
            int cNo;
            if (int.TryParse(cel, out cNo))
            {
                rtn = cNo;
            }

            return rtn;
        }
        
        ///------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、nullではないときそのまま値を返す</returns>
        ///------------------------------------------------------------------------
        private string rangeNullToString(object rng)
        {
            if (rng == null)
            {
                return string.Empty;
            }
            else
            {
                return rng.ToString();
            }
        }
        
        ///--------------------------------------------------------------
        /// <summary>
        ///     稼働予定データが登録済みか調べる </summary>
        /// <param name="cPData">
        ///     個人情報クラス</param>
        /// <returns>
        ///     true:登録済み, false:未登録</returns>
        ///--------------------------------------------------------------
        private bool getScheRec(clsPersonalData cPData)
        {
            //var s = db.会員稼働予定.Where(a => a.カード番号 == cPData.cNo && a.年 == cPData.eYear && a.月 == cPData.eMonth);
            //if (s.Count() != 0) return true;
            //else return false;

            var s = db.会員稼働予定.Where(a => a.カード番号 == cPData.cNo && a.年 == cPData.eYear && a.月 == cPData.eMonth);
            if (db.会員稼働予定.Any(a => a.カード番号 == cPData.cNo && a.年 == cPData.eYear && a.月 == cPData.eMonth))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     ログ配列作成 </summary>
        /// <param name="logTxt">
        ///     ログ文字列</param>
        ///--------------------------------------------------------------
        private void arrayLog(string logTxt)
        {
            if (logCnt > 0)
            {
                Array.Resize(ref log, logCnt + 1);
            }

            log[logCnt] = logTxt;
            logCnt++;
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     ログファイル書き出し </summary>
        /// <param name="sFile">
        ///     ログファイルパス</param>
        /// <param name="arrayLg">
        ///     ログ配列</param>
        ///---------------------------------------------------------------
        private void writeLog(string sFile, string[] arrayLg)
        {
            // ログ出力
            System.IO.File.AppendAllLines(sFile, arrayLg, System.Text.Encoding.GetEncoding(932));
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     会員稼働予定テーブル：稼働日数、自己申告日数を更新　</summary>
        /// -------------------------------------------------------------------   
        private void updateDays()
        {
            DateTime dt = DateTime.Parse(DateTime.Today.Year.ToString() + "/01/01"); 

            jfgDataClassDataContext db = new jfgDataClassDataContext();

            var s = db.会員稼働予定.OrderBy(a => a.カード番号);

            foreach (会員稼働予定 t in s)
            {
                // 稼働日数を取得
                var saQuery = db.アサイン
                    .Where(a => a.カード番号 == t.カード番号 && a.稼働日1 >= dt)
                    .GroupBy(a => a.カード番号)
                    .Select(g => new
                    {
                        cNum = g.Key,
                        nissu = g.Sum(a => a.日数.Value)
                    });

                int kDays = 0;
                foreach (var ti in saQuery)
                {
                    kDays = int.Parse(ti.nissu.ToString());
                }

                // 自己申告日数を取得
                var siQuery = db.アサイン
                    .Where(a => a.カード番号 == t.カード番号 && a.稼働日1 >= dt && 
                           a.アサイン担当.Contains("自己申告"));
                int jikoShinDays = siQuery.Count();

                // 稼働日数、自己申告日数を更新
                t.稼働日数 = kDays;
                t.自己申告日数 = jikoShinDays;
                db.SubmitChanges();
            }
        }
    }
}
