using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace xlsScheoutput
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            xlsPut();
        }

        private void xlsPut()
        {
            DialogResult ret;
            int sCnt = 0;
            int dCnt = 0;

            //ダイアログボックスの初期設定
            string pathName = string.Empty;
            ret = folderBrowserDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.OK)
            {
                pathName = folderBrowserDialog1.SelectedPath;
                if (MessageBox.Show("出力先は以下のフォルダでよろしいですか？" + Environment.NewLine+Environment.NewLine + pathName,"確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes);
            }
            else
            {
                return;
            }

            //マウスポインタを待機にする
            this.Cursor = Cursors.WaitCursor;

            // Excelテンプレートシート開く
            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.xlsTempPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

            Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

            try
            {
                // 会員情報を取得
                JfgDataClassDataContext db = new JfgDataClassDataContext();
                var s = db.会員情報.OrderBy(a => a.カード番号);

                //オーナーフォームを無効にする
                this.Enabled = false;

                //プログレスバーを表示する
                frmPrg frmP = new frmPrg(s.Count(), 0);
                frmP.Owner = this;
                frmP.Show();

                // 会員情報を読み込む
                foreach (var item in s)
                {
                    //プログレスバー表示
                    dCnt++;
                    frmP.Text = "エクセル予定申告書作成中　" + dCnt.ToString() + "/" + s.Count().ToString();
                    frmP.progressValue = dCnt;
                    frmP.ProgressStep();
                    
                    // 選択条件
                    if (chkGen.CheckState == CheckState.Checked && item.JFG会員歴 == 1 || 
                        chkTai.CheckState == CheckState.Checked && item.JFG会員歴 == 2 || 
                        chkKyu.CheckState == CheckState.Checked && item.JFG会員歴 == 3 || 
                        chkFumei.CheckState == CheckState.Checked && item.JFG会員歴 == 4 || 
                        chkHi.CheckState == CheckState.Checked && item.JFG会員歴 ==5)
                    {
                        // シートの保護を解除する
                        oxlsSheet.Unprotect();

                        // セルに値を渡す
                        oxlsSheet.Cells[2, 7] = item.氏名 + "さん";
                        oxlsSheet.Cells[2, 11] = item.カード番号.ToString();
                        oxlsSheet.Cells[5, 2] = DateTime.Today.Year.ToString();
                        oxlsSheet.Cells[5, 3] = DateTime.Today.Month.ToString();
                        oxlsSheet.Cells[5, 6] = DateTime.Today.AddMonths(1).Year.ToString();
                        oxlsSheet.Cells[5, 7] = DateTime.Today.AddMonths(1).Month.ToString();
                        oxlsSheet.Cells[5, 10] = DateTime.Today.AddMonths(2).Year.ToString();
                        oxlsSheet.Cells[5, 11] = DateTime.Today.AddMonths(2).Month.ToString();
                        oxlsSheet.Cells[5, 14] = DateTime.Today.AddMonths(3).Year.ToString();
                        oxlsSheet.Cells[5, 15] = DateTime.Today.AddMonths(3).Month.ToString();
                        oxlsSheet.Cells[5, 18] = DateTime.Today.AddMonths(4).Year.ToString();
                        oxlsSheet.Cells[5, 19] = DateTime.Today.AddMonths(4).Month.ToString();
                        oxlsSheet.Cells[5, 22] = DateTime.Today.AddMonths(5).Year.ToString();
                        oxlsSheet.Cells[5, 23] = DateTime.Today.AddMonths(5).Month.ToString();

                        // 申告欄クリア
                        for (int iRow = 8; iRow < 39; iRow++)
                        {
                            oxlsSheet.Cells[iRow, 3] = string.Empty;
                            oxlsSheet.Cells[iRow, 7] = string.Empty;
                            oxlsSheet.Cells[iRow, 11] = string.Empty;
                            oxlsSheet.Cells[iRow, 15] = string.Empty;
                            oxlsSheet.Cells[iRow, 19] = string.Empty;
                            oxlsSheet.Cells[iRow, 23] = string.Empty;
                        }

                        // ウィンドウを非表示にする
                        oXls.Visible = false;

                        // 保存処理
                        oXls.DisplayAlerts = false;

                        // シートの保護
                        oxlsSheet.Protect(Type.Missing, false, true, false, false, false, false, false, false, false, false, false, false, false, false, false);

                        // シートの保存
                        string fileName = pathName + @"\" + item.カード番号.ToString() + " " + item.氏名 + ".xlsx";
                        oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        // カウント
                        sCnt++;
                    }
                }

                // いったんオーナーをアクティブにする
                this.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                this.Enabled = true;

                // 終了メッセージ
                MessageBox.Show(sCnt.ToString() + "人の予定申告シートを出力しました", "予定申告シート出力完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "エクセル予定申告シート出力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }
    }
}
