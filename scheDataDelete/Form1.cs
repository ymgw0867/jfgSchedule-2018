using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace scheDataDelete
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("退会者の稼働予定データと予定申告シートを削除します。" + Environment.NewLine + "よろしいですか。","実行確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // カーソルを待機中
            Cursor.Current = Cursors.WaitCursor;

            // 削除処理
            int n = scheDel();

            // カーソルをデフォルトに戻す
            Cursor.Current = Cursors.Default;

            // 結果表示
            if (n > 0)
            {
                MessageBox.Show(n.ToString() + "件処理しました", "稼働予定データ削除処理", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("該当データはありませんでした", "稼働予定データ削除処理", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            button1.Enabled = false;
        }

        private int scheDel()
        {
            int result = 0;
            
            DataClasses1DataContext db = new DataClasses1DataContext();

            var k = db.会員稼働予定.Select(a => new 
                { 
                    カード番号 = a.カード番号
                }).Distinct();

            foreach (var m in k)
	        {
                // 会員情報の「JFG会員歴」を参照
                if (db.会員情報.Any(a => a.カード番号 == m.カード番号))
                {
                    var s = db.会員情報.Single(a => a.カード番号 == m.カード番号);

                    // JFG会員歴が[1]（現）ではなかったら
                    if (s.JFG会員歴 != 1)
                    {
                        // 削除する
                        foreach (var d in db.会員稼働予定.Where(a => a.カード番号 == m.カード番号))
	                    {
                            // 会員稼働予定データ削除
                            db.会員稼働予定.DeleteOnSubmit(d);

                            // 件数カウント
                            result++;

                            // listBoxに表示
                            listBox1.Items.Add(m.カード番号.ToString() + " " + s.氏名 + "：稼働予定テーブルより削除しました。");
	                    }

                        // 予定申告シートファイル名
                        string exlFileName = m.カード番号.ToString() + " " + s.氏名 + ".xlsx";

                        // 予定申告シートを削除する
                        System.IO.File.Delete(Properties.Settings.Default.xlsxPath + @"\" + exlFileName);
                        listBox1.Items.Add(exlFileName + " を削除しました。");
                    }
                }
	        }
            
            try 
	        {
	            // データベースを確定します
                if (result > 0)
                {
		            db.SubmitChanges();
                }
	        }
	        catch (Exception e)
	        {
		        MessageBox.Show(e.ToString());
	        }

            return result;
        }
    }
}

