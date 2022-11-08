using ClosedXML.Excel;
namespace ClosedXmlTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IXLTable tbl;
            using (var selectBook = new XLWorkbook(@"C:\\JFGWORKS\ホテル向けガイドリスト(英語) 2022.xlsx"))
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

            foreach (var row in tbl.Rows())
            {
                var card = row.Cell(1).Value;
                MessageBox.Show(card.ToString());
            }


        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}