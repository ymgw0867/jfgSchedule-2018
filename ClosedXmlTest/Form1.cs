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
            using (var selectBook = new XLWorkbook(@"C:\\JFGWORKS\�z�e�������K�C�h���X�g(�p��) 2022.xlsx"))
            using (var selSheet = selectBook.Worksheet(1))
            {
                // �J�[�h�ԍ��J�n�Z��
                var cell1 = selSheet.Cell("A4");
                // �ŏI�s���擾
                var lastRow = selSheet.LastRowUsed().RowNumber();
                // �J�[�h�ԍ��ŏI�Z��
                var cell2 = selSheet.Cell(lastRow, 1);
                // �J�[�h�ԍ����e�[�u���Ŏ擾
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