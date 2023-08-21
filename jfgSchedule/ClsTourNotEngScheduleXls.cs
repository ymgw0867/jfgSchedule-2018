using ClosedXML.Excel;
using System;

namespace jfgSchedule
{
    /// <summary>
    /// ツアー向けガイド稼働表・組合員情報クラス
    /// </summary>
    public class ClsTourNotEngScheduleXls
    {
        public int Year { get; set; }
        public int Month { get; set; }
        [ColumnName("A", ColumnIndex = 1, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 10, HeaderFontSize = 9)]
        public string カード番号 { get; set; }
        [ColumnName("B", ColumnIndex = 2, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 20, HeaderFontSize = 9)]
        public string 氏名 { get; set; }
        [ColumnName("C", ColumnIndex = 3, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 20, HeaderFontSize = 9)]
        public string フリガナ { get; set; }
        [ColumnName("D", ColumnIndex = 4, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 6, HeaderFontSize = 9)]
        public string 言語 { get; set; }
        [ColumnName("E", ColumnIndex = 4, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 20, HeaderFontSize = 9)]
        public string 携帯電話 { get; set; }
        [ColumnName("F", ColumnIndex = 5, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 8, HeaderFontSize = 9)]
        public string ホテル業務対応可 { get; set; }
        [ColumnName("G", ColumnIndex = 6, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 9, HeaderFontSize = 9)]
        public string 団体インセンティブ対応可 { get; set; }
        [ColumnName("H", ColumnIndex = 7, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 9, HeaderFontSize = 9)]
        public string 一般のツアー対応可 { get; set; }
        [ColumnName("I", ColumnIndex = 8, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 11, HeaderFontSize = 9)]
        public string 東京都福祉衛生局対応 { get; set; }
        [ColumnName("J", ColumnIndex = 9, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 14, HeaderFontSize = 9)]
        public string クレーム履歴 { get; set; }
        [ColumnName("K", ColumnIndex = 10, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 10, HeaderFontSize = 9)]
        public string 生まれ年 { get; set; }
        [ColumnName("L", ColumnIndex = 11, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 16, HeaderFontSize = 9)]
        public string プレゼン面談年月 { get; set; }
        [ColumnName("M", ColumnIndex = 12, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 11, HeaderFontSize = 9)]
        public string 都道府県 { get; set; }
        [ColumnName("N", ColumnIndex = 13, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 14, HeaderFontSize = 9)]
        public string 市区町村 { get; set; }
        [ColumnName("O", ColumnIndex = 14, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 40, HeaderFontSize = 9)]
        public string メールアドレス { get; set; }
        [ColumnName("P", ColumnIndex = 15, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 8, HeaderFontSize = 9)]
        public string 他言語 { get; set; }
        [ColumnName("Q", ColumnIndex = 16, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 20, HeaderFontSize = 9)]
        public string 得意分野 { get; set; }
        [ColumnName("R", ColumnIndex = 17, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 8, HeaderFontSize = 9)]
        public string JFG加入年 { get; set; }
        [ColumnName("S", ColumnIndex = 18, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Right, Width = 9, HeaderFontSize = 9)]
        public string 稼働日数2020 { get; set; }
        [ColumnName("T", ColumnIndex = 19, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Right, Width = 9, HeaderFontSize = 9)]
        public string 稼働日数2023 { get; set; }
        [ColumnName("U", ColumnIndex = 21, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 26, HeaderFontSize = 9)]
        public string 更新日 { get; set; }
    }
}
