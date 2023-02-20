using ClosedXML.Excel;
using System;

namespace jfgSchedule
{
    public class ColumnNameAttribute : Attribute
    {
        public string ColumnName { get; set; }
        public int ColumnIndex { get; set; }
        public XLAlignmentHorizontalValues AlignHorizon { get; set; }
        public XLAlignmentVerticalValues AlignVertial { get; set; }
        public float Width { get; set; }
        public int HeaderFontSize { get; set; }
        public ColumnNameAttribute(string name)
        {
            this.ColumnName = name;
        }
    }
    /// <summary>
    /// 新ホテル向けガイド稼働表・組合員情報クラス
    /// </summary>
    public class ClsHotelScheduleXls
    {
        public int Year { get; set; }
        public int Month { get; set; }
        [ColumnName("A", ColumnIndex = 1, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 10, HeaderFontSize = 9)]
        public string カード番号 { get; set; }
        [ColumnName("B", ColumnIndex = 2, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 20, HeaderFontSize = 9)]
        public string 氏名 { get; set; }
        [ColumnName("C", ColumnIndex = 3, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 20, HeaderFontSize = 9)]
        public string フリガナ { get; set; }
        [ColumnName("D", ColumnIndex = 4, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 20, HeaderFontSize = 9)]
        public string 携帯電話 { get; set; }
        [ColumnName("E", ColumnIndex = 5, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 8, HeaderFontSize = 9)]
        public string 分類 { get; set; }
        [ColumnName("F", ColumnIndex = 6, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Right, Width = 9, HeaderFontSize = 9)]
        public string アサイン2019 { get; set; }
        [ColumnName("G", ColumnIndex = 7, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Right, Width = 9, HeaderFontSize = 9)]
        public string アサイン2020 { get; set; }
        [ColumnName("H", ColumnIndex = 8, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 14, HeaderFontSize = 9)]
        public string クレーム履歴 { get; set; }
        [ColumnName("I", ColumnIndex = 9, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 13.6f, HeaderFontSize = 9)]
        public string プレゼン面談年月 { get; set; }
        [ColumnName("J", ColumnIndex = 10, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 20, HeaderFontSize = 9)]
        public string 得意分野 { get; set; }
        [ColumnName("K", ColumnIndex = 11, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 9, HeaderFontSize = 9)]
        public string 保険加入 { get; set; }
        [ColumnName("L", ColumnIndex = 12, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 11, HeaderFontSize = 9)]
        public string 都道府県 { get; set; }
        [ColumnName("M", ColumnIndex = 13, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 13, HeaderFontSize = 9)]
        public string 市区町村 { get; set; }
        [ColumnName("N", ColumnIndex = 14, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 40, HeaderFontSize = 9)]
        public string メールアドレス { get; set; }
        [ColumnName("O", ColumnIndex = 15, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 8, HeaderFontSize = 9)]
        public string 他言語 { get; set; }
        [ColumnName("P", ColumnIndex = 16, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 8, HeaderFontSize = 9)]
        public string FIT { get; set; }
        [ColumnName("Q", ColumnIndex = 17, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 8, HeaderFontSize = 9)]
        public string マンダリン { get; set; }
        [ColumnName("R", ColumnIndex = 18, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 8, HeaderFontSize = 9)]
        public string ペニンシュラ { get; set; }
        [ColumnName("S", ColumnIndex = 19, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 8, HeaderFontSize = 9)]
        public string 稼働日数 { get; set; }
        [ColumnName("T", ColumnIndex = 20, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 26, HeaderFontSize = 9)]
        public string 更新日 { get; set; }
    }
    /// <summary>
    ///     新ホテル向けガイド稼働表・予定表クラス
    /// </summary>
    public class ClsScheduleDays
    {
        [ColumnName("ScheduledDate", AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 4)]
        public string 予定 { get; set; }
    }
}
