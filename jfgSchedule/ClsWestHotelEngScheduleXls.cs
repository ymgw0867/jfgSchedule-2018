using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace jfgSchedule
{
    public class ClsWestHotelEngScheduleXls
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
        [ColumnName("E", ColumnIndex = 5, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 9, HeaderFontSize = 9)]
        public string ホテルアサイン2019 { get; set; }
        [ColumnName("F", ColumnIndex = 8, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 16, HeaderFontSize = 9)]
        public string クレーム履歴 { get; set; }
        [ColumnName("G", ColumnIndex = 9, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 32, HeaderFontSize = 9)]
        public string 備考 { get; set; }
        [ColumnName("H", ColumnIndex = 10, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 10, HeaderFontSize = 9)]
        public string 生まれ年 { get; set; }
        [ColumnName("I", ColumnIndex = 13, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 11, HeaderFontSize = 9)]
        public string 都道府県 { get; set; }
        [ColumnName("J", ColumnIndex = 14, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 11, HeaderFontSize = 9)]
        public string 市区町村 { get; set; }
        [ColumnName("K", ColumnIndex = 15, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 40, HeaderFontSize = 9)]
        public string メールアドレス { get; set; }
        [ColumnName("L", ColumnIndex = 16, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Center, Width = 8, HeaderFontSize = 9)]
        public string 他言語 { get; set; }
        [ColumnName("M", ColumnIndex = 19, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Right, Width = 9, HeaderFontSize = 9)]
        public string 稼働日数2022 { get; set; }
        [ColumnName("N", ColumnIndex = 20, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Right, Width = 9, HeaderFontSize = 9)]
        public string 稼働日数2023 { get; set; }
        [ColumnName("O", ColumnIndex = 21, AlignVertial = XLAlignmentVerticalValues.Center, AlignHorizon = XLAlignmentHorizontalValues.Left, Width = 26, HeaderFontSize = 9)]
        public string 更新日 { get; set; }
    }
}
