using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace scheDataDelete.Common
{
    class utility
    {
        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     文字型をIntへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     Int型の値</returns>
        /// --------------------------------------------------------------------------------
        public static int StrtoInt(string tempStr)
        {
            if (NumericCheck(tempStr)) return int.Parse(tempStr);
            else return 0;
        }
        
        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     文字列の値が数字かチェックする </summary>
        /// <param name="tempStr">
        ///     検証する文字列</param>
        /// <returns>
        ///     数字:true,数字でない:false</returns>
        /// ------------------------------------------------------------------------------
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }
        
    }
}
