using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;

namespace jfgSchedule
{
    class Utility
    {
        public const int configKey = 1;     // �ғ��ݒ�e�[�u���F���R�[�h�L�[

        public class DBConnect
        {
            OleDbConnection cn = new OleDbConnection();

            public OleDbConnection Cn
            {
                get
                {
                    return cn;
                }
            }

            private string sServerName;
            private string sLogin;
            private string sPass;
            private string sDatabase;

            public DBConnect()
            {
                try
                {
                    // MySeting���ڂ̎擾
                    // �T�[�o��
                    sServerName = Properties.Settings.Default.ServerName;

                    // ���O�C����
                    sLogin = Properties.Settings.Default.Login;

                    // �p�X���[�h
                    sPass = Properties.Settings.Default.Pass;

                    // �f�[�^�x�[�X��
                    sDatabase = Properties.Settings.Default.Database;

                    // �f�[�^�x�[�X�ڑ�������
                    cn.ConnectionString = "";
                    cn.ConnectionString += "Provider=SQLOLEDB;";
                    cn.ConnectionString += "SERVER=" + sServerName + ";";
                    cn.ConnectionString += "DataBase=" + sDatabase + ";";
                    cn.ConnectionString += "UID=" + sLogin + ";";
                    cn.ConnectionString += "PWD=" + sPass + ";";
                    //cn.ConnectionString += "WSID=";

                    cn.Open();

                }

                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        /// <summary>
        /// ������̒l���������`�F�b�N����
        /// </summary>
        /// <param name="tempStr">���؂��镶����</param>
        /// <returns>����:true,�����łȂ�:false</returns>
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     �J�[�h�ԍ����� �idouble�^�ϊ��\�Ȓl�����؁j</summary>
        /// <param name="cel">
        ///     �Z���̒l</param>
        /// <returns>
        ///     double�^�̂Ƃ��l��Ԃ��Adouble�^�ϊ��G���[�̂Ƃ�-1��Ԃ�</returns>
        ///-------------------------------------------------------------------------
        public static double cNumberCheck(string cel)
        {
            double rtn = -1;
            double cNo;
            if (double.TryParse(cel, out cNo))
            {
                rtn = cNo;
            }

            return rtn;
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     �I�u�W�F�N�g��string�^�ɕϊ����ĕԂ��@</summary>
        /// <param name="obj">
        ///     �Ώۂ̃I�u�W�F�N�g</param>
        /// <returns>
        ///     string�^�̖߂�l</returns>
        /// -------------------------------------------------------------------------
        public static string nulltoString(object obj)
        {
            string sVal = string.Empty;

            if (obj == null)
            {
                sVal = string.Empty;
            }
            else
            {
                sVal = obj.ToString();
            }

            return sVal;
        }
    }
}
