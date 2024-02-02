using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

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
        ///-------------------------------------------------------------------------
        /// <summary>
        ///     Excel�t�@�C�����p�X���[�h�t���ŃI�[�v���E�N���[�Y���� </summary>
        /// <param name="sPath">
        ///     Excel�t�@�C���p�X</param>
        /// <param name="rPw">
        ///     �ǂݍ��݃p�X���[�h</param>
        /// <param name="wPw">
        ///     �������݃p�X���[�h</param>
        /// <param name="logFile">
        ///     ���O�t�@�C���p�X</param>
        /// <returns>
        ///     �����Ftrue, ���s�Ffalse</returns>
        ///-------------------------------------------------------------------------
        public static bool PwdXlsFile(string sPath, string rPw, string wPw, string logFile)
        {
            if (rPw == string.Empty)
            {
                return true;
            }

            // ���O�o��
            System.IO.File.AppendAllText(logFile, Form1.GetNowTime(" Excel���N�����Ă��܂�..."), Encoding.GetEncoding(932));

            System.Threading.Thread.Sleep(100);
            Application.DoEvents();

            // �G�N�Z���I�u�W�F�N�g
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = null;

            try
            {
                //if (wPw != string.Empty)
                //{
                //    lblMsg.Text = sPath + " �̃p�X���[�h���������Ă��܂�...";
                //}
                //else
                //{
                //    lblMsg.Text = sPath + " ���J���Ă��܂�...";
                //}

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                // Excel�t�@�C�����J��
                oXlsBook = (oXls.Workbooks.Open(sPath, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, wPw, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                    //if (wPw != string.Empty)
                    //{
                    //    lblMsg.Text = sPath + " �̃p�X���[�h����������܂���...";
                    //}
                    //else
                    //{
                    //    lblMsg.Text = sPath + " ���J���܂���...";
                    //}

                    System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                oXls.DisplayAlerts = false;

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                //// Excel�t�@�C����������
                //oXlsBook.SaveAs(sPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, rPw,
                //                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                //                Type.Missing, Type.Missing);

                // Excel�t�@�C����������
                oXlsBook.SaveAs(sPath, Type.Missing, rPw, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);

                //lblMsg.Text = sPath + " ��ۑ����܂���...";

                // ���O�o��
                if (rPw != string.Empty)
                {
                    System.IO.File.AppendAllText(logFile, Form1.GetNowTime(" " + sPath + " ���p�X���[�h�t���ŕۑ����܂���..."), Encoding.GetEncoding(932));
                }

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                // Book���N���[�Y
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                //lblMsg.Text = "Excel���I�����܂���...";
                System.IO.File.AppendAllText(logFile, Form1.GetNowTime(" Excel���I�����܂���..."), Encoding.GetEncoding(932));

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {
                // Excel���I��
                oXls.Quit();

                // COM �I�u�W�F�N�g�̎Q�ƃJ�E���g��������� 
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                if (oXlsBook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;

                GC.Collect();
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     ��������w�蕶�������l�`�w�Ƃ��ĕԂ��܂�</summary>
        /// <param name="s">
        ///     ������</param>
        /// <param name="n">
        ///     ������</param>
        /// <returns>
        ///     �������͈͓��̕�����</returns>
        /// --------------------------------------------------------------------
        public static string GetStringSubMax(string s, int n)
        {
            string val;

            // �����Ԃ̃X�y�[�X������ 2015/03/10
            s = s.Replace(" ", "");

            if (s.Length > n) val = s.Substring(0, n);
            else val = s;

            return val;
        }
    }
}
