using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace scheTask
{
    class Control
    {
        /// <summary>
        /// DataControl�N���X�̊�{�N���X
        /// </summary>
        public class BaseControl
        {
            private Utility.DBConnect DBConnect;

            //BaseControl�̃R���X�g���N�^�BDBConnect�N���X�̃C���X�^���X���쐬���܂��B
            public BaseControl()
            {
                DBConnect = new Utility.DBConnect();
            }

            public OleDbConnection GetConnection()
            {
                return DBConnect.Cn;
            }
        }

        public class DataControl : BaseControl
        {
            public OleDbConnection Cn = new OleDbConnection();

            //�f�[�^�R���g���[���N���X�̃R���X�g���N�^
            public DataControl()
            {
            }

            /// <summary>
            /// �f�[�^�x�[�X�ڑ�����
            /// </summary>
            public void Close()
            {
                if (Cn.State == ConnectionState.Open)
                {
                    Cn.Close();
                }
            }

            /// -------------------------------------------------------------
            /// <summary>
            ///     �f�[�^���[�_�[���擾���� </summary>
            /// <param name="tempSQL">
            ///     SQL��    </param>
            /// <returns>
            ///     �f�[�^���[�_�[</returns>
            /// -------------------------------------------------------------
            public OleDbDataReader FreeReader(string tempSQL)
            {
                Cn = GetConnection();
                OleDbCommand sCom = new OleDbCommand();
                sCom.Connection = Cn;
                sCom.CommandText = tempSQL;
                return sCom.ExecuteReader();
            }
        }
    }
}
