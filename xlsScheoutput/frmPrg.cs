using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace xlsScheoutput
{
    public partial class frmPrg : Form
    {
        public frmPrg(int pMax, int pMin)
        {
            InitializeComponent();
            prgBar.Maximum = pMax;
            prgBar.Minimum = pMin;
        }

        private void frmPrg_Load(object sender, EventArgs e)
        {

        }

        public int progressValue { get; set; }

        public void ProgressStep()
        {
            prgBar.Value = progressValue;
        }

        private void frmPrg_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }
    }
}
