using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Buste
{
    public partial class frmMonoprodotto : Form
    {
        protected Access boAccess;

        public frmMonoprodotto()
        {
            InitializeComponent();
        }

        private void frmMonoprodotto_Load(object sender, EventArgs e)
        {
            boAccess = new Access();

            DataTable dt = boAccess.TotaliMonoprodotto();
            dgMonoprodotto.DataSource = dt;
        }
    }
}
