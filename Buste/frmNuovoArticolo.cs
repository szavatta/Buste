using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Buste
{
    public partial class frmNuovoArticolo : Form
    {
        public frmNuovoArticolo(string codice, string descrizione)
        {
            InitializeComponent();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            this.Tag = tbCodiceBase.Text;
            this.Close();
        }
    }
}
