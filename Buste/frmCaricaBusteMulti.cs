using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Buste
{
    public partial class frmCaricaBusteMulti : Form
    {
        public frmCaricaBusteMulti()
        {
            InitializeComponent();
        }

        private void tbBuste_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnCarica_Click(object sender, EventArgs e)
        {
            string elenco = tbBuste.Text;
            foreach (string item in tbBuste.Text.Split('\n'))
            {
                string busta = item.Split(',')[0];
                int qta = 1;
                try
                {
                    qta = Convert.ToInt32(item.Split(',')[1]);
                }
                catch { }
                
                
                
            }
        }
    }
}