using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Buste
{
    public partial class fmConteggiaArticolo : Form
    {
        protected Access boAccess;
        public Form frmBuste;

        public fmConteggiaArticolo(ref Form frm)
        {
            InitializeComponent();
            
        }

        private void fmAnalizzaBusta_Load(object sender, EventArgs e)
        {
            boAccess = new Access();

            //tbCodiceBusta.Select();
            //tbCodiceBusta.Focus();
            inizializzaDgMovimentiBusta();
            if (this.Tag != null)
            {
                //tbCodiceBusta.Text = this.Tag.ToString();
                bOk_Click(null, null);
            }
        }

        private void inizializzaDgMovimentiBusta()
        {
            dgTotaliArticoli.Columns.Add("idArticolo", "Articolo");
            dgTotaliArticoli.Columns["idArticolo"].Width = 60;
            dgTotaliArticoli.Columns.Add("descrizione", "Descrizione");
            dgTotaliArticoli.Columns["descrizione"].Width = 300;
            dgTotaliArticoli.Columns.Add("totale", "Totale");
            dgTotaliArticoli.Columns["totale"].Width = 70;
            dgTotaliArticoli.Columns.Add("importo", "Importo");
            dgTotaliArticoli.Columns["importo"].Width = 70;
        }

        private void AggiungeMovimentoBusta(string tipo, string idArticolo, string descrArticolo, DateTime dataMov, string numBolla, int qtaCodice, string idBatch)
        {
            object[] objMov = new object[] { tipo, dataMov.ToShortDateString(), idArticolo, descrArticolo, qtaCodice, idBatch, numBolla };
            dgTotaliArticoli.Rows.Add(objMov);
        }

        private void bOk_Click(object sender, EventArgs e)
        {
            dgTotaliArticoli.Rows.Clear();
            if (dtDaData.Text != string.Empty)
            {
                DataTable dbTotali = boAccess.TotaliArticoli(dtDaData.Value.Date, dtAData.Value.Date, tbArticolo.Text, tbBolla.Text, ckNonRaggruppare.Checked);
                if (dbTotali != null)
                {
                    foreach (DataRow dr in dbTotali.Rows)
                    {
                        object[] objMov = new object[] { dr["idArticolo"], dr["descrizione"], dr["totale"], dr["importo"] };
                        dgTotaliArticoli.Rows.Add(objMov);
                    }
                }

                //dgMovimentiBusta.Sort(dgMovimentiBusta.Columns["data"], ListSortDirection.Ascending);
            }

        }

        private void fmAnalizzaBusta_Resize(object sender, EventArgs e)
        {
            try
            {
                dgTotaliArticoli.Width = fmBuste.ActiveForm.Width - 28;
                dgTotaliArticoli.Height = fmBuste.ActiveForm.Height - 109;
            }
            catch
            {
            }
        }

        private void dgMovimentiBusta_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            bOk_Click(null, null);
        }

    }
}