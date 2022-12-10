using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Buste
{
    public partial class fmAnalizzaBusta : Form
    {
        protected Access boAccess;
        public Form frmBuste;

        public fmAnalizzaBusta(ref Form frm)
        {
            InitializeComponent();
            
        }

        private void fmAnalizzaBusta_Load(object sender, EventArgs e)
        {
            boAccess = new Access();

            tbCodiceBusta.Select();
            tbCodiceBusta.Focus();
            inizializzaDgMovimentiBusta();
            if (this.Tag != null)
            {
                tbCodiceBusta.Text = this.Tag.ToString();
                bOk_Click(null, null);
            }
        }

        private void inizializzaDgMovimentiBusta()
        {
            dgMovimentiBusta.Columns.Add("tipo", "Tipo");
            dgMovimentiBusta.Columns["tipo"].Width = 50;
            dgMovimentiBusta.Columns.Add("data", "Data");
            dgMovimentiBusta.Columns["data"].Width = 80;
            dgMovimentiBusta.Columns.Add("idArticolo", "Articolo");
            dgMovimentiBusta.Columns["idArticolo"].Width = 60;
            dgMovimentiBusta.Columns.Add("descrizione", "Descrizione");
            dgMovimentiBusta.Columns["descrizione"].Width = 200;
            dgMovimentiBusta.Columns.Add("qta", "Qta");
            dgMovimentiBusta.Columns["qta"].Width = 50;
            dgMovimentiBusta.Columns.Add("batch", "Batch");
            dgMovimentiBusta.Columns["batch"].Width = 100;
            dgMovimentiBusta.Columns.Add("numBolla", "NumBolla");
            dgMovimentiBusta.Columns["numBolla"].Width = 60;
            dgMovimentiBusta.Columns.Add("tipoOrdine", "tipoOrdine");
            dgMovimentiBusta.Columns["tipoOrdine"].Width = 80;
            dgMovimentiBusta.Columns.Add("numProdotti", "numProdotti");
            dgMovimentiBusta.Columns["numProdotti"].Width = 80;
        }

        private void AggiungeMovimentoBusta(string tipo, string idArticolo, string descrArticolo, DateTime dataMov, string numBolla, int qtaCodice, string idBatch, string tipoOrdine = "", string numProdotti = "")
        {
            object[] objMov = new object[] { tipo, dataMov.ToShortDateString(), idArticolo, descrArticolo, qtaCodice, idBatch, numBolla, tipoOrdine, numProdotti };
            dgMovimentiBusta.Rows.Add(objMov);
        }

        private void bOk_Click(object sender, EventArgs e)
        {
            dgMovimentiBusta.Rows.Clear();

            string fornitore = boAccess.leggiAppSettings("Fornitore");
            string busta = tbCodiceBusta.Text;
            bool isMascherina = false;
            if (busta.StartsWith("F") && busta.EndsWith("F"))
            {
                busta = busta.Replace("F", "");
                isMascherina = true;
            }
            if (fornitore == "FotoEvolution" && busta.Length > 6)
                busta = Convert.ToString(Convert.ToInt64(busta.Substring(0, 11)));

            if (busta != string.Empty)
            {
                DataTable dbBustaIn = boAccess.ListaMovimentiBusta(busta, "In", isMascherina);
                DataTable dbBustaOut = boAccess.ListaMovimentiBusta(busta, "Out", isMascherina);
                DataTable dbBustaResi = boAccess.ListaMovimentiBusta(busta, "Resi", isMascherina);
                foreach (DataRow dr in dbBustaIn.Rows)
                {
                    AggiungeMovimentoBusta("In", dr["idArticolo"].ToString(), dr["descrArticolo"].ToString(), Convert.ToDateTime(dr["data"]), dr["numBolla"].ToString(), Convert.ToInt32(dr["qta"]), dr["idBatch"].ToString(), dr["tipoOrdine"].ToString(), dr["numProdotti"].ToString() == "MONOPRODOTTO" ? dr["numProdotti"].ToString() : "");
                }
                foreach (DataRow dr in dbBustaOut.Rows)
                {
                    AggiungeMovimentoBusta("Out", dr["idArticolo"].ToString(), dr["descrArticolo"].ToString(), Convert.ToDateTime(dr["data"]), dr["numBolla"].ToString(), Convert.ToInt32(dr["qta"]), dr["idBatch"].ToString(), dr["tipoOrdine"].ToString(), dr["numProdotti"].ToString() == "MONOPRODOTTO" ? dr["numProdotti"].ToString() : "");
                }
                foreach (DataRow dr in dbBustaResi.Rows)
                {
                    AggiungeMovimentoBusta("Resi", dr["idArticolo"].ToString(), dr["descrArticolo"].ToString(), Convert.ToDateTime(dr["data"]), dr["numBolla"].ToString(), Convert.ToInt32(dr["qta"]), dr["idBatch"].ToString());
                }
                foreach (DataGridViewRow dr in ((fmBuste)frmBuste).dgOutMovimenti.Rows)
                {
                    if (dr.Cells["idBusta"].Value.ToString() == busta)
                    {
                        AggiungeMovimentoBusta("In lettura", dr.Cells["idArticolo"].Value.ToString(), dr.Cells["descrizione"].Value.ToString().Trim(), Convert.ToDateTime(dr.Cells["data"].Value), dr.Cells["numBolla"].Value.ToString(), Convert.ToInt32(dr.Cells["qta"].Value), dr.Cells["batch"].Value.ToString());
                    }
                }
                if (dgMovimentiBusta.Rows.Count == 0)
                {
                    MessageBox.Show("Codice busta non trovato", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                //dgMovimentiBusta.Sort(dgMovimentiBusta.Columns["data"], ListSortDirection.Ascending);
            }
            tbCodiceBusta.Select(0, tbCodiceBusta.Text.Length);
            tbCodiceBusta.Focus();
        }

        private void fmAnalizzaBusta_Resize(object sender, EventArgs e)
        {
            try
            {
                dgMovimentiBusta.Width = fmBuste.ActiveForm.Width - 28;
                dgMovimentiBusta.Height = fmBuste.ActiveForm.Height - 109;
            }
            catch
            {
            }
        }

        private void dgMovimentiBusta_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            string tipo = e.Row.Cells["tipo"].Value.ToString();
            if (tipo == "In lettura")
            {
                MessageBox.Show("Non è possibile cancellare buste in lettura");
                return;
            }

            string msg = "Conferma cancellazione?";
            DialogResult dr = MessageBox.Show(msg, "", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (dr == DialogResult.Yes)
            {
                string idBusta = tbCodiceBusta.Text;
                string idArticolo = e.Row.Cells["idArticolo"].Value.ToString();
                int qta = Convert.ToInt32(e.Row.Cells["qta"].Value);
                int ret = boAccess.CancellaRighe(tipo, idBusta, idArticolo, qta);
            }
        }

        private void dgMovimentiBusta_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            bOk_Click(null, null);
        }

    }
}