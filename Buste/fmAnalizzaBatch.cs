using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Buste
{
    public partial class fmAnalizzaBatch : Form
    {
        protected Access boAccess;

        public fmAnalizzaBatch()
        {
            InitializeComponent();
        }

        private void fmAnalizzaBusta_Load(object sender, EventArgs e)
        {
            boAccess = new Access();

            tbCodiceBatch.Select();
            tbCodiceBatch.Focus();
            inizializzaDgMovimentiBatch();
        }

        private void inizializzaDgMovimentiBatch()
        {
            dgMovimentiBatch.Columns.Add("busta", "busta");
            dgMovimentiBatch.Columns["busta"].Width = 100;
            dgMovimentiBatch.Columns.Add("idArticolo", "idArticolo");
            dgMovimentiBatch.Columns["idArticolo"].Width = 60;
            dgMovimentiBatch.Columns.Add("descrizione", "descrizione");
            dgMovimentiBatch.Columns["descrizione"].Width = 200;
            dgMovimentiBatch.Columns.Add("qta", "qta");
            dgMovimentiBatch.Columns["qta"].Width = 50;
            dgMovimentiBatch.Columns.Add("data", "dataIn");
            dgMovimentiBatch.Columns["data"].Width = 80;
            dgMovimentiBatch.Columns.Add("dataout", "dataOut");
            dgMovimentiBatch.Columns["dataout"].Width = 80;
            dgMovimentiBatch.Columns.Add("tipoOrdine", "tipoOrdine");
            dgMovimentiBatch.Columns["tipoOrdine"].Width = 80;
            dgMovimentiBatch.Columns.Add("numProdotti", "numProdotti");
            dgMovimentiBatch.Columns["numProdotti"].Width = 80;
        }
                                            
        private void AggiungeMovimentoBatch(string idBusta, string idArticolo, string descrArticolo, string dataMov, int qtaCodice, string dataOut, string tipoOrdine= "", string numProdotti = "")
        {
            object[] objMov = new object[] { idBusta, idArticolo, descrArticolo, qtaCodice, dataMov, dataOut, tipoOrdine, numProdotti };
            dgMovimentiBatch.Rows.Add(objMov);
        }

        private void bOk_Click(object sender, EventArgs e)
        {
            dgMovimentiBatch.Rows.Clear();
            string codBatch = tbCodiceBatch.Text;
            if(tbCodiceBatch.Text != string.Empty)
            {
                if(codBatch.StartsWith("F") && codBatch.StartsWith("F"))
                    codBatch = codBatch.Replace("F", "");
                DataTable dbBatch = boAccess.ListaMovimentiBatch(codBatch);
                foreach (DataRow dr in dbBatch.Rows)
                {
                    string dataIn = string.Empty;
                    try { dataIn = Convert.ToDateTime(dr["dataIn"]).ToShortDateString(); }
                    catch { }
                    string dataOut = string.Empty;
                    try { dataOut = Convert.ToDateTime(dr["dataOut"]).ToShortDateString(); }
                    catch { }
                    AggiungeMovimentoBatch(dr["idBusta"].ToString(), dr["idArticolo"].ToString(), dr["descrizione"].ToString(), dataIn, Convert.ToInt32(dr["qta"]), dataOut, dr["tipoOrdine"].ToString(), dr["numProdotti"].ToString() == "MONOPRODOTTO" ? dr["numProdotti"].ToString() : "");
                }

                dgMovimentiBatch.Sort(dgMovimentiBatch.Columns["data"], ListSortDirection.Ascending);
            }
            tbCodiceBatch.Select(0,tbCodiceBatch.Text.Length);
            tbCodiceBatch.Focus();
        }

        private void fmAnalizzaBusta_Resize(object sender, EventArgs e)
        {
            try
            {
                dgMovimentiBatch.Width = fmBuste.ActiveForm.Width - 28;
                dgMovimentiBatch.Height = fmBuste.ActiveForm.Height - 109;
            }
            catch
            {
            }
        }

    }
}