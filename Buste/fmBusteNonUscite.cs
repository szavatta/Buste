using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Buste
{
    public partial class fmBusteNonUscite : Form
    {
        protected Access boAccess;
        public Form frmBuste;

        public fmBusteNonUscite(ref Form frm)
        {
            InitializeComponent();
        }

        private void fmBusteNonUscite_Load(object sender, EventArgs e)
        {
            boAccess = new Access();

            inizializzaDgBusteNonUscite();
            int totQta;
            caricaDgBusteNonUscite(out totQta);
            lblTotBuste.Text = dgBusteNonUscite.Rows.Count.ToString();
            lblTotQta.Text = totQta.ToString();
        }

        private void inizializzaDgBusteNonUscite()
        {
            DataGridViewButtonColumn buttons = new DataGridViewButtonColumn();
            buttons.HeaderText = "Vis";
            buttons.UseColumnTextForButtonValue = true;
            buttons.FlatStyle = FlatStyle.Standard;
            buttons.CellTemplate.Style.BackColor = Color.Honeydew;
            buttons.DisplayIndex = 0;
            buttons.Width = 30;

            dgBusteNonUscite.Columns.Add("idBusta", "Busta");
            dgBusteNonUscite.Columns["idBusta"].Width = 80;
            dgBusteNonUscite.Columns.Add(buttons);
            dgBusteNonUscite.Columns.Add("data", "Data");
            dgBusteNonUscite.Columns["data"].Width = 80;
            dgBusteNonUscite.Columns["data"].ValueType = typeof(DateTime);
            dgBusteNonUscite.Columns.Add("idArticolo", "Articolo");
            dgBusteNonUscite.Columns["idArticolo"].Width = 60;
            dgBusteNonUscite.Columns.Add("descrizione", "Descrizione");
            dgBusteNonUscite.Columns["descrizione"].Width = 200;
            dgBusteNonUscite.Columns.Add("qta", "Qta");
            dgBusteNonUscite.Columns["qta"].Width = 50;
            dgBusteNonUscite.Columns.Add("idBatch", "Batch");
            dgBusteNonUscite.Columns["idBatch"].Width = 60;
            dgBusteNonUscite.Columns.Add("inLettura", "In lettura");
            dgBusteNonUscite.Columns["inLettura"].Width = 60;
            dgBusteNonUscite.Columns["inLettura"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void caricaDgBusteNonUscite(out int totQta)
        {
            dgBusteNonUscite.Rows.Clear();
            totQta = 0;
            DataTable dbBuste = boAccess.ListaBusteNonUscite(null);
            foreach (DataRow dr in dbBuste.Rows)
            {
                string inlettura = string.Empty;
                foreach (DataGridViewRow drBuste in ((fmBuste)frmBuste).dgOutMovimenti.Rows)
                {
                    if (drBuste.Cells["idBusta"].Value.ToString() == dr["idBusta"].ToString()
                        && drBuste.Cells["idArticolo"].Value.ToString() == dr["idArticolo"].ToString()
                        && drBuste.Cells["qta"].Value.ToString() == dr["qta"].ToString())
                    {
                        inlettura = "X";
                    }
                }
                DateTime data = Convert.ToDateTime(dr["data"]);
                int qta = Convert.ToInt32(dr["qta"]);
                if (!(CBEscludiLettura.Checked && inlettura == "X" 
                    || CBAData.Checked && data > dtAData.Value))
                {
                    object[] objMov = new object[] { dr["idBusta"].ToString(), null, data, dr["idArticolo"].ToString(), dr["descrizione"].ToString(), qta, dr["idBatch"].ToString(), inlettura };
                    dgBusteNonUscite.Rows.Add(objMov);
                }
                totQta += qta;
            }
        }

        private void fmBusteNonUscite_Resize(object sender, EventArgs e)
        {
            try
            {
                dgBusteNonUscite.Width = fmBuste.ActiveForm.Width - 28;
                dgBusteNonUscite.Height = fmBuste.ActiveForm.Height - 74;
            }
            catch
            {
            }
        }

        private void CBAData_CheckedChanged(object sender, EventArgs e)
        {
            dtAData.Enabled = CBAData.Checked;
        }

        private void BRicarica_Click(object sender, EventArgs e)
        {
            int totQta;
            caricaDgBusteNonUscite(out totQta);
            lblTotBuste.Text = dgBusteNonUscite.Rows.Count.ToString();
            lblTotQta.Text = totQta.ToString();
        }

        private void dgBusteNonUscite_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                string idBusta = dgBusteNonUscite.Rows[e.RowIndex].Cells["idBusta"].Value.ToString();

                Form formBuste = this;
                fmAnalizzaBusta fmBusta = new fmAnalizzaBusta(ref formBuste);
                fmBusta.frmBuste = frmBuste;
                fmBusta.Tag = idBusta;
                fmBusta.Show();
            }
        }
    }
}