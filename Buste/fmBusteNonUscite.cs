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
            DataGridViewButtonColumn button1 = new DataGridViewButtonColumn();
            button1.Text = "?";
            button1.ToolTipText = "Visualizza le entrate e uscite di questo articolo";
            button1.UseColumnTextForButtonValue = true;
            button1.FlatStyle = FlatStyle.Standard;
            button1.CellTemplate.Style.BackColor = Color.Honeydew;
            button1.DisplayIndex = 0;
            button1.Width = 30;

            DataGridViewButtonColumn button2 = new DataGridViewButtonColumn();
            button2.Text = "+";
            button2.ToolTipText = "Aggiunge negli articoli in uscita";
            button2.UseColumnTextForButtonValue = true;
            button2.FlatStyle = FlatStyle.Standard;
            button2.CellTemplate.Style.BackColor = Color.Honeydew;
            button2.DisplayIndex = 0;
            button2.Width = 30;

            dgBusteNonUscite.Columns.Add("idBusta", "Busta");
            dgBusteNonUscite.Columns["idBusta"].Width = 80;
            dgBusteNonUscite.Columns.Add(button1);
            dgBusteNonUscite.Columns.Add(button2);
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
                    object[] objMov = new object[] { dr["idBusta"].ToString(), null, null, data, dr["idArticolo"].ToString(), dr["descrizione"].ToString(), qta, dr["idBatch"].ToString(), inlettura };
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
            else if (e.ColumnIndex == 2)
            {
                var row = dgBusteNonUscite.Rows[e.RowIndex];
                ((fmBuste)frmBuste).tbOutCodice.Text = row.Cells["idBusta"].Value.ToString();
                ((fmBuste)frmBuste).btOutInserisce.PerformClick();

                ((fmBuste)frmBuste).tbOutCodice.Text = row.Cells["idArticolo"].Value.ToString() + "*" + row.Cells["qta"].Value.ToString();
                ((fmBuste)frmBuste).btOutInserisce.PerformClick();

                this.Focus();


            }
        }
    }
}