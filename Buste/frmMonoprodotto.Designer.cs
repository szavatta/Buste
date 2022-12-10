namespace Buste
{
    partial class frmMonoprodotto
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dgMonoprodotto = new System.Windows.Forms.DataGridView();
            this.Data = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Categoria = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Totale = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgMonoprodotto)).BeginInit();
            this.SuspendLayout();
            // 
            // dgMonoprodotto
            // 
            this.dgMonoprodotto.AllowUserToAddRows = false;
            this.dgMonoprodotto.AllowUserToDeleteRows = false;
            this.dgMonoprodotto.AllowUserToOrderColumns = true;
            this.dgMonoprodotto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.dgMonoprodotto.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgMonoprodotto.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Data,
            this.Categoria,
            this.Totale});
            this.dgMonoprodotto.Location = new System.Drawing.Point(12, 12);
            this.dgMonoprodotto.Name = "dgMonoprodotto";
            this.dgMonoprodotto.Size = new System.Drawing.Size(352, 247);
            this.dgMonoprodotto.TabIndex = 0;
            // 
            // Data
            // 
            this.Data.DataPropertyName = "data";
            this.Data.HeaderText = "Data";
            this.Data.Name = "Data";
            this.Data.ReadOnly = true;
            this.Data.Width = 80;
            // 
            // Categoria
            // 
            this.Categoria.DataPropertyName = "categoria";
            this.Categoria.HeaderText = "Categoria";
            this.Categoria.Name = "Categoria";
            this.Categoria.ReadOnly = true;
            this.Categoria.Width = 150;
            // 
            // Totale
            // 
            this.Totale.DataPropertyName = "num";
            this.Totale.HeaderText = "Totale";
            this.Totale.Name = "Totale";
            this.Totale.ReadOnly = true;
            this.Totale.Width = 50;
            // 
            // frmMonoprodotto
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(376, 271);
            this.Controls.Add(this.dgMonoprodotto);
            this.Name = "frmMonoprodotto";
            this.Text = "Monoprodotto";
            this.Load += new System.EventHandler(this.frmMonoprodotto_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgMonoprodotto)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgMonoprodotto;
        private System.Windows.Forms.DataGridViewTextBoxColumn Data;
        private System.Windows.Forms.DataGridViewTextBoxColumn Categoria;
        private System.Windows.Forms.DataGridViewTextBoxColumn Totale;
    }
}