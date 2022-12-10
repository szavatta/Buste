namespace Buste
{
    partial class fmBusteNonUscite
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
            this.dgBusteNonUscite = new System.Windows.Forms.DataGridView();
            this.lbTotBuste = new System.Windows.Forms.Label();
            this.lblTotBuste = new System.Windows.Forms.Label();
            this.CBEscludiLettura = new System.Windows.Forms.CheckBox();
            this.BRicarica = new System.Windows.Forms.Button();
            this.CBAData = new System.Windows.Forms.CheckBox();
            this.dtAData = new System.Windows.Forms.DateTimePicker();
            this.lblTotQta = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgBusteNonUscite)).BeginInit();
            this.SuspendLayout();
            // 
            // dgBusteNonUscite
            // 
            this.dgBusteNonUscite.AllowUserToAddRows = false;
            this.dgBusteNonUscite.AllowUserToDeleteRows = false;
            this.dgBusteNonUscite.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgBusteNonUscite.Location = new System.Drawing.Point(12, 28);
            this.dgBusteNonUscite.Name = "dgBusteNonUscite";
            this.dgBusteNonUscite.Size = new System.Drawing.Size(680, 366);
            this.dgBusteNonUscite.TabIndex = 3;
            this.dgBusteNonUscite.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgBusteNonUscite_CellContentClick);
            // 
            // lbTotBuste
            // 
            this.lbTotBuste.AutoSize = true;
            this.lbTotBuste.Location = new System.Drawing.Point(12, 6);
            this.lbTotBuste.Name = "lbTotBuste";
            this.lbTotBuste.Size = new System.Drawing.Size(78, 13);
            this.lbTotBuste.TabIndex = 4;
            this.lbTotBuste.Text = "Tot. buste/qta:";
            // 
            // lblTotBuste
            // 
            this.lblTotBuste.AutoSize = true;
            this.lblTotBuste.Location = new System.Drawing.Point(92, 6);
            this.lblTotBuste.Name = "lblTotBuste";
            this.lblTotBuste.Size = new System.Drawing.Size(13, 13);
            this.lblTotBuste.TabIndex = 5;
            this.lblTotBuste.Text = "0";
            // 
            // CBEscludiLettura
            // 
            this.CBEscludiLettura.AutoSize = true;
            this.CBEscludiLettura.Checked = true;
            this.CBEscludiLettura.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CBEscludiLettura.Location = new System.Drawing.Point(176, 4);
            this.CBEscludiLettura.Name = "CBEscludiLettura";
            this.CBEscludiLettura.Size = new System.Drawing.Size(103, 17);
            this.CBEscludiLettura.TabIndex = 6;
            this.CBEscludiLettura.Text = "Escludi in lettura";
            this.CBEscludiLettura.UseVisualStyleBackColor = true;
            // 
            // BRicarica
            // 
            this.BRicarica.Location = new System.Drawing.Point(596, 1);
            this.BRicarica.Name = "BRicarica";
            this.BRicarica.Size = new System.Drawing.Size(96, 23);
            this.BRicarica.TabIndex = 11;
            this.BRicarica.Text = "Ricarica";
            this.BRicarica.UseVisualStyleBackColor = true;
            this.BRicarica.Click += new System.EventHandler(this.BRicarica_Click);
            // 
            // CBAData
            // 
            this.CBAData.AutoSize = true;
            this.CBAData.Location = new System.Drawing.Point(293, 4);
            this.CBAData.Name = "CBAData";
            this.CBAData.Size = new System.Drawing.Size(79, 17);
            this.CBAData.TabIndex = 13;
            this.CBAData.Text = "Fino a data";
            this.CBAData.UseVisualStyleBackColor = true;
            this.CBAData.CheckedChanged += new System.EventHandler(this.CBAData_CheckedChanged);
            // 
            // dtAData
            // 
            this.dtAData.Enabled = false;
            this.dtAData.Location = new System.Drawing.Point(378, 2);
            this.dtAData.Name = "dtAData";
            this.dtAData.Size = new System.Drawing.Size(200, 20);
            this.dtAData.TabIndex = 14;
            // 
            // lblTotQta
            // 
            this.lblTotQta.AutoSize = true;
            this.lblTotQta.Location = new System.Drawing.Point(138, 6);
            this.lblTotQta.Name = "lblTotQta";
            this.lblTotQta.Size = new System.Drawing.Size(13, 13);
            this.lblTotQta.TabIndex = 16;
            this.lblTotQta.Text = "0";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(126, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(12, 13);
            this.label1.TabIndex = 17;
            this.label1.Text = "/";
            // 
            // fmBusteNonUscite
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(704, 406);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblTotQta);
            this.Controls.Add(this.dtAData);
            this.Controls.Add(this.CBAData);
            this.Controls.Add(this.BRicarica);
            this.Controls.Add(this.CBEscludiLettura);
            this.Controls.Add(this.lblTotBuste);
            this.Controls.Add(this.lbTotBuste);
            this.Controls.Add(this.dgBusteNonUscite);
            this.Name = "fmBusteNonUscite";
            this.Text = "Buste non uscite";
            this.Load += new System.EventHandler(this.fmBusteNonUscite_Load);
            this.Resize += new System.EventHandler(this.fmBusteNonUscite_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dgBusteNonUscite)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgBusteNonUscite;
        private System.Windows.Forms.Label lbTotBuste;
        private System.Windows.Forms.Label lblTotBuste;
        private System.Windows.Forms.CheckBox CBEscludiLettura;
        private System.Windows.Forms.Button BRicarica;
        private System.Windows.Forms.CheckBox CBAData;
        private System.Windows.Forms.DateTimePicker dtAData;
        private System.Windows.Forms.Label lblTotQta;
        private System.Windows.Forms.Label label1;
    }
}