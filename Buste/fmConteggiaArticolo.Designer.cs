namespace Buste
{
    partial class fmConteggiaArticolo
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
            this.bOk = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.dgTotaliArticoli = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.dtDaData = new System.Windows.Forms.DateTimePicker();
            this.dtAData = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.tbArticolo = new System.Windows.Forms.TextBox();
            this.tbBolla = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.ckNonRaggruppare = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgTotaliArticoli)).BeginInit();
            this.SuspendLayout();
            // 
            // bOk
            // 
            this.bOk.Location = new System.Drawing.Point(592, 14);
            this.bOk.Name = "bOk";
            this.bOk.Size = new System.Drawing.Size(75, 23);
            this.bOk.TabIndex = 0;
            this.bOk.Text = "OK";
            this.bOk.UseVisualStyleBackColor = true;
            this.bOk.Click += new System.EventHandler(this.bOk_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Da data";
            // 
            // dgTotaliArticoli
            // 
            this.dgTotaliArticoli.AllowUserToAddRows = false;
            this.dgTotaliArticoli.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgTotaliArticoli.Location = new System.Drawing.Point(12, 63);
            this.dgTotaliArticoli.Name = "dgTotaliArticoli";
            this.dgTotaliArticoli.Size = new System.Drawing.Size(655, 331);
            this.dgTotaliArticoli.TabIndex = 3;
            this.dgTotaliArticoli.UserDeletedRow += new System.Windows.Forms.DataGridViewRowEventHandler(this.dgMovimentiBusta_UserDeletedRow);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(179, 17);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "A data";
            // 
            // dtDaData
            // 
            this.dtDaData.CustomFormat = "";
            this.dtDaData.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtDaData.Location = new System.Drawing.Point(63, 14);
            this.dtDaData.Name = "dtDaData";
            this.dtDaData.Size = new System.Drawing.Size(104, 20);
            this.dtDaData.TabIndex = 15;
            // 
            // dtAData
            // 
            this.dtAData.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtAData.Location = new System.Drawing.Point(223, 13);
            this.dtAData.Name = "dtAData";
            this.dtAData.Size = new System.Drawing.Size(109, 20);
            this.dtAData.TabIndex = 16;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(348, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(42, 13);
            this.label3.TabIndex = 17;
            this.label3.Text = "Articolo";
            // 
            // tbArticolo
            // 
            this.tbArticolo.Location = new System.Drawing.Point(393, 13);
            this.tbArticolo.Name = "tbArticolo";
            this.tbArticolo.Size = new System.Drawing.Size(70, 20);
            this.tbArticolo.TabIndex = 18;
            // 
            // tbBolla
            // 
            this.tbBolla.Location = new System.Drawing.Point(514, 13);
            this.tbBolla.Name = "tbBolla";
            this.tbBolla.Size = new System.Drawing.Size(51, 20);
            this.tbBolla.TabIndex = 20;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(480, 17);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(30, 13);
            this.label4.TabIndex = 19;
            this.label4.Text = "Bolla";
            // 
            // ckNonRaggruppare
            // 
            this.ckNonRaggruppare.AutoSize = true;
            this.ckNonRaggruppare.Location = new System.Drawing.Point(397, 40);
            this.ckNonRaggruppare.Name = "ckNonRaggruppare";
            this.ckNonRaggruppare.Size = new System.Drawing.Size(106, 17);
            this.ckNonRaggruppare.TabIndex = 21;
            this.ckNonRaggruppare.Text = "Non raggruppare";
            this.ckNonRaggruppare.UseVisualStyleBackColor = true;
            // 
            // fmConteggiaArticolo
            // 
            this.AcceptButton = this.bOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(679, 406);
            this.Controls.Add(this.ckNonRaggruppare);
            this.Controls.Add(this.tbBolla);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.tbArticolo);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dtAData);
            this.Controls.Add(this.dtDaData);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dgTotaliArticoli);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.bOk);
            this.Name = "fmConteggiaArticolo";
            this.Text = "Conteggia articoli";
            this.Load += new System.EventHandler(this.fmAnalizzaBusta_Load);
            this.Resize += new System.EventHandler(this.fmAnalizzaBusta_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dgTotaliArticoli)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bOk;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dgTotaliArticoli;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtDaData;
        private System.Windows.Forms.DateTimePicker dtAData;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbArticolo;
        private System.Windows.Forms.TextBox tbBolla;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox ckNonRaggruppare;
    }
}