namespace Buste
{
    partial class fmAnalizzaBatch
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
            this.tbCodiceBatch = new System.Windows.Forms.TextBox();
            this.dgMovimentiBatch = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgMovimentiBatch)).BeginInit();
            this.SuspendLayout();
            // 
            // bOk
            // 
            this.bOk.Location = new System.Drawing.Point(226, 21);
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
            this.label1.Location = new System.Drawing.Point(12, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Codice batch";
            // 
            // tbCodiceBatch
            // 
            this.tbCodiceBatch.Location = new System.Drawing.Point(87, 23);
            this.tbCodiceBatch.Name = "tbCodiceBatch";
            this.tbCodiceBatch.Size = new System.Drawing.Size(133, 20);
            this.tbCodiceBatch.TabIndex = 2;
            // 
            // dgMovimentiBatch
            // 
            this.dgMovimentiBatch.AllowUserToAddRows = false;
            this.dgMovimentiBatch.AllowUserToDeleteRows = false;
            this.dgMovimentiBatch.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgMovimentiBatch.Location = new System.Drawing.Point(12, 63);
            this.dgMovimentiBatch.Name = "dgMovimentiBatch";
            this.dgMovimentiBatch.Size = new System.Drawing.Size(791, 331);
            this.dgMovimentiBatch.TabIndex = 3;
            // 
            // fmAnalizzaBatch
            // 
            this.AcceptButton = this.bOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(815, 406);
            this.Controls.Add(this.dgMovimentiBatch);
            this.Controls.Add(this.tbCodiceBatch);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.bOk);
            this.Name = "fmAnalizzaBatch";
            this.Text = "Analizza batch";
            this.Load += new System.EventHandler(this.fmAnalizzaBusta_Load);
            this.Resize += new System.EventHandler(this.fmAnalizzaBusta_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dgMovimentiBatch)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bOk;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbCodiceBatch;
        private System.Windows.Forms.DataGridView dgMovimentiBatch;
    }
}