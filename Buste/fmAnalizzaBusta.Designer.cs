namespace Buste
{
    partial class fmAnalizzaBusta
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
            this.tbCodiceBusta = new System.Windows.Forms.TextBox();
            this.dgMovimentiBusta = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgMovimentiBusta)).BeginInit();
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
            this.label1.Size = new System.Drawing.Size(69, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Codice busta";
            // 
            // tbCodiceBusta
            // 
            this.tbCodiceBusta.Location = new System.Drawing.Point(87, 23);
            this.tbCodiceBusta.Name = "tbCodiceBusta";
            this.tbCodiceBusta.Size = new System.Drawing.Size(133, 20);
            this.tbCodiceBusta.TabIndex = 2;
            // 
            // dgMovimentiBusta
            // 
            this.dgMovimentiBusta.AllowUserToAddRows = false;
            this.dgMovimentiBusta.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgMovimentiBusta.Location = new System.Drawing.Point(12, 63);
            this.dgMovimentiBusta.Name = "dgMovimentiBusta";
            this.dgMovimentiBusta.Size = new System.Drawing.Size(774, 331);
            this.dgMovimentiBusta.TabIndex = 3;
            this.dgMovimentiBusta.UserDeletedRow += new System.Windows.Forms.DataGridViewRowEventHandler(this.dgMovimentiBusta_UserDeletedRow);
            this.dgMovimentiBusta.UserDeletingRow += new System.Windows.Forms.DataGridViewRowCancelEventHandler(this.dgMovimentiBusta_UserDeletingRow);
            // 
            // fmAnalizzaBusta
            // 
            this.AcceptButton = this.bOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(798, 406);
            this.Controls.Add(this.dgMovimentiBusta);
            this.Controls.Add(this.tbCodiceBusta);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.bOk);
            this.Name = "fmAnalizzaBusta";
            this.Text = "Analizza busta";
            this.Load += new System.EventHandler(this.fmAnalizzaBusta_Load);
            this.Resize += new System.EventHandler(this.fmAnalizzaBusta_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dgMovimentiBusta)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bOk;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbCodiceBusta;
        private System.Windows.Forms.DataGridView dgMovimentiBusta;
    }
}