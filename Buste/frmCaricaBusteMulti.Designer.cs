namespace Buste
{
    partial class frmCaricaBusteMulti
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
            this.tbBuste = new System.Windows.Forms.TextBox();
            this.btnCarica = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tbBuste
            // 
            this.tbBuste.Location = new System.Drawing.Point(13, 51);
            this.tbBuste.Multiline = true;
            this.tbBuste.Name = "tbBuste";
            this.tbBuste.Size = new System.Drawing.Size(316, 200);
            this.tbBuste.TabIndex = 0;
            // 
            // btnCarica
            // 
            this.btnCarica.Location = new System.Drawing.Point(13, 257);
            this.btnCarica.Name = "btnCarica";
            this.btnCarica.Size = new System.Drawing.Size(75, 23);
            this.btnCarica.TabIndex = 1;
            this.btnCarica.Text = "Carica";
            this.btnCarica.UseVisualStyleBackColor = true;
            this.btnCarica.Click += new System.EventHandler(this.btnCarica_Click);
            // 
            // frmCaricaBusteMulti
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(341, 292);
            this.Controls.Add(this.btnCarica);
            this.Controls.Add(this.tbBuste);
            this.Name = "frmCaricaBusteMulti";
            this.Text = "Carica buste multi";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbBuste;
        private System.Windows.Forms.Button btnCarica;
    }
}