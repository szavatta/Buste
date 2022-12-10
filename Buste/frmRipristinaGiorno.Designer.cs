namespace Buste
{
    partial class frmRipristinaGiorno
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
            this.dtDataRipristino = new System.Windows.Forms.DateTimePicker();
            this.btnRipristinaGiorno = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // dtDataRipristino
            // 
            this.dtDataRipristino.Location = new System.Drawing.Point(37, 50);
            this.dtDataRipristino.Name = "dtDataRipristino";
            this.dtDataRipristino.Size = new System.Drawing.Size(200, 20);
            this.dtDataRipristino.TabIndex = 16;
            // 
            // btnRipristinaGiorno
            // 
            this.btnRipristinaGiorno.Location = new System.Drawing.Point(37, 118);
            this.btnRipristinaGiorno.Name = "btnRipristinaGiorno";
            this.btnRipristinaGiorno.Size = new System.Drawing.Size(75, 23);
            this.btnRipristinaGiorno.TabIndex = 17;
            this.btnRipristinaGiorno.Text = "Ripristina";
            this.btnRipristinaGiorno.UseVisualStyleBackColor = true;
            this.btnRipristinaGiorno.Click += new System.EventHandler(this.btnRipristinaGiorno_Click);
            // 
            // frmRipristinaGiorno
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 194);
            this.Controls.Add(this.btnRipristinaGiorno);
            this.Controls.Add(this.dtDataRipristino);
            this.Name = "frmRipristinaGiorno";
            this.Text = "Ripristina giorno";
            this.Load += new System.EventHandler(this.frmRipristinaGiorno_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dtDataRipristino;
        private System.Windows.Forms.Button btnRipristinaGiorno;
    }
}