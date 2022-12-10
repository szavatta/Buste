using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Diagnostics;
using System.IO;

namespace Buste
{
    public partial class frmRipristinaGiorno : Form
    {
        protected Access boAccess;

        public frmRipristinaGiorno()
        {
            InitializeComponent();
        }

        private void frmRipristinaGiorno_Load(object sender, EventArgs e)
        {
            boAccess = new Access();
        }

        private void btnRipristinaGiorno_Click(object sender, EventArgs e)
        {
            string DirPhotosi = boAccess.leggiAppSettings("DirLavorazioni");
            string DirPhotosiMascherine = DirPhotosi.Contains("Photosi") ? DirPhotosi + "Mascherine" : "";
            //DirPhotosi = "c:\\temp\\photosi";

            string path = DirPhotosi + "\\" + dtDataRipristino.Value.ToString("yyyy-MM-dd");
            if (!Directory.Exists(path))
            {
                MessageBox.Show("Cartella inesistente", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string pathmascherine = DirPhotosiMascherine + "\\" + dtDataRipristino.Value.ToString("yyyy-MM-dd");

            boAccess.EliminaMovimentiIn(dtDataRipristino.Value.Date);

            Process myProcess = new Process();
            myProcess.StartInfo.FileName = "cmd.exe";
            myProcess.StartInfo.Arguments = "/c del " + path + "\\job.txt /s";
            myProcess.StartInfo.UseShellExecute = false;
            myProcess.StartInfo.ErrorDialog = false;
            myProcess.StartInfo.RedirectStandardOutput = true;
            myProcess.StartInfo.CreateNoWindow = true;
            myProcess.StartInfo.RedirectStandardError = true;
            myProcess.Start();
            string processError = myProcess.StandardError.ReadToEnd();
            myProcess.WaitForExit();
            myProcess.Close();

            myProcess = new Process();
            myProcess.StartInfo.FileName = "cmd.exe";
            myProcess.StartInfo.Arguments = "/c del " + pathmascherine + "\\job.txt /s";
            myProcess.StartInfo.UseShellExecute = false;
            myProcess.StartInfo.ErrorDialog = false;
            myProcess.StartInfo.RedirectStandardOutput = true;
            myProcess.StartInfo.CreateNoWindow = true;
            myProcess.StartInfo.RedirectStandardError = true;
            myProcess.Start();
            processError = myProcess.StandardError.ReadToEnd();
            myProcess.WaitForExit();
            myProcess.Close();

            MessageBox.Show("Operazione terminata. Ora eseguire una lettura delle buste in entrata.", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
