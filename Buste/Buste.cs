using System;
using System.Collections.Generic;
using System.Configuration;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using ExcelLibrary;
using ExcelLibrary.SpreadSheet;

namespace Buste
{
    public partial class fmBuste : Form
    {

        private bool modificatoOut = false;
        private bool modificatoIn = false;
        private bool modificatoResi = false;
        private int numMovimentoOut = 0;
        private int numMovimentoIn = 0;
        private int numMovimentoReso = 0;
        private string ultimoCodiceLetto = string.Empty;
        private string ultimoArticoloLetto = string.Empty;
        private string ultimaQtaLetta = string.Empty;
        private string ultimaBustaLetta = string.Empty;
        DateTime ultimoCodiceLettoTime = DateTime.Now;
        private int totQta = 0;
        private articolo articolo;
        protected Access boAccess;
        protected Common boCommon;
        protected ListBox ConsuntiviImport = new ListBox();
        private bool ancheIngresso = false;
        private DateTime inizioLettura = DateTime.Now;
        private string messaggioDurata = string.Empty;

        private string suonoErrore = "";
        private string suonoConferma = "";
        private string suonoOKInserimento = "";
        private string suonoOKInserimentoMulti = "";
        private string suonoOKInserimentoMonoprodotto = "";

        private void frmBuste_Load(object sender, EventArgs e)
        {
            boAccess = new Access();
            boCommon = new Common();

            suonoErrore = ConfigurationSettings.AppSettings["SuonoErrore"].ToString();
            if(string.IsNullOrEmpty(suonoErrore))
                suonoErrore = boAccess.leggiAppSettings("SuonoErrore");
            suonoConferma = ConfigurationSettings.AppSettings["SuonoConferma"].ToString();
            if (string.IsNullOrEmpty(suonoConferma))
                suonoConferma = boAccess.leggiAppSettings("SuonoConferma");
            suonoOKInserimento = ConfigurationSettings.AppSettings["SuonoOkInserimento"].ToString();
            if (string.IsNullOrEmpty(suonoOKInserimento))
                suonoOKInserimento = boAccess.leggiAppSettings("SuonoOkInserimento");
            suonoOKInserimentoMulti = ConfigurationSettings.AppSettings["SuonoOkInserimentoMulti"].ToString();
            if (string.IsNullOrEmpty(suonoOKInserimentoMulti))
                suonoOKInserimentoMulti = boAccess.leggiAppSettings("SuonoOkInserimentoMulti");
            suonoOKInserimentoMonoprodotto = ConfigurationSettings.AppSettings["SuonoOKInserimentoMonoprodotto"];
            if (string.IsNullOrEmpty(suonoOKInserimentoMonoprodotto))
                suonoOKInserimentoMonoprodotto = boAccess.leggiAppSettings("SuonoOKInserimentoMonoprodotto");
            //DataTable dt = classifica();

            lblFornitore.Text = boAccess.leggiAppSettings("Fornitore");

            if (boAccess.leggiAppSettings("VerificaConsuntivi") == "1")
            {
                string messaggio = VerificaConsuntivi();
                if (!string.IsNullOrEmpty(messaggio))
                {
                    PlaySound(suonoErrore, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                    DialogResult dr = MessageBox.Show(messaggio, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            tab1.SelectedTab = tabOut;
            PulisceMascheraOut();
            PulisceMascheraIn();
            PulisceMascheraResi();
            tbOutCodice.Select();
            tbOutCodice.Focus();
            
            dgArticoli.DataSource = boAccess.ListaArticoli();

            int secondiIntervalloCaricaBuste = 0; //un ora di intervallo
            try
            {
                secondiIntervalloCaricaBuste = Convert.ToInt32(boAccess.leggiAppSettings("SecondiIntervalloCaricaBuste"));
            }
            catch { }
            if (secondiIntervalloCaricaBuste == 0)
                timerCaricaBuste.Enabled = false;
            else
                timerCaricaBuste.Interval = secondiIntervalloCaricaBuste * 1000;

            try
            {
                if (boAccess.leggiAppSettings("EseguiCaricaBusteAllaPartenza") == "1")
                {
                    int numbuste = CaricaBuste(tab1);
                    if (numbuste > 0)
                    {
                        StampaMessaggio("Sono state inserite " + numbuste.ToString() + " buste in entrata.", false, TipoSuono.NoSuono);
                    }
                    else
                    {
                        StampaMessaggio("Non sono state trovate nuove buste in entrata.", false, TipoSuono.NoSuono);
                    }
                }
            }
            catch { }
        }

        public fmBuste()
        {
            InitializeComponent();
        }

        private RegistryKey key1;
        private RegistryKey key2;


        // collection of user-defined sound events
        private PropertyCollection events;

        // PlaySound()
        [DllImport("winmm.dll", SetLastError = true,
                                CallingConvention = CallingConvention.Winapi)]
        static extern bool PlaySound(
            string pszSound,
            IntPtr hMod,
            SoundFlags sf);

        // Flags for playing sounds.  For this example, we are reading 
        // the sound from a filename, so we need only specify 
        // SND_FILENAME | SND_ASYNC
        [Flags]
        public enum SoundFlags : int
        {
            SND_SYNC = 0x0000,  // play synchronously (default) 
            SND_ASYNC = 0x0001,  // play asynchronously 
            SND_NODEFAULT = 0x0002,  // silence (!default) if sound not found 
            SND_MEMORY = 0x0004,  // pszSound points to a memory file
            SND_LOOP = 0x0008,  // loop the sound until next sndPlaySound 
            SND_NOSTOP = 0x0010,  // don't stop any currently playing sound 
            SND_NOWAIT = 0x00002000, // don't wait if the driver is busy 
            SND_ALIAS = 0x00010000, // name is a registry alias 
            SND_ALIAS_ID = 0x00110000, // alias is a predefined ID
            SND_FILENAME = 0x00020000, // name is file name 
            SND_RESOURCE = 0x00040004  // name is resource name or atom 
        }

        private void PulisceMascheraOut()
        {
            btAzzeraPagina_Click(btOutInserisce, null);
            inizializzaDgOutMovimenti();
            inizializzaDgOutTotali();
            tbOutCodice.Focus();
            tbNumBolla.Text = Convert.ToString(Convert.ToInt32(boAccess.leggiAppSettings("UltimoNumeroBolla")) + 1);
        }

        private void PulisceMascheraIn()
        {
            btAzzeraPagina_Click(btAzzeraPaginaIn, null);
            inizializzaDgInBuste();
            inizializzaDgInTotali();
        }

        private void PulisceMascheraResi()
        {
            btAzzeraPagina_Click(btAzzeraPaginaResi, null);
            inizializzaDgResiBuste();
        }

        private void btOutInserisce_Click(object sender, EventArgs e)
        {
            string fornitore = boAccess.leggiAppSettings("Fornitore");
            lblErrore.Text = string.Empty;
            string codArticolo = string.Empty;
            int qtaArticolo = 0;
            decimal scontoQtaArticolo = 0;
            int secondiDoppiaLettura = 0;
            string valore = boAccess.leggiAppSettings("SecondiDoppiaLettura");
            if (valore != string.Empty)
                secondiDoppiaLettura = Convert.ToInt32(valore);
            else
                secondiDoppiaLettura = 0;

            string[] CodiciFittiziPromozione = null;
            valore = boAccess.leggiAppSettings("CodiciFittiziPromozione");
            if (valore != string.Empty)
                CodiciFittiziPromozione = valore.Split(',');
            
            //sistema la data
            if (DateTime.Now.ToShortDateString() != dtData.Value.ToShortDateString()
                && dgOutMovimenti.Rows.Count == 0)
                dtData.Value = DateTime.Now;
            
            if (lblArticoloOut.Text != string.Empty && lblBustaOut.Text != string.Empty && !lblKitOut.Visible)
            {
                lblArticoloOut.Text = string.Empty;
                lblDesArticoloOut.Text = string.Empty;
                lblQtaOut.Text = string.Empty;
                lblBustaOut.Text = string.Empty;
                lblBatchOut.Text = string.Empty;
            }

            lblMonoprodotto.Visible = false;
            bool isMascherine = false;
            string codice = tbOutCodice.Text;
            if (codice == string.Empty)
                return;
            if (codice.Substring(0, 1) == "C")
                codice=codice.Substring(1);
            if(codice.Substring(codice.Length-1,1) == "C")
                codice=codice.Substring(0,codice.Length-1);

            TimeSpan duration = (DateTime.Now - ultimoCodiceLettoTime);
            if (codice == ultimoCodiceLetto && duration.Seconds < secondiDoppiaLettura)
            {
                StampaMessaggio("Errore. Doppia lettura codice.", true, TipoSuono.NoSuono);
                return;
            }
            ultimoCodiceLetto = tbOutCodice.Text;
            ultimoCodiceLettoTime = DateTime.Now;

            if (codice == string.Empty)
            {
                tbOutCodice.Focus();
                return;
            }
            if (codice.ToLower() == "azzerapagina")
            {
                btAzzeraPagina_Click(btOutInserisce, null);
                return;
            }
            else if (codice.ToLower() == "eliminaultimo")
            {
                btEliminaUltimo_Click(null, null);
                return;
            }
            else if (codice.ToLower() == "iniziokit")
            {
                if (lblBustaOut.Text != string.Empty || lblArticoloOut.Text != string.Empty) //busta piena -> errore
                {
                    StampaMessaggio("Errore. Il codice Kit va letto con pagina vuota.", true, TipoSuono.Errore);
                    return;
                }
                StampaMessaggio("", true, TipoSuono.Kit);
                lblKitOut.Visible = true;
                tbOutCodice.Text = string.Empty;
                tbOutCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "elultimomov")
            {
                PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                DialogResult dr = MessageBox.Show("Si desidera eliminare l'ultimo movimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                if (dr == DialogResult.Yes)
                {
                    eliminaUltimoMovimento();
                }
                else
                {
                    tbOutCodice.Text = string.Empty;
                    tbOutCodice.Focus();
                    return;
                }
            }
            else if (codice.ToLower() == "finekit")
            {
                if (!lblKitOut.Visible)
                {
                    StampaMessaggio("Codice finekit senza iniziokit.", true, TipoSuono.Errore);
                    return;
                }
                StampaMessaggio("", true, TipoSuono.Kit);
                lblKitOut.Visible = false;
                tbOutCodice.Text = string.Empty;
                tbOutCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "salvatemp")
            {
                btSalvaTemp_Click(null, null);
                tbOutCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "caricatemp")
            {
                btCaricaTemp_Click(null, null);
                tbOutCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "out")
            {
                tbOutCodice.Text = string.Empty;
                tbOutCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "in")
            {
                tbOutCodice.Text = string.Empty;
                tbOutCodice.Focus();
                tab1.SelectedTab = tabIn;
                return;
            }
            else if (codice.ToLower() == "aggscura" &&
                ((lblArticoloOut.Text == string.Empty && lblBustaOut.Text == string.Empty) ||
                lblKitOut.Visible))
            {
                codArticolo = boAccess.leggiAppSettings("CodStampaScura");
                articolo articolo = boAccess.LeggiArticolo(codArticolo);
                lblArticoloOut.Text = codArticolo;
                lblDesArticoloOut.Text = articolo.descrizione;
                lblQtaOut.Text = ultimaQtaLetta;
                lblArticoloIntOut.Text = articolo.codiceIntNum;
                lblBustaOut.Text = ultimaBustaLetta;
            }
            //else if (boCommon.ContainsString(CodiciFittiziPromozione, codice.ToLower()))
            //{
            //}
            else if (codice.Substring(codice.Length - 1, 1) == "*")
            {
                try
                {
                    qtaArticolo = Convert.ToInt32(codice.Substring(0, codice.Length - 1));
                }
                catch
                {
                    StampaMessaggio("Errore. Quantità non valida.", true, TipoSuono.Errore);
                    return;
                }
                return;
            }
            else if (codice.StartsWith("F") && codice.EndsWith("F"))
            {
                codice = codice.Replace("F", "");
                DataTable dbBatch = boAccess.ListaMovimentiBatchIn(codice);
                if(dbBatch.Rows.Count == 0)
                {
                    StampaMessaggio("Errore. Codice busta non trovato.", true, TipoSuono.Errore);
                    return;
                }
                DataTable dbBatchOut = boAccess.ListaMovimentiBatchOut(codice);
                if (dbBatchOut.Rows.Count > 0 || isBatchEsistenteGridOut(codice))
                {
                    PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                    DialogResult dr = MessageBox.Show("Codice busta già letta. Confermi ugualmente l'inserimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                    if (dr == DialogResult.No)
                    {
                        tbOutCodice.Text = string.Empty;
                        return;
                    }
                }
                int cont = 0, qtam = 0;
                string codArt = "", idBusta = "", idBatch = "";
                foreach (DataRow drc in dbBatch.Rows)
                {
                    cont++;
                    codArt = drc["idArticolo"].ToString();
                    idBusta = drc["idBusta"].ToString();
                    idBatch = drc["idBatch"].ToString();
                    string qta = drc["qta"].ToString();
                    string data = drc["data"].ToString();
                    qtam += Convert.ToInt32(qta);

                    articolo articolo = boAccess.LeggiArticolo(codArt);
                    if (articolo.codice == string.Empty)
                    {
                        string oldcodart = codArt;
                        string artms = boAccess.leggiAppSettings("ArticoliMascherineS");
                        if (artms.Contains(";" + codArt + ";"))
                            codArt = boAccess.leggiAppSettings("ArticoliMascherineS").Split(';')[1];
                        artms = boAccess.leggiAppSettings("ArticoliMascherineM");
                        if (artms.Contains(";" + codArt + ";"))
                            codArt = boAccess.leggiAppSettings("ArticoliMascherineM").Split(';')[1];
                        artms = boAccess.leggiAppSettings("ArticoliMascherineL");
                        if (artms.Contains(";" + codArt + ";"))
                            codArt = boAccess.leggiAppSettings("ArticoliMascherineL").Split(';')[1];

                        if (codArt == oldcodart)
                        {
                            codArt = "";
                            PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                            DialogResult dr = MessageBox.Show("Errore. Codice articolo " + codArt + " inesistente.");
                            tbOutCodice.Text = string.Empty;
                        }
                        else
                        {
                            articolo = boAccess.LeggiArticolo(codArt);
                        }
                    }

                    lblArticoloOut.Text = codArt;
                    lblDesArticoloOut.Text = articolo.descrizione;
                    lblArticoloIntOut.Text = articolo.codiceIntNum;
                    lblBustaOut.Text = idBusta;
                    lblBatchOut.Text = idBatch;
                    //if (cont < dbBatch.Rows.Count)
                    //{
                    //    AggiungeDGOutMovimento(codArt, lblDesArticoloOut.Text, idBusta, dtData.Value, Convert.ToInt32(tbNumBolla.Text), Convert.ToInt32(qta), Convert.ToDecimal("0" + lblScontoOut.Text));
                    //}
                }
                //int qtaf = qtam / 36 + (qtam % 36 != 0 ? 1 : 0);
                lblQtaOut.Text = qtam.ToString();
                //AggiungeDGOutMovimento(codArt, lblDesArticoloOut.Text, idBusta, dtData.Value, Convert.ToInt32(tbNumBolla.Text), Convert.ToInt32(qtaf), Convert.ToDecimal("0" + lblScontoOut.Text));
            }
            else if (fornitore.ToLower() == "photosi" && (codice.Length == 11 || codice.Length == 4 || codice.IndexOf("*") > 0) ||
                    fornitore.ToLower() == "fotoevolution" && (codice.Length == 12 && codice.Substring(0, 5) != "00000") || (codice.Length == 6 && lblBustaOut.Text != string.Empty) || codice.IndexOf("*") > 0) //codice articolo
            {
                if (lblArticoloOut.Text != string.Empty && !lblKitOut.Visible) //articolo pieno -> errore
                {
                    StampaMessaggio("Errore. Lettura doppia codice articolo.", true, TipoSuono.Errore);
                    return;
                }
                if (lblKitOut.Visible && lblBustaOut.Text == string.Empty)
                {
                    StampaMessaggio("Errore. Con il Kit leggere per primo il codice busta.", true, TipoSuono.Errore);
                    return;
                }

                int posX = codice.IndexOf("*");
                if (posX >= 4)
                {
                    codArticolo = codice.Substring(0, posX);
                    try
                    {
                        qtaArticolo = Convert.ToInt32(codice.Substring(posX + 1));
                    }
                    catch
                    {
                        qtaArticolo = 1;
                    }
                }
                else if (posX < 4 && posX > 0)
                {
                    codArticolo = codice.Substring(posX + 1, 4);
                    try
                    {
                        qtaArticolo = Convert.ToInt32(codice.Substring(0, posX));
                    }
                    catch
                    {
                        qtaArticolo = 1;
                    }
                }
                else if (fornitore.ToLower() == "photosi" && codice.Length == 11)
                {
                    codArticolo = codice.Substring(5, 4);
                    try
                    {
                        qtaArticolo = Convert.ToInt32(codice.Substring(0, 5));
                    }
                    catch { return; }
                }
                else if (fornitore.ToLower() == "photosi" && codice.Length == 4)
                {
                    codArticolo = codice;
                    qtaArticolo = 1;
                }
                else if (fornitore.ToLower() == "fotoevolution" && codice.Length == 12)
                {
                    codArticolo = codice.Substring(0, 6);
                    try
                    {
                        qtaArticolo = Convert.ToInt32(codice.Substring(6, 5));
                    }
                    catch { return; }
                }
                else if (fornitore.ToLower() == "fotoevolution" && codice.Length == 6)
                {
                    codArticolo = codice;
                    qtaArticolo = 1;
                }

                if (qtaArticolo < 1)
                {
                    StampaMessaggio("Errore. Quantita a zero.", true, TipoSuono.Errore);
                    return;
                }
                
                articolo articolo = boAccess.LeggiArticolo(codArticolo);
                if (articolo.codice == string.Empty)
                {
                    PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                    DialogResult dr = MessageBox.Show("Errore. Codice articolo " + codArticolo + " inesistente. Confermi inserimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                    if (dr == DialogResult.Yes)
                    {
                        frmNuovoArticolo form = new frmNuovoArticolo(codArticolo, "");
                        //form.Tag
                    }
                    else
                    {
                        tbOutCodice.Text = string.Empty;
                        return;
                    }

                    //StampaMessaggio("Errore. Codice articolo " + codArticolo + " inesistente.", true, TipoSuono.Errore);
                    return;
                }
                if (boAccess.leggiAppSettings("MinQtaMessaggio").ToString() != string.Empty && qtaArticolo > Convert.ToInt32(boAccess.leggiAppSettings("MinQtaMessaggio")))
                {
                    PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                    DialogResult dr = MessageBox.Show("Attenzione, quantità = " + qtaArticolo.ToString() + ". Confermi?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                    if (dr == DialogResult.No)
                    {
                        tbOutCodice.Text = string.Empty;
                        return;
                    }
                }

                if (boAccess.leggiAppSettings("VerificaBusteEntrata") == "1" &&
                    codArticolo != boAccess.leggiAppSettings("BustaFittizia") &&
                    lblBustaOut.Text != string.Empty &&
                    !ancheIngresso)
                {
                    if (!controlloCoerenzaBuste(lblBustaOut.Text, codArticolo, qtaArticolo, "Out"))
                    {
                        tbOutCodice.Text = "";
                        return;
                    }
                }

                lblArticoloOut.Text = codArticolo;
                lblDesArticoloOut.Text = articolo.descrizione;
                lblQtaOut.Text = qtaArticolo.ToString();
                lblArticoloIntOut.Text = articolo.codiceIntNum;
                scontoQtaArticolo = calcolaSconto(articolo, qtaArticolo);
                if (scontoQtaArticolo > 0)
                {
                    lblScontoOut.Text = scontoQtaArticolo.ToString();
                    lblPercentuale.Visible = true;
                }
                else
                {
                    lblScontoOut.Text = string.Empty;
                    lblPercentuale.Visible = false;
                }

                ultimoArticoloLetto = codArticolo;
                ultimaQtaLetta = qtaArticolo.ToString();

                if (lblBustaOut.Text == string.Empty) //articolo vuoto e busta vuota
                {
                    tbOutCodice.Text = string.Empty;
                    tbOutCodice.Focus();
                    return;
                }
            }
            else if (isCodiceBusta(codice)) //codice busta
            {
                if (fornitore.ToLower() == "fotoevolution" && codice.Length > 6)
                    codice = Convert.ToString(Convert.ToInt32(codice.Substring(0, 11)));
                if (boAccess.leggiAppSettings("VerificaBusteEntrata") == "1" &&
                    lblArticoloOut.Text != boAccess.leggiAppSettings("BustaFittizia") &&
                    lblArticoloOut.Text != string.Empty)
                {
                    if (!controlloCoerenzaBuste(codice, lblArticoloOut.Text, Convert.ToInt32(lblQtaOut.Text), "Out"))
                    {
                        tbOutCodice.Text = "";
                        return;
                    }
                }
                if (!inserisceCodiceBustaOut(codice))
                    return;
            }
            else
            {
                //StampaMessaggio("Codice non riconosciuto. E' un codice busta?", false, TipoSuono.Domanda);
                //lCodErrore.Text = "1";
                //tbErrore.Enabled = true;
                //this.AcceptButton = btErrore;
                //tbErrore.Text = string.Empty;
                //tbErrore.Select();
                //tbErrore.Focus();
                //return;
                PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                DialogResult dr = MessageBox.Show("Codice busta non riconosciuto. Confermi ugualmente?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.No)
                {
                    tbOutCodice.Text = string.Empty;
                    return;
                }
                else
                {
                    lblBustaOut.Text = codice;
                    ultimaBustaLetta = codice;
                }
            }

            if (lblQtaOut.Text == string.Empty)
                qtaArticolo = 0;
            else
                qtaArticolo = Convert.ToInt32(lblQtaOut.Text);

            codArticolo = lblArticoloOut.Text;
            string codBusta = lblBustaOut.Text;
            if (qtaArticolo > 0)
            {
                if (ancheIngresso)
                {
                    AggiungeDGInBuste("", codArticolo, lblDesArticoloOut.Text, codBusta, dtData.Value, qtaArticolo);
                    bool ret = boAccess.ScriveMovimentiIn(dgInBuste);
                    AzzeraInBuste();
                }

                busta? bustain = boAccess.leggiBustaDB(codBusta, "In");
                string tipoOrdine = bustain != null ? bustain.Value.tipoOrdine : "";
                string numProdotti = bustain != null ? bustain.Value.numProdotti : "";

                lblMonoprodotto.Visible = tipoOrdine == "MAILORDER" && numProdotti == "MONOPRODOTTO";

                AggiungeDGOutMovimento(codArticolo, 
                    lblDesArticoloOut.Text, 
                    codBusta, 
                    dtData.Value, 
                    Convert.ToInt32(tbNumBolla.Text), 
                    qtaArticolo, 
                    Convert.ToDecimal("0" + lblScontoOut.Text), 
                    lblBatchOut.Text,
                    tipoOrdine,
                    numProdotti == "MONOPRODOTTO" ? numProdotti : ""
                    );
                //AggiungeCodiceArticoloInTotali(codArticolo, lblDesArticoloOut.Text, qtaArticolo, lblArticoloIntOut.Text, "Out");
                //AggiungeCodiceArticoloInTotaliGrup(codArticolo, lblDesArticoloOut.Text, qtaArticolo, lblArticoloIntOut.Text, "Out");
                TipoSuono tiposuono = qtaArticolo > 1 ? TipoSuono.OkInserimentoMulti : TipoSuono.OkInserimento;
                if (lblMonoprodotto.Visible)
                    tiposuono = TipoSuono.OkInserimentoMonoprodotto;
                StampaMessaggio("", true, tiposuono);
                modificatoOut = true;
            }
            tbOutCodice.Text = string.Empty;
            tbOutCodice.Focus();
        }

        private decimal calcolaSconto(articolo art, int quantita)
        {
            decimal sconto = 0;
            if (art.scaglione5 > 0 && quantita >= art.scaglione5)
                sconto = art.sconto5;
            else if (art.scaglione4 > 0 && quantita >= art.scaglione4)
                sconto = art.sconto4;
            else if (art.scaglione3 > 0 && quantita >= art.scaglione3)
                sconto = art.sconto3;
            else if (art.scaglione2 > 0 && quantita >= art.scaglione2)
                sconto = art.sconto2;
            else if (art.scaglione1 > 0 && quantita >= art.scaglione1)
                sconto = art.sconto1;

            return sconto;
        }

        private bool controlloCoerenzaBuste(string codBusta, string codArticolo, int qta, string tipoOp)
        {
            //controllo coerenza con buste in ingresso
            bool ret = true;
            //if (boAccess.coerenzaBusta(codBusta, codArticolo, qta))
            //    return true;

            //busta BustaIn = boAccess.leggiBustaDB(codBusta, "In");
            //if (BustaIn.idBusta != "0")
            //{
            //    if (BustaIn.codArticolo != codArticolo)
            //    {
            //        PlaySound(boAccess.leggiAppSettings("SuonoConferma"), IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
            //        DialogResult dr = MessageBox.Show("Attenzione, codice articolo diverso dalla busta entrata. " + BustaIn.codArticolo + " -> " + codArticolo, "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
            //        if (dr == DialogResult.No)
            //            ret = false;
            //    }
                //if (BustaIn.quantita != qta)
                //{
                //    PlaySound(boAccess.leggiAppSettings("SuonoConferma"), IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                //    DialogResult dr = MessageBox.Show("Attenzione, quantità diversa dalla busta entrata. " + BustaIn.quantita.ToString() + " -> " + qta.ToString(), "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                //    if (dr == DialogResult.No)
                //        ret = false;
                //}
                int qtaRif = qta;
                if (tipoOp == "Out" || tipoOp == "Resi")
                    qtaRif = qta * -1;
                int DiffQta = boAccess.QtaDiffArticolo(codBusta, codArticolo) + qtaRif;
                if (DiffQta != 0)
                {
                    //Non corrisponde la quantità
                    PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                    DialogResult dr = MessageBox.Show("Quantità incoerente (" + DiffQta + "). Confermi ugualmente?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                    if (dr == DialogResult.No)
                    {
                        tbOutCodice.Text = string.Empty;
                        ret = false;
                    }
                }


            //}
            
            return ret;

        }

        private bool inserisceCodiceBustaOut(string codice)
        {
            if (lblBustaOut.Text != string.Empty) //busta piena -> errore
            {
                StampaMessaggio("Errore. Lettura doppia codice busta.", true, TipoSuono.Errore);
                return false;
            }
            ancheIngresso = false;
            if (!boAccess.isBustaEsistenteDBIn(codice) 
                && !isBustaEsistenteGridOut(codice)
                && boAccess.leggiAppSettings("VerificaBusteEntrata") == "1")
            {
                PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                DialogResult dr = MessageBox.Show("Codice busta non entrata. Vuoi inserire anche l'entrata?", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3);
                if (dr == DialogResult.Cancel)
                {
                    tbOutCodice.Text = string.Empty;
                    return false;
                }
                else if (dr == DialogResult.Yes)
                {
                    ancheIngresso = true;
                }
            }
            
            if (isBustaEsistenteGridOut(codice)
                && boAccess.leggiAppSettings("VerificaBusteUscita") == "1"
                && codice != boAccess.leggiAppSettings("BustaFittizia")
                && boAccess.leggiAppSettings("BustaFittizia") != string.Empty
                && codice.Substring(0,1) != "8")
            {
                PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                DialogResult dr = MessageBox.Show("Codice busta già letta. Confermi ugualmente l'inserimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.No)
                {
                    tbOutCodice.Text = string.Empty;
                    return false;
                }
            }

            lblBustaOut.Text = codice;
            ultimaBustaLetta = codice;
            return true;
            
            if (lblArticoloOut.Text == string.Empty) //busta vuota e articolo pieno
            {
                tbOutCodice.Text = string.Empty;
                tbOutCodice.Focus();
            }
        }

        //private bool inserisceCodiceBustaIn_Old(string codice)
        //{
        //    bool ok = true;
        //    if (boAccess.isBustaEsistenteDBIn(codice)
        //        || isBustaEsistenteGridIn(codice))
        //    {
        //        PlaySound(boAccess.leggiAppSettings("SuonoConferma"), IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
        //        DialogResult dr = MessageBox.Show("Codice busta esistente. Confermi ugualmente l'inserimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
        //        if (dr == DialogResult.No)
        //        {
        //            ok = false;
        //        }
        //    }

        //    lblBustaIn.Text = codice;
        //    ultimaBustaLetta = codice;
        //    return true;

        //    return ok;

        //}

        private bool inserisceCodiceBustaIn(string codice, bool isRifacimento)
        {
            if (lblBustaIn.Text != string.Empty) //busta piena -> errore
            {
                StampaMessaggio("Errore. Lettura doppia codice busta.", true, TipoSuono.Errore);
                return false;
            }
            if (boAccess.leggiAppSettings("VerificaBusteEntrata") == "1"
                && codice != boAccess.leggiAppSettings("BustaFittizia")
                && boAccess.leggiAppSettings("BustaFittizia") != string.Empty
                && codice.Substring(0, 1) != "8")
            {
                if (boAccess.isBustaEsistenteDBIn(codice) && !isRifacimento)
                {
                    busta Busta = boAccess.leggiBustaDB(codice, "In");
                    if (Busta.idBusta == codice)
                    {
                        PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                        DialogResult dr = MessageBox.Show("Codice busta già entrata. Confermi ugualmente l'inserimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                        if (dr == DialogResult.No)
                        {
                            tbInCodice.Text = string.Empty;
                            return false;
                        }
                    }
                }

                if (!boAccess.isBustaEsistenteDBIn(codice) && isRifacimento)
                {
                    PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                    DialogResult dr = MessageBox.Show("Codice busta rifacimento mai entrata. Confermi ugualmente l'inserimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                    if (dr == DialogResult.No)
                    {
                        tbInCodice.Text = string.Empty;
                        return false;
                    }
                }

                if (isBustaEsistenteGridIn(codice))
                {
                    PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                    DialogResult dr = MessageBox.Show("Codice busta presente in lettura. Confermi ugualmente l'inserimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                    if (dr == DialogResult.No)
                    {
                        tbInCodice.Text = string.Empty;
                        return false;
                    }
                }

            }

            lblBustaIn.Text = codice;
            ultimaBustaLetta = codice;
            return true;

            if (lblArticoloIn.Text == string.Empty) //busta vuota e articolo pieno
            {
                tbInCodice.Text = string.Empty;
                tbInCodice.Focus();
            }
        }

        private bool inserisceCodiceBustaResi(string codice)
        {
            if (lblBustaResi.Text != string.Empty) //busta piena -> errore
            {
                StampaMessaggio("Errore. Lettura doppia codice busta.", true, TipoSuono.Errore);
                return false;
            }
            if (!boAccess.isBustaEsistenteDBIn(codice)
                && !isBustaEsistenteGridResi(codice)
                && boAccess.leggiAppSettings("VerificaBusteEntrata") == "1")
            {
                PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                DialogResult dr = MessageBox.Show("Codice busta non entrata. Confermi ugualmente l'inserimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.No)
                {
                    tbCodiceResi.Text = string.Empty;
                    return false;
                }
            }
            if (boAccess.isBustaEsistenteDBResi(codice)
                && boAccess.leggiAppSettings("VerificaBusteEntrata") == "1"
                && codice != boAccess.leggiAppSettings("BustaFittizia")
                && boAccess.leggiAppSettings("BustaFittizia") != string.Empty)
            {
                busta Busta = boAccess.leggiBustaDB(codice, "Resi");
                if (Busta.data != Convert.ToDateTime("01/01/1900"))
                {
                    PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                    DialogResult dr = MessageBox.Show("Codice busta già resa. Confermi ugualmente l'inserimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                    if (dr == DialogResult.No)
                    {
                        tbCodiceResi.Text = string.Empty;
                        return false;
                    }
                }
            }
            if (boAccess.isBustaEsistenteDBOut(codice)
                && boAccess.leggiAppSettings("VerificaBusteEntrata") == "1"
                && codice != boAccess.leggiAppSettings("BustaFittizia")
                && boAccess.leggiAppSettings("BustaFittizia") != string.Empty)
            {
                busta Busta = boAccess.leggiBustaDB(codice, "Resi");
                if (Busta.data != Convert.ToDateTime("01/01/1900"))
                {
                    PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                    DialogResult dr = MessageBox.Show("Codice busta già uscita. Confermi ugualmente l'inserimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                    if (dr == DialogResult.No)
                    {
                        tbCodiceResi.Text = string.Empty;
                        return false;
                    }
                }
            }
            if (isBustaEsistenteGridResi(codice)
                && boAccess.leggiAppSettings("VerificaBusteEntrata") == "1"
                && codice != boAccess.leggiAppSettings("BustaFittizia")
                && boAccess.leggiAppSettings("BustaFittizia") != string.Empty)
            {
                PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                DialogResult dr = MessageBox.Show("Codice busta già letta. Confermi ugualmente l'inserimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.No)
                {
                    tbCodiceResi.Text = string.Empty;
                    return false;
                }
            }

            lblBustaResi.Text = codice;
            ultimaBustaLetta = codice;
            return true;

            if (lblArticoloResi.Text == string.Empty) //busta vuota e articolo pieno
            {
                tbCodiceResi.Text = string.Empty;
                tbCodiceResi.Focus();
            }
        }

        private void tbFine_Click(object sender, EventArgs e)
        {
            if (dgOutMovimenti.Rows.Count == 0)
            {
                StampaMessaggio("Non ci sono buste da salvare.", true, TipoSuono.Errore);
                return;
            }
            System.Media.SystemSounds.Question.Play();
            DialogResult dr = MessageBox.Show("Confermi chiusura bolla?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.No)
                return;

            if ((lblArticoloOut.Text == string.Empty && lblBustaOut.Text != string.Empty) || 
                (lblArticoloOut.Text != string.Empty && lblBustaOut.Text == string.Empty) || 
                lblKitOut.Visible)
            {
                StampaMessaggio("Errore. Busta aperta.", true, TipoSuono.Errore);
                return;
            }

            RicalcolaTotali();
            bool ret = boAccess.ScriveMovimentiOut(dgOutMovimenti);
            string nomeFile = string.Empty;
            if (ret)
            {
                //DialogResult dr = MessageBox.Show("Scrittura eseguita con successo. Vuoi generare i file?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                //if (dr == DialogResult.Yes)
                //{
                string aa = boAccess.leggiAppSettings("DirFilePhotosi");
                if (!GeneraFilePhotosi(boAccess.leggiAppSettings("DirFilePhotosi")))
                    MessageBox.Show("Errore in fase di scrittura file Photosi", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (!GeneraFileEsatto(boAccess.leggiAppSettings("DirFileEsatto"), boAccess.leggiAppSettings("Fornitore")))
                    MessageBox.Show("Errore in fase di scrittura file Esatto", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (!GeneraFileExcel(boAccess.leggiAppSettings("DirFileExcel"), boAccess.leggiAppSettings("Fornitore")))
                    MessageBox.Show("Errore in fase di scrittura file Esatto", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (!GeneraFileTotali(boAccess.leggiAppSettings("DirFileBolla")))
                    MessageBox.Show("Errore in fase di scrittura file totali", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (!GeneraFileTotaliGrup(boAccess.leggiAppSettings("DirFileBolla"), out nomeFile))
                    MessageBox.Show("Errore in fase di scrittura file totaligrup", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (!GeneraFileMovimenti(boAccess.leggiAppSettings("DirFileBolla")))
                    MessageBox.Show("Errore in fase di scrittura file movimenti", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //}
                boAccess.scriviAppSettings("UltimoNumeroBolla", tbNumBolla.Text);
                AzzeraTutto();
                modificatoOut = false;
            }
            else
            {
                MessageBox.Show("Errore in fase di scrittura dei movimenti", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            string msg = "Scrittura eseguita con successo.";
            if (boAccess.leggiAppSettings("DirFilePhotosi") == "1")
            {
                msg += " Sarà aperto il file con i totali.";
                MessageBox.Show(msg, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Process.Start("Notepad.exe", boAccess.leggiAppSettings("DirFileBolla") + "\\" + nomeFile);
            }
            else
            {
                MessageBox.Show(msg, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private bool isBustaEsistenteGridOut(string idBusta)
        {
            bool ret = false;

            foreach (DataGridViewRow dr in dgOutMovimenti.Rows)
            {
                if (dr.Cells["idBusta"].Value.ToString() == idBusta)
                {
                    ret = true;
                    break;
                }
            }

            return ret;
        }

        private bool isBatchEsistenteGridOut(string idBatch)
        {
            bool ret = false;

            foreach (DataGridViewRow dr in dgOutMovimenti.Rows)
            {
                if (dr.Cells["batch"].Value.ToString() == idBatch)
                {
                    ret = true;
                    break;
                }
            }

            return ret;
        }

        private bool isBustaEsistenteGridIn(string idBusta)
        {
            bool ret = false;

            foreach (DataGridViewRow dr in dgInBuste.Rows)
            {
                if (dr.Cells["idBusta"].Value.ToString() == idBusta)
                {
                    ret = true;
                    break;
                }
            }

            return ret;
        }

        private bool isBustaEsistenteGridResi(string idBusta)
        {
            bool ret = false;

            foreach (DataGridViewRow dr in dgResiBuste.Rows)
            {
                if (dr.Cells["idBusta"].Value.ToString() == idBusta)
                {
                    ret = true;
                    break;
                }
            }

            return ret;
        }

        private void btAzzeraPagina_Click(object sender, EventArgs e)
        {
            if (sender == btAzzeraPaginaOut || sender == btOutInserisce)
            {
                lblArticoloOut.Text = string.Empty;
                lblDesArticoloOut.Text = string.Empty;
                lblBustaOut.Text = string.Empty;
                lblQtaOut.Text = string.Empty;
                lblErrore.Text = string.Empty;
                tbOutCodice.Text = string.Empty;
                lblKitOut.Visible = false;
                lblMonoprodotto.Visible = false;
                lblScontoOut.Text = string.Empty;
                lblPercentuale.Visible = false;
                tbOutCodice.Focus();
            }
            else if (sender == btAzzeraPaginaIn || sender == btInInserisce)
            {
                lblArticoloIn.Text = string.Empty;
                lblDesArticoloIn.Text = string.Empty;
                lblBustaIn.Text = string.Empty;
                lblQtaIn.Text = string.Empty;
                tbInCodice.Text = string.Empty;
                lblKitIn.Visible = false;
                ckRifacimento.Checked = false;
                tbInCodice.Focus();
            }
            else if (sender == btAzzeraPaginaResi || sender == btInserisceResi)
            {
                lblArticoloResi.Text = string.Empty;
                lblDesArticoloResi.Text = string.Empty;
                lblBustaResi.Text = string.Empty;
                lblQtaResi.Text = string.Empty;
                tbCodiceResi.Text = string.Empty;
                lblKitResi.Visible = false;
                tbCodiceResi.Focus();
            }
        }

        private void DgOutMovimenti_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            //RicalcolaTotali();
        }

        private void RicalcolaTotali()
        {
            AzzeraTotali();
            foreach (DataGridViewRow dr in dgOutMovimenti.Rows)
            {
                if (dr.Index < dgOutMovimenti.Rows.Count)
                {
                    articolo starticolo = boAccess.LeggiArticolo(dr.Cells["idArticolo"].Value.ToString());
                    AggiungeCodiceArticoloInTotali(dr.Cells["idArticolo"].Value.ToString(), starticolo.descrizione, Convert.ToInt32(dr.Cells["qta"].Value), Convert.ToDecimal(dr.Cells["sconto"].Value), starticolo.codiceInt, "Out");
                    AggiungeCodiceArticoloInTotaliGrup(dr.Cells["idArticolo"].Value.ToString(), starticolo.descrizione, Convert.ToInt32(dr.Cells["qta"].Value), starticolo.codiceIntNum, "Out");
                }
            }
        }

        private void AzzeraTotali()
        {
            dgOutTotali.Rows.Clear();
            dgOutTotaliGrup.Rows.Clear();
            totQta = 0;
        }

        private void AzzeraMovimenti()
        {
            dgOutMovimenti.Rows.Clear();
            numMovimentoOut = 0;
        }

        private void AzzeraInTotali()
        {
            dgInTotali.Rows.Clear();
        }

        private void AzzeraInBuste()
        {
            dgInBuste.Rows.Clear();
            numMovimentoIn = 0;
        }

        private void AzzeraResiBuste()
        {
            dgResiBuste.Rows.Clear();
            numMovimentoReso = 0;
        }

        private void AggiungeCodiceArticoloInTotaliGrup(string codArticolo, string desArticolo, int qtaCodice, string codArticoloInt, string InOut)
        {
            DataGridView datagrid = null;
            if (InOut.ToLower() == "out")
                datagrid = dgOutTotaliGrup;

            foreach (DataGridViewRow dr in datagrid.Rows)
            {
                if (dr.Cells["idArticolo"].Value.ToString() == codArticoloInt)
                {
                    int quantita = Convert.ToInt32(dr.Cells["qta"].Value);
                    quantita += qtaCodice;
                    dr.Cells["qta"].Value = quantita.ToString();
                    return;
                }
            }
            object[] objArtGrup = new object[] { codArticoloInt, desArticolo, qtaCodice.ToString() };
            datagrid.Rows.Add(objArtGrup);
        }

        private void AggiungeCodiceArticoloInTotali(string codArticolo, string desArticolo, int qtaCodice, decimal sconto, string codArticoloInt, string InOut)
        {

            DataGridView datagrid = null;
            if (InOut.ToLower() == "out")
                datagrid = dgOutTotali;
            else if (InOut.ToLower() == "in")
                datagrid = dgInTotali;

            articolo articolo = boAccess.LeggiArticolo(codArticolo);

            int omg = 0;
            if (InOut.ToLower() == "out")
            {
                foreach (DataGridViewRow dr in datagrid.Rows)
                {
                    if (dr.Cells["idArticolo"].Value.ToString() == codArticolo &&
                        dr.Cells["sconto"].Value.ToString() == sconto.ToString())
                    {
                        int quantita = Convert.ToInt32(dr.Cells["qta"].Value);
                        quantita += qtaCodice;
                        dr.Cells["qta"].Value = quantita.ToString();
                        dr.Cells["sconto"].Value = sconto.ToString();
                        int omaggio = Convert.ToInt32(dr.Cells["omaggio"].Value);
                        omaggio += omg;
                        dr.Cells["omaggio"].Value = omaggio.ToString();
                        return;
                    }
                }
            }
            object[] objArt = new object[] { codArticolo, desArticolo, qtaCodice.ToString(), sconto.ToString(), omg.ToString() };
            datagrid.Rows.Add(objArt);
        }

        private void AggiungeDGOutMovimento(string codArticolo, string descrArticolo, string codBusta, DateTime dataMov, long numBolla, int qtaCodice, decimal sconto, string codBatch = "", string tipoOrdine = "", string numProdotti = "")
        {
            numMovimentoOut += 1;
            object[] objMov = new object[] { numMovimentoOut, codArticolo, descrArticolo, codBusta, codBatch, dataMov.ToShortDateString(), numBolla, qtaCodice, sconto, tipoOrdine, numProdotti };
            dgOutMovimenti.Rows.Insert(0, objMov);
            totQta += qtaCodice;
            lblTotQta.Text = totQta.ToString();
        }

        private void AggiungeDGInBuste(string idBatch, string codArticolo, string descrArticolo, string codBusta, DateTime dataBusta, int qtaCodice, string tipoOrdine = "", string numProdotti = "")
        {
            numMovimentoIn += 1;
            object[] objMov = new object[] { numMovimentoIn, codArticolo, descrArticolo, codBusta, idBatch, dataBusta.ToShortDateString(), qtaCodice, tipoOrdine, numProdotti };
            dgInBuste.Rows.Insert(0, objMov);
        }

        private void AggiungeDGResiBuste(string idBatch, string codArticolo, string descrArticolo, string codBusta, DateTime dataBusta, int qtaCodice)
        {
            numMovimentoReso += 1;
            object[] objMov = new object[] { numMovimentoReso, codArticolo, descrArticolo, codBusta, idBatch, dataBusta.ToShortDateString(), qtaCodice };
            dgResiBuste.Rows.Insert(0, objMov);
        }

        private void btAzzeraTutto_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Confermi azzeramento totale?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.Yes)
                AzzeraTutto();
        }

        private void AzzeraTutto()
        {
            string numbolla = string.Empty;
            AzzeraTotali();
            AzzeraMovimenti();
            btAzzeraPagina_Click(btAzzeraPaginaOut, null);
            numbolla = boAccess.leggiAppSettings("UltimoNumeroBolla");
            if(!string.IsNullOrEmpty(numbolla))
                tbNumBolla.Text = Convert.ToString(Convert.ToInt32(numbolla) + 1);
            modificatoOut = false;
            totQta = 0;
            lblTotQta.Text = "0";
        }

        private void frmBuste_Resize(object sender, EventArgs e)
        {
            try
            {
                tab1.Width = fmBuste.ActiveForm.Width - 42;
                tab1.Height = fmBuste.ActiveForm.Height - 83;
                tabOutInt.Width = tab1.Width - 339;
                tabOutInt.Height = tab1.Height - 182;
                tabInInt.Width = tab1.Width - 339;
                tabInInt.Height = tab1.Height - 182;
                tabResiInt.Width = tab1.Width - 339;
                tabResiInt.Height = tab1.Height - 182;
                dgOutMovimenti.Width = tabOutInt.Width - 20;
                dgOutMovimenti.Height = tabOutInt.Height - 38;
                dgOutTotali.Width = tabOutInt.Width - 20;
                dgOutTotali.Height = tabOutInt.Height - 38;
                dgOutTotaliGrup.Width = tabOutInt.Width - 20;
                dgOutTotaliGrup.Height = tabOutInt.Height - 38;
                dgResiBuste.Width = tabOutInt.Width - 20;
                dgResiBuste.Height = tabOutInt.Height - 38;
                gbOut1.Width = tab1.Width - 35;
                gbIn1.Width = tab1.Width - 35;
                gbResi1.Width = tab1.Width - 35;
                dgInBuste.Width = tabInInt.Width - 20;
                dgInBuste.Height = tabInInt.Height - 38;
                tabResiInt.Width = tab1.Width - 339;
                tabResiInt.Height = tab1.Height - 182;
                lblKitOut.Location = new Point(tab1.Width - 88, lblKitOut.Location.Y);
            }
            catch
            {
            }
        }

        private bool GeneraFilePhotosi(string strDir)
        {
            bool ret = true;

            string[] dirPhotosi = strDir.Split(',');

            if (!string.IsNullOrEmpty(dirPhotosi[0]))
            {

                string nomeFile = dtData.Value.Year.ToString().PadLeft(4, '0') +
                    dtData.Value.Month.ToString().PadLeft(2, '0') +
                    dtData.Value.Day.ToString().PadLeft(2, '0') +
                    "-Bolla" + tbNumBolla.Text.PadLeft(4, '0') +
                    ".txt";

                FileStream fs = new FileStream(dirPhotosi[0] + "\\" + nomeFile, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs, Encoding.ASCII);

                try
                {
                    foreach (DataGridViewRow dr in dgOutMovimenti.Rows)
                    {
                        if (dr.Cells["idArticolo"].Value.ToString() != boAccess.leggiAppSettings("CodStampaScura"))
                        {
                            string riga = "  " +
                                dtData.Value.Year.ToString().Substring(2, 2) +
                                dtData.Value.Month.ToString().PadLeft(2, '0') +
                                dtData.Value.Day.ToString().PadLeft(2, '0') +
                                "     " +
                                "000" +
                                boAccess.leggiAppSettings("idFornitore").PadLeft(6, '0') +
                                dr.Cells["idBusta"].Value.ToString().PadLeft(9, '0') +
                                dr.Cells["idArticolo"].Value.ToString().PadLeft(4, '0') +
                                dr.Cells["qta"].Value.ToString().PadLeft(5, '0');

                            sw.WriteLine(riga);
                        }
                    }
                }
                catch (Exception e)
                {
                    ret = false;
                }

                sw.Flush();
                sw.Close();
                fs.Close();

                try
                {
                    for (int i = 1; i < dirPhotosi.Length; i++)
                    {
                        File.Copy(dirPhotosi[0] + "\\" + nomeFile, dirPhotosi[i] + "\\" + nomeFile);
                    }
                }
                catch
                {
                    //se da errore la copia continua ugualmente
                }

            }

            return ret;

        }

        private bool GeneraFileMovimenti(string strDir)
        {
            bool ret = true;

            string nomeFile = dtData.Value.Year.ToString().PadLeft(4, '0') +
                dtData.Value.Month.ToString().PadLeft(2, '0') +
                dtData.Value.Day.ToString().PadLeft(2, '0') +
                "-Movimenti" + tbNumBolla.Text.PadLeft(4, '0') +
                ".txt";

            FileStream fs = new FileStream(strDir + "\\" + nomeFile, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.ASCII);

            try
            {
                foreach (DataGridViewRow dr in dgOutMovimenti.Rows)
                {
                    string riga = dr.Cells["idBusta"].Value.ToString().PadRight(9, ' ') +
                        ";" +
                        dr.Cells["qta"].Value.ToString().PadLeft(5, '0') +
                        ";" +
                        dr.Cells["idArticolo"].Value.ToString().PadLeft(4, '0') +
                        ";" +
                        dr.Cells["descrizione"].Value.ToString() +
                        ";" +
                        dr.Cells["sconto"].Value.ToString() +
                        ";" +
                        dr.Cells["batch"].Value.ToString() +
                        ";" +
                        dr.Cells["tipoOrdine"].Value.ToString() +
                        ";" +
                        dr.Cells["numProdotti"].Value.ToString();

                    sw.WriteLine(riga);
                }
            }
            catch (Exception e)
            {
                ret = false;
            }

            sw.Flush();
            sw.Close();
            fs.Close();
            return ret;

        }

        private bool GeneraFileTotali(string strDir)
        {
            bool ret = true;

            string nomeFile = dtData.Value.Year.ToString().PadLeft(4, '0') +
                dtData.Value.Month.ToString().PadLeft(2, '0') +
                dtData.Value.Day.ToString().PadLeft(2, '0') +
                "-Totali" + tbNumBolla.Text.PadLeft(4, '0') +
                ".txt";

            FileStream fs = new FileStream(strDir + "\\" + nomeFile, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.ASCII);

            try
            {
                foreach (DataGridViewRow dr in dgOutTotali.Rows)
                {
                    if (Convert.ToInt32(dr.Cells["qta"].Value) > 0)
                    {
                        string riga = dr.Cells["idArticolo"].Value.ToString() +
                            ";" +
                            dr.Cells["qta"].Value.ToString().PadLeft(5, ' ') +
                            ";" +
                            dr.Cells["descrizione"].Value.ToString();

                        sw.WriteLine(riga);
                    }
                }
                sw.WriteLine("Totale quantita': " + lblTotQta.Text);
            }
            catch (Exception e)
            {
                ret = false;
            }

            sw.Flush();
            sw.Close();
            fs.Close();
            return ret;

        }

        private bool GeneraFileEsatto(string strDir, string fornitore)
        {
            bool ret = true;

            string[] dirEsatto = strDir.Split(',');

            if (!string.IsNullOrEmpty(dirEsatto[0]))
            {

                string nomeFile = dtData.Value.Year.ToString().PadLeft(4, '0') +
                    dtData.Value.Month.ToString().PadLeft(2, '0') +
                    dtData.Value.Day.ToString().PadLeft(2, '0') +
                    "-" + fornitore + tbNumBolla.Text.PadLeft(4, '0') +
                    "-Qta" + lblTotQta.Text.PadLeft(4, '0') +
                    ".txt";

                FileStream fs = new FileStream(dirEsatto[0] + "\\" + nomeFile, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs, Encoding.ASCII);

                DataGridView dg = dgOutTotali;
                if (boAccess.leggiAppSettings("BollaRaggruppata") == "1")
                    dg = dgOutTotaliGrup;

                try
                {
                    dg.Sort(dg.Columns["idArticolo"], ListSortDirection.Ascending);

                    foreach (DataGridViewRow dr in dg.Rows)
                    {
                        if (Convert.ToInt32(dr.Cells["qta"].Value) > 0)
                        {
                            int qta = Convert.ToInt32(dr.Cells["qta"].Value);
                            int omaggio = Convert.ToInt32(dr.Cells["omaggio"].Value);
                            if (omaggio != 0)
                                qta -= omaggio;
                            string riga = dr.Cells["idArticolo"].Value.ToString() +
                                "|" +
                                qta.ToString();
                            string sconto = dr.Cells["sconto"].Value.ToString();
                            if (!string.IsNullOrEmpty(sconto) && sconto != "0")
                                riga += "|||||||" + sconto;

                            sw.WriteLine(riga);

                            if (omaggio != 0)
                            {
                                riga = dr.Cells["idArticolo"].Value.ToString() +
                                    "|" +
                                    omaggio.ToString() +
                                    "|||||||||||820";

                                sw.WriteLine(riga);
                            }

                        }
                    }
                }
                catch (Exception e)
                {
                    ret = false;
                }

                sw.Flush();
                sw.Close();
                fs.Close();

                try
                {
                    for (int i = 1; i < dirEsatto.Length; i++)
                    {
                        File.Copy(dirEsatto[0] + "\\" + nomeFile, dirEsatto[i] + "\\" + nomeFile);
                    }
                }
                catch 
                { 
                    //se da errore la copia continua ugualmente
                }
            }

            return ret;

        }

        private bool GeneraFileExcel(string strDir, string fornitore)
        {
            bool ret = true;

            string[] dirFile = strDir.Split(',');

            if (!string.IsNullOrEmpty(dirFile[0]))
            {

                string nomeFile = dtData.Value.Year.ToString().PadLeft(4, '0') +
                    dtData.Value.Month.ToString().PadLeft(2, '0') +
                    dtData.Value.Day.ToString().PadLeft(2, '0') +
                    "-" + fornitore + tbNumBolla.Text.PadLeft(4, '0') +
                    "-Qta" + lblTotQta.Text.PadLeft(4, '0') +
                    ".xls";

                Workbook workbook = new Workbook();
                Worksheet sheet = new Worksheet("Test");
                workbook.Worksheets.Add(sheet);
                sheet.Cells[0, 0] = new Cell("Cod.");
                sheet.Cells[0, 1] = new Cell("Q.tà");
                int row = 0;

                DataGridView dg = dgOutTotali;
                if (boAccess.leggiAppSettings("BollaRaggruppata") == "1")
                    dg = dgOutTotaliGrup;

                try
                {
                    dg.Sort(dg.Columns["idArticolo"], ListSortDirection.Ascending);

                    foreach (DataGridViewRow dr in dg.Rows)
                    {
                        if (Convert.ToInt32(dr.Cells["qta"].Value) > 0)
                        {
                            articolo articolo = boAccess.LeggiArticolo(dr.Cells["idArticolo"].Value.ToString());

                            string codArticolo = dr.Cells["idArticolo"].Value.ToString();
                            if (!string.IsNullOrEmpty(articolo.codiceFatturazione))
                                codArticolo = articolo.codiceFatturazione;

                            int qta = Convert.ToInt32(dr.Cells["qta"].Value);
                            int omaggio = Convert.ToInt32(dr.Cells["omaggio"].Value);
                            if (omaggio != 0)
                                qta -= omaggio;
                            
                            row++;
                            sheet.Cells[row, 0] = new Cell(codArticolo);
                            sheet.Cells[row, 1] = new Cell(qta);
                        }
                    }
                }
                catch (Exception e)
                {
                    ret = false;
                }

                try
                {
                    workbook.Save(dirFile[0] + "\\" + nomeFile);

                    for (int i = 1; i < dirFile.Length; i++)
                    {
                        File.Copy(dirFile[0] + "\\" + nomeFile, dirFile[i] + "\\" + nomeFile);
                    }
                }
                catch
                {
                    //se da errore la copia continua ugualmente
                }
            }

            return ret;

        }

        private bool GeneraFileTotaliGrup(string strDir, out string nomeFile)
        {
            bool ret = true;

            nomeFile = dtData.Value.Year.ToString().PadLeft(4, '0') +
                dtData.Value.Month.ToString().PadLeft(2, '0') +
                dtData.Value.Day.ToString().PadLeft(2, '0') +
                "-TotaliGrup" + tbNumBolla.Text.PadLeft(4, '0') +
                ".txt";

            FileStream fs = new FileStream(strDir + "\\" + nomeFile, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.ASCII);

            try
            {
                foreach (DataGridViewRow dr in dgOutTotaliGrup.Rows)
                {
                    if (Convert.ToInt32(dr.Cells["qta"].Value) > 0)
                    {
                        string riga = dr.Cells["idArticolo"].Value.ToString() +
                            ";" +
                            dr.Cells["qta"].Value.ToString().PadLeft(5, ' ') +
                            ";" +
                            dr.Cells["descrizione"].Value.ToString();

                        sw.WriteLine(riga);
                    }
                }
                sw.WriteLine("Totale quantita': " + lblTotQta.Text);
            }
            catch (Exception e)
            {
                ret = false;
            }

            sw.Flush();
            sw.Close();
            fs.Close();
            return ret;

        }

        private bool GeneraFileBolla(string strDir)
        {
            bool ret = true;

            string nomeFile = dtData.Value.Year.ToString().PadLeft(4, '0') +
                dtData.Value.Month.ToString().PadLeft(2, '0') +
                dtData.Value.Day.ToString().PadLeft(2, '0') +
                "-Bolla" + tbNumBolla.Text.PadLeft(4, '0') +
                ".cvs";

            FileStream fs = new FileStream(strDir + "\\" + nomeFile, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.ASCII);

            try
            {
                sw.WriteLine("delay(3000)");
                foreach (DataGridViewRow dr in dgOutMovimenti.Rows)
                {
                    for (int i = 0; i < dr.Cells["idArticolo"].Value.ToString().Length; i++)
                    {
                        sw.WriteLine("KD(" + dr.Cells["idArticolo"].Value.ToString().Substring(i,1) + ")");
                        sw.WriteLine("delay(30)");
                    }
                    sw.WriteLine("KB_CLK(enter)");
                    sw.WriteLine("delay(30)");
                    for (int i = 0; i < dr.Cells["qta"].Value.ToString().Length; i++)
                    {
                        sw.WriteLine("KD(" + dr.Cells["qta"].Value.ToString().Substring(i, 1) + ")");
                        sw.WriteLine("delay(30)");
                    }
                    sw.WriteLine("KB_CLK(enter)");
                    sw.WriteLine("delay(30)");
                    sw.WriteLine("KB_CLK(enter)");
                    sw.WriteLine("delay(30)");
                    sw.WriteLine("KB_CLK(enter)");
                    sw.WriteLine("delay(30)");
                    sw.WriteLine("KB_CLK(enter)");
                    sw.WriteLine("delay(30)");
                    sw.WriteLine("KB_CLK(enter)");
                    sw.WriteLine("delay(30)");
                    sw.WriteLine("KB_CLK(enter)");
                    sw.WriteLine("delay(30)");
                }
            }
            catch (Exception e)
            {
                ret = false;
            }

            sw.Flush();
            sw.Close();
            fs.Close();
            return ret;

        }

        private bool GeneraFileBusteIn(string strDir)
        {
            bool ret = true;

            string nomeFile = dtData.Value.Year.ToString().PadLeft(4, '0') +
                dtData.Value.Month.ToString().PadLeft(2, '0') +
                dtData.Value.Day.ToString().PadLeft(2, '0') +
                "-BusteIn.txt";

            try
            {
                FileMode fm = FileMode.Create;
                if (File.Exists(strDir + "\\" + nomeFile))
                    fm = FileMode.Append;
                FileStream fs = new FileStream(strDir + "\\" + nomeFile, fm);
                StreamWriter sw = new StreamWriter(fs, Encoding.ASCII);

                try
                {
                    foreach (DataGridViewRow dr in dgInBuste.Rows)
                    {
                        string riga = dr.Cells["idBusta"].Value.ToString().PadRight(9, ' ') +
                            ";" +
                            dr.Cells["batch"].Value.ToString().PadLeft(5, '0') +
                            ";" +
                            dr.Cells["qta"].Value.ToString().PadLeft(5, '0') +
                            ";" +
                            dr.Cells["idArticolo"].Value.ToString().PadLeft(4, '0') +
                            ";" +
                            dr.Cells["descrizione"].Value.ToString();

                        sw.WriteLine(riga);
                    }
                }
                catch (Exception e)
                {
                    ret = false;
                }

                sw.Flush();
                sw.Close();
                fs.Close();
            }
            catch (Exception e)
            {
                ret = false;
            }

            return ret;

        }

        private bool GeneraFileBusteResi(string strDir)
        {
            bool ret = true;

            string nomeFile = dtData.Value.Year.ToString().PadLeft(4, '0') +
                dtData.Value.Month.ToString().PadLeft(2, '0') +
                dtData.Value.Day.ToString().PadLeft(2, '0') +
                "-BusteResi.txt";

            FileMode fm = FileMode.Create;
            if (File.Exists(strDir + "\\" + nomeFile))
                fm = FileMode.Append;
            FileStream fs = new FileStream(strDir + "\\" + nomeFile, fm);
            StreamWriter sw = new StreamWriter(fs, Encoding.ASCII);

            try
            {
                foreach (DataGridViewRow dr in dgInBuste.Rows)
                {
                    string riga = dr.Cells["idBusta"].Value.ToString().PadRight(9, ' ') +
                        ";" +
                        dr.Cells["batch"].Value.ToString().PadLeft(5, '0') +
                        ";" +
                        dr.Cells["qta"].Value.ToString().PadLeft(5, '0') +
                        ";" +
                        dr.Cells["idArticolo"].Value.ToString().PadLeft(4, '0') +
                        ";" +
                        dr.Cells["descrizione"].Value.ToString();

                    sw.WriteLine(riga);
                }
            }
            catch (Exception e)
            {
                ret = false;
            }

            sw.Flush();
            sw.Close();
            fs.Close();
            return ret;

        }

        private void inizializzaDgOutMovimenti()
        {
            dgOutMovimenti.Columns.Add("N", "N");
            dgOutMovimenti.Columns["N"].Width = 30;
            dgOutMovimenti.Columns.Add("idArticolo", "idArticolo");
            dgOutMovimenti.Columns.Add("descrizione", "descrizione");
            dgOutMovimenti.Columns.Add("idBusta", "idBusta");
            dgOutMovimenti.Columns.Add("batch", "batch");
            dgOutMovimenti.Columns.Add("data", "data");
            dgOutMovimenti.Columns.Add("numBolla", "numBolla");
            dgOutMovimenti.Columns.Add("qta", "qta");
            dgOutMovimenti.Columns.Add("sconto", "sconto");
            dgOutMovimenti.Columns.Add("tipoOrdine", "tipoOrdine");
            dgOutMovimenti.Columns.Add("numProdotti", "numProdotti");
        }

        private void inizializzaDgOutTotali()
        {
            dgOutTotali.Columns.Add("idArticolo", "idArticolo");
            dgOutTotali.Columns.Add("descrizione", "descrizione");
            dgOutTotali.Columns.Add("qta", "qta");
            dgOutTotali.Columns.Add("sconto", "sconto");
            dgOutTotali.Columns.Add("omaggio", "omaggio");

            dgOutTotaliGrup.Columns.Add("idArticolo", "idArticolo");
            dgOutTotaliGrup.Columns.Add("descrizione", "descrizione");
            dgOutTotaliGrup.Columns.Add("qta", "qta");
        }

        private void inizializzaDgInBuste()
        {
            dgInBuste.Columns.Add("N", "N");
            dgInBuste.Columns["N"].Width = 40;
            dgInBuste.Columns.Add("idArticolo", "idArticolo");
            dgInBuste.Columns.Add("descrizione", "descrizione");
            dgInBuste.Columns.Add("idBusta", "idBusta");
            dgInBuste.Columns.Add("batch", "batch");
            dgInBuste.Columns.Add("data", "data");
            dgInBuste.Columns.Add("qta", "qta");
            dgInBuste.Columns.Add("tipoOrdine", "tipoOrdine");
            dgInBuste.Columns.Add("numProdotti", "numProdotti");
        }

        private void inizializzaDgInTotali()
        {
            dgInTotali.Columns.Add("idArticolo", "idArticolo");
            dgInTotali.Columns.Add("descrizione", "descrizione");
            dgInTotali.Columns.Add("qta", "qta");
        }

        private void inizializzaDgResiBuste()
        {
            dgResiBuste.Columns.Add("N", "N");
            dgResiBuste.Columns["N"].Width = 40;
            dgResiBuste.Columns.Add("idArticolo", "idArticolo");
            dgResiBuste.Columns.Add("descrizione", "descrizione");
            dgResiBuste.Columns.Add("idBusta", "idBusta");
            dgResiBuste.Columns.Add("batch", "batch");
            dgResiBuste.Columns.Add("data", "data");
            dgResiBuste.Columns.Add("qta", "qta");
        }

        private void StampaMessaggio(string messaggio, bool focus, TipoSuono tipoSuono)
        {
            try
            {
                if (tipoSuono == TipoSuono.Errore)
                    PlaySound(suonoErrore, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                else if (tipoSuono == TipoSuono.OkInserimento)
                    PlaySound(suonoOKInserimento, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                else if (tipoSuono == TipoSuono.OkInserimentoMulti)
                    PlaySound(suonoOKInserimentoMulti, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                else if (tipoSuono == TipoSuono.OkInserimentoMonoprodotto)
                    PlaySound(suonoOKInserimentoMonoprodotto, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                else if (tipoSuono == TipoSuono.Domanda)
                    PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                else if (tipoSuono == TipoSuono.Kit)
                    PlaySound(boAccess.leggiAppSettings("SuonoKit"), IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
            }
            catch { }
            if(messaggio != string.Empty)
                lblErrore.Text = messaggio;
            
            if (focus)
            {
                tbOutCodice.Text = string.Empty;
                tbOutCodice.Focus();
            }
        }

        private void frmBuste_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (modificatoOut || modificatoIn || modificatoResi)
            {
                PlaySound(suonoErrore, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                DialogResult dr = MessageBox.Show("Attenzione, non sono state salvate le modifiche. Vuoi uscire ugualmente?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if(dr == DialogResult.No)
                    e.Cancel = true;
                else
                    boAccess.ChiudiConnessione();
            }
            else
                boAccess.ChiudiConnessione();
        }

        private void tabBuste_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tab1.SelectedTab == this.tabIn)
            {
                tbInCodice.Focus();
                this.AcceptButton = btInInserisce;
            }
            else if (tab1.SelectedTab == this.tabOut)
            {
                tbOutCodice.Focus();
                this.AcceptButton = btOutInserisce;
            }
            else if (tab1.SelectedTab == this.tabResi)
            {
                tbCodiceResi.Focus();
                this.AcceptButton = btInserisceResi;
            }
        }

        private void tabArticoli_Click(object sender, EventArgs e)
        {
            int a = 1;
        }

        private void eliminaUltimoMovimento()
        {
            DataGridViewRow row;
            if (dgOutMovimenti.Rows.Count > 0)
            {
                row = dgOutMovimenti.Rows[dgOutMovimenti.Rows.Count - 1];
                dgOutMovimenti.Rows.Remove(row);
            }
        }

        private void btSalvaTemp_Click(object sender, EventArgs e)
        {
            string nomeFile;
            string DirTemp = ConfigurationSettings.AppSettings["DirTemp"];
            if (string.IsNullOrEmpty(DirTemp))
                DirTemp = boAccess.leggiAppSettings("DirTemp");
            if (!GeneraFilePhotosi(DirTemp) || !GeneraFileBolla(DirTemp) || !GeneraFileTotali(DirTemp) || !GeneraFileTotaliGrup(DirTemp, out nomeFile) || !GeneraFileMovimenti(DirTemp))
                StampaMessaggio("Errore in fase di scrittura file temporaneo.", true, TipoSuono.Errore);
            else
                StampaMessaggio("File salvati in cartella temporanea.", true, TipoSuono.NoSuono);
        }

        private void btCaricaTemp_Click(object sender, EventArgs e)
        {
            if (dgOutMovimenti.Rows.Count > 0)
                StampaMessaggio("Attenzione. Esistono già movimenti inseriti.", true, TipoSuono.Errore);
            else
            {
                string DirTemp = ConfigurationSettings.AppSettings["DirTemp"];
                if (string.IsNullOrEmpty(DirTemp))
                    DirTemp = boAccess.leggiAppSettings("DirTemp");

                ofdFileTemp.Filter = "Movimenti .txt|?????????Movimenti*.txt";
                ofdFileTemp.InitialDirectory = DirTemp;
                ofdFileTemp.ShowDialog();
                if (ofdFileTemp.FileName != "")
                {
                    try
                    {
                        using (StreamReader sr = new StreamReader(ofdFileTemp.FileName))
                        {
                            String line;
                            while ((line = sr.ReadLine()) != null)
                            {
                                string[] split = line.Split(';');
                                string idBatch = split.Length > 5 ? split[5] : "";
                                AggiungeDGOutMovimento(split[2], split[3], split[0], dtData.Value, Convert.ToInt32(tbNumBolla.Text), Convert.ToInt32(split[1]), Convert.ToDecimal(split[4]), idBatch, split[6], split[7]);
                            }
                        }
                        modificatoOut = true;
                        //RicalcolaTotali();
                        tbOutCodice.Focus();
                    }
                    catch (Exception ex)
                    {
                        StampaMessaggio("Errore in lettura del file. " + ex.Message, true, TipoSuono.Errore);
                    }
                }
            }
        }

        private void btInInserisce_Click(object sender, EventArgs e)
        {
            string fornitore = boAccess.leggiAppSettings("Fornitore");
            lblErrore.Text = string.Empty;
            string codArticolo = string.Empty;
            int qtaArticolo = 0;
            int secondiDoppiaLettura = 0;
            string valore = boAccess.leggiAppSettings("SecondiDoppiaLettura");
            if (valore != string.Empty)
                secondiDoppiaLettura = Convert.ToInt32(valore);
            else
                secondiDoppiaLettura = 0;

            //sistema la data
            if (DateTime.Now.ToShortDateString() != dtData.Value.ToShortDateString()
                && dgInBuste.Rows.Count == 0)
                dtData.Value = DateTime.Now;

            if (lblArticoloIn.Text != string.Empty && lblBustaIn.Text != string.Empty && !lblKitIn.Visible)
            {
                lblArticoloIn.Text = string.Empty;
                lblDesArticoloIn.Text = string.Empty;
                lblQtaIn.Text = string.Empty;
                lblBustaIn.Text = string.Empty;
            }
            string codice = tbInCodice.Text;
            if (codice == string.Empty)
                return;

            TimeSpan duration = (DateTime.Now - ultimoCodiceLettoTime);
            if (codice == ultimoCodiceLetto && duration.Seconds < secondiDoppiaLettura)
            {
                StampaMessaggio("Errore. Doppia lettura codice.", true, TipoSuono.NoSuono);
                return;
            }
            ultimoCodiceLetto = tbInCodice.Text;
            ultimoCodiceLettoTime = DateTime.Now;

            if (codice == string.Empty)
            {
                tbInCodice.Focus();
                return;
            }
            if (codice.ToLower() == "azzerapagina")
            {
                btAzzeraPagina_Click(btAzzeraPaginaIn, null);
                return;
            }
            else if (codice.ToLower() == "eliminaultimo")
            {
                btEliminaUltimo_Click(btInInserisce, null);
                return;
            }
            else if (codice.ToLower() == "iniziokit")
            {
                if (lblBustaIn.Text != string.Empty || lblArticoloIn.Text != string.Empty) //busta piena -> errore
                {
                    StampaMessaggio("Errore. Il codice Kit va letto con pagina vuota.", true, TipoSuono.Errore);
                    return;
                }
                StampaMessaggio("", true, TipoSuono.Kit);
                lblKitIn.Visible = true;
                tbInCodice.Text = string.Empty;
                tbInCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "elultimomov")
            {
                PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                DialogResult dr = MessageBox.Show("Si desidera eliminare l'ultimo movimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                if (dr == DialogResult.Yes)
                {
                    eliminaUltimoMovimento();
                }
                else
                {
                    tbInCodice.Text = string.Empty;
                    tbInCodice.Focus();
                    return;
                }
            }
            else if (codice.ToLower() == "finekit")
            {
                if (!lblKitIn.Visible)
                {
                    StampaMessaggio("Codice finekit senza iniziokit.", true, TipoSuono.Errore);
                    return;
                }
                StampaMessaggio("", true, TipoSuono.Kit);
                lblKitIn.Visible = false;
                tbInCodice.Text = string.Empty;
                tbInCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "salvatemp")
            {
                btSalvaTemp_Click(null, null);
                tbInCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "caricatemp")
            {
                btCaricaTemp_Click(null, null);
                tbInCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "in")
            {
                tbInCodice.Text = string.Empty;
                tbInCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "out")
            {
                tab1.SelectedTab = tabOut;
                tbInCodice.Text = string.Empty;
                tbInCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "fine")
            {
                TBFineIn_Click(tbInCodice, null);
                return;
            }
            else if (codice.ToLower() == "aggscura" &&
                ((lblArticoloIn.Text == string.Empty && lblBustaIn.Text == string.Empty) ||
                lblKitIn.Visible))
            {
                codArticolo = boAccess.leggiAppSettings("CodStampaScura");
                articolo articolo = boAccess.LeggiArticolo(codArticolo);
                lblArticoloIn.Text = codArticolo;
                lblDesArticoloIn.Text = articolo.descrizione;
                lblQtaIn.Text = ultimaQtaLetta;
                lblBustaIn.Text = ultimaBustaLetta;
            }
            else if (codice.Substring(codice.Length - 1, 1) == "*")
            {
                try
                {
                    qtaArticolo = Convert.ToInt32(codice.Substring(0, codice.Length - 1));
                }
                catch
                {
                    StampaMessaggio("Errore. Quantità non valida.", true, TipoSuono.Errore);
                    return;
                }
                return;
            }
            else if (fornitore.ToLower() == "photosi" && (codice.Length == 11 || codice.Length == 4 || codice.IndexOf("*") > 0) ||
                    fornitore.ToLower() == "fotoevolution" && (codice.Length == 12 && codice.Substring(0, 5) != "00000") || (codice.Length == 6 && lblBustaOut.Text != string.Empty) || codice.IndexOf("*") > 0) //codice articolo
            {
                if (lblArticoloIn.Text != string.Empty && !lblKitIn.Visible) //articolo pieno -> errore
                {
                    StampaMessaggio("Errore. Lettura doppia codice articolo.", true, TipoSuono.Errore);
                    return;
                }
                if (lblKitIn.Visible && lblBustaIn.Text == string.Empty)
                {
                    StampaMessaggio("Errore. Con il Kit leggere per primo il codice busta.", true, TipoSuono.Errore);
                    return;
                }

                int posX = codice.IndexOf("*");
                if (posX == 4)
                {
                    codArticolo = codice.Substring(0, 4);
                    try
                    {
                        qtaArticolo = Convert.ToInt32(codice.Substring(5));
                    }
                    catch
                    {
                        qtaArticolo = 1;
                    }
                }
                if (posX < 4 && posX > 0)
                {
                    codArticolo = codice.Substring(posX + 1, 4);
                    try
                    {
                        qtaArticolo = Convert.ToInt32(codice.Substring(0, posX));
                    }
                    catch
                    {
                        qtaArticolo = 1;
                    }
                }
                else if (fornitore.ToLower() == "photosi" && codice.Length == 11)
                {
                    codArticolo = codice.Substring(5, 4);
                    try
                    {
                        qtaArticolo = Convert.ToInt32(codice.Substring(0, 5));
                    }
                    catch { return; }
                }
                else if (fornitore.ToLower() == "photosi" && codice.Length == 4)
                {
                    codArticolo = codice;
                    qtaArticolo = 1;
                }
                else if (fornitore.ToLower() == "fotoevolution" && codice.Length == 12)
                {
                    codArticolo = codice.Substring(0, 6);
                    try
                    {
                        qtaArticolo = Convert.ToInt32(codice.Substring(6, 5));
                    }
                    catch { return; }
                }
                else if (fornitore.ToLower() == "fotoevolution" && codice.Length == 6)
                {
                    codArticolo = codice;
                    qtaArticolo = 1;
                }

                if (qtaArticolo < 1)
                {
                    StampaMessaggio("Errore. Quantita a zero.", true, TipoSuono.Errore);
                    return;
                }

                articolo articolo = boAccess.LeggiArticolo(codArticolo);
                if (articolo.codice == string.Empty)
                {
                    StampaMessaggio("Errore. Codice articolo " + codArticolo + " inesistente.", true, TipoSuono.Errore);
                    return;
                }
                if (boAccess.leggiAppSettings("MinQtaMessaggio").ToString() != string.Empty && qtaArticolo > Convert.ToInt32(boAccess.leggiAppSettings("MinQtaMessaggio")))
                {
                    PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                    DialogResult dr = MessageBox.Show("Attenzione, quantità = " + qtaArticolo.ToString() + ". Confermi?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                    if (dr == DialogResult.No)
                    {
                        tbInCodice.Text = string.Empty;
                        return;
                    }
                }
                lblArticoloIn.Text = codArticolo;
                lblDesArticoloIn.Text = articolo.descrizione;
                lblQtaIn.Text = qtaArticolo.ToString();

                ultimoArticoloLetto = codArticolo;
                ultimaQtaLetta = qtaArticolo.ToString();

                if (lblBustaIn.Text == string.Empty) //articolo vuoto e busta vuota
                {
                    tbInCodice.Text = string.Empty;
                    tbInCodice.Focus();
                    return;
                }
            }
            else if (isCodiceBusta(codice)) //codice busta
            {
                if (fornitore.ToLower() == "fotoevolution" && codice.Length > 6)
                    codice = Convert.ToString(Convert.ToInt32(codice.Substring(0, 11)));
                if (!inserisceCodiceBustaIn(codice, ckRifacimento.Checked))
                    return;
            }
            else
            {
                //StampaMessaggio("Codice non riconosciuto. E' un codice busta?", false, TipoSuono.Domanda);
                //lCodErrore.Text = "1";
                //tbErrore.Enabled = true;
                //this.AcceptButton = btErrore;
                //tbErrore.Text = string.Empty;
                //tbErrore.Select();
                //tbErrore.Focus();
                //return;
                PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                DialogResult dr = MessageBox.Show("Codice busta non riconosciuto. Confermi ugualmente?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.No)
                {
                    tbInCodice.Text = string.Empty;
                    return;
                }
                else
                {
                    lblBustaIn.Text = codice;
                    ultimaBustaLetta = codice;
                }
            }

            if (lblQtaIn.Text == string.Empty)
                qtaArticolo = 0;
            else
                qtaArticolo = Convert.ToInt32(lblQtaIn.Text);

            codArticolo = lblArticoloIn.Text;
            string codBusta = lblBustaIn.Text;
            if (qtaArticolo > 0)
            {
                string idBatch = string.Empty;
                if (ckRifacimento.Checked)
                    idBatch = "999999";
                AggiungeDGInBuste(idBatch, codArticolo, lblDesArticoloIn.Text, codBusta, dtData.Value, qtaArticolo);
                AggiungeCodiceArticoloInTotali(codArticolo, lblDesArticoloIn.Text, qtaArticolo, 0, "", "In");
                //AggiungeCodiceArticoloInTotaliGrup(codArticolo, lblDesArticoloIn.Text, qtaArticolo, "");
                StampaMessaggio("", true, TipoSuono.OkInserimento);
                modificatoIn = true;
            }
            tbInCodice.Text = string.Empty;
            tbInCodice.Focus();


            //lblInErrore.Text = string.Empty;
            //bool ok = true;

            //string codice = tbInCodice.Text;
            //if (codice == string.Empty)
            //{
            //    tbInCodice.Focus();
            //    return;
            //}

            //if (isCodiceBusta(codice)) //codice busta
            //{
            //    if (!inserisceCodiceBustaIn(codice))
            //        ok = false;
            //}
            //else
            //{
            //    PlaySound(boAccess.leggiAppSettings("SuonoConferma"), IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
            //    DialogResult dr = MessageBox.Show("Codice busta non riconosciuto. Confermi ugualmente?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
            //    if (dr == DialogResult.No)
            //    {
            //        ok = false;
            //    }
            //}

            //if (ok)
            //{
            //    lblBustaIn.Text = tbInCodice.Text;
            //    AggiungeDGInBuste("", "", "", tbInCodice.Text, dtData.Value, 0);
            //    StampaMessaggio("", true, TipoSuono.OkInserimento);
            //    modificatoIn = true;
            //}

            //tbInCodice.Text = string.Empty;
            //tbInCodice.Focus();
        }

        private void btInInserisceOld_Click(object sender, EventArgs e)
        {
            lblInErrore.Text = string.Empty;
            bool ok = true;

            string codice = tbInCodice.Text;
            if (codice == string.Empty)
            {
                tbInCodice.Focus();
                return;
            }

            if (isCodiceBusta(codice)) //codice busta
            {
                if (!inserisceCodiceBustaIn(codice, ckRifacimento.Checked))
                    ok = false;
            }
            else
            {
                PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                DialogResult dr = MessageBox.Show("Codice busta non riconosciuto. Confermi ugualmente?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.No)
                {
                    ok = false;
                }
            }

            if (ok)
            {
                lblBustaIn.Text = tbInCodice.Text;
                AggiungeDGInBuste("", "", "", tbInCodice.Text, dtData.Value, 0);
                StampaMessaggio("", true, TipoSuono.OkInserimento);
                modificatoIn = true;
            }

            tbInCodice.Text = string.Empty;
            tbInCodice.Focus();
        }

        private void btEliminaUltimo_Click(object sender, EventArgs e)
        {
            if (dgOutMovimenti.RowCount > 0)
            {
                DataGridViewRow row = dgOutMovimenti.Rows[dgOutMovimenti.RowCount - 1];
                dgOutMovimenti.Rows.Remove(row);
                StampaMessaggio("Eliminato ultimo movimento. ", true, TipoSuono.NoSuono);
            }
        }

        private void tabOutInt_TabIndexChanged(object sender, EventArgs e)
        {
            tbOutCodice.Focus();
        }

        private void tabOutInt_SelectedIndexChanged(object sender, EventArgs e)
        {
            tbOutCodice.Focus();
        }

        private bool isCodiceBusta(string codice)
        {
            if (!boCommon.isNumerico(codice))
                return false;
            string[] MaschereBuste = boAccess.leggiAppSettings("MaschereBuste").Split(',');
            for (int i1 = 0; i1 < MaschereBuste.Length; i1++)
            {
                if (codice.Length == MaschereBuste[i1].Length)
                {
                    bool okCodice = true;
                    for (int i2 = 0; i2 < MaschereBuste[i1].Length; i2++)
                    {
                        if (MaschereBuste[i1].Substring(i2, 1) != "*" && MaschereBuste[i1].Substring(i2, 1) != codice.Substring(i2, 1))
                        {
                            okCodice = false;
                            break;
                        }
                            
                    }
                    if (okCodice)
                        return true;
                }
            }
            
            return false;

            //if(codice.Length == 9 && (codice.Substring(0,1)=="5" || codice.Substring(0,1) == "8") && boCommon.isNumerico(codice))
            //    return true;
            //if (codice.Length == 8 && (codice.Substring(0,1)=="1") && boCommon.isNumerico(codice))
            //    return true;
            //if (codice == boAccess.leggiAppSettings("BustaFittizia"))
            //    return true;

            //return false;

        }

        private void btErrore_Click(object sender, EventArgs e)
        {
            if (lCodErrore.Text == "1")
            {
                if (tbErrore.Text.Substring(0,1).ToLower() == "s")
                {
                    inserisceCodiceBustaOut(tbOutCodice.Text);
                    tbOutCodice.Text = string.Empty;
                    tbOutCodice.Focus();
                }
                else if (tbErrore.Text.Substring(0, 1).ToLower() == "n")
                {
                    tbOutCodice.Text = string.Empty;
                    tbOutCodice.Focus();
                }
                if (tbErrore.Text.Substring(0, 1).ToLower() == "s" ||
                    tbErrore.Text.Substring(0, 1).ToLower() == "n")
                {
                    this.AcceptButton = btOutInserisce;
                    lblErrore.Text = "";
                    tbErrore.Enabled = false;
                    tbErrore.Focus();
                }
                tbErrore.Text = "";
            }
        }

        private void TBFineIn_Click(object sender, EventArgs e)
        {
            if (dgInBuste.Rows.Count == 0)
            {
                StampaMessaggio("Non ci sono buste da salvare.", true, TipoSuono.Errore);
                return;
            }
            bool ret = ScriveMovimentiIn(dgInBuste);
            string nomeFile = string.Empty;
            if (ret)
            {
                //DialogResult dr = MessageBox.Show("Scrittura eseguita con successo. Vuoi generare i file?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                //if (dr == DialogResult.Yes)
                //{
                if (!GeneraFileBusteIn(boAccess.leggiAppSettings("DirFileBolla")))
                    if (sender != timerCaricaBuste)
                        MessageBox.Show("Errore in fase di scrittura", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                //}
                AzzeraInBuste();
                btAzzeraPagina_Click(btAzzeraPaginaIn, null);
                CopiaFileConsuntivi();
                modificatoIn = false;
                ConsuntiviImport.Items.Clear();
                if (sender != tbInCodice)
                {
                    if (sender != timerCaricaBuste)
                        MessageBox.Show("Scrittura eseguita con successo.", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    StampaMessaggio("", true, TipoSuono.OkInserimento);
                }
            }
            else
            {
                if (sender != timerCaricaBuste)
                    MessageBox.Show("Errore in fase di scrittura", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public bool ScriveMovimentiIn(DataGridView dgBuste)
        {
            bool ret = true;

            fmProgress fmPbar = new fmProgress();

            Label pBarLabel = ((Label)fmPbar.Controls["Label1"]);
            string testoLabel = "Lettura in corso buste in entrata. Attendere.\r\nBusta {0} di {1}.";
            ProgressBar pBarMovimenti = ((ProgressBar)fmPbar.Controls["progressBar1"]);
            fmPbar.Show();
            this.Cursor = Cursors.WaitCursor;

            pBarMovimenti.Minimum = 0;
            pBarMovimenti.Maximum = dgBuste.Rows.Count;
            pBarMovimenti.Value = 0;

            foreach (DataGridViewRow dr in dgBuste.Rows)
            {
                pBarMovimenti.Value += 1;
                pBarLabel.Text = string.Format(testoLabel, pBarMovimenti.Value, pBarMovimenti.Maximum);
                fmPbar.Refresh();

                //DateTime adesso = DateTime.Now;
                //DateTime inizio = adesso.AddSeconds(3);
                //while (adesso < inizio) { adesso = DateTime.Now; }

                try
                {
                    movimento mov = new movimento();
                    mov.idBusta = dr.Cells["idBusta"].Value.ToString();
                    mov.idBatch = dr.Cells["batch"].Value.ToString();
                    mov.idArticolo = dr.Cells["idArticolo"].Value.ToString();
                    mov.desArticolo = dr.Cells["descrizione"].Value.ToString();
                    mov.data = Convert.ToDateTime(dr.Cells["data"].Value);
                    mov.qta = Convert.ToInt32(dr.Cells["qta"].Value);
                    mov.tipoOrdine = dr.Cells["tipoOrdine"].Value.ToString();
                    mov.numProdotti = dr.Cells["numProdotti"].Value.ToString();

                    if (!boAccess.ScriveMovimentoIn(mov, false))
                        throw new Exception();
                }

                //Some usual exception handling
                catch (Exception e)
                {
                    //tx.Rollback();
                    //aConnection.Close();
                    Console.WriteLine("Error: {0}", e.Message);
                    ret = false;
                }
            }

            fmPbar.Dispose();
            //tx.Commit();
            //aConnection.Close();
            this.Cursor = Cursors.Default;

            return ret;

        }


        private void CopiaFileConsuntivi()
        {
            foreach (string item in ConsuntiviImport.Items)
            {
                string fileBackup = item.Replace(".csv", "") + ".txt";
                FileStream fs = File.Create(fileBackup);
                fs.Close();
            }
        }

        private void CopiaFileConsuntiviOld()
        {
            foreach (string item in ConsuntiviImport.Items)
            {
                File.Copy(boAccess.leggiAppSettings("DirConsuntivi") + "\\" + item, boAccess.leggiAppSettings("DirConsuntivi") + "\\imp\\" + item, true);
            }
        }

        private void btInCaricaBuste_Click(object sender, EventArgs e)
        {
            string DirTemp = ConfigurationSettings.AppSettings["DirTemp"];
            DateTime? DaDataLettura = null;
            try
            {
                DaDataLettura = Convert.ToDateTime(ConfigurationSettings.AppSettings["DaDataLettura"]);
            }
            catch { }

            if (string.IsNullOrEmpty(DirTemp))
                DirTemp = boAccess.leggiAppSettings("DirTemp");

            string DirPhotosi = boAccess.leggiAppSettings("DirLavorazioni");
            string DirPhotosiMascherine = DirPhotosi.Contains("Photosi") ? DirPhotosi + "Mascherine" : "";
            string fileDirJob = DirTemp + "\\dirjob.txt";
            Process myProcess = new Process();
            myProcess.StartInfo.FileName = "cmd.exe";
            myProcess.StartInfo.Arguments = "/c dir " + DirPhotosi + "\\job.csv " + DirPhotosiMascherine + "\\job.csv /s/b>" + fileDirJob;
            myProcess.StartInfo.UseShellExecute = false;
            myProcess.StartInfo.ErrorDialog = false;
            myProcess.StartInfo.RedirectStandardOutput = true;
            myProcess.StartInfo.CreateNoWindow = true;
            myProcess.StartInfo.RedirectStandardError = true;
            if (sender == btInCaricaBuste)
            {
                DialogResult dr = MessageBox.Show(myProcess.StartInfo.Arguments + "\r\n\r\nConfermi comando?", "Buste", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.No)
                    return;
            }
            myProcess.Start();
            string processError = myProcess.StandardError.ReadToEnd();
            myProcess.WaitForExit();

            //if (!string.IsNullOrEmpty(processError) && sender != timerCaricaBuste)
            //{
            //    MessageBox.Show("Errore in fase di caricamento buste in entrata.\r\n" + processError + myProcess.StartInfo.Arguments, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    return;
            //}

            StreamReader sr = null;
            try
            {
                sr = File.OpenText(fileDirJob);
                String input = String.Empty;
                while ((input = sr.ReadLine()) != null)
                {
                    string pathfilecsv = input;
                    int posiz = pathfilecsv.LastIndexOf('\\');
                    string datastr = pathfilecsv.Substring(0, posiz);
                    posiz = datastr.LastIndexOf('\\');
                    datastr = datastr.Substring(0, posiz);
                    posiz = datastr.LastIndexOf('\\');
                    datastr = datastr.Substring(posiz+1);
                    DateTime? databatch = null;
                    try
                    {
                        databatch = Convert.ToDateTime(datastr);
                    }
                    catch { }
                    if (databatch != null && databatch < DaDataLettura)
                        continue;

                    string fileBackup = pathfilecsv.Replace(".csv", "") + ".txt";
                    if (!File.Exists(fileBackup))
                    {
                        bool isBatchMascherine = fileBackup.Contains(DirPhotosiMascherine) ? true : false;
                        string[] riga = input.Split('\\');
                        string batch = riga[riga.Length - 2];
                        string giorno = riga[riga.Length - 3];

                        string idBusta = string.Empty;
                        string idBatch = string.Empty;
                        string idArticolo = string.Empty;
                        string desArticolo = string.Empty;
                        string qta = "";
                        int qtaMascherine = 0;
                        DateTime data = DateTime.Now;
                        string tipoOrdine = string.Empty;
                        string numProdotti = string.Empty;

                        StreamReader srJob = File.OpenText(pathfilecsv);
                        String inputJob = String.Empty;
                        while ((inputJob = srJob.ReadLine()) != null)
                        {
                            string[] rigaJob = inputJob.Split(';');
                            if (rigaJob.Length == 3)
                            {
                                idBusta = rigaJob[0];
                                desArticolo = rigaJob[1];
                                qta = rigaJob[2];
                                data = Convert.ToDateTime(giorno);
                            }
                            else if (rigaJob.Length >= 6)
                            {
                                idBatch = rigaJob[0];
                                idBusta = rigaJob[1];
                                idArticolo = rigaJob[3];
                                desArticolo = rigaJob[4];
                                qta = rigaJob[5];
                                data = Convert.ToDateTime(giorno);
                                if (isBatchMascherine)
                                    qtaMascherine += Convert.ToInt32(qta);

                            }
                            if (rigaJob.Length >= 8)
                            {
                                tipoOrdine = rigaJob[6];
                                numProdotti = rigaJob[7];
                            }
                            if (qtaMascherine == 0)
                                AggiungeDGInBuste(idBatch, 
                                    idArticolo, 
                                    desArticolo, 
                                    idBusta, 
                                    Convert.ToDateTime(data),
                                    Convert.ToInt32(qta),
                                    tipoOrdine,
                                    numProdotti);
                        }

                        if (qtaMascherine > 0)
                        {
                            int qtaFogli = qtaMascherine / 36;
                            if (qtaMascherine % 36 != 0)
                                qtaFogli += 1;
                            AggiungeDGInBuste(idBatch, 
                                idArticolo, 
                                desArticolo, 
                                idBusta, 
                                Convert.ToDateTime(data), 
                                qtaFogli);
                        }


                        srJob.Close();
                        ConsuntiviImport.Items.Add(pathfilecsv);
                        modificatoIn = true;
                    }
                }

                sr.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore in fase di caricamento buste in entrata.\r\n" + ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            finally
            {
                sr.Close();
            }
            if (ConsuntiviImport.Items.Count <= 0)
                if (sender == btInCaricaBuste)
                    MessageBox.Show("Non sono stati trovati file da importare.", "", MessageBoxButtons.OK, MessageBoxIcon.Information);


            //foreach (string file in Directory.GetFiles(DirConsuntivi,"*.csv"))
            //{
            //    FileInfo fi = new FileInfo(file);
            //    string nomeFile = fi.Name;
            //    string idBusta = string.Empty;
            //    string idBatch = string.Empty;
            //    string idArticolo = string.Empty;
            //    string desArticolo = string.Empty;
            //    string qta = "";
            //    DateTime data = DateTime.Now;
            //    if (!File.Exists(DirConsuntivi + "\\imp\\" + fi.Name))
            //    {
            //        StreamReader sr = File.OpenText(fi.FullName);
            //        String input = String.Empty;
            //        while ((input = sr.ReadLine()) != null)
            //        {
            //            string[] riga = input.Split(';');
            //            if (riga.Length == 3)
            //            {
            //                idBusta = riga[0];
            //                desArticolo = riga[1];
            //                qta = riga[2];
            //                data = Convert.ToDateTime(fi.Name.Replace(fi.Extension, ""));
            //            }
            //            else if (riga.Length == 6)
            //            {
            //                idBatch = riga[0];
            //                idBusta = riga[1];
            //                idArticolo = riga[3];
            //                desArticolo = riga[4];
            //                qta = riga[5];
            //                //data = Convert.ToDateTime(riga[2]);
            //                data = Convert.ToDateTime(fi.Name.Replace(fi.Extension, ""));
            //            }

            //            AggiungeDGInBuste(idBatch, idArticolo, desArticolo, idBusta, Convert.ToDateTime(data), Convert.ToInt32(qta));
            //            ConsuntiviImport.Items.Add(fi.Name);
            //            modificatoIn = true;
            //        }
            //    }
            //}

        }

        private string VerificaConsuntivi()
        {
            string messaggio = string.Empty;
            string DirPhotosi = boAccess.leggiAppSettings("DirLavorazioni");
            for (int i = -1; i > -7; i--)
            {
                DateTime giorno = DateTime.Now.AddDays(i);
                string giornoStr = giorno.ToString("yyyy-MM-dd");
                if (Directory.Exists(DirPhotosi + "\\" + giornoStr))
                {
                    if(File.Exists(DirPhotosi + "\\Consuntivi\\" + giornoStr + ".csv"))
                    {
                        List<busta> listaBuste = CaricaBuste(DirPhotosi + "\\" + giornoStr, "dir");
                        List<busta> listaBusteCons = CaricaBuste(DirPhotosi + "\\Consuntivi\\" + giornoStr + ".csv", "file");

                        if (listaBuste.Count != listaBusteCons.Count)
                            messaggio += "Consuntivi del giorno " + giornoStr + " non corrispondono: " + listaBuste.Count.ToString() + " - " + listaBusteCons.Count.ToString() + "\r\n";

                        foreach (busta bustac in listaBusteCons)
                        {
                            bool trovato = false;
                            foreach (busta busta in listaBuste)
                            {
                                if (bustac.idBatch == busta.idBatch && bustac.idBusta == busta.idBusta)
                                {
                                    trovato = true;
                                    break;
                                }
                            }
                            //if (!listaBuste.Contains(bustac))
                            if(!trovato)
                                messaggio += "Busta non trovata nei consuntivi: busta " + bustac.idBusta + " batch " + bustac.idBatch + "\r\n";
                        }
                    }
                    else
                    {
                        messaggio += "File consuntivi " + giornoStr + " inesistente.\r\n";
                    }
                }
            }

            return messaggio;   
        }

        private List<busta> CaricaBuste(string path, string tipo)
        {
            List<busta> listaBuste = new List<busta>();
            
            string DirTemp = ConfigurationSettings.AppSettings["DirTemp"];
            if (string.IsNullOrEmpty(DirTemp))
                DirTemp = boAccess.leggiAppSettings("DirTemp");

            string fileDirJob = DirTemp + "\\dirjob.txt";
            Process myProcess = new Process();
            myProcess.StartInfo.FileName = "cmd.exe";
            if(tipo == "dir")
                myProcess.StartInfo.Arguments = "/c dir " + path + "\\job.csv /s/b>" + fileDirJob;
            else if (tipo == "file")
                myProcess.StartInfo.Arguments = "/c dir " + path + " /s/b>" + fileDirJob;
            myProcess.StartInfo.UseShellExecute = false;
            myProcess.StartInfo.ErrorDialog = false;
            myProcess.StartInfo.RedirectStandardOutput = true;
            myProcess.StartInfo.CreateNoWindow = true;
            myProcess.StartInfo.RedirectStandardError = true;
            myProcess.Start();
            string processError = myProcess.StandardError.ReadToEnd();
            myProcess.WaitForExit();

            StreamReader sr = null;
            StreamReader srJob = null;
            try
            {
                sr = File.OpenText(fileDirJob);
                String input = String.Empty;
                while ((input = sr.ReadLine()) != null)
                {
                    string pathfilecsv = input;
                    string[] riga = input.Split('\\');
                    string batch = riga[riga.Length - 2];
                    string giorno = riga[riga.Length - 3];

                    DateTime data = DateTime.Now;

                    srJob = File.OpenText(pathfilecsv);
                    String inputJob = String.Empty;
                    busta busta = new busta();
                    while ((inputJob = srJob.ReadLine()) != null)
                    {
                        string[] rigaJob = inputJob.Split(';');
                        if (rigaJob.Length == 3)
                        {
                        }
                        else if (rigaJob.Length == 6)
                        {
                            busta.idBatch = rigaJob[0];
                            busta.idBusta = rigaJob[1];
                            busta.codArticolo = rigaJob[3];
                            busta.quantita = Convert.ToInt32(rigaJob[5]);
                        }
                        listaBuste.Add(busta);
                    }
                    srJob.Close();
                }

                sr.Close();
            }
            catch (Exception ex)
            {
            }
            finally
            {
                sr.Close();
                srJob.Close();
            }

            return listaBuste;
        }

        private void btInCaricaBuste_ClickOld(object sender, EventArgs e)
        {
            string DirConsuntivi = boAccess.leggiAppSettings("DirConsuntivi");
            foreach (string file in Directory.GetFiles(DirConsuntivi, "*.csv"))
            {
                FileInfo fi = new FileInfo(file);
                string nomeFile = fi.Name;
                string idBusta = string.Empty;
                string idBatch = string.Empty;
                string idArticolo = string.Empty;
                string desArticolo = string.Empty;
                string qta = "";
                DateTime data = DateTime.Now;
                if (!File.Exists(DirConsuntivi + "\\imp\\" + fi.Name))
                {
                    StreamReader sr = File.OpenText(fi.FullName);
                    String input = String.Empty;
                    while ((input = sr.ReadLine()) != null)
                    {
                        string[] riga = input.Split(';');
                        if (riga.Length == 3)
                        {
                            idBusta = riga[0];
                            desArticolo = riga[1];
                            qta = riga[2];
                            data = Convert.ToDateTime(fi.Name.Replace(fi.Extension, ""));
                        }
                        else if (riga.Length == 6)
                        {
                            idBatch = riga[0];
                            idBusta = riga[1];
                            idArticolo = riga[3];
                            desArticolo = riga[4];
                            qta = riga[5];
                            //data = Convert.ToDateTime(riga[2]);
                            data = Convert.ToDateTime(fi.Name.Replace(fi.Extension, ""));
                        }

                        AggiungeDGInBuste(idBatch, idArticolo, desArticolo, idBusta, Convert.ToDateTime(data), Convert.ToInt32(qta));
                        ConsuntiviImport.Items.Add(fi.Name);
                        modificatoIn = true;
                    }
                }
            }

            if (ConsuntiviImport.Items.Count <= 0)
                MessageBox.Show("Non sono stati trovati file da importare.", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btInserisceResi_Click(object sender, EventArgs e)
        {
            string fornitore = boAccess.leggiAppSettings("Fornitore");
            lblErrore.Text = string.Empty;
            string codArticolo = string.Empty;
            int qtaArticolo = 0;
            int secondiDoppiaLettura = 0;
            string valore = boAccess.leggiAppSettings("SecondiDoppiaLettura");
            if (valore != string.Empty)
                secondiDoppiaLettura = Convert.ToInt32(valore);
            else
                secondiDoppiaLettura = 0;

            //sistema la data
            if (DateTime.Now.ToShortDateString() != dtData.Value.ToShortDateString()
                && dgInBuste.Rows.Count == 0)
                dtData.Value = DateTime.Now;

            if (lblArticoloResi.Text != string.Empty && lblBustaResi.Text != string.Empty && !lblKitResi.Visible)
            {
                lblArticoloResi.Text = string.Empty;
                lblDesArticoloResi.Text = string.Empty;
                lblQtaResi.Text = string.Empty;
                lblBustaResi.Text = string.Empty;
            }
            string codice = tbCodiceResi.Text;
            if (codice == string.Empty)
                return;

            TimeSpan duration = (DateTime.Now - ultimoCodiceLettoTime);
            if (codice == ultimoCodiceLetto && duration.Seconds < secondiDoppiaLettura)
            {
                StampaMessaggio("Errore. Doppia lettura codice.", true, TipoSuono.NoSuono);
                return;
            }
            ultimoCodiceLetto = tbInCodice.Text;
            ultimoCodiceLettoTime = DateTime.Now;

            if (codice == string.Empty)
            {
                tbInCodice.Focus();
                return;
            }
            if (codice.ToLower() == "azzerapagina")
            {
                btAzzeraPagina_Click(btAzzeraPaginaIn, null);
                return;
            }
            else if (codice.ToLower() == "eliminaultimo")
            {
                btEliminaUltimo_Click(btInInserisce, null);
                return;
            }
            else if (codice.ToLower() == "iniziokit")
            {
                if (lblBustaResi.Text != string.Empty || lblArticoloResi.Text != string.Empty) //busta piena -> errore
                {
                    StampaMessaggio("Errore. Il codice Kit va letto con pagina vuota.", true, TipoSuono.Errore);
                    return;
                }
                StampaMessaggio("", true, TipoSuono.Kit);
                lblKitResi.Visible = true;
                tbInCodice.Text = string.Empty;
                tbInCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "elultimomov")
            {
                PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                DialogResult dr = MessageBox.Show("Si desidera eliminare l'ultimo movimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                if (dr == DialogResult.Yes)
                {
                    eliminaUltimoMovimento();
                }
                else
                {
                    tbInCodice.Text = string.Empty;
                    tbInCodice.Focus();
                    return;
                }
            }
            else if (codice.ToLower() == "finekit")
            {
                if (!lblKitResi.Visible)
                {
                    StampaMessaggio("Codice finekit senza iniziokit.", true, TipoSuono.Errore);
                    return;
                }
                StampaMessaggio("", true, TipoSuono.Kit);
                lblKitResi.Visible = false;
                tbInCodice.Text = string.Empty;
                tbInCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "salvatemp")
            {
                btSalvaTemp_Click(null, null);
                tbInCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "caricatemp")
            {
                btCaricaTemp_Click(null, null);
                tbInCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "in")
            {
                tbInCodice.Text = string.Empty;
                tbInCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "out")
            {
                tab1.SelectedTab = tabOut;
                tbInCodice.Text = string.Empty;
                tbInCodice.Focus();
                return;
            }
            else if (codice.ToLower() == "fine")
            {
                TBFineIn_Click(tbInCodice, null);
                return;
            }
            else if (codice.ToLower() == "aggscura" &&
                ((lblArticoloResi.Text == string.Empty && lblBustaResi.Text == string.Empty) ||
                lblKitResi.Visible))
            {
                codArticolo = boAccess.leggiAppSettings("CodStampaScura");
                articolo articolo = boAccess.LeggiArticolo(codArticolo);
                lblArticoloResi.Text = codArticolo;
                lblDesArticoloResi.Text = articolo.descrizione;
                lblQtaResi.Text = ultimaQtaLetta;
                lblBustaResi.Text = ultimaBustaLetta;
            }
            else if (codice.Substring(codice.Length - 1, 1) == "*")
            {
                try
                {
                    qtaArticolo = Convert.ToInt32(codice.Substring(0, codice.Length - 1));
                }
                catch
                {
                    StampaMessaggio("Errore. Quantità non valida.", true, TipoSuono.Errore);
                    return;
                }
                return;
            }
            else if (fornitore.ToLower() == "photosi" && (codice.Length == 11 || codice.Length == 4 || codice.IndexOf("*") > 0) ||
                    fornitore.ToLower() == "fotoevolution" && (codice.Length == 12 && codice.Substring(0, 5) != "00000") || (codice.Length == 6 && lblBustaOut.Text != string.Empty) || codice.IndexOf("*") > 0) //codice articolo
            {
                if (lblArticoloIn.Text != string.Empty && !lblKitIn.Visible) //articolo pieno -> errore
                {
                    StampaMessaggio("Errore. Lettura doppia codice articolo.", true, TipoSuono.Errore);
                    return;
                }
                if (lblKitIn.Visible && lblBustaIn.Text == string.Empty)
                {
                    StampaMessaggio("Errore. Con il Kit leggere per primo il codice busta.", true, TipoSuono.Errore);
                    return;
                }

                int posX = codice.IndexOf("*");
                if (posX == 4)
                {
                    codArticolo = codice.Substring(0, 4);
                    try
                    {
                        qtaArticolo = Convert.ToInt32(codice.Substring(5));
                    }
                    catch
                    {
                        qtaArticolo = 1;
                    }
                }
                if (posX < 4 && posX > 0)
                {
                    codArticolo = codice.Substring(posX + 1, 4);
                    try
                    {
                        qtaArticolo = Convert.ToInt32(codice.Substring(0, posX));
                    }
                    catch
                    {
                        qtaArticolo = 1;
                    }
                }
                else if (fornitore.ToLower() == "photosi" && codice.Length == 11)
                {
                    codArticolo = codice.Substring(5, 4);
                    try
                    {
                        qtaArticolo = Convert.ToInt32(codice.Substring(0, 5));
                    }
                    catch { return; }
                }
                else if (fornitore.ToLower() == "photosi" && codice.Length == 4)
                {
                    codArticolo = codice;
                    qtaArticolo = 1;
                }
                else if (fornitore.ToLower() == "fotoevolution" && codice.Length == 12)
                {
                    codArticolo = codice.Substring(0, 6);
                    try
                    {
                        qtaArticolo = Convert.ToInt32(codice.Substring(6, 5));
                    }
                    catch { return; }
                }
                else if (fornitore.ToLower() == "fotoevolution" && codice.Length == 6)
                {
                    codArticolo = codice;
                    qtaArticolo = 1;
                }

                if (qtaArticolo < 1)
                {
                    StampaMessaggio("Errore. Quantita a zero.", true, TipoSuono.Errore);
                    return;
                }

                articolo articolo = boAccess.LeggiArticolo(codArticolo);
                if (articolo.codice == string.Empty)
                {
                    StampaMessaggio("Errore. Codice articolo " + codArticolo + " inesistente.", true, TipoSuono.Errore);
                    return;
                }
                if (boAccess.leggiAppSettings("MinQtaMessaggio").ToString() != string.Empty && qtaArticolo > Convert.ToInt32(boAccess.leggiAppSettings("MinQtaMessaggio")))
                {
                    PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                    DialogResult dr = MessageBox.Show("Attenzione, quantità = " + qtaArticolo.ToString() + ". Confermi?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                    if (dr == DialogResult.No)
                    {
                        tbInCodice.Text = string.Empty;
                        return;
                    }
                }
                lblArticoloResi.Text = codArticolo;
                lblDesArticoloResi.Text = articolo.descrizione;
                lblQtaResi.Text = qtaArticolo.ToString();

                ultimoArticoloLetto = codArticolo;
                ultimaQtaLetta = qtaArticolo.ToString();

                if (lblBustaResi.Text == string.Empty) //articolo vuoto e busta vuota
                {
                    tbCodiceResi.Text = string.Empty;
                    tbCodiceResi.Focus();
                    return;
                }
            }
            else if (isCodiceBusta(codice)) //codice busta
            {
                if (fornitore.ToLower() == "fotoevolution" && codice.Length > 6)
                    codice = Convert.ToString(Convert.ToInt32(codice.Substring(0, 11)));
                if (boAccess.leggiAppSettings("VerificaBusteEntrata") == "1" &&
                    lblArticoloResi.Text != boAccess.leggiAppSettings("BustaFittizia") &&
                    lblArticoloResi.Text != string.Empty)
                {
                    if (!controlloCoerenzaBuste(codice, lblArticoloResi.Text, Convert.ToInt32(lblQtaResi.Text), "Resi"))
                    {
                        tbCodiceResi.Text = "";
                        return;
                    }
                }
                if (!inserisceCodiceBustaResi(codice))
                    return;
            }
            else
            {
                //StampaMessaggio("Codice non riconosciuto. E' un codice busta?", false, TipoSuono.Domanda);
                //lCodErrore.Text = "1";
                //tbErrore.Enabled = true;
                //this.AcceptButton = btErrore;
                //tbErrore.Text = string.Empty;
                //tbErrore.Select();
                //tbErrore.Focus();
                //return;
                PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                DialogResult dr = MessageBox.Show("Codice busta non riconosciuto. Confermi ugualmente?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.No)
                {
                    tbInCodice.Text = string.Empty;
                    return;
                }
                else
                {
                    lblBustaResi.Text = codice;
                    ultimaBustaLetta = codice;
                }
            }

            if (lblQtaResi.Text == string.Empty)
                qtaArticolo = 0;
            else
                qtaArticolo = Convert.ToInt32(lblQtaResi.Text);

            codArticolo = lblArticoloResi.Text;
            string codBusta = lblBustaResi.Text;
            if (qtaArticolo > 0)
            {
                AggiungeDGResiBuste("", codArticolo, lblDesArticoloResi.Text, codBusta, dtData.Value, qtaArticolo);
                StampaMessaggio("", true, TipoSuono.OkInserimento);
                modificatoIn = true;
            }
            tbCodiceResi.Text = string.Empty;
            tbCodiceResi.Focus();

        }

        private void btInserisceResiOld_Click(object sender, EventArgs e)
        {
            bool ok = true;

            string codice = tbCodiceResi.Text;
            if (codice == string.Empty)
            {
                tbCodiceResi.Focus();
                return;
            }

            if (!isCodiceBusta(codice)) //codice busta
            {
                PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                DialogResult dr = MessageBox.Show("Codice busta non riconosciuto. Confermi ugualmente?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.No)
                {
                    ok = false;
                }
            }
            else
            {
                if (boAccess.leggiAppSettings("VerificaBusteEntrata") == "1" &&
                    !boAccess.isBustaEsistenteDBIn(codice))
                {
                    PlaySound(suonoConferma, IntPtr.Zero, SoundFlags.SND_FILENAME | SoundFlags.SND_ASYNC);
                    DialogResult dr = MessageBox.Show("Codice busta non entrata. Confermi ugualmente l'inserimento?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2);
                    if (dr == DialogResult.No)
                    {
                        ok = false;
                    }
                }
            }

            if (ok)
            {
                lblBustaResi.Text = tbCodiceResi.Text;
                AggiungeDGResiBuste("", "", "", tbCodiceResi.Text, dtData.Value, 0);
                StampaMessaggio("", true, TipoSuono.OkInserimento);
                modificatoResi = true;
            }

            tbCodiceResi.Text = string.Empty;
            tbCodiceResi.Focus();
        }

        private void tbFineResi_Click(object sender, EventArgs e)
        {
            if (dgResiBuste.Rows.Count == 0)
            {
                StampaMessaggio("Non ci sono buste da salvare.", true, TipoSuono.Errore);
                return;
            }
            bool ret = boAccess.ScriveMovimentiResi(dgResiBuste);
            string nomeFile = string.Empty;
            if (ret)
            {
                //DialogResult dr = MessageBox.Show("Scrittura eseguita con successo. Vuoi generare i file?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                //if (dr == DialogResult.Yes)
                //{
                if (!GeneraFileBusteResi(boAccess.leggiAppSettings("DirFileBolla")))
                    MessageBox.Show("Errore in fase di scrittura", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //}
                AzzeraResiBuste();
                btAzzeraPagina_Click(btAzzeraPaginaIn, null);
                modificatoResi = false;
                MessageBox.Show("Scrittura eseguita con successo.", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Errore in fase di scrittura", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsmAnalizzaBusta_Click(object sender, EventArgs e)
        {
            Form formBuste = this;
            fmAnalizzaBusta fmBusta = new fmAnalizzaBusta(ref formBuste);
            fmBusta.frmBuste = this;
            fmBusta.Show();
        }

        private void tsmAnalizzaBatch_Click(object sender, EventArgs e)
        {
            fmAnalizzaBatch fmBatch = new fmAnalizzaBatch();
            fmBatch.Show();
        }

        static void WalkDirectoryTree(System.IO.DirectoryInfo root)
        {
            System.IO.FileInfo[] files = null;
            System.IO.DirectoryInfo[] subDirs = null;

            // First, process all the files directly under this folder
            try
            {
                files = root.GetFiles("*.*");
            }
            // This is thrown if even one of the files requires permissions greater
            // than the application provides.
            catch (UnauthorizedAccessException e)
            {
                // This code just writes out the message and continues to recurse.
                // You may decide to do something different here. For example, you
                // can try to elevate your privileges and access the file again.
                Console.WriteLine(e.Message);
            }

            catch (System.IO.DirectoryNotFoundException e)
            {
                Console.WriteLine(e.Message);
            }

            if (files != null)
            {
                foreach (System.IO.FileInfo fi in files)
                {
                    // In this example, we only access the existing FileInfo object. If we
                    // want to open, delete or modify the file, then
                    // a try-catch block is required here to handle the case
                    // where the file has been deleted since the call to TraverseTree().
                    Console.WriteLine(fi.FullName);
                }

                // Now find all the subdirectories under this directory.
                subDirs = root.GetDirectories();

                foreach (System.IO.DirectoryInfo dirInfo in subDirs)
                {
                    // Resursive call for each subdirectory.
                    WalkDirectoryTree(dirInfo);
                }
            }
        }

        private void timerCaricaBuste_Tick(object sender, EventArgs e)
        {
            if (tab1.SelectedTab != tabIn && dgInBuste.Rows.Count == 0)
            {
                int numbuste = CaricaBuste(timerCaricaBuste);
                if(numbuste > 0)
                    StampaMessaggio("Buste in entrata aggiornate.", false, TipoSuono.NoSuono);
            }
        }

        private int CaricaBuste(object sender)
        {
            btInCaricaBuste_Click(sender, null);
            int numBuste = dgInBuste.Rows.Count;
            if (numBuste > 0)
                TBFineIn_Click(timerCaricaBuste, null);
            return numBuste;
        }

        private void tsmBusteNonUscite_Click(object sender, EventArgs e)
        {
            Form formBuste = this;
            fmBusteNonUscite fmBusta = new fmBusteNonUscite(ref formBuste);
            fmBusta.frmBuste = this;
            fmBusta.Show();
        }

        private void tsmConteggiaArticolo_Click(object sender, EventArgs e)
        {
            Form formBuste = this;
            fmConteggiaArticolo form = new fmConteggiaArticolo(ref formBuste);
            form.frmBuste = this;
            form.Show();
        }

        private void ckRifacimento_CheckedChanged(object sender, EventArgs e)
        {
            tbInCodice.Text = string.Empty;
            tbInCodice.Focus();
        }

        private void btRicalcolaTotali_Click(object sender, EventArgs e)
        {
            RicalcolaTotali();
        }

        private void leggiBusteInEntrataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tab1.SelectedTab != tabIn && dgInBuste.Rows.Count == 0)
            {
                int numbuste = CaricaBuste(tab1);
                if (numbuste > 0)
                {
                    MessageBox.Show("Sono state inserite " + numbuste.ToString() + " buste in entrata.");
                }
                else
                {
                    MessageBox.Show("Non sono state trovate nuove buste in entrata.");
                }
            }
        }

        private void totaliDelGiornoToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void tsmRipristinaGiorno_Click(object sender, EventArgs e)
        {
            frmRipristinaGiorno form = new frmRipristinaGiorno();
            form.Show();
        }

        private void totakiMonoprodottoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmMonoprodotto form = new frmMonoprodotto();
            form.Show();
        }

        private void caricaBusteMultiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //frmCaricaBusteMulti form = new frmCaricaBusteMulti();
            //form.Show();
            CaricaBusteMulti();
        }

        private void CaricaBusteMulti()
        {
            string elenco = "";
            foreach (string item in elenco.Split(';'))
            {
                string busta = item.Split(',')[0];
                string codice = item.Split(',')[1];
                string qta = "";
                try
                {
                    qta = item.Split(',')[2];
                }
                catch { }

                tbOutCodice.Text = busta;
                btOutInserisce_Click(null, null);
                tbOutCodice.Text = codice + (!string.IsNullOrEmpty(qta) ? "*" + qta : "");
                btOutInserisce_Click(null, null);
            }
        }

    }

}