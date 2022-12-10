using System;
using System.Collections.Generic;
using System.Configuration;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;


namespace Buste
{
    [Serializable()]
    public class Access
    {
        protected OleDbConnection aConnection;
        protected OleDbCommand aCommand;
        //private OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");

        public Access()
        {
            string database = string.Empty;
            try
            {
                //string dirdatabase = ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb";
                database = ConfigurationSettings.AppSettings["PathDatabase"].ToString();
                //database = "\\\\server7\\buste\\Buste\\Buste.mdb";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            string provider = "Microsoft.ACE.OLEDB.12.0";
            provider = "Microsoft.Jet.OLEDB.4.0";
            string connString = "Provider=" + provider + ";Data Source=" + database + ";Persist Security Info=False;";
            aConnection = new OleDbConnection(connString);
            aConnection.Open();
        }

        public void ChiudiConnessione()
        {
            if (aConnection.State == ConnectionState.Open)
                aConnection.Close();
        }

        public DataTable LeggiCodici()
        {
            //create the database connection
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");

            //create the command object and store the sql query
            aCommand = new OleDbCommand("select * from Articoli", aConnection);
            try
            {
                DataTable dtCodici = new DataTable();
                dtCodici.Columns.Add("idArticolo");
                dtCodici.Columns.Add("Descrizione");
                dtCodici.Columns.Add("qta");
                DataRow dr;

                aConnection.Open();

                //create the datareader object to connect to table
                OleDbDataReader aReader = aCommand.ExecuteReader();

                //Iterate throuth the database
                while(aReader.Read())
                {
                    dr = dtCodici.NewRow();
                    dr["idArticolo"] = aReader["idArticolo"].ToString();
                    dr["descrizione"] = aReader["Descrizione"].ToString();
                    dr["qta"] = "0";
                    dtCodici.Rows.Add(dr);
                }

                //close the reader 
                aReader.Close();

                //close the connection Its important.
                //aConnection.Close();

                return dtCodici;
            }

            //Some usual exception handling
            catch(OleDbException e)
            {
                //aConnection.Close();
                Console.WriteLine("Error: {0}", e.Errors[0].Message);
                return null;
            }
        }

        public articolo LeggiArticolo(string codArticolo)
        {
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");
            //aConnection.Open();

            string query = "SELECT * FROM Articoli WHERE idArticolo=@idArticolo";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@idArticolo", codArticolo);

            articolo starticolo;
            starticolo.codice = string.Empty;
            starticolo.descrizione = string.Empty;
            starticolo.codiceInt = string.Empty;
            starticolo.codiceIntNum = string.Empty;
            starticolo.codiceFatturazione = string.Empty;
            starticolo.scaglione1 = 0;
            starticolo.sconto1 = 0;
            starticolo.scaglione2 = 0;
            starticolo.sconto2 = 0;
            starticolo.scaglione3 = 0;
            starticolo.sconto3 = 0;
            starticolo.scaglione4 = 0;
            starticolo.sconto4 = 0;
            starticolo.scaglione5 = 0;
            starticolo.sconto5 = 0;
            starticolo.omaggio = 0;

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                if (aReader.HasRows)
                {
                    aReader.Read();
                    starticolo.codice = aReader["idArticolo"].ToString();
                    starticolo.descrizione = aReader["Descrizione"].ToString();
                    starticolo.codiceInt = aReader["CodiceInterno"].ToString();
                    starticolo.codiceIntNum = aReader["CodiceInternoNum"].ToString();
                    starticolo.codiceFatturazione = aReader["CodiceFatturazione"].ToString();
                    if (aReader["Scaglione1"] != DBNull.Value)
                        starticolo.scaglione1 = Convert.ToInt32(aReader["Scaglione1"].ToString());
                    if (aReader["Sconto1"] != DBNull.Value)
                        starticolo.sconto1 = Convert.ToDecimal(aReader["Sconto1"].ToString());
                    if (aReader["Scaglione2"] != DBNull.Value)
                        starticolo.scaglione2 = Convert.ToInt32(aReader["Scaglione2"].ToString());
                    if (aReader["Sconto2"] != DBNull.Value)
                        starticolo.sconto2 = Convert.ToDecimal(aReader["Sconto2"].ToString());
                    if (aReader["Scaglione3"] != DBNull.Value)
                        starticolo.scaglione3 = Convert.ToInt32(aReader["Scaglione3"].ToString());
                    if (aReader["Sconto3"] != DBNull.Value)
                        starticolo.sconto3 = Convert.ToDecimal(aReader["Sconto3"].ToString());
                    if (aReader["Scaglione4"] != DBNull.Value)
                        starticolo.scaglione4 = Convert.ToInt32(aReader["Scaglione4"].ToString());
                    if (aReader["Sconto4"] != DBNull.Value)
                        starticolo.sconto4 = Convert.ToDecimal(aReader["Sconto4"].ToString());
                    if (aReader["Scaglione5"] != DBNull.Value)
                        starticolo.scaglione5 = Convert.ToInt32(aReader["Scaglione5"].ToString());
                    if (aReader["Sconto5"] != DBNull.Value)
                        starticolo.sconto5 = Convert.ToDecimal(aReader["Sconto5"].ToString());
                    if (aReader["Omaggio"] != DBNull.Value)
                        starticolo.omaggio = Convert.ToInt32(aReader["Omaggio"].ToString());
                }
            }
            catch (OleDbException e)
            {
                starticolo.codice = string.Empty;
            }

            //aConnection.Close();

            return starticolo;

        }

        public DataTable ListaArticoli()
        {
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");
            //aConnection.Open();

            string query = "SELECT * FROM Articoli";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);

            DataTable dt = new DataTable();

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                dt.Load(aReader);
                aReader.Close();
            }
            catch (OleDbException e)
            {
                return null;
            }

            //aConnection.Close();

            return dt;

        }

        public DataTable LeggiBatch(DateTime data)
        {
            //create the database connection
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");
            //aConnection.Open();

            //create the command object and store the sql query
            OleDbCommand aCommand = new OleDbCommand("select * from Batch where data = @data", aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@data", data);
            
            try
            {
                DataTable dtBatch = new DataTable();
                dtBatch.Columns.Add("idArticolo");
                dtBatch.Columns.Add("idBusta");
                dtBatch.Columns.Add("qta");
                DataRow dr;


                //create the datareader object to connect to table
                OleDbDataReader aReader = aCommand.ExecuteReader();

                //Iterate throuth the database
                while (aReader.Read())
                {
                    dr = dtBatch.NewRow();
                    dr["idArticolo"] = aReader["idArticolo"].ToString();
                    dr["idBusta"] = aReader["idBusta"].ToString();
                    dr["batch"] = aReader["batch"].ToString();
                    dr["data"] = aReader["data"].ToString();
                    dr["qta"] = aReader["qta"].ToString();
                    dtBatch.Rows.Add(dr);
                }

                //close the reader 
                aReader.Close();

                //close the connection Its important.
                //aConnection.Close();

                return dtBatch;
            }

            //Some usual exception handling
            catch (OleDbException e)
            {
                //aConnection.Close();
                Console.WriteLine("Error: {0}", e.Errors[0].Message);
                return null;
            }
        }

        public bool ScriveMovimentiOut(DataGridView dgMovimenti)
        {
            //create the database connection
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");
            //aConnection.Open();
            //OleDbTransaction tx = aConnection.BeginTransaction();

            //create the command object and store the sql query
            foreach (DataGridViewRow dr in dgMovimenti.Rows)
            {
                int qta = Convert.ToInt32(dr.Cells["qta"].Value);
                if (qta > 0)
                {
                    try
                    {
                        string query = "INSERT INTO MovimentiOut ([idArticolo], [idBusta], [data], [numBolla], [qta], [note], [sconto], [idBatch], [tipoOrdine], [numProdotti]) " +
                                    "VALUES (@idArticolo,@idBusta,@data,@numBolla,@qta,@note,@sconto,@idBatch,@tipoOrdine,@numProdotti)";

                        OleDbCommand aCommand = new OleDbCommand(query, aConnection);
                        //aCommand.Transaction = tx;
                        aCommand.Parameters.Clear();
                        aCommand.Parameters.Add("@idArticolo", dr.Cells["idArticolo"].Value);
                        aCommand.Parameters.Add("@idBusta", dr.Cells["idBusta"].Value);
                        aCommand.Parameters.Add("@data", dr.Cells["data"].Value);
                        aCommand.Parameters.Add("@numBolla", dr.Cells["numBolla"].Value);
                        aCommand.Parameters.Add("@qta", dr.Cells["qta"].Value);
                        aCommand.Parameters.Add("@note", "");
                        aCommand.Parameters.Add("@sconto", dr.Cells["sconto"].Value);
                        aCommand.Parameters.Add("@idBatch", dr.Cells["batch"].Value);
                        aCommand.Parameters.Add("@tipoOrdine", dr.Cells["tipoOrdine"].Value);
                        aCommand.Parameters.Add("@numProdotti", dr.Cells["numProdotti"].Value);
                        aCommand.ExecuteNonQuery();

                        //query = "SELECT * FROM Buste WHERE idBusta=@idBusta";
                        //OleDbCommand aCommand2 = new OleDbCommand(query, aConnection);
                        ////aCommand2.Transaction = tx;
                        //aCommand2.Parameters.Add("@idBusta", dr.Cells["idBusta"].Value);
                        //OleDbDataReader aReader = aCommand2.ExecuteReader();
                        //if (aReader.HasRows)
                        //{
                        //    query = "UPDATE Buste SET dataOut='" + dr.Cells["data"].Value + "' WHERE idBusta=@idBusta";
                        //    //query = "UPDATE Buste SET dataOut='03/11/2008' WHERE idBusta='506488681'";
                        //    //query = "UPDATE Buste SET dataOut=@dataOut WHERE idBusta=@idBusta";
                        //    OleDbCommand aCommand3 = new OleDbCommand(query, aConnection);
                        //    //aCommand3.Transaction = tx;
                        //    aCommand3.Parameters.Add("@idBusta", dr.Cells["idBusta"].Value);
                        //    aCommand3.Parameters.Add("@dataOut", Convert.ToDateTime(dr.Cells["data"].Value));
                        //    aCommand3.ExecuteNonQuery();
                        //}
                        //else
                        //{
                        //    query = "INSERT INTO Buste ([idBusta], [dataOut]) " +
                        //                "VALUES (@idBusta,@dataOut)";
                        //    OleDbCommand aCommand4 = new OleDbCommand(query, aConnection);
                        //    //aCommand4.Transaction = tx;
                        //    aCommand4.Parameters.Add("@idBusta", dr.Cells["idBusta"].Value);
                        //    aCommand4.Parameters.Add("@dataOut", dr.Cells["data"].Value);
                        //    aCommand4.Parameters.Add("@dataIn", DBNull.Value);
                        //    aCommand4.Parameters.Add("@idBatch", "");
                        //    aCommand4.ExecuteNonQuery();
                        //}

                    }

                    //Some usual exception handling
                    catch (OleDbException e)
                    {
                        //tx.Rollback();
                        //aConnection.Close();
                        Console.WriteLine("Error: {0}", e.Errors[0].Message);
                        return false;
                    }
                }
            }
            //tx.Commit();
            //aConnection.Close();

            return true;

        }

        public bool ScriveMovimentoIn(movimento mov, bool verificaEsistenza)
        {
            bool ret = true;

            try
            {
                bool inserisce = true;
                if (verificaEsistenza && isBustaEsistenteDBIn(mov))
                    inserisce = false;
                
                if (inserisce)
                {
                    string query = "INSERT INTO MovimentiIn ([idBusta], [idBatch], [idArticolo], [descrizione], [data], [qta], [tipoOrdine], [numProdotti]) " +
                                "VALUES (@idBusta,@idBatch,@idArticolo,@desArticolo,@data,@qta,@tipoOrdine,@numProdotti)";

                    OleDbCommand aCommand = new OleDbCommand(query, aConnection);
                    //aCommand.Transaction = tx;
                    aCommand.Parameters.Clear();
                    aCommand.Parameters.Add("@idBusta", mov.idBusta);
                    aCommand.Parameters.Add("@idBatch", mov.idBatch);
                    aCommand.Parameters.Add("@idArticolo", mov.idArticolo);
                    aCommand.Parameters.Add("@desArticolo", mov.desArticolo);
                    aCommand.Parameters.Add("@data", mov.data);
                    aCommand.Parameters.Add("@qta", mov.qta);
                    aCommand.Parameters.Add("@tipoOrdine", mov.tipoOrdine);
                    aCommand.Parameters.Add("@numProdotti", mov.numProdotti);
                    aCommand.ExecuteNonQuery();
                }
            }
            catch (OleDbException e)
            {
                //aConnection.Close();
                Console.WriteLine("Error: {0}", e.Errors[0].Message);
                ret = false;
            }

            return ret;
        }

        public bool ScriveMovimentiIn(DataGridView dgBuste)
        {
            //create the database connection
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");
            //aConnection.Open();
            //OleDbTransaction tx = aConnection.BeginTransaction();

            //create the command object and store the sql query
            foreach (DataGridViewRow dr in dgBuste.Rows)
            {
                try
                {
                    movimento mov = new movimento();
                    mov.idBusta = dr.Cells["idBusta"].Value.ToString();
                    mov.idBatch = dr.Cells["batch"].Value.ToString();
                    mov.idArticolo = dr.Cells["idArticolo"].Value.ToString();
                    mov.desArticolo = dr.Cells["descrizione"].Value.ToString();
                    mov.data = Convert.ToDateTime(dr.Cells["data"].Value);
                    mov.qta = Convert.ToInt32(dr.Cells["qta"].Value);
                    ScriveMovimentoIn(mov, true);
                }

                //Some usual exception handling
                catch (OleDbException e)
                {
                    //tx.Rollback();
                    //aConnection.Close();
                    Console.WriteLine("Error: {0}", e.Errors[0].Message);
                    return false;
                }
            }
            //tx.Commit();
            //aConnection.Close();

            return true;

        }

        public bool ScriveMovimentiResi(DataGridView dataGrid)
        {
            //create the database connection
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");
            //aConnection.Open();
            //OleDbTransaction tx = aConnection.BeginTransaction();

            //create the command object and store the sql query
            foreach (DataGridViewRow dr in dataGrid.Rows)
            {
                try
                {
                    string query = "INSERT INTO MovimentiResi ([idBusta], [idBatch], [idArticolo], [descrizione], [data], [qta]) " +
                                "VALUES (@idBusta,@idBatch,@idArticolo,@desArticolo,@data,@qta)";

                    OleDbCommand aCommand = new OleDbCommand(query, aConnection);
                    //aCommand.Transaction = tx;
                    aCommand.Parameters.Clear();
                    aCommand.Parameters.Add("@idBusta", dr.Cells["idBusta"].Value);
                    aCommand.Parameters.Add("@idBatch", dr.Cells["batch"].Value);
                    aCommand.Parameters.Add("@idArticolo", dr.Cells["idArticolo"].Value);
                    aCommand.Parameters.Add("@desArticolo", dr.Cells["descrizione"].Value);
                    aCommand.Parameters.Add("@data", dr.Cells["data"].Value);
                    aCommand.Parameters.Add("@qta", dr.Cells["qta"].Value);
                    aCommand.ExecuteNonQuery();

                }

                //Some usual exception handling
                catch (OleDbException e)
                {
                    //tx.Rollback();
                    //aConnection.Close();
                    Console.WriteLine("Error: {0}", e.Errors[0].Message);
                    return false;
                }
            }
            //tx.Commit();
            //aConnection.Close();

            return true;

        }

        public bool isBustaEsistenteDBOut(string idBusta)
        {
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");
            //aConnection.Open();

            string query = "SELECT * FROM MovimentiOut WHERE idBusta=@idBusta";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@idBusta", idBusta);

            bool ret = false;

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                if (aReader.HasRows)
                    ret = true;
            }
            catch (OleDbException e)
            {
                ret = false;
            }

            //aConnection.Close();

            return ret;
        }

        public bool isBustaEsistenteDBIn(string idBusta)
        {
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");
            //aConnection.Open();

            string query = "SELECT * FROM MovimentiIn WHERE idBusta=@idBusta";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@idBusta", idBusta);

            bool ret = false;

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                if (aReader.HasRows)
                    ret = true;
            }
            catch (OleDbException e)
            {
                ret = false;
            }

            //aConnection.Close();

            return ret;
        }

        public bool isBustaEsistenteDBIn(movimento mov)
        {
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");
            //aConnection.Open();

            string query = "SELECT * FROM MovimentiIn WHERE idBusta=@idBusta AND idBatch=@idBatch AND idArticolo=@idArticolo AND data=@data AND qta=@qta";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@idBusta", mov.idBusta);
            aCommand.Parameters.Add("@idBatch", mov.idBatch);
            aCommand.Parameters.Add("@idArticolo", mov.idArticolo);
            aCommand.Parameters.Add("@data", mov.data);
            aCommand.Parameters.Add("@qta", mov.qta);

            bool ret = false;

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                if (aReader.HasRows)
                    ret = true;
            }
            catch (OleDbException e)
            {
                ret = false;
            }

            //aConnection.Close();

            return ret;
        }

        public bool isBustaEsistenteDBResi(string idBusta)
        {
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");
            //aConnection.Open();

            string query = "SELECT * FROM MovimentiResi WHERE idBusta=@idBusta";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@idBusta", idBusta);

            bool ret = false;

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                if (aReader.HasRows)
                    ret = true;
            }
            catch (OleDbException e)
            {
                ret = false;
            }

            //aConnection.Close();

            return ret;
        }

        public busta leggiBustaDB(string idBusta, string InOut)
        {
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");
            //aConnection.Open();

            string query = string.Empty;
            if(InOut.ToLower() == "out")
                query = "SELECT * FROM MovimentiOut WHERE idBusta=@idBusta";
            else if (InOut.ToLower() == "in")
                query = "SELECT * FROM MovimentiIn WHERE idBusta=@idBusta";
            else if (InOut.ToLower() == "resi")
                query = "SELECT * FROM MovimentiResi WHERE idBusta=@idBusta";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@idBusta", idBusta);

            busta stbusta = new busta();
            stbusta.idBusta = "0";
            stbusta.idBatch = "";
            stbusta.data = DateTime.Today;

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                if (aReader.HasRows)
                {
                    aReader.Read();
                    stbusta.idBusta = idBusta;
                    stbusta.idBatch = aReader["idBatch"].ToString();
                    stbusta.data = Convert.ToDateTime(aReader["data"]);
                    stbusta.codArticolo = aReader["idArticolo"].ToString();
                    stbusta.quantita = Convert.ToInt32(aReader["qta"]);
                    stbusta.tipoOrdine = aReader["tipoOrdine"].ToString();
                    stbusta.numProdotti = aReader["numProdotti"].ToString();
                }
            }
            catch (OleDbException e)
            {
                stbusta.idBusta = "0";
            }

            //aConnection.Close();

            return stbusta;
        }

        public bool coerenzaBusta(string idBusta, string idArticolo, int qta)
        {
            bool ret = false;

            string query = string.Empty;
            query = "SELECT * FROM MovimentiIn WHERE idBusta=@idBusta AND idArticolo=@idArticolo AND qta=@qta ";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@idBusta", idBusta);
            aCommand.Parameters.Add("@idArticolo", idArticolo);
            aCommand.Parameters.Add("@qta", qta);

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                if (aReader.HasRows)
                    ret = true;
                
                aReader.Close();
                aReader.Dispose();
            }
            catch (OleDbException e)
            {
                ret = false;
            }

            //aConnection.Close();

            return ret;
        }

        public string leggiAppSettings(string chiave)
        {
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");
            //aConnection.Open();

            string query = "SELECT * FROM AppSettings WHERE chiave=@chiave";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@chiave", chiave);

            string valore = "";

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                if (aReader.HasRows)
                {
                    aReader.Read();
                    valore = aReader["valore"].ToString();
                }
                aReader.Close();
                aReader.Dispose();
            }
            catch (OleDbException e)
            {
                valore = "";
            }

            aCommand.Dispose();

            return valore;
        }

        public int scriviAppSettings(string chiave, string valore)
        {
            //OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ConfigurationSettings.AppSettings["DirDatabase"].ToString() + "\\Buste.mdb");
            //aConnection.Open();

            int ret = 0;

            try
            {
                string query = "UPDATE AppSettings SET valore='" + valore + "' WHERE chiave='" + chiave +"'";
                OleDbCommand aCommand = new OleDbCommand(query, aConnection);
                aCommand.Parameters.Clear();
                aCommand.Parameters.Add("@chiave", chiave);
                aCommand.Parameters.Add("@valore", valore);
                ret = aCommand.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                ret = 0;
            }

            //aConnection.Close();

            return ret;
        }

        public DataTable ListaMovimentiBusta(string idBusta, string tipo, bool isMascherina = false)
        {
            string strNumBolla;
            if (tipo == "Out") strNumBolla = "numBolla";
            else strNumBolla = "'' as numBolla";

            string query = "SELECT mov.idArticolo, Articoli.descrizione as descrArticolo, data, " + strNumBolla + ", qta, idBatch";
            query += tipo == "In" || tipo == "Out" ? ", mov.tipoOrdine, mov.numProdotti" : "";
            query += " FROM Movimenti" + tipo + " as mov";
            query += " LEFT OUTER JOIN Articoli on mov.idArticolo = Articoli.idArticolo";
            if(!isMascherina)
                query += " WHERE idBusta=@idBusta";
            else
                query += " WHERE idBatch=@idBusta";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@idBusta", idBusta);

            DataTable dt = new DataTable();

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                dt.Load(aReader);
                aReader.Close();
            }
            catch (OleDbException e)
            {
                return null;
            }
            return dt;
        }

        public DataTable ListaMovimentiArticolo(string idBusta, string idArticolo, string tipo, DateTime ?daData)
        {
            string strNumBolla;
            if (tipo == "Out") strNumBolla = "numBolla";
            else strNumBolla = "'' as numBolla";

            string query = "SELECT Movimenti" + tipo + ".idArticolo, Articoli.descrizione as descrArticolo, data, " + strNumBolla + ", qta, idBatch ";
            query += " FROM Movimenti" + tipo;
            query += " LEFT OUTER JOIN Articoli on Movimenti" + tipo + ".idArticolo = Articoli.idArticolo";
            query += " WHERE Movimenti" + tipo + ".idBusta=@idBusta AND Movimenti" + tipo + ".idArticolo=@idArticolo";
            if (daData != null)
                query += " AND Movimenti" + tipo + ".data>=@data";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@idBusta", idBusta);
            aCommand.Parameters.Add("@idArticolo", idArticolo);
            if (daData != null)
                aCommand.Parameters.Add("@data", daData);

            DataTable dt = new DataTable();

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                dt.Load(aReader);
                aReader.Close();
            }
            catch (OleDbException e)
            {
                return null;
            }
            return dt;
        }

        public DataTable ListaMovimentiBatch(string idBatch)
        {

            string query = "SELECT MovimentiIn.idBusta, MovimentiIn.idArticolo, Articoli.Descrizione, MovimentiIn.qta, MovimentiIn.data AS dataIn, MovimentiOut.data AS dataOut, MovimentiIn.tipoOrdine, MovimentiIn.numProdotti ";
            query += "FROM (MovimentiIn LEFT JOIN Articoli ON MovimentiIn.idArticolo = Articoli.idArticolo) LEFT JOIN MovimentiOut ON MovimentiIn.idBusta = MovimentiOut.idBusta ";
            query += "WHERE MovimentiIn.idBatch=@idBatch";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@idBatch", idBatch);

            DataTable dt = new DataTable();

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                dt.Load(aReader);
                aReader.Close();
            }
            catch (OleDbException e)
            {
                dt = null;
            }
            return dt;
        }

        public DataTable ListaMovimentiBatchIn(string idBatch)
        {
            string query = "SELECT idBusta, idArticolo, qta, data, idBatch";
            query += " FROM MovimentiIn";
            query += " WHERE idBatch=@idBatch";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@idBatch", idBatch);

            DataTable dt = new DataTable();

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                dt.Load(aReader);
                aReader.Close();
            }
            catch (OleDbException e)
            {
                return null;
            }
            return dt;
        }

        public DataTable ListaMovimentiBatchOut(string idBatch)
        {
            string query = "SELECT idBusta, idArticolo, qta, data, idBatch";
            query += " FROM MovimentiOut";
            query += " WHERE idBatch=@idBatch";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@idBatch", idBatch);

            DataTable dt = new DataTable();

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                dt.Load(aReader);
                aReader.Close();
            }
            catch (OleDbException e)
            {
                return null;
            }
            return dt;
        }

        public DataTable ListaBusteNonUscite(DateTime? daData)
        {

            string query = "SELECT * ";
            query += " FROM BusteNonUscite ";
            //query += " WHERE 1=1 ";
            //if (daData == null)
            //    daData = Convert.ToDateTime("01/01/2010");
            //query += " AND data >= @data ";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            //aCommand.Parameters.Add("@data", daData);

            DataTable dt = new DataTable();

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                dt.Load(aReader);
                aReader.Close();
            }
            catch (OleDbException e)
            {
                return null;
            }
            return dt;
        }

        public DataTable TotaliArticoli(DateTime? daData, DateTime? aData, string idArticolo, string bolla, bool isNonRaggruppare = false)
        {
            string codiceIntNum = string.Empty;
            if (!string.IsNullOrEmpty(idArticolo))
            {
                articolo articolo = LeggiArticolo(idArticolo);
                codiceIntNum = articolo.codiceIntNum;
            }

            string dataini = daData.ToString().Substring(3, 2) + "/" + daData.ToString().Substring(0, 2) + "/" + daData.ToString().Substring(6,4) + " 00.00";
            string datafin = aData.ToString().Substring(3, 2) + "/" + aData.ToString().Substring(0, 2) + "/" + aData.ToString().Substring(6,4) + " 23.59";
            string query = "";

            if(isNonRaggruppare)
            {
                query = "SELECT DISTINCTROW MovimentiOut.idArticolo, Articoli.descrizione, Sum(MovimentiOut.qta) AS Totale, Count(*) AS numMovimenti, Sum([MovimentiOut].[qta]*[Articoli].[PrezzoPhotosi]) AS Importo  ";
                query += " FROM MovimentiOut ";
                query += " INNER JOIN Articoli ON MovimentiOut.idArticolo = Articoli.idArticolo ";
                query += " WHERE (((MovimentiOut.data)>=#" + dataini + "# And (MovimentiOut.data)<=#" + datafin + "#)) ";
                if (!string.IsNullOrEmpty(bolla))
                    query += " AND MovimentiOut.numBolla = " + bolla;
                query += "GROUP BY MovimentiOut.idArticolo, Articoli.descrizione ";
                query += "Having MovimentiOut.idArticolo = '" + idArticolo + "'";
            }
            else
            {
                query = "SELECT DISTINCTROW Articoli.CodiceInternoNum AS idArticolo, (select top 1 descrizione from articoli as art1 where art1.CodiceInternoNum=articoli.CodiceInternoNum) AS descrizione, Sum(MovimentiOut.qta) AS Totale, Count(*) AS numMovimenti, Sum([MovimentiOut].[qta]*[Articoli].[PrezzoPhotosi]) AS Importo ";
                query += " FROM MovimentiOut ";
                query += " INNER JOIN Articoli ON MovimentiOut.idArticolo = Articoli.idArticolo ";
                query += " WHERE (((MovimentiOut.data)>=#" + dataini + "# And (MovimentiOut.data)<=#" + datafin + "#)) ";
                if (!string.IsNullOrEmpty(bolla))
                    query += " AND MovimentiOut.numBolla = " + bolla;
                query += " GROUP BY Articoli.CodiceInternoNum ";
                if (!string.IsNullOrEmpty(codiceIntNum))
                    query += " Having Articoli.CodiceInternoNum = '" + codiceIntNum + "'";
            }

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            //if (daData != null)
            //    aCommand.Parameters.Add("@daData", daData);
            //if (aData != null)
            //    aCommand.Parameters.Add("@aData", aData);

            DataTable dt = new DataTable();

            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                dt.Load(aReader);
                aReader.Close();
            }
            catch (OleDbException e)
            {
                return null;
            }
            return dt;
        }

        public int CancellaRighe(string tipo, string idBusta, string idArticolo, int qta)
        {
            string tabella = "Movimenti" + tipo;

            string query = "DELETE FROM " + tabella +
                " WHERE idBusta = '" + idBusta + "' AND idArticolo = '" + idArticolo + "' AND qta = " + qta;

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@tipo", tipo);
            aCommand.Parameters.Add("@idBusta", idBusta);
            aCommand.Parameters.Add("@idArticolo", idArticolo);
            aCommand.Parameters.Add("@qta", qta);
            int ret = aCommand.ExecuteNonQuery();

            return ret;
        }

        public int QtaDiffArticolo(string idBusta, string idArticolo)
        {
            string data = leggiAppSettings("DaDataLettura");
            DateTime ?daData = null;
            if(!string.IsNullOrEmpty(data))
                daData = Convert.ToDateTime(data);

            DataTable dtIn = ListaMovimentiArticolo(idBusta, idArticolo, "In", daData);
            DataTable dtOut = ListaMovimentiArticolo(idBusta, idArticolo, "Out", daData);
            DataTable dtResi = ListaMovimentiArticolo(idBusta, idArticolo, "Resi", daData);

            int QtaDiff = 0;

            foreach (DataRow dr in dtIn.Rows)
            {
                try
                {
                    QtaDiff += (int)dr["qta"];
                }
                catch { }
            }

            foreach (DataRow dr in dtOut.Rows)
            {
                try
                {
                    QtaDiff -= (int)dr["qta"];
                }
                catch { }
            }

            foreach (DataRow dr in dtResi.Rows)
            {
                try
                {
                    QtaDiff -= (int)dr["qta"];
                }
                catch { }
            }

            return QtaDiff;
        }

        public bool isArticoloMascherina(string codice)
        {
            bool ret = false;
            string codicimascherina = leggiAppSettings("CodiciMascherina");
            if (string.IsNullOrEmpty(codicimascherina))
                codicimascherina = "3894,3895,3896,3514";
            if (codicimascherina.IndexOf(codice) >= 0)
                ret = true;

            return ret;
        }

        public bool EliminaMovimentiIn(DateTime giorno)
        {
            string query = "delete * from MovimentiIn where data=@data";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            aCommand.Parameters.Clear();
            aCommand.Parameters.Add("@data", giorno.ToShortDateString());
            aCommand.ExecuteNonQuery();

            return true;

        }

        public DataTable TotaliMonoprodotto()
        {
            string query = "select * from [Monoprodotto totali giornalieri]";

            OleDbCommand aCommand = new OleDbCommand(query, aConnection);
            //aCommand.Parameters.Clear();
            //aCommand.Parameters.Add("@chiave", chiave);

            DataTable dt = new DataTable();
            try
            {
                OleDbDataReader aReader = aCommand.ExecuteReader();
                dt.Load(aReader);
                aReader.Close();
            }
            catch (Exception e)
            {
                dt = null;
            }
            return dt;

            aCommand.Dispose();
        }


    }
}
