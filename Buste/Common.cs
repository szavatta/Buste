using System;
using System.Collections.Generic;
using System.Text;

namespace Buste
{
    public struct busta
    {
        public string idBusta;
        public DateTime data;
        public string idBatch;
        public string codArticolo;
        public int quantita;
        public string tipoOrdine;
        public string numProdotti;
    }

    public struct articolo
    {
        public string codice;
        public string descrizione;
        public string codiceInt;
        public string codiceIntNum;
        public string codiceFatturazione;
        public int scaglione1;
        public decimal sconto1;
        public int scaglione2;
        public decimal sconto2;
        public int scaglione3;
        public decimal sconto3;
        public int scaglione4;
        public decimal sconto4;
        public int scaglione5;
        public decimal sconto5;
        public int omaggio;
    }

    public struct movimento
    {
        public string idBusta;
        public string idBatch;
        public string idArticolo;
        public string desArticolo;
        public DateTime data;
        public int qta;
        public string tipoOrdine;
        public string numProdotti;
    }

    public enum TipoSuono
    {
        Errore,
        OkInserimento,
        OkInserimentoMulti,
        OkInserimentoMonoprodotto,
        Domanda,
        Kit,
        NoSuono
    }

    [Serializable()]
    public class Common
    {

        public Common()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        protected void Beep()
        {
            Console.Beep();
        }

        public bool isNumerico(string codice)
        {
            bool ret = true;
            for (int i = 0; i < codice.Length; i++)
            {
                if ("0123456789".LastIndexOf(codice.Substring(i, 1)) < 0)
                {
                    ret = false;
                    break;
                }
            }
            return ret;
        }

        public bool ContainsString(string[] arr, string testval)
        {
            if (arr == null)
                return false;
            for (int i = arr.Length - 1; i >= 0; i--)
                if (arr[i] == testval)
                    return true;
            return false;
        }

    }

}
