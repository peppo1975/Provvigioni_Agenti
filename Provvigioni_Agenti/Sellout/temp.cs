using ClosedXML.Excel;
using Provvigioni_Agenti.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Provvigioni_Agenti.Sellout
{
    internal class _temp
        : ITrasferitiService
    {
        private List<Trasferito> _trasferito = null;
        private List<Trasferito> _trasferitoMese = null; // 2025-05-05 trasferito mensile
        public _temp(string anno, List<string> mesi)
        {
            _trasferito = new List<Trasferito>();

            this.leggiAgenzia(anno, mesi);
        }

        private void leggiAgenzia(string anno, List<string> mesi)
        {
            // vedo se ci sono files in barcella

            foreach (string mese in mesi)
            {
                List<string> listFiles = IntefaceClass.elencaFiles(anno, mese, TrasferitiAgenzie.Barcella);

                _trasferitoMese = new List<Trasferito>(); // 2025-05-05 

                if (listFiles.Count == 0)
                {
                    continue;
                }

                List<Citta> citta = null;
                List<string> nuoveCitta = new List<string>();
                string path = $"../trasferiti/{anno}/{mese}/{TrasferitiAgenzie.Barcella}";

                //apri xml citta
                XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
                using (TextReader tr = new StreamReader("citta_mc_elettrici.xml"))
                {
                    citta = (List<Citta>)xmlsd.Deserialize(tr);
                }

                // apri 
                string file = listFiles[0];
                var wb = new XLWorkbook($"{path}/{file}");
                var ws = wb.Worksheet(1);

            
                IntefaceClass.serializzaXml(anno, mese, TrasferitiAgenzie.Sonepar, $"{TrasferitiAgenzie.Sonepar}", _trasferitoMese);

            }

        }

        public IList<Trasferito> Trasferito => _trasferito;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}
