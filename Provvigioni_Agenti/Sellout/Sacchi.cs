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
    internal class Sacchi : ITrasferitiService
    {
        private List<Trasferito> _trasferito = null;
        private List<Trasferito> _trasferitoMese = null; // 2025-05-05 trasferito mensile

        List<string> _nuoveCitta = null;
        public Sacchi(string anno, List<string> mesi)
        {
            _trasferito = new List<Trasferito>();

            this.leggiAgenzia(anno, mesi);
        }

        private void leggiAgenzia(string anno, List<string> mesi)
        {
            // vedo se ci sono files in barcella
            _nuoveCitta = new List<string>();

            foreach (string mese in mesi)
            {
                List<string> listFiles = IntefaceClass.elencaFiles(anno, mese, TrasferitiAgenzie.Sacchi);

                _trasferitoMese = new List<Trasferito>(); // 2025-05-05 

                if (listFiles.Count == 0)
                {
                    continue;
                }

                List<Citta> citta = null;

                string path = $"../trasferiti/{anno}/{mese}/{TrasferitiAgenzie.Sacchi}";

                //apri xml citta
                XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
                using (TextReader tr = new StreamReader("citta_sacchi.xml"))
                {
                    citta = (List<Citta>)xmlsd.Deserialize(tr);
                }

                // apri 
                string file = listFiles[0];
                var wb = new XLWorkbook($"{path}/{file}");
                var ws = wb.Worksheet(1);


                IntefaceClass.serializzaXml(anno, mese, TrasferitiAgenzie.Sacchi, $"{TrasferitiAgenzie.Sacchi}", _trasferitoMese);

            }

        }

        public IList<Trasferito> Trasferito => _trasferito;  //elemento pubblico che da modo di visualizzare un elemento privato
        public IList<String> NuoveCitta => _nuoveCitta;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}
