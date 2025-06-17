using ClosedXML.Excel;
using Provvigioni_Agenti.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Provvigioni_Agenti.Sellout
{
    internal class McElettrici : ITrasferitiService
    {
        private List<Trasferito> _trasferito = null;
        private List<Trasferito> _trasferitoMese = null; // 2025-05-05 trasferito mensile
        private List<string> _nuoveCitta = null;
        public McElettrici(string anno, List<string> mesi, int scheda)
        {
            _trasferito = new List<Trasferito>();

            this.leggiAgenzia(anno, mesi, scheda);
        }

        private void leggiAgenzia(string anno, List<string> mesi, int scheda)
        {
            // vedo se ci sono files in barcella
            _nuoveCitta = new List<string>();

            foreach (string mese in mesi)
            {
                _trasferitoMese = new List<Trasferito>(); // 2025-05-05 

                List<string> listFiles = IntefaceClass.elencaFiles(anno, mese, TrasferitiAgenzie.McElettrici);

                if (listFiles.Count == 0)
                {
                    continue;
                }

                List<Citta> citta = null;

                string path = $"../trasferiti/{anno}/{mese}/{TrasferitiAgenzie.McElettrici}";

                CultureInfo culture = null;
                culture = CultureInfo.CreateSpecificCulture("fr-FR");


                //apri xml citta
                XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
                using (TextReader tr = new StreamReader("citta_mc_elettrici.xml"))
                {
                    citta = (List<Citta>)xmlsd.Deserialize(tr);
                }

                string file = listFiles[0];
                var wb = new XLWorkbook($"{path}/{file}");
                var ws = wb.Worksheet(scheda);

                int colonnaRegione = 0;
                int colonnaVenduto = 0;
                int rigaDati = 1000;




                string nomeRegione = string.Empty;
                string venduto = string.Empty;

                Trasferito mce = null;

                double sales = 0;
                double scontato = 0.9;

                foreach (var row in ws.Rows())
                {
                    foreach (var cell in row.Cells())
                    {
                        int r = cell.Address.RowNumber;
                        int c = cell.Address.ColumnNumber;

                        string cellValue = cell.Value.ToString().Trim(' ');

                        if (cellValue.Contains("IMP_SPED"))
                        {
                            colonnaVenduto = c;
                        }

                        if (cellValue.Contains("PRODREG"))
                        {
                            colonnaRegione = c;
                            rigaDati = r + 1;
                        }
                    }

                    if (colonnaVenduto > 0 && colonnaRegione > 0)
                    {
                        break;
                    }
                }


                foreach (var row in ws.Rows())
                {
                    int r = row.RowNumber();

                    if (r >= rigaDati)
                    {
                        venduto = ws.Cell(r, colonnaVenduto).Value.ToString().Trim(' ');
                        nomeRegione = ws.Cell(r, colonnaRegione).Value.ToString().Trim(' ');

                        var regioneProvincia = citta.Find(x => x.Comune == nomeRegione);

                        //2025-04-22
                        if (regioneProvincia == null)
                        {
                            _nuoveCitta.Add(nomeRegione);

                            continue;
                        }

                        var result = _trasferito.Find(x => x.Regione == regioneProvincia.Regione);
                        var resultMese = _trasferitoMese.Find(x => x.Regione == regioneProvincia.Regione);// 2025-05-05

                        if (result == null)
                        {
                            _trasferito.Add(new Trasferito() { Regione = regioneProvincia.Regione });
                        }

                        if (resultMese == null)// 2025-05-05
                        {
                            _trasferitoMese.Add(new Trasferito() { Regione = regioneProvincia.Regione });
                        }// --------------------------------------------------------------------------------------

                        result = _trasferito.Find(x => x.Regione == regioneProvincia.Regione);
                        result.Venduto += Double.Parse(venduto, culture) * scontato; // sconto 10% fisso 

                        // 2025-05-05 ------------------------------------------------------------------------------------
                        resultMese = _trasferitoMese.Find(x => x.Regione == regioneProvincia.Regione);
                        resultMese.Venduto += Double.Parse(venduto, culture) * scontato; // sconto 10% fisso 
                                                                                         // -----------------------------------------------------------------------------------------------


                    }
                }

                // salva xml
                string trasferitoName = scheda == 1 ? "_magazzino" : "_vendita_diretta";
                IntefaceClass.serializzaXml(anno, mese, TrasferitiAgenzie.McElettrici, $"{TrasferitiAgenzie.McElettrici}{trasferitoName}", _trasferitoMese);

            }

        }

        public IList<Trasferito> Trasferito => _trasferito;  //elemento pubblico che da modo di visualizzare un elemento privato
        public IList<String> NuoveCitta => _nuoveCitta;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}
