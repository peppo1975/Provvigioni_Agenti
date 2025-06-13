using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
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
    internal class Comoli : ITrasferitiService
    {
        private List<Trasferito> _trasferito = null;
        private List<Trasferito> _trasferitoMese = null; // 2025-05-05 trasferito mensile
        public Comoli(string anno, List<string> mesi)
        {
            _trasferito = new List<Trasferito>();

            this.leggiAgenzia(anno, mesi);
        }

        private void leggiAgenzia(string anno, List<string> mesi)
        {
            // vedo se ci sono files in barcella

            foreach (string mese in mesi)
            {
                List<string> listFiles = IntefaceClass.elencaFiles(anno, mese, TrasferitiAgenzie.Comoli);

                _trasferitoMese = new List<Trasferito>(); // 2025-05-05 

                if (listFiles.Count == 0)
                {
                    continue;
                }

                List<Citta> citta = null;
                List<string> nuoveCitta = new List<string>();
                string path = $"../trasferiti/{anno}/{mese}/{TrasferitiAgenzie.Comoli}";

                //apri xml citta
                XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
                using (TextReader tr = new StreamReader("citta_comoli.xml"))
                {
                    citta = (List<Citta>)xmlsd.Deserialize(tr);
                }

                // apri 
                string file = listFiles[0];
                var wb = new XLWorkbook($"{path}/{file}");
                var ws = wb.Worksheet(1);

                int colonnaProvincia = 0;
                int rigaDati = 0;
                int colonnaVenduto = 0;
                string provincia = string.Empty;
                string regione = string.Empty;
                Citta regioneProvincia = new Citta();
                Trasferito cml = null;
                Trasferito cmlMese = null;

                foreach (var row in ws.Rows())
                {
                    foreach (var cell in row.Cells())
                    {
                        if (!cell.IsEmpty())
                        {
                            int r = cell.Address.RowNumber;
                            int c = cell.Address.ColumnNumber;

                            string cellValue = cell.Value.ToString().Trim(' ');
                            if (cellValue.Contains("Provincia"))
                            {
                                colonnaProvincia = c;
                                rigaDati = r + 1;
                            }

                            if (cellValue.Contains("Fatt al Costo"))
                            {
                                colonnaVenduto = c;
                            }
                        }
                    }

                    if (colonnaProvincia > 0 && colonnaVenduto > 0)
                    {
                        break;
                    }
                }


                foreach (var row in ws.Rows())
                {
                    int r = row.RowNumber();
                    if (rigaDati > r)
                    {
                        continue;
                    }
                    provincia = ws.Cell(r, colonnaProvincia).Value.ToString().Trim(' ');
                    regioneProvincia = citta.Find(x => x.Comune == provincia);
                    regione = regioneProvincia.Regione;

                    if (regioneProvincia == null && regione != string.Empty)
                    {
                        nuoveCitta.Add(regione.ToString());
                        continue;
                    }


                    var result = _trasferito.Find(x => x.Regione == regione);
                    var resultMese = _trasferitoMese.Find(x => x.Regione == regione);// 2025-05-05



                    if (result == null)
                    {

                        cml = new Trasferito();
                        cml.Regione = regione;
                        _trasferito.Add(cml);
                    }



                    if (resultMese == null)
                    {

                        cmlMese = new Trasferito();
                        cmlMese.Regione = regione;
                        _trasferitoMese.Add(cmlMese);
                    }


                    var venduto = ws.Cell(r, colonnaVenduto).Value.ToString();

                    result = _trasferito.Find(x => x.Regione == regione);
                    result.Venduto += Double.Parse(venduto.Trim(' '));

                    resultMese = _trasferitoMese.Find(x => x.Regione == regione);
                    resultMese.Venduto += Double.Parse(venduto.Trim(' '));






                    //_trasferito.Add(new Trasferito() { Regione = regioneProvincia.Regione, Venduto = Double.Parse(venduto.ToString()) });
                    //_trasferitoMese.Add(new Trasferito() { Regione = regioneProvincia.Regione, Venduto = Double.Parse(venduto.ToString()) });
                }


                IntefaceClass.serializzaXml(anno, mese, TrasferitiAgenzie.Sonepar, $"{TrasferitiAgenzie.Sonepar}", _trasferitoMese);

            }

        }

        public IList<Trasferito> Trasferito => _trasferito;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}
