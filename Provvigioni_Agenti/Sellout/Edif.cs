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
    internal class Edif : ITrasferitiService
    {
        private List<Trasferito> _trasferito = null;
        private List<Trasferito> _trasferitoMese = null; // 2025-05-05 trasferito mensile
        public Edif(string anno, List<string> mesi)
        {
            _trasferito = new List<Trasferito>();

            this.leggiAgenzia(anno, mesi);
        }

        private void leggiAgenzia(string anno, List<string> mesi)
        {
            // vedo se ci sono files in barcella

            foreach (string mese in mesi)
            {
                List<string> listFiles = IntefaceClass.elencaFiles(anno, mese, TrasferitiAgenzie.Edif);

                _trasferitoMese = new List<Trasferito>(); // 2025-05-05 

                if (listFiles.Count == 0)
                {
                    continue;
                }

                List<Citta> citta = null;
                List<string> nuoveCitta = new List<string>();
                string path = $"../trasferiti/{anno}/{mese}/{TrasferitiAgenzie.Edif}";

                //apri xml citta
                XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
                using (TextReader tr = new StreamReader("citta_edif.xml"))
                {
                    citta = (List<Citta>)xmlsd.Deserialize(tr);
                }

                // apri 
                string file = listFiles[0];
                var wb = new XLWorkbook($"{path}/{file}");
                var ws = wb.Worksheet(1);


                int colonnaArticoli = 0;
                int colonnaFinale = 0;


                string nomeRegione = string.Empty;

                var statoLettura = EdifStatoLettura.Init;

                foreach (var row in ws.Rows())
                {
                    foreach (var cell in row.Cells())
                    {
                        if (!cell.IsEmpty())
                        {
                            int r = cell.Address.RowNumber;
                            int c = cell.Address.ColumnNumber;
                            string cellValue = cell.Value.ToString().Trim(' ');

                            if (cellValue.Contains("Costo Netto del Venduto"))
                            {
                                colonnaArticoli = c;
                                statoLettura = EdifStatoLettura.ScorriRiga;
                                //  IndiceRegione = new List<IndiceRegione>();
                            }

                            switch (statoLettura)
                            {
                                case EdifStatoLettura.ScorriRiga:
                                    statoLettura = EdifStatoLettura.LeggiRegione;
                                    break;




                                case EdifStatoLettura.LeggiRegione:
                                    if (cellValue.Contains("TOTALE"))
                                    {
                                        statoLettura = EdifStatoLettura.LeggiValori;
                                        colonnaFinale = c;
                                    }
                                    else
                                    {
                                        nomeRegione = cellValue;
                                        var regioneProvincia = citta.Find(x => x.Comune == nomeRegione);
                                        //2025-04-22
                                        if (regioneProvincia == null)
                                        {
                                            nuoveCitta.Add(nomeRegione);

                                            continue;
                                        }

                                        //var result = _trasferito.Find(x => x.Regione == regioneProvincia.Regione);
                                        //var resultMese = _trasferitoMese.Find(x => x.Regione == regioneProvincia.Regione);// 2025-05-05
                                        _trasferito.Add(new Trasferito() { Regione = regioneProvincia.Regione });
                                        _trasferitoMese.Add(new Trasferito() { Regione = regioneProvincia.Regione });
                                    }
                                    break;




                                case EdifStatoLettura.LeggiValori:
                                    if (c == colonnaArticoli)
                                    {
                                        if (cellValue.Contains("TOTALE")) // arriva all'ultima riga
                                        {
                                            statoLettura = EdifStatoLettura.Exit;
                                        }
                                    }
                                    else
                                    {

                                        //Console.WriteLine(c);


                                        if (c < colonnaFinale)
                                        {
                                            _trasferito[c - 2].Venduto += Double.Parse(cellValue);
                                            _trasferitoMese[c - 2].Venduto += Double.Parse(cellValue);
                                        }
                                    }
                                    break;
                            }

                        }
                    }
                }


                IntefaceClass.serializzaXml(anno, mese, TrasferitiAgenzie.Edif, $"{TrasferitiAgenzie.Edif}", _trasferitoMese);

            }

        }

        public IList<Trasferito> Trasferito => _trasferito;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}
