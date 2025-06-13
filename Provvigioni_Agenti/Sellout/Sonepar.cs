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
    internal class Sonepar : ITrasferitiService
    {
        private List<Trasferito> _trasferito = null;
        private List<Trasferito> _trasferitoMese = null; // 2025-05-05 trasferito mensile
        public Sonepar(string anno, List<string> mesi)
        {
            _trasferito = new List<Trasferito>();

            this.leggiAgenzia(anno, mesi);
        }


        private void leggiAgenzia(string anno, List<string> mesi)
        {
            // vedo se ci sono files in barcella

            foreach (string mese in mesi)
            {
                _trasferitoMese = new List<Trasferito>(); // 2025-05-05 

                List<string> listFiles = IntefaceClass.elencaFiles(anno, mese, TrasferitiAgenzie.Sonepar);

                string path = $"../trasferiti/{anno}/{mese}/{TrasferitiAgenzie.Sonepar}";

                if (listFiles.Count == 0)
                {
                    continue;
                }

                List<Citta> citta = null;
                //2025-04-22
                List<string> nuoveCitta = new List<string>();
                //apri xml citta
                XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
                using (TextReader tr = new StreamReader("citta_sonepar.xml"))
                {
                    citta = (List<Citta>)xmlsd.Deserialize(tr);
                }

                string file = listFiles[0];
                var wb = new XLWorkbook($"{path}/{file}");


                var ws = wb.Worksheet(1);

                int colonnaRegione = 0;
                int colonnaValore = 0;
                int rowRead = 0;

                // Trasferito snp = null;

                string nomeRegione = string.Empty;

                var statoLettura = SoneparStatoLettura.Init;

                foreach (var row in ws.Rows())
                {
                    foreach (var cell in row.Cells())
                    {
                        if (!cell.IsEmpty())
                        {
                            int r = cell.Address.RowNumber;
                            int c = cell.Address.ColumnNumber;

                            string cellValue = cell.Value.ToString().Trim(' ');

                            if (cellValue.Contains("Regione"))
                            {
                                colonnaRegione = c;
                                colonnaValore = colonnaRegione + 4;

                                statoLettura = SoneparStatoLettura.LeggiValori;

                                rowRead = r + 1;

                            }



                            switch (statoLettura)
                            {
                                case SoneparStatoLettura.LeggiValori:

                                    if (c == colonnaRegione && r >= rowRead)
                                    {





                                        var regioneProvincia = citta.Find(x => x.Comune == cellValue);
                                        //2025-04-22
                                        if (regioneProvincia == null)
                                        {
                                            nuoveCitta.Add(cellValue);

                                            continue;
                                        }


                                        var result = _trasferito.Find(x => x.Regione == regioneProvincia.Regione);

                                        if (result == null)
                                        {
                                            _trasferito.Add(new Trasferito() { Regione = regioneProvincia.Regione });
                                        }

                                        var resultMese = _trasferitoMese.Find(x => x.Regione == regioneProvincia.Regione);

                                        if (resultMese == null)
                                        {
                                            _trasferitoMese.Add(new Trasferito() { Regione = regioneProvincia.Regione });
                                        }

                                        result = _trasferito.Find(x => x.Regione == regioneProvincia.Regione);
                                        result.Venduto += Double.Parse(ws.Cell(r, colonnaValore).Value.ToString());

                                        resultMese = _trasferitoMese.Find(x => x.Regione == regioneProvincia.Regione);
                                        resultMese.Venduto += Double.Parse(ws.Cell(r, colonnaValore).Value.ToString());


                                    }

                                    break;
                            }

                        }
                        else
                        {

                        }
                    }
                }

                IntefaceClass.serializzaXml(anno, mese, TrasferitiAgenzie.Sonepar, $"{TrasferitiAgenzie.Sonepar}", _trasferitoMese);


            }

        }

        public IList<Trasferito> Trasferito => _trasferito;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}