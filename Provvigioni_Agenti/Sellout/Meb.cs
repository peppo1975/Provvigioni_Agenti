using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
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
    internal class Meb
        : ITrasferitiService
    {
        private List<Trasferito> _trasferito = null;
        private List<Trasferito> _trasferitoMese = null; // 2025-05-05 trasferito mensile
        public Meb(string anno, List<string> mesi)
        {
            _trasferito = new List<Trasferito>();

            this.leggiAgenzia(anno, mesi);
        }

        private void leggiAgenzia(string anno, List<string> mesi)
        {
            // vedo se ci sono files in barcella

            foreach (string mese in mesi)
            {
                List<string> listFiles = IntefaceClass.elencaFiles(anno, mese, TrasferitiAgenzie.Meb);

                _trasferitoMese = new List<Trasferito>(); // 2025-05-05 

                if (listFiles.Count == 0)
                {
                    continue;
                }

                List<Citta> citta = null;
                List<string> nuoveCitta = new List<string>();
                string path = $"../trasferiti/{anno}/{mese}/{TrasferitiAgenzie.Meb}";

                //apri xml citta
                XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
                using (TextReader tr = new StreamReader("citta_meb.xml"))
                {
                    citta = (List<Citta>)xmlsd.Deserialize(tr);
                }

                // apri 
                string file = listFiles[0];
                var wb = new XLWorkbook($"{path}/{file}");
                var ws = wb.Worksheet(1);


                int colonnaRegione = 0;
                int rigaRegione = 1000;
                int colvenduto = 0;


                foreach (var row in ws.Rows())
                {
                    foreach (var cell in row.Cells())
                    {
                        if (!cell.IsEmpty())
                        {
                            int r = cell.Address.RowNumber;
                            int c = cell.Address.ColumnNumber;

                            string cellValue = cell.Value.ToString().Trim(' ');
                            if (cellValue.Contains("Regione") && r == 1)
                            {
                                colonnaRegione = c;
                                rigaRegione = r + 1;
                            }

                            if (r >= rigaRegione)
                            {
                                var reg = ws.Cell(r, colonnaRegione).Value;
                                var regioneProvincia = citta.Find(x => x.Comune == reg.ToString());

                                if (regioneProvincia == null)
                                {
                                    nuoveCitta.Add(reg.ToString());

                                    continue;
                                }

                          

                                if (mese == Mesi.Gennaio)
                                {
                                    colvenduto = colonnaRegione + 6;
                                }

                                if (mese == Mesi.Febbraio)
                                {
                                    colvenduto = colonnaRegione + 7;
                                }

                                if (mese == Mesi.Marzo)
                                {
                                    colvenduto = colonnaRegione + 8;
                                }

                                if (mese == Mesi.Aprile)
                                {
                                    colvenduto = colonnaRegione + 9;
                                }

                                if (mese == Mesi.Maggio)
                                {
                                    colvenduto = colonnaRegione + 10;
                                }

                                if (mese == Mesi.Giugno)
                                {
                                    colvenduto = colonnaRegione + 11;
                                }

                                if (mese == Mesi.Luglio)
                                {
                                    colvenduto = colonnaRegione + 12;
                                }

                                if (mese == Mesi.Agosto)
                                {
                                    colvenduto = colonnaRegione + 13;
                                }

                                if (mese == Mesi.Settembre)
                                {
                                    colvenduto = colonnaRegione + 14;
                                }

                                if (mese == Mesi.Ottobre)
                                {
                                    colvenduto = colonnaRegione + 15;
                                }

                                if (mese == Mesi.Novembre)
                                {
                                    colvenduto = colonnaRegione + 16;
                                }

                                if (mese == Mesi.Dicembre)
                                {
                                    colvenduto = colonnaRegione + 17;
                                }

                                var venduto = ws.Cell(r, colvenduto).Value.ToString();

                                var resultMese = _trasferitoMese.Find(x => x.Regione == regioneProvincia.Regione);
                                if (resultMese == null)
                                {

                                    _trasferitoMese.Add(new() { Regione = regioneProvincia.Regione });
                                }


                                var resultAll = _trasferito.Find(x => x.Regione == regioneProvincia.Regione);
                                if (resultAll == null)
                                {

                                    _trasferito.Add(new() { Regione = regioneProvincia.Regione });
                                }




                                resultMese = _trasferitoMese.Find(x => x.Regione == regioneProvincia.Regione);
                                resultAll = _trasferito.Find(x => x.Regione == regioneProvincia.Regione);

                            
                                resultMese.Venduto += Double.Parse(venduto);
                                resultAll.Venduto += Double.Parse(venduto);
                                break;
                            }


                        }


                    }

                }

                IntefaceClass.serializzaXml(anno, mese, TrasferitiAgenzie.Meb, $"{TrasferitiAgenzie.Meb}", _trasferitoMese);
            }
        }

        public IList<Trasferito> Trasferito => _trasferito;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}
