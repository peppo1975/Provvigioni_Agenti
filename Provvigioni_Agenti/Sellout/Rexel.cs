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
    internal class Rexel
        : ITrasferitiService
    {
        private List<Trasferito> _trasferito = null;
        private List<Trasferito> _trasferitoMese = null; // 2025-05-05 trasferito mensile
        List<string> _nuoveCitta = null;
        public Rexel(string anno, List<string> mesi)
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
                List<string> listFiles = IntefaceClass.elencaFiles(anno, mese, TrasferitiAgenzie.Rexel);

                _trasferitoMese = new List<Trasferito>(); // 2025-05-05 

                if (listFiles.Count == 0)
                {
                    continue;
                }

                List<Citta> citta = null;

                string path = $"../trasferiti/{anno}/{mese}/{TrasferitiAgenzie.Rexel}";

                //apri xml citta
                XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
                using (TextReader tr = new StreamReader("citta_rexel.xml"))
                {
                    citta = (List<Citta>)xmlsd.Deserialize(tr);
                }

                // apri 
                string file = listFiles[0];
                var wb = new XLWorkbook($"{path}/{file}");
                var ws = wb.Worksheet(1);



                int colonnaComune = 0;
                int colonnaAnno = 0;
                int rowRead = 0;

                string nomeRegione = string.Empty;

                var statoLettura = RexelStatoLettura.Init;
                int mVenduto = 0;



                foreach (var row in ws.Rows())
                {
                    foreach (var cell in row.Cells())
                    {
                        if (!cell.IsEmpty())
                        {
                            int r = cell.Address.RowNumber;
                            int c = cell.Address.ColumnNumber;

                            string cellValue = cell.Value.ToString().Trim(' ');

                            if (cellValue.Contains("Punto Vendita"))
                            {
                                colonnaComune = c;
                                colonnaAnno = colonnaComune + 1;


                                statoLettura = RexelStatoLettura.StabilisciTriemestre;

                                rowRead = r + 1;

                            }

                            switch (statoLettura)
                            {
                                case RexelStatoLettura.StabilisciTriemestre:


                                    if (mese == Mesi.Gennaio)
                                    {
                                        mVenduto = colonnaComune + 2;
                                    }

                                    if (mese == Mesi.Febbraio)
                                    {
                                        mVenduto = colonnaComune + 3;
                                    }

                                    if (mese == Mesi.Marzo)
                                    {
                                        mVenduto = colonnaComune + 4;
                                    }

                                    if (mese == Mesi.Aprile)
                                    {
                                        mVenduto = colonnaComune + 5;
                                    }

                                    if (mese == Mesi.Maggio)
                                    {
                                        mVenduto = colonnaComune + 6;
                                    }

                                    if (mese == Mesi.Giugno)
                                    {
                                        mVenduto = colonnaComune + 7;
                                    }

                                    if (mese == Mesi.Luglio)
                                    {
                                        mVenduto = colonnaComune + 8;
                                    }

                                    if (mese == Mesi.Agosto)
                                    {
                                        mVenduto = colonnaComune + 9;
                                    }

                                    if (mese == Mesi.Settembre)
                                    {
                                        mVenduto = colonnaComune + 10;
                                    }

                                    if (mese == Mesi.Ottobre)
                                    {
                                        mVenduto = colonnaComune + 11;
                                    }

                                    if (mese == Mesi.Novembre)
                                    {
                                        mVenduto = colonnaComune + 12;
                                    }

                                    if (mese == Mesi.Dicembre)
                                    {
                                        mVenduto = colonnaComune + 13;
                                    }

                                    statoLettura = RexelStatoLettura.LeggiValori;

                                    break;

                                case RexelStatoLettura.LeggiValori:

                                    if (r >= rowRead)
                                    {
                                        if ((c == colonnaComune) && (ws.Cell(r, colonnaAnno).Value.ToString() == anno))
                                        {
                                            var cittaNome = citta.Find(x => x.Comune == cellValue);


                                            //2025-04-22
                                            if (cittaNome == null)
                                            {
                                                _nuoveCitta.Add(cellValue);
                                                continue;
                                            }





                                            var resMese = _trasferitoMese.Find(x => x.Regione == cittaNome.Regione);
                                            if (resMese == null)
                                            {
                                                _trasferitoMese.Add(new() { Regione = cittaNome.Regione });
                                            }
                                            resMese = _trasferitoMese.Find(x => x.Regione == cittaNome.Regione);
                                            resMese.Venduto += Double.Parse(ws.Cell(r, mVenduto).Value.ToString());




                                            var resTot = _trasferitoMese.Find(x => x.Regione == cittaNome.Regione);
                                            if (resTot == null)
                                            {
                                                _trasferito.Add(new() { Regione = cittaNome.Regione });
                                            }
                                            resTot = _trasferitoMese.Find(x => x.Regione == cittaNome.Regione);
                                            resTot.Venduto += Double.Parse(ws.Cell(r, mVenduto).Value.ToString());
                                        }
                                    }
                                    break;
                            }
                        }
                        else
                        {

                        }
                    }
                }

                IntefaceClass.serializzaXml(anno, mese, TrasferitiAgenzie.Rexel, $"{TrasferitiAgenzie.Rexel}", _trasferitoMese);

            }

        }

        public IList<Trasferito> Trasferito => _trasferito;  //elemento pubblico che da modo di visualizzare un elemento privato
        public IList<String> NuoveCitta => _nuoveCitta;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}
