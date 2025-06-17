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

    //public interface ITrasferitiService
    //{
    //    IList<Final> Trasferiti { get; }
    //}

    internal class Acmei : ITrasferitiService
    {
        private List<Trasferito> _trasferito = null;
        private List<Trasferito> _trasferitoMese = null; // 2025-05-05 trasferito mensile
        private List<string> _nuoveCitta = null;
        public Acmei(string anno, List<string> mesi)
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
                List<string> listFiles = IntefaceClass.elencaFiles(anno, mese, TrasferitiAgenzie.Acmei);
                _trasferitoMese = new List<Trasferito>(); // 2025-05-05 


                if (listFiles.Count == 0)
                {
                    continue;
                }

                List<Citta> citta = null;

                string path = $"../trasferiti/{anno}/{mese}/{TrasferitiAgenzie.Acmei}";

                //apri xml citta
                XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
                using (TextReader tr = new StreamReader("citta_acmei.xml"))
                {
                    citta = (List<Citta>)xmlsd.Deserialize(tr);
                }

                // apri 
                string file = listFiles[0];
                var wb = new XLWorkbook($"{path}/{file}");
                var ws = wb.Worksheet(1);



                int colonnaRegione = 0;
                int rigaRegione = 0;
                int colonnaVenduto = 0;
                string regione = "";
                Citta regioneNome = new Citta();

                AcmeiStatoLettura statoLettura = AcmeiStatoLettura.Init;

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
                                colonnaRegione = c;
                                rigaRegione = r + 1;
                                //colonnaVenduto = c + 2;
                                statoLettura = AcmeiStatoLettura.LeggiNomeRegione;
                            }
                            if (cellValue.Contains("Venduto"))
                            {
                                colonnaVenduto = c;
                                statoLettura = AcmeiStatoLettura.LeggiNomeRegione;
                            }

                        }
                    }
                    if (colonnaRegione > 0 && rigaRegione > 0 && colonnaVenduto > 0)
                        break;
                }


                foreach (var row in ws.Rows())
                {
                    int r = row.RowNumber();
                    if (r < rigaRegione)
                    {
                        continue;
                    }
                    switch (statoLettura)
                    {
                        case AcmeiStatoLettura.LeggiNomeRegione:
                            regione = ws.Cell(r, colonnaRegione).Value.ToString().Trim(' ');
                            regioneNome = citta.Find(x => x.Comune == regione);


                            //if (regioneNome == null && regione != string.Empty)
                            if (regioneNome == null)
                            {
                                if (regione == string.Empty)
                                {
                                    statoLettura = AcmeiStatoLettura.LeggiNomeRegione;
                                    continue;
                                }
                                else
                                {
                                    _nuoveCitta.Add(regione.ToString());
                                    statoLettura = AcmeiStatoLettura.AttendiFineCittaNonSalvata;
                                    continue;

                                }


                            }


                            statoLettura = AcmeiStatoLettura.AttendiFine;

                            break;

                        case AcmeiStatoLettura.AttendiFineCittaNonSalvata:
                            if (ws.Cell(r, colonnaRegione).Value.ToString().Trim(' ').Contains("Totale"))
                            {
                                statoLettura = AcmeiStatoLettura.LeggiNomeRegione;
                            }
                            break;

                        case AcmeiStatoLettura.AttendiFine:
                            if (ws.Cell(r, colonnaRegione).Value.ToString().Trim(' ').Contains("Totale"))
                            {
                                var venduto = ws.Cell(r, colonnaVenduto).Value.ToString();


                                var result = _trasferito.Find(x => x.Regione == regioneNome.Regione);

                                if (result == null)
                                {
                                    _trasferito.Add(new Trasferito() { Regione = regioneNome.Regione });
                                }

                                var resultMese = _trasferitoMese.Find(x => x.Regione == regioneNome.Regione);

                                if (resultMese == null)
                                {
                                    _trasferitoMese.Add(new Trasferito() { Regione = regioneNome.Regione });
                                }



                                result = _trasferito.Find(x => x.Regione == regioneNome.Regione);
                                result.Venduto += Double.Parse(ws.Cell(r, colonnaVenduto).Value.ToString());

                                resultMese = _trasferitoMese.Find(x => x.Regione == regioneNome.Regione);
                                resultMese.Venduto += Double.Parse(ws.Cell(r, colonnaVenduto).Value.ToString());


                                //_trasferito.Add(new Trasferito() { Regione = regioneNome.Regione, Venduto = Double.Parse(venduto.ToString()) });
                                //_trasferitoMese.Add(new Trasferito() { Regione = regioneNome.Regione, Venduto = Double.Parse(venduto.ToString()) });

                                statoLettura = AcmeiStatoLettura.LeggiNomeRegione;
                            }

                            break;

                    }

                }
                IntefaceClass.serializzaXml(anno, mese, TrasferitiAgenzie.Acmei, $"{TrasferitiAgenzie.Acmei}", _trasferitoMese);
            }

        }

        public IList<Trasferito> Trasferito => _trasferito;  //elemento pubblico che da modo di visualizzare un elemento privato
        public IList<String> NuoveCitta => _nuoveCitta;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}
