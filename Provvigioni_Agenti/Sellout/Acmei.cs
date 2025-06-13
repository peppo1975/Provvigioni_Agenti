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
        public Acmei(string anno, List<string> mesi)
        {
            _trasferito = new List<Trasferito>();

            this.leggiAgenzia(anno, mesi);
        }

        private void leggiAgenzia(string anno, List<string> mesi)
        {
            // vedo se ci sono files in barcella

            foreach (string mese in mesi)
            {
                List<string> listFiles = IntefaceClass.elencaFiles(anno, mese, TrasferitiAgenzie.Acmei);
                _trasferitoMese = new List<Trasferito>(); // 2025-05-05 


                if (listFiles.Count == 0)
                {
                    continue;
                }

                List<Citta> citta = null;
                List<string> nuoveCitta = new List<string>();
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

                            if (cellValue.Contains("Regione Nazione"))
                            {
                                colonnaRegione = c;
                                rigaRegione = r + 1;
                                colonnaVenduto = c + 2;
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
                            statoLettura = AcmeiStatoLettura.AttendiFine;

                            if (regioneNome == null && regione != string.Empty)
                            {
                                nuoveCitta.Add(regione.ToString());
                                continue;
                            }

                            break;

                        case AcmeiStatoLettura.AttendiFine:
                            if(ws.Cell(r, colonnaRegione).Value.ToString().Trim(' ').Contains("Totale"))
                            {
                                var venduto = ws.Cell(r, colonnaVenduto).Value.ToString();
                                _trasferito.Add(new Trasferito() { Regione = regioneNome.Regione, Venduto = Double.Parse(venduto.ToString()) });

                                _trasferitoMese.Add(new Trasferito() { Regione = regioneNome.Regione, Venduto = Double.Parse(venduto.ToString()) });

                                statoLettura = AcmeiStatoLettura.LeggiNomeRegione;
                            }

                            break;

                    }

                }
                IntefaceClass.serializzaXml(anno, mese, TrasferitiAgenzie.Acmei, $"{TrasferitiAgenzie.Acmei}", _trasferitoMese);
            }

        }

        public IList<Trasferito> Trasferito => _trasferito;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}
