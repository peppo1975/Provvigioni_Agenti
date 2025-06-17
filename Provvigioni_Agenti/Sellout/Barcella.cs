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

    internal class Barcella : ITrasferitiService
    {
        private List<Trasferito> _trasferito = null;
        private List<Trasferito> _trasferitoMese = null; // 2025-05-05 trasferito mensile
        private List<string> _nuoveCitta = null;
        public Barcella(string anno, List<string> mesi)
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
                _trasferitoMese = new List<Trasferito>(); // 2025-05-05 

                List<string> listFiles = IntefaceClass.elencaFiles(anno, mese, TrasferitiAgenzie.Barcella);

               

                List<Citta> citta = new List<Citta>();

                string path = $"../trasferiti/{anno}/{mese}/{TrasferitiAgenzie.Barcella}";

                //2025-04-22
            

                if (listFiles.Count == 0)
                {
                   continue;
                }

                string file = listFiles[0];
                var wb = new XLWorkbook($"{path}/{file}");

                //apri xml citta
                XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
                using (TextReader tr = new StreamReader("citta_barcella.xml"))
                {
                    citta = (List<Citta>)xmlsd.Deserialize(tr);
                }



              
                {
                    var ws = wb.Worksheet(1);

                    int colonnaRegione = 0;
                    int rigaRegione = 0;
                    int colonnaVenduto = 0;


                    Trasferito brc = null;
                    Trasferito brcMese = null;

                    string nomeRegione = string.Empty;

                    foreach (var row in ws.Rows())
                    {
                        foreach (var cell in row.Cells())
                        {
                            if (!cell.IsEmpty())
                            {
                                int r = cell.Address.RowNumber;
                                int c = cell.Address.ColumnNumber;

                                string cellValue = cell.Value.ToString().Trim(' ');
                                if (cellValue.Contains("Punto Vendita") || cellValue.Contains("Provincia Cliente"))
                                {
                                    colonnaRegione = c;
                                    rigaRegione = r + 1;
                                    colonnaVenduto = c + 1;

                                }

                                if (r >= rigaRegione && c == colonnaRegione)
                                {

                                    var regioneCitta = citta.Find(x => x.Comune == cellValue);

                                    //2025-04-22
                                    if (regioneCitta == null)
                                    {
                                        _nuoveCitta.Add(cellValue);
                                        break;
                                        //continue;
                                    }

                                    string regione = regioneCitta.Regione;

                                    var result = _trasferito.Find(x => x.Regione == regione);
                                    var resultMese = _trasferitoMese.Find(x => x.Regione == regione);// 2025-05-05

                                    nomeRegione = regione;

                                    if (result == null)
                                    {

                                        brc = new Trasferito();

                                        brc.Regione = nomeRegione;
                                        _trasferito.Add(brc);
                                    }



                                    if (resultMese == null)
                                    {

                                        brcMese = new Trasferito();

                                        brcMese.Regione = nomeRegione;
                                        _trasferitoMese.Add(brcMese);
                                    }


                                    result = _trasferito.Find(x => x.Regione == nomeRegione);
                                    result.Venduto += Double.Parse(ws.Cell(r, colonnaVenduto).Value.ToString().Trim(' '));

                                    resultMese = _trasferitoMese.Find(x => x.Regione == nomeRegione);
                                    resultMese.Venduto += Double.Parse(ws.Cell(r, colonnaVenduto).Value.ToString().Trim(' '));
                                    break;
                                }

                            }
                            else
                            {

                            }
                        }
                    }
                }

                IntefaceClass.serializzaXml(anno, mese, TrasferitiAgenzie.Barcella, $"{TrasferitiAgenzie.Barcella}", _trasferitoMese);

            }

        }

        public IList<Trasferito> Trasferito => _trasferito;  //elemento pubblico che da modo di visualizzare un elemento privato
        public IList<String> NuoveCitta => _nuoveCitta;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}
