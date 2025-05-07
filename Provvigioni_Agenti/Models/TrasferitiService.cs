using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Serialization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Spreadsheet;
using Provvigioni_Agenti.Controllers;

namespace Provvigioni_Agenti.Models
{

    public interface ITrasferitiService
    {
        IList<Final> Trasferiti { get; }
    }

    internal class TrasferitiService : ITrasferitiService
    {
        private List<Final> _trasferiti = null;
        List<Regione> regione = new List<Regione>();
        string annoCorrente = string.Empty;
        string trimestre = string.Empty;
        List<string> elencoTrasferiti = new List<string>();
        public TrasferitiService(List<Regione> regione, string annoCorrente, string trimestre, List<string> elencoTrasferiti)
        {
            _trasferiti = new List<Final>();
            this.regione = regione;
            this.annoCorrente = annoCorrente;
            this.trimestre = trimestre;
            this.elencoTrasferiti = elencoTrasferiti;

            leggiDirectory();
        }

        private void leggiDirectory()
        {
            if (trimestre.Length == 0)
            {
                return;
            }

            string path = $"../trasferiti/{annoCorrente}/{trimestre}";

            List<Trasferito> acmei = new List<Trasferito>();
            List<Trasferito> meb = new List<Trasferito>();
            List<Trasferito> mc_elettrici = new List<Trasferito>();
            List<Trasferito> barcella = new List<Trasferito>();
            List<Trasferito> comoli = new List<Trasferito>();
            List<Trasferito> edif = new List<Trasferito>();
            List<Trasferito> rexel = new List<Trasferito>();
            List<Trasferito> sonepar = new List<Trasferito>();
            List<Trasferito> sacchi = new List<Trasferito>();

            List<Final> finale = new List<Final>();

            Final estratti = null;
            Trasferito result = null;
            string nomeFornitore = string.Empty;

            try
            {



            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }


            foreach (string trasferito in elencoTrasferiti)
            {
                string pathTrasferito = $"{path}/{trasferito}";
                List<string> elencoFiles = General.elencoFile(pathTrasferito);

                elencoFiles.Remove(elencoFiles.Find(x => x.Contains(".xml")));

                //acmei ha solo un file
                switch (trasferito)
                {
                    case "acmei": //
                        acmei = elaboraAcmei(pathTrasferito, elencoFiles);
                        nomeFornitore = "ACMEI";
                        result = cercaInList(this.regione, acmei);
                        serializzaXml(annoCorrente, trimestre, trasferito, trasferito, acmei);
                        break;

                    case "meb": //
                        meb = elaboraMeb(pathTrasferito, elencoFiles, trimestre);
                        nomeFornitore = "MEB";
                        result = cercaInList(this.regione, meb);
                        serializzaXml(annoCorrente, trimestre, trasferito, trasferito, meb);
                        break;

                    case "mc_elettrici": //
                        mc_elettrici = elaboraMcElettrici(pathTrasferito, elencoFiles, 1);
                        nomeFornitore = "MC ELETTRICI (Magazzino)";
                        result = cercaInList(this.regione, mc_elettrici);
                        serializzaXml(annoCorrente, trimestre, trasferito, $"{trasferito}_magazzino", mc_elettrici);
                        if (result != null)
                        {
                            estratti = new Final();
                            estratti.Fornitore = nomeFornitore;
                            estratti.Valore = result.Venduto.ToString("C", CultureInfo.CurrentCulture);
                            estratti.ValoreEuro = result.Venduto;
                            finale.Add(estratti);

                        }
                        mc_elettrici = elaboraMcElettrici(pathTrasferito, elencoFiles, 2);
                        nomeFornitore = "MC ELETTRICI (Consegna diretta)";
                        result = cercaInList(this.regione, mc_elettrici);
                        serializzaXml(annoCorrente, trimestre, trasferito, $"{trasferito}_consegna_diretta", mc_elettrici);
                        break;

                    case "barcella": //
                        barcella = elaboraBarcella(pathTrasferito, elencoFiles);
                        nomeFornitore = "BARCELLA";
                        result = cercaInList(this.regione, barcella);
                        serializzaXml(annoCorrente, trimestre, trasferito, trasferito, barcella);
                        break;


                    case "comoli": //
                        nomeFornitore = "COMOLI";
                        comoli = elaboraComoli(pathTrasferito, elencoFiles);
                        result = cercaInList(this.regione, comoli);
                        serializzaXml(annoCorrente, trimestre, trasferito, trasferito, comoli);
                        break;

                    case "edif": //
                        nomeFornitore = "EDIF";
                        edif = elaboraEdif(pathTrasferito, elencoFiles);
                        result = cercaInList(this.regione, edif);
                        serializzaXml(annoCorrente, trimestre, trasferito, trasferito, edif);
                        break;


                    case "rexel": //
                        nomeFornitore = "REXEL";
                        rexel = elaboraRexel(pathTrasferito, elencoFiles, annoCorrente, trimestre);
                        result = cercaInList(this.regione, rexel);
                        serializzaXml(annoCorrente, trimestre, trasferito, trasferito, rexel);
                        break;

                    case "sonepar": //
                        nomeFornitore = "SONEPAR";
                        sonepar = elaboraSonepar(pathTrasferito, elencoFiles);
                        result = cercaInList(this.regione, sonepar);
                        serializzaXml(annoCorrente, trimestre, trasferito, trasferito, sonepar);
                        break;

                    case "sacchi": //
                        nomeFornitore = "SACCHI";
                        sacchi = elaboraSacchi(pathTrasferito, elencoFiles);
                        result = cercaInList(this.regione, sacchi);
                        serializzaXml(annoCorrente, trimestre, trasferito, trasferito, sacchi);
                        break;
                }

                if (result != null)
                {
                    estratti = new Final();
                    estratti.Fornitore = nomeFornitore;
                    estratti.Valore = result.Venduto.ToString("C", CultureInfo.CurrentCulture);
                    estratti.ValoreEuro = result.Venduto;
                    finale.Add(estratti);
                }

                result = null;


            }

            _trasferiti = finale;

        }


        private void serializzaXml(string annoCorrente, string trimestre, string trasferito, string trasferitoName, List<Trasferito> sellout)
        {
            string path = $"../trasferiti/{annoCorrente}/{trimestre}/{trasferito}/{trasferitoName}.xml";

            XmlSerializer xmls = new XmlSerializer(typeof(List<Trasferito>));

            using (TextWriter writer = new StreamWriter(path))
            {
                xmls.Serialize(writer, sellout);
            }
        }

        private Trasferito cercaInList(List<Regione> regioni, List<Trasferito> daExcel)
        {
            Trasferito res = null;

            foreach (Regione regione in regioni)
            {
                string regioneNome = regione.Nome;

                var result = daExcel.Find(x => x.Regione == regioneNome);

                if (result != null)
                {
                    if (res == null)
                    {
                        res = new Trasferito();
                    }

                    res.Regione = result.Regione;
                    res.Venduto += result.Venduto;
                }
            }


            return res;
        }

        private List<Trasferito> elaboraAcmei(string path, List<string> elencoFiles)
        {
            List<Trasferito> acmei = new List<Trasferito>();

            AcmeiStatoLettura statoLettura = AcmeiStatoLettura.Init;

            if (elencoFiles.Count == 0)
            {
                return acmei;
            }
            // apri 
            string file = elencoFiles[0];
            var wb = new XLWorkbook($"{path}/{file}");
            var ws = wb.Worksheet(1);

            int colonnaRegione = 0;
            int rigaRegione = 0;
            int colonnaVenduto = 0;



            List<Citta> citta = null;
            //2025-04-22
            List<string> nuoveCitta = new List<string>();
            //apri xml citta
            XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
            using (TextReader tr = new StreamReader("citta_acmei.xml"))
            {
                citta = (List<Citta>)xmlsd.Deserialize(tr);
            }

            var regione = "";


            Trasferito ac = null;
            bool leggiVenduto = false;
            bool readData = false;
            foreach (var row in ws.Rows())
            {
                foreach (var cell in row.Cells())
                {
                    if (!cell.IsEmpty())
                    {
                        int r = cell.Address.RowNumber;
                        int c = cell.Address.ColumnNumber;

                        string cellValue = cell.Value.ToString().Trim(' ');

                        switch (statoLettura)
                        {
                            case AcmeiStatoLettura.Init:
                                if (cellValue.Contains("Regione Nazione"))
                                {
                                    colonnaRegione = c;
                                    rigaRegione = r + 1;
                                    colonnaVenduto = c + 2;
                                    ac = new Trasferito();
                                    statoLettura = AcmeiStatoLettura.LeggiNomeRegione;
                                }
                                break;

                            case AcmeiStatoLettura.LeggiNomeRegione:
                                regione = ws.Cell(r, colonnaRegione).Value.ToString();

                                statoLettura = AcmeiStatoLettura.AttendiFine;

                                var regioneProvincia = citta.Find(x => x.Comune == regione.ToString());
                                //2025-04-22
                                if (regioneProvincia == null && regione != string.Empty)
                                {
                                    nuoveCitta.Add(regione.ToString());

                                    continue;
                                }




                                break;

                            case AcmeiStatoLettura.AttendiFine:
                                if (cellValue == $"{regione} Totale")
                                {

                                    regioneProvincia = citta.Find(x => x.Comune == regione.ToString());
                                    regione = regioneProvincia.Regione;
                                    var venduto = ws.Cell(r, colonnaVenduto).Value.ToString();
                                    acmei.Add(new Trasferito() { Regione = regione, Venduto = Double.Parse(venduto.ToString()) });
                                    //ac = new Trasferito();
                                    //rigaRegione = r + 1;
                                    //leggiVenduto = false;
                                    statoLettura = AcmeiStatoLettura.LeggiNomeRegione;
                                }
                                break;

                        }


                        if (statoLettura != AcmeiStatoLettura.Init)
                        {
                            break;
                        }

                        //   break;
                        continue;

                        if (readData == true)
                        {
                            if (cellValue == $"{ac.Regione} Totale")
                            {
                                acmei.Add(ac);
                                ac = new Trasferito();
                                rigaRegione = r + 1;
                                leggiVenduto = false;
                            }
                        }



                        if (r == rigaRegione && c == colonnaRegione)
                        {


                            var regioneProvincia = citta.Find(x => x.Comune == cellValue);
                            //2025-04-22
                            if (regioneProvincia == null)
                            {
                                nuoveCitta.Add(cellValue);

                                continue;
                            }


                            ac.Regione = cellValue;
                            leggiVenduto = true;
                        }

                        if (leggiVenduto && c == colonnaVenduto)
                        {
                            ac.Venduto += Double.Parse(cellValue);
                        }



                    }
                    else
                    {

                    }
                }
            }

            acmei.Remove(acmei.SingleOrDefault(x => x.Venduto == 0));
            aggiornaExcelCitta(nuoveCitta, "citta_acmei.xlsx");
            return acmei;
        }

        private List<Trasferito> elaboraMcElettrici(string path, List<string> elencoFiles, int scheda)
        {
            List<Trasferito> mc_elettrici = new List<Trasferito>();
            if (elencoFiles.Count == 0)
            {
                return mc_elettrici;
            }

            //List<string> campania = new List<string>();

            List<Citta> citta = null;
            //2025-04-22
            List<string> nuoveCitta = new List<string>();

            if (elencoFiles.Count == 0)
            {
                return mc_elettrici;
            }


            CultureInfo culture = null;
            culture = CultureInfo.CreateSpecificCulture("fr-FR");

            //apri xml citta
            XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
            using (TextReader tr = new StreamReader("citta_mc_elettrici.xml"))
            {
                citta = (List<Citta>)xmlsd.Deserialize(tr);
            }

            // apri 
            string file = elencoFiles[0];
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
                    if (!cell.IsEmpty())
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

                        if (r >= rigaDati)
                        {

                            venduto = ws.Cell(r, colonnaVenduto).Value.ToString().Trim(' ');
                            nomeRegione = ws.Cell(r, colonnaRegione).Value.ToString().Trim(' ');

                            var regioneProvincia = citta.Find(x => x.Comune == nomeRegione);

                            //if(regioneProvincia.Regione == "CAMPANIA")
                            //{
                            //    campania.Add(venduto);
                            //}

                            //2025-04-22
                            if (regioneProvincia == null)
                            {
                                nuoveCitta.Add(nomeRegione);

                                continue;
                            }

                            var result = mc_elettrici.Find(x => x.Regione == regioneProvincia.Regione);

                            if (result == null)
                            {
                                mc_elettrici.Add(new Trasferito() { Regione = regioneProvincia.Regione });
                            }

                            result = mc_elettrici.Find(x => x.Regione == regioneProvincia.Regione);
                            result.Venduto += Double.Parse(venduto, culture) * scontato; // sconto 10% fisso 

                            break;
                        }
                    }
                }
            }

            mc_elettrici.Remove(mc_elettrici.SingleOrDefault(x => x.Venduto == 0));
            aggiornaExcelCitta(nuoveCitta, "citta_mc_elettrici.xlsx");
            return mc_elettrici;
        }


        private List<Trasferito> elaboraMeb(string path, List<string> elencoFiles, string trimestre)
        {
            List<Trasferito> meb = new List<Trasferito>();

            if (elencoFiles.Count == 0)
            {
                return meb;
            }
            // apri 
            string file = elencoFiles[0];
            var wb = new XLWorkbook($"{path}/{file}");
            var ws = wb.Worksheet(1);

            int colonnaRegione = 0;
            int rigaRegione = 1000;
            List<int> colonnaVenduto = new List<int>();


            List<Citta> citta = null;
            //2025-04-22
            List<string> nuoveCitta = new List<string>();
            //apri xml citta
            XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
            using (TextReader tr = new StreamReader("citta_meb.xml"))
            {
                citta = (List<Citta>)xmlsd.Deserialize(tr);
            }

            Trasferito mb = null;

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
                        if (cellValue.Contains("Regione") && r == 1)
                        {
                            //result = meb.Find(x => x.Regione == "xy");

                            colonnaRegione = c;
                            rigaRegione = r + 1;
                            switch (trimestre)
                            {
                                case "t_1":
                                    colonnaVenduto.Add(6 + c);
                                    colonnaVenduto.Add(7 + c);
                                    colonnaVenduto.Add(8 + c);
                                    break;
                                case "t_2":
                                    colonnaVenduto.Add(9 + c);
                                    colonnaVenduto.Add(10 + c);
                                    colonnaVenduto.Add(11 + c);
                                    break;
                                case "t_3":
                                    colonnaVenduto.Add(12 + c);
                                    colonnaVenduto.Add(13 + c);
                                    colonnaVenduto.Add(14 + c);
                                    break;
                                case "t_4":
                                    colonnaVenduto.Add(15 + c);
                                    colonnaVenduto.Add(16 + c);
                                    colonnaVenduto.Add(17 + c);
                                    break;

                            }

                            break;

                        }

                        if (r >= rigaRegione)
                        {

                            var reg = ws.Cell(r, colonnaRegione).Value;

                            var regioneProvincia = citta.Find(x => x.Comune == reg.ToString());
                            //2025-04-22
                            if (regioneProvincia == null)
                            {
                                nuoveCitta.Add(reg.ToString());

                                continue;
                            }


                            var result = meb.Find(x => x.Regione == regioneProvincia.Regione);

                            nomeRegione = regioneProvincia.Regione;
                            if (result == null)
                            {

                                mb = new Trasferito();

                                mb.Regione = nomeRegione;
                                meb.Add(mb);
                            }

                            result = meb.Find(x => x.Regione == nomeRegione);

                            colonnaVenduto.ForEach((colvenduto) =>
                            {

                                var venduto = ws.Cell(r, colvenduto).Value.ToString();

                                result.Venduto += Double.Parse(venduto);
                            });

                            break;

                        }

                        //if (r >= rigaRegione && c == colonnaVenduto)
                        //{
                        //    var result = meb.Find(x => x.Regione == nomeRegione);
                        //    result.Venduto += Double.Parse(cellValue);
                        //}

                    }
                    else
                    {

                    }
                }
            }

            meb.Remove(meb.SingleOrDefault(x => x.Venduto == 0));
            aggiornaExcelCitta(nuoveCitta, "citta_meb.xlsx");
            return meb;
        }



        private List<Trasferito> elaboraBarcella(string path, List<string> elencoFiles)
        {
            List<Trasferito> barcella = new List<Trasferito>();

            List<Citta> citta = new List<Citta>();

            //2025-04-22
            List<string> nuoveCitta = new List<string>();

            if (elencoFiles.Count == 0)
            {
                return barcella;
            }

            // apri 
            string file = elencoFiles[0];
            var wb = new XLWorkbook($"{path}/{file}");

            //apri xml citta
            XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
            using (TextReader tr = new StreamReader("citta_barcella.xml"))
            {
                citta = (List<Citta>)xmlsd.Deserialize(tr);
            }

            for (int i = 1; i <= 3; i++)
            {
                var ws = wb.Worksheet(i);

                int colonnaRegione = 0;
                int rigaRegione = 0;
                int colonnaVenduto = 0;


                Trasferito brc = null;

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
                                    nuoveCitta.Add(cellValue);
                                    break;
                                    //continue;
                                }

                                string regione = regioneCitta.Regione;

                                var result = barcella.Find(x => x.Regione == regione);
                                nomeRegione = regione;

                                if (result == null)
                                {

                                    brc = new Trasferito();

                                    brc.Regione = nomeRegione;
                                    barcella.Add(brc);
                                }



                                result = barcella.Find(x => x.Regione == nomeRegione);
                                result.Venduto += Double.Parse(ws.Cell(r, colonnaVenduto).Value.ToString().Trim(' '));
                                break;
                                //result = meb.Find(x => x.Regione == cellValue);
                            }

                            //if (r >= rigaRegione && c == colonnaVenduto)
                            //{
                            //    var result = barcella.Find(x => x.Regione == nomeRegione);
                            //    result.Venduto += Double.Parse(cellValue);
                            //}

                        }
                        else
                        {

                        }
                    }
                }
            }

            aggiornaExcelCitta(nuoveCitta, "citta_barcella.xlsx");

            barcella.Remove(barcella.SingleOrDefault(x => x.Venduto == 0));

            return barcella;
        }



        private List<Trasferito> elaboraComoli(string path, List<string> elencoFiles)
        {


            List<Trasferito> comoli = new List<Trasferito>();

            if (elencoFiles.Count == 0)
            {
                return comoli;
            }
            // apri 
            string file = elencoFiles[0];
            var wb = new XLWorkbook($"{path}/{file}");
            var ws = wb.Worksheet(1);

            List<Citta> citta = new List<Citta>();
            //2025-04-22
            List<string> nuoveCitta = new List<string>();

            int colonnaProvincia = 0;
            int rigaProvincia = 0;
            int colonnaVenduto = 0;

            CultureInfo culture = null;
            culture = CultureInfo.CreateSpecificCulture("fr-FR");

            //apri xml citta
            XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
            using (TextReader tr = new StreamReader("citta_comoli.xml"))
            {
                citta = (List<Citta>)xmlsd.Deserialize(tr);
            }

            Trasferito cml = null;

            string nomeProvincia = string.Empty;

            var statoLettura = ComoliStatoLettura.Init;

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
                            //result = meb.Find(x => x.Regione == "xy");

                            colonnaProvincia = c;

                            //mb = new Trasferito();

                        }

                        if (cellValue.Contains("Fatt al Costo"))
                        {
                            colonnaVenduto = c;
                            statoLettura = ComoliStatoLettura.AttendiDati;
                        }

                        if (c == colonnaVenduto)
                        {
                            switch (statoLettura)
                            {
                                case ComoliStatoLettura.AttendiDati:
                                    statoLettura = ComoliStatoLettura.LeggiValori;
                                    break;
                                case ComoliStatoLettura.LeggiValori:

                                    var venduto = ws.Cell(r, colonnaVenduto).Value.ToString().Trim(' ');

                                    var provincia = ws.Cell(r, colonnaProvincia).Value.ToString().Trim(' ');



                                    var regioneProvincia = citta.Find(x => x.Comune == provincia);


                                    //2025-04-22
                                    if (regioneProvincia == null)
                                    {
                                        nuoveCitta.Add(provincia);

                                        continue;
                                    }


                                    var result = comoli.Find(x => x.Regione == regioneProvincia.Regione);


                                    if (result == null)
                                    {




                                        cml = new Trasferito();

                                        cml.Regione = regioneProvincia.Regione;
                                        comoli.Add(cml);
                                    }

                                    result = comoli.Find(x => x.Regione == regioneProvincia.Regione);

                                    result.Venduto += Double.Parse(venduto, culture);

                                    break;
                            }
                        }
                    }
                    else
                    {

                    }
                }
            }

            comoli.Remove(comoli.SingleOrDefault(x => x.Venduto == 0));
            aggiornaExcelCitta(nuoveCitta, "citta_comoli.xlsx");
            return comoli;
        }




        private List<Trasferito> elaboraEdif(string path, List<string> elencoFiles)
        {
            List<Trasferito> edif = new List<Trasferito>();
            List<IndiceRegione> IndiceRegione = null;
            if (elencoFiles.Count == 0)
            {
                return edif;
            }
            // apri 
            string file = elencoFiles[0];
            var wb = new XLWorkbook($"{path}/{file}");
            var ws = wb.Worksheet(1);



            List<Citta> citta = null;
            //2025-04-22
            List<string> nuoveCitta = new List<string>();
            //apri xml citta
            XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
            using (TextReader tr = new StreamReader("citta_edif.xml"))
            {
                citta = (List<Citta>)xmlsd.Deserialize(tr);
            }


            int colonnaArticoli = 0;

            Trasferito edf = null;

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
                            IndiceRegione = new List<IndiceRegione>();
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
                                }
                                else
                                {
                                    edf = new Trasferito();
                                    nomeRegione = cellValue;

                                    //if (nomeRegione == "ROMAGNA")
                                    //{
                                    //    nomeRegione = "EMILIA ROMAGNA";
                                    //}

                                    IndiceRegione.Add(new IndiceRegione() { Index = c, Regione = nomeRegione });


                                    var regioneProvincia = citta.Find(x => x.Comune == nomeRegione);
                                    //2025-04-22
                                    if (regioneProvincia == null)
                                    {
                                        nuoveCitta.Add(nomeRegione);

                                        continue;
                                    }


                                    edf.Regione = regioneProvincia.Regione;
                                    edif.Add(edf);
                                }
                                break;

                            case EdifStatoLettura.LeggiValori:
                                if (c == colonnaArticoli)
                                {
                                    if (cellValue == "TOTALE")
                                    {
                                        statoLettura = EdifStatoLettura.Exit;
                                    }
                                }
                                else
                                {
                                    var indexRes = IndiceRegione.Find(x => x.Index == c);
                                    if (indexRes != null)
                                    {
                                        var regioneProvincia = citta.Find(x => x.Comune == indexRes.Regione);

                                        var result = edif.Find(x => x.Regione == regioneProvincia.Regione);
                                        result.Venduto += Double.Parse(cellValue);
                                    }
                                }

                                break;

                            case EdifStatoLettura.Exit:

                                break;

                        }



                    }
                    else
                    {

                    }
                }
            }

            edif.Remove(edif.SingleOrDefault(x => x.Venduto == 0));
            aggiornaExcelCitta(nuoveCitta, "citta_edif.xlsx");
            return edif;
        }




        private List<Trasferito> elaboraRexel(string path, List<string> elencoFiles, string annoCorrente, string trimestre)
        {
            List<Trasferito> rexel = new List<Trasferito>();

            List<Citta> citta = new List<Citta>();

            //2025-04-22
            List<string> nuoveCitta = new List<string>();

            if (elencoFiles.Count == 0)
            {
                return rexel;
            }
            // apri 
            string file = elencoFiles[0];
            var wb = new XLWorkbook($"{path}/{file}");
            var ws = wb.Worksheet(1);


            int colonnaComune = 0;
            int colonnaAnno = 0;
            int rowRead = 0;

            Trasferito rxl = null;

            string nomeRegione = string.Empty;

            var statoLettura = RexelStatoLettura.Init;
            int m1 = 0;
            int m2 = 0;
            int m3 = 0;

            //apri xml citta
            XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
            using (TextReader tr = new StreamReader("citta_rexel.xml"))
            {
                citta = (List<Citta>)xmlsd.Deserialize(tr);
            }


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

                                switch (trimestre)
                                {
                                    case "t_1":
                                        m1 = colonnaComune + 2;
                                        m2 = colonnaComune + 3;
                                        m3 = colonnaComune + 4;
                                        break;

                                    case "t_2":
                                        m1 = colonnaComune + 5;
                                        m2 = colonnaComune + 6;
                                        m3 = colonnaComune + 7;
                                        break;

                                    case "t_3":
                                        m1 = colonnaComune + 8;
                                        m2 = colonnaComune + 9;
                                        m3 = colonnaComune + 10;
                                        break;

                                    case "t_4":
                                        break;
                                        m1 = colonnaComune + 11;
                                        m2 = colonnaComune + 12;
                                        m3 = colonnaComune + 13;
                                }


                                statoLettura = RexelStatoLettura.LeggiValori;

                                break;

                            case RexelStatoLettura.LeggiValori:
                                if (r >= rowRead)
                                {
                                    if ((c == colonnaComune) && (ws.Cell(r, colonnaAnno).Value.ToString() == annoCorrente))
                                    {
                                        //var s = cellValue.Split('(');
                                        //var sSearch = s[0].ToString().Trim(' ');
                                        var result = citta.Find(x => x.Comune == cellValue);


                                        //2025-04-22
                                        if (result == null)
                                        {
                                            nuoveCitta.Add(cellValue);

                                            continue;
                                        }

                                        if (result == null)
                                        {
                                            continue;
                                        }


                                        var result_2 = rexel.Find(x => x.Regione == result.Regione);
                                        if (result_2 == null)
                                        {
                                            rxl = new Trasferito();
                                            rxl.Regione = result.Regione;
                                            rexel.Add(rxl);
                                        }

                                        result_2 = rexel.Find(x => x.Regione == result.Regione);

                                        result_2.Venduto += Double.Parse(ws.Cell(r, m1).Value.ToString());
                                        result_2.Venduto += Double.Parse(ws.Cell(r, m2).Value.ToString());
                                        result_2.Venduto += Double.Parse(ws.Cell(r, m3).Value.ToString());
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

            rexel.Remove(rexel.SingleOrDefault(x => x.Venduto == 0));
            aggiornaExcelCitta(nuoveCitta, "citta_rexel.xlsx");
            return rexel;
        }





        private List<Trasferito> elaboraSonepar(string path, List<string> elencoFiles)
        {
            List<Trasferito> sonepar = new List<Trasferito>();

            if (elencoFiles.Count == 0)
            {
                return sonepar;
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



            // apri 
            string file = elencoFiles[0];
            var wb = new XLWorkbook($"{path}/{file}");

            for (int scheda = 1; scheda <= 3; scheda++)
            {
                var ws = wb.Worksheet(scheda);

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


                                        var result = sonepar.Find(x => x.Regione == regioneProvincia.Regione);

                                        if (result == null)
                                        {
                                            sonepar.Add(new Trasferito() { Regione = regioneProvincia.Regione });
                                        }

                                        result = sonepar.Find(x => x.Regione == regioneProvincia.Regione);
                                        result.Venduto += Double.Parse(ws.Cell(r, colonnaValore).Value.ToString());


                                    }

                                    break;
                            }

                        }
                        else
                        {

                        }
                    }
                }
            }

            sonepar.Remove(sonepar.SingleOrDefault(x => x.Venduto == 0));
            aggiornaExcelCitta(nuoveCitta, "citta_sonepar.xlsx");
            return sonepar;
        }


        private List<Trasferito> elaboraSacchi(string path, List<string> elencoFiles)
        {
            List<Trasferito> sacchi = new List<Trasferito>();
            List<Citta> citta = new List<Citta>();
            //2025-04-22
            List<string> nuoveCitta = new List<string>();

            if (elencoFiles.Count == 0)
            {
                return sacchi;
            }
            // apri 
            string file = elencoFiles[0];
            var wb = new XLWorkbook($"{path}/{file}");


            //apri xml citta
            XmlSerializer xmlsd = new XmlSerializer(typeof(List<Citta>));
            using (TextReader tr = new StreamReader("citta_sacchi.xml"))
            {
                citta = (List<Citta>)xmlsd.Deserialize(tr);
            }

            // for (int scheda = 1; scheda <= 3; scheda++)
            //{
            string scheda = "Table";


            var ws = wb.Worksheet(scheda);

            int lastRow = ws.Column("N").CellsUsed().Count();

            int colonnaRegione = 0;
            int colonnaValore = 0;
            int rowRead = 0;

            int riga = 0;
            int colonna = 0;

            // Trasferito snp = null;

            string nomeRegione = string.Empty;

            var statoLettura = SacchiStatoLettura.Init;

            foreach (var row in ws.Rows())
            {
                foreach (var cell in row.Cells())
                {
                    if (!cell.IsEmpty())
                    {
                        int r = cell.Address.RowNumber;
                        int c = cell.Address.ColumnNumber;

                        riga = r;
                        colonna = c;
                        string cellValue = cell.Value.ToString().Trim(' ');

                        //if (cellValue.Contains("Regione"))
                        if (cellValue.IndexOf("Regione") == 0)
                        {
                            if (c > colonnaRegione)
                                colonnaRegione = c + 1;
                        }


                        if (cellValue.Contains("Costo Finito A.C."))
                        {

                            switch (colonnaValore)
                            {
                                case 0:
                                    colonnaValore = c;
                                    Console.WriteLine(cell.WorksheetColumn().ColumnLetter());
                                    break;

                            }


                        }



                        switch (statoLettura)
                        {
                            case SacchiStatoLettura.Init:

                                break;

                            case SacchiStatoLettura.LeggiValori:

                                if (c == colonnaRegione)
                                {

                                    var regioneRegione = citta.Find(x => x.Comune == cellValue);

                                    //2025-04-22
                                    if (regioneRegione == null)
                                    {
                                        nuoveCitta.Add(cellValue);

                                        continue;
                                    }

                                    var result = sacchi.Find(x => x.Regione == regioneRegione.Regione);
                                    if (result == null)
                                    {
                                        sacchi.Add(new Trasferito() { Regione = regioneRegione.Regione });
                                    }
                                    result = sacchi.Find(x => x.Regione == regioneRegione.Regione);

                                    if (ws.Cell(r, colonnaValore).Value.ToString().Trim(' ') == string.Empty)
                                    {
                                        result.Venduto += 0;
                                    }
                                    else
                                    {
                                        result.Venduto += Double.Parse(ws.Cell(r, colonnaValore).Value.ToString());
                                    }


                                }

                                break;
                        }

                    }
                    else
                    {

                    }
                }

                if (colonnaRegione > 0)
                {
                    statoLettura = SacchiStatoLettura.LeggiValori;
                }

            }
            //}

            aggiornaExcelCitta(nuoveCitta, "citta_sacchi.xlsx");
            sacchi.Remove(sacchi.SingleOrDefault(x => x.Venduto == 0));
            return sacchi;
        }


        private void aggiornaExcelCitta(List<string> nuoveCitta, string fileExcel)
        {
            if (nuoveCitta.Count == 0)
            {
                return;
            }


            List<string> noDupes = nuoveCitta.Distinct().ToList();

            string path = $"../citta_regione/{fileExcel}";

            var wb = new XLWorkbook(path);
            var ws = wb.Worksheet("comuni");

            int lastRow = ws.Column("A").CellsUsed().Count();

            lastRow++;

            noDupes.ForEach((val) =>
            {

                Console.WriteLine("ciao");
                ws.Cell(lastRow, 1).Value = val;
                lastRow++;
            });

            wb.Save();

            Process.Start("explorer.exe", System.IO.Path.GetFullPath($"{path}"));

        }

        public IList<Final> Trasferiti => _trasferiti;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}