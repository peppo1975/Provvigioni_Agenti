using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2019.Drawing.Animation;
using DocumentFormat.OpenXml.Spreadsheet;
using Provvigioni_Agenti.Models;
using System.Collections.Generic;
using System.Diagnostics;
using System.Formats.Tar;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml.Serialization;

namespace Provvigioni_Agenti.Controllers
{
    internal class General
    {
        static public string valuta(double valore)
        {
            string res = string.Empty;
            res = valore.ToString("C", CultureInfo.CurrentCulture);
            return res;
        }

        static public string percentuale(double valore)
        {
            string res = string.Empty;
            res = Math.Round(valore * 100, 2).ToString("N2") + "%";
            return res;
        }


        static public void coloraVariazioni(double valoreCorrente, double valoreRiferimento, TextBlock textBlock)
        {
            textBlock.Background = Brushes.GreenYellow;
            textBlock.Foreground = Brushes.Black;
            if ((valoreRiferimento - valoreCorrente) > 0)
            {
                textBlock.Background = Brushes.Red;
                textBlock.Foreground = Brushes.White;
            }

            if ((valoreRiferimento - valoreCorrente) == 0)
            {
                textBlock.Background = null;
            }
        }


        static public List<string> directoryTrasferiti(string annoCorrente)
        {
            DateTime localDate = DateTime.Now;

            //string year = localDate.Year.ToString();
            string year = annoCorrente;
            //leggo il file agenti.xml
            AgentiService agentiService = new Models.AgentiService();

            List<string> trasferiti = new List<string>();

            trasferiti = Models.TrasferitiAgenzie.ToArray();

            List<string> mesi = new List<string>(); // { "m_01", "m_02", "m_03", "m_04", "m_05", "m_06", "m_07", "m_08", "m_09", "m_10", "m_11", "m_12" };

            mesi = Models.Mesi.toArray();

            string path = "../trasferiti";

            if (!File.Exists(path))
            {
                DirectoryInfo di = Directory.CreateDirectory(path);

            }

            string pathYear = $"{path}/{year}";

            if (!File.Exists(pathYear))
            {
                DirectoryInfo di = Directory.CreateDirectory(pathYear);

            }

            foreach (string mese in mesi)
            {
                string subPathTrimestre = $"{pathYear}/{mese}";

                if (!File.Exists(subPathTrimestre))
                {
                    DirectoryInfo diA = Directory.CreateDirectory(subPathTrimestre);
                }

                foreach (string trasferito in trasferiti)
                {
                    string subPath = $"{subPathTrimestre}/{trasferito}";
                    if (!File.Exists(subPath))
                    {
                        DirectoryInfo diA = Directory.CreateDirectory(subPath);
                    }
                }
            }

            return trasferiti;

        }


        public static List<string> elencoFile(string path)
        {
            List<string> fileEntries = new List<string>();
            try
            {
                var fullPath = System.IO.Path.GetFullPath(path);

                DirectoryInfo di = new DirectoryInfo(fullPath);
                Console.WriteLine("No search pattern returns:");
                foreach (var fi in di.GetFiles())
                {
                    Console.WriteLine(fi.Name);
                    fileEntries.Add(fi.Name.ToString());
                }

            }
            catch (Exception e)
            {
            }

            return fileEntries;
        }


        // that are found, and process the files they contain.
        public static void ProcessDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
            {
                ProcessFile(fileName);
            }

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
            {
                ProcessDirectory(subdirectory);
            }

        }

        // Insert logic for processing found files here.
        public static void ProcessFile(string path)
        {
            Console.WriteLine("Processed file '{0}'.", path);
        }


        public static void leggiAgenti()
        {
            var wb = new XLWorkbook($"../agenti/agenti.xlsx");
            var ws = wb.Worksheet("agenti");

            List<Agente> list = new List<Agente>();
            Agente a = null;
            Agente all = new Agente();
            List<string> allId = new List<string>();
            bool rigaValida = false;

            foreach (var row in ws.Rows())
            {
                if (row.RowNumber() > 1)
                {
                    rigaValida = false;
                    a = new Agente();
                }
                foreach (var cell in row.Cells())
                {
                    if (!cell.IsEmpty())
                    {
                        int r = cell.Address.RowNumber;
                        int c = cell.Address.ColumnNumber;

                        if (r > 1)
                        {
                            string cellValue = cell.Value.ToString().Trim(' ');
                            rigaValida = true;


                            switch (c)
                            {
                                case 1: // id_agente	
                                    a.ID = cellValue;
                                    allId.Add(cellValue);
                                    break;

                                case 2: // nome_lista	
                                    a.NikName = cellValue;
                                    break;

                                case 3: // nome_completo
                                    a.Nome = cellValue;
                                    break;

                                default: // regioni
                                    //Regione reg = 
                                    a.Regione.Add(new Regione() { Nome = cellValue });
                                    break;

                            }
                        }

                    }
                }


                if (row.RowNumber() > 1 && rigaValida == true)
                {
                    list.Add(a);
                }
            }

            list.OrderBy(o => o.NikName).ToList();


            all.ID = string.Join("#", allId);
            all.NikName = "-- TUTTI GLI AGENTI --";
            all.Nome = "TUTTI GLI AGENTI";
            all.Regione.Add(new Regione() { Nome = "ITALIA" });

            list.Add(all);

            XmlSerializer xmls = new XmlSerializer(typeof(List<Agente>));

            using (TextWriter writer = new StreamWriter(@"agenti.xml"))
            {
                xmls.Serialize(writer, list.OrderBy(o => o.NikName).ToList());
            }

        }


        public static void leggiRegioni()
        {
            var wb = new XLWorkbook($"../regioni/regioni.xlsx");
            var ws = wb.Worksheet(1);
            List<string> regioni = new List<string>();

            foreach (var row in ws.Rows())
            {
                foreach (var cell in row.Cells())
                {
                    if (!cell.IsEmpty())
                    {

                        int r = cell.Address.RowNumber;
                        int c = cell.Address.ColumnNumber;
                        string cellValue = cell.Value.ToString().Trim(' ');
                        switch (c)
                        {
                            case 1:
                                regioni.Add(cellValue);
                                break;
                        }
                    }
                }
            }

            wb = new XLWorkbook($"../agenti/agenti.xlsx");
            ws = wb.Worksheet("regioni");
            ws.Column(1).Delete(); // Range starts on C2

            regioni.Sort();
            int index = 1;
            foreach (var item in regioni)
            {
                ws.Cell($"A{index}").Value = item;
                index++;
            }


            wb.Save();

            List<string> elencoCittaRegione = elencoFile("../citta_regione");

            foreach (var item in elencoCittaRegione)
            {
                wb = new XLWorkbook($"../citta_regione/{item}");
                ws = wb.Worksheet("regioni");
                ws.Column(1).Delete(); // Range starts on C2
                index = 1;
                foreach (var itemR in regioni)
                {
                    ws.Cell($"A{index}").Value = itemR;
                    index++;
                }


                wb.Save();

            }


        }


        public static void generaXmlCitta()
        {

            List<string> cittaXml = new List<string>();

            cittaXml.Add("citta_acmei");
            cittaXml.Add("citta_barcella");
            cittaXml.Add("citta_comoli");
            cittaXml.Add("citta_edif");
            cittaXml.Add("citta_mc_elettrici");
            cittaXml.Add("citta_meb");
            cittaXml.Add("citta_rexel");
            cittaXml.Add("citta_sacchi");
            cittaXml.Add("citta_sonepar");

            cittaXml.ForEach((x) =>
            {
                generaCittaXml(x);
            });

        }


        private static void generaCittaXml(string fileXlsx)
        {
            var wb = new XLWorkbook($"../citta_regione/{fileXlsx}.xlsx");
            var ws = wb.Worksheet(1);

            List<Citta> citta = new List<Citta>();
            Citta comune = null;



            foreach (var row in ws.Rows())
            {
                foreach (var cell in row.Cells())
                {
                    if (!cell.IsEmpty())
                    {

                        int r = cell.Address.RowNumber;
                        int c = cell.Address.ColumnNumber;
                        string cellValue = cell.Value.ToString().Trim(' ');
                        switch (c)
                        {
                            case 1:
                                comune = new Citta();
                                comune.Comune = cellValue;
                                break;

                            case 2:
                                comune.Regione = cellValue;
                                citta.Add(comune);
                                comune = null;
                                break;
                        }
                    }
                }
            }



            XmlSerializer xmls = new XmlSerializer(typeof(List<Citta>));

            using (TextWriter writer = new StreamWriter($"{fileXlsx}.xml"))
            {
                xmls.Serialize(writer, citta);
            }
        }



        public static void generaExcelTrasferiti(Agente agente, string annoCorrente, string annoRiferimento, List<string> mese, IList<ClienteResponse> clienteResponse, IList<Final> Trasferiti, List<GruppoStatistico> GruppoStatistico, List<GruppoStatisticoRiepilogo> GruppoStatisticoProgressivo, List<GruppoStatisticoRiepilogo> GruppoStatisticoTrimestre)
        {

            var esplPer = esplicitaPeriodo(mese);


            IList<ClienteResponse> clienteResponseFiltered = clienteResponse.Where(x => x.TotaleVendutoCorrente > 0 || x.TotaleVendutoRiferimento > 0).ToList();

            string path = "../excelAgenti";
            string pathFile = $"{path}/{annoCorrente}-{esplPer["excelNomeFile"]}--{agente.NikName}.xlsx";
            string fullPath = System.IO.Path.GetFullPath(pathFile);
            int index = 9;

            int indexRowProvvigioneAgente = 0;

            int indexRowProvvigioneAgenteSellout = 0;

            if (!File.Exists(path))
            {
                DirectoryInfo di = Directory.CreateDirectory(path);

            }

            using var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Provvigioni");

            int rowTotDb = provvigioniPassPartout(workbook, index, agente.NikName, agente.Nome, annoCorrente, annoRiferimento, esplPer["excel"], clienteResponse);

            int rowTotSellOut = provvigioniSellout(workbook, rowTotDb, Trasferiti);

            int rowProgressivo = gruppiStatistici(workbook, rowTotSellOut, annoCorrente, annoRiferimento, clienteResponse, GruppoStatistico, GruppoStatisticoProgressivo, "PROGRESSIVO");

            int rowTrimestre = gruppiStatistici(workbook, rowProgressivo, annoCorrente, annoRiferimento, clienteResponse, GruppoStatistico, GruppoStatisticoTrimestre, esplPer["excel"]);

            int rowGrStClienteProgr = gruppiStatisticiClienti(workbook, rowTrimestre, GruppoStatistico, GruppoStatisticoProgressivo, clienteResponse, "PROGRESSIVO");

            int rowGrStClienteTrim = gruppiStatisticiClienti(workbook, rowGrStClienteProgr, GruppoStatistico, GruppoStatisticoTrimestre, clienteResponse, esplPer["excel"]);


            worksheet.Cell("G4").FormulaA1 = $"G{rowTotDb}+G{rowTotSellOut}";
            worksheet.Cell("G4").Style.NumberFormat.Format = "#,##0.00 €";
            worksheet.Cell("G4").Style.Font.Bold = true;
            worksheet.Cell("G4").Style.Font.FontSize = 15;


            workbook.SaveAs(pathFile);

            //Process.Start("explorer.exe", System.IO.Path.GetFullPath($"{path}"));

            Process.Start("explorer.exe", fullPath);

            return;


        }



        private static Dictionary<string, string> esplicitaPeriodo(List<string> mese)
        {
            Dictionary<string, string> esplicitaExcel = new Dictionary<string, string>();
            Dictionary<string, string> esplicitaExcelNomeFile = new Dictionary<string, string>();
            Dictionary<string, string> result = new Dictionary<string, string>();


            string cerca = String.Join("", mese);

            esplicitaExcel.Add(Mesi.Gennaio + Mesi.Febbraio + Mesi.Marzo, "1° TRIMESTRE");
            esplicitaExcel.Add(Mesi.Aprile + Mesi.Maggio + Mesi.Giugno, "2° TRIMESTRE");
            esplicitaExcel.Add(Mesi.Luglio + Mesi.Agosto + Mesi.Settembre, "3° TRIMESTRE");
            esplicitaExcel.Add(Mesi.Ottobre + Mesi.Novembre + Mesi.Dicembre, "4° TRIMESTRE");
            esplicitaExcel.Add(Mesi.Gennaio, "GENNAIO");
            esplicitaExcel.Add(Mesi.Febbraio, "FEBBRAIO");
            esplicitaExcel.Add(Mesi.Marzo, "MARZO");
            esplicitaExcel.Add(Mesi.Aprile, "APRILE");
            esplicitaExcel.Add(Mesi.Maggio, "MAGGIO");
            esplicitaExcel.Add(Mesi.Giugno, "GIUGNO");
            esplicitaExcel.Add(Mesi.Luglio, "LUGLIO");
            esplicitaExcel.Add(Mesi.Agosto, "AGOSTO");
            esplicitaExcel.Add(Mesi.Settembre, "SETTEMBRE");
            esplicitaExcel.Add(Mesi.Ottobre, "OTTOBRE");
            esplicitaExcel.Add(Mesi.Novembre, "NOVEMBRE");
            esplicitaExcel.Add(Mesi.Dicembre, "DICEMBRE");

            esplicitaExcelNomeFile.Add(Mesi.Gennaio + Mesi.Febbraio + Mesi.Marzo, "1-TRIMESTRE");
            esplicitaExcelNomeFile.Add(Mesi.Aprile + Mesi.Maggio + Mesi.Giugno, "2-TRIMESTRE");
            esplicitaExcelNomeFile.Add(Mesi.Luglio + Mesi.Agosto + Mesi.Settembre, "3-TRIMESTRE");
            esplicitaExcelNomeFile.Add(Mesi.Ottobre + Mesi.Novembre + Mesi.Dicembre, "4-TRIMESTRE");
            esplicitaExcelNomeFile.Add(Mesi.Gennaio, "GENNAIO");
            esplicitaExcelNomeFile.Add(Mesi.Febbraio, "FEBBRAIO");
            esplicitaExcelNomeFile.Add(Mesi.Marzo, "MARZO");
            esplicitaExcelNomeFile.Add(Mesi.Aprile, "APRILE");
            esplicitaExcelNomeFile.Add(Mesi.Maggio, "MAGGIO");
            esplicitaExcelNomeFile.Add(Mesi.Giugno, "GIUGNO");
            esplicitaExcelNomeFile.Add(Mesi.Luglio, "LUGLIO");
            esplicitaExcelNomeFile.Add(Mesi.Agosto, "AGOSTO");
            esplicitaExcelNomeFile.Add(Mesi.Settembre, "SETTEMBRE");
            esplicitaExcelNomeFile.Add(Mesi.Ottobre, "OTTOBRE");
            esplicitaExcelNomeFile.Add(Mesi.Novembre, "NOVEMBRE");
            esplicitaExcelNomeFile.Add(Mesi.Dicembre, "DICEMBRE");

            result.Add("excel", esplicitaExcel[cerca]);
            result.Add("excelNomeFile", esplicitaExcelNomeFile[cerca]);

            return result;
        }


        private static int provvigioniPassPartout(XLWorkbook workbook, int indexInit, string agente, string agenteFullName, string annoCorrente, string annoRiferimento, string trimestre, IList<ClienteResponse> clienteResponse)
        {


            Dictionary<string, string> trimestriSuExcel = new Dictionary<string, string>() { { "t_1", "1° TRIM" }, { "t_2", "2° TRIM" }, { "t_3", "3° TRIM" }, { "t_4", "4° TRIM" } };

            IList<ClienteResponse> clienteResponseFiltered = clienteResponse.Where(x => x.TotaleVendutoCorrente > 0 || x.TotaleVendutoRiferimento > 0).ToList();

            int index = indexInit;


            var worksheet = workbook.Worksheet("Provvigioni");

            var imagePath = @"../logo.jpg";

            var image = worksheet.AddPicture(imagePath)
                .MoveTo(worksheet.Cell("A2"))
                .Scale(0.3); // optional: resize picture

            int indexRowProvvigioneAgente = 0;


            worksheet.Column("A").Width = 10;
            worksheet.Column("B").Width = 45;
            worksheet.Column("C").Width = 17;
            worksheet.Column("D").Width = 17;
            worksheet.Column("E").Width = 17;
            worksheet.Column("F").Width = 17;
            worksheet.Column("G").Width = 17;
            worksheet.Column("H").Width = 17;


            worksheet.Cell("C2").Value = agenteFullName;
            worksheet.Cell("C2").Style.Font.FontSize = 20;
            worksheet.Range("C2:G2").Merge();


            worksheet.Cell("C4").Value = $"{trimestre} {annoCorrente} - PROVVIGIONE TOTALE: ";
            worksheet.Cell("C4").Style.Font.FontSize = 15;
            worksheet.Range("C4:F4").Merge();
            //worksheet.Cell("H3").Value

            worksheet.Cell("A8").Value = $"Codice";
            worksheet.Cell("B8").Value = $"Descrizione";
            worksheet.Cell("C8").Value = $"Imp. periodo " + annoRiferimento;
            worksheet.Cell("D8").Value = $"Imp. periodo " + annoCorrente;
            worksheet.Cell("E8").Value = $"Delta imp.";
            worksheet.Cell("F8").Value = $"Delta imp. %";
            worksheet.Cell("G8").Value = $"Provvigione";



            worksheet.Range($"A8:H8").Style.Font.Bold = true;

            //foreach (ClienteResponse cliente in clienteResponse)
            foreach (ClienteResponse cliente in clienteResponseFiltered)
            {
                double[] result = new double[2];
                result = calcolaPercentuale(cliente.TotaleVendutoRiferimento, cliente.TotaleVendutoCorrente);

                worksheet.Cell(index, 1).Value = cliente.IdCliente;
                worksheet.Cell(index, 2).Value = cliente.NomeCliente;
                worksheet.Cell(index, 3).Value = cliente.TotaleVendutoRiferimento;
                worksheet.Cell(index, 4).Value = cliente.TotaleVendutoCorrente;
                worksheet.Cell(index, 5).Value = result[0];
                worksheet.Cell(index, 6).Value = result[1]; // percentule
                worksheet.Cell(index, 7).Value = cliente.ProvvigioneCorrente;




                worksheet.Cell(index, 3).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 4).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 5).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 6).Style.NumberFormat.Format = "0.00%";
                worksheet.Cell(index, 7).Style.NumberFormat.Format = "#,##0.00 €";



                if (result[0] < 0)
                {
                    //worksheet.Cell(index, 2).Style.Fill.BackgroundColor = XLColor.Red;
                    //worksheet.Cell(index, 5).Style.Fill.BackgroundColor = XLColor.Red;
                    //worksheet.Cell(index, 6).Style.Fill.BackgroundColor = XLColor.Red;
                    worksheet.Cell(index, 2).Style.Fill.BackgroundColor = XLColor.RedRyb;
                    worksheet.Cell(index, 5).Style.Fill.BackgroundColor = XLColor.RedRyb;
                    worksheet.Cell(index, 6).Style.Fill.BackgroundColor = XLColor.RedRyb;

                    worksheet.Cell(index, 2).Style.Font.Bold = true;
                    worksheet.Cell(index, 5).Style.Font.Bold = true;
                    worksheet.Cell(index, 6).Style.Font.Bold = true;
                }

                index++;
            }


            worksheet.Range($"A9:G{index - 1}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"A9:G{index - 1}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"A9:G{index - 1}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"A9:G{index - 1}").Style.Border.RightBorder = XLBorderStyleValues.Thin;


            //index += 1;


            worksheet.Range($"A{index}:H{index}").Style.Font.Bold = true;
            worksheet.Range($"C{index}:H{index}").Style.NumberFormat.Format = "#,##0.00 €";
            worksheet.Cell($"F{index}").Style.NumberFormat.Format = "0.00%";

            worksheet.Cell(index, 2).Value = "TOTALE"; // vendite

            worksheet.Cell(index, 3).FormulaA1 = $"SUM(C9:C{index - 1})";
            worksheet.Cell(index, 4).FormulaA1 = $"SUM(D9:D{index - 1})";
            worksheet.Cell(index, 5).FormulaA1 = $"SUM(E9:E{index - 1})";
            worksheet.Cell(index, 6).FormulaA1 = $"(D{index}-C{index})/C{index}";

            if ((double)worksheet.Cell(index, 5).Value < 0)
            {
                worksheet.Cell(index, 5).Style.Fill.BackgroundColor = XLColor.Red;
                worksheet.Cell(index, 6).Style.Fill.BackgroundColor = XLColor.Red;
            }

            worksheet.Cell(index, 7).FormulaA1 = $"SUM(G9:G{index - 1})";

            return index;

        }


        private static int provvigioniSellout(XLWorkbook workbook, int indexInit, IList<Final> Trasferiti)
        {

            var worksheet = workbook.Worksheet("Provvigioni");

            int indexSellout = 0;

            int index = indexInit;

            index += 2;

            //worksheet.Range($"A{index}:G{index}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;  // bordo

            index += 3;


            worksheet.Cell($"C{index}").Value = "SELLOUT";
            worksheet.Cell($"C{index}").Style.Font.FontSize = 20;
            worksheet.Range($"C{index}:G{index}").Merge();

            index += 2;

            indexSellout = index;
            worksheet.Cell(index, 2).Value = "SELLOUT";
            worksheet.Cell(index, 4).Value = "VENDUTO";
            worksheet.Cell(index, 7).Value = "PROVVIGIONE";
            worksheet.Range($"A{index}:G{index}").Style.Font.Bold = true;
            index++;
            foreach (Final f in Trasferiti)
            {
                if (f.Fornitore == " - - - TOTALE: ")
                {
                    continue;
                }

                worksheet.Cell(index, 2).Value = f.Fornitore;
                worksheet.Cell(index, 4).Value = f.ValoreEuro;
                worksheet.Cell(index, 4).Style.NumberFormat.Format = "#,##0.00 €";

                worksheet.Cell(index, 7).Value = f.ValoreEuro * 0.02;
                worksheet.Cell(index, 7).Style.NumberFormat.Format = "#,##0.00 €";

                index++;
            }



            worksheet.Range($"B{indexSellout}:G{index - 1}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"B{indexSellout}:G{index - 1}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"B{indexSellout}:G{index - 1}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"B{indexSellout}:G{index - 1}").Style.Border.RightBorder = XLBorderStyleValues.Thin;


            worksheet.Cell(index, 2).Value = "TOTALE"; // sellout
            worksheet.Cell(index, 2).Style.Font.Bold = true;

            worksheet.Cell(index, 4).FormulaA1 = $"SUM(D{indexSellout}:D{index - 1})";
            worksheet.Cell(index, 4).Style.Font.Bold = true;
            worksheet.Cell(index, 4).Style.NumberFormat.Format = "#,##0.00 €";

            //worksheet.Cell(index, 6).Value = "TOTALE"; // sellout provvigione
            worksheet.Cell(index, 6).Style.Font.Bold = true;

            worksheet.Cell(index, 7).FormulaA1 = $"SUM(G{indexSellout}:G{index - 1})";
            worksheet.Cell(index, 7).Style.Font.Bold = true;
            worksheet.Cell(index, 7).Style.NumberFormat.Format = "#,##0.00 €";
            //indexRowProvvigioneAgenteSellout = index;

            return index;
        }

        private static int gruppiStatistici(XLWorkbook workbook, int indexInit, string annoCorrente, string annoRiferimento, IList<ClienteResponse> clienteResponse, List<GruppoStatistico> GruppoStatistico, List<GruppoStatisticoRiepilogo> GruppoStatisticoPeriodo, string tipo)
        {
            List<ClienteResponse> cliente = (List<ClienteResponse>)clienteResponse;


            int index = indexInit;
            var worksheet = workbook.Worksheet("Provvigioni");
            int initTabella = 0;
            int endTabella = 0;
            index += 3;

            index += 2;
            worksheet.Cell($"C{index}").Value = $"GRUPPI STATISTICI - {tipo}";
            worksheet.Cell($"C{index}").Style.Font.FontSize = 20;
            worksheet.Range($"C{index}:G{index}").Merge();
            index += 2;
            initTabella = index;
            worksheet.Cell($"A{index}").Value = "Codice";
            worksheet.Cell($"B{index}").Value = "Descrizione";
            worksheet.Cell($"C{index}").Value = $"{annoRiferimento}";
            worksheet.Cell($"D{index}").Value = $"{annoCorrente}";
            worksheet.Cell($"E{index}").Value = "Delta imp.";

            worksheet.Range($"A{index}:F{index}").Style.Font.Bold = true;

            index++;
            GruppoStatistico.ForEach((x) =>
            {
                Console.WriteLine(x);


                var rif = GruppoStatisticoPeriodo.Where(y => y.CKY_MERC == x.CKY_MERC).ToList();

                //if (rif[0].ValoreRiferimento == 0 && rif[0].ValoreCorrente == 0)
                //{
                //    return;
                //}

                worksheet.Cell(index, 1).Value = x.CKY_MERC;
                worksheet.Cell(index, 2).Value = x.CDS_MERC;
                worksheet.Cell(index, 3).Value = rif[0].ValoreRiferimento;
                worksheet.Cell(index, 4).Value = rif[0].ValoreCorrente;
                worksheet.Cell(index, 5).FormulaA1 = $"D{index}-C{index}";

                worksheet.Cell(index, 3).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 4).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 5).Style.NumberFormat.Format = "#,##0.00 €";

                if (rif[0].ValoreRiferimento > rif[0].ValoreCorrente)
                {
                    worksheet.Cell(index, 5).Style.Font.FontColor = XLColor.RedRyb;
                }


                index++;
            });
            endTabella = index - 1;

            worksheet.Range($"A{initTabella}:E{endTabella}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"A{initTabella}:E{endTabella}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"A{initTabella}:E{endTabella}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"A{initTabella}:E{endTabella}").Style.Border.RightBorder = XLBorderStyleValues.Thin;



            worksheet.Cell(index, 3).FormulaA1 = $"SUM(C{initTabella + 1}:C{endTabella})";
            worksheet.Cell(index, 4).FormulaA1 = $"SUM(D{initTabella + 1}:D{endTabella})";
            worksheet.Cell(index, 5).FormulaA1 = $"SUM(E{initTabella + 1}:E{endTabella})";
            worksheet.Range($"A{index}:F{index}").Style.Font.Bold = true;
            worksheet.Range($"A{index}:F{index}").Style.NumberFormat.Format = "#,##0.00 €";

            if (Double.Parse(worksheet.Cell(index, 5).Value.ToString()) < 0)
            {
                worksheet.Cell(index, 5).Style.Font.FontColor = XLColor.RedRyb;
            }


            index++;

            return index;
        }

        private static int gruppiStatisticiClienti(XLWorkbook workbook, int indexInit, List<GruppoStatistico> GruppoStatistico, List<GruppoStatisticoRiepilogo> GruppoStatisticoPeriodo, IList<ClienteResponse> clienteResponse, string tipo)
        {
            //List<ClienteResponse> cliente = clienteResponse.Where(x => x.TotaleVendutoCorrente > 0 || x.TotaleVendutoRiferimento > 0).ToList(); ;
            List<ClienteResponse> cliente = clienteResponse.ToList();

            int index = indexInit + 5;


            var worksheet = workbook.Worksheet("Provvigioni");
            int initTabella = 0;
            int endTabella = 0;
            int col = 3;


            List<string> removeGruppoStatistico = new List<string>();
            GruppoStatistico.ForEach((x) =>
            {
                Console.WriteLine(x);


                var rif = GruppoStatisticoPeriodo.Where(y => y.CKY_MERC == x.CKY_MERC).ToList();

                //if (rif[0].ValoreRiferimento == 0 && rif[0].ValoreCorrente == 0)
                //{
                //    removeGruppoStatistico.Add(x.CKY_MERC);
                //    //return;
                //}

            });

            //removeGruppoStatistico.ForEach((x) => {
            //    GruppoStatistico.RemoveAll(y=>y.CKY_MERC == x);
            //});



            worksheet.Cell($"C{index}").Value = $"GRUPPI STATISTICI - CLIENTI - {tipo}";
            worksheet.Cell($"C{index}").Style.Font.FontSize = 20;
            worksheet.Range($"C{index}:G{index}").Merge();
            index += 2;
            initTabella = index;

            worksheet.Row(index).Style.Font.Bold = true;
            worksheet.Row(index).Style.Alignment.WrapText = true;
            worksheet.Cell(index, 1).Value = "Codice";
            worksheet.Cell(index, 2).Value = "Descrizione";

            GruppoStatistico.ForEach((i) =>
            {

                worksheet.Cell(index, col).Value = i.CDS_MERC.Trim();
                worksheet.Column(col).Width = 17;
                col++;
            });
            worksheet.Column(col).Width = 17;

            worksheet.Row(index).Height = 30;

            index++;



            cliente.ForEach((x) =>
            {

                worksheet.Cell(index, 1).Value = x.IdCliente.Trim();
                worksheet.Cell(index, 2).Value = x.NomeCliente.Trim();

                int colTemp = 3;

                GruppoStatistico.ForEach((i) =>
                {

                    var res = x.GruppoStatisticoCorrente.Find(x => x.CKY_MERC == i.CKY_MERC);

                    switch (tipo)
                    {
                        case "TRIMESTRE":
                            res = x.GruppoStatisticoCorrente.Find(x => x.CKY_MERC == i.CKY_MERC);
                            break;

                        case "PROGRESSIVO":
                            res = x.GruppoStatisticoCorrenteProgressivo.Find(x => x.CKY_MERC == i.CKY_MERC);
                            break;
                    }

                    if (res != null)
                    {
                        worksheet.Cell(index, colTemp).Value = res.Valore;
                    }
                    else
                    {
                        worksheet.Cell(index, colTemp).Value = "";
                    }

                    colTemp++;
                });

                index++;

            });

            endTabella = index - 1;

            worksheet.Range(worksheet.Cell(initTabella, 1).Address, worksheet.Cell(endTabella, col - 1).Address).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            worksheet.Range(worksheet.Cell(initTabella, 1).Address, worksheet.Cell(endTabella, col - 1).Address).Style.Border.TopBorder = XLBorderStyleValues.Thin;
            worksheet.Range(worksheet.Cell(initTabella, 1).Address, worksheet.Cell(endTabella, col - 1).Address).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Range(worksheet.Cell(initTabella, 1).Address, worksheet.Cell(endTabella, col - 1).Address).Style.Border.RightBorder = XLBorderStyleValues.Thin;

            worksheet.Range(worksheet.Cell(initTabella + 1, 3).Address, worksheet.Cell(endTabella + 1, col).Address).Style.NumberFormat.Format = "#,##0.00 €";
            worksheet.Row(index).Style.Font.Bold = true;

            for (int c = 3; c < col; c++)
            {
                var indInit = worksheet.Cell(initTabella + 1, c).Address.ToString();
                var indEnd = worksheet.Cell(endTabella, c).Address.ToString();
                worksheet.Cell(endTabella + 1, c).FormulaA1 = $"SUM({indInit}:{indEnd})";
            }


            for (int r = initTabella + 1; r <= endTabella; r++)
            {
                var indInit = worksheet.Cell(r, 3).Address.ToString();
                var indEnd = worksheet.Cell(r, col - 1).Address.ToString();
                worksheet.Cell(r, col).FormulaA1 = $"SUM({indInit}:{indEnd})";
                worksheet.Cell(r, col).Style.Font.Bold = true;
            }

            return index;
        }




        private static int _gruppiStatisticiClienti(XLWorkbook workbook, int indexInit, List<GruppoStatistico> GruppoStatistico, List<GruppoStatisticoRiepilogo> GruppoStatisticoPeriodo, IList<ClienteResponse> clienteResponse, string tipo)
        {
            List<ClienteResponse> cliente = clienteResponse.Where(x => x.TotaleVendutoCorrente > 0 || x.TotaleVendutoRiferimento > 0).ToList(); ;

            int index = indexInit + 5;


            var worksheet = workbook.Worksheet("Provvigioni");
            int initTabella = 0;
            int endTabella = 0;
            int col = 3;

            worksheet.Cell($"C{index}").Value = $"GRUPPI STATISTICI - CLIENTI - {tipo}";
            worksheet.Cell($"C{index}").Style.Font.FontSize = 20;
            worksheet.Range($"C{index}:G{index}").Merge();
            index += 2;
            initTabella = index;

            worksheet.Row(index).Style.Font.Bold = true;
            worksheet.Row(index).Style.Alignment.WrapText = true;
            worksheet.Cell(index, 1).Value = "Codice";
            worksheet.Cell(index, 2).Value = "Descrizione";

            GruppoStatisticoPeriodo.ForEach((i) =>
            {
                if (i.ValoreCorrente == 0)
                {
                    return;
                }
                worksheet.Cell(index, col).Value = i.CDS_MERC.Trim();
                worksheet.Column(col).Width = 17;
                col++;
            });

            worksheet.Row(index).Height = 30;

            index++;



            cliente.ForEach((x) =>
            {

                worksheet.Cell(index, 1).Value = x.IdCliente.Trim();
                worksheet.Cell(index, 2).Value = x.NomeCliente.Trim();

                int colTemp = 3;
                GruppoStatisticoPeriodo.ForEach((i) =>
                {
                    if (i.ValoreCorrente == 0)
                    {
                        return;
                    }

                    var res = x.GruppoStatisticoCorrente.Find(x => x.CKY_MERC == i.CKY_MERC);
                    if (res != null)
                    {
                        worksheet.Cell(index, colTemp).Value = res.Valore;
                    }
                    else
                    {
                        worksheet.Cell(index, colTemp).Value = "";
                    }

                    colTemp++;
                });

                index++;

            });

            endTabella = index - 1;

            worksheet.Range(worksheet.Cell(initTabella, 1).Address, worksheet.Cell(endTabella, col - 1).Address).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            worksheet.Range(worksheet.Cell(initTabella, 1).Address, worksheet.Cell(endTabella, col - 1).Address).Style.Border.TopBorder = XLBorderStyleValues.Thin;
            worksheet.Range(worksheet.Cell(initTabella, 1).Address, worksheet.Cell(endTabella, col - 1).Address).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Range(worksheet.Cell(initTabella, 1).Address, worksheet.Cell(endTabella, col - 1).Address).Style.Border.RightBorder = XLBorderStyleValues.Thin;

            worksheet.Range(worksheet.Cell(initTabella + 1, 3).Address, worksheet.Cell(endTabella + 1, col).Address).Style.NumberFormat.Format = "#,##0.00 €";
            worksheet.Row(index).Style.Font.Bold = true;

            for (int c = 3; c < col; c++)
            {
                var indInit = worksheet.Cell(initTabella + 1, c).Address.ToString();
                var indEnd = worksheet.Cell(endTabella, c).Address.ToString();
                worksheet.Cell(endTabella + 1, c).FormulaA1 = $"SUM({indInit}:{indEnd})";
            }


            for (int r = initTabella + 1; r <= endTabella; r++)
            {
                var indInit = worksheet.Cell(r, 3).Address.ToString();
                var indEnd = worksheet.Cell(r, col - 1).Address.ToString();
                worksheet.Cell(r, col).FormulaA1 = $"SUM({indInit}:{indEnd})";
                worksheet.Cell(r, col).Style.Font.Bold = true;
            }

            return index;
        }

        public static void generaExcelTotale(string annoCorrente, string annoRiferimento, string trimestre, IList<AgenteRiepilogo> AgentiRiepilogo, List<GruppoStatistico> GruppoStatistico)
        {

            Dictionary<string, string> trimestri = new Dictionary<string, string>() { { "t_1", "TRIM-1" }, { "t_2", "TRIM-2" }, { "t_3", "TRIM-3" }, { "t_4", "TRIM-4" } };
            Dictionary<string, string> trimestriSuExcel = new Dictionary<string, string>() { { "t_1", "1° TRIM" }, { "t_2", "2° TRIM" }, { "t_3", "3° TRIM" }, { "t_4", "4° TRIM" } };

            string path = "../excelAgenti";
            string pathFile = $"{path}/{annoCorrente}-{trimestri[trimestre]}--RIEPILOGO-AGENTI.xlsx";
            string fullPath = System.IO.Path.GetFullPath(pathFile);
            int index = 9;
            int rowInit = index;


            // legge xml
            List<Agente> cc = new List<Agente>();
            XmlSerializer xmlsd = new XmlSerializer(typeof(List<Agente>));
            using (TextReader tr = new StreamReader(@"agenti.xml"))
            {
                cc = (List<Agente>)xmlsd.Deserialize(tr);
            }

            // TRASFERITI --------------------------------------------
            var elencoTrasferiti = General.directoryTrasferiti(annoCorrente);
            //   var trs = new TrasferitiService(Regione, annoCorrente, trimestre, elencoTrasferiti);
            // ------------------------------------------------------

            List<string> titleTable = new List<string>();
            titleTable.Add("Codice");
            titleTable.Add("Descrizione");
            titleTable.Add($"Imp. periodo {annoRiferimento}");
            titleTable.Add($"Imp. periodo {annoCorrente}");
            titleTable.Add("Delta imp.");
            titleTable.Add("Delta imp. %");
            titleTable.Add("Imp. Sellout");
            titleTable.Add("Provvigioni Passepartout");
            titleTable.Add("Provvigioni Sellout");
            titleTable.Add("Provvigioni TOTALI");

            int indexRowProvvigioneAgente = 0;
            int indexRowProvvigioneAgenteSellout = 0;

            if (!File.Exists(path))
            {
                DirectoryInfo di = Directory.CreateDirectory(path);

            }

            using var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Resoconto");


            var imagePath = @"../logo.jpg";

            var image = worksheet.AddPicture(imagePath)
                .MoveTo(worksheet.Cell("A2"))
                .Scale(0.3); // optional: resize picture

            int indexCol = 1;

            titleTable.ForEach((x) =>
            {

                int index = titleTable.IndexOf(x);

                worksheet.Cell(8, index + 1).Value = x;
            });

            worksheet.Column("A").Width = 10;
            worksheet.Column("B").Width = 45;
            worksheet.Column("C").Width = 14;
            worksheet.Column("D").Width = 14;
            worksheet.Column("E").Width = 14;
            worksheet.Column("F").Width = 14;
            worksheet.Column("G").Width = 14;
            worksheet.Column("H").Width = 14;
            worksheet.Column("I").Width = 14;
            worksheet.Column("J").Width = 14;


            worksheet.Cell("C2").Value = "FT_AGENTI";
            worksheet.Cell("C2").Style.Font.FontSize = 20;
            worksheet.Range("C2:G2").Merge();


            worksheet.Cell("C3").Value = $"{trimestriSuExcel[trimestre]} {annoCorrente}";
            worksheet.Cell("C3").Style.Font.FontSize = 20;
            worksheet.Range("C3:G3").Merge();

            worksheet.Row(2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            worksheet.Row(3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);



            worksheet.Range($"A8:Z8").Style.Font.Bold = true;

            //foreach (ClienteResponse cliente in clienteResponse)
            foreach (AgenteRiepilogo cliente in AgentiRiepilogo)
            {
                double[] result = new double[2];



                worksheet.Cell(index, 3).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 4).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 5).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 6).Style.NumberFormat.Format = "0.00%";
                worksheet.Cell(index, 7).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 8).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 9).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 10).Style.NumberFormat.Format = "#,##0.00 €";

                worksheet.Cell(index, 1).Value = cliente.ID;
                worksheet.Cell(index, 2).Value = cliente.Nome;
                worksheet.Cell(index, 3).Value = cliente.VendutoRiferimento;
                worksheet.Cell(index, 4).Value = cliente.VendutoCorrente;
                worksheet.Cell(index, 5).Value = cliente.Delta;
                worksheet.Cell(index, 6).Value = cliente.DeltaPercent == double.PositiveInfinity ? 0 : cliente.DeltaPercent;
                worksheet.Cell(index, 7).Value = cliente.VendutoSellout;
                worksheet.Cell(index, 8).Value = cliente.ProvvigioneCorrente;
                worksheet.Cell(index, 9).Value = cliente.ProvvigioneSellout;

                string columnString1 = worksheet.Cell(index, 8).WorksheetColumn().ColumnLetter();
                string columnString2 = worksheet.Cell(index, 9).WorksheetColumn().ColumnLetter();

                worksheet.Cell(index, 10).FormulaA1 = $"SUM({columnString1}{index}:{columnString2}{index})";

                worksheet.Row(index).Height = 20;

                index++;
            }

            worksheet.Range($"C{index}:Z{index}").Style.NumberFormat.Format = "#,##0.00 €";
            worksheet.Range($"C{index}:Z{index}").Style.Font.Bold = true;


            for (int col = 3; col <= 10; col++)
            {
                string columnString = worksheet.Cell(index, col).WorksheetColumn().ColumnLetter();
                switch (columnString)
                {
                    case "F":
                        worksheet.Cell(index, col).FormulaA1 = $"AVERAGE({columnString}{rowInit}:{columnString}{index - 1})";
                        worksheet.Cell(index, col).Style.NumberFormat.Format = "0.00%";
                        break;

                    default:
                        worksheet.Cell(index, col).FormulaA1 = $"SUM({columnString}{rowInit}:{columnString}{index - 1})";
                        break;
                }


            }



            worksheet.Row(8).Height = 30;
            worksheet.Row(8).Style.Alignment.WrapText = true;
            worksheet.Row(8).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            for (int rw = 9; rw <= index; rw++)
            {
                worksheet.Row(rw).Height = 20;
            }


            string initCol = "";
            string endCol = "J";
            for (int rw = 8; rw <= index; rw++)
            {
                initCol = "A";


                if (rw == index)
                {
                    initCol = "C";
                }

                worksheet.Range($"{initCol}{rw}:{endCol}{rw}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range($"{initCol}{rw}:{endCol}{rw}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range($"{initCol}{rw}:{endCol}{rw}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Range($"{initCol}{rw}:{endCol}{rw}").Style.Border.RightBorder = XLBorderStyleValues.Thin;

            }

            index += 3;
            worksheet.Cell(index, 2).Value = "SELL OUT";
            worksheet.Cell(index, 2).Style.Font.FontSize = 20;
            index += 2;


            foreach (AgenteRiepilogo cliente in AgentiRiepilogo)
            {
                string id = cliente.ID;
                var Regione = cc.Find(x => x.ID == id).Regione.ToList();
                var trs = new TrasferitiService(Regione, annoCorrente, trimestre, elencoTrasferiti);
                worksheet.Cell(index, 1).Value = cliente.ID;
                worksheet.Cell(index, 2).Value = cliente.Nome;

                worksheet.Range($"A{index}:Z{index}").Style.Font.Bold = true;

                index++;
                int indexInit = index;
                foreach (Final final in (List<Final>)trs.Trasferiti)
                {
                    worksheet.Cell(index, 2).Value = final.Fornitore;
                    worksheet.Cell(index, 3).Value = final.ValoreEuro;
                    worksheet.Cell(index, 3).Style.NumberFormat.Format = "#,##0.00 €";
                    index++;
                }
                worksheet.Cell(index, 3).FormulaA1 = $"SUM(C{indexInit}:C{index - 1})";
                worksheet.Cell(index, 3).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 3).Style.Font.Bold = true;




                index++;
                index++;
            }

            //   
            workbook.SaveAs(pathFile);

            Process.Start("explorer.exe", fullPath);
        }

        private static double[] calcolaPercentuale(double annoPrecedente, double annoCorrente)
        {
            double[] result = new double[2];
            double percent = 0;
            double delta = annoCorrente - annoPrecedente;
            result[0] = delta;
            if (annoPrecedente > 0)
            {
                percent = delta / annoPrecedente;
            }
            result[1] = percent;
            return result;
        }


        public static void OpenUrl(string url)
        {
            try
            {
                Process.Start(url);
            }
            catch
            {
                // hack because of this: https://github.com/dotnet/corefx/issues/10361
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    url = url.Replace("&", "^&");
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Process.Start("xdg-open", url);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Process.Start("open", url);
                }
                else
                {
                    throw;
                }
            }
        }

        public static List<Trasferito> estraiXmlSellout(List<string> listDir)
        {
            List<Trasferito> a = new List<Trasferito>();

            List<Trasferito> cc = null;

            List<string> fileXml = new List<string>();

            foreach (var item in listDir)
            {
                string[] subs = item.Split('\\');

                string nameFile = subs[subs.Length - 1];

                if (nameFile == "mc_elettrici")
                {
                    fileXml.Add($"{item}/mc_elettrici_vendita_diretta.xml");
                    fileXml.Add($"{item}/mc_elettrici_magazzino.xml");
                }
                else
                {
                    fileXml.Add($"{item}/{nameFile}.xml");
                }
            }


            foreach (var item in fileXml)
            {


                if (!File.Exists(item))
                    continue;

                // legge xml
                XmlSerializer xmlsd = new XmlSerializer(typeof(List<Trasferito>));
                using (TextReader tr = new StreamReader($"{item}"))
                {
                    cc = (List<Trasferito>)xmlsd.Deserialize(tr);
                }

                cc.ForEach((x) =>
                {

                    var item = a.Find(y => y.Regione == x.Regione);

                    if (item == null)
                    {
                        a.Add(new Trasferito() { Regione = x.Regione });
                    }

                    item = a.Find(y => y.Regione == x.Regione);
                    item.Venduto += x.Venduto;

                });
            }

            return a;
        }


        private static string xmlPeriodoHome = "trimestreSelezionatoHome.xml";
        public static void salvaPeriodoHome(string trimestre = null, string annoCorrente = null, string annoRiferimento = null)
        {
            PeriodoStart p = leggiPeriodoHome();

            if (trimestre != null)
            {
                p.Trimestre = trimestre;
            }

            if (annoCorrente != null)
            {
                p.AnnoCorrente = annoCorrente;
            }

            if (annoRiferimento != null)
            {
                p.AnnoRiferimento = annoRiferimento;
            }

            XmlSerializer x = new XmlSerializer(p.GetType());
            using (TextWriter writer = new StreamWriter(xmlPeriodoHome))
            {
                x.Serialize(writer, p);
            }
        }

        public static PeriodoStart leggiPeriodoHome()
        {
            PeriodoStart p = new PeriodoStart();

            if (!File.Exists(xmlPeriodoHome))
            {
                XmlSerializer x = new XmlSerializer(p.GetType());
                using (TextWriter writer = new StreamWriter(xmlPeriodoHome))
                {
                    x.Serialize(writer, p);
                }
            }

            // legge xml
            XmlSerializer xmlsd = new XmlSerializer(p.GetType());
            using (TextReader tr = new StreamReader(xmlPeriodoHome))
            {
                p = (PeriodoStart)xmlsd.Deserialize(tr);
            }

            return p;
        }


        // ----------------------------------------------------------------------------------------------

        private static string xmlPeriodoSelezionatoHome = "periodoSelezionatoHome.xml";
        public static PeriodoSelezionato periodoHome()
        {
            PeriodoSelezionato p = new PeriodoSelezionato();

            if (!File.Exists(xmlPeriodoSelezionatoHome))
            {
                XmlSerializer x = new XmlSerializer(p.GetType());
                using (TextWriter writer = new StreamWriter(xmlPeriodoSelezionatoHome))
                {
                    x.Serialize(writer, p);
                }
            }

            // legge xml
            XmlSerializer xmlsd = new XmlSerializer(p.GetType());
            using (TextReader tr = new StreamReader(xmlPeriodoSelezionatoHome))
            {
                p = (PeriodoSelezionato)xmlsd.Deserialize(tr);
            }

            return p;

        }

        public static void periodoHomeSave(PeriodoSelezionato p)
        {
            XmlSerializer x = new XmlSerializer(p.GetType());
            using (TextWriter writer = new StreamWriter(xmlPeriodoSelezionatoHome))
            {
                x.Serialize(writer, p);
            }
        }


        public static List<string> cercaMese(ComboBox mese, ComboBox trimestre)
        {
            List<string> list = new List<string>();

            if (mese.SelectedItem != null)
            {

                list.Add(((PeriodoList)mese.SelectedItem).Valore);
            }

            if (trimestre.SelectedItem != null)
            {

                string trimesteSelezionato = ((PeriodoList)trimestre.SelectedItem).Valore;

                if (trimesteSelezionato == Trimestri.T1)
                {
                    list.Add(Mesi.Gennaio);
                    list.Add(Mesi.Febbraio);
                    list.Add(Mesi.Marzo);
                }

                if (trimesteSelezionato == Trimestri.T2)
                {
                    list.Add(Mesi.Aprile);
                    list.Add(Mesi.Maggio);
                    list.Add(Mesi.Giugno);
                }
                if (trimesteSelezionato == Trimestri.T3)
                {
                    list.Add(Mesi.Luglio);
                    list.Add(Mesi.Agosto);
                    list.Add(Mesi.Settembre);
                }
                if (trimesteSelezionato == Trimestri.T4)
                {
                    list.Add(Mesi.Luglio);
                    list.Add(Mesi.Agosto);
                    list.Add(Mesi.Settembre);
                }

                //switch (trimesteSelezionato)
                //{

                //    case "t_1":
                //        list.Add(Mesi.Gennaio);
                //        list.Add(Mesi.Febbraio);
                //        list.Add(Mesi.Marzo);
                //        break;

                //    case "t_2":
                //        list.Add(Mesi.Aprile);
                //        list.Add(Mesi.Maggio);
                //        list.Add(Mesi.Giugno);
                //        break;

                //    case "t_3":
                //        list.Add(Mesi.Lugio);
                //        list.Add(Mesi.Agosto);
                //        list.Add(Mesi.Settembre);
                //        break;

                //    case "t_4":
                //        list.Add(Mesi.Ottobre);
                //        list.Add(Mesi.Novembre);
                //        list.Add(Mesi.Dicembre);
                //        break;

                //}



            }

            return list;
        }

        public static List<ClientiContactDiretti> ClientiContactDiretti(string annoCorrente, string annoRiferimento)
        {
            List<ClientiContactDiretti> list = new List<ClientiContactDiretti>();

            string query = $"SELECT CKY_CNT_CLFR FROM MMA_M WHERE CKY_CNT_AGENTE = '415.01009'  and NGB_ANNO_DOC in ('{annoRiferimento}','{annoCorrente}') GROUP BY CKY_CNT_CLFR";

            List<Anno> A = new List<Anno>();

            list = Controllers.Database.SELECT_GET_LIST<ClientiContactDiretti>(query);

            return list;
        }

    }
}
