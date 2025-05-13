using System.Globalization;
using System.Windows.Controls;
using Provvigioni_Agenti.Models;
using System.Windows.Media;
using System.IO;
using ClosedXML.Excel;
using System.Xml.Serialization;
using System.Diagnostics;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Bibliography;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;

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

            List<string> trasferiti = new List<string>() { "acmei", "barcella", "comoli", "edif", "mc_elettrici", "meb", "rexel", "sacchi", "sonepar" };
            List<string> trimestri = new List<string>() { "t_1", "t_2", "t_3", "t_4" };
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

            foreach (string trimestre in trimestri)
            {
                string subPathTrimestre = $"{pathYear}/{trimestre}";

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

                //string[] fileEntries = {"ciao"};

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
            //generaCittaBarcella();
            //generaCittaRexel();
            //generaCittaSacchi();


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

            //generaCittaXml("citta_acmei");
            //generaCittaXml("citta_barcella");
            //generaCittaXml("citta_rexel");
            //generaCittaXml("citta_sacchi");
            //generaCittaXml("citta_barcella");
            //generaCittaXml("citta_comoli");
            //generaCittaXml("citta_mc_elettrici");

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



        public static void generaExcelTrasferiti(string agente, string agenteFullName, string annoCorrente, string annoRiferimento, string trimestre, IList<ClienteResponse> clienteResponse, IList<Final> Trasferiti, IList<CategoriaStatistica> categorieStatistiche, IList<CategoriaStatistica> categorieStatisticheTotaleProgressivo)
        {
            Dictionary<string, string> trimestri = new Dictionary<string, string>() { { "t_1", "TRIM-1" }, { "t_2", "TRIM-2" }, { "t_3", "TRIM-3" }, { "t_4", "TRIM-4" } };
            Dictionary<string, string> trimestriSuExcel = new Dictionary<string, string>() { { "t_1", "1° TRIM" }, { "t_2", "2° TRIM" }, { "t_3", "3° TRIM" }, { "t_4", "4° TRIM" } };

            IList<ClienteResponse> clienteResponseFiltered = clienteResponse.Where(x => x.TotaleVendutoCorrente > 0 || x.TotaleVendutoRiferimento > 0).ToList();

            string path = "../excelAgenti";
            string pathFile = $"{path}/{annoCorrente}-{trimestri[trimestre]}--{agente}.xlsx";
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

            var imagePath = @"../logo.jpg";

            var image = worksheet.AddPicture(imagePath)
                .MoveTo(worksheet.Cell("A2"))
                .Scale(0.3); // optional: resize picture



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


            worksheet.Cell("D3").Value = $"{trimestriSuExcel[trimestre]} {annoCorrente} - PROVVIGIONE TOTALE: ";
            worksheet.Cell("D3").Style.Font.FontSize = 15;
            worksheet.Range("D3:G3").Merge();
            //worksheet.Cell("H3").Value

            worksheet.Cell("A8").Value = $"Codice";
            worksheet.Cell("B8").Value = $"Descrizione";
            worksheet.Cell("C8").Value = $"Imp. periodo " + annoRiferimento;
            worksheet.Cell("D8").Value = $"Imp. periodo " + annoCorrente;
            worksheet.Cell("E8").Value = $"Delta imp.";
            worksheet.Cell("F8").Value = $"Delta imp. %";
            worksheet.Cell("H8").Value = $"Provvigione";



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
                worksheet.Cell(index, 8).Value = cliente.ProvvigioneCorrente;




                worksheet.Cell(index, 3).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 4).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 5).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 6).Style.NumberFormat.Format = "0.00%";
                worksheet.Cell(index, 8).Style.NumberFormat.Format = "#,##0.00 €";



                if (result[0] < 0)
                {
                    worksheet.Cell(index, 2).Style.Fill.BackgroundColor = XLColor.Red;
                    worksheet.Cell(index, 5).Style.Fill.BackgroundColor = XLColor.Red;
                    worksheet.Cell(index, 6).Style.Fill.BackgroundColor = XLColor.Red;

                    worksheet.Cell(index, 2).Style.Font.Bold = true;
                    worksheet.Cell(index, 5).Style.Font.Bold = true;
                    worksheet.Cell(index, 6).Style.Font.Bold = true;
                }

                index++;
            }


            worksheet.Range($"A9:H{index - 1}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"A9:H{index - 1}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"A9:H{index - 1}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"A9:H{index - 1}").Style.Border.RightBorder = XLBorderStyleValues.Thin;


            //index += 1;


            worksheet.Range($"A{index}:H{index}").Style.Font.Bold = true;
            worksheet.Range($"C{index}:H{index}").Style.NumberFormat.Format = "#,##0.00 €";
            worksheet.Cell($"F{index}").Style.NumberFormat.Format = "0.00%";

            worksheet.Cell(index, 2).Value = "TOTALE"; // vendite

            worksheet.Cell(index, 3).FormulaA1 = $"SUM(C9:C{index - 1})";
            worksheet.Cell(index, 4).FormulaA1 = $"SUM(D9:D{index - 1})";
            worksheet.Cell(index, 5).FormulaA1 = $"SUM(E9:E{index - 1})";
            //worksheet.Cell(index, 6).FormulaA1 = $"(SUM(D9:D{index - 1})-SUM(C9:C{index - 1}))/SUM(C9:C{index - 1})";
            worksheet.Cell(index, 6).FormulaA1 = $"(D{index}-C{index})/C{index}";

            if ((double)worksheet.Cell(index, 5).Value < 0)
            {
                worksheet.Cell(index, 5).Style.Fill.BackgroundColor = XLColor.Red;
                worksheet.Cell(index, 6).Style.Fill.BackgroundColor = XLColor.Red;
            }



            worksheet.Cell(index, 8).FormulaA1 = $"SUM(H9:H{index - 1})";
            indexRowProvvigioneAgente = index; // prooviogne db

            index += 2;

            worksheet.Range($"A{index}:H{index}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;  // bordo

            index += 2;

            worksheet.Cell(index, 2).Value = "SELLOUT";
            worksheet.Cell(index, 8).Value = $"Provvigione";
            worksheet.Range($"A{index}:H{index}").Style.Font.Bold = true;
            //worksheet.Range($"A{index}:H{index}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;  // bordo
            index++;

            int indexSellout = index;

            foreach (Final f in Trasferiti)
            {
                worksheet.Cell(index, 2).Value = f.Fornitore;
                worksheet.Cell(index, 4).Value = f.ValoreEuro;
                worksheet.Cell(index, 4).Style.NumberFormat.Format = "#,##0.00 €";

                worksheet.Cell(index, 8).Value = f.ValoreEuro * 0.02;
                worksheet.Cell(index, 8).Style.NumberFormat.Format = "#,##0.00 €";

                index++;
            }


            worksheet.Range($"B{indexSellout}:H{index - 1}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"B{indexSellout}:H{index - 1}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"B{indexSellout}:H{index - 1}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"B{indexSellout}:H{index - 1}").Style.Border.RightBorder = XLBorderStyleValues.Thin;


            worksheet.Cell(index, 3).Value = "TOTALE"; // sellout
            worksheet.Cell(index, 3).Style.Font.Bold = true;

            worksheet.Cell(index, 4).FormulaA1 = $"SUM(D{indexSellout}:D{index - 1})";
            worksheet.Cell(index, 4).Style.Font.Bold = true;
            worksheet.Cell(index, 4).Style.NumberFormat.Format = "#,##0.00 €";

            worksheet.Cell(index, 7).Value = "TOTALE"; // sellout provvigione
            worksheet.Cell(index, 7).Style.Font.Bold = true;

            worksheet.Cell(index, 8).FormulaA1 = $"SUM(H{indexSellout}:H{index - 1})";
            worksheet.Cell(index, 8).Style.Font.Bold = true;
            worksheet.Cell(index, 8).Style.NumberFormat.Format = "#,##0.00 €";
            indexRowProvvigioneAgenteSellout = index;


            worksheet.Cell("H3").FormulaA1 = $"SUM(H{indexRowProvvigioneAgente}+H{indexRowProvvigioneAgenteSellout})";
            worksheet.Cell("H3").Style.NumberFormat.Format = "#,##0.00 €";
            worksheet.Cell("H3").Style.Font.Bold = true;
            worksheet.Cell("H3").Style.Font.FontSize = 15;

            index += 2;
            worksheet.Range($"A{index}:H{index}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;

            index += 2;

            worksheet.Cell(index, 2).Value = "CATEGORIE STATISTICHE";
            worksheet.Cell(index, 2).Style.Font.Bold = true;

            index++;
            int indexCatStat = index;
            foreach (CategoriaStatistica catStatTitle in categorieStatistiche)
            {
                worksheet.Cell(index, 2).Value = catStatTitle.Categoria.Trim(' ');
                worksheet.Cell(index, 4).Value = catStatTitle.ValoreCorrente;
                worksheet.Cell(index, 4).Style.NumberFormat.Format = "#,##0.00 €";
                index++;
            }


            worksheet.Cell(index, 3).Value = "TOTALE"; // categorie statistiche
            worksheet.Cell(index, 3).Style.Font.Bold = true;

            worksheet.Cell(index, 4).FormulaA1 = $"SUM(D{indexCatStat}:D{index - 1})";
            worksheet.Cell(index, 4).Style.Font.Bold = true;
            worksheet.Cell(index, 4).Style.NumberFormat.Format = "#,##0.00 €";


            worksheet.Range($"B{indexCatStat}:D{index - 1}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"B{indexCatStat}:D{index - 1}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"B{indexCatStat}:D{index - 1}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"B{indexCatStat}:D{index - 1}").Style.Border.RightBorder = XLBorderStyleValues.Thin;



            index += 2;
            worksheet.Range($"A{index}:H{index}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            index += 2;

            worksheet.Cell(index, 2).Value = "CATEGORIE STATISTICHE PROGRESSIVO";
            worksheet.Cell(index, 2).Style.Font.Bold = true;
            index++;
            int indexCatStatProgr = index;
            foreach (CategoriaStatistica catStatTitle in categorieStatisticheTotaleProgressivo)
            {
                worksheet.Cell(index, 2).Value = catStatTitle.Categoria.Trim(' ');
                worksheet.Cell(index, 4).Value = catStatTitle.ValoreCorrente;
                worksheet.Cell(index, 4).Style.NumberFormat.Format = "#,##0.00 €";
                index++;
            }
            worksheet.Cell(index, 3).Value = "TOTALE"; // categorie statistiche progressivo
            worksheet.Cell(index, 3).Style.Font.Bold = true;

            worksheet.Cell(index, 4).FormulaA1 = $"SUM(D{indexCatStatProgr}:D{index - 1})";
            worksheet.Cell(index, 4).Style.Font.Bold = true;
            worksheet.Cell(index, 4).Style.NumberFormat.Format = "#,##0.00 €";



            worksheet.Range($"B{indexCatStatProgr}:D{index - 1}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"B{indexCatStatProgr}:D{index - 1}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"B{indexCatStatProgr}:D{index - 1}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"B{indexCatStatProgr}:D{index - 1}").Style.Border.RightBorder = XLBorderStyleValues.Thin;

            index += 2;
            worksheet.Range($"A{index}:H{index}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            index += 2;

            int indexCol = 3;
            foreach (CategoriaStatistica catStatTitle in categorieStatistiche)
            {
                worksheet.Cell(index, indexCol).Value = catStatTitle.Categoria;
                worksheet.Cell(index, indexCol).Style.Font.Bold = true;

                indexCol++;
            }

            index++;
            int indexCatStatExpl = index;
            foreach (ClienteResponse cliente in clienteResponse)
            {


                worksheet.Cell(index, 1).Value = cliente.IdCliente;
                worksheet.Cell(index, 2).Value = cliente.NomeCliente;

                // ciclare \\
                int titleCatStat = 3;
                foreach (CategoriaStatistica catStatTitle in categorieStatistiche)
                {
                    var res = cliente.CategoriaStatistica.Find(x => x.Categoria == catStatTitle.Categoria);
                    if (res != null)
                    {
                        worksheet.Cell(index, titleCatStat).Value = res.ValoreCorrente;
                        worksheet.Cell(index, titleCatStat).Style.NumberFormat.Format = "#,##0.00 €";
                    }
                    titleCatStat++;

                }
                string colinit = worksheet.Cell(index, 3).WorksheetColumn().ColumnLetter();
                string colEnd = worksheet.Cell(index, titleCatStat - 1).WorksheetColumn().ColumnLetter();

                worksheet.Cell(index, titleCatStat).FormulaA1 = $"SUM({colinit}{index}:{colEnd}{index})";
                worksheet.Cell(index, titleCatStat).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, titleCatStat).Style.Font.Bold = true;

                worksheet.Cell(index, indexCol).Style.Fill.BackgroundColor = XLColor.Yellow;


                index++;
            }
            worksheet.Cell(index, 2).Value = "TOTALE"; // categorie statistiche expl
            worksheet.Cell(index, 2).Style.Font.Bold = true;




            indexCol = 3;
            foreach (CategoriaStatistica catStatTitle in categorieStatistiche)
            {

                string col = worksheet.Cell(index, indexCol).WorksheetColumn().ColumnLetter();

                worksheet.Cell(index, indexCol).FormulaA1 = $"SUM({col}{indexCatStatExpl}:{col}{index - 1})";
                worksheet.Cell(index, indexCol).Style.Font.Bold = true;
                worksheet.Cell(index, indexCol).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, indexCol).Style.Fill.BackgroundColor = XLColor.Yellow;
                indexCol++;
            }

            string colTot = worksheet.Cell(index, indexCol).WorksheetColumn().ColumnLetter();

            worksheet.Cell(index, indexCol).FormulaA1 = $"SUM({colTot}{indexCatStatExpl}:{colTot}{index - 1})";
            worksheet.Cell(index, indexCol).Style.Font.Bold = true;
            worksheet.Cell(index, indexCol).Style.NumberFormat.Format = "#,##0.00 €";
            worksheet.Cell(index, indexCol).Style.Fill.BackgroundColor = XLColor.GreenYellow;


            worksheet.Range($"A{indexCatStatExpl}:{colTot}{index - 1}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"A{indexCatStatExpl}:{colTot}{index - 1}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"A{indexCatStatExpl}:{colTot}{index - 1}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Range($"A{indexCatStatExpl}:{colTot}{index - 1}").Style.Border.RightBorder = XLBorderStyleValues.Thin;


            workbook.SaveAs(pathFile);


            Process.Start("explorer.exe", System.IO.Path.GetFullPath($"{path}"));

            Process.Start("explorer.exe", fullPath);
        }

        public static void generaExcelTotale(string annoCorrente, string annoRiferimento, string trimestre, IList<AgenteRiepilogo> AgentiRiepilogo)
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
                    worksheet.Cell(index,2).Value = final.Fornitore;
                    worksheet.Cell(index,3).Value = final.ValoreEuro;
                    worksheet.Cell(index, 3).Style.NumberFormat.Format = "#,##0.00 €";
                    index++;
                }
                worksheet.Cell(index, 3).FormulaA1 = $"SUM(C{indexInit}:C{index-1})";
                worksheet.Cell(index, 3).Style.NumberFormat.Format = "#,##0.00 €";
                worksheet.Cell(index, 3).Style.Font.Bold = true;




                index++;
                index++;
            }

            //   
            workbook.SaveAs(pathFile);

            Process.Start("explorer.exe", System.IO.Path.GetFullPath($"{path}"));
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

        public static List<Trasferito> estraiXmlSellout(string[] listDir)
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
                    fileXml.Add($"{item}/mc_elettrici_consegna_diretta.xml");
                    fileXml.Add($"{item}/mc_elettrici_magazzino.xml");
                }
                else
                {
                    fileXml.Add($"{item}/{nameFile}.xml");
                }
            }


            foreach (var item in fileXml)
            {

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

    }
}
