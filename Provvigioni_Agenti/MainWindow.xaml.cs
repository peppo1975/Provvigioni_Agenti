
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection.Metadata;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Drawing;
using Provvigioni_Agenti.Controllers;
using Provvigioni_Agenti.Models;

namespace Provvigioni_Agenti
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Models.AgentiService agentiService = null;
        private Models.StoricoService storicoService = null;
        private Models.PeriodoService periodoService = null;
        private Models.CheckService checkService = null;
        private Models.Agente agente = null;
        private bool controlDate = false;

        string annoCorrenteTxt = string.Empty;
        string annoRiferimentoTxt = string.Empty;

        string trimestre = string.Empty;
        List<string> elencoTrasferiti = new List<string>();
        TrasferitiService trs = null;
        IList<ClienteResponseDatagrid> clienteResponseDatagrids = null;

        IList<ClienteResponse> clienteResponse = null;
        ClienteRiepilogoVendite ClientiRiepilogoVendite { get; set; }
        IList<CategoriaStatistica> categorieStatistiche = null;
        IList<CategoriaStatistica> categorieStatisticheProgressivo = null;
        IList<CategoriaStatistica> categorieStatisticheTotaleProgressivo = null;
        IList<AgenteRiepilogo> AgentiRiepilogo = null;


        Style HorizontalCenterStyle = null;

        public MainWindow()
        {
            InitializeComponent();

            General.leggiRegioni();

            General.leggiAgenti();

            apriImpostazioni.Visibility = Visibility.Hidden;

            List<Anno> A = new List<Anno>();

            A = Controllers.Database.SELECT_GET_LIST<Anno>("SELECT NGB_ANNO_DOC FROM MMA_M GROUP BY NGB_ANNO_DOC  ORDER BY NGB_ANNO_DOC DESC");

            foreach (Anno anno in A)
            {
                annoCorrente.Items.Add(anno.NGB_ANNO_DOC);
                annoRiferimento.Items.Add(anno.NGB_ANNO_DOC);
            }

            annoCorrente.SelectedItem = A.Max(t => t.NGB_ANNO_DOC);

            annoRiferimento.SelectedItem = A.Max(t => t.NGB_ANNO_DOC) - 1;

            annoCorrenteTxt = annoCorrente.SelectedItem.ToString();

            agentiService = new Models.AgentiService();

            elencoAgenti.ItemsSource = agentiService.Agenti;

            elencoAgenti.DisplayMemberPath = "NikName";

            controlDate = true;

            checkService = new Models.CheckService(t_1, t_2, t_3, t_4);

            buttonElabora.IsEnabled = false;

            General.directoryTrasferiti(A.Max(t => t.NGB_ANNO_DOC).ToString());

            HorizontalCenterStyle = new Style();

            HorizontalCenterStyle.Setters.Add(new Setter(HorizontalAlignmentProperty, HorizontalAlignment.Right));



        }

        private void elencoAgenti_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (elencoAgenti.SelectedItem != null)
            {
                {

                    buttonElabora.IsEnabled = true;
                    buttonElabora.Background = Brushes.GreenYellow;

                    agente = (Models.Agente)elencoAgenti.SelectedItem;

                    agenteInfo.Text = $"{agente.Nome} - {String.Join(", ", agente.Regione.Select(x => x.Nome))}";

                    refreshInfo();
                }

            }



            corr.Text = "";
            totInfo.Text = "";


            rif.Text = "";
            totInfoPrec.Text = "";
        }






        private void ElaboraAgente()
        {
            bool initCalc = true;

            General.generaXmlCitta();

            if (!initCalc)
                return;

            refreshInfo();



            annoCorrenteTxt = annoCorrente.SelectedItem.ToString();
            annoRiferimentoTxt = annoRiferimento.SelectedItem.ToString();

            // TRASFERITI --------------------------------------------
            elencoTrasferiti = General.directoryTrasferiti(annoCorrenteTxt);
            trs = new TrasferitiService(agente.Regione, annoCorrenteTxt, trimestre, elencoTrasferiti);
            // ------------------------------------------------------


            string query = Controllers.Query.clientiAgente(agente.ID, annoCorrenteTxt, annoRiferimentoTxt);
            var st = Controllers.Database.SELECT_GET_LIST<Storico>(query);




            string fromDateTxt = fromDate.ToString();
            string toDateTxt = toDate.ToString();

            storicoService = new Models.StoricoService();
            periodoService = new Models.PeriodoService(annoCorrente.SelectedItem.ToString(), annoRiferimento.SelectedItem.ToString(), fromDate, toDate);

            if (agente.ID.Contains('#'))
            {
                string[] dirs = Directory.GetDirectories(trimestreSelezionato());
                var elencoRiepilogo = new Models.AgentiRiepilogoService(st, periodoService.Periodo[0], dirs);
                AgentiRiepilogo = elencoRiepilogo.AgentiRiepilogo;

                dataGridVendite.ItemsSource = elencoRiepilogo.AgentiRiepilogo;


                dataGridVendite.Columns[7].Header = annoRiferimento.SelectedItem;
                dataGridVendite.Columns[8].Header = annoCorrente.SelectedItem;
                dataGridVendite.Columns[9].Header = "Δ (€)";
                dataGridVendite.Columns[10].Header = "Δ (%)";
                dataGridVendite.Columns[11].Header = $"Provv. {annoCorrenteTxt}";
                dataGridVendite.Columns[13].Header = $"Vend. sellout";
                dataGridVendite.Columns[15].Header = $"Provv. sellout";

                dataGridVendite.Columns[2].Visibility = Visibility.Collapsed;
                dataGridVendite.Columns[3].Visibility = Visibility.Collapsed;
                dataGridVendite.Columns[4].Visibility = Visibility.Collapsed;
                dataGridVendite.Columns[5].Visibility = Visibility.Collapsed;
                dataGridVendite.Columns[6].Visibility = Visibility.Collapsed;
                dataGridVendite.Columns[12].Visibility = Visibility.Collapsed;
                dataGridVendite.Columns[14].Visibility = Visibility.Collapsed;

                for (int i = 2; i <= 15; i++)
                {
                    dataGridVendite.Columns[i].CellStyle = HorizontalCenterStyle;
                }


                var riepilogoVendutoCorrente = elencoRiepilogo.AgentiRiepilogo.Sum(x => x.VendutoCorrente);
                var riepilogoVendutoRiferimento = elencoRiepilogo.AgentiRiepilogo.Sum(x => x.VendutoRiferimento);
                var riepilogoProvvigioni = elencoRiepilogo.AgentiRiepilogo.Sum(x => x.ProvvigioneCorrente);
                var riepilogoProvvigioniSellout = elencoRiepilogo.AgentiRiepilogo.Sum(x => x.ProvvigioneSellout);

                corr.Text = annoCorrenteTxt;
                rif.Text = annoRiferimentoTxt;
                totInfo.Text = riepilogoVendutoCorrente.ToString("C", CultureInfo.CurrentCulture);
                totInfoPrec.Text = riepilogoVendutoRiferimento.ToString("C", CultureInfo.CurrentCulture);
                totProvvigioneCorrente.Text = riepilogoProvvigioni.ToString("C", CultureInfo.CurrentCulture);
                totProvvigioneSellout.Text = riepilogoProvvigioniSellout.ToString("C", CultureInfo.CurrentCulture);

                deltaTrimestre.Text = (riepilogoVendutoCorrente - riepilogoVendutoRiferimento).ToString("C", CultureInfo.CurrentCulture);
                deltaTrimestrePercent.Text = String.Format("{0:P2}", ((riepilogoVendutoCorrente - riepilogoVendutoRiferimento) / riepilogoVendutoRiferimento));

                return;

            }



            var elenco = new Models.ClientiService(st, periodoService.Periodo[0]);

            ClientiRiepilogoVendite = elenco.ClientiRiepilogoVendite[0];
            categorieStatisticheTotaleProgressivo = elenco.CategorieStatitischeTotaleProgressivo;


            double total = trs.Trasferiti.Sum(x => x.ValoreEuro);

            trs.Trasferiti.Add(new Final() { Fornitore = " - - - TOTALE: ", Valore = total.ToString("C", CultureInfo.CurrentCulture), ValoreEuro = total });

            dataGridTrasferiti.ItemsSource = trs.Trasferiti;
            dataGridTrasferiti.Columns[1].CellStyle = HorizontalCenterStyle;

            dataGridTrasferiti.Columns[0].Width = 190;
            dataGridTrasferiti.Columns[1].Width = 80;

            dataGridTrasferiti.Columns[2].Visibility = Visibility.Collapsed;

            clienteResponseDatagrids = elenco.ClientiResponseDatagrid;
            categorieStatistiche = elenco.CategorieStatitischeTotale;

            clienteResponse = elenco.ClientiResponse;

            dataGridVendite.ItemsSource = clienteResponseDatagrids;

            dataGridVendite.Columns[2].Header = annoRiferimento.SelectedItem;
            dataGridVendite.Columns[3].Header = annoCorrente.SelectedItem;
            dataGridVendite.Columns[4].Header = "Δ (€)";
            dataGridVendite.Columns[5].Header = "Δ (%)";
            dataGridVendite.Columns[6].Header = $"Provv. {annoRiferimentoTxt}";
            dataGridVendite.Columns[7].Header = $"Provv. {annoCorrenteTxt}";
            dataGridVendite.Columns[8].Header = "Provv. %";

            dataGridVendite.Columns[8].Visibility = Visibility.Collapsed;
            dataGridVendite.Columns[6].Visibility = Visibility.Collapsed;

            // dataGridCategorieStatCliente.ItemsSource = trs.Trasferiti;



            for (int i = 2; i <= 7; i++)
            {
                dataGridVendite.Columns[i].CellStyle = HorizontalCenterStyle;
            }

            Console.WriteLine(dataGridVendite.Items.Count);

            //   var LastRow = dataGridVendite.Items[dataGridVendite.Items.Count - 1];
            // LastRow.Font.Bold = true;


            corr.Text = annoCorrenteTxt;

            totInfo.Text = General.valuta(ClientiRiepilogoVendite.TotaleVendutoCorrente);
            totInfoProgressivoCorrente.Text = General.valuta(ClientiRiepilogoVendite.ProgressivoCorrente);

            double provvigioneCorrente = clienteResponse.Sum(x => x.ProvvigioneCorrente);
            totProvvigioneCorrente.Text = General.valuta(provvigioneCorrente);


            double provvigioneTrasferiti = (double)trs.Trasferiti.Sum(x => x.ValoreEuro) * 0.02;
            totProvvigioneSellout.Text = General.valuta(provvigioneTrasferiti);

            totProvvigioneTrimestre.Text = General.valuta(provvigioneCorrente + provvigioneTrasferiti);

            rif.Text = annoRiferimentoTxt;
            totInfoPrec.Text = General.valuta(ClientiRiepilogoVendite.TotaleVendutoRiferimento);
            totInfoProgressivoRiferimento.Text = General.valuta(ClientiRiepilogoVendite.ProgressivoRiferimento);

            double deltaProgressivoValue = ClientiRiepilogoVendite.ProgressivoCorrente - ClientiRiepilogoVendite.ProgressivoRiferimento;
            deltaProgressivo.Text = General.valuta(deltaProgressivoValue);

            double deltaProgressivoPercentValue = deltaProgressivoValue / ClientiRiepilogoVendite.ProgressivoRiferimento;
            deltaProgressivoPercent.Text = General.percentuale(deltaProgressivoPercentValue);

            General.coloraVariazioni(ClientiRiepilogoVendite.ProgressivoCorrente, ClientiRiepilogoVendite.ProgressivoRiferimento, totInfoProgressivoCorrente);


            double deltaTrimestreValue = ClientiRiepilogoVendite.TotaleVendutoCorrente - ClientiRiepilogoVendite.TotaleVendutoRiferimento;
            deltaTrimestre.Text = General.valuta(deltaTrimestreValue);

            double deltaTrimestrePercentoValue = deltaTrimestreValue / ClientiRiepilogoVendite.TotaleVendutoRiferimento;
            deltaTrimestrePercent.Text = General.percentuale(deltaTrimestrePercentoValue);

            General.coloraVariazioni(ClientiRiepilogoVendite.TotaleVendutoCorrente, ClientiRiepilogoVendite.TotaleVendutoRiferimento, totInfo);

            buttonElabora.Background = Brushes.GreenYellow;


        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {


            try
            {
                ElaboraAgente();

            }
            catch (Exception exeption)
            {
                MessageBox.Show(exeption.Message);
            }


            //      });

        }




        void SelezionaTrimestre(object sender, RoutedEventArgs e)
        {
            dataGridVendite.ItemsSource = null;

            RadioButton li = (sender as RadioButton);

            string name = li.Name.ToString();

            string anno = annoCorrente.SelectedItem.ToString();

            string ricercaDa = string.Empty;

            string ricercaFinoA = string.Empty;

            string titoloTimestreText = string.Empty;

            List<string> trimestreSelezionato = new List<string>();

            trimestre = name;


            trimestreSelezionato.Add(name);

            XmlSerializer xmls = new XmlSerializer(typeof(List<string>));

            using (TextWriter writer = new StreamWriter($"trimestreSelezionato.xml"))
            {
                xmls.Serialize(writer, trimestreSelezionato);
            }



            switch (name)
            {
                case "t_1":
                    ricercaDa = $"{anno}-01-01";
                    ricercaFinoA = $"{anno}-03-31";
                    titoloTimestreText = "1° Trimestre";
                    // t_1.IsChecked = true;

                    break;

                case "t_2":
                    ricercaDa = $"{anno}-04-01";
                    ricercaFinoA = $"{anno}-06-30";
                    titoloTimestreText = "2° Trimestre";
                    //t_2.IsChecked = true;
                    break;

                case "t_3":
                    ricercaDa = $"{anno}-07-01";
                    ricercaFinoA = $"{anno}-09-30";
                    titoloTimestreText = "3° Trimestre";
                    //t_3.IsChecked = true;
                    break;

                case "t_4":
                    ricercaDa = $"{anno}-10-01";
                    ricercaFinoA = $"{anno}-12-31";
                    titoloTimestreText = "4° Trimestre";
                    // t_4.IsChecked = true;
                    break;

            }

            fromDate.SelectedDate = DateTime.Parse(ricercaDa);
            toDate.SelectedDate = DateTime.Parse(ricercaFinoA);
            //dataGridVendite.ItemsSource = null;
            li.IsChecked = true;
            titoloTrimestre.Text = titoloTimestreText;
            trimestre = name;

            refreshInfo();
        }



        void refreshInfo()
        {
            totInfo.Text = string.Empty;
            totInfoPrec.Text = string.Empty;
            corr.Text = string.Empty;
            rif.Text = string.Empty;

            dataGridVendite.ItemsSource = null;
            dataGridTrasferiti.ItemsSource = null;
            dataGridCategorieStatCliente.ItemsSource = null;
            nomeClienteCategoriaLabel.Text = string.Empty;
            totInfoProgressivoCorrente.Text = string.Empty;
            totInfoProgressivoRiferimento.Text = string.Empty;
            deltaTrimestre.Text = string.Empty;
            deltaTrimestrePercent.Text = string.Empty;
            totProvvigioneCorrente.Text = string.Empty;
            totProvvigioneTrimestre.Text = string.Empty;
            totProvvigioneSellout.Text = string.Empty;

            deltaProgressivo.Text = string.Empty;
            deltaProgressivoPercent.Text = string.Empty;

            totInfo.Background = null;
            totInfoProgressivoCorrente.Background = null;


        }

        void deselezionaTrimestre()
        {
            t_1.IsChecked = false;
            t_2.IsChecked = false;
            t_3.IsChecked = false;
            t_4.IsChecked = false;

            trimestre = string.Empty;

        }


        private void annoCorrente_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            dataGridVendite.ItemsSource = null;
            string anno = annoCorrente.SelectedItem.ToString();
            string ricercaDa = string.Empty;
            string ricercaFinoA = string.Empty;
            bool checkedTrimestre = false;

            General.directoryTrasferiti(anno);

            if (t_1.IsChecked == true)
            {
                ricercaDa = $"{anno}-01-01";
                ricercaFinoA = $"{anno}-03-31";
                checkedTrimestre = true;
            }


            if (t_2.IsChecked == true)
            {
                ricercaDa = $"{anno}-04-01";
                ricercaFinoA = $"{anno}-06-30";
                checkedTrimestre = true;
            }

            if (t_3.IsChecked == true)
            {
                ricercaDa = $"{anno}-07-01";
                ricercaFinoA = $"{anno}-09-30";
                checkedTrimestre = true;
            }

            if (t_4.IsChecked == true)
            {
                ricercaDa = $"{anno}-10-01";
                ricercaFinoA = $"{anno}-12-31";
                checkedTrimestre = true;
            }

            if (checkedTrimestre)
            {
                fromDate.SelectedDate = DateTime.Parse(ricercaDa);
                toDate.SelectedDate = DateTime.Parse(ricercaFinoA);
            }

        }

        private void dataGridVendite_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            try
            {

                int RowIndex = dataGridVendite.SelectedIndex;

                if (RowIndex >= 0)
                {

                    if (agente.ID.Contains('#')) // TUTTI GLI AGENTI
                    {
                        // legge xml
                        List<Agente> cc = new List<Agente>();
                        XmlSerializer xmlsd = new XmlSerializer(typeof(List<Agente>));
                        using (TextReader tr = new StreamReader(@"agenti.xml"))
                        {
                            cc = (List<Agente>)xmlsd.Deserialize(tr);
                        }

                        var g = AgentiRiepilogo[RowIndex];

                        var Regione = cc.Find(x => x.ID == g.ID).Regione.ToList();

                        // TRASFERITI --------------------------------------------
                        elencoTrasferiti = General.directoryTrasferiti(annoCorrenteTxt);
                        trs = new TrasferitiService(Regione, annoCorrenteTxt, trimestre, elencoTrasferiti);
                        // ------------------------------------------------------

                        double total = trs.Trasferiti.Sum(x => x.ValoreEuro);

                        trs.Trasferiti.Add(new Final() { Fornitore = " - - - TOTALE: ", Valore = total.ToString("C", CultureInfo.CurrentCulture), ValoreEuro = total });

                        dataGridTrasferiti.ItemsSource = trs.Trasferiti;

                        dataGridTrasferiti.Columns[1].CellStyle = HorizontalCenterStyle;

                        dataGridTrasferiti.Columns[0].Width = 190;

                        dataGridTrasferiti.Columns[1].Width = 80;

                        dataGridTrasferiti.Columns[2].Visibility = Visibility.Collapsed;

                        int nRows = dataGridTrasferiti.Items.Count;

                        // dataGridTrasferiti.SelectedIndex = nRows - 2;

                        //  dataGridTrasferiti.CurrentCell = new DataGridCellInfo(dataGridTrasferiti.Items[nRows - 2], dataGridTrasferiti.Columns[1]);

                        // dataGridTrasferiti.SelectedCells.Add(dataGridTrasferiti.CurrentCell);



                        //       dataGridTrasferiti.SelectedCells.FontWeight = FontWeights.Bold;

                        return;
                    }

                    ClienteResponseDatagrid customer = (ClienteResponseDatagrid)dataGridVendite.SelectedItem;
                    string idCliente = customer.IdCliente;
                    string nomeCliente = customer.NomeCliente;

                    nomeClienteCategoriaLabel.Text = nomeCliente;

                    var clickCliente = clienteResponse.Where(x => x.IdCliente == idCliente).ToArray()[0];
                    //List<CategoriaStatisticaDettaglio> cstdL = new List<CategoriaStatisticaDettaglio>();
                    List<GruppoStatistico> cstdL = new List<GruppoStatistico>();

                    //foreach (var item in clickCliente.CategoriaStatistica)
                    foreach (var item in clickCliente.GruppoStatisticoCorrente)
                    {
                        //cstdL.Add(new CategoriaStatisticaDettaglio() { Categoria = item.Categoria, ValoreCorrente = General.valuta(item.ValoreCorrente) });
                        cstdL.Add(new GruppoStatistico() { CKY_MERC = item.CKY_MERC.Trim(' '), CDS_MERC = item.CDS_MERC.Trim(' '), ValoreString = item.Valore.ToString("C", CultureInfo.CurrentCulture) });
                    }

                    dataGridCategorieStatCliente.ItemsSource = cstdL;

                    dataGridCategorieStatCliente.Columns[0].Width = 70;
                    dataGridCategorieStatCliente.Columns[1].Width = 190;

                    dataGridCategorieStatCliente.Columns[2].Visibility = Visibility.Collapsed;

                    dataGridCategorieStatCliente.Columns[0].CellStyle = HorizontalCenterStyle;
                    dataGridCategorieStatCliente.Columns[3].CellStyle = HorizontalCenterStyle;
                }
                else
                {
                    // dataGridCategorieStatCliente.Items.Clear();

                    //Console.WriteLine("sdfsdf");
                }

            }
            catch
            {
                //dataGridCategorieStatCliente.Items.Clear();
            }



        }

        private void DateChange(object sender, SelectionChangedEventArgs e)
        {
            if (fromDate.ToString() != "" && toDate.ToString() != "" && controlDate == true)
            {

            }
            else
            {
                return;
            }

            string da = DateTime.Parse(fromDate.ToString()).ToString("MM-dd");
            string al = DateTime.Parse(toDate.ToString()).ToString("MM-dd");

            bool one = ((da == "01-01") || (da == "04-01") || (da == "07-01") || (da == "10-01"));
            bool two = ((al == "03-31") || (al == "07-31") || (al == "09-30") || (al == "12-31"));

            if (!(one && two))
            {
                deselezionaTrimestre();
            }
        }

        private void annoRiferimento_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            dataGridVendite.ItemsSource = null;
        }

        private void creaExcelButton_Click(object sender, RoutedEventArgs e)
        {
            agente = (Models.Agente)elencoAgenti.SelectedItem;
            if (agente.ID.Contains('#'))
            {
                General.generaExcelTotale(annoCorrenteTxt, annoRiferimentoTxt, trimestre, AgentiRiepilogo);
            }
            else
            {
                General.generaExcelTrasferiti(agente.NikName, agente.Nome, annoCorrenteTxt, annoRiferimentoTxt, trimestre, clienteResponse, trs.Trasferiti, categorieStatistiche, categorieStatisticheTotaleProgressivo);
            }

        }


        private void directoryExcelFinale_Click(object sender, RoutedEventArgs e)
        {
            string path = $"../excelAgenti/";

            string fullPath = System.IO.Path.GetFullPath(path);

            Process.Start("explorer.exe", fullPath);
        }


        private string trimestreSelezionato()
        {
            RadioButton rb = null;

            if (t_1.IsChecked == true)
            {
                rb = t_1;
            }
            else if (t_2.IsChecked == true)
            {
                rb = t_2;
            }
            else if (t_3.IsChecked == true)
            {
                rb = t_3;
            }
            else if (t_4.IsChecked == true)
            {
                rb = t_4;
            }

            if (rb == null)
            {
                return "";
            }




            string path = $"../trasferiti/{annoCorrente.SelectedItem.ToString()}/{rb.Name}";

            return path;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            RadioButton rb = null;

            if (t_1.IsChecked == true)
            {
                rb = t_1;
            }
            else if (t_2.IsChecked == true)
            {
                rb = t_2;
            }
            else if (t_3.IsChecked == true)
            {
                rb = t_3;
            }
            else if (t_4.IsChecked == true)
            {
                rb = t_4;
            }

            if (rb == null)
            {
                return;
            }




            string path = $"../trasferiti/{annoCorrente.SelectedItem.ToString()}/{rb.Name}";

            string fullPath = System.IO.Path.GetFullPath(path);

            File.Create($"{fullPath}/{annoCorrente.SelectedItem.ToString()}--{rb.Name}").Dispose();

            Process.Start("explorer.exe", fullPath);
        }
        Process p = null;
        private void apriImpostazioni_Click(object sender, RoutedEventArgs e)
        {

            Process currentProcess = Process.GetCurrentProcess();
            Process[] localByName = Process.GetProcessesByName("UwAmp");

            if (localByName.Count() == 0)
            {
                string path = "../UwAmp/UwAmp.exe";

                p = new Process();
                p.Exited += new EventHandler(p_Exited);
                p.StartInfo.FileName = System.IO.Path.GetFullPath(path);
                p.EnableRaisingEvents = true;
                p.Start();


            }

            General.OpenUrl("http://127.0.0.1:6768/settingsProvvigioni");

        }




        void p_Exited(object sender, EventArgs e)
        {
            //  MessageBox.Show("Process exited");
            // p.Close();
        }
    }

}