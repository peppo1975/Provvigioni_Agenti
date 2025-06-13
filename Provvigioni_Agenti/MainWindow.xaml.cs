
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using Provvigioni_Agenti.Controllers;
using Provvigioni_Agenti.Models;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Printing.IndexedProperties;
using System.Reflection.Metadata;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml.Serialization;

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
        List<string> mese = null;
        List<string> elencoTrasferiti = new List<string>();
        TrasferitiService trs = null;
        IList<ClienteResponseDatagrid> clienteResponseDatagrids = null;

        IList<ClienteResponse> clienteResponse = null;
        ClienteRiepilogoVendite ClientiRiepilogoVendite { get; set; }
        IList<CategoriaStatistica> categorieStatistiche = null;
        IList<CategoriaStatistica> categorieStatisticheProgressivo = null;
        IList<CategoriaStatistica> categorieStatisticheTotaleProgressivo = null;
        IList<AgenteRiepilogo> AgentiRiepilogo = null;

        List<GruppoStatistico> GruppoStatistico = null;
        List<GruppoStatisticoRiepilogo> GruppoStatisticoProgessivo = null;
        List<GruppoStatisticoRiepilogo> GruppoStatisticoTrimestre = null;

        List<PeriodoList> trimestreComboBox = new List<PeriodoList>();
        List<PeriodoList> meseComboBox = null;
        PeriodoSelezionato periodoSelezionato = null;
        List<ClientiContactDiretti> clientiContactDiretti = new List<ClientiContactDiretti>();



        Style HorizontalRightStyle = null;

        public MainWindow()
        {
            InitializeComponent();

            General.leggiRegioni();

            General.leggiAgenti();

            apriImpostazioni.Visibility = Visibility.Hidden;

            // ------------------------------------------------------------------------------
            periodoSelezionato = new PeriodoSelezionato(); // xml del periodo selezionato


            trimestreComboBox = new List<PeriodoList>()
            {
                new PeriodoList() { Valore = Trimestri.T1, Testo = "1° TRIM" } ,
                new PeriodoList() { Valore = Trimestri.T2, Testo = "2° TRIM" } ,
                new PeriodoList() { Valore = Trimestri.T3, Testo = "3° TRIM" } ,
                new PeriodoList() { Valore = Trimestri.T4, Testo = "4° TRIM" }
            };

            trimestreComboBox.ForEach(trimestre =>
            {
                trimestreList.Items.Add(trimestre);

            });

            meseComboBox = new List<PeriodoList>()
            {
                new PeriodoList() { Valore = Mesi.Gennaio, Testo = "GENNAIO" } ,
                new PeriodoList() { Valore = Mesi.Febbraio, Testo = "FEBBRAIO" } ,
                new PeriodoList() { Valore = Mesi.Marzo, Testo = "MARZO" } ,
                new PeriodoList() { Valore = Mesi.Aprile, Testo = "APRILE" },
                new PeriodoList() { Valore = Mesi.Maggio, Testo = "MAGGIO" },
                new PeriodoList() { Valore = Mesi.Giugno, Testo = "GIUGNO" },
                new PeriodoList() { Valore = Mesi.Luglio, Testo = "LUGLIO" },
                new PeriodoList() { Valore = Mesi.Agosto, Testo = "AGOSTO" },
                new PeriodoList() { Valore = Mesi.Settembre, Testo = "SETTEMBRE" },
                new PeriodoList() { Valore = Mesi.Ottobre, Testo = "OTTOBRE" },
                new PeriodoList() { Valore = Mesi.Novembre, Testo = "NOVEMBRE" },
                new PeriodoList() { Valore = Mesi.Dicembre, Testo = "DICEMBRE" },
            };
            meseComboBox.ForEach(x =>
            {
                meseList.Items.Add(x);
            });


            periodoSelezionato = General.periodoHome();

            // ------------------------------------------------------------------------------



            List<Anno> A = new List<Anno>();

            A = Controllers.Database.SELECT_GET_LIST<Anno>("SELECT NGB_ANNO_DOC FROM MMA_M GROUP BY NGB_ANNO_DOC  ORDER BY NGB_ANNO_DOC DESC");




            foreach (Anno anno in A)
            {
                annoCorrente.Items.Add(anno.NGB_ANNO_DOC);
                annoRiferimento.Items.Add(anno.NGB_ANNO_DOC);
            }

            if (periodoSelezionato.AnnoCorrente == string.Empty)
                annoCorrente.SelectedItem = A.Max(t => t.NGB_ANNO_DOC);


            if (periodoSelezionato.AnnoRiferimento == string.Empty)
                annoRiferimento.SelectedItem = A.Max(t => t.NGB_ANNO_DOC) - 1;

            annoCorrenteTxt = A.Max(t => t.NGB_ANNO_DOC).ToString(); //annoCorrente.SelectedItem.ToString();
            annoRiferimentoTxt = (A.Max(t => t.NGB_ANNO_DOC) - 1).ToString(); //annoCorrente.SelectedItem.ToString();

            agentiService = new Models.AgentiService();

            elencoAgenti.ItemsSource = agentiService.Agenti;

            elencoAgenti.DisplayMemberPath = "NikName";

            controlDate = true;

            buttonElabora.IsEnabled = false;

            General.directoryTrasferiti(A.Max(t => t.NGB_ANNO_DOC).ToString());

            HorizontalRightStyle = new Style();

            HorizontalRightStyle.Setters.Add(new Setter(HorizontalAlignmentProperty, HorizontalAlignment.Right));

            // vedo se sono salvate delle date -----------------------------------

            if (periodoSelezionato.AnnoCorrente != string.Empty)
            {
                annoCorrente.SelectedItem = Int32.Parse(periodoSelezionato.AnnoCorrente);
            }

            if (periodoSelezionato.AnnoRiferimento != string.Empty)
            {
                annoRiferimento.SelectedItem = Int32.Parse(periodoSelezionato.AnnoRiferimento);
            }


            var m = meseComboBox.FindIndex(e => e.Valore == periodoSelezionato.Mese);
            var t = trimestreComboBox.FindIndex(e => e.Valore == periodoSelezionato.Trimestre);

            trimestreList.SelectedIndex = t;
            meseList.SelectedIndex = m;

            // -------------------------------------------------------------------





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

            mese = new List<string>();

            mese = General.cercaMese(meseList, trimestreList);

            General.generaXmlCitta();

            if (!initCalc)
                return;

            refreshInfo();

            annoCorrenteTxt = annoCorrente.SelectedItem.ToString();

            annoRiferimentoTxt = annoRiferimento.SelectedItem.ToString();

            // TRASFERITI --------------------------------------------
            elencoTrasferiti = General.directoryTrasferiti(annoCorrenteTxt);
            trs = new TrasferitiService(agente.Regione, annoCorrenteTxt, trimestre, elencoTrasferiti, mese);
            // -------------------------------------------------------

            clientiContactDiretti = General.ClientiContactDiretti(annoCorrenteTxt, annoRiferimentoTxt);

            string query = Controllers.Query.clientiAgente(agente.ID, annoCorrenteTxt, annoRiferimentoTxt, clientiContactDiretti);
            var st = Controllers.Database.SELECT_GET_LIST<Storico>(query);




            string fromDateTxt = fromDate.ToString();
            string toDateTxt = toDate.ToString();

            storicoService = new Models.StoricoService();
            periodoService = new Models.PeriodoService(annoCorrente.SelectedItem.ToString(), annoRiferimento.SelectedItem.ToString(), fromDate, toDate);

            if (agente.ID.Contains('#')) // TUTTI GLI AGENTI -------------------------------------------------------
            {

                List<string> listDirs = new List<string>();

                foreach (var item in mese)
                {
                    string[] dirs2 = Directory.GetDirectories($"../trasferiti/{annoCorrente.SelectedItem.ToString()}/{item}");
                    listDirs.AddRange(dirs2);
                }



                // string[] dirs = Directory.GetDirectories(periodoSelezionatoL()); // intervallo mesi selezionato (non più trimestre)
                var elencoRiepilogo = new Models.AgentiRiepilogoService(st, periodoService.Periodo[0], listDirs);
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
                    dataGridVendite.Columns[i].CellStyle = HorizontalRightStyle;
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

            GruppoStatistico = (List<GruppoStatistico>)elenco.GruppoStatistico;
            GruppoStatisticoProgessivo = (List<GruppoStatisticoRiepilogo>)elenco.GruppoStatisticoProgressivo;
            GruppoStatisticoTrimestre = (List<GruppoStatisticoRiepilogo>)elenco.GruppoStatisticoTrimestre;

            ClientiRiepilogoVendite = elenco.ClientiRiepilogoVendite[0];
            categorieStatisticheTotaleProgressivo = elenco.CategorieStatitischeTotaleProgressivo;


            double totalTrasferiti = trs.Trasferiti.Sum(x => x.ValoreEuro);

            //trs.Trasferiti.Add(new Final() { Fornitore = " - - - TOTALE: ", Valore = total.ToString("C", CultureInfo.CurrentCulture), ValoreEuro = total });

            var trsTrasferiti = trs.Trasferiti;

            trsTrasferiti.Add(new Final() { Fornitore = " - - - TOTALE: ", Valore = totalTrasferiti.ToString("C", CultureInfo.CurrentCulture), ValoreEuro = totalTrasferiti });

            totSellout.Text = totalTrasferiti.ToString("C", CultureInfo.CurrentCulture);

            dataGridTrasferiti.ItemsSource = trsTrasferiti;

            dataGridTrasferiti.Columns[1].CellStyle = HorizontalRightStyle;

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
                dataGridVendite.Columns[i].CellStyle = HorizontalRightStyle;
            }

            Console.WriteLine(dataGridVendite.Items.Count);

            //   var LastRow = dataGridVendite.Items[dataGridVendite.Items.Count - 1];
            // LastRow.Font.Bold = true;


            corr.Text = annoCorrenteTxt;

            totInfo.Text = General.valuta(ClientiRiepilogoVendite.TotaleVendutoCorrente);
            totInfoProgressivoCorrente.Text = General.valuta(ClientiRiepilogoVendite.ProgressivoCorrente);

            double provvigioneCorrente = clienteResponse.Sum(x => x.ProvvigioneCorrente);
            totProvvigioneCorrente.Text = General.valuta(provvigioneCorrente);


            double provvigioneTrasferiti = totalTrasferiti * 0.05;
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

                creaExcelButton.IsEnabled = true;

            }
            catch (Exception exeption)
            {
                MessageBox.Show(exeption.Message);
            }


        }





        void refreshInfo()
        {
            totInfo.Text = string.Empty;
            totInfoPrec.Text = string.Empty;
            corr.Text = string.Empty;
            rif.Text = string.Empty;

            dataGridVendite.ItemsSource = null;
            dataGridTrasferiti.ItemsSource = null;
            dataGridGruppiStatisticiProgressivoCliente.ItemsSource = null;
            dataGridGruppiStatisticiTrimestreCliente.ItemsSource = null;
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
            creaExcelButton.IsEnabled = false;
            totSellout.Text = string.Empty;
        }

        void deselezionaTrimestre()
        {
            trimestreList.Text = "";
            meseList.Text = "";
            titoloTrimestre.Text = "";
        }


        private void annoCorrente_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string anno = annoCorrente.SelectedItem.ToString();
      
            periodoSelezionato.AnnoCorrente = anno;

            General.periodoHomeSave(periodoSelezionato);

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

                        dataGridTrasferiti.Columns[1].CellStyle = HorizontalRightStyle;

                        dataGridTrasferiti.Columns[0].Width = 190;

                        dataGridTrasferiti.Columns[1].Width = 80;

                        dataGridTrasferiti.Columns[2].Visibility = Visibility.Collapsed;

                        int nRows = dataGridTrasferiti.Items.Count;

                        return;
                    }

                    ClienteResponseDatagrid customer = (ClienteResponseDatagrid)dataGridVendite.SelectedItem;
                    string idCliente = customer.IdCliente;
                    string nomeCliente = customer.NomeCliente;

                    nomeClienteCategoriaLabel.Text = nomeCliente;

                    var clickCliente = clienteResponse.Where(x => x.IdCliente == idCliente).ToArray()[0];
                    //List<CategoriaStatisticaDettaglio> cstdL = new List<CategoriaStatisticaDettaglio>();
                    //List<GruppoStatistico> cstdL = new List<GruppoStatistico>();
                    List<GruppoStatisticoDataGrid> grpStatProgr = new List<GruppoStatisticoDataGrid>();

                    //foreach (var item in clickCliente.CategoriaStatistica)
                    foreach (var item in clickCliente.GruppoStatisticoDataGridProgressivo)
                    {

                        if ((double)item.ValoreRiferimento == 0 && (double)item.ValoreCorrente == 0)
                        {
                            continue;
                        }

                        grpStatProgr.Add(new GruppoStatisticoDataGrid()
                        {
                            CKY_MERC = item.CKY_MERC.Trim(' '),
                            CDS_MERC = item.CDS_MERC.Trim(' '),
                            ValoreRiferimentoString = item.ValoreRiferimento.ToString("C", CultureInfo.CurrentCulture),
                            ValoreCorrenteString = item.ValoreCorrente.ToString("C", CultureInfo.CurrentCulture)
                        });
                    }

                    dataGridGruppiStatisticiProgressivoCliente.ItemsSource = grpStatProgr;

                    dataGridGruppiStatisticiProgressivoCliente.Columns[0].Width = 70;
                    dataGridGruppiStatisticiProgressivoCliente.Columns[1].Width = 190;

                    dataGridGruppiStatisticiProgressivoCliente.Columns[2].Visibility = Visibility.Collapsed;
                    dataGridGruppiStatisticiProgressivoCliente.Columns[3].Visibility = Visibility.Collapsed;

                    dataGridGruppiStatisticiProgressivoCliente.Columns[0].CellStyle = HorizontalRightStyle;
                    dataGridGruppiStatisticiProgressivoCliente.Columns[4].CellStyle = HorizontalRightStyle;
                    dataGridGruppiStatisticiProgressivoCliente.Columns[5].CellStyle = HorizontalRightStyle;

                    dataGridGruppiStatisticiProgressivoCliente.Columns[4].Header = annoRiferimentoTxt;
                    dataGridGruppiStatisticiProgressivoCliente.Columns[5].Header = annoCorrenteTxt;

                    dataGridGruppiStatisticiProgressivoCliente.Columns[4].Width = 90;
                    dataGridGruppiStatisticiProgressivoCliente.Columns[5].Width = 90;



                    List<GruppoStatisticoDataGrid> grpStatTrim = new List<GruppoStatisticoDataGrid>();

                    //foreach (var item in clickCliente.CategoriaStatistica)
                    foreach (var item in clickCliente.GruppoStatisticoDataGridTrimestre)
                    {

                        if ((double)item.ValoreRiferimento == 0 && (double)item.ValoreCorrente == 0)
                        {
                            continue;
                        }

                        grpStatTrim.Add(new GruppoStatisticoDataGrid()
                        {
                            CKY_MERC = item.CKY_MERC.Trim(' '),
                            CDS_MERC = item.CDS_MERC.Trim(' '),
                            ValoreRiferimentoString = item.ValoreRiferimento.ToString("C", CultureInfo.CurrentCulture),
                            ValoreCorrenteString = item.ValoreCorrente.ToString("C", CultureInfo.CurrentCulture)
                        });
                    }

                    dataGridGruppiStatisticiTrimestreCliente.ItemsSource = grpStatTrim;

                    dataGridGruppiStatisticiTrimestreCliente.Columns[0].Width = 70;
                    dataGridGruppiStatisticiTrimestreCliente.Columns[1].Width = 190;

                    dataGridGruppiStatisticiTrimestreCliente.Columns[2].Visibility = Visibility.Collapsed;
                    dataGridGruppiStatisticiTrimestreCliente.Columns[3].Visibility = Visibility.Collapsed;

                    dataGridGruppiStatisticiTrimestreCliente.Columns[0].CellStyle = HorizontalRightStyle;
                    dataGridGruppiStatisticiTrimestreCliente.Columns[4].CellStyle = HorizontalRightStyle;
                    dataGridGruppiStatisticiTrimestreCliente.Columns[5].CellStyle = HorizontalRightStyle;

                    dataGridGruppiStatisticiTrimestreCliente.Columns[4].Header = annoRiferimentoTxt;
                    dataGridGruppiStatisticiTrimestreCliente.Columns[5].Header = annoCorrenteTxt;

                    dataGridGruppiStatisticiTrimestreCliente.Columns[4].Width = 90;
                    dataGridGruppiStatisticiTrimestreCliente.Columns[5].Width = 90;


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

            bool one = ((da == "01-01") || (da == "02-01") || (da == "03-01") || (da == "04-01") || (da == "05-01") || (da == "06-01") || (da == "07-01") || (da == "08-01") || (da == "09-01") || (da == "10-01") || (da == "11-01") || (da == "12-01"));
            bool two = ((al == "01-31") || (al == "02-28") || (al == "02-29") || (al == "03-31") || (al == "04-30") || (al == "05-31") || (al == "06-30") || (al == "07-31") || (al == "08-31") || (al == "09-30") || (al == "10-31") || (al == "11-30") || (al == "12-31"));

            if (!(one && two))
            {
                deselezionaTrimestre();
            }
        }

        private void annoRiferimento_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            dataGridVendite.ItemsSource = null;

      
            periodoSelezionato.AnnoRiferimento = annoRiferimentoTxt = annoRiferimento.SelectedItem.ToString();

            General.periodoHomeSave(periodoSelezionato);
        }

        private void creaExcelButton_Click(object sender, RoutedEventArgs e)
        {
            agente = (Models.Agente)elencoAgenti.SelectedItem;

            Console.WriteLine(mese);
            if (agente.ID.Contains('#'))
            {   // tutti gli agenti
                General.generaExcelTotale(annoCorrenteTxt, annoRiferimentoTxt, trimestre, AgentiRiepilogo, GruppoStatistico);
            }
            else
            {   // singolo agente
                General.generaExcelTrasferiti(agente, annoCorrenteTxt, annoRiferimentoTxt, mese, clienteResponse, trs.Trasferiti, GruppoStatistico, GruppoStatisticoProgessivo, GruppoStatisticoTrimestre);
            }

        }


        private void directoryExcelFinale_Click(object sender, RoutedEventArgs e)
        {
            string path = $"../excelAgenti/";

            string fullPath = System.IO.Path.GetFullPath(path);

            Process.Start("explorer.exe", fullPath);
        }


        private string periodoSelezionatoL()
        {
            RadioButton rb = null;

            //if (t_1.IsChecked == true)
            //{
            //    rb = t_1;
            //}
            //else if (t_2.IsChecked == true)
            //{
            //    rb = t_2;
            //}
            //else if (t_3.IsChecked == true)
            //{
            //    rb = t_3;
            //}
            //else if (t_4.IsChecked == true)
            //{
            //    rb = t_4;
            //}

            //if (rb == null)
            //{
            //    return "";
            //}

            string path = $"../trasferiti/{annoCorrente.SelectedItem.ToString()}/{rb.Name}";

            return path;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            RadioButton rb = null;

            //if (t_1.IsChecked == true)
            //{
            //    rb = t_1;
            //}
            //else if (t_2.IsChecked == true)
            //{
            //    rb = t_2;
            //}
            //else if (t_3.IsChecked == true)
            //{
            //    rb = t_3;
            //}
            //else if (t_4.IsChecked == true)
            //{
            //    rb = t_4;
            //}

            //if (rb == null)
            //{
            //    return;
            //}

            string path = string.Empty;

            var mese = meseList.SelectedItem;
            var trimeste = trimestreList.SelectedItem;

            if (mese != null)
            {


                int selectedIndex = meseList.SelectedIndex;
                PeriodoList selectedValue = (PeriodoList)meseList.SelectedValue;

                path = $"../trasferiti/{annoCorrente.SelectedItem.ToString()}/{selectedValue.Valore.ToString()}";
            }

            if (trimeste != null)
            {
                path = $"../trasferiti/{annoCorrente.SelectedItem.ToString()}";
            }




            //string path = $"../trasferiti/{annoCorrente.SelectedItem.ToString()}/{rb.Name}";


            string fullPath = System.IO.Path.GetFullPath(path);

            if (mese != null)
            {
                int selectedIndex = meseList.SelectedIndex;
                PeriodoList selectedValue = (PeriodoList)meseList.SelectedValue;

                File.Create($"{fullPath}/{annoCorrente.SelectedItem.ToString()}_{selectedValue.Valore.ToString()}").Dispose();
            }
            if (trimeste != null)
            {
                File.Create($"{fullPath}/{annoCorrente.SelectedItem.ToString()}").Dispose();
            }


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

        private void trimestreList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            ComboBox cmb = (ComboBox)sender;
            int selectedIndex = cmb.SelectedIndex;
            PeriodoList selectedValue = (PeriodoList)cmb.SelectedValue;

            if (selectedValue == null)
            {
                return;
            }

            dataGridVendite.ItemsSource = null;

            string name = selectedValue.Valore.ToString();

            string anno = annoCorrente.SelectedItem.ToString();

            string ricercaDa = string.Empty;

            string ricercaFinoA = string.Empty;

            string titoloTimestreText = string.Empty;

            List<string> trimestreSelezionato = new List<string>();

            trimestre = name;

            string numTrimestre = name.Split("_")[1];

            ricercaDa = DateTime.Today.ToString("d");
            ricercaFinoA = DateTime.Today.ToString("d");
            titoloTimestreText = DateTime.Today.ToString("d");

            int end = Int32.Parse(numTrimestre) * 3; // mese finale trimestre

            if (end > 0)
            {
                int init = end - 2; // mese iniziale trimestre
                var days = DateTime.DaysInMonth(Int32.Parse(anno), end);

                ricercaDa = $"{anno}-{init}-01";
                ricercaFinoA = $"{anno}-{end}-{days}";
                titoloTimestreText = $"{selectedValue.Testo}";
            }

            XmlSerializer xmls = new XmlSerializer(typeof(List<string>));

            using (TextWriter writer = new StreamWriter($"periodoSelezionato.xml"))
            {
                xmls.Serialize(writer, trimestreSelezionato);
            }

            fromDate.SelectedDate = DateTime.Parse(ricercaDa);
            toDate.SelectedDate = DateTime.Parse(ricercaFinoA);

            dataGridVendite.ItemsSource = null;

            titoloTrimestre.Text = titoloTimestreText;
            trimestre = name;

            periodoSelezionato.Trimestre = name;
            periodoSelezionato.Mese = "";

            General.periodoHomeSave(periodoSelezionato);

            refreshInfo();

            meseList.Text = "";

        }

        private void meseList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cmb = (ComboBox)sender;
            int selectedIndex = cmb.SelectedIndex;
            PeriodoList selectedValue = (PeriodoList)cmb.SelectedValue;

            if (selectedValue == null)
            {
                return;
            }


            bool inserisciData = true;

            dataGridVendite.ItemsSource = null;

            string name = selectedValue.Valore.ToString();

            string anno = annoCorrente.SelectedItem.ToString();

            string ricercaDa = string.Empty;

            string ricercaFinoA = string.Empty;

            string titoloTimestreText = string.Empty;

            List<string> trimestreSelezionato = new List<string>();

            string mese = name.Split("_")[1];

            var days = DateTime.DaysInMonth(Int32.Parse(anno), Int32.Parse(mese));

            XmlSerializer xmls = new XmlSerializer(typeof(List<string>));

            using (TextWriter writer = new StreamWriter($"meseSelezionato.xml"))
            {
                xmls.Serialize(writer, trimestreSelezionato);
            }

            ricercaDa = $"{anno}-{mese}-01";
            ricercaFinoA = $"{anno}-{mese}-{days}";
            titoloTimestreText = $"{selectedValue.Testo}";



            fromDate.SelectedDate = DateTime.Parse(ricercaDa);
            toDate.SelectedDate = DateTime.Parse(ricercaFinoA);


            dataGridVendite.ItemsSource = null;
            titoloTrimestre.Text = titoloTimestreText;
            periodoSelezionato.Mese = name;
            periodoSelezionato.Trimestre = "";

            General.periodoHomeSave(periodoSelezionato);

            refreshInfo();

            trimestreList.Text = "";

        }
    }

}