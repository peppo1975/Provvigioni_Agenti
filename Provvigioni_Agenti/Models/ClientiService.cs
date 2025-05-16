using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Xml.Serialization;

namespace Provvigioni_Agenti.Models
{

    public interface IClientiService
    {
        IList<ClienteResponse> ClientiResponse { get; }
        IList<ClienteResponseDatagrid> ClientiResponseDatagrid { get; }
        IList<ClienteRiepilogoVendite> ClientiRiepilogoVendite { get; }
        IList<string> CategorieStatitischeRichiamate { get; }
        IList<CategoriaStatistica> CategorieStatitischeTotale { get; }
        IList<CategoriaStatistica> CategorieStatitischeTotaleProgressivo { get; }

    }
    internal class ClientiService : IClientiService
    {
        private List<ClienteResponse> _clientiResponse = null;

        private List<ClienteResponseDatagrid> _clientiResponseDatagrid = null;
        private List<ClienteRiepilogoVendite> _clientiRiepilogoVendite = null;
        private List<string> _categorieStatitischeRichiamate = null;
        private List<CategoriaStatistica> _categorieStatitischeTotale = null;
        private List<CategoriaStatistica> _categorieStatitischeTotaleProgressivo = null;


        // -------------------------------------------------------------------------
        private List<ClienteResponse> _clientiResponse2 = null;
        private IList<Storico> Storico2 = null;
        private List<ClienteResponseDatagrid> _clientiResponseDatagrid2 = null;
        // -------------------------------------------------------------------------

        public ClientiService(IList<Storico> Storico, Periodo p)
        {
            _clientiResponse = new List<ClienteResponse>();
            _clientiResponse2 = new List<ClienteResponse>();
            _clientiResponseDatagrid = new List<ClienteResponseDatagrid>();
            _clientiRiepilogoVendite = new List<ClienteRiepilogoVendite>();
            _categorieStatitischeRichiamate = new List<string>();
            _categorieStatitischeTotale = new List<CategoriaStatistica>();
            _categorieStatitischeTotaleProgressivo = new List<CategoriaStatistica>();

            double totaleVenduto = 0;
            ClienteRiepilogoVendite crp = new ClienteRiepilogoVendite();


            // -------------------------------------------------------------------------
            var clientiID = Storico.DistinctBy(x => x.CKY_CNT).ToList();
            IList<Storico> idGruppiMerceologici = Storico.DistinctBy(x => x.CKY_MERC).ToList(); // estrapolo tutti i gruppi merceologici nei 2 anni
            Storico2 = Storico;
            _clientiResponseDatagrid2 = new List<ClienteResponseDatagrid>();

            ClienteResponseDatagrid clDg = null;

            foreach (var clienteID in clientiID)
            {
                // ClienteResponse cl = new ClienteResponse() { IdCliente = clienteID.CKY_CNT, NomeCliente = clienteID.CDS_CNT_RAGSOC };

                ClienteResponse cl = estrapola(clienteID.CKY_CNT, p, idGruppiMerceologici);

                cl.IdCliente = clienteID.CKY_CNT;
                cl.NomeCliente = clienteID.CDS_CNT_RAGSOC;
                _clientiResponse2.Add(cl);

                if (cl.TotaleVendutoCorrente != 0 || cl.TotaleVendutoRiferimento != 0)
                {
                    clDg = new ClienteResponseDatagrid();
                    clDg.IdCliente = clienteID.CKY_CNT;
                    clDg.NomeCliente = clienteID.CDS_CNT_RAGSOC;
                    clDg.ProvvigioneCorrente = cl.ProvvigioneCorrente.ToString("C", CultureInfo.CurrentCulture);
                    clDg.TotaleVenduto = cl.TotaleVendutoCorrente.ToString("C", CultureInfo.CurrentCulture);
                    clDg.totaleAnnoPrecedente = cl.TotaleVendutoRiferimento.ToString("C", CultureInfo.CurrentCulture);
                    clDg.Delta = (cl.TotaleVendutoCorrente - cl.TotaleVendutoRiferimento).ToString("C", CultureInfo.CurrentCulture);
                    clDg.DeltaPercento = ((cl.TotaleVendutoCorrente - cl.TotaleVendutoRiferimento) / cl.TotaleVendutoRiferimento).ToString("P2");
                    _clientiResponseDatagrid2.Add(clDg);
                }


            }

            var trimestreCorrente = _clientiResponse2.Sum(x => x.TotaleVendutoCorrente);
            var trimestreRiferimento = _clientiResponse2.Sum(x => x.TotaleVendutoRiferimento);
            var progressivoCorrente = _clientiResponse2.Sum(x => x.TotaleVendutoCorrenteProgressivo);
            var progressivoRiferimento = _clientiResponse2.Sum(x => x.TotaleVendutoRiferimentoProgressivo);
            var provvigione = _clientiResponse2.Sum(x => x.ProvvigioneCorrente);

            var deltaTrimestre = trimestreCorrente - trimestreRiferimento;
            var deltaProgressivo = progressivoCorrente - progressivoRiferimento;

            var deltaTrimestrePercent = deltaTrimestre / trimestreRiferimento;
            var deltaProgressivoPercent = deltaProgressivo / progressivoRiferimento;

            _clientiResponseDatagrid2.Add(new ClienteResponseDatagrid()
            {
                NomeCliente = " * * * * * TOTALE * * * * * ",
                totaleAnnoPrecedente = trimestreRiferimento.ToString("C", CultureInfo.CurrentCulture),
                TotaleVenduto = trimestreCorrente.ToString("C", CultureInfo.CurrentCulture),
                Delta = (_clientiResponse2.Sum(x => x.TotaleVendutoCorrente) - _clientiResponse2.Sum(x => x.TotaleVendutoRiferimento)).ToString("C", CultureInfo.CurrentCulture),
                DeltaPercento = ((_clientiResponse2.Sum(x => x.TotaleVendutoCorrente) - _clientiResponse2.Sum(x => x.TotaleVendutoRiferimento)) / _clientiResponse2.Sum(x => x.TotaleVendutoRiferimento)).ToString("P2"),
                ProvvigioneCorrente = _clientiResponse2.Sum(x => x.ProvvigioneCorrente).ToString("C", CultureInfo.CurrentCulture),
            });


            XmlSerializer xmls = new XmlSerializer(typeof(List<ClienteResponse>));

            using (TextWriter writer = new StreamWriter(@"../_clientiResponse2.xml"))
            {
                //xmls.Serialize(writer, _clientiResponse2.Where(x => ((x.TotaleVendutoCorrente != 0) || (x.TotaleVendutoRiferimento != 0))).ToList());
                xmls.Serialize(writer, _clientiResponse2);
            }

            // -------------------------------------------------------------------------




            foreach (var storico in Storico) // storico è quello che prendo in 1 anno
            {
                var result = _clientiResponse.Find(x => x.IdCliente == storico.CKY_CNT.ToString());

                if (result == null)
                {
                    ClienteResponse c = new ClienteResponse();

                    c.IdCliente = storico.CKY_CNT;
                    c.NomeCliente = storico.CDS_CNT_RAGSOC;

                    _clientiResponse.Add(c);
                }
                else
                {

                }

                ClienteResponse resultItem = _clientiResponse.Find(x => x.IdCliente.Trim() == storico.CKY_CNT.ToString().Trim());

                //storico.DTT_DOC
                //totaleVenduto = Double.Parse(storico.NMP_VALMOV_UM1);
                totaleVenduto = storico.CSG_DOC == "FT" ? Double.Parse(storico.NMP_VALMOV_UM1) : -Double.Parse(storico.NMP_VALMOV_UM1);


                var findCatStat = _categorieStatitischeRichiamate.Contains(storico.CDS_CAT_STAT_ART.Trim(' '));
                if (findCatStat == false)
                {
                    CategorieStatitischeRichiamate.Add(storico.CDS_CAT_STAT_ART.Trim(' '));
                }

                var cstatTot = _categorieStatitischeTotale.Find(x => x.Categoria == storico.CDS_CAT_STAT_ART.Trim(' '));
                if (cstatTot == null)
                {
                    CategoriaStatistica statistica = new CategoriaStatistica();
                    statistica.Categoria = storico.CDS_CAT_STAT_ART.Trim(' ');
                    _categorieStatitischeTotale.Add(statistica);
                }

                if (storico.ANNO == p.annoCorrente.ToString())
                {
                    // progressivo
                    if ((DateTime.Parse(storico.DTT_DOC) <= DateTime.Parse(p.dataFineCorrente)))
                    {
                        crp.ProgressivoCorrente += totaleVenduto;

                        var cStat = resultItem.CategoriaStatisticaProgressiva.Find(x => x.Categoria == storico.CDS_CAT_STAT_ART.Trim(' '));

                        if (cStat == null)
                        {
                            CategoriaStatistica statistica = new CategoriaStatistica();
                            statistica.Categoria = storico.CDS_CAT_STAT_ART.Trim(' ');
                            statistica.ValoreCorrente += totaleVenduto;
                            resultItem.CategoriaStatisticaProgressiva.Add(statistica);
                        }
                        else
                        {
                            cStat.ValoreCorrente += totaleVenduto;
                        }


                        var cstatTotProgr = _categorieStatitischeTotaleProgressivo.Find(x => x.Categoria == storico.CDS_CAT_STAT_ART.Trim(' '));
                        if (cstatTotProgr == null)
                        {
                            CategoriaStatistica statistica = new CategoriaStatistica();
                            statistica.Categoria = storico.CDS_CAT_STAT_ART.Trim(' ');
                            statistica.ValoreCorrente += totaleVenduto;
                            _categorieStatitischeTotaleProgressivo.Add(statistica);
                        }
                        else
                        {
                            cstatTotProgr.ValoreCorrente += totaleVenduto;
                        }

                    }

                    if ((DateTime.Parse(storico.DTT_DOC) >= DateTime.Parse(p.dataInizioCorrente)) && (DateTime.Parse(storico.DTT_DOC) <= DateTime.Parse(p.dataFineCorrente)))
                    {


                        resultItem.ProvvigioneCorrente += storico.CSG_DOC == "FT" ? Double.Parse(storico.NMP_VALPRO_UM1) : -Double.Parse(storico.NMP_VALPRO_UM1);
                        resultItem.TotaleVendutoCorrente += totaleVenduto;



                        var cStat = resultItem.CategoriaStatistica.Find(x => x.Categoria == storico.CDS_CAT_STAT_ART.Trim(' '));
                        if (cStat == null)
                        {
                            CategoriaStatistica statistica = new CategoriaStatistica();
                            statistica.Categoria = storico.CDS_CAT_STAT_ART.Trim(' ');
                            statistica.ValoreCorrente += totaleVenduto;
                            resultItem.CategoriaStatistica.Add(statistica);

                        }
                        else
                        {
                            cStat.ValoreCorrente += totaleVenduto;
                        }


                        cstatTot = _categorieStatitischeTotale.Find(x => x.Categoria == storico.CDS_CAT_STAT_ART.Trim(' '));
                        cstatTot.ValoreCorrente += totaleVenduto;

                    }



                }

                if (storico.ANNO == p.annoRiferimento.ToString())
                {

                    if (DateTime.Parse(storico.DTT_DOC) <= DateTime.Parse(p.dataFineRiferimento))
                    {
                        crp.ProgressivoRiferimento += totaleVenduto;
                    }

                    if ((DateTime.Parse(storico.DTT_DOC) >= DateTime.Parse(p.dataInizioRiferimento)) && (DateTime.Parse(storico.DTT_DOC) <= DateTime.Parse(p.dataFineRiferimento)))
                    {
                        resultItem.ProvvigioneRiferimento += Double.Parse(storico.NMP_VALPRO_UM1);
                        resultItem.TotaleVendutoRiferimento += totaleVenduto;
                    }


                }


            }

            // ----------------------------------------------------------------------------------------------------------
            XmlSerializer xmls1 = new XmlSerializer(typeof(List<ClienteResponse>));

            using (TextWriter writer = new StreamWriter(@"../_clientiResponse.xml"))
            {
                xmls1.Serialize(writer, _clientiResponse);
            }
            // ----------------------------------------------------------------------------------------------------------


            foreach (var t in _clientiResponse)
            {
                ClienteResponseDatagrid cd = new ClienteResponseDatagrid();


                cd.IdCliente = t.IdCliente.ToString();
                cd.NomeCliente = t.NomeCliente.ToString();
                cd.totaleAnnoPrecedente = t.TotaleVendutoRiferimento.ToString("C", CultureInfo.CurrentCulture);
                cd.TotaleVenduto = t.TotaleVendutoCorrente.ToString("C", CultureInfo.CurrentCulture);
                cd.Delta = (t.TotaleVendutoCorrente - t.TotaleVendutoRiferimento).ToString("C", CultureInfo.CurrentCulture);

                cd.DeltaPercento = t.TotaleVendutoRiferimento == 0 ? (t.TotaleVendutoCorrente == 0 ? "" : "∞") : Math.Round((((t.TotaleVendutoCorrente - t.TotaleVendutoRiferimento) / t.TotaleVendutoRiferimento) * 100), 2).ToString("N2") + "%";

                cd.ProvvigioneCorrente = t.ProvvigioneCorrente.ToString("C", CultureInfo.CurrentCulture);
                cd.ProvvigioneRiferimento = t.ProvvigioneRiferimento.ToString("C", CultureInfo.CurrentCulture);
                cd.Percentuale = t.ProvvigioneCorrente == 0 ? "0.00%" : Math.Round((t.ProvvigioneCorrente / t.TotaleVendutoCorrente) * 100, 2).ToString("N2") + "%";

                crp.TotaleVendutoCorrente += t.TotaleVendutoCorrente;
                crp.TotaleVendutoRiferimento += t.TotaleVendutoRiferimento;
                _clientiResponseDatagrid.Add(cd);

            }


            _clientiRiepilogoVendite.Add(crp);
        }



        private ClienteResponse estrapola(string idCliente, Periodo p, IList<Storico> idGruppiMerceologici)
        {
            ClienteResponse res = new ClienteResponse();

            List<GruppoStatisticoDataGrid> gruppoStatisticoDataGridTrimestre = new List<GruppoStatisticoDataGrid>();
            List<GruppoStatisticoDataGrid> gruppoStatisticoDataGridProgressivo = new List<GruppoStatisticoDataGrid>();

            IList<Storico> cliente = Storico2.Where(x => x.CKY_CNT == idCliente).ToList();
            IList<Storico> entrate = cliente.Where(x => x.CSG_DOC == "FT").ToList();
            IList<Storico> entrate_not = cliente.Where(x => x.CSG_DOC != "FT").ToList();


            //anno corrente
            IList<Storico> correnteProgressivo = entrate.Where(x => (DateTime.Parse(x.DTT_DOC) <= DateTime.Parse(p.dataFineCorrente))).ToList();
            IList<Storico> correnteTrimestre = correnteProgressivo.Where(x => (DateTime.Parse(x.DTT_DOC) >= DateTime.Parse(p.dataInizioCorrente))).ToList();

            IList<Storico> correnteProgressivo_not = entrate_not.Where(x => (DateTime.Parse(x.DTT_DOC) <= DateTime.Parse(p.dataFineCorrente))).ToList();
            IList<Storico> correnteTrimestre_not = correnteProgressivo_not.Where(x => (DateTime.Parse(x.DTT_DOC) >= DateTime.Parse(p.dataInizioCorrente))).ToList();


            EstrapolaDatiCliente esvalCorr = estraiValori(cliente, p.dataInizioCorrente, p.dataFineCorrente, p.annoCorrente);

            res.TotaleVendutoCorrente = esvalCorr.totaleVenduto;
            res.TotaleVendutoCorrenteProgressivo = esvalCorr.totaleVendutoProgressivo;
            res.ProvvigioneCorrente = esvalCorr.provvigione;
            res.GruppoStatisticoCorrente = esvalCorr.GruppoStatisticoTrimestre;


            EstrapolaDatiCliente esvalRif = estraiValori(cliente, p.dataInizioRiferimento, p.dataFineRiferimento, p.annoRiferimento);

            res.TotaleVendutoRiferimento = esvalRif.totaleVenduto;
            res.TotaleVendutoRiferimentoProgressivo = esvalRif.totaleVendutoProgressivo;
            res.ProvvigioneRiferimento = esvalRif.provvigione;
            res.GruppoStatisticoRiferimento = esvalCorr.GruppoStatisticoTrimestre;


            foreach (var item in idGruppiMerceologici)
            {
                var grPrCorr = esvalCorr.GruppoStatisticoProgressivo.Find(x => x.CKY_MERC == item.CKY_MERC);
                var grPrRif = esvalRif.GruppoStatisticoProgressivo.Find(x => x.CKY_MERC == item.CKY_MERC);

                var grTrCorr = esvalCorr.GruppoStatisticoTrimestre.Find(x => x.CKY_MERC == item.CKY_MERC);
                var grTrRif = esvalRif.GruppoStatisticoTrimestre.Find(x => x.CKY_MERC == item.CKY_MERC);

                if (grPrCorr != null || grPrRif != null)
                {
                    gruppoStatisticoDataGridProgressivo.Add(new GruppoStatisticoDataGrid() // creo il prtogressivo gruppi statistici per il datagrid
                    {
                        CKY_MERC = item.CKY_MERC,
                        CDS_MERC = item.CDS_MERC,
                        ValoreCorrente = grPrCorr == null ? 0 : grPrCorr.Valore,
                        ValoreRiferimento = grPrRif == null ? 0 : grPrRif.Valore,
                        ValoreCorrenteString = grPrCorr == null ? "" : grPrCorr.Valore.ToString("C", CultureInfo.CurrentCulture),
                        ValoreRiferimentoString = grPrRif == null ? "" : grPrRif.Valore.ToString("C", CultureInfo.CurrentCulture),
                    });

                    gruppoStatisticoDataGridTrimestre.Add(new GruppoStatisticoDataGrid() // creo il prtogressivo gruppi statistici per il datagrid
                    {
                        CKY_MERC = item.CKY_MERC,
                        CDS_MERC = item.CDS_MERC,
                        ValoreCorrente = grTrCorr == null ? 0 : grTrCorr.Valore,
                        ValoreRiferimento = grTrRif == null ? 0 : grTrRif.Valore,
                        ValoreCorrenteString = grTrCorr == null ? "" : grTrCorr.Valore.ToString("C", CultureInfo.CurrentCulture),
                        ValoreRiferimentoString = grTrRif == null ? "" : grTrRif.Valore.ToString("C", CultureInfo.CurrentCulture),
                    });
                }
            }

            res.GruppoStatisticoDataGridProgressivo = gruppoStatisticoDataGridProgressivo;
            res.GruppoStatisticoDataGridTrimestre = gruppoStatisticoDataGridTrimestre;

            return res;
        }

        private EstrapolaDatiCliente estraiValori(IList<Storico> clienteAll, string dataInizio, string dataFine, string anno)
        {
            // viene estratto il trimestre il progressivo delle vendite e dei gruppi statistici
            EstrapolaDatiCliente res = new EstrapolaDatiCliente();

            IList<Storico> cliente = clienteAll.Where(x => x.ANNO == anno).ToList();

            IList<Storico> entrate = cliente.Where(x => x.CSG_DOC == "FT").ToList();
            IList<Storico> entrateNot = cliente.Where(x => x.CSG_DOC != "FT").ToList(); // vengono detratti dalle vendite

            IList<Storico> Progressivo = entrate.Where(x => (DateTime.Parse(x.DTT_DOC) <= DateTime.Parse(dataFine))).ToList();
            IList<Storico> Trimestre = Progressivo.Where(x => (DateTime.Parse(x.DTT_DOC) >= DateTime.Parse(dataInizio))).ToList();


            IList<Storico> ProgressivoNot = entrateNot.Where(x => (DateTime.Parse(x.DTT_DOC) <= DateTime.Parse(dataFine))).ToList();
            IList<Storico> TrimestreNot = ProgressivoNot.Where(x => (DateTime.Parse(x.DTT_DOC) >= DateTime.Parse(dataInizio))).ToList();

            double totaleVenduto = Trimestre.Sum(x => Double.Parse(x.NMP_VALMOV_UM1)) - TrimestreNot.Sum(x => Double.Parse(x.NMP_VALMOV_UM1));
            double totaleVendutoProgressivo = Progressivo.Sum(x => Double.Parse(x.NMP_VALMOV_UM1)) - ProgressivoNot.Sum(x => Double.Parse(x.NMP_VALMOV_UM1));
            double provvigione = Trimestre.Sum(x => Double.Parse(x.NMP_VALPRO_UM1)) - TrimestreNot.Sum(x => Double.Parse(x.NMP_VALPRO_UM1));


            res.totaleVenduto = totaleVenduto;
            res.totaleVendutoProgressivo = totaleVendutoProgressivo;
            res.provvigione = provvigione;


            var groupStatId = cliente.DistinctBy(x => x.CKY_MERC).ToList();

            foreach (var cat in groupStatId)
            {
                var allCat = Trimestre.Where(x => x.CKY_MERC == cat.CKY_MERC).ToList();
                var allcatPos = allCat.Where(x => x.CSG_DOC == "FT");
                var allcatNeg = allCat.Where(x => x.CSG_DOC != "FT");
                double valueTrim = allcatPos.Sum(x => Double.Parse(x.NMP_VALMOV_UM1)) - allcatNeg.Sum(x => Double.Parse(x.NMP_VALMOV_UM1));
                res.GruppoStatisticoTrimestre.Add(new GruppoStatistico() { CKY_MERC = cat.CKY_MERC, CDS_MERC = cat.CDS_MERC, Valore = valueTrim, ValoreString = valueTrim.ToString("C", CultureInfo.CurrentCulture) });



                var allCatProg = Progressivo.Where(x => x.CKY_MERC == cat.CKY_MERC).ToList();
                var allcatProgrPos = allCatProg.Where(x => x.CSG_DOC == "FT");
                var allcatProgrNeg = allCatProg.Where(x => x.CSG_DOC != "FT");
                double valueProgr = allcatProgrPos.Sum(x => Double.Parse(x.NMP_VALMOV_UM1)) - allcatProgrNeg.Sum(x => Double.Parse(x.NMP_VALMOV_UM1));
                res.GruppoStatisticoProgressivo.Add(new GruppoStatistico() { CKY_MERC = cat.CKY_MERC, CDS_MERC = cat.CDS_MERC, Valore = valueProgr, ValoreString = valueProgr.ToString("C", CultureInfo.CurrentCulture) });


            }

            return res;
        }


        //public IList<ClienteResponse> ClientiResponse => _clientiResponse;  //ele
        //public IList<ClienteResponseDatagrid> ClientiResponseDatagrid => _clientiResponseDatagrid.Where(x => ((x.TotaleVenduto != "0,00 €") || (x.totaleAnnoPrecedente != "0,00 €"))).ToList();  //ele

        public IList<ClienteResponse> ClientiResponse => _clientiResponse2;  //ele
        public IList<ClienteResponseDatagrid> ClientiResponseDatagrid => _clientiResponseDatagrid2;  //ele



        public IList<ClienteRiepilogoVendite> ClientiRiepilogoVendite => _clientiRiepilogoVendite;  //ele
        public IList<String> CategorieStatitischeRichiamate => _categorieStatitischeRichiamate;  //ele
        public IList<CategoriaStatistica> CategorieStatitischeTotale => _categorieStatitischeTotale;  //ele
        public IList<CategoriaStatistica> CategorieStatitischeTotaleProgressivo => _categorieStatitischeTotaleProgressivo;  //ele
    }
}
