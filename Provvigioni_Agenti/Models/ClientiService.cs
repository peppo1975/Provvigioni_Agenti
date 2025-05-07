using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;

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

        //ClienteResponse resultItem = null;
        public ClientiService(IList<Storico> Storico, Periodo p)
        {
            _clientiResponse = new List<ClienteResponse>();
            _clientiResponseDatagrid = new List<ClienteResponseDatagrid>();
            _clientiRiepilogoVendite = new List<ClienteRiepilogoVendite>();
            _categorieStatitischeRichiamate = new List<string>();
            _categorieStatitischeTotale = new List<CategoriaStatistica>();
            _categorieStatitischeTotaleProgressivo = new List<CategoriaStatistica>();

            double totaleVenduto = 0;
            ClienteRiepilogoVendite crp = new ClienteRiepilogoVendite();

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
                        resultItem.TotaleVenduto += totaleVenduto;



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

                        //if (cstatTot == null)
                        //{
                        //    CategoriaStatistica statistica = new CategoriaStatistica();
                        //    statistica.Categoria = storico.CDS_CAT_STAT_ART;
                        //    statistica.Valore += totaleVenduto;
                        //}
                        //else
                        //{
                        //    cstatTot.Valore += totaleVenduto;
                        //}

                        if (storico.CategoriaMerce.ToString() == "5_0")
                        {
                            resultItem.RepSolare += totaleVenduto;
                        }
                        else if (storico.CategoriaMerce.ToString().Trim() == "5_1")
                        {
                            resultItem.QuadriFV += totaleVenduto;
                        }
                        else
                        {
                            resultItem.Commercializzato += totaleVenduto;
                        }

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
                        resultItem.totaleAnnoPrecedente += totaleVenduto;
                    }


                }

                //result2.
            }



            foreach (var t in _clientiResponse)
            {
                ClienteResponseDatagrid cd = new ClienteResponseDatagrid();

                //double maxVal = Math.Max(t.totaleAnnoPrecedente, t.TotaleVenduto);

                cd.IdCliente = t.IdCliente.ToString();
                cd.NomeCliente = t.NomeCliente.ToString();
                cd.totaleAnnoPrecedente = t.totaleAnnoPrecedente.ToString("C", CultureInfo.CurrentCulture);
                cd.TotaleVenduto = t.TotaleVenduto.ToString("C", CultureInfo.CurrentCulture);
                cd.Delta = (t.TotaleVenduto - t.totaleAnnoPrecedente).ToString("C", CultureInfo.CurrentCulture);

                //cd.DeltaPercento = t.TotaleVenduto == 0 ? "-100.00%" : Math.Round((((t.TotaleVenduto - t.totaleAnnoPrecedente) / t.TotaleVenduto) * 100), 2).ToString("N2") + "%";
                //cd.DeltaPercento = t.TotaleVenduto == 0 ? ((t.TotaleVenduto - t.totaleAnnoPrecedente) == 0 ? "-" : "-100%") : Math.Round((((t.TotaleVenduto - t.totaleAnnoPrecedente) / t.TotaleVenduto) * 100), 2).ToString("N2") + "%";



                cd.DeltaPercento = t.totaleAnnoPrecedente == 0 ? (t.TotaleVenduto == 0 ? "" : "∞") : Math.Round((((t.TotaleVenduto - t.totaleAnnoPrecedente) / t.totaleAnnoPrecedente) * 100), 2).ToString("N2") + "%";

                cd.ProvvigioneCorrente = t.ProvvigioneCorrente.ToString("C", CultureInfo.CurrentCulture);
                cd.ProvvigioneRiferimento = t.ProvvigioneRiferimento.ToString("C", CultureInfo.CurrentCulture);
                cd.Percentuale = t.ProvvigioneCorrente == 0 ? "0.00%" : Math.Round((t.ProvvigioneCorrente / t.TotaleVenduto) * 100, 2).ToString("N2") + "%";

                crp.TotaleVendutoCorrente += t.TotaleVenduto;
                crp.TotaleVendutoRiferimento += t.totaleAnnoPrecedente;
                _clientiResponseDatagrid.Add(cd);

            }

            //var fil = _clientiResponseDatagrid.Where(x => (x.TotaleVenduto > 0 && x.totaleAnnoPrecedente > 0));
            //List<ClienteResponseDatagrid> fil = _clientiResponseDatagrid.Where(x => ((x.TotaleVenduto != "0,00 €") && (x.totaleAnnoPrecedente != "0,00 €"))).ToList();

            _clientiRiepilogoVendite.Add(crp);
        }




        public IList<ClienteResponse> ClientiResponse => _clientiResponse;  //ele
        public IList<ClienteResponseDatagrid> ClientiResponseDatagrid => _clientiResponseDatagrid.Where(x => ((x.TotaleVenduto != "0,00 €") || (x.totaleAnnoPrecedente != "0,00 €"))).ToList();  //ele
        public IList<ClienteRiepilogoVendite> ClientiRiepilogoVendite => _clientiRiepilogoVendite;  //ele
        public IList<String> CategorieStatitischeRichiamate => _categorieStatitischeRichiamate;  //ele
        public IList<CategoriaStatistica> CategorieStatitischeTotale => _categorieStatitischeTotale;  //ele
        public IList<CategoriaStatistica> CategorieStatitischeTotaleProgressivo => _categorieStatitischeTotaleProgressivo;  //ele
    }
}
