using DocumentFormat.OpenXml.Wordprocessing;
using Provvigioni_Agenti.Controllers;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Provvigioni_Agenti.Models
{
    public interface IAgentiRiepilogoService
    {
        IList<AgenteRiepilogo> AgentiRiepilogo { get; }
    }
    internal class AgentiRiepilogoService : IAgentiRiepilogoService
    {
        private List<AgenteRiepilogo> _agentiRiepilogo = null;
        private List<AgenteRiepilogo> _agentiRiepilogo2 = null;

        public AgentiRiepilogoService(IList<Storico> Storico, Periodo p, string[] dirs)
        {
            _agentiRiepilogo = new List<AgenteRiepilogo>();
            _agentiRiepilogo2 = new List<AgenteRiepilogo>();

        https://josipmisko.com/posts/c-sharp-unique-list
            var agentiId = Storico.DistinctBy(x => x.CKY_CNT_AGENTE).ToList();

            foreach (var item in agentiId)
            {
                // _agentiRiepilogo2.Add(filtra((List<Storico>)Storico, p, item.CKY_CNT_AGENTE, (Storico)item));

                _agentiRiepilogo.Add(new AgenteRiepilogo() { ID = item.CKY_CNT_AGENTE, Nome = item.NomeAgente });
            }

            foreach (var item in Storico)
            {

                var venduto = item.CSG_DOC == "FT" ? Double.Parse(item.NMP_VALMOV_UM1) : -Double.Parse(item.NMP_VALMOV_UM1);
                var provvigione = item.CSG_DOC == "FT" ? Double.Parse(item.NMP_VALPRO_UM1) : -Double.Parse(item.NMP_VALPRO_UM1);

                var response = _agentiRiepilogo.Find(x => x.ID == item.CKY_CNT_AGENTE);
                if (response == null)
                {

                    _agentiRiepilogo.Add(new AgenteRiepilogo() { ID = item.CKY_CNT_AGENTE, Nome = item.NomeAgente });
                }
                response = _agentiRiepilogo.Find(x => x.ID == item.CKY_CNT_AGENTE);

                if (item.ANNO == p.annoCorrente)
                {

                    if ((DateTime.Parse(item.DTT_DOC) >= DateTime.Parse(p.dataInizioCorrente)) && (DateTime.Parse(item.DTT_DOC) <= DateTime.Parse(p.dataFineCorrente)))
                    {
                        response.VendutoCorrente += venduto;
                        response.ProvvigioneCorrente += provvigione;
                    }


                }

                if (item.ANNO == p.annoRiferimento)
                {
                    if ((DateTime.Parse(item.DTT_DOC) >= DateTime.Parse(p.dataInizioRiferimento)) && (DateTime.Parse(item.DTT_DOC) <= DateTime.Parse(p.dataFineRiferimento)))
                    {
                        response.VendutoRiferimento += venduto;
                    }
                }

                response.Delta = response.VendutoCorrente - response.VendutoRiferimento;

                response.DeltaPercent = response.Delta / response.VendutoRiferimento;


                response.VendutoCorrenteString = response.VendutoCorrente.ToString("C", CultureInfo.CurrentCulture);
                response.VendutoRiferimentoString = response.VendutoRiferimento.ToString("C", CultureInfo.CurrentCulture);
                response.ProvvigioneCorrenteString = response.ProvvigioneCorrente.ToString("C", CultureInfo.CurrentCulture);
                response.DeltaString = response.Delta.ToString("C", CultureInfo.CurrentCulture);
                response.DeltaPercentString = String.Format("{0:P2}", response.DeltaPercent);

            }


            // legge xml
            List<Agente> ag = null;
            XmlSerializer xmlsd = new XmlSerializer(typeof(List<Agente>));
            using (TextReader tr = new StreamReader(@"agenti.xml"))
            {
                ag = (List<Agente>)xmlsd.Deserialize(tr);
            }

            var itemToRemove = ag.Single(r => r.Nome == "TUTTI GLI AGENTI");
            ag.Remove(itemToRemove);

            List<Trasferito> a = General.estraiXmlSellout(dirs);



            foreach (var item in agentiId)
            {
                List<Regione> agenteRegione = ag.Find(r => r.ID == item.CKY_CNT_AGENTE).Regione.ToList();

                double sellout = 0;

                agenteRegione.ForEach((r) =>
                {
                    Trasferito resnum = a.Find(z => z.Regione == r.Nome);
                    sellout += resnum.Venduto;
                });

                _agentiRiepilogo2.Add(filtra((List<Storico>)Storico, p, item.CKY_CNT_AGENTE, (Storico)item, sellout));
            }

            _agentiRiepilogo.ForEach((x) =>
        {
            //x.VendutoRiferimento = 1.3;
            string id = x.ID;

            var res = ag.Find(x => x.ID == id);

            res.Regione.ForEach((y) =>
            {
                string nome = y.Nome;
                var resnum = a.Find(z => z.Regione == nome);

                x.VendutoSellout += resnum.Venduto;
                x.VendutoSelloutString = x.VendutoSellout.ToString("C", CultureInfo.CurrentCulture);

                x.ProvvigioneSellout = x.VendutoSellout * 0.02;
                x.ProvvigioneSelloutString = x.ProvvigioneSellout.ToString("C", CultureInfo.CurrentCulture);
            });
        });

        }


        private AgenteRiepilogo filtra(List<Storico> Storico, Periodo p, string AgenteId, Storico item, double sellout)
        {
            // new AgenteRiepilogo() { ID = item.CKY_CNT_AGENTE, Nome = item.NomeAgente }

            AgenteRiepilogo res = new AgenteRiepilogo();

            List<Storico> aCorrPos = Storico.Where((x) => x.CKY_CNT_AGENTE.ToString() == item.CKY_CNT_AGENTE && x.ANNO.ToString() == p.annoCorrente.ToString() && (DateTime.Parse(x.DTT_DOC) >= DateTime.Parse(p.dataInizioCorrente)) && (DateTime.Parse(x.DTT_DOC) <= DateTime.Parse(p.dataFineCorrente)) && x.CSG_DOC == "FT").ToList();
            List<Storico> aCorrNeg = Storico.Where((x) => x.CKY_CNT_AGENTE.ToString() == item.CKY_CNT_AGENTE && x.ANNO.ToString() == p.annoCorrente.ToString() && (DateTime.Parse(x.DTT_DOC) >= DateTime.Parse(p.dataInizioCorrente)) && (DateTime.Parse(x.DTT_DOC) <= DateTime.Parse(p.dataFineCorrente)) && x.CSG_DOC != "FT").ToList();

            double vendutoCorrente = aCorrPos.Sum(x => (Double.Parse(x.NMP_VALMOV_UM1))) - aCorrNeg.Sum(x => (Double.Parse(x.NMP_VALMOV_UM1)));
            double provvigioneCorrente = aCorrPos.Sum(x => (Double.Parse(x.NMP_VALPRO_UM1))) - aCorrNeg.Sum(x => (Double.Parse(x.NMP_VALPRO_UM1)));

            List<Storico> aRifPos = Storico.Where((x) => x.CKY_CNT_AGENTE.ToString() == item.CKY_CNT_AGENTE && x.ANNO.ToString() == p.annoRiferimento.ToString() && (DateTime.Parse(x.DTT_DOC) >= DateTime.Parse(p.dataInizioRiferimento)) && (DateTime.Parse(x.DTT_DOC) <= DateTime.Parse(p.dataFineRiferimento)) && x.CSG_DOC == "FT").ToList();
            List<Storico> aRifNeg = Storico.Where((x) => x.CKY_CNT_AGENTE.ToString() == item.CKY_CNT_AGENTE && x.ANNO.ToString() == p.annoRiferimento.ToString() && (DateTime.Parse(x.DTT_DOC) >= DateTime.Parse(p.dataInizioRiferimento)) && (DateTime.Parse(x.DTT_DOC) <= DateTime.Parse(p.dataFineRiferimento)) && x.CSG_DOC != "FT").ToList();

            double vendutoRiferimento = aRifPos.Sum(x => (Double.Parse(x.NMP_VALMOV_UM1))) - aRifNeg.Sum(x => (Double.Parse(x.NMP_VALMOV_UM1)));
            //double provvigioneRiferimento = aRifPos.Sum(x => (Double.Parse(x.NMP_VALPRO_UM1))) - aRifNeg.Sum(x => (Double.Parse(x.NMP_VALPRO_UM1)));

            res.ID = item.CKY_CNT_AGENTE;
            res.Nome = item.NomeAgente;
            res.VendutoCorrente = vendutoCorrente;
            res.VendutoRiferimento = vendutoRiferimento;
            res.ProvvigioneCorrente = provvigioneCorrente;
            res.Delta = vendutoCorrente - vendutoRiferimento;
            res.DeltaPercent = res.Delta / vendutoRiferimento;
            res.VendutoSellout = sellout;
            res.ProvvigioneSellout = sellout * 0.02;


            res.VendutoCorrenteString = res.VendutoCorrente.ToString("C", CultureInfo.CurrentCulture);
            res.VendutoRiferimentoString = res.VendutoRiferimento.ToString("C", CultureInfo.CurrentCulture);
            res.ProvvigioneCorrenteString = res.ProvvigioneCorrente.ToString("C", CultureInfo.CurrentCulture);
            res.VendutoSelloutString = res.VendutoSellout.ToString("C", CultureInfo.CurrentCulture);
            res.ProvvigioneSelloutString = res.ProvvigioneSellout.ToString("C", CultureInfo.CurrentCulture);
            res.ProvvigioneCorrenteString = res.ProvvigioneCorrente.ToString("C", CultureInfo.CurrentCulture);
            res.DeltaString = res.Delta.ToString("C", CultureInfo.CurrentCulture);
            res.DeltaPercentString = String.Format("{0:P2}", res.DeltaPercent);

            //return new AgenteRiepilogo() { ID = item.CKY_CNT_AGENTE, Nome = item.NomeAgente, VendutoCorrente = vendutoCorrente, VendutoRiferimento = vendutoRiferimento };
            return res;

        }

        public IList<AgenteRiepilogo> AgentiRiepilogo => _agentiRiepilogo2;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}
