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

        public AgentiRiepilogoService(IList<Storico> Storico, Periodo p, string[] dirs)
        {
            _agentiRiepilogo = new List<AgenteRiepilogo>();

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


            _agentiRiepilogo.ForEach((x) =>
            {
                //x.VendutoRiferimento = 1.3;
                string id = x.ID;

               var res = ag.Find(x => x.ID == id);

                res.Regione.ForEach((y) =>
                {
                    string nome = y.Nome;
                    var resnum = a.Find(z=>z.Regione == nome);

                    x.VendutoSellout += resnum.Venduto;
                    x.VendutoSelloutString = x.VendutoSellout.ToString("C", CultureInfo.CurrentCulture);

                    x.ProvvigioneSellout = x.VendutoSellout * 0.02;
                    x.ProvvigioneSelloutString = x.ProvvigioneSellout.ToString("C", CultureInfo.CurrentCulture);
                });
            });

        }

        public IList<AgenteRiepilogo> AgentiRiepilogo => _agentiRiepilogo;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}
