using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Bibliography;
using Provvigioni_Agenti.Controllers;

namespace Provvigioni_Agenti.Models
{
    internal class ElaborazioniService
    {
    }

    public interface IStoricoService
    {
        IList<Storico> Storico { get; }
    }

    internal class StoricoService : IStoricoService
    {
        private List<Storico> _storico = null;
        private List<StoricoTotal> _storicoTotal = null;

        public StoricoService()
        {

            _storico = new List<Storico>();

            //_agenti.Add(new Agente() { Nome = "Peppe", Regione = "Puglia", NikName = "GiuP", ID = "0001" });

            XmlSerializer xmlsd = new XmlSerializer(typeof(List<Storico>));

            using (TextReader tr = new StreamReader(@"Xml.xml"))
            {
                _storico = (List<Storico>)xmlsd.Deserialize(tr);
            }

        }

        public IList<Storico> Storico => _storico;  //elemento pubblico che da modo di visualizzare un elemento privato
        public IList<StoricoTotal> StoricoTotal => _storicoTotal;  //elemento pubblico che da modo di visualizzare un elemento privato

    }

    public interface IPeriodoService
    {
        IList<Periodo> Periodo { get; }
    }
    internal class PeriodoService : IPeriodoService
    {

        private List<Periodo> _periodo = null;

        public PeriodoService(string annoCorrente, string annoRiferimento, DatePicker fromDate, DatePicker toDate)
        {
            _periodo = new List<Periodo>();

            Periodo p = new Periodo();

            string dataInizioCorrente = string.Empty;
            string dataFineCorrente = string.Empty;
            string dataInizioRiferimento = string.Empty;
            string dataFineRiferimento = string.Empty;

            string da = DateTime.Parse(fromDate.ToString()).ToString("MM-dd");
            string al = DateTime.Parse(toDate.ToString()).ToString("MM-dd");

            dataInizioCorrente = $"{annoCorrente}-{da}";
            dataFineCorrente = $"{annoCorrente}-{al}";
            dataInizioRiferimento = $"{annoRiferimento}-{da}";
            dataFineRiferimento = $"{annoRiferimento}-{al}";

            p.annoCorrente = annoCorrente;
            p.annoRiferimento = annoRiferimento;
            p.dataInizioCorrente = dataInizioCorrente;
            p.dataFineCorrente = dataFineCorrente;
            p.dataInizioRiferimento = dataInizioRiferimento;
            p.dataFineRiferimento = dataFineRiferimento;

            _periodo.Add(p);
        }
        public IList<Periodo> Periodo => _periodo;  //elemento pubblico che da modo di visualizzare un elemento privato
    }

    internal class CheckService
    {
        public CheckService(RadioButton t_1, RadioButton t_2, RadioButton t_3, RadioButton t_4)
        {


            try
            {
                //apri xml trimestre
                List<string> trimestreSelezionato = new List<string>();
                XmlSerializer xmlsd = new XmlSerializer(typeof(List<string>));
                using (TextReader tr = new StreamReader("trimestreSelezionato.xml"))
                {
                    trimestreSelezionato = (List<string>)xmlsd.Deserialize(tr);
                }


                string t = trimestreSelezionato[0];

                var p = General.leggiPeriodoHome();

                switch (p.Trimestre)
                {
                    case "t_1":
                        t_1.IsChecked = true;
                        break;
                    case "t_2":
                        t_2.IsChecked = true;
                        break;
                    case "t_3":
                        t_3.IsChecked = true;
                        break;
                    case "t_4":
                        t_4.IsChecked = true;
                        break;
                }

                if (p.Trimestre != string.Empty)
                {
                    return;
                }



            }
            catch
            {

            }



            int mont = DateTime.Now.Month;

            if (mont >= 1 && mont <= 3)
            {
                t_1.IsChecked = true;

            }
            if (mont >= 4 && mont <= 6)
            {
                t_2.IsChecked = true;
            }
            if (mont >= 7 && mont <= 9)
            {
                t_3.IsChecked = true;

            }
            if (mont >= 10 && mont <= 12)
            {
                t_4.IsChecked = true;
            }


        }
    }
}
