using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Provvigioni_Agenti.Models
{
    public interface IAgentiService
    {
        IList<Agente> Agenti { get; }
    }
    internal class AgentiService:IAgentiService
    {
        private List<Agente> _agenti = null;

        public AgentiService()
        {


            _agenti = new List<Agente>();

            //_agenti.Add(new Agente() { Nome = "Peppe", Regione = "Puglia", NikName = "GiuP", ID = "0001" });

            XmlSerializer xmlsd = new XmlSerializer(typeof(List<Agente>));

            using (TextReader tr = new StreamReader(@"agenti.xml"))
            {
                _agenti = (List<Agente>)xmlsd.Deserialize(tr);
            }

        }

        public IList<Agente> Agenti => _agenti;  //elemento pubblico che da modo di visualizzare un elemento privato
    }
}
