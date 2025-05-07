using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Provvigioni_Agenti.Models;

namespace Provvigioni_Agenti.Controllers
{


    internal class ElaboraXml
    {
        private List<ClienteResponse> _clientiResponse = null;

        public static List<StoricoTotal> s(string agenteId, IList<StoricoTotal> t)
        {

            List<StoricoTotal> _storico = new List<StoricoTotal>();
           // _clientiResponse = new List<ClienteResponse>();

            foreach (var storico in t)
            {

                // var result = storico.Find(x => x.IdCliente == storico.CKY_CNT.ToString());
                var result = _storico.Find(x => x.CKY_CNT_AGENTE == storico.CKY_CNT_AGENTE.ToString());
                ClienteResponse c = new ClienteResponse();

                c.IdCliente = storico.CKY_CNT;
                c.NomeCliente = storico.CDS_CNT_RAGSOC;

              //  _clientiResponse.Add(c);
            }


            List<Storico> cc = new List<Storico>();

            XmlSerializer xmls = new XmlSerializer(typeof(List<Storico>));

            using (TextWriter writer = new StreamWriter(@"C:\Users\Peppe\Desktop\Xml.xml"))
            {
                xmls.Serialize(writer, cc);
            }


            // legge xml
            XmlSerializer xmlsd = new XmlSerializer(typeof(List<Storico>));
            using (TextReader tr = new StreamReader(@"C:\Users\Peppe\Desktop\Xml.xml"))
            {
                cc = (List<Storico>)xmlsd.Deserialize(tr);
            }


            return _storico;
        }
    }



}
