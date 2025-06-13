using Provvigioni_Agenti.Controllers;
using Provvigioni_Agenti.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Provvigioni_Agenti.Sellout
{
    public interface ITrasferitiService
    {
        IList<Trasferito> Trasferito { get; }
    }
    internal class IntefaceClass
    {
        public static List<string> elencaFiles(string anno, string mese, string nomeAgenzia)
        {
            List<string> list = new List<string>();

            string pathTrasferito = $"../trasferiti/{anno}/{mese}/{nomeAgenzia}";

            List<string> el = General.elencoFile(pathTrasferito);

            var elencoFiles = el.FindAll(el => el.Contains(".xlsx"));

            return elencoFiles;
        }


        public static void serializzaXml(string annoCorrente, string mese, string trasferito, string trasferitoName, List<Trasferito> sellout)
        {
            string path = $"../trasferiti/{annoCorrente}/{mese}/{trasferito}/{trasferitoName}.xml";

            XmlSerializer xmls = new XmlSerializer(typeof(List<Trasferito>));

            using (TextWriter writer = new StreamWriter(path))
            {
                xmls.Serialize(writer, sellout);
            }
        }

        public static Trasferito cercaInList(List<Regione> regioni, List<Trasferito> daExcel)
        {
            Trasferito res = null;

            foreach (Regione regione in regioni)
            {
                string regioneNome = regione.Nome;

                var result = daExcel.Find(x => x.Regione == regioneNome);

                if (result != null)
                {
                    if (res == null)
                    {
                        res = new Trasferito();
                    }

                    res.Regione = result.Regione;
                    res.Venduto += result.Venduto;
                }
            }

            return res;
        }

    }

}