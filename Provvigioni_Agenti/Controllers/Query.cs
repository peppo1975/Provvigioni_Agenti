using Provvigioni_Agenti.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Provvigioni_Agenti.Controllers
{
    internal class Query
    {

        public static string clientiAgente(string agente, string annoCorrente, string annoRiferimento,List<ClientiContactDiretti> clientiContactDiretti)
        {
            string query = string.Empty;

            List<string> list = new List<string>();
            list = clientiContactDiretti.Select(x => x.CKY_CNT_CLFR).ToList();

            string direttiContact = String.Join(',',list);

            if (agente.Contains('#'))
            {
                List<string> idAll = new List<string>();
                string[] subs = agente.Split('#');

                foreach (string s in subs)
                {
                    idAll.Add($"'{s}'");
                }

                agente = string.Join(',', idAll);
            
            }
            else
            {
                agente = $"'{agente}'";
            }

            string readText = File.ReadAllText("query.sql");

            query = readText.Replace("{annoCorrente}", annoCorrente).Replace("{annoRiferimento}", annoRiferimento).Replace("{agente}", agente).Replace("{direttiContact}", direttiContact);

            return query;
        }

    }
}
