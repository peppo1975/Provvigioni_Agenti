using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Provvigioni_Agenti.Models
{
    public class Agente
    {
        public string ID { get; set; }
        public string Nome { get; set; }
        public string NikName { get; set; }
        public List<Regione> Regione { get; set; } = new List<Regione>();
    }

    public class Regione
    {
        public string Nome { get; set; }


    }



}
