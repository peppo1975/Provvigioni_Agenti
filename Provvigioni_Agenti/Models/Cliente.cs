using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Provvigioni_Agenti.Models
{
    public class Cliente
    {
    }

    public class ClienteResponse
    {
        public string IdCliente { get; set; }
        public string NomeCliente { get; set; }

        public double RepSolare { get; set; } = 0;
        public double QuadriFV { get; set; } = 0;
        public double Commercializzato { get; set; } = 0;

        public List<CategoriaStatistica> CategoriaStatisticaProgressiva { get; set; } = new List<CategoriaStatistica>();
        public List<CategoriaStatistica> CategoriaStatistica { get; set; } = new List<CategoriaStatistica>();

        //public double ProgessivoRepSolare { get; set; } = 0;
        //public double ProgessivoQuadriFV { get; set; } = 0;
        //public double ProgessivoCommercializzato { get; set; } = 0;


        public double TotaleVenduto { get; set; } = 0;
        public double ProvvigioneCorrente { get; set; } = 0;
        public double ProvvigioneRiferimento { get; set; } = 0;
        public double Percentuale { get; set; } = 0;
        public double totaleAnnoPrecedente { get; set; } = 0;
    }



    public class CategoriaStatistica
    {
        public string Categoria { get; set; }
        public double ValoreCorrente { get; set; } = 0;
        public double ValoreRiferimento { get; set; } = 0;
    }


    public class CategoriaStatisticaDettaglio
    {
        public string Categoria { get; set; }
        public string ValoreCorrente { get; set; }
        public string ValoreRiferimento { get; set; }
    }

    public class ClienteResponseDatagrid
    {
        public string IdCliente { get; set; }
        public string NomeCliente { get; set; }
        public string totaleAnnoPrecedente { get; set; } = "";
        public string TotaleVenduto { get; set; } = "";

        public string Delta { get; set; } = "";
        public string DeltaPercento { get; set; } = "";

        public string ProvvigioneRiferimento { get; set; } = "";
        public string ProvvigioneCorrente { get; set; } = "";

        public string Percentuale { get; set; } = "";

    }


    public class ClienteRiepilogoVendite
    {
        public double TotaleVendutoCorrente { get; set; }
        public double TotaleVendutoRiferimento { get; set; } = 0; // anno precedente
        public double ProgressivoRiferimento { get; set; } = 0;
        public double ProgressivoCorrente { get; set; } = 0;
    }
}
