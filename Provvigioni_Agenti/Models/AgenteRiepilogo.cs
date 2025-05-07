using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Provvigioni_Agenti.Models
{
    public class AgenteRiepilogo
    {
        public string ID { get; set; }
        public string Nome { get; set; }


        public double VendutoRiferimento { get; set; } = 0;
        public double VendutoCorrente { get; set; } = 0;


        public double ProvvigioneCorrente { get; set; } = 0;
      

        public double Delta { get; set; } = 0;

        public double DeltaPercent { get; set; } = 0;


        public string VendutoRiferimentoString { get; set; } = "";
        public string VendutoCorrenteString { get; set; } = "";

        public string DeltaString { get; set; } = "";
        public string DeltaPercentString { get; set; } = "";

        public string ProvvigioneCorrenteString { get; set; } = "";

        public double VendutoSellout { get; set; } = 0;
        public string VendutoSelloutString { get; set; } = "";
        public double ProvvigioneSellout { get; set; } = 0;
        public string ProvvigioneSelloutString { get; set; } = "";





    }

}
