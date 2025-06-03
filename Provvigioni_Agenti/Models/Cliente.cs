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
        public string IdCliente { get; set; } = string.Empty;
        public string NomeCliente { get; set; } = string.Empty;

        public List<CategoriaStatistica> CategoriaStatisticaProgressiva { get; set; } = new List<CategoriaStatistica>();
        public List<CategoriaStatistica> CategoriaStatistica { get; set; } = new List<CategoriaStatistica>();

        public List<GruppoStatisticoVendita> GruppoStatisticoCorrente { get; set; } = new List<GruppoStatisticoVendita>();
        public List<GruppoStatisticoVendita> GruppoStatisticoRiferimento { get; set; } = new List<GruppoStatisticoVendita>();
        public List<GruppoStatisticoVendita> GruppoStatisticoCorrenteProgressivo { get; set; } = new List<GruppoStatisticoVendita>();
        public List<GruppoStatisticoVendita> GruppoStatisticoRiferimentoProgressivo { get; set; } = new List<GruppoStatisticoVendita>();

        public List<GruppoStatisticoDataGrid> GruppoStatisticoDataGridProgressivo { get; set; } = new List<GruppoStatisticoDataGrid>();
        public List<GruppoStatisticoDataGrid> GruppoStatisticoDataGridTrimestre { get; set; } = new List<GruppoStatisticoDataGrid>();

        public double TotaleVendutoCorrente { get; set; } = 0;
        public double TotaleVendutoCorrenteProgressivo { get; set; } = 0;
        public double ProvvigioneCorrente { get; set; } = 0;
        public double ProvvigioneRiferimento { get; set; } = 0;
        public double Percentuale { get; set; } = 0;
        public double TotaleVendutoRiferimento { get; set; } = 0;
        public double TotaleVendutoRiferimentoProgressivo { get; set; } = 0;
    }




    public class CategoriaStatistica
    {
        public string Categoria { get; set; } = string.Empty;
        public double ValoreCorrente { get; set; } = 0;
        public double ValoreRiferimento { get; set; } = 0;
    }


    public class CategoriaStatisticaDettaglio
    {
        public string Categoria { get; set; } = string.Empty;
        public string ValoreCorrente { get; set; } = string.Empty;
        public string ValoreRiferimento { get; set; } = string.Empty;
    }

    public class ClienteResponseDatagrid
    {
        public string IdCliente { get; set; } = string.Empty;
        public string NomeCliente { get; set; } = string.Empty;
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


    public class EstrapolaDatiCliente
    {
        public double totaleVenduto { get; set; } = 0;
        public double totaleVendutoProgressivo { get; set; } = 0;
        public double provvigione { get; set; } = 0;

        public List<GruppoStatisticoVendita> GruppoStatisticoTrimestre { get; set; } = new List<GruppoStatisticoVendita>();
        public List<GruppoStatisticoVendita> GruppoStatisticoProgressivo { get; set; } = new List<GruppoStatisticoVendita>();
    }


    public class GruppoStatistico
    {
        public string CKY_MERC { get; set; } = string.Empty;
        public string CDS_MERC { get; set; } = string.Empty;
    }

    public class GruppoStatisticoVendita 
    {
        public string CKY_MERC { get; set; } = string.Empty;
        public string CDS_MERC { get; set; } = string.Empty;

        public double Valore { get; set; } = 0;
        public string ValoreString { get; set; } = string.Empty;
    }

    public class GruppoStatisticoDataGrid  
    {
        public string CKY_MERC { get; set; } = string.Empty;
        public string CDS_MERC { get; set; } = string.Empty;


        public double ValoreRiferimento { get; set; } = 0;
        public double ValoreCorrente { get; set; } = 0;
        public string ValoreRiferimentoString { get; set; } = string.Empty;
        public string ValoreCorrenteString { get; set; } = string.Empty;

    }

    public class GruppoStatisticoRiepilogo
    {
        public string CKY_MERC { get; set; } = string.Empty;
        public string CDS_MERC { get; set; } = string.Empty;

        public double ValoreRiferimento { get; set; } = 0;
        public double ValoreCorrente { get; set; } = 0;

    }

}
