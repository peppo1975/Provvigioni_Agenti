using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Provvigioni_Agenti.Models
{
    public class Elaborazione
    {
    }

    public class Storico
    {
        public string CKY_CNT_AGENTE { get; set; }
        public string NomeAgente { get; set; }
        public string NPC_PROVV { get; set; }
        public string NMP_VALPRO_UM1 { get; set; } //provvigione
        public string NMP_VALMOV_UM1 { get; set; } //venduto
        public string CDS_CNT_RAGSOC { get; set; } // nome
        public string CKY_CNT { get; set; } // id cliente
        public string ANNO { get; set; } // anno

        public string DTT_DOC { get; set; } //data

        public string CDS_CAT_STAT_ART { get; set; }
        public string CSG_DOC { get; set; }

        public string CKY_MERC { get; set; }

        public string CDS_MERC { get; set; }

    }
    public class StoricoTotal
    {
        public string CKY_CNT_AGENTE { get; set; }//agente
        public string NPC_PROVV { get; set; }
        public string NMP_VALPRO_UM1 { get; set; } //provvigione
        public string NMP_VALMOV_UM1 { get; set; } //venduto
        public string CDS_CNT_RAGSOC { get; set; } // nome
        public string CKY_CNT { get; set; } // id cliente
        public string ANNO { get; set; } // anno

        public string DTT_DOC { get; set; } //data

        public string CDS_CAT_STAT_ART { get; set; }
        public string CategoriaMerce { get; set; }
        public string CSG_DOC { get; set; }
    }

    public class Periodo
    {
        public string dataInizioRiferimento = "1971-01-01";
        public string dataFineRiferimento = "1971-01-31";

        public string dataInizioCorrente = "1970-01-01";
        public string dataFineCorrente = "1970-01-31";

        public string annoRiferimento = "1970";
        public string annoCorrente = "1971";
    }

    public class Anno
    {

        public int NGB_ANNO_DOC { get; set; }

    }


    public class ClientiContactDiretti
    {
        public string CKY_CNT_CLFR { get; set; } = string.Empty;
    }
}
