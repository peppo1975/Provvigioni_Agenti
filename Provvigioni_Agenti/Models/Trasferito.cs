using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Provvigioni_Agenti.Models
{
    public class Trasferito
    {
        public string Regione { get; set; }
        public double Venduto { get; set; } = 0;
    }

    public class Final
    {
        public string Fornitore { get; set; }
        public string Valore { get; set; }
        public double ValoreEuro { get; set; }
        //public double Valore { get; set; } = 0;
    }

    public class Citta
    {
        public string Comune { get; set; }
        public string Regione { set; get; }
    }


    public class IndiceRegione
    {
        public int Index { get; set; }
        public string Regione { get; set; }
    }

    public enum ComoliStatoLettura : byte
    {
        Init,
        RigaSuccessiva,
        AttendiDati,
        LeggiValori,

    }

    public enum AcmeiStatoLettura : byte
    {
        Init,
        LeggiNomeRegione,
        AttendiFine,

    }

    public enum EdifStatoLettura : byte
    {
        Init,
        ScorriRiga,
        LeggiRegione,
        LeggiValori,
        Exit,
    }

    public enum RexelStatoLettura : byte
    {
        Init,
        StabilisciTriemestre,
        LeggiRegione,
        LeggiValori,
        Exit,
    }



    public enum SoneparStatoLettura : byte
    {
        Init,
        LeggiValori,
        Exit,
    }



    public enum SacchiStatoLettura : byte
    {
        Init,
        LeggiValori,
        Exit,
    }

}
