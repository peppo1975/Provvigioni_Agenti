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


    public static class TrasferitiAgenzie
    {
        // List<string> trasferiti = new List<string>() { "amei", "barcella", "comoli", "edif", "mc_elettrici", "meb", "rexel", "sacchi", "sonepar" };

        public static string Acmei { get; set; } = "acmei";
        public static string Barcella { get; set; } = "barcella";
        public static string Comoli { get; set; } = "comoli";
        public static string Edif { get; set; } = "edif";
        public static string McElettrici { get; set; } = "mc_elettrici";
        public static string Meb { get; set; } = "meb";
        public static string Rexel { get; set; } = "rexel";
        public static string Sacchi { get; set; } = "sacchi";
        public static string Sonepar { get; set; } = "sonepar";

        public static  List<string> ToArray()
        {
            List<string> list = new List<string>(); 

            list.Add(Acmei);
            list.Add(Barcella);
            list.Add(Comoli);
            list.Add(Edif);
            list.Add(McElettrici);
            list.Add(Meb);
            list.Add(Rexel);
            list.Add(Sacchi);
            list.Add(Sonepar);

            return list;
        }

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
