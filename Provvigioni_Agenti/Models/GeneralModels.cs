using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Provvigioni_Agenti.Models
{
    internal class GeneralModels
    {
    }
    public class PeriodoStart
    {
        public string Trimestre { get; set; } = string.Empty;
        public string AnnoCorrente { get; set; } = string.Empty;
        public string AnnoRiferimento { get; set; } = string.Empty;
    }


    public class PeriodoList
    {
        public string Valore { get; set; } = string.Empty;
        public string Testo { get; set; } = string.Empty;
        public override string ToString()
        {
            return Testo;
        }

    }

    public class PeriodoSelezionato
    {
        public string Trimestre { get; set; } = string.Empty;
        public string Mese { get; set; } = string.Empty;
        public string AnnoCorrente { get; set; } = string.Empty;
        public string AnnoRiferimento { get; set; } = string.Empty;

    }


    public static class Mesi
    {
        public static string Gennaio { get; set; } = "m_01";
        public static string Febbraio { get; set; } = "m_02";
        public static string Marzo { get; set; } = "m_03";
        public static string Aprile { get; set; } = "m_04";
        public static string Maggio { get; set; } = "m_05";
        public static string Giugno { get; set; } = "m_06";
        public static string Luglio { get; set; } = "m_07";
        public static string Agosto { get; set; } = "m_08";
        public static string Settembre { get; set; } = "m_09";
        public static string Ottobre { get; set; } = "m_10";
        public static string Novembre { get; set; } = "m_11";
        public static string Dicembre { get; set; } = "m_12";

        public static List<string> toArray()
        {
            List<string> list = new List<string>();

            list.Add(Gennaio);
            list.Add(Febbraio);
            list.Add(Marzo);
            list.Add(Aprile);
            list.Add(Maggio);
            list.Add(Giugno);
            list.Add(Luglio);
            list.Add(Agosto);
            list.Add(Settembre);
            list.Add(Ottobre);
            list.Add(Novembre);
            list.Add(Dicembre);
            return list;
        }
    }


    public static class Trimestri
    {
        public static string T1 { get; set; } = "t_1";
        public static string T2 { get; set; } = "t_2";
        public static string T3 { get; set; } = "t_3";
        public static string T4 { get; set; } = "t_4";

        public static List<string> toArray()
        {
            List<string> list = new List<string>();

            list.Add(T1);
            list.Add(T2);
            list.Add(T3);
            list.Add(T4);
          
            return list;
        }

       
    }

}
