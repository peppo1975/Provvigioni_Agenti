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
        public string AnnoCorrente {  get; set; } = string.Empty;   
        public string AnnoRiferimento {  get; set; } = string.Empty;
    }
}
