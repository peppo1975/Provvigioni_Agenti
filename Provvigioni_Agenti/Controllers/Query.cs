using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Provvigioni_Agenti.Controllers
{
    internal class Query
    {

        public static string clientiAgente(string agente, string annoCorrente, string annoRiferimento)
        {
            string query = string.Empty;

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



            query =
                               " SELECT MMA_M.CKY_CNT_AGENTE, " +
                               "       Agenti.CDS_CNT_RAGSOC AS NomeAgente, " +
                               "       MMA_D_PROVV.NPC_PROVV," +
                               "       MMA_D_PROVV.NMP_VALPRO_UM1," +
                               "       MMA_D.NMP_VALMOV_UM1," +
                               "       Clienti.CDS_CNT_RAGSOC," +
                               "       Clienti.CKY_CNT," +
                               "       MMA_M.ANNO," +
                               "       MMA_M.DTT_DOC," +
                               "       TBSA.CDS_CAT_STAT_ART, " +
                               "       CONCAT(Articoli.CSG_CAT_STAT_ART,'_',Articoli.NGB_CAT_STAT_ART) AS CategoriaMerce, " +
                               "       MMA_D.CSG_DOC " +
                               "FROM MMA_M" +
                               "    JOIN MMA_D" +
                               "        ON MMA_D.ANNO = MMA_M.ANNO" +
                               "           AND MMA_D.CSG_DOC = MMA_M.CSG_DOC" +
                               "           AND MMA_D.CKY_SAZ_DOC = MMA_M.CKY_SAZ_DOC" +
                               "           AND MMA_D.NGB_SR_DOC = MMA_M.NGB_SR_DOC" +
                               "           AND MMA_D.NGL_DOC = MMA_M.NGL_DOC" +
                               "           AND MMA_D.CKY_CNT_CLFR = MMA_M.CKY_CNT_CLFR" +
                               "           AND MMA_D.NPR_DOC = MMA_M.NPR_DOC" +
                               "           AND MMA_D.NGB_ANNO_DOC = MMA_M.NGB_ANNO_DOC" +
                               //"           AND MMA_D.IFL_DATIMYDB = MMA_M.IFL_DATIMYDB" +
                               "    JOIN MMA_D_PROVV" +
                               "        ON MMA_D.ID_RIGA = MMA_D_PROVV.ID_RIGA" +
                               "           AND MMA_D.AZIENDA = MMA_D_PROVV.AZIENDA" +
                               "           AND MMA_D.ANNO = MMA_D_PROVV.ANNO" +
                               "           AND MMA_D.CSG_DOC = MMA_D_PROVV.CSG_DOC" +
                               "           AND MMA_D.CKY_SAZ_DOC = MMA_D_PROVV.CKY_SAZ_DOC" +
                               "           AND MMA_D.NGB_SR_DOC = MMA_D_PROVV.NGB_SR_DOC" +
                               "           AND MMA_D.NGL_DOC = MMA_D_PROVV.NGL_DOC" +
                               "           AND MMA_D.CKY_CNT_CLFR = MMA_D_PROVV.CKY_CNT_CLFR" +
                               "           AND MMA_D.NPR_DOC = MMA_D_PROVV.NPR_DOC" +
                               "           AND MMA_D.NGB_ANNO_DOC = MMA_D_PROVV.NGB_ANNO_DOC" +
                               "           AND MMA_D.NPC_PROVV = MMA_D_PROVV.NPC_PROVV" +
                               "           AND MMA_D.NMP_VALPRO_UM1 = MMA_D_PROVV.NMP_VALPRO_UM1 " +
                               "    JOIN Clienti" +
                               "        ON MMA_M.CKY_CNT_CLFR = Clienti.CKY_CNT" +
                               "    JOIN Articoli" +
                               "        ON Articoli.CodArticolo = MMA_D.CKY_ART" +
                               "    JOIN TBSA" +
                               "        ON TBSA.CKY_CAT_STAT_ART = Articoli.CSG_CAT_STAT_ART " +
                               "           AND TBSA.NKY_CAT_STAT_ART = Articoli.NGB_CAT_STAT_ART " +
                               "     JOIN Agenti " +
                               "        ON Agenti.CKY_CNT = MMA_M.CKY_CNT_AGENTE " +
                               "WHERE (" +
                               $"          MMA_M.ANNO = '{annoCorrente}' " +
                               $"          OR MMA_M.ANNO = '{annoRiferimento}' " +
                               "      )" +
                               $"      AND MMA_M.CKY_CNT_AGENTE IN ({agente}) " +
                               $"      AND (MMA_D.CSG_DOC = 'NC' OR MMA_D.CSG_DOC = 'FT' ) " +
                               "       ORDER BY Clienti.CDS_CNT_RAGSOC ASC, MMA_M.DTT_DOC ASC";


            return query;
        }

    }
}
