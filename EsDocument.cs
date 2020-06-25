using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _1111111
{
    class EsDocument
    {
        public string Ls { get; set; }
        public decimal ItogVal { get; set; }
        public decimal FirstVal { get; set; }
        public decimal EndValue { get; set; }
        public string Postav { get; set; }
        public string NumSchet { get; set; }

        public string Source { get; set; }


        public static List<EsDocument> ParseEs(DataTable table)
        {
            var result2 = new List<EsDocument>();

            for (int i = 1; i < table.Rows.Count; i++)
            {

                var numLs = table.Rows[i][0].ToString().Replace(" ", "").OracleTrim();
                var numSchet = table.Rows[i][20].ToString().Replace( " ", "").OracleTrim();
                var itogVal = decimal.TryParse(table.Rows[i][33].ToString(), out decimal b) ? b : 0;
                var firstVal = decimal.TryParse(table.Rows[i][23].ToString(), out decimal first_b) ? first_b : 0;
                var endVal = decimal.TryParse(table.Rows[i][26].ToString(), out decimal end_b) ? end_b : 0;
                var dictrict = table.Rows[i][5].ToString();
                var postav = table.Rows[i][6].ToString() + " " + dictrict ;

                var source = table.Rows[i][16].ToString();

                var Document = new EsDocument();

                Document.Ls = numLs;
                Document.NumSchet = numSchet;
                Document.ItogVal = itogVal;
                Document.FirstVal = firstVal;
                Document.EndValue = endVal;
                Document.Postav = postav;

                Document.Source = source;


                result2.Add(Document);
            }
            return result2;

        }
    }
}
