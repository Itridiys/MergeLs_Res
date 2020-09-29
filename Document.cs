using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _1111111
{
    class Document
    {
        public string Ls { get; set; }
        public string NumSchet { get; set; }
        public string Res { get; set; }
        public decimal ItogVal { get; set; }
        public decimal FirstVal { get; set; }
        public decimal EndValue { get; set; }
        public string Pes { get; set; }
        public string Source { get; set; }



        public static List<Document> ParseDoc(DataTable table)
        {
            var result2 = new List<Document>();
            var itogo = 0;
            var numSch = 0;
            var ls = 0;
            var first = 0;
            var end = 0;
            var _res = 0;
            var _pes = 0;
            var _source = 0;


            for (int i = 3; i < 5; i++)
            {
                try
                {

                    for (int j = 0; j < 38; j++)
                    {
                        var val = table.Rows[i][j].ToString().Replace(" ", "").OracleTrim().ToLower();

                        switch (val)
                        {
                            case "итоговыйрасход(квт/ч)":
                                itogo = j;
                                break;

                            case "рэс":
                                _res = j;
                                break;

                            case "заводской№":
                                numSch = j;
                                break;

                            case "лицевойсчет":
                                ls = j;
                                break;

                            case "начальноепоказание":
                                first = j;
                                break;

                            case "конечноепоказание":
                                end = j;
                                break;

                            case "пэс":
                                _pes = j;
                                break;

                            case "источник":
                                _source = j;
                                break;
                                
                        }
                    }

                }
                catch (Exception ex)
                {
                    continue; //TODO: Возможна ошибка (Невозможно найти столбец 3 - Решение: !!!!!!!!Выбран Файл НЕ РЕС!!!!!!!) 
                }
            }


            for (int i = 5; i < table.Rows.Count; i++)
            {
                var numLs = table.Rows[i][ls].ToString().Replace(" ","").OracleTrim();
                var numSchet = table.Rows[i][numSch].ToString().Replace(" ","").OracleTrim();
                var res = table.Rows[i][_res].ToString();
                var itogVal = decimal.TryParse(table.Rows[i][itogo].ToString(), out decimal b) ? b : 0;
                var firstVal = decimal.TryParse(table.Rows[i][first].ToString(), out decimal first_b) ? first_b : 0;
                var endVal = decimal.TryParse(table.Rows[i][end].ToString(), out decimal end_b) ? end_b : 0;

                var pes = table.Rows[i][_pes].ToString();

                var source = table.Rows[i][_source].ToString();

                var Document = new Document();

                Document.Ls = numLs;
                Document.NumSchet = numSchet;
                Document.ItogVal = itogVal;
                Document.FirstVal = firstVal;
                Document.EndValue = endVal;
                Document.Res = res;

                Document.Pes = pes;

                Document.Source = source;

                result2.Add(Document);
            }
            return result2;


        }
    }
}
