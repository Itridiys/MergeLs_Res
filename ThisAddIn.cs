using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;
using System.Collections;
using System.Data;

namespace _1111111
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public string getFileExtension(string fileName)
        {
            return fileName.Substring(fileName.LastIndexOf(".") + 1);
        }



        public void Main()
        {

            try
            {
                #region input
                
               
                var parth2 = Helper.GetPath("Укажите общий файл Сбыт");
                var arrayPath = new string[43];


                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog1.Multiselect = true;

                    openFileDialog1.Title = "Укажите путь к файлам РЭС: ";
                    if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        string[] kol = openFileDialog1.FileNames;
                        int count = kol.Length;
                        if (count <= 43)
                        {
                            int i = 0;
                            foreach (string File in openFileDialog1.FileNames.AsParallel())
                            {
                                var endex = getFileExtension(File);
                                if (endex == "xlsx" || endex =="xls")
                                {
                                    arrayPath[i] = File;
                                    i++;
                                }
                                else
                                {
                                    throw new Exception("Некорректный Формат файла" + File);
                                }
                            }
                        }
                        else
                        {
                            throw new Exception("Колличество Файлов Привысело Лимит");
                        }
                    }
                    else
                    {
                        throw new Exception("Отменено пользователем");
                    }


                    var arrayTable = new DataTable[43];

                    for (int i = 0; i < arrayPath.Length; i++)
                    {
                        if (arrayPath[i] != null)
                        {
                            arrayTable[i] = Helper.GetTableFromPathByIndex(arrayPath[i], 1);
                        }
                    }

                    var arrayDoc = new List<Document>[43];

                    for (int i = 0; i < arrayTable.Length; i++)
                    {
                        if (arrayTable[i] != null)
                        {
                          arrayDoc[i] = Document.ParseDoc(arrayTable[i]);
                        }
                    }

                    var Docum = arrayDoc[0];

                    for (int i = 1; i < arrayDoc.Length; i++)
                    {
                        if (arrayDoc[i] != null)
                        {
                            Docum.AddRange(arrayDoc[i].AsParallel());
                        }
                    }

               

                //var parth = Helper.GetPath("Укажите общий файл МРСК");
                //var parth2 = Helper.GetPath("Укажите общий файл Сбыт");


                //var MRSTable = Helper.GetTableFromPathByIndex(parth, 1);
                var ESTable = Helper.GetTableFromPathByIndex(parth2, 1);

                    //var Docum = Document.ParseDoc(MRSTable);
                    var EsDocum = EsDocument.ParseEs(ESTable);   

                    var group1 = (from o in Docum
                                  group o by o.Ls into g
                                  select new Document
                                  {
                                      Ls = g.Key,
                                      NumSchet = string.Join(" , ", g.Select(n => n.NumSchet).Distinct()),
                                      ItogVal = g.Sum(u => u.ItogVal),
                                      FirstVal = g.Sum(a => a.FirstVal),
                                      EndValue = g.Sum(b => b.EndValue),
                                      Res = string.Join(" , ", g.Select(r => r.Res).Distinct()),
                                      
                                      Pes= g.OrderBy(r=> r.Pes).First().Pes,

                                      Source= string.Join(" , ", g.Select(n => n.Source).Distinct()),

                                      /*NumDog = g.Sum(a => a.NumDog)*/
                                  }).ToList();

                    var groupEs = (from o in EsDocum
                                   group o by o.Ls into g
                                   select new EsDocument
                                   {
                                       Ls = g.Key,
                                       NumSchet = string.Join(" , ", g.Select(n => n.NumSchet).Distinct()),
                                       ItogVal = g.Sum(u => u.ItogVal),
                                       FirstVal = g.Sum(a => a.FirstVal),
                                       EndValue = g.Sum(b => b.EndValue),
                                       Postav = string.Join(" , ", g.Select(_n => _n.Postav).Distinct()),
                                       Source = string.Join(" , ", g.Select(_n => _n.Source).Distinct()),

                                   }).ToList();

                

                var leftOuterJoin =
                    (from es in groupEs
                     join mrsk in group1
                     on new { es.Ls } equals new { mrsk.Ls } into temp
                     from mrsk in temp.DefaultIfEmpty(new Document
                     {
                         Ls = "В MRSK НЕТУ Лицевого Счета",
                         NumSchet = "Нет значения",
                         FirstVal = 0,
                         EndValue = 0,
                         ItogVal = 0,
                         Res = "Нет Значения",
                         Pes = "Нет Значения",
                         Source = "Нет Значения"

                     })
                     select new
                     {
                         d = " " + es.Ls,
                         schet = " " + es.NumSchet,
                         first = es.FirstVal,
                         end = es.EndValue,
                         itog = es.ItogVal,
                         r1 = es.Postav,
                         s = es.Source,

                         d2 = " " + mrsk.Ls + " ",
                         schet2 = " " + mrsk.NumSchet + " ",
                         firstEs = mrsk.FirstVal,
                         endEs = mrsk.EndValue,
                         itogEs = mrsk.ItogVal,
                         r2 = mrsk.Res,
                         p2 = mrsk.Pes,
                         s2 =mrsk.Source

                     }).ToList();

                var RightOuterJoin =
                    (from mrsk in  group1
                     join es in groupEs
                     on new { mrsk.Ls } equals new { es.Ls } into temp
                     from es in temp.DefaultIfEmpty(new EsDocument
                     {
                         Ls = "В MRSK НЕТУ Лицевого Счета",
                         NumSchet = "Нет значения",
                         FirstVal = 0,
                         EndValue = 0,
                         ItogVal = 0,
                         Postav = "Нет Значения",                         
                         Source = "Нет Значения"
                         

                     })
                     select new
                     {
                         d = " " + es.Ls,
                         schet = " " + es.NumSchet,
                         first = es.FirstVal,
                         end = es.EndValue,
                         itog = es.ItogVal,
                         r1 = es.Postav,
                         s = es.Source,

                         d2 = " " + mrsk.Ls + " ",
                         schet2 = " " + mrsk.NumSchet + " ",
                         firstEs = mrsk.FirstVal,
                         endEs = mrsk.EndValue,
                         itogEs = mrsk.ItogVal,
                         r2 = mrsk.Res,
                         p2 = mrsk.Pes,
                         s2 = mrsk.Source

                     }).ToList();

                var FullOuter = leftOuterJoin.Union(RightOuterJoin).Distinct().ToList();

                #endregion



                #region output
                Globals.ThisAddIn.Application.Columns[1].NumberFormat = "@";
                Globals.ThisAddIn.Application.Columns[2].NumberFormat = "@";
                Globals.ThisAddIn.Application.Columns[8].NumberFormat = "@";
                Globals.ThisAddIn.Application.Columns[9].NumberFormat = "@";


                var result = group1.To2DArrayProp();
                //var resultOuter = leftOuterJoin.To2DArrayProp();
                var resultOuter = FullOuter.To2DArrayProp();
                //var rightResult = RightOuterJoin.To2DArrayProp();



                Helper.InsertArrayToWorksheet(new string[] {
                "Ls ES",
                "Номер Счетчка",
                "Начальные показания",
                "Конечные показание",
                "Итоговый Расход",
                "РЭС",
                "Показания",
                "Ls MRSK",
                "Номер Счетчика",
                "Начальные показания",
                "Конечные показание",
                "Итоговый Расход",
                "РЭС" }.To2DArrayProp(true), 0, 1, 1, false);

                Helper.InsertArrayToWorksheet(resultOuter, 0, 2, 1, false);
                



                //Globals.ThisAddIn.Application.Sheets.Add(After: Globals.ThisAddIn.Application.Sheets[Globals.ThisAddIn.Application.Sheets.Count]);
                //Globals.ThisAddIn.Application.Sheets[Index: Globals.ThisAddIn.Application.Sheets.Count].Name = "Мрск";
                //var lastSheetIndex = Globals.ThisAddIn.Application.Sheets.Count;

                //Helper.InsertArrayToWorksheet(new string[] {
                //    "Ls ES",
                //    "Номер Счетчка",
                //    "Начальные показания",
                //    "Конечные показание",
                //    "Итоговый Расход",
                //    "РЭС",
                //    "Ls MRSK",
                //    "Номер Счетчика",
                //    "Начальные показания",
                //    "Конечные показание",
                //    "Итоговый Расход",
                //    "РЭС" }.To2DArrayProp(true), 0, 1, 1, false);
                //Helper.InsertArrayToWorksheet(rightResult, lastSheetIndex, 2, 1, false);



                #endregion

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            MessageBox.Show("Готово", "Сообщение",
                   MessageBoxButtons.OK,
                     MessageBoxIcon.Exclamation,
                     MessageBoxDefaultButton.Button1,
                     MessageBoxOptions.DefaultDesktopOnly);
        }


        public int? GetSheetIndex(string sheetName)
        {
            string iNumber = System.Convert.ToString(sheetName);
            for (int i = 1; i <= Globals.ThisAddIn.Application.Sheets.Count; i++)
            {
                if (Globals.ThisAddIn.Application.Sheets[i].Name.ToLower() == sheetName)
                {
                    return i;
                }
            }
            throw new Exception(sheetName + " Такого названия листа нет в книге");
        }

        static void OpenFile(string parth)
        {
            string path = System.IO.Path.GetFullPath(parth);
            string mySheet = path;
            var ExcelApp = new Excel.Application();
            Excel.Worksheet wss = ExcelApp.ActiveWorkbook.Worksheets[1];
            ExcelApp.WindowState = Excel.XlWindowState.xlMaximized;
            Microsoft.Office.Interop.Excel.Workbook book = ExcelApp.Workbooks.Open(path);
            ExcelApp.Visible = true;
        }



        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
