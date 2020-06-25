using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace _1111111
{
    public static class Helper
    {
        private static OleDbConnection GetOleDbConnection(string fileName, string hdr = "no")
        {
            string con = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + fileName + "; Extended Properties =\"Excel 12.0 Macro;HDR=" + hdr + ";IMEX=1;ImportMixedTypes=Text\"";
            return new OleDbConnection(con);
        }

        /// <summary>
        /// Возвращает путь установки надстройки
        /// </summary>
        /// <returns></returns>
        public static string GetLocalPath()
        {
            return Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath);
        }

        /// <summary>
        /// Возвращает таблицу открытого документа Excel по индексу листа
        /// </summary>
        /// <param name="sheetindex"></param>
        /// <param name="hdr"></param>
        /// <returns></returns>
        public static DataTable GetTableByIndex(int sheetindex = 0, string hdr = "no", string range = "", string fields = "*", string condition = "") // загрузка по индексу страницы
        {
            DataTable sheetData = new DataTable();
            string url = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;

            int wsCount = Globals.ThisAddIn.Application.Sheets.Count;
            if (sheetindex > wsCount)
                throw new Exception($"Ошибка! Документ не содержит лист номер {sheetindex}. Колистество листов в документе {wsCount}");

            string sheetname = sheetindex == 0 ? Globals.ThisAddIn.Application.ActiveSheet.Name : Globals.ThisAddIn.Application.Sheets[sheetindex].Name;

            using (OleDbConnection conn = GetOleDbConnection(url, hdr))
            {
                conn.Open();

                fields = fields.Replace(".", "#");
                string query = "select " + fields + " from [" + sheetname + "$" + range + "]" + condition;

                OleDbDataAdapter sheetAdapter = new OleDbDataAdapter(query, conn);
                sheetAdapter.Fill(sheetData);
            }

            return sheetData;
        }

        /// <summary>
        /// Возвращает таблицу по имени листа Excel
        /// </summary>
        /// <param name="sheetname"></param>
        /// <param name="hdr"></param>
        /// <returns></returns>
        public static DataTable GetTableByName(string sheetname, string hdr = "no", string range = "", string fields = "*", string condition = "")
        {
            DataTable sheetData = new DataTable();
            string url = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;

            using (OleDbConnection conn = GetOleDbConnection(url, hdr))
            {
                conn.Open();

                fields = fields.Replace(".", "#");
                string query = "select " + fields + " from [" + sheetname + "$" + range + "]" + condition;

                OleDbDataAdapter sheetAdapter = new OleDbDataAdapter(query, conn);
                sheetAdapter.Fill(sheetData);
            }
            return sheetData;
        }

        /// <summary>
        /// Возвращает таблицу выбранного листа Excel по его индексу
        /// </summary>
        /// <param name="message"></param>
        /// <param name="sheetnumber"></param>
        /// <param name="hdr"></param>
        /// <returns></returns>
        public static DataTable GetTableFromFileByIndex(string targetFileName, int sheetnumber = 1, string hdr = "no", string range = "", string fields = "*", string condition = "")
        {
            DataTable sheetData = new DataTable();

            string url = GetPath(targetFileName);


            List<string> sheetnames = new List<string>();

            using (OleDbConnection conn = GetOleDbConnection(url, hdr))
            {
                conn.Open();
                DataTable schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow s in schema.Rows)
                {
                    if (s["TABLE_NAME"].ToString().Replace("'", "").EndsWith("$"))
                    {
                        sheetnames.Add(s["TABLE_NAME"].ToString());
                    }
                }

                fields = fields.Replace(".", "#");
                string query = "select " + fields + " from [" + sheetnames[sheetnumber - 1] + range + "]" + condition;

                OleDbDataAdapter sheetAdapter = new OleDbDataAdapter(query, conn);
                sheetAdapter.Fill(sheetData);
            }


            return sheetData;
        }

        /// <summary>
        /// Возвращает таблицу выбранного листа Excel по его индексу
        /// </summary>
        /// <param name="message"></param>
        /// <param name="sheetnumber"></param>
        /// <param name="hdr"></param>
        /// <returns></returns>
        public static DataTable GetTableFromPathByIndex(string url, int sheetnumber = 1, string hdr = "no", string range = "", string fields = "*", string condition = "")
        {
            DataTable sheetData = new DataTable();

            List<string> sheetnames = new List<string>();

            using (OleDbConnection conn = GetOleDbConnection(url, hdr))
            {
                conn.Open();
                DataTable schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow s in schema.Rows)
                {
                    if (s["TABLE_NAME"].ToString().Replace("'", "").EndsWith("$"))
                    {
                        sheetnames.Add(s["TABLE_NAME"].ToString());
                    }
                }

                fields = fields.Replace(".", "#");
                string query = "select " + fields + " from [" + sheetnames[sheetnumber - 1] + range + "]" + condition;

                OleDbDataAdapter sheetAdapter = new OleDbDataAdapter(query, conn);
                sheetAdapter.Fill(sheetData);
            }

            return sheetData;
        }

        /// <summary>
        /// Возвращает таблицу выбранного листа Excel по его имени
        /// </summary>
        /// <param name="message"></param>
        /// <param name="sheetnumber"></param>
        /// <param name="hdr"></param>
        /// <returns></returns>
        public static DataTable GetTableFromPathByName(string url, string tableName, string hdr = "no", string range = "", string fields = "*", string condition = "")
        {
            DataTable sheetData = new DataTable();

            List<string> sheetnames = new List<string>();

            using (OleDbConnection conn = GetOleDbConnection(url, hdr))
            {
                conn.Open();
                DataTable schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow s in schema.Rows)
                {
                    if (s["TABLE_NAME"].ToString().Replace("'", "").EndsWith("$"))
                    {
                        sheetnames.Add(s["TABLE_NAME"].ToString());
                    }
                }

                fields = fields.Replace(".", "#");
                string query = "select " + fields + " from [" + tableName + "$" + range + "]" + condition;

                OleDbDataAdapter sheetAdapter = new OleDbDataAdapter(query, conn);
                sheetAdapter.Fill(sheetData);
            }

            return sheetData;
        }

        /// <summary>
        /// Читалка для CSV
        /// </summary>
        /// <param name="url"></param>
        /// <param name="separator"></param>
        /// <param name="rows"></param>
        /// <returns></returns>
        public static List<object[]> GetDataFromCSV(string url, char separator = ';', params string[] rows)
        {
            List<object[]> result = new List<object[]>();

            List<string> columns = new List<string>();

            using (var reader = new StreamReader(url))
            {
                var line = reader.ReadLine();
                var values = line.Split(separator);

                foreach (var value in values)
                {
                    columns.Add(value.Replace("\"", ""));
                }
            }

            List<int> index = new List<int>();
            List<string> exceptionCols = new List<string>();

            bool allContains = true;
            foreach (var row in rows)
            {
                if (columns.Contains(row))
                {
                    index.Add(columns.IndexOf(row));
                }
                else
                {
                    allContains = false;
                    exceptionCols.Add(row);
                }
            }

            if (allContains)
            {
                using (var reader = new StreamReader(url))
                {
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(separator);

                        List<object> row = new List<object>();

                        foreach (int i in index)
                        {
                            row.Add(values[i].Replace("\"", ""));
                        }

                        object[] array = row.ToArray();
                        result.Add(array);
                    }
                }

                return result;
            }
            else
                throw new Exception($"Файл не содержит столбцы: {string.Join(", ", exceptionCols)}");
        }


        /// <summary>
        /// Возвращает полный путь выбранного файла
        /// </summary>
        /// <param name="targetFileName"></param>
        /// <returns></returns>
        public static string GetPath(string targetFileName)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();           

            openFileDialog1.Title = "Укажите путь к файлу: " + targetFileName;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog1.FileName;
            }           
            else
            {
                throw new Exception("Отменено пользователем");
            }
        }

        /// <summary>
        /// Возвращает путь выбранной папки
        /// </summary>
        /// <param name="folderFilesDescription"></param>
        /// <returns></returns>

        public static string GetFoled(string folderFilesDescription)
        {
            FolderBrowserDialog openFolder = new FolderBrowserDialog();
            openFolder.Description = folderFilesDescription;

            if (openFolder.ShowDialog() == DialogResult.OK)
            {
                return openFolder.SelectedPath;
            }
            else
            {
                throw new Exception("Отменено пользователем");
            }
        }


        /// <summary>
        /// Вставка массива на лист
        /// </summary>
        /// <param name="array"></param>
        /// <param name="sheetindex"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="borders"></param>
        public static void InsertArrayToWorksheet(object[,] array, int sheetindex = 0, int row = 1, int col = 1, bool borders = true)
        {
            try
            {
                int wsCount = Globals.ThisAddIn.Application.Sheets.Count;
                if (sheetindex > wsCount)
                    throw new Exception($"Документ не содержит лист номер {sheetindex}. Колистество листов в документе {wsCount}");

                var ws = sheetindex == 0 ? Globals.ThisAddIn.Application.ActiveSheet : Globals.ThisAddIn.Application.Sheets[sheetindex];

                var firstcell = ws.Cells[row, col];

                int rows = array.GetLength(0);
                int cols = array.GetLength(1);

                var range = ws.Range[firstcell, ws.Cells[firstcell.Row + rows - 1, firstcell.Column + cols - 1]];

                range.value = array;

                if (borders)
                    range.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            }
            catch (Exception ex)
            {
                throw new Exception("Ошибка метода InsertArrayToWorksheet(" + ex.Message + ")");
            }
        }

        /// <summary>
        /// Вставка массива на лист со смещением вниз
        /// </summary>
        /// <param name="application"></param>
        /// <param name="array"></param>
        /// <param name="sheetindex"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="borders"></param>
        public static void InsertArrayToWorksheetDown(Excel.Application application, object[,] array, int sheetindex = 0, int row = 1, int col = 1, bool borders = true)
        {
            try
            {
                int wsCount = Globals.ThisAddIn.Application.Sheets.Count;
                if (sheetindex > wsCount)
                    throw new Exception($"Документ не содержит лист номер {sheetindex}. Колистество листов в документе {wsCount}");

                var ws = sheetindex == 0 ? Globals.ThisAddIn.Application.ActiveSheet : Globals.ThisAddIn.Application.Sheets[sheetindex];


                int rows = array.GetLength(0);
                int cols = array.GetLength(1);

                //вставка пустого range на одну строку короче (чтобы без пустых строк)
                var emptyFirstCell = ws.Cells[row, col];
                Excel.Range empty = ws.Range[emptyFirstCell, ws.Cells[emptyFirstCell.Row + rows - 2, emptyFirstCell.Column + cols - 1]];
                empty.Insert(Excel.XlInsertShiftDirection.xlShiftDown);


                //вставка данных из массива
                var firstCell = ws.Cells[row, col];
                Excel.Range range = ws.Range[firstCell, ws.Cells[firstCell.Row + rows - 1, firstCell.Column + cols - 1]];
                range.Value = array;


                if (borders)
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            }
            catch (Exception ex)
            {
                throw new Exception("Ошибка метода InsertArrayToWorksheet(" + ex.Message + ")");
            }
        }

        /// <summary>
        /// Принимает порядковый номер столбца Excel и возвращает имя 
        /// </summary>
        /// <param name="n"></param>
        /// <returns></returns>
        public static string GetExcelColumnName(int n)
        {
            string result = string.Empty;

            string[] alphabet = new string[26] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

            int div = (n - 1) / alphabet.Length;
            int mod = (n - 1) % alphabet.Length;

            return result = div > 0 ? alphabet[div - 1] + alphabet[mod] : alphabet[mod];
        }


        /// <summary>
        /// Чтение файлов из папки
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        public static string[] GetAnotherFilesFromDirectory(Excel.Application excel, int count = 1)
        {
            var folder = excel.ActiveWorkbook.Path;
            var appFullName = excel.ActiveWorkbook.FullName;

            DirectoryInfo dir = new DirectoryInfo(folder);

            var files = dir.GetFiles()
                .Where(x => !x.Name.StartsWith("~$"))
                .Where(x => x.Name.ToUpper().EndsWith(".CSV") || x.Name.ToUpper().EndsWith(".XLS") || x.Name.ToUpper().EndsWith(".XLSX"))
                .Where(x => x.FullName != appFullName);


            return files.Select(x => x.FullName).ToArray();
        }

        /// <summary>
        /// Конвертер для дат типа 43643 => 27.06.2019
        /// </summary>
        /// <param name="s"></param>
        /// <param name="date"></param>
        /// <returns></returns>
        public static bool TryParseAODate(string s, out DateTime? date)
        {
            bool success = double.TryParse(s, out double value);

            if (success)
            {
                date = DateTime.FromOADate(value);
                return true;
            }
            else
            {
                date = null;
                return false;
            }
        }

        /// <summary>
        /// Поворот массива(матрицы) по/против часовой стрелки
        /// </summary>
        /// <param name="array"></param>
        /// <param name="clockwise"></param>
        /// <returns></returns>
        public static object[,] RotateArray(object[,] array, bool clockwise = true)
        {
            var result = new object[array.GetLength(1), array.GetLength(0)];

            if (clockwise)
            {
                for (int i = 0; i < result.GetLength(0); i++)
                {
                    for (int j = 0; j < result.GetLength(1); j++)
                    {
                        result[i, j] = array[result.GetLength(1) - j - 1, i];
                    }
                }
            }
            else
            {
                for (int i = 0; i < result.GetLength(0); i++)
                {
                    for (int j = 0; j < result.GetLength(1); j++)
                    {
                        result[i, j] = array[j, result.GetLength(0) - 1 - i];
                    }
                }
            }

            return result;
        }
    }


    public static class Extention
    {
        /// <summary>
        /// Убирает пустой символ из отчетов Оракл
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string OracleTrim(this string s)
        {
            return s.Replace(" ", "").Trim();
        }

        /// <summary>
        /// Удаляет лишние пробелы в строке
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string ReplaceDoubleSpaces(this string s) //хз как работает или нет
        {
            if (s.Contains("  "))
            {
                return ReplaceDoubleSpaces(s.Replace("  ", " "));
            }
            else
            {
                return s;
            }
        }

        /// <summary>
        /// Переводит коллекцию в двумерный массив для вставки в Excel
        /// </summary>
        /// <param name="list"></param>
        /// <param name="isArray"></param>
        /// <returns></returns>
        public static object[,] To2DArrayProp(this IList list, bool isArray = false)
        {
            try
            {
                if (isArray)
                {
                    var result = new object[1, list.Count];

                    for (int i = 0; i < result.GetLength(1); i++)
                    {
                        result[0, i] = list[i];
                    }

                    return result;
                }
                else
                {
                    var properties = list[0].GetType().GetProperties().Select(x => x.Name).ToList();

                    var result = new object[list.Count, properties.Count];

                    for (int i = 0; i < result.GetLength(0); i++)
                    {
                        for (int j = 0; j < result.GetLength(1); j++)
                        {
                            result[i, j] = list[i].GetType().GetProperty(properties[j]).GetValue(list[i], null);
                        }
                    }

                    return result;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Ошибка метода To2DArray(" + ex.Message + ")");
            }
        }

        /// <summary>
        /// Получает пути для файлов, содержащихся в той же директории, что и открытый документ Excel
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        public static string[] GetAnotherFilesFromDirectory(this Excel.Application excel, int count = 1)
        {
            var folder = excel.ActiveWorkbook.Path;
            var appFullName = excel.ActiveWorkbook.FullName;

            DirectoryInfo dir = new DirectoryInfo(folder);

            var files = dir.GetFiles()
                .Where(x => !x.Name.StartsWith("~$"))
                .Where(x => x.Name.ToUpper().EndsWith(".CSV") || x.Name.ToUpper().EndsWith(".XLS") || x.Name.ToUpper().EndsWith(".XLSX"))
                .Where(x => x.FullName != appFullName);


            return files.Select(x => x.FullName).ToArray();

        }


    }
}
