using Aspose.Cells;
using Aspose.Cells.Properties;
using OfficeOpenXml.ConditionalFormatting;
using Workbook = Aspose.Cells.Workbook;
using System.Text.RegularExpressions;

namespace withdrawingDataExcel
{
    class Program
    {
        Workbook wb;

        //путь к файлу 
        public string path { get; set; }
        //название листа  
        public int numberSheets { get; set; }
        //название группы 
        public string nameGroup { get; set; }


        public Program(string path, int numberSheets, string nameGroup)
        {
            this.path = path;
            this.numberSheets = numberSheets;
            this.nameGroup = nameGroup;
            wb = new Workbook(path);
        }

        //метод возвращающий список листов
        public List<string> sheetsList()
        {
            WorksheetCollection collection = wb.Worksheets;
            List<string> sheets = new List<string>();
            for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
            {
                sheets.Add(collection[worksheetIndex].Name.ToString());
            }
            return sheets;
        }

        //поиск по индексу, принимает индекс формата A0, возвращает ряд и колонку
        public string findIndex(string nameIndex)
        {
            WorksheetCollection collection = wb.Worksheets;
            Worksheet worksheet = collection[numberSheets];

            int row;
            int column;

            Aspose.Cells.CellsHelper.CellNameToIndex(nameIndex, out row, out column);

            string indexValue = Aspose.Cells.CellsHelper.CellIndexToName(row, column);
            return (string)worksheet.Cells[indexValue].Value + $" row: {row}, column: {column}";
        }

        //поиск по колонке и строке, принимает массив в виде столбца и строки. Возвращает индекс формата A0
        public string findRowAndColumn(int[] rowcol)
        {

            string name = Aspose.Cells.CellsHelper.CellIndexToName(rowcol[0], rowcol[1]);

            return name;
        }

        //метод по поиску нужной строки, который возрващает массив строки и столбца 
        public int[] findStrings(string nameString)
        {
            //значение которое ищет метод 
            string value = null;
            int row = 0;
            int column = 0;

            WorksheetCollection collection = wb.Worksheets;

            // Получить рабочий лист, используя его индекс
            Worksheet worksheet = collection[numberSheets];

            // Получить количество строк и столбцов
            int rows = worksheet.Cells.MaxDataRow;
            int cols = worksheet.Cells.MaxDataColumn;

            // Цикл по строкам
            for (int i = 0; i < rows; i++)
            {

                // Перебрать каждый столбец в выбранной строке
                for (int j = 0; j < cols; j++)
                {

                    string stringForList = (string)worksheet.Cells[i, j].Value;
                    List<string> stringsList = new List<string>();

                    if (stringForList != null)
                    {
                        stringsList = stringForList.Split(' ').ToList();
                    }

                    foreach (var item in stringsList)
                    {
                        if (item.ToLower() == nameString.ToLower())
                        {
                            value = String.Join(' ', stringsList);
                            row = i;
                            column = j;

                            break;
                        }
                    }

                }
            }
            if (value != null)
            {
                return new int[] { Convert.ToInt32(row), Convert.ToInt32(column) };
            }
            else
            { 
                return new int[] { 0 };
            }
        }

        //метод принимает строковое значение и возврщает данные формата A0
        public int[] returnIndex(string findString)
        {
            return findStrings(findString);
        }

        //поиск индекса дня недели, возвращает массив из индексов начала недели 
        public int[] findIndexWeek()
        {
            string[] week = new string[7] { "ПНД", "ВТР", "СРД", "ЧТВ", "ПТН", "СБТ", "СБТ"};

            int[] indexWeek = new int[7];
            for (int i = 0; i < week.Length; i++)
            {
                int j = 0;

                string index = findRowAndColumn(findStrings(week[i]));

                index = Regex.Match(index, @"\d+").Value;

                bool isNumeric = int.TryParse(index, out j);
                indexWeek[i] = j;
            }

            return indexWeek;
        }

        //метод принимает массив индексов дней недели. Возвращает список пар 
        public List<string> returnListPair(int[] indexWeek)
        {
            int[] column = new int[0];
            //column = 
            List<string> test = new List<string>();

            string indexGroup = findRowAndColumn(findStrings(nameGroup));

            //задаёт последний индекс субботы
            indexWeek[7] = 38;

            //for (int i = 0; i < indexWeek.Length; i++)
            //{
            //    int index = 1;
            //    for (int j = indexWeek[i]; j < indexWeek[index]; j++)
            //    {
            //        Console.WriteLine(j);
            //        index++;
            //    }
            //}

            int index = 1;

            for (int i = 0; i < indexWeek.Length - 1; i++)
            {
                while (indexWeek[i] < indexWeek[index])
                {
                    Console.WriteLine();
                }
            }

            Console.WriteLine(nameGroup);
            return test;
        }
    }
}