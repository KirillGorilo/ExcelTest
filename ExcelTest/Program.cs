using Aspose.Cells;
using Aspose.Cells.Properties;
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

        //словарь который хрванит массивы со значениями 
        public Dictionary<int, List<string>> infoWeek = 
            new Dictionary<int, List<string>>()
        {
            {1, new List<string>() },
            {2, new List<string>() },
            {3, new List<string>() },
            {4, new List<string>() },
            {5, new List<string>() },
            {6, new List<string>() },
        };

        public Program(string path, int numberSheets, string nameGroup)
        {
            this.path = path;
            this.numberSheets = numberSheets;
            this.nameGroup = nameGroup;

            wb = new Workbook(path);


            testingCourse(numberSheets);
            testingGroup(nameGroup);
            returnListPair();
        }

        #region Обработка исключения 
        void testingGroup(string nameGroup)
        {
            if (findStrings(nameGroup.ToLower()).Length == 1)
            {
                Console.WriteLine("Нет такой группы!!!");
                Environment.Exit(1);
            }
        }

        void testingCourse(int nameSheets)
        {
            for (int i = 0; i < sheetsList().Count; i++)
            {
                if (i == nameSheets)
                {
                    break;
                }
            }
            Console.WriteLine("Не найден такой курс!!!");
            Environment.Exit(1);
        }

        #endregion

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
        public int[] findIndex(string nameIndex)
        {
            WorksheetCollection collection = wb.Worksheets;
            Worksheet worksheet = collection[numberSheets];

            int row;
            int column;

            Aspose.Cells.CellsHelper.CellNameToIndex(nameIndex, out row, out column);

            string indexValue = Aspose.Cells.CellsHelper.CellIndexToName(row, column);

            int[] rowsColumns = new int[2] { row, column };
            return rowsColumns;
        }

        //поиск по колонке и строке, принимает массив в виде столбца и строки. Возвращает индекс формата A0
        public string findRowAndColumn(int[] rowcol)
        {
            string name;
            try
            {
                name = Aspose.Cells.CellsHelper.CellIndexToName(rowcol[0], rowcol[1]);

            }
            catch (IndexOutOfRangeException)
            {

                return "A0";
            }
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
        public string returnIndex(string findString)
        {
            return findRowAndColumn(findStrings(findString));
        }

        //поиск индекса дня недели, возвращает массив из индексов начала недели 
        private int[] findIndexWeek()
        {
            string[] week = new string[7] { "ПНД", "ВТР", "СРД", "ЧТВ", "ПТН", "СБТ", "СБТ" };

            int[] indexWeek = new int[7];
            for (int i = 0; i < week.Length; i++)
            {
                if (i == 6)
                {
                    indexWeek[6] = indexWeek[5] + 5;
                    break;
                }
                int j = 0;

                string index = returnIndex(week[i]);

                index = Regex.Match(index, @"\d+").Value;

                bool isNumeric = int.TryParse(index, out j);
                indexWeek[i] = j;
            }
            return indexWeek;
        }

        //метод принимает массив индексов дней недели. Возвращает список пар 
        private void returnListPair()
        {
            // Получить количество строк и столбцов
            int cols = findStrings(nameGroup)[1];

            WorksheetCollection collection = wb.Worksheets;

            // Получить рабочий лист, используя его индекс
            Worksheet worksheet = collection[numberSheets];

            int n = 1;

            int indexPair = 1;
            for (int j = 0; j < findIndexWeek().Length - 1; j++)
            {
                for (int i = findIndexWeek()[j]; i < findIndexWeek()[n]; i++)
                {
                    string stringForList = (string)worksheet.Cells[i - 1, cols].Value;
                    string teacher = (string)worksheet.Cells[i - 1, cols + 1].Value;
                    string cabinet = (string)worksheet.Cells[i - 1, cols + 2].Value;
                    string dataTime = (string)worksheet.Cells[i - 1, findStrings("часы")[1]].Value;

                    string all = $"{stringForList} {teacher} {cabinet}".Trim(' ');
                    all = all.Insert(0, indexPair.ToString() + ". ");

                    if (dataTime != null && stringForList != null)
                    {
                        if (stringForList.Replace(" ", "") != "")
                        {
                            all = all.Insert(3, dataTime.ToString());
                        }

                    }
                    if (all.Length < 13)
                    {
                        all += "Нет пары";
                    }

                    infoWeek[n].Add(all);
                    indexPair++;
                }
                indexPair = 1;
                n++;
            }
        }
    }
}