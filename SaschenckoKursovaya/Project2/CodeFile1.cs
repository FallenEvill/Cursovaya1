using System;
using System.Windows;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Movement
{
    class Optimization
    {
        public List<List<double>> data = new List<List<double>>(); // Значения фукнции в каждый момент времени до 29.5 сек
        public List<List<double>> pData = new List<List<double>>(); // Значения функции, разбитые на промежутки
        double gLim; // Верхняя граница перебора по гамме
        readonly double a = 5; //нижняя граница интегрирования
        readonly double c = 15; //верхняя граница интегрирования 
        double b; // Предел по тау
        readonly double h = 0.01; //шаг интегрирования
        double g, tau1, tau2, tau3 = 0.00; // перебираемые значения
        [STAThread]
        public void Transfer(int cn) // Перенос данных из Excel таблицы в массив значений
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.   
            string pathToFile = @"C:\Users\sashe\Pictures\SK_Moschnost.xlsx";
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(pathToFile, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            // Указываем номер столбца (таблицы Excel) из которого будут считываться данные.
            int numCol = cn;

            Microsoft.Office.Interop.Excel.Range usedColumn = ObjWorkSheet.UsedRange.Columns[numCol];
            System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
            string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

            // Выходим из программы Excel.
            ObjExcel.Quit();
            for (int i = 1; i < strArray.Length; i++)
                data[cn - 1].Add(Convert.ToDouble(strArray[i]));
            Console.WriteLine("Well");
        }

        public double F(double z) // расчет значения
        {
            double k1 = data[1][Convert.ToInt32((z + 0) * 100) - 1];
            double k2 = data[1][Convert.ToInt32((z + tau1) * 100) - 1];
            double k3 = data[1][Convert.ToInt32((z + tau2) * 100) - 1];
            double k4 = data[1][Convert.ToInt32((z + tau3) * 100) - 1];
            return (g - k1 - k2 - k3 - k4);
        }

        public void Minmax() // функция оптимизации
        {
            double max, min = 0; // Максимум по t минимум по гамма, тау1,2,3
            double t = 0;
            double tau1opt = 0, tau2opt = 0, tau3opt = 0, gmax = 0; // Оптимизированные значения тау и гамма
            for (double i = a; i < c; i += 0.01) // Цикл поиска максимального значения по t
            {
                max = Math.Abs(F(i));
                if (max > min)
                {
                    min = max;
                    t = i;
                }
            }
            b = 29.5 - c; // Определение верхнего предела тау
            Partition(); // Разделение значений на промежутки для ускорения перебора
            //do
            for (g = 0; g < gLim; g += 1000) // Поиск предварительного минимального значения по гамма и тау1,2,3
            {
                tau1 = 0;
                for (int i = 0; i < pData.Count() - 1; i++) // Цикл поиска минимального значения по первым значениям промежутков по тау1
                {
                    tau2 = 0;
                    for (int j = 0; j < pData.Count() - 1; j++)
                    {
                        tau3 = 0;
                        for (int k = 0; k < pData.Count() - 1; k++)
                        {
                            max = Math.Abs(F(t)); // Вычисление значения оптимизируемой функции
                            if (max.CompareTo(min) < 0) // Проверка на минимальность
                            {
                                min = max; // Присвоение оптимизированных значений
                                gmax = g;
                                tau1opt = tau1;
                                tau2opt = tau2;
                                tau3opt = tau3;
                            }
                            tau3 += 0.01 * pData[k].Count(); // Переход к первому значению в следующем промежутке
                        }
                        tau2 += 0.01 * pData[j].Count();
                    }
                    tau1 += 0.01 * pData[i].Count();
                }
            }
            double r, x, y, z;
            for (g = gmax, r = 1; (g >= 0) && (g <= gmax + 2500); Math.Abs(g += (Math.Pow(-1, r) * r * 500)), r += 1) // Углублённый поиск оптимальных значений гаммы и тау
            { // Math.Abs используется ради случая когда одно или несколько преварительно оптимизированных значений равно 0
                for (tau1 = tau1opt, x = 1; (tau1 >= 0) && (tau1 <= b) && (x < 251); Math.Abs(tau1 += (Math.Pow(-1, x) * x / 100)), x += 1)
                { // Ограничение итераций цикла применяется из-за необходимости сохранять производительность
                    for (tau2 = tau2opt, y = 1; (tau2 >= 0) && (tau2 <= b) && (y < 251); Math.Abs(tau2 += (Math.Pow(-1, y) * y / 100)), y += 1)
                    { // Перебор одновременно идёт по значениям большим и меньшим предварительно найденных
                        if (tau3opt != 0)
                            for (tau3 = tau3opt, z = 1; (tau3 >= 0) && (tau3 <= b) && (z < 251); Math.Abs(tau3 += (Math.Pow(-1, z) * z / 100)), z += 1)
                            {
                                max = Math.Abs(F(t));
                                if (max.CompareTo(min) < 0) 
                                {
                                    min = max; 
                                    gmax = g;
                                    tau1opt = tau1;
                                    tau2opt = tau2;
                                    tau3opt = tau3;
                                }
                            }
                    }
                }
            }
            StreamWriter n = new StreamWriter("results.txt");
            Console.WriteLine("Минимальное значение:{0} ", min); // Вывод результатов на экран
            Console.WriteLine("При значениях гамма: {0}, тау1: {1}, тау2: {2}, тау3: {3}, t:  {4}", gmax, tau1opt, tau2opt, tau3opt, t);
            n.WriteLine("Минимальное значение:{0} ", min); // Экспорт результатов в текстовый файл
            n.WriteLine("При значениях гамма: {0}, тау1: {1}, тау2: {2}, тау3: {3}, t:  {4}", gmax, tau1opt, tau2opt, tau3opt, t);
            for (double i = a; i <= c; i += 0.01) // Вывод в текстовый файл значений оптимизированной функции в каждый момент времени от a до c
            {
                n.WriteLine(data[1][Convert.ToInt32(i * 100 - 1)] + data[1][Convert.ToInt32((i + tau1opt) * 100 - 1)] + data[1][Convert.ToInt32((i + tau2opt) * 100 - 1)] + data[1][Convert.ToInt32((i + tau3opt) * 100 - 1)]);
            }
            n.Close();
        }

        public void Int() // функция оптимизации
        {
            b = 29.5 - c; // Определение верхнего предела тау
            Partition(); // Разделение значений на промежутки для ускорения перебора
            int st = Convert.ToInt32(a * 100);
            double tau1min = 0, tau2min = 0, tau3min = 0, gmin = 0, Integral; // Оптимизированные значения гамма тау
            double min = 0; // Минимальное значение интеграла
            for (g = 0; g < gLim; g += 1000) // Поиск предварительного минимального значения по гамма и тау1,2,3
            {
                tau1 = 0;
                for (int i = 0; i < pData.Count() - 1; i++) // Цикл поиска минимального значения по первым значениям промежутков по тау1
                {
                    tau2 = 0;
                    for (int j = 0; j < pData.Count() - 1; j++)
                    {
                        tau3 = 0;
                        for (int k = 0; k < pData.Count() - 1; k++)
                        {
                            Integral = 0.00;
                            for (double l = st; l <= (c / h); l++) // Вычисление значения оптимизируемой функции
                            {
                                Integral += Math.Pow(F(l / 100), 2) * 0.01;
                            }
                            if ((min == 0) || (Integral < min)) // Проверка на минимальность
                            {
                                min = Integral; // Присвоение оптимизированных значений
                                gmin = g;
                                tau1min = tau1;
                                tau2min = tau2;
                                tau3min = tau3;
                            }
                            tau3 += 0.01 * pData[k].Count(); // Переход к первому значению в следующем промежутке
                        }
                        tau2 += 0.01 * pData[j].Count();
                    }
                    tau1 += 0.01 * pData[i].Count();
                }
                if (g > gmin)
                    break;
            }
            double r, x, y, z;
                for (g = gmin, r = 1; (g >= 0) && (g <= gmin + 2500); Math.Abs(g += (Math.Pow(-1, r) * r * 500)), r += 1) // Углублённый поиск оптимальных значений гаммы и тау
                { // Math.Abs используется ради случая когда одно или несколько преварительно оптимизированных значений равно 0
                    for (tau1 = tau1min, x = 1; (tau1 >= 0) && (tau1 <= b) && (x < 101); Math.Abs(tau1 += (Math.Pow(-1, x) * x / 100)), x += 1)
                    { // Ограничение итераций цикла применяется из-за необходимости сохранять производительность
                        for (tau2 = tau2min, y = 1; (tau2 >= 0) && (tau2 <= b) && (y < 101); Math.Abs(tau2 += (Math.Pow(-1, y) * y / 100)), y += 1)
                        { // Перебор одновременно идёт по значениям большим и меньшим предварительно найденных
                            for (tau3 = tau3min, z = 1; (tau3 >= 0) && (tau3 <= b) && (z < 101); Math.Abs(tau3 += (Math.Pow(-1, z) * z / 100)), z += 1)
                            { 
                                Integral = 0;
                                for (double l = st; l <= (((c - a) / h) + st); l++)
                                {
                                    Integral += Math.Pow(F(l / 100), 2) * 0.01;
                                }
                                if (Integral < min)
                                {
                                    min = Integral;
                                    gmin = g;
                                    tau1min = tau1;
                                    tau2min = tau2;
                                    tau3min = tau3;
                                }
                            }
                        }
                    }
                }
            StreamWriter n = new StreamWriter("testau1234.txt");
            Console.WriteLine("Минимальное значение интеграла:{0} ", min); // Вывод результатов на экран
            Console.WriteLine("При значениях гамма: {0}, тау1: {1}, тау2: {2}, тау3: {3}", gmin, tau1min, tau2min, tau3min);
            n.WriteLine("Минимальное значение интеграла:{0} ", min); // Экспорт результатов в текстовый файл
            n.WriteLine("При значениях гамма: {0}, тау1: {1}, тау2: {2}, тау3: {3}", gmin, tau1min, tau2min, tau3min); 
            for (double i = a; i <= c; i += 0.01) // Вывод в текстовый файл значений оптимизированной функции в каждый момент времени от a до c
            {
                n.WriteLine(data[1][Convert.ToInt32(i * 100 - 1)] + data[1][Convert.ToInt32((i + tau1min) * 100 - 1)] + data[1][Convert.ToInt32((i + tau2min) * 100 - 1)] + data[1][Convert.ToInt32((i + tau3min) * 100 - 1)]);
            }
            n.Close();
        }
        public void Partition() // Функция разделения значений на премежутки
        {
            int count = 0; //Счётчик промежутков
            pData.Add(new List<double>());
            for (int i = Convert.ToInt32(a*100); i <= Convert.ToInt32(b / h); i++) //Цикл разбиения на промежутки
            {
                if (count + 1 == pData.Count()) //Условие увеличения List
                    pData.Add(new List<double>());
                pData[count].Add(data[1][i - 1]);
                if (data[1][i] >= data[1][i - 1]) //Если функции на промежутке возрастает
                {
                    for (int j = 1; ((data[1][i] >= data[1][i - 1]) && (i < Convert.ToInt32((b / h))) && (j <= 49)); j++)
                    {
                        pData[count].Add(data[1][i]);
                        i++;
                    }
                    count += 1;
                }
                else
                {
                    for (int j = 1; ((data[1][i] <= data[1][i - 1]) && (i < Convert.ToInt32((b / h))) && (j <= 49)); j++)
                    {
                        pData[count].Add(data[1][i]);
                        i++;
                    }
                    count += 1;
                }
            }
        }
        public Optimization(int cc) //Создание объекта класса
        {
            for (int i = 1; i <= cc; i++) //Импорт данных из Excel таблицы
            {
                data.Add(new List<double>());
                Transfer(i);
            }
        }
        static void Main()
        {
            Optimization e1 = new Optimization(2);
            Optimization e2 = new Optimization(2);
            e1.gLim = 70000;
            e1.Int();
            e2.gLim = 40000;
            e2.Minmax();
        }
    }
}
