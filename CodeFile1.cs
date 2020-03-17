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
    class Office
    {
        public List<List<double>> data = new List<List<double>>();
        double fvalue = 0.00; //значение функции f
        double min = 0.00;
        double bGamma = 100000;
        double a = 0.01; //нижняя граница интегрирования
        double b = 5; //верхняя граница интегрирования 
        double h = 0.01; //шаг интегрирования
        double Integral = 0.00; //значение интеграла //число разбиений
        double g, T1, T2, T3 = 0.00;
        double gmin, T1min, T2min, T3min = 0.00;
        List<List<double>> check = new List<List<double>>();
        [STAThread]
        public void Transfer(int cn)
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
                data[cn-1].Add(Convert.ToDouble(strArray[i]));
            Console.WriteLine("Well");
        }
        public double f(double z)
        {
            double figa1 = data[1][Convert.ToInt32((z+0)*100)];
            double figa2 = data[1][Convert.ToInt32((z + T1) * 100)];
            double figa3 = data[1][Convert.ToInt32((z + T2) * 100)];
            double figa4 = data[1][Convert.ToInt32((z + T3) * 100)];
            return fvalue = (g - figa1 - figa2 - figa3 - figa4) * (g - figa1 - figa2 - figa3 - figa4);
        }


        public void Int()
        {
            double n = (b - a) / h;
            for (g = -100000; g < bGamma; g += 1000)
            {
                for (T1 = 0.01; T1 < b; T1 += 0.01)
                {
                    for (T2 = 0.01; T2 < b; T2 += 0.01)
                    {
                        for (T3 = 0.01; T3 < b; T3 += 0.01)
                        {
                            //check.Add(new List<double>());
                            Integral = 0.00;
                            for (double i = 0.01; i <= 0.1; i += 0.01)
                            {
                                Integral = Integral + f(i) * 0.01;
                                /*Integral = Integral + h * f(a + h * (i - 0.5));*/
                            }
                            //check[Convert.ToInt32((T3 * 100) - 1)].Add(Integral);
                            if ((min == 0) || (Integral < min))
                            {
                                min = Integral;
                                gmin = g;
                                T1min = T1;
                                T2min = T2;
                                T3min = T3;
                            }
                        }
                    }
                }
            }
            StreamWriter p = new StreamWriter("test123.txt");
            Console.WriteLine("Минимальное значение интеграла:{0} ", min);
            Console.WriteLine("При значениях гамма: {0}, тау1: {1}, тау2: {2}, тау3: {3}", gmin, T1min, T2min, T3min);
            p.WriteLine("Минимальное значение интеграла:{0} ", min);
            p.WriteLine("При значениях гамма: {0}, тау1: {1}, тау2: {2}, тау3: {3}", gmin, T1min, T2min, T3min);
            p.Close();
        }

        public Office(int cc)
        {
            for (int i = 1; i <= cc; i++)
            {
                data.Add(new List<double>());
                Transfer(i);
            }
            Int();
        }
        static void Main(string[] args)
        {
            Office e1 = new Office(2);
            Console.WriteLine("Cool");
        }
    }
 }