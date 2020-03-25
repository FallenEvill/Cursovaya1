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
        public List<List<double>> pData = new List<List<double>>();
        double fvalue = 0.00; //значение функции f
        double min = 0.00;
        double bGL,bGH;
        double a = 0.01; //нижняя граница интегрирования
        double c = 1;
        double b = 1; //верхняя граница интегрирования 
        double h = 0.01; //шаг интегрирования
        double Integral = 0.00; //значение интеграла //число разбиений
        double g, T1, T2, T3 = 0.00;
        double gmin, T1min, T2min, T3min = 0.00;
        List<List<double>> check = new List<List<double>>();
        List<List<List<List<double>>>> summa = new List<List<List<List<double>>>>();
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
            fvalue = 0;
            double figa1 = data[1][Convert.ToInt32((z+0)*100)-1];
            double figa2 = data[1][Convert.ToInt32((z + T1) * 100)-1];
            double figa3 = data[1][Convert.ToInt32((z + T2) * 100)-1];
            double figa4 = data[1][Convert.ToInt32((z + T3) * 100)-1];
            return fvalue = Math.Pow((g - figa1 - figa2 - figa3 - figa4),2);
        }
        public void GammaCheck()
        {
            double min = 10000000;
            double max = -1000000;
            double mid = 0;
            for (T1 = 0; T1 <= b; T1 += 0.01)
            {
                for (T2 = 0; T2 <= b; T2 += 0.01)
                {
                    for (T3 = 0; T3 <= b; T3 += 0.01)
                    {
                        for (double t = a; t <= c; t += 0.01)
                        {
                            if (data[1][Convert.ToInt32((t + 0) * 100)] + data[1][Convert.ToInt32((t + T1) * 100)] + data[1][Convert.ToInt32((t + T3) * 100)] > max)
                                max = data[1][Convert.ToInt32((t + 0) * 100)] + data[1][Convert.ToInt32((t + T1) * 100)] + data[1][Convert.ToInt32((t + T3) * 100)];
                            if (data[1][Convert.ToInt32((t + 0) * 100)] + data[1][Convert.ToInt32((t + T1) * 100)] + data[1][Convert.ToInt32((t + T3) * 100)] < min)
                                min = data[1][Convert.ToInt32((t + 0) * 100)] + data[1][Convert.ToInt32((t + T1) * 100)] + data[1][Convert.ToInt32((t + T3) * 100)];
                            mid = (mid + data[1][Convert.ToInt32((t + 0) * 100)] + data[1][Convert.ToInt32((t + T1) * 100)] + data[1][Convert.ToInt32((t + T3) * 100)]) / (T1 * 100 * T2 * 100 * T3 * 100 * t * 100);
                        }
                    }
                }
            }
            bGL = min;
            bGH = max;
        }


        public void Int()
        {
            min = 0;
            gmin = 0;
            T1min = 0;
            T2min = 0;
            T3min = 0;
            for (g = bGL; g < bGH; g += 1000)
            {
                for (T1 = 0; T1 <= b; T1 += 0.01)
                {
                    int hahah = 0;
                    for (T2 = 0; T2 <= b; T2 += 0.01)
                    {
                        for (T3 = 0; T3 <= b; T3 += 0.01)
                        {
                            //check.Add(new List<double>());
                            Integral = 0.00;
                            for (double i = 1; i <= ((c - a) / h)+1; i ++)
                            {
                                Integral = Integral + f(i/100) * 0.01;
                                //Integral = Integral + h * f(a + h * (i - 0.5));
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
            for(double i = 0.01; i <= c; i+=0.01)
            {
                p.WriteLine(gmin - data[1][Convert.ToInt32(i*100-1)]- data[1][Convert.ToInt32((i+T1min) * 100 - 1)] - data[1][Convert.ToInt32((i+T2min) * 100 - 1)] - data[1][Convert.ToInt32((i+T3min) * 100 - 1)]);
            }
            p.Close();
        }
        public void NewInt()
        {
            int t1Part = 0;
            int t2Part = 0;
            int t3Part = 0;
            double T1pmin = 0, T2pmin = 0, T3pmin = 0;
            for (g = bGL; g < bGH; g += 1000)
            {
                T1 = 0;
                for (int i = 0; i < pData.Count()-1; i++)
                {
                    T2 = 0;
                    for( int j = 0; j < pData.Count() - 1; j++)
                    {
                        T3 = 0;
                        for (int k = 0; k < pData.Count() - 1; k++)
                        {
                            //check.Add(new List<double>());
                            Integral = 0.00;
                            for (double l = 1; l <= (((c - a) / h)+1); l++)
                            {
                                Integral = Integral + f(l / 100) * 0.01;
                                //Integral = Integral + h * f(a + h * (i - 0.5));
                            }
                            //check[Convert.ToInt32((T3 * 100) - 1)].Add(Integral);
                            if ((min == 0) || (Integral < min))
                            {
                                min = Integral;
                                gmin = g;
                                T1pmin = T1;
                                T2pmin = T2;
                                T3pmin = T3;
                                t1Part = i;
                                t2Part = j;
                                t3Part = k;
                            }
                            int lol = pData[k].Count();
                            int kek = pData.Count();
                            T3 += 0.01 * pData[k].Count();
                        }
                        T2 += 0.01 * pData[j].Count();
                    }
                    T1 += 0.01 * pData[i].Count();
                }
            }
            for (g = bGL; g < bGH; g += 1000)
            {
                T1 = T1pmin - pData[t1Part - 1].Count() * 0.01;
                for (int i = 0; i <= (pData[t1Part].Count() + pData[t1Part-1].Count()); i++)
                {
                    T2 = T2pmin - pData[t2Part - 1].Count() * 0.01;
                    for (int j = 0; j <= (pData[t2Part].Count() + pData[t2Part - 1].Count()); j++)
                    {
                        T3 = T3pmin - pData[t2Part - 1].Count() * 0.01;
                        for (int k = 0; k <= (pData[t3Part].Count() + pData[t3Part - 1].Count()); k++)
                        {
                            //check.Add(new List<double>());
                            Integral = 0.00;
                            for (double l = 1; l <= (((c - a) / h)+1); l++)
                            {
                                Integral = Integral + f(l / 100) * 0.01;
                                //Integral = Integral + h * f(a + h * (i - 0.5));
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
                            T3 += 0.01;
                        }
                        T2 += 0.01;
                    }
                    T1 += 0.01;
                }
            }
            StreamWriter n = new StreamWriter("test1234.txt");
            Console.WriteLine("Минимальное значение интеграла:{0} ", min);
            Console.WriteLine("При значениях гамма: {0}, тау1: {1}, тау2: {2}, тау3: {3}", gmin, T1min, T2min, T3min);
            n.WriteLine("Минимальное значение интеграла:{0} ", min);
            n.WriteLine("При значениях гамма: {0}, тау1: {1}, тау2: {2}, тау3: {3}", gmin, T1min, T2min, T3min);
            for (double i = 0.01; i <= c; i += 0.01)
            {
                n.WriteLine(gmin - data[1][Convert.ToInt32(i * 100 - 1)] - data[1][Convert.ToInt32((i + T1min) * 100 - 1)] - data[1][Convert.ToInt32((i + T2min) * 100 - 1)] - data[1][Convert.ToInt32((i + T3min) * 100 - 1)]);
            }
            n.Close();
        }
        public void Partition()
        {
            int count = 0;
            pData.Add(new List<double>());
            for(int i = 1; i <= Convert.ToInt32((c-a)/h)+1; i++)
            {
                if (count + 1 == pData.Count())
                    pData.Add(new List<double>());
                pData[count].Add(data[1][i-1]);
                if (data[1][i] >= data[1][i-1])
                {
                    for(int j = 1; ((data[1][i] >= data[1][i - 1])&&(i < Convert.ToInt32((c - a) / h)+1)); j++)
                    {
                        pData[count].Add(data[1][i]);
                        i ++;
                    }
                    count += 1;
                }
                else
                {
                    for (int j = 1; ((data[1][i] <= data[1][i - 1]) && (i < Convert.ToInt32((c - a) / h)+1)); j++)
                    {
                        pData[count].Add(data[1][i]);
                        i++;
                    }
                    count += 1;
                }
            }
        }


        public Office(int cc)
        {
            for (int i = 1; i <= cc; i++)
            {
                data.Add(new List<double>());
                Transfer(i);
            }
        }
        static void Main(string[] args)
        {
            Office e1 = new Office(2);
            e1.Partition();
            Console.WriteLine("Partition Done");
            e1.GammaCheck();
            Console.WriteLine("Gamma Limits Decided");
            e1.NewInt();
            e1.Int();
            Console.WriteLine("Cool");
        }
    }
 }