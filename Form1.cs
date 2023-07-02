using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.LinkLabel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Application = Microsoft.Office.Interop.Word.Application;
using DocumentFormat.OpenXml;

namespace Teorver
{
    public partial class Form1 : Form
    {
        private List<string> tasks;

        private List<string> tasks_1;
        private List<string> tasks_2;
        private List<string> tasks_3;
        private List<string> tasks_4;
        private List<string> tasks_5;
        private List<string> tasks_6;
        private List<string> tasks_7;
        private List<string> tasks_8;
        private List<string> tasks_9;
        private List<string> tasks_10;
        private List<string> tasks_11;
        private List<string> tasks_12;
        private List<string> answer_1;
        private List<string> answer_2;
        private List<string> answer_3;
        private List<string> answer_4;
        private List<string> answer_5;
        private List<string> answer_6;
        private List<string> answer_7;
        private List<string> answer_8;
        private List<string> answer_9;
        private List<string> answer_10;
        private List<string> answer_11;
        private List<string> answer_12;
        public Form1()
        {
            InitializeComponent();

            tasks_1 = new List<string>();
            tasks_1.Add("Задача 1. Медиана вариационного ряда 10, 12, 12, 14, 14, 17, x7, 20, 23, 24, 25, 25 равна 17. Тогда значение варианты x7 равно:\r\n" +
                "1) 20                          2) 18                        3) 17                     4) 21.");
            tasks_1.Add("Задача 1. Медиана вариационного ряда 8, 10, 10, 11, 12, 15, x7, 18, 21, 22, 23, 23 равна 15. Тогда значение варианты x7 равно:\r\n" +
                "1) 8                          2) 15                        3) 11                     4) 10.");
            tasks_1.Add("Задача 1. Медиана вариационного ряда 5, 7, 7, 8, 9, 11, x7, 18, 21, 22, 23, 23 равна 11. Тогда значение варианты x7 равно:\r\n" +
                "1) 8                          2) 15                        3) 21                     4) 11.");
            tasks_1.Add("Задача 1. Медиана вариационного ряда 20, 22, 22, 23, 24, 27, x7, 30, 33, 34, 35, 35 равна 27. Тогда значение варианты x7 равно:\r\n" +
                "1) 27                          2) 20                       3) 30                     4) 35.");
            tasks_1.Add("Задача 1. Медиана вариационного ряда 7, 9, 9, 10, 11, 14, x7, 17, 20, 21, 22, 22 равна 14. Тогда значение варианты x7 равно:\r\n" +
                "1) 9                           2) 17                       3) 7                     4) 14.");
            tasks_1.Add("Задача 1. Медиана вариационного ряда 5, 6, 7, 7, 9, 11, x7, 15, 16, 17, 19, 19 равна 12. Тогда значение варианты x7 равно:\r\n" +
                "1) 12                           2) 15                       3) 10                     4) 5.");
            tasks_1.Add("Задача 1. Медиана вариационного ряда 7, 9, 9, 11, 12, 14, x7, 17, 18, 20, 21, 21 равна 16. Тогда значение варианты x7 равно:\r\n" +
                "1) 17                           2) 18                       3) 11                     4) 16.");
            tasks_1.Add("Задача 1. Медиана вариационного ряда 1, 2, 4, 5, 5, 7, x7, 10, 12, 14, 15, 15 равна 9. Тогда значение варианты x7 равно:\r\n" +
                "1) 9                           2) 5                      3) 10                     4) 15.");
            tasks_1.Add("Задача 1. Медиана вариационного ряда 5, 6, 6, 7, 8, 9, x7, 12, 15, 16, 18, 18 равна 10. Тогда значение варианты x7 равно:\r\n" +
                "1) 15                           2) 5                      3) 7                     4) 10.");
            tasks_1.Add("Задача 1. Медиана вариационного ряда 16, 17, 17, 18, 19, 20, x7, 23, 25, 26, 28, 28 равна 21. Тогда значение варианты x7 равно:\r\n" +
                "1) 21                           2) 17                      3) 28                     4) 25.");

            tasks_2 = new List<string>();
            tasks_2.Add("Задача 2. Размах варьирования вариационного ряда 5, 7, 8, 9, 9, 11, 12, 15, 16, 17, x11 равен 14. Тогда значение x11 равно:\r\n" +
                "1) 19                       2) 9                            3) 5                       4) 21.");
            tasks_2.Add("Задача 2. Размах варьирования вариационного ряда 7, 8, 9, 10, 10, 14, 15, 17, 18, 19, x11 равен 15. Тогда значение x11 равно:\r\n" +
                "1) 7                      2) 10                            3) 22                       4) 15.");
            tasks_2.Add("Задача 2. Размах варьирования вариационного ряда 1, 2, 3, 4, 4, 5, 6, 7, 9, 11, x11 равен 11. Тогда значение x11 равно:\r\n" +
                "1) 5                      2) 15                            3) 11                       4) 12.");
            tasks_2.Add("Задача 2. Размах варьирования вариационного ряда 10, 11, 12, 14, 14, 15, 16, 17, 19, 21, x11 равен 11. Тогда значение x11 равно:\r\n" +
                "1) 15                      2) 21                            3) 10                       4) 25.");
            tasks_2.Add("Задача 2. Размах варьирования вариационного ряда  6, 7, 8, 10, 10, 11, 12, 14, 15, 16, x11 равен 12. Тогда значение x11 равно:\r\n" +
                "1) 10                      2) 8                            3) 15                       4) 18.");

            tasks_3 = new List<string>();
            tasks_3.Add("Задача 3. Если все варианты xi исходного вариационного ряда уменьшить в четыре раза, то выборочное среднее квадратическое отклонение σВ:\r\n" +
                "1) уменьшится в четыре раза;\r\n" +
                "2) увеличится в четыре раза;\r\n" +
                "3) не изменится;\r\n" +
                "4) уменьшится в два раза.");
            tasks_3.Add("Задача 3. Если все варианты xi исходного вариационного ряда уменьшить в пять раз, то выборочное среднее квадратическое отклонение σВ:\r\n" +
                "1) увеличится в пять раз;\r\n" +
                "2) уменьшится в пять раз;\r\n" +
                "3) не изменится;\r\n" +
                "4) уменьшится в три раза.");
            tasks_3.Add("Задача 3. Если все варианты xi исходного вариационного ряда уменьшить в семь раз, то выборочное среднее квадратическое отклонение σВ:\r\n" +
                "1) увеличится в семь раз;\r\n" +
                "2) не изменится;\r\n" +
                "3) уменьшится в семь раз;\r\n" +
                "4) уменьшится в пять раз.");
            tasks_3.Add("Задача 3. Если все варианты xi исходного вариационного ряда уменьшить в восемь раз, то выборочное среднее квадратическое отклонение σВ:\r\n" +
                "1) уменьшится в восемь раз;\r\n" +
                "2) не изменится;\r\n" +
                "3) увеличится в восемь раз;\r\n" +
                "4) уменьшится в шесть раз.");
            tasks_3.Add("Задача 3. Если все варианты xi исходного вариационного ряда уменьшить в десять раз, то выборочное среднее квадратическое отклонение σВ:\r\n" +
                "1) уменьшится в восемь раз;\r\n" +
                "2) не изменится;\r\n" +
                "3) увеличится в десять раз;\r\n" +
                "4) уменьшится в десять раз.");

            tasks_4 = new List<string>();
            tasks_4.Add("Задача 4. Если все варианты xi исходного вариационного ряда увеличить в четыре раза, то выборочное среднее  x ̅в:\r\n" +
                "1) увеличится в четыре раза;\r\n" +
                "2) увеличится в восемь раза;\r\n" +
                "3) не изменится;\r\n" +
                "4) увеличится на четыре единицы.\r\n");
            tasks_4.Add("Задача 4. Если все варианты xi исходного вариационного ряда увеличить в пять раз, то выборочное среднее  x ̅в:\r\n" +
                "1) увеличится в десять раз;\r\n" +
                "2) увеличится в пять раз;\r\n" +
                "3) не изменится;\r\n" +
                "4) увеличится на пять единиц.\r\n");
            tasks_4.Add("Задача 4. Если все варианты xi исходного вариационного ряда увеличить в семь раз, то выборочное среднее  x ̅в:\r\n" +
                "1) увеличится в четырнадцать раз;\r\n" +
                "2) увеличится на семь единиц;\r\n" +
                "3) не изменится;\r\n" +
                "4) увеличится в семь раз.\r\n");
            tasks_4.Add("Задача 4. Если все варианты xi исходного вариационного ряда увеличить в восемь раз, то выборочное среднее  x ̅в:\r\n" +
                "1) увеличится на восемь единиц;\r\n" +
                "2) увеличится в шестнадцать раз;\r\n" +
                "3) увеличится в восемь раз;\r\n" +
                "4) не изменится.\r\n");
            tasks_4.Add("Задача 4. Если все варианты xi исходного вариационного ряда увеличить в десять раз, то выборочное среднее  x ̅в:\r\n" +
                "1) увеличится на десять единиц;\r\n" +
                "2) увеличится в пять раз;\r\n" +
                "3) не изменится;\r\n" +
                "4) увеличится в десять раз.\r\n");

            tasks_5 = new List<string>();
            tasks_5.Add("Задача 5. Дан доверительный интервал (20,4; 25,6) для оценки математического ожидания нормально распределенного количественного признака при известном среднем квадратическом отклонении генеральной совокупности. Тогда при увеличении объема выборки в четыре раза этот доверительный интервал примет вид:\r\n" +
                "1) (21,7; 24,3);     \r\n" +
                "2) (17,4; 27,0);     \r\n" +
                "3) (21,45; 24,15);     \r\n" +
                "4) (10,0; 30,6).     \r\n");
            tasks_5.Add("Задача 5. Дан доверительный интервал (20,2; 25,4) для оценки математического ожидания нормально распределенного количественного признака при известном среднем квадратическом отклонении генеральной совокупности. Тогда при увеличении объема выборки в четыре раза этот доверительный интервал примет вид:\r\n" +
                "1) (21,47; 24,15);     \r\n" +
                "2) (18,4; 20,0);     \r\n" +
                "3) (21,5; 24,1);     \r\n" +
                "4) (15,0; 30,5).     \r\n");
            tasks_5.Add("Задача 5. Дан доверительный интервал (24,6; 26,8) для оценки математического ожидания нормально распределенного количественного признака при известном среднем квадратическом отклонении генеральной совокупности. Тогда при увеличении объема выборки в четыре раза этот доверительный интервал примет вид:\r\n" +
                "1) (23,5;27,9);     \r\n" +
                "2) (25,15; 26,25);     \r\n" +
                "3) (21,3; 30,1);     \r\n" +
                "4) (15,0; 30,5).     \r\n");
            tasks_5.Add("Задача 5. Дан доверительный интервал (16,4; 19,6) для оценки математического ожидания нормально распределенного количественного признака при известном среднем квадратическом отклонении генеральной совокупности. Тогда при увеличении объема выборки в четыре раза этот доверительный интервал примет вид:\r\n" +
                "1) (23,5;27,9);     \r\n" +
                "2) (15,15; 16,25);     \r\n" +
                "3) (21,3; 30,1);     \r\n" +
                "4) (17,2; 18,8).     \r\n");
            tasks_5.Add("Задача 5. Дан доверительный интервал (27,18; 32,75) для оценки математического ожидания нормально распределенного количественного признака при известном среднем квадратическом отклонении генеральной совокупности. Тогда при увеличении объема выборки в четыре раза этот доверительный интервал примет вид:\r\n" +
                "1) (23,5;27,9);     \r\n" +
                "2) (15,17; 16,25);     \r\n" +
                "3) (21,3; 30,1);     \r\n" +
                "4) (28,57; 31,36).     \r\n");

            tasks_6 = new List<string>();
            tasks_6.Add("Задача 6. Дан доверительный интервал (16,78; 19,04) для оценки математического ожидания нормально распределенного количественного признака. Тогда при увеличении объема выборки этот доверительный интервал может принять вид:\r\n" +
                "1) (17,19; 18,36);     \r\n" +
                "2) (16,15; 19,41);     \r\n" +
                "3) (17,13; 18,99);     \r\n" +
                "4) (16,15; 18,38).     \r\n");
            tasks_6.Add("Задача 6. Дан доверительный интервал (16,64; 18,92) для оценки математического ожидания нормально распределенного количественного признака. Тогда при увеличении объема выборки этот доверительный интервал может принять вид:\r\n" +
                "1) (16,15; 18,38);     \r\n" +
                "2) (17,18; 18,92);     \r\n" +
                "3) (16,15; 19,41);     \r\n" +
                "4) (17,18; 18,38).     \r\n");
            tasks_6.Add("Задача 6. Дан доверительный интервал (24,6; 26,8) для оценки математического ожидания нормально распределенного количественного признака. Тогда при увеличении объема выборки этот доверительный интервал может принять вид:\r\n" +
                "1) (24,1; 26);    \r\n" +
                "2) (25,4; 26);     \r\n" +
                "3) (24,1; 27,15);    \r\n" +
                "4) (25,4; 26,8).    \r\n");
            tasks_6.Add("Задача 6. Дан доверительный интервал (20,2; 25,4) для оценки математического ожидания нормально распределенного количественного признака. Тогда при увеличении объема выборки этот доверительный интервал может принять вид:\r\n" +
                "1) (22,3; 25,4);    \r\n" +
                "2) (20,1; 25,4);    \r\n" +
                "3) (22,3; 23,3);    \r\n" +
                "4) (20,1; 26,5).    \r\n");
            tasks_6.Add("Задача 6. Дан доверительный интервал (20,4; 25,6) для оценки математического ожидания нормально распределенного количественного признака. Тогда при увеличении объема выборки этот доверительный интервал может принять вид:\r\n" +
                "1) (20,1; 24,6);    \r\n" +
                "2) (21,4; 25,6);    \r\n" +
                "3) (20,1; 25,6);    \r\n" +
                "4) (21,4; 24,6).    \r\n");


            tasks_7 = new List<string>();
            tasks_7.Add("Задача 7. Основная гипотеза имеет вид H0: σ^2 = 4,2. Тогда конкурирующей может являться гипотеза …\r\n" +
                "1) H1: σ^2 < 4,2;     \r\n" +
                "2) H1: σ^2 ≤ 4,2;     \r\n" +
                "3) H1: σ^2 ≥ 4,2;     \r\n" +
                "4) H1: σ^2 > 4.     \r\n");
            tasks_7.Add("Задача 7. Основная гипотеза имеет вид H0: σ^2 = 5. Тогда конкурирующей может являться гипотеза …\r\n" +
               "1) H1: σ^2 < 6;     \r\n" +
               "2) H1: σ^2 ≤ 5;     \r\n" +
               "3) H1: σ^2 ≥ 5;     \r\n" +
               "4) H1: σ^2 > 5.     \r\n");
            tasks_7.Add("Задача 7. Основная гипотеза имеет вид H0: σ^2 = 7. Тогда конкурирующей может являться гипотеза …\r\n" +
              "1) H1: σ^2 < 8;     \r\n" +
              "2) H1: σ^2 > 7;     \r\n" +
              "3) H1: σ^2 ≥ 7;     \r\n" +
              "4) H1: σ^2 > 5.     \r\n");
            tasks_7.Add("Задача 7. Основная гипотеза имеет вид H0: σ^2 = 9. Тогда конкурирующей может являться гипотеза …\r\n" +
               "1) H1: σ^2 < 10;     \r\n" +
               "2) H1: σ^2 ≤ 9;     \r\n" +
               "3) H1: σ^2 < 9;     \r\n" +
               "4) H1: σ^2 > 5.     \r\n");
            tasks_7.Add("Задача 7. Основная гипотеза имеет вид H0: σ^2 = 17. Тогда конкурирующей может являться гипотеза …\r\n" +
              "1) H1: σ^2 < 19;     \r\n" +
              "2) H1: σ^2 ≤ 17;     \r\n" +
              "3) H1: σ^2 ≥ 17;     \r\n" +
              "4) H1: σ^2 > 17.     \r\n");

            tasks_8 = new List<string>();
            tasks_8.Add("Задача 8. Соотношением вида P(K < -2,18) = 0,036 можно определить:\r\n" +
                "1) левостороннюю критическую область;     \r\n" +
                "2) правостороннюю критическую область;     \r\n" +
                "3) двустороннюю критическую область;     \r\n" +
                "4) область принятия гипотезы.     \r\n");
            tasks_8.Add("Задача 8. Соотношением вида P(K > 1,49) = 0,05 можно определить:\r\n" +
                "1) левостороннюю критическую область;     \r\n" +
                "2) правостороннюю критическую область;     \r\n" +
                "3) двустороннюю критическую область;     \r\n" +
                "4) область принятия гипотезы.     \r\n");
            tasks_8.Add("Задача 8. Соотношением вида P(K < -2,78) + P(K > 2,78)= 0,01 можно определить:\r\n" +
                "1) левостороннюю критическую область;     \r\n" +
                "2) правостороннюю критическую область;     \r\n" +
                "3) двустороннюю критическую область;     \r\n" +
                "4) область принятия гипотезы.     \r\n");
            tasks_8.Add("Задача 8. Соотношением вида P(K < -2,44) = 0,05 можно определить:\r\n" +
                "1) область принятия гипотезы;     \r\n" +
                "2) правостороннюю критическую область;     \r\n" +
                "3) двустороннюю критическую область;     \r\n" +
                "4) левостороннюю критическую область.     \r\n");
            tasks_8.Add("Задача 8. Соотношением вида P(K < -2,09) = 0,025 можно определить:\r\n" +
                "1) область принятия гипотезы;    \r\n" +
                "2) правостороннюю критическую область;     \r\n" +
                "3) двустороннюю критическую область;     \r\n" +
                "4) левостороннюю критическую область.    \r\n");

            tasks_9 = new List<string>();
            tasks_9.Add("Задача 9. Выборочное уравнение прямой линии регрессии Y на X имеет вид  y = -7,0 - 2,9x . Тогда выборочный коэффициент регрессии равен:\r\n" +
                "1) -2,9                            2) 2,9                         3) 7,0                     4) -0,25\r\n");
            tasks_9.Add("Задача 9. Выборочное уравнение прямой линии регрессии Y на X имеет вид  y = -4,8 + 1,2x . Тогда выборочный коэффициент регрессии равен:\r\n" +
                "1) -1,2                           2) -4,8                        3) 4,8                     4) -0,75\r\n");
            tasks_9.Add("Задача 9. Выборочное уравнение прямой линии регрессии Y на X имеет вид  y = 3,0 - 1,5x . Тогда выборочный коэффициент регрессии равен:\r\n" +
                "1) 1,5                           2) -3,0                        3) -1,5                   4) 0,15\r\n");
            tasks_9.Add("Задача 9. Выборочное уравнение прямой линии регрессии Y на X имеет вид  y = -5,0 + 2,5x . Тогда выборочный коэффициент регрессии равен:\r\n" +
               "1) 5,5                           2) -2,5                        3) 5,0                  4) -5,0\r\n");
            tasks_9.Add("Задача 9. Выборочное уравнение прямой линии регрессии Y на X имеет вид  y = 5,8 - 1,5x . Тогда выборочный коэффициент регрессии равен:\r\n" +
               "1) 5,8                           2) -5,8                        3) 5,0                  4) 1,5\r\n");

            tasks_10 = new List<string>();
            tasks_10.Add("Задача 10. Выборочное уравнение прямой линии регрессии X на Y\r\nимеет вид x ̅y = 38.3 - 2.13y, а выборочные средние квадратические отклонения равны: σx= 5.0,  σy=1.8. Тогда выборочный коэффициент корреляции rв равен:\r\n" +
                "1) -0.7668                      2) 0.7668                     3) -3.83                     4) 3.83\r\n");
            tasks_10.Add("Задача 10. Выборочное уравнение прямой линии регрессии X на Y\r\nимеет вид x ̅y = 34.5 - 2.44y, а выборочные средние квадратические отклонения равны: σx= 6.0,  σy=1.5. Тогда выборочный коэффициент корреляции rв равен:\r\n" +
                "1) 0.61                     2) -0.61                   3) -9.76                     4) 9.76\r\n");
            tasks_10.Add("Задача 10. Выборочное уравнение прямой линии регрессии X на Y\r\nимеет вид x ̅y = 15.7 - 1.5y, а выборочные средние квадратические отклонения равны: σx= 5.0,  σy=1.7. Тогда выборочный коэффициент корреляции rв равен:\r\n" +
                "1) 0.17                     2) 0.51                   3) -0.51                     4) -0.17\r\n");
            tasks_10.Add("Задача 10. Выборочное уравнение прямой линии регрессии X на Y\r\nимеет вид x ̅y = 19.5 - 1.7y, а выборочные средние квадратические отклонения равны: σx= 8.0,  σy=1.5. Тогда выборочный коэффициент корреляции rв равен:\r\n" +
                "1) 0.39875                    2) -0.39875                   3) 0.31875                     4) -0.31875\r\n");
            tasks_10.Add("Задача 10. Выборочное уравнение прямой линии регрессии X на Y\r\nимеет вид x ̅y = 39.5 - 0.9y, а выборочные средние квадратические отклонения равны: σx= 5.0,  σy=0.8. Тогда выборочный коэффициент корреляции rв равен:\r\n" +
                "1) 0.57                     2) -0.57                  3) 0.144                     4) -0.144\r\n");

            
            tasks_11 = new List<string>();
            tasks_11.Add("Задача 11. Чему равно математическое ожидание постоянной?\r\n" +
                "1)Этой постоянной;     \r\n" +
                "2)0;     \r\n" +
                "3)Единицы;     \r\n" +
                "4)Постоянной в квадрате;     \r\n");
            //tasks_11.Add("Корреляционные зависимости принято разделять на:" +
            //    "1) постоянные и непостоянные;" +
            //    "2) ");

            answer_1 = new List<string>();
            answer_1.Add("Задача 1. 3) 17");
            answer_1.Add("Задача 1. 2) 15");
            answer_1.Add("Задача 1. 4) 11");
            answer_1.Add("Задача 1. 1) 27");
            answer_1.Add("Задача 1. 4) 14");
            answer_1.Add("Задача 1. 1) 12");
            answer_1.Add("Задача 1. 4) 16");
            answer_1.Add("Задача 1. 1) 9");
            answer_1.Add("Задача 1. 4) 10");
            answer_1.Add("Задача 1. 1) 21");

            answer_2 = new List<string>();
            answer_2.Add("Задача 2. 1) 19");
            answer_2.Add("Задача 2. 3) 22");
            answer_2.Add("Задача 2. 4) 12");
            answer_2.Add("Задача 2. 2) 21");
            answer_2.Add("Задача 2. 4) 18");

            answer_3 = new List<string>();
            answer_3.Add("Задача 3. 1) уменьшится в четыре раза");
            answer_3.Add("Задача 3. 2) уменьшится в пять раз");
            answer_3.Add("Задача 3. 3) уменьшится в семь раз");
            answer_3.Add("Задача 3. 1) уменьшится в восемь раз");
            answer_3.Add("Задача 3. 4) уменьшится в десять раз");

            answer_4 = new List<string>();
            answer_4.Add("Задача 4. 1) увеличится в четыре раза");
            answer_4.Add("Задача 4. 2) увеличится в пять раз");
            answer_4.Add("Задача 4. 4) увеличится в семь раз");
            answer_4.Add("Задача 4. 3) увеличится в восемь раз");
            answer_4.Add("Задача 4. 4) увеличится в десять раз");

            answer_5 = new List<string>();
            answer_5.Add("Задача 5. 1) (21,7; 24,3)");
            answer_5.Add("Задача 5. 3) (21,5; 24,1)");
            answer_5.Add("Задача 5. 2) (25,15; 26,25)");
            answer_5.Add("Задача 5. 4) (17,2; 18,8)");
            answer_5.Add("Задача 5. 4) (28,57; 31,36)");

            answer_6 = new List<string>();
            answer_6.Add("Задача 6. 1) (17,19; 18,36)");
            answer_6.Add("Задача 6. 4) (17,18; 18,38)");
            answer_6.Add("Задача 6. 2) (25,4; 26)");
            answer_6.Add("Задача 6. 3) (22,3; 23,3)");
            answer_6.Add("Задача 6. 4) (21,4; 24,6)");

            answer_7 = new List<string>();
            answer_7.Add("Задача 7. 1) H1: σ^2 < 4,2");
            answer_7.Add("Задача 7. 4) H1: σ^2 > 5");
            answer_7.Add("Задача 7. 2) H1: σ^2 > 7");
            answer_7.Add("Задача 7. 3) H1: σ^2 < 9");
            answer_7.Add("Задача 7. 4) H1: σ^2 > 17");

            answer_8 = new List<string>();
            answer_8.Add("Задача 8. 1) левостороннюю критическую область");
            answer_8.Add("Задача 8. 2) правостороннюю критическую область");
            answer_8.Add("Задача 8. 3) двустороннюю критическую область");
            answer_8.Add("Задача 8. 4) левостороннюю критическую область");
            answer_8.Add("Задача 8. 4) левостороннюю критическую область");

            answer_9 = new List<string>();
            answer_9.Add("Задача 9. 1) -2.9");
            answer_9.Add("Задача 9. 2) -4,8");
            answer_9.Add("Задача 9. 3) -1,5");
            answer_9.Add("Задача 9. 4) -5,0");
            answer_9.Add("Задача 9. 1) 5,8");

            answer_10 = new List<string>();
            answer_10.Add("Задача 10. 1) -0.7668");
            answer_10.Add("Задача 10. 2) -0.61");
            answer_10.Add("Задача 10. 3) -0.51");
            answer_10.Add("Задача 10. 4) -0.31875");
            answer_10.Add("Задача 10. 4) -0.144");

        
            //answer_11 = new List<string>();
            //answer_11.Add("Задача 11. 1)Этой постоянной");
        }


        private void button1_Click(object sender, EventArgs e)
        {
            
            string TestNumber = textBox2.Text;

            

            // код, который будет выполнен, если NumberTasks можно преобразовать в целое число

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents|*.docx";
            openFileDialog.Title = "Выберите файл с задачами";

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Text Files (*.docx)|*.docx";
            saveFileDialog.Title = "Выберите файл для сохранения задач";

            if (int.TryParse(TestNumber, out int numOfTest) && numOfTest >= 1)
            {
            List<string> task_save = new List<string>();
            List<string> answer = new List<string>();

           
            int kount1 = 1;
            Random random = new Random();

            while (kount1 <= numOfTest)
            {
                task_save.Clear();
                answer.Clear();
                if (checkBox1.Checked)
                {
                    int index = random.Next(tasks_1.Count);
                    string selectedTask = tasks_1[index];
                    task_save.Add(selectedTask);
                    string selectedAnswer = answer_1[index];
                    answer.Add(selectedAnswer);
                    
                }
                if (checkBox2.Checked)
                {
                    int index = random.Next(tasks_2.Count);
                    string selectedTask = tasks_2[index];
                    task_save.Add(selectedTask);
                    string selectedAnswer = answer_2[index];
                    answer.Add(selectedAnswer);
                    
                }
                    if (checkBox3.Checked)
                    {

                        int index = random.Next(tasks_3.Count);
                        string selectedTask = tasks_3[index];
                        task_save.Add(selectedTask);
                        string selectedAnswer = answer_3[index];
                        answer.Add(selectedAnswer);
                        
                    }
                    if (checkBox4.Checked)
                    {

                        int index = random.Next(tasks_4.Count);
                        string selectedTask = tasks_4[index];
                        task_save.Add(selectedTask);
                        string selectedAnswer = answer_4[index];
                        answer.Add(selectedAnswer);
                        
                    }
                    if (checkBox5.Checked)
                    {

                        int index = random.Next(tasks_5.Count);
                        string selectedTask = tasks_5[index];
                        task_save.Add(selectedTask);
                        string selectedAnswer = answer_5[index];
                        answer.Add(selectedAnswer);
                        
                    }
                    if (checkBox6.Checked)
                    {

                        int index = random.Next(tasks_6.Count);
                        string selectedTask = tasks_6[index];
                        task_save.Add(selectedTask);
                        string selectedAnswer = answer_6[index];
                        answer.Add(selectedAnswer);
                        
                    }
                    if (checkBox7.Checked)
                    {

                        int index = random.Next(tasks_7.Count);
                        string selectedTask = tasks_7[index];
                        task_save.Add(selectedTask);
                        string selectedAnswer = answer_7[index];
                        answer.Add(selectedAnswer);
                       
                    }
                    if (checkBox8.Checked)
                    {

                        int index = random.Next(tasks_8.Count);
                        string selectedTask = tasks_8[index];
                        task_save.Add(selectedTask);
                        string selectedAnswer = answer_8[index];
                        answer.Add(selectedAnswer);
                      
                    }
                    if (checkBox9.Checked)
                    {

                        int index = random.Next(tasks_9.Count);
                        string selectedTask = tasks_9[index];
                        task_save.Add(selectedTask);
                        string selectedAnswer = answer_9[index];
                        answer.Add(selectedAnswer);
                        
                    }
                    if (checkBox10.Checked)
                    {

                        int index = random.Next(tasks_10.Count);
                        string selectedTask = tasks_10[index];
                        task_save.Add(selectedTask);
                        string selectedAnswer = answer_10[index];
                        answer.Add(selectedAnswer);
                        
                    }
                    
                    
                    if (kount1 == 1)
                    {
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string saveFilePath = saveFileDialog.FileName;
                        SaveTasksToDocx(task_save, saveFilePath, TestNumber);
                        MessageBox.Show("Задачи сохранены в файл.");
                    }
                    string selectedFilePath = saveFileDialog.FileName;
                    string directoryPath = Path.GetDirectoryName(selectedFilePath);
                    string newFileName = "Ответы " + kount1 + ".docx";
                    string newFilePath = Path.Combine(directoryPath, newFileName);
                    SaveListToDocx(answer, newFilePath);
                    }
               }
                else
                {
                    string selectedFilePath = saveFileDialog.FileName;
                    string directoryPath = Path.GetDirectoryName(selectedFilePath);
                    string newFileName = "Тест " + kount1 + ".docx";
                    string newFilePath = Path.Combine(directoryPath, newFileName);
                    SaveTasksToDocx(task_save, newFilePath,TestNumber);
                    string newFileName1 = "Ответы " + kount1 + ".docx";
                    string newFilePath1 = Path.Combine(directoryPath, newFileName1);
                    SaveListToDocx(answer, newFilePath1);
                }
             
                kount1++;
                }
            }
            else
            {
                // код, который будет выполнен, если numOfTest не является целым числом или меньше 1
                MessageBox.Show("Некорректное значение числа тестов!");
            }
        }


        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click_1(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
        private static List<string> ReadTasksFromDocx(string filePath, int num)
        {
            List<string> tasks = new List<string>();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
            {
                Body body = doc.MainDocumentPart.Document.Body;

                int count = 0;
                foreach (var paragraph in body.Elements<Paragraph>())
                {
                    string text = paragraph.InnerText.Trim();

                    if (text.StartsWith("Задача"))
                    {
                        tasks.Add(text);
                        count++;

                        if (count >= num)
                        {
                            break;
                        }
                    }
                }
            }

            return tasks;
        }


        private void SaveTasksToDocx(List<string> tasks, string filePath, string num)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                //Paragraph paragraph_1 = new Paragraph(new Run(new Text("Тест №" + num)));
                //body.AppendChild(paragraph_1);

                foreach (string task in tasks)
                {
                    Paragraph paragraph = new Paragraph(new Run(new Text(task)));
                    body.AppendChild(paragraph);
                }
                doc.Save();
            }
        }

        public static void SaveListToDocx(List<string> list, string outputPath)
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    Paragraph paragraph_1 = new Paragraph(new Run(new Text("Ответы:")));
                    body.AppendChild(paragraph_1);

                    foreach (string item in list)
                    {
                        Paragraph paragraph = new Paragraph(new Run(new Text(item)));
                        body.AppendChild(paragraph);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при сохранении документа: " + ex.Message);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox12.Checked)
            {
                checkBox1.Checked = true;
                checkBox2.Checked = true;
                checkBox3.Checked = true;
                checkBox4.Checked = true;
                checkBox5.Checked = true;
                checkBox6.Checked = true;
                checkBox7.Checked = true;
                checkBox8.Checked = true;
                checkBox9.Checked = true;
                checkBox10.Checked = true;

            }
            if (!checkBox12.Checked)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
 
            }
        }
    }
}






//private void SaveTasksToFile(string filePath, List<string> tasks)
//{

//    using (StreamWriter writer = new StreamWriter(filePath))
//    {
//        foreach (string task in tasks)
//        {
//            writer.WriteLine(task);
//        }
//    }
//}




//// Нумерация задач в новом списке
//string pattern = @"\d+\|";
//for (int i = 1; i <= selectedTasks.Count; i++)
//{
//    string line = selectedTasks[i - 1];
//    string replaced = Regex.Replace(line, pattern, m => i.ToString() + ")");
//    selectedTasks[i - 1] = replaced;
//}



//    // Чтение задач из файла в список
//    string filePath = @"file.txt"; // путь к файлу с задачами
//    List<string> tasks = new List<string>(); // список для хранения задач из файла
//    using (StreamReader reader = new StreamReader(filePath))
//    {
//        string line;
//        while ((line = reader.ReadLine()) != null)
//        {
//            tasks.Add(line);
//        }
//    }

//    int startingIndex = 0; // индекс, с которого начать выбор задач
//    List<string> selectedTasks = new List<string>(); // список для хранения выбранных задач

//    for (int i = startingIndex; i < startingIndex + numOfTasks && i < tasks.Count; i++)
//    {
//        selectedTasks.Add(tasks[i]);
//    }


//    // Создаем объект класса SaveFileDialog
//    SaveFileDialog saveFileDialog1 = new SaveFileDialog();

//    // Задаем свойства для объекта класса SaveFileDialog
//    saveFileDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
//    saveFileDialog1.FilterIndex = 1;
//    saveFileDialog1.RestoreDirectory = true;

//    // Если пользователь выбрал место сохранения файла и нажал кнопку "ОК"
//    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
//    {
//        // Получаем путь и имя файла
//        string NameFile = saveFileDialog1.FileName;

//        // Сохранение выбранных задач в новый файл
//        using (StreamWriter writer = new StreamWriter(NameFile))
//        {
//            writer.WriteLine($"Тест №{TestNumber}");
//            foreach (string task in selectedTasks)
//            {
//                writer.WriteLine(task);
//            }
//        }
//    }


//using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
//{
//    Body body = doc.MainDocumentPart.Document.Body;

//    foreach (var paragraph in body.Elements<Paragraph>())
//    {
//        string text = paragraph.InnerText.Trim();
//        tasks.Add(text);
//    }



//foreach (var paragraph in body.Elements<Paragraph>())
//{
//    string text = paragraph.InnerText.Trim();

//    // Проверка на начало и окончание регулярных выражений "Задача"
//    if (Regex.IsMatch(text, @"Задача", RegexOptions.IgnoreCase))
//    {
//        startReading = true;
//    }
//    else if (Regex.IsMatch(text, @"Задача", RegexOptions.IgnoreCase) && startReading)
//    {
//        break;
//    }

//    if (startReading)
//    {
//        tasks.Add(text);
//    }
//}




//    List<string> selectedTasks = new List<string>(); // список для хранения выбранных задач
//    int startingIndex = 0; // индекс, с которого начать выбор задач
//    bool startReading = false;
//    for (int i = startingIndex; i < startingIndex + num && i < tasks.Count; i++)
//    {
//        foreach (var paragraph in body.Elements<Paragraph>())
//        {
//            string text = paragraph.InnerText.Trim();

//            // Проверка на начало и окончание регулярных выражений "Задача"
//            if (Regex.IsMatch(text, @"Задача", RegexOptions.IgnoreCase))
//            {
//                startReading = true;
//            }
//            else if (Regex.IsMatch(text, @"Задача", RegexOptions.IgnoreCase) && startReading)
//            {
//                break;
//            }

//            if (startReading)
//            {
//                tasks.Add(text);
//            }
//        }




//        selectedTasks.Add(tasks[i]);
//    }

//    return selectedTasks;
//}