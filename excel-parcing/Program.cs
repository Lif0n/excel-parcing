using excel_parcing.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace excel_parcing
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.Write("Введите путь файла: ");
            string path = Console.ReadLine();
            path = @"C:\Users\home\Documents\kopiya-kopiya-bolshoe-raspisanie-p-1-semestr.xls";
            //создание и запуск таймера
            Stopwatch stopwatch = Stopwatch.StartNew();
            Parsing parsing = new Parsing(path);
            Console.WriteLine("Парсинг начался");
            List<Task> tasks = parsing.ParseAllDataAsync();
            //ожидание завершения всех задач
            Task.WaitAll(tasks.ToArray());
            //вывод всей информации
            //parsing.ParseAllData();
            parsing.OutputAllData();
            stopwatch.Stop();
            Console.WriteLine($"Время выполнения: {stopwatch.ElapsedMilliseconds.ToString()}");
            Console.ReadKey();
            
        }
    }
}
