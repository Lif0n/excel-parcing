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
        static void Main(string[] args)
        {
            Console.Write("Введите путь файла: ");
            string path = Console.ReadLine();
            Parsing parsing = new Parsing(path);
            parsing.ParseAllData();
            parsing.OutputAllData();
            Console.ReadKey();
        }
    }
}
