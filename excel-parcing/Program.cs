using API.Models;
using API.Database;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
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
            //string path = Console.ReadLine();
            string projectDirectory = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName;

            string path = Path.Combine(projectDirectory, @"bolshoe-raspisanie-p-2-semestr.xls");
            //создание и запуск таймера
            Stopwatch stopwatch = Stopwatch.StartNew();
            Parsing parsing = new Parsing(path, 1, new Schedule
            {
                AcademicYear = 2024,
                Semester = 2,
                ScheduleStatusId = 1
            });
            Console.WriteLine("Парсинг начался. Путь: " + path.ToString());
            List<Task> tasks = parsing.ParseAllDataAsync();
            //ожидание завершения всех задач
            Task.WaitAll(tasks.ToArray());
            parsing.ParseLessonGroupTeachers();
            parsing.ParseLessons();
            //parsing.ParseTeacherSubject();
            //вывод всей информации
            //parsing.OutputAllData();
            stopwatch.Stop();
            parsing.CloseApp();
            Console.WriteLine($"Время выполнения: {stopwatch.ElapsedMilliseconds.ToString()}");
            //ParserContext.Instance.LessonTeacher.AddRange(parsing.LessonTeachers);
            //ParserContext.Instance.GroupTeacher.AddRange(parsing.GroupTeachers);
            //ParserContext.Instance.MainLessons.AddRange(parsing.Main_Lessons);
            //ParserContext.Instance.TeacherSubject.AddRange(parsing.Teacher_Subjects);
            //ParserContext.Instance.SaveChanges();
            Console.ReadKey();
            
        }
    }
}
