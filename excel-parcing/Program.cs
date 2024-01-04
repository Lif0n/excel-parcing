using excel_parcing.Models;
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

            string path = Path.Combine(projectDirectory, @"raspisanie.xls");
            //создание и запуск таймера
            Stopwatch stopwatch = Stopwatch.StartNew();
            Parsing parsing = new Parsing(path);
            Console.WriteLine("Парсинг начался. Путь: " + path.ToString());
            List<Task> tasks = parsing.ParseAllDataAsync();
            //ожидание завершения всех задач
            Task.WaitAll(tasks.ToArray());
            parsing.ParseLessons();
            List<Task> continueTasks = new List<Task>();
            continueTasks.Add(new Task(parsing.ParseTeacherLesson));
            continueTasks.Add(new Task(parsing.ParseGroupTeacher));
            continueTasks.Add(new Task(parsing.ParseTeacherSubject));
            foreach (Task task in continueTasks)
            {
                task.Start();
            }
            Task.WaitAll(continueTasks.ToArray());
            //вывод всей информации
            //parsing.OutputAllData();
            stopwatch.Stop();
            parsing.CloseApp();
            Console.WriteLine($"Время выполнения: {stopwatch.ElapsedMilliseconds.ToString()}");
            ParserContext.Instance.LessonTeachers.AddRange(parsing.LessonTeachers);
            ParserContext.Instance.GroupTeachers.AddRange(parsing.GroupTeachers);
            ParserContext.Instance.MainLessons.AddRange(parsing.Main_Lessons);
            ParserContext.Instance.TeacherSubjects.AddRange(parsing.Teacher_Subjects);
            ParserContext.Instance.SaveChanges();
            Console.ReadKey();
            
        }
    }
}
