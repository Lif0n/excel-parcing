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
            Stopwatch stopwatch = new Stopwatch();
            List<Course> courses = new List<Course>();
            List<Group> groups = new List<Group>();
            List<Teacher> teachers = new List<Teacher>();

            Console.Write("Введите путь файла: ");
            string path = Console.ReadLine();
            stopwatch.Start();
            object rOnly = true;
            object SaveChanges = false;
            object MissingObj = System.Reflection.Missing.Value;


            Excel.Application app = new Excel.Application();
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;

            try
            {
                workbooks = app.Workbooks;
                workbook = workbooks.Open(path, MissingObj, rOnly, MissingObj, MissingObj,
                                            MissingObj, MissingObj, MissingObj, MissingObj, MissingObj,
                                            MissingObj, MissingObj, MissingObj, MissingObj, MissingObj);

                // Получение всех страниц докуента
                sheets = workbook.Sheets;

                foreach (Excel.Worksheet worksheet in sheets)
                {
                    // Получаем диапазон используемых на странице ячеек
                    Excel.Range UsedRange = worksheet.UsedRange;

                    courses = GetCourses(worksheet, UsedRange);
                    groups = GetGroups(courses, UsedRange);
                    teachers = GetTeachers(UsedRange);

                }
                Console.WriteLine("Направления");
                foreach (Course course in courses)
                    Console.WriteLine($"{course.Id.ToString()} {course.Name} {course.Shortname}");
                Console.WriteLine("Группы");
                foreach (Group group in groups)
                    Console.WriteLine($"{group.Id.ToString()} {group.CourseId.ToString()} {group.Code}");
                Console.WriteLine("Преподаватели");
                foreach (Teacher teacher in teachers)
                    Console.WriteLine($"{teacher.Id} {teacher.Surname} {teacher.Name} {teacher.Patronymic}");
                stopwatch.Stop();
                Console.WriteLine($"Время выполнения {stopwatch.ElapsedMilliseconds.ToString()}");
            }

            catch (Exception ex) 
            {
                Console.WriteLine(ex);
            }
            Console.ReadKey();
        }
        public static List<Course> GetCourses(Worksheet worksheet, Range UsedRange)
        {
            List<Course> courses = new List<Course>();

            int IdCourse = 1;
            for (int i = 3; ; i++)
            {
                Excel.Range CellRange = UsedRange.Cells[9, i];
                string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                    (CellRange as Excel.Range).Value2.ToString();
                if (CellText != null)
                {
                    string[] group = CellText.Split(new char[] { '-' });
                    if (courses.Where(x => x.Shortname == group[0]).FirstOrDefault() == null)
                    {
                        courses.Add(new Course { Id = IdCourse, Name = group[0], Shortname = group[0] });
                        IdCourse++;
                    }
                }
                else break;
            }
            return courses;
        }
        public static List<Models.Group> GetGroups(List<Course> courses, Range UsedRange)
        {
            List<Group> groups = new List<Group>();

            int IdCourse = 1;
            int IdGroup = 1;
            for (int i = 3; ; i++)
            {
                Excel.Range CellRange = UsedRange.Cells[9, i];
                string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                    (CellRange as Excel.Range).Value2.ToString();
                if (CellText != null)
                {
                    string[] group = CellText.Split(new char[] { '-' });
                    if (courses.Where(x => x.Shortname == group[0]).FirstOrDefault() == null)
                    {
                        courses.Add(new Course { Id = IdCourse, Name = group[0], Shortname = group[0] });
                        IdCourse++;
                    }
                    Course currentCourse = courses.Where(x => x.Name == group[0]).FirstOrDefault();
                    groups.Add(new Group { Id = IdGroup, CourseId = currentCourse.Id, Code = group[1], Course = currentCourse });
                    IdGroup++;
                }
                else break;
            }
            return groups;
        }
        public static List<Teacher> GetTeachers(Range UsedRange)
        {
            List<Teacher> teachers = new List<Teacher>();
            int teacherId = 1;
            for (int x = 10; x < 159 ; x++)
            {
                for (int y = 3; y < 36 ; y++)
                {
                    Range CellRange = UsedRange.Cells[x,y];
                    string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                        (CellRange as Excel.Range).Value2.ToString();
                    if (CellText != null)
                    {
                        if (CellText.Contains(" ") && CellText.Contains("."))
                        {
                            string[] teacher = CellText.Split(new char[] { ' ', '.' }, StringSplitOptions.RemoveEmptyEntries);
                            if (teacher.Length == 3 || teacher.Length == 5)
                            {
                                if (teachers.Where(t => t.Name == teacher[1] && t.Patronymic == teacher[2] && t.Surname == teacher[0]).FirstOrDefault() == null
                                    && teacher[1].Length == 1 && teacher[2].Length == 1 && teacher[0].Length > 2)
                                {
                                    teachers.Add(new Teacher
                                    {
                                        Id = teacherId,
                                        Name = teacher[1],
                                        Surname = teacher[0],
                                        Patronymic = teacher[2],
                                    }) ;
                                    teacherId++;
                                }
                            }
                        }
                    }
                }
            }
            return teachers;
        }
        public static List<Cabinet> GetCabinets(Range UsedRange)
        {
            List<Cabinet> cabinets = new List<Cabinet>();

            return cabinets;
        }
    }
}
