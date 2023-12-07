using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using excel_parcing.Models;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace excel_parcing
{
    public class Parsing
    {
        public string Path { get; set; }
        public List<Cabinet> Cabinets = new List<Cabinet>();
        public List<Course> Courses = new List<Course>();
        public List<Models.Group> Groups = new List<Models.Group>();
        public List<Subject> Subjects = new List<Subject>();
        public List<Teacher> Teachers = new List<Teacher>();
        public List<Teacher_Subject> Teacher_Subjects = new List<Teacher_Subject>();
        Excel.Range UsedRange = null;
        public Parsing(string path)
        {
            Path = path;
        }

        public void ParseAllData()
        {
            object rOnly = true;
            object SaveChanges = false;
            object MissingObj = System.Reflection.Missing.Value;

            Excel.Application app = new Excel.Application();
            Excel.Workbooks workbooks = app.Workbooks;
            Excel.Workbook workbook = workbooks.Open(Path, MissingObj, rOnly, MissingObj, MissingObj,
                                        MissingObj, MissingObj, MissingObj, MissingObj, MissingObj,
                                        MissingObj, MissingObj, MissingObj, MissingObj, MissingObj);

            // Получение всех страниц докуента
            Excel.Sheets sheets = workbook.Sheets;
            foreach (Excel.Worksheet worksheet in sheets)
            {
                UsedRange = worksheet.UsedRange;

                ParseCoursesAndGroups();
                ParseTeachers();
                ParseCabinets();
                
            }
        }
        //парсинг всей информации
        public List<Task> ParseAllDataAsync()
        {
            object rOnly = true;
            object SaveChanges = false;
            object MissingObj = System.Reflection.Missing.Value;

            Excel.Application app = new Excel.Application();
            Excel.Workbooks workbooks = app.Workbooks;
            Excel.Workbook workbook = workbooks.Open(Path, MissingObj, rOnly, MissingObj, MissingObj,
                                        MissingObj, MissingObj, MissingObj, MissingObj, MissingObj,
                                        MissingObj, MissingObj, MissingObj, MissingObj, MissingObj);

            // Получение всех страниц докуента
            Excel.Sheets sheets = workbook.Sheets;
            foreach (Excel.Worksheet worksheet in sheets)
            {
                UsedRange = worksheet.UsedRange;

                List<Task> tasks= new List<Task>()
                {
                    new Task(ParseCabinets),
                    new Task(ParseCoursesAndGroups),
                    new Task(ParseTeachers),
                    new Task(ParseCabinets)
                };
                foreach (Task task in tasks)
                {
                    task.Start();
                }
                return tasks;
            }
            return null;
        }
        //вывод всех данных
        public void OutputAllData()
        {
            CoursesPasport();
            GroupsPasport();
            TeachersPassport();
            CabinetsPassport();
        }
        //парсинг направлений и групп
        public void ParseCoursesAndGroups()
        {
            int CourseId = 1;
            int GroupId = 1;
            for (int i = 3; ; i++)
            {
                Excel.Range CellRange = UsedRange[9, i];
                string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                    (CellRange as Excel.Range).Value2.ToString();
                if (CellText != null)
                {
                    string[] group = CellText.Split(new char[] { '-' });
                    if (Courses.Where(x => x.Shortname == group[0]).FirstOrDefault() == null)
                    {
                        Courses.Add(new Course
                        {
                            Id = CourseId,
                            Name = group[0],
                            Shortname = group[0]
                        });
                        CourseId++;
                    }
                    Course CurrentCourse = Courses.Where(x => x.Name == group[0]).FirstOrDefault();
                    Groups.Add(new Models.Group
                    {
                        Id = GroupId,
                        Course = CurrentCourse,
                        CourseId = CurrentCourse.Id,
                        Code = group[1]
                    });
                    GroupId++;
                }
                else break;
            }
        }
        //парсинг преподов
        public void ParseTeachers()
        {
            int teacherId = 1;
            Regex regex = new Regex(@"[А-ЯЁ][а-яё]*([-][А-ЯЁ][а-яё]*)?\\s[А-ЯЁ]\\.?[А-ЯЁ]\\.?");

			for (int x = 3; x < 36; x++)
            {
                for (int y = 10; y < 159; y++)
                {
                    Range CellRange = UsedRange.Cells[x, y];
                    string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                        (CellRange as Excel.Range).Value2.ToString();
                    if (CellText != null)
                    {
                        MatchCollection matches = regex.Matches(CellText);
                        if (matches.Count > 0)
                        {
                            foreach (Match match in matches)
                            {
                                string[] s = match.ToString().Trim().Split(' ', '.');
                                if (Teachers.Where(z=>z.Surname == s[0] && z.Name == s[1] && z.Patronymic == s[2]).Count() == 0)
                                {
                                    Teachers.Add(new Teacher
                                    {
                                        Id = teacherId,
                                        Surname = s[0],
                                        Name = s[1],
                                        Patronymic = s[2]
                                    });
                                }
                            }
                        }
/*                        if (CellText.Contains(" ") && CellText.Contains("."))
                        {
                            string[] teacher = CellText.Split(new char[] { ' ', '.' }, StringSplitOptions.RemoveEmptyEntries);
                            if (teacher.Length == 3 || teacher.Length == 5)
                            {
                                if (Teachers.Where(t => t.Name == teacher[1] && t.Patronymic == teacher[2] && t.Surname == teacher[0]).FirstOrDefault() == null
                                    && teacher[1].Length == 1 && teacher[2].Length == 1 && teacher[0].Length > 2)
                                {
                                    Teachers.Add(new Teacher
                                    {
                                        Id = teacherId,
                                        Name = teacher[1],
                                        Surname = teacher[0],
                                        Patronymic = teacher[2],
                                    });
                                    teacherId++;
                                }
                            }
                        }*/
                    }
                }
            }
        }
        //парсинг кабинетов
        public void ParseCabinets()
        {
            int cabinetId = 1;
            Regex regex = new Regex("ауд\\. \\d+");
            for (int x = 3; x < 36; x++)
            {
                for (int y = 10; y < 159 ;y++)
                {
                    Range CellRange = UsedRange.Cells[x, y];
                    string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                        (CellRange as Excel.Range).Value2.ToString();
                    if (CellText != null)
                    {
                        MatchCollection matches = regex.Matches(CellText);
                        if (matches.Count > 0)
                        {
                            foreach(Match match in matches)
                            {
                                string[] s = match.ToString().Trim().Split(' ');
                                if (Cabinets.Where(z=>z.Number == s[1]).Count() == 0)
                                {
                                    Cabinets.Add(new Cabinet
                                    {
                                        Id = cabinetId,
                                        Number = s[1],
                                    });
                                    cabinetId++;
                                }
                            }
                        }
                    }
                }
            }
        }
        public void ParseObjects()
        {
            int subjectId = 1;
            for (int x = 3; x < 36; x++)
            {
                for (int y = 10; y < 159; y++)
                {
                    Range CellRange = UsedRange.Cells[x, y];
                    string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                        (CellRange as Excel.Range).Value2.ToString();
                    if (CellText != null)
                    {

                    }
                }
            }
        }
        //вывод всех направлений
        public void CoursesPasport()
        {
            Console.WriteLine("Направления");
            foreach (Course course in Courses)
                Console.WriteLine($"{course.Id.ToString()} {course.Name} {course.Shortname}");
        }
        //вывод всех групп
        public void GroupsPasport()
        {
            Console.WriteLine("Группы");
            foreach (var group in Groups)
                Console.WriteLine($"{group.Id.ToString()} {group.CourseId.ToString()} {group.Code}");
        }
        //вывод всех преподов
        public void TeachersPassport()
        {
            Console.WriteLine("Преподаватели");
            foreach (Teacher teacher in Teachers)
                Console.WriteLine($"{teacher.Id} {teacher.Surname} {teacher.Name} {teacher.Patronymic}");
        }
        //вывод всех кабинетов
        public void CabinetsPassport()
        {
            Console.WriteLine("Кабинеты");
            foreach (Cabinet cabinet in Cabinets)
                Console.WriteLine($"{cabinet.Id} {cabinet.Number}");
        }
    }
}
