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
using System.Runtime.CompilerServices;
using System.Net.Http.Headers;

namespace excel_parcing
{
    public class ParseConfig
    {
        public int StartColumn { get; set; }
        public int StartRow { get; set; }
        public int[] SkipRows { get; set; }
        public int EndColumn { get; set; }
        public int CountRows { get; set; }
    }
    public class Parsing
    {
        private ParseConfig _config;
        public Parsing(string path)
        {
            var str = LevenshteinDistance("Введение в прф. деятельность", "Введение в проф. деятельность");
            var str1 = LevenshteinDistance("Введение в проф. деятельность", "Введение в проф. деятельность");
            var str2 = LevenshteinDistance("Введение в проф деятельность", "Введение в проф. деятельность");
            var str3 = LevenshteinDistance("Ввведение в проф деятельность", "Введение в проф. деятельность");

            Path = path;
            _config = new ParseConfig
            {
                StartColumn = 3,
                StartRow = 10,
                EndColumn = 36,
                SkipRows = new[] { 34, 59, 84, 109, 134 },
                CountRows = 159
            };
        }

        public string Path { get; set; }
        public List<Cabinet> Cabinets = new List<Cabinet>();
        public List<Course> Courses = new List<Course>();
        public List<Models.Group> Groups = new List<Models.Group>();
        public List<Subject> Subjects = new List<Subject>();
        public List<Teacher> Teachers = new List<Teacher>();
        public List<Teacher_Subject> Teacher_Subjects = new List<Teacher_Subject>();
        public List<Main_Lesson> Main_Lessons = new List<Main_Lesson>();
        public List<Main_Teacher_Lesson> Main_Teacher_Lessons = new List<Main_Teacher_Lesson>();
        Excel.Range UsedRange = null;

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

                /*                ParseCoursesAndGroups();
                                ParseTeachers();
                                ParseCabinets();*/
                ParseLessons();
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

                List<Task> tasks = new List<Task>()
                {
                    new Task(ParseCoursesAndGroups),
                    new Task(ParseTeachers),
                    new Task(ParseCabinets),
                    new Task(ParseSubjects),
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
            SubjectsPassport();
        }
        //парсинг направлений и групп
        public void ParseCoursesAndGroups()
        {
            int CourseId = 1;
            int GroupId = 1;
            int lastRow = _config.StartRow + _config.CountRows;
            for (int i = 3; i <= lastRow; i++)
            {
                Range CellRange = UsedRange[9, i];
                string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                    (CellRange as Excel.Range).Value2.ToString();
                if (CellText == null)
                {
                    continue;
                }
                string[] group = CellText.Split(new char[] { '-' });
                if (Courses.Where(x => x.Shortname == group[0]).FirstOrDefault() != null)
                {
                    continue;
                }
                Courses.Add(new Course
                {
                    Id = CourseId,
                    Name = group[0],
                    Shortname = group[0]
                });
                CourseId++;
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
        }
        //парсинг преподов
        public void ParseTeachers()
        {
            int teacherId = 1;
            Regex regex = new Regex(@"[А-ЯЁа-яё\-]+ [А-ЯЁ]\.\s*[А-ЯЁ]\.*");


            for (int x = _config.StartColumn; x < _config.EndColumn; x++)
            {
                for (int y = _config.StartRow; y < _config.CountRows; y++)
                {
                    Range CellRange = UsedRange.Cells[y, x];
                    string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                        (CellRange).Value2.ToString();
                    if (CellText == null)
                    {
                        continue;
                    }

                    MatchCollection matches = regex.Matches(CellText);

                    if (matches.Count == 0)
                    {
                        continue;
                    }

                    foreach (Match match in matches)
                    {
                        string[] s = match.ToString().Trim().Split(' ', '.');
                        if (Teachers.Any(z => z.Surname == s[0] && z.Name == s[1] && z.Patronymic == s[2]))
                        {
                            continue;
                        }
                        Teachers.Add(new Teacher
                        {
                            Id = teacherId,
                            Surname = s[0],
                            Name = s[1],
                            Patronymic = s[2]
                        });
                        teacherId++;
                    }
                }
            }
        }
        //парсинг кабинетов
        public void ParseCabinets()
        {
            int cabinetId = 1;
            Regex regex = new Regex(@"ауд\.\s*\d+");
            for (int x = 3; x < 36; x++)
            {
                for (int y = 10; y < 159; y++)
                {
                    Range CellRange = UsedRange.Cells[y, x];
                    string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                        (CellRange as Excel.Range).Value2.ToString();
                    if (CellText == null)
                    {
                        continue;
                    }
                    MatchCollection matches = regex.Matches(CellText);
                    if (matches.Count == 0)
                    {
                        continue;
                    }
                    foreach (Match match in matches)
                    {
                        string[] s = match.ToString().Trim().Split(' ', '.').Where(m => m != "").ToArray();
                        if (Cabinets.Where(z => z.Number == s[1]).Count() == 0)
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
        public void ParseSubjects()
        {
            Regex regCab = new Regex(@"ауд\.\s*\d+");
            Regex regTeacher = new Regex(@"[А-ЯЁа-яё\-]+ [А-ЯЁ]\.\s*[А-ЯЁ]\.*");
            for (int x = 3; x < 36; x++)
            {
                for (int y = 10; y < 159; y += 2)
                {
                    if (_config.SkipRows.Contains(y))
                        y++;

                    Range CellRange = UsedRange.Cells[y, x];
                    Range NextCellRange = UsedRange.Cells[y + 1, x];

                    string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                        (CellRange as Excel.Range).Value2.ToString();

                    string NextCellText = (NextCellRange == null || NextCellRange.Value2 == null) ? null :
                        (NextCellRange as Excel.Range).Value2.ToString();

                    string s = String.Concat(CellText, " ", NextCellText);

                    MatchCollection cabMatches = regCab.Matches(s);
                    MatchCollection teacherMatches = regTeacher.Matches(s);
                    if (CellRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight == 2)
                    {
                        continue;
                    }

                    foreach (Match m in cabMatches)
                    {
                        s = s.Replace(m.ToString(), "");
                    }

                    foreach (Match m in teacherMatches)
                    {
                        s = s.Replace(m.ToString(), "");
                    }
                    
                    s = s.Contains("Космонавта Комарова 55") ? s.Replace("Космонавта Комарова 55", "") : s;

                    foreach (Match m in new Regex(@"\d{3}").Matches(s))
                        s = s.Replace(m.ToString(), "");

                    s = s.Trim();
                    if (String.IsNullOrEmpty(s))
                    {
                        continue;
                    }
                    GetOrCreate(s);
                }
            }
        }



















































        public void ParseLessons()
        {
            Regex regCab = new Regex(@"ауд\.\s*\d+");
            Regex regTeacher = new Regex(@"[А-ЯЁа-яё\-]+ [А-ЯЁ]\.\s*[А-ЯЁ]\.*");
            int id = 1;
            int lessonNumber = 1;
            for (int x = 3; x < 36; x++)
            {
                Models.Group group = CheckGroup(UsedRange.Cells[9, x].Value2 == null ? "" : UsedRange.Cells[9, x].Value2.ToString());
                int weekday = 1;
                for (int y = 10; y < 159;)
                {
                    if (lessonNumber >6)
                    {
                        y++;
                        weekday++;
                        lessonNumber = 1;
                    }
                    string s = "";
                    string s1 = "";
                    bool isOne = true;
                    for (int i = 0; i < 4; i++)
                    {
                        isOne = true;
                        Range CellRange = UsedRange.Cells[y, x];
						//-4138
						s += CellRange.Value2 == null ? "" : CellRange.Value2.ToString() + " ";
						if (i == 1)
                        {
                            if (CellRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight == -4138)
                            {
                                isOne = false;
                                s1 += UsedRange.Cells[y + 1, x].Value2 == null ? "" : UsedRange.Cells[y + 1, x].Value2.ToString() + " ";
                                s1 += UsedRange.Cells[y + 2, x].Value2 == null ? "" : UsedRange.Cells[y + 2, x].Value2.ToString() + " ";
                                y += 3;
                                break;
                            }
                        }
                        y++;
					}
					lessonNumber++;
					if (!string.IsNullOrWhiteSpace(s) || !string.IsNullOrWhiteSpace(s1))
                    {
                        if (!string.IsNullOrEmpty(s))
                        {
							Cabinet cabinet = CheckCabinet(s);
							s = DeleteCabinet(s);
							Teacher[] lessonsTeachers = CheckTeacher(s);
							s = DeleteTeacher(s);
                            Subject subject= CheckSubject(s);
							Main_Lesson lesson = new Main_Lesson
							{
								Id = Main_Lessons.Count() + 1,
								LessonNumber = lessonNumber,
								Weekday = weekday,
								isDistantсe = false,
								GroupId = group.Id,
                                SubjectId = subject.Id,
							};
							if (isOne == false)
							{
								lesson.WeekNumber = 1;
							}

                            if (cabinet == null)
                                lesson.CabinetId = null;
                            else
                                lesson.CabinetId = cabinet.Id;

                            Main_Lessons.Add(lesson);
						}
                        if (!string.IsNullOrEmpty(s1))
                        {
							Cabinet cabinet = CheckCabinet(s1);
							s = DeleteCabinet(s1);
							Teacher[] lessonsTeachers = CheckTeacher(s1);
							s = DeleteTeacher(s1);
							Subject subject = CheckSubject(s1);
							Main_Lesson lesson = new Main_Lesson
							{
								Id = Main_Lessons.Count() + 1,
								LessonNumber = lessonNumber,
								Weekday = weekday,
								isDistantсe = false,
								GroupId = group.Id,
								CabinetId = cabinet.Id,
								SubjectId = subject.Id,
                                WeekNumber = 2
							};

							if (cabinet == null)
								lesson.CabinetId = null;
							else
								lesson.CabinetId = cabinet.Id;

							Main_Lessons.Add(lesson);
						}
					}
                }
            }
        }






        private Models.Group CheckGroup(string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                string[] groupText = text.Split(new char[] { '-' });
                if (Courses.Any(x => x.Shortname == groupText[0]) == false)
                {
                    Courses.Add(new Course
                    {
                        Id = Courses.Count() + 1,
                        Name = groupText[0],
                        Shortname = groupText[0]
                    });
                }
                Course CurrentCourse = Courses.Where(x => x.Name == groupText[0]).FirstOrDefault();
                Models.Group group = new Models.Group
                {
                    Id = Groups.Count() + 1,
                    CourseId = CurrentCourse.Id,
                    Code = groupText[1]
                };
                Groups.Add(group);
                return group;
            }
            else return null;
        }

        private Cabinet CheckCabinet(string s)
        {
            Cabinet cabinet = new Cabinet();
			Regex regex = new Regex(@"ауд\.\s*\d+");
			MatchCollection matches = regex.Matches(s);
			if (matches.Count == 0)
			{
				return null;
			}
			foreach (Match match in matches)
			{
				string[] cab = match.ToString().Trim().Split(' ', '.').Where(m => m != "").ToArray();
				if (Cabinets.Any(z => z.Number == cab[1]) == false)
				{
					cabinet = new Cabinet
					{
						Id = Cabinets.Count()+1,
						Number = cab[1],
					};
					Cabinets.Add(cabinet);
				}
                else
                {
                    cabinet = Cabinets.Where(z => z.Number == cab[1]).FirstOrDefault();
                }
			}
			return cabinet;
        }
        private string DeleteCabinet(string s)
        {
			Regex regex = new Regex(@"ауд\.\s*\d+");
			MatchCollection matches = regex.Matches(s);
			foreach (Match match in matches)
            {
                s = s.Replace(match.ToString(), "");
            }
            return s.Trim();
		}
        private Subject CheckSubject(string s)
        {
			s = s.Contains("Космонавта Комарова 55") ? s.Replace("Космонавта Комарова 55", "") : s;

			foreach (Match m in new Regex(@"\d{3}").Matches(s))
				s = s.Replace(m.ToString(), "");

			s = s.Trim();
			return GetOrCreate(s);
		}
        private Teacher[] CheckTeacher(string CellText)
        {
			Regex regex = new Regex(@"[А-ЯЁа-яё\-]+ [А-ЯЁ]\.\s*[А-ЯЁ]\.*");
			MatchCollection matches = regex.Matches(CellText);
            Teacher[] lessonTeachers = new Teacher[2];
			foreach (Match match in matches)
			{
				string[] s = match.ToString().Trim().Split(' ', '.');
				if (Teachers.Any(z => z.Surname == s[0] && z.Name == s[1] && z.Patronymic == s[2]))
				{
                    if (lessonTeachers[0] == null)
                    {
                        lessonTeachers[0] = Teachers.Where(z => z.Surname == s[0] && z.Name == s[1] && z.Patronymic == s[2]).FirstOrDefault();
					}
                    else
                    {
						lessonTeachers[1] = Teachers.Where(z => z.Surname == s[0] && z.Name == s[1] && z.Patronymic == s[2]).FirstOrDefault();
					}
				}
                else
                {
					Teachers.Add(new Teacher
					{
						Id = Teachers.Count()+1,
						Surname = s[0],
						Name = s[1],
						Patronymic = s[2]
					});
                    if (lessonTeachers[0] == null)
                    {
                        lessonTeachers[0] = Teachers.Last();
                    }
				}
			}
            return lessonTeachers;
		}
        private string DeleteTeacher(string s)
        {
			Regex regex = new Regex(@"[А-ЯЁа-яё\-]+ [А-ЯЁ]\.\s*[А-ЯЁ]\.*");
			MatchCollection matches = regex.Matches(s);
            foreach (Match match in matches) 
            {
                s = s.Replace(match.ToString(), "");
            }
            s.Trim();
            return s;
		}





















        private Subject GetOrCreate(string word)
        {
            foreach (var item in Subjects)
            {
                if (word.Contains(item.Name))
                {
                    return item;
                }
                if (IsMatch(item.Name, word))
                {
                    return item;
                }
            }
            Subject sbj = new Subject()
            {
                Name = word,
                Shortname = word
            };
            Subjects.Add(sbj);
            sbj.Id = Subjects.Max(x => x.Id) + 1;
            return sbj;
        }
        private bool IsMatch(string source, string target)
        {
            return LevenshteinDistance(source.ToLower(), target.ToLower()) < source.Length*0.2;
        }
        public int LevenshteinDistance(string source, string target)
        {
            if (String.IsNullOrEmpty(source))
            {
                if (String.IsNullOrEmpty(target)) return 0;
                return target.Length;
            }
            if (String.IsNullOrEmpty(target)) return source.Length;

            var m = target.Length;
            var n = source.Length;
            var distance = new int[2, m + 1];
            // Initialize the distance 'matrix'
            for (var j = 1; j <= m; j++) distance[0, j] = j;

            var currentRow = 0;
            for (var i = 1; i <= n; ++i)
            {
                currentRow = i & 1;
                distance[currentRow, 0] = i;
                var previousRow = currentRow ^ 1;
                for (var j = 1; j <= m; j++)
                {
                    var cost = (target[j - 1] == source[i - 1] ? 0 : 1);
                    distance[currentRow, j] = Math.Min(Math.Min(
                                distance[previousRow, j] + 1,
                                distance[currentRow, j - 1] + 1),
                                distance[previousRow, j - 1] + cost);
                }
            }
            return distance[currentRow, m];
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
            //Teachers = Teachers.OrderBy(x=>x.Surname).ToList();
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
        public void SubjectsPassport()
        {
            Console.WriteLine("Предметы");
            foreach (Subject subject in Subjects)
                Console.WriteLine($"{subject.Id} {subject.Name}");
        }
    }
}
