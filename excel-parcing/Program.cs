using excel_parcing.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace excel_parcing
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<Course> courses = new List<Course>();
            List<Group> groups = new List<Group>();

            Console.Write("Введите путь файла: ");
            string path = Console.ReadLine();
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
                    // Получаем строки в используемом диапазоне
                    Excel.Range urRows = UsedRange.Rows;
                    // Получаем столбцы в используемом диапазоне
                    Excel.Range urColums = UsedRange.Columns;

                    int IdCourse = 1;
                    int IdGroup = 1;
                    for (int i = 3; ;i++ )
                    {
                        Excel.Range CellRange = UsedRange.Cells[9,i];
                        string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                        (CellRange as Excel.Range).Value2.ToString();
                        //Test
                        if (CellText != null)
                        {
                            string[] group = CellText.Split(new char[] { '-'});
                            if (courses.Where(x => x.Shortname == group[0]).FirstOrDefault() == null)
                            {
                                courses.Add(new Course { Id = IdCourse, Name = group[0], Shortname = group[0] });
                                IdCourse++;
                            }

                            groups.Add(new Group {Id = IdGroup, IdCourse = courses.Where(x=>x.Name == group[0]).FirstOrDefault().Id, Code = group[1] });
                            IdGroup++;

                        }
                        else break;
                    }
                }
                Console.WriteLine("Направления");
                foreach (Course course in courses)
                    Console.WriteLine($"{course.Id.ToString()} {course.Name} {course.Shortname}");
                Console.WriteLine("Группы");
                foreach (Group group1 in groups)
                    Console.WriteLine($"{group1.Id.ToString()} {group1.IdCourse.ToString()} {group1.Code}");
            }

            catch (Exception ex) 
            {
                Console.WriteLine(ex);
            }
            Console.ReadKey();
        }
    }
}
