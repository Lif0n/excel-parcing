using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_parcing.Models
{
    public class LessonTeacher
    {
        public int Id { get; set; }
        [Column("Scheduled-lesson")]
        public Main_Lesson Lesson { get; set; }
        public Teacher Teacher { get; set; }
        public bool IsGeneral { get; set; }
    }
}
