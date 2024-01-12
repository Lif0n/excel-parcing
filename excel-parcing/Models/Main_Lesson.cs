using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_parcing.Models
{
    [Table("lesson_plan")]
    public class Main_Lesson
	{

		public int Id { get; set; }
		public int LessonNumber { get; set; }
		public Subject Subject { get; set; }
		public Cabinet? Audience { get; set; }
		public Group Group { get; set; }
		public bool isDistantce { get; set; }
		public int Weekday { get; set; }
		public int? WeekNumber { get; set; }
		[NotMapped]
        public List<Teacher> Teachers { get; set; }
    }
}
