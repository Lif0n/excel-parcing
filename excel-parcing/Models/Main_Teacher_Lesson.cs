using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_parcing.Models
{
	public class Main_Teacher_Lesson
	{
		public int Id { get; set; }
		public int LessonId { get; set; }
		public Main_Lesson Main_Lesson{get;set; }
		public int TeacherId { get; set; }
		public Teacher Teacher { get; set; }
		public bool IsMain { get; set; }
	}
}
