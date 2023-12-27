using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_parcing.Models
{
	public class Main_Lesson
	{
		public int Id { get; set; }
		public int LessonNumber { get; set; }
		public int SubjectId { get;set; }
		public Subject Subject { get; set; }
		public int? CabinetId { get; set; }
		public Cabinet Cabinet { get; set; }
		public int GroupId { get; set; }
		public Group Group { get; set; }
		public bool isDistantсe { get; set; }
		public int Weekday { get; set; }
		public int? WeekNumber { get; set; }

	}
}
