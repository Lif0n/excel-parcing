using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_parcing.Models
{
    [Table("Speciality")]
    public class Course
    {
        public int Id { get; set; }
        public string? Code { get; set; }
        public string Name { get; set; }
        public string Shortname { get; set; }
    }
}
