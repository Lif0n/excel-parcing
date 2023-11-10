using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_parcing.Models
{
    public class Course
    {
        public int Id { get; set; }

        public int? IdSpeciality = null;
        public string Name { get; set; }
        public string Shortname { get; set; }
    }
}
