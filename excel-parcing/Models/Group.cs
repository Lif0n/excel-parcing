using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_parcing.Models
{
    public class Group
    {
        public int Id { get; set; }
        public Course Speciality { get; set; }
        public string Name { get; set; }
    }
}
