using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_parcing.Models
{
    [Table("Audience")]
    public class Cabinet
    {
        public int Id { get; set; }
        public string Number { get; set; }
        public CabinetType? AudienceType { get; set; } = null;
    }
}
