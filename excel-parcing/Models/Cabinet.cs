using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_parcing.Models
{
    public class Cabinet
    {
        public int Id { get; set; }
        public string Number { get; set; }
        public int? CabinetTypeId = null;
        public CabinetType CabinetType { get; set; } = null;
    }
}
