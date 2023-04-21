using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML
{
    public class Contragent
    {
        public int Id { get; set; }
        public bool IsMain { get; set; }
        public string Name { get; set; }
        public string INN { get; set; }
        public string KPP { get; set; }
        public string DirectorName { get; set; }
        public string DirectorNameR { get; set; }
    }
}
