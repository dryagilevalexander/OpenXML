using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace OpenXML
{
    public class ContragentsService
    {
        ApplicationContext db;
        public ContragentsService()
        {
            db = new ApplicationContext();
        }
        public Contragent GetContragentById(int id)
        {
            return db.Contragents.FirstOrDefault(p => p.Id == id);
        }
        public Contragent GetMainOrganization()
        {
            return db.Contragents.FirstOrDefault(p => p.IsMain == true); 
        }
    }
}
