using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace OpenXML
{
    public class ConditionsService
    {
        ApplicationContext db;
        public ConditionsService()
        {
            db = new ApplicationContext();
        }
        public Condition GetConditionById(int id)
        {
            return db.Conditions.Include(p => p.SubConditions).ThenInclude(c => c.SubConditionParagraphs).FirstOrDefault(p => p.Id == id);
        }
    }
}
