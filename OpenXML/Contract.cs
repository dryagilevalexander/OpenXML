﻿using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML
{
    public class Contract
    {
        public Contract()
        {
            IsPerpetual = false;
            RegulationParagraph = 0;
        }

        public int ContractType { get; set; }   //1 - подряд, 2 - услуги, 3 - поставка, 4 - аренда 
        public int ContractTemplateId { get; set; }
        public bool IsCustomer { get; set; } //Кто заказчик
        public int RegulationType { get; set; } //1 - ГК, 2 - 223-ФЗ, 3 - 44-ФЗ
        public int RegulationParagraph { get; set; } //Только для 44-фз: 1 - п.4 ст. 93, 2 п.8 ст. 93 по умолчанию 0
        public bool IsPerpetual { get; set; }
        public Contragent Customer { get; set; }
        public Contragent Executor { get; set; }
        public string SubjectOfContract { get; set; }
        public string DateStart { get; set; }
        public string DateEnd { get; set; }
        public List<Condition> Conditions { get; set; }
        public Dictionary<string, string> CustomerProp { get; set; }
        public Dictionary<string, string> ExecutorProp { get; set; }
    }
}
