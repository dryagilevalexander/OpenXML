using DocumentFormat.OpenXml.Office2010.Excel;
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
        public int RegulationType { get; set; } //1 - ГК, 2 - 223-ФЗ, 3 - 44-ФЗ
        public int RegulationParagraph { get; set; } //Только для 44-фз: 1 - п.4 ст. 93, 2 п.8 ст. 93 по умолчанию 0
        public bool IsPerpetual { get; set; }
        public Contragent Contragent { get; set; }
        public Contragent MainOrganization { get; set; }
        public string SubjectOfContract { get; set; }
        public string DateStart { get; set; }
        public string DateEnd { get; set; }
        public List<Condition> Conditions { get; set; }
        public Dictionary<string, string> MainProp { get; set; }
        public Dictionary<string, string> ContragentProp { get; set; }

        public List<Condition> CreateConditions(ConditionsService conditionService)
        {
            List<Condition> conditions = new List<Condition>();
            ContractTemplate contractTemplate = conditionService.GetContractTemplateId(ContractTemplateId);
            foreach(var condition in contractTemplate.Conditions) 
            { 
                if(condition.RegulationType == 4)
                {
                    conditions.Add(condition);
                }
            
                if(RegulationType == 3)
                {
                    if (condition.RegulationType == 3)
                    {
                        conditions.Add(condition);
                    }
                }
            }
            
            return conditions;
        }

        //Метод получения реквизитов
        public Dictionary<string, string> GetRequisites(Contragent contragent)
        {
            Dictionary<string, string> props = new Dictionary<string, string>()
            {
                    {contragent.ShortName,""},                   
                    {"ИНН", contragent.INN},
                    {"КПП", contragent.KPP},
                    {"ОГРН", contragent.OGRN},
                    {"Адрес", contragent.Address},
                    {"Банк", contragent.Bank},
                    {"БИК", contragent.BIK},
                    {"р/с", contragent.Account},
                    {"к/с", contragent.CorrespondentAccount},
                    { "Директор _________ ", contragent.DirectorName}
            };
            return props;
        }
    }
}
