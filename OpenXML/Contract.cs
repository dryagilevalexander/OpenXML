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

        public List<Condition> CreateConditions(Contract contract, ConditionsService conditionService)
        {
            string contractType = "";
            string contractName = "";
            string baseOfContract = "";
            string paragraphBaseOfContract = "";
            List<Condition> conditions = new List<Condition>();

            //Получаем тип договора
            switch (contract.ContractType)
            {
                case 1:
                    contractType = "подряда";
                    break;
                case 2:
                    contractType = "оказания услуг";
                    break;
                case 3:
                    contractType = "поставки";
                    break;
                case 4:
                    contractType = "аренды";
                    break;
            }

            //Получаем пункт основания заключения контракта (для 44-ФЗ)
            if (contract.RegulationType == 3)
            {
                switch (contract.RegulationParagraph)
                {
                    case 1:
                        paragraphBaseOfContract = "п. 4 ст. 93 ";
                        break;
                    case 2:
                        paragraphBaseOfContract = "п. 8 ст. 93 ";
                        break;
                }
            }

            //Получаем фактическое наименование контракта и основание заключения
            switch (contract.RegulationType)
            {
                case 1:
                    contractName = "Договор";
                    break;
                case 2:
                    contractName = "Договор";
                    baseOfContract = "на основании федерального закона \"О закупках товаров, работ, услуг отдельными видами юридических лиц\" от 18.07.2011 N 223-ФЗ,";
                    break;
                case 3:
                    contractName = "Контракт";
                    baseOfContract = "на основании " + paragraphBaseOfContract + "федерального закона \"О контрактной системе в сфере закупок товаров, работ, услуг для обеспечения государственных и муниципальных нужд\" от 05.04.2013 N 44-ФЗ,";
                    break;
            }

            //Определяем шапку договора
            Condition headOfContract = new Condition()
            {
                Id = 1,
                TypeOfCondition = 1,
                Name = contractName + " " + contractType + " № __"
            };
            conditions.Add(headOfContract);

            Condition preambleOfContract = new Condition()
            {
                Id = 2,
                TypeOfCondition = 2,
                Text = contract.MainOrganization.Name + " именуемое в дальнейшем \"Заказчик\", в лице директора " + contract.MainOrganization.DirectorNameR + ", действующего на основании Устава, с одной стороны, и " + contract.Contragent.Name + ", именуемое в дальнейшем \"Подрядчик\", в лице директора " + contract.Contragent.DirectorNameR + ", действующего на основании Устава, с другой стороны, " + baseOfContract + " заключили настоящий " + contractName + " о нижеследующем:"
            };

            conditions.Add(preambleOfContract);

            //1. Предмет контракта
            if (contract.ContractType == 1)
            {
                Condition subjectOfContract = conditionService.GetConditionById(2);
                subjectOfContract.Name = subjectOfContract.Name.Replace("договор", contractName);
                if (subjectOfContract.SubConditions != null)
                {
                    foreach (var subCondition in subjectOfContract.SubConditions)
                    {
                        subCondition.Text = subCondition.Text.Replace("<subjectOfContract>", contract.SubjectOfContract);
                        subCondition.Text = subCondition.Text.Replace("договор", contractName);
                    }
                }
                conditions.Add(subjectOfContract);
            }

            //2. Права и обязанности сторон
            Condition rightsAndDuties = conditionService.GetConditionById(3);
            rightsAndDuties.Name = rightsAndDuties.Name.Replace("договор", contractName);
            if (rightsAndDuties.SubConditions != null)
            {
                foreach (var subCondition in rightsAndDuties.SubConditions)
                {
                    subCondition.Text = subCondition.Text.Replace("договор", contractName);
                    if (subCondition.SubConditionParagraphs != null)
                    {
                        foreach (var subConditionParagraph in subCondition.SubConditionParagraphs)
                        {
                            subConditionParagraph.Text = subConditionParagraph.Text.Replace("<dateEnd>", contract.DateEnd);
                            subConditionParagraph.Text = subConditionParagraph.Text.Replace("договор", contractName);
                        }
                    }
                }
            }
            conditions.Add(rightsAndDuties);


            //3. Включаем ответственность сторон
            if(contract.RegulationType == 3) //44-ФЗ
            {
                conditions.Add(conditionService.GetConditionById(1));
            }

            return conditions;
        }


        //Метод получения реквизитов
        public Dictionary<string, string> GetRequisites(Contragent contragent)
        {
            Dictionary<string, string> props = new Dictionary<string, string>()
            {
                    { "Наименование:", contragent.Name},
                    { "ИНН", contragent.INN},
                    { "КПП", contragent.KPP},
                    { "Директор _________ ", contragent.DirectorName}
            };
            return props;
        }
    }
}
