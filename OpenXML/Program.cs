using OpenXML;

ContragentsService contragentsService = new ContragentsService();
ConditionsService conditionsService = new ConditionsService();

Contragent contragentMain = contragentsService.GetMainOrganization();
Contragent contragent = contragentsService.GetContragentById(2); 

Contract contract = new Contract()
{
    ContractType = 1,
    RegulationType = 2,
    RegulationParagraph = 2,
    Contragent = contragent,
    MainOrganization = contragentMain,
    SubjectOfContract = "Работы по ремонту теплотрассы в р.п. Некрасовское",
    DateStart = new DateTime(2023, 3, 20).ToShortDateString(),
    DateEnd = new DateTime(2023,12,31).ToShortDateString()

};

contract.Conditions = contract.CreateConditions(contract, conditionsService);
contract.MainProp = contract.GetRequisites(contragentMain);
contract.ContragentProp = contract.GetRequisites(contragent);


new DocumentGenerator().CreateContract(@"C:\AIS\Output.docx", contract);
