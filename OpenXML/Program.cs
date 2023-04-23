using OpenXML;

ContragentsService contragentsService = new ContragentsService();
ContractService contractService = new ContractService();

Contragent mainOrganization = contragentsService.GetMainOrganization();
Contragent contragent = contragentsService.GetContragentById(2);

Contract contract = new Contract()
{
    ContractType = 2,
    ContractTemplateId = 3,
    IsCustomer = true,
    RegulationType = 3,
    RegulationParagraph = 2,
    SubjectOfContract = "Оказание услуг по ремонту офисной техники",
    PlaceOfContract = "рп. Некрасовское",
    DateStart = new DateTime(2023, 3, 20).ToShortDateString(),
    DateEnd = new DateTime(2023,12,31).ToShortDateString()
};

contractService.CreateConditions(contract);
contractService.SetContractRequisites(contract, mainOrganization, contragent);

new DocumentGenerator().CreateContract(@"C:\AIS\Output.docx", contract);
