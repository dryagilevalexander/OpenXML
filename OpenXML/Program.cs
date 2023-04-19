using OpenXML;

Contragent contragentMain = new Contragent()
{
    id = 1,
    IsMain = true,
    Name = "ООО \"Альфа\"",
    INN = "7777777",
    KPP = "7000001",
    DirectorName = "Иванов И.И.",
    DirectorNameR = "Иванова И.И."
};

Contragent contragent = new Contragent()
{
    id = 1,
    IsMain = false,
    Name = "ООО \"Бетта\"",
    INN = "7555555",
    KPP = "7000002",
    DirectorName = "Петров А.А.",
    DirectorNameR = "Петрова А.А."
};

Contract contract = new Contract()
{
    ContractType = 1,
    RegulationType = 3,
    RegulationParagraph = 2,
    Contragent = contragent,
    MainOrganization = contragentMain,
    SubjectOfContract = "Работы по ремонту теплотрассы в р.п. Некрасовское",
    DateStart = new DateTime(2023, 3, 20).ToShortDateString(),
    DateEnd = new DateTime(2023,12,31).ToShortDateString()

};

new GeneratedClass().CreateWordDocument(@"C:\AIS\Output.docx", contract);
