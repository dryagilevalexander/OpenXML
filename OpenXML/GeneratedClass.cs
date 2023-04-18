using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Math;
using JustificationValues = DocumentFormat.OpenXml.Wordprocessing.JustificationValues;
using Justification = DocumentFormat.OpenXml.Wordprocessing.Justification;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;
using StyleValues = DocumentFormat.OpenXml.Wordprocessing.StyleValues;
using System;

namespace OpenXML
{

    public class GeneratedClass
    {

        public void CreateWordDocument(string filepath)
        {
            //Создаем новый документ
            WordprocessingDocument wordDoc = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document);
            //Создаем корень документа
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
            //Создаем Body
            Body body = mainPart.Document.AppendChild(new Body());

            //Определяем стили используемые в документе
            StyleDefinitionsPart part = mainPart.StyleDefinitionsPart;
            if (part == null) part = AddStylesPartToPackage(wordDoc);
            CreateAndAddNoSpaceParagraphStyle(part, "a3");

            //Создаем абзацы

            createParagraph(wordDoc, "a3", "24", false, 0, 0, "center", true, "Договор № __");
            CreateDateAndPlaceTable(wordDoc, "рп. Некрасовское", "__.__.202_");
            createParagraph(wordDoc, "a3", "24", false, 0, 0, "center", false, "");
            string condition = "ООО Тест, именуемое в дальнейшем \"Заказчик\", в лице директора Иванова И.И., действующего на основании Устава, с одной стороны, и Муниципальное унитарное предприятие Некрасовского муниципального района «Энергетический ресурс», именуемое в дальнейшем \"Подрядчик\", в лице директора Голубева В.В., действующего на основании Устава, с другой стороны,  заключили настоящий договор о нижеследующем:";
            createParagraph(wordDoc, "a3", "24", false, 0, 0, "both", false, condition);
            createParagraph(wordDoc, "a3", "24", false, 0, 0, "center", false, "");
            createParagraph(wordDoc, "a3", "24", true, 0, 1, "center", true, "Предмет договора и общие условия");
            condition = "Подрядчик обязуется выполнить по заданию Заказчика работу, указанную в пункте 1.2 настоящего договора, и сдать ее результат Заказчику, а Заказчик обязуется принять результат работы и оплатить его.";
            createParagraph(wordDoc, "a3", "24", true, 1, 1, "both", false, condition);
            condition = "Подрядчик обязуется выполнить следующую работу: тестовая работа, именуемую в дальнейшем \"Работа\".";
            createParagraph(wordDoc, "a3", "24", true, 1, 1, "both", false, condition);
            createParagraph(wordDoc, "a3", "24", false, 0, 0, "center", false, "");
            createParagraph(wordDoc, "a3", "24", false, 0, 1, "center", true, "Реквизиты");
            CreateTable(wordDoc, MainProp, ContragentProp);
            //Закрываем файл
            wordDoc.Close();
        }

        //Метод создания абзаца
        //Параметры:
        //wordDoc - ссылка на документ
        //styleId - id стиля применяемого к абзацу
        //paragraphfontSize - размер шрифта
        //isNumbering - применять ли нумерацию
        //numId - id группировки нумерации
        //justificationValue - горизонтальное выравнивание
        //txt - текст абзаца

        public void createParagraph(WordprocessingDocument wordDoc, string styleId, string paragraphfontSize, bool isNumbering, int numLevelReference, int numId, string justificationValue, bool isBold, string txt)
        {
            //Получаем корень документа
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;
            //Получаем Body
            Body body = mainPart.Document.Body;
            
            //Создаем абзац и его свойства
            Paragraph para = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties = new ParagraphProperties();


            //Определяем шрифты
            ParagraphMarkRunProperties paragraphMarkRunProperties = new ParagraphMarkRunProperties();
            RunFonts runFonts = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
            FontSize fontSize1 = new FontSize() { Val = paragraphfontSize };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = paragraphfontSize };
            paragraphMarkRunProperties.Append(fontSize1);
            paragraphMarkRunProperties.Append(fontSizeComplexScript1);
            paragraphMarkRunProperties.Append(runFonts);
            //если шрифт жирный
            if(isBold)
            {
                Bold bold = new Bold();
                paragraphMarkRunProperties.Append(bold);
            }

            paragraphProperties.Append(paragraphMarkRunProperties);
 
            if (isNumbering)
            {
            //Определяем нумерацию
            NumberingProperties numberingProperties = new NumberingProperties();
            NumberingLevelReference numberingLevelReference = new NumberingLevelReference() { Val = numLevelReference };
            NumberingId numberingId = new NumberingId() { Val = numId };
            numberingProperties.Append(numberingLevelReference);
            numberingProperties.Append(numberingId);
            paragraphProperties.Append(numberingProperties);
            }

            Justification justification = new Justification();
            switch (justificationValue)
            {
                case "left":
                    justification.Val = JustificationValues.Left;
                break;
                case "center":
                    justification.Val = JustificationValues.Center;
                break;
                case "right":
                    justification.Val = JustificationValues.Right;
                break;
                case "both":
                    justification.Val = JustificationValues.Both;
                break;
            }
            paragraphProperties.Append(justification);

            if (paragraphProperties.ParagraphStyleId == null) paragraphProperties.ParagraphStyleId = new ParagraphStyleId();
            paragraphProperties.ParagraphStyleId.Val = styleId;

            //Добавляем сформированные свойства в абзац
            para.Append(paragraphProperties);


            //Добавляем текст с определенными свойствами в абзац
            Run run = para.AppendChild(new Run());
            RunProperties runProperties = new RunProperties();
            //Тип и размер шрифта текста
            RunFonts runFonts1 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
            FontSize fontSize2 = new FontSize() { Val = paragraphfontSize };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = paragraphfontSize };
            runProperties.Append(runFonts1);
            runProperties.Append(fontSize2);
            runProperties.Append(fontSizeComplexScript2);
            //если шрифт жирный
            if (isBold)
            {
                Bold bold = new Bold();
                runProperties.Append(bold);
            }

            run.Append(runProperties);
            run.AppendChild(new Text(txt));
        }


        public static Paragraph GetParagraph(WordprocessingDocument wordDoc, string styleId, string paragraphfontSize, bool isNumbering, int numId, string justificationValue, string txt)
        {
            //Получаем корень документа
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;
            //Получаем Body
            Body body = mainPart.Document.Body;

            //Создаем абзац и его свойства
            Paragraph para = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();


            //Определяем шрифты
            ParagraphMarkRunProperties paragraphMarkRunProperties = new ParagraphMarkRunProperties();
            RunFonts runFonts = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
            FontSize fontSize1 = new FontSize() { Val = paragraphfontSize };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = paragraphfontSize };
            paragraphMarkRunProperties.Append(fontSize1);
            paragraphMarkRunProperties.Append(fontSizeComplexScript1);
            paragraphMarkRunProperties.Append(runFonts);
            paragraphProperties.Append(paragraphMarkRunProperties);

            if (isNumbering == true)
            {
                //Определяем нумерацию
                NumberingProperties numberingProperties = new NumberingProperties();
                NumberingLevelReference numberingLevelReference = new NumberingLevelReference() { Val = 0 };
                NumberingId numberingId = new NumberingId() { Val = numId };
                numberingProperties.Append(numberingLevelReference);
                numberingProperties.Append(numberingId);
                paragraphProperties.Append(numberingProperties);
            }

            Justification justification = new Justification();
            switch (justificationValue)
            {
                case "left":
                    justification.Val = JustificationValues.Left;
                    break;
                case "center":
                    justification.Val = JustificationValues.Center;
                    break;
                case "right":
                    justification.Val = JustificationValues.Right;
                    break;
            }
            paragraphProperties.Append(justification);

            if (paragraphProperties.ParagraphStyleId == null) paragraphProperties.ParagraphStyleId = new ParagraphStyleId();
            paragraphProperties.ParagraphStyleId.Val = styleId;

            //Добавляем сформированные свойства в абзац
            para.Append(paragraphProperties);


            //Добавляем текст с определенными свойствами в абзац
            Run run = para.AppendChild(new Run());
            RunProperties runProperties = new RunProperties();
            //Тип и размер шрифта текста
            RunFonts runFonts1 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
            FontSize fontSize2 = new FontSize() { Val = paragraphfontSize };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = paragraphfontSize };
            runProperties.Append(runFonts1);
            runProperties.Append(fontSize2);
            runProperties.Append(fontSizeComplexScript2);
            run.Append(runProperties);
            run.AppendChild(new Text(txt));
            return para;
        }

        //Метод создания стиля "No Spacing"
        public static void CreateAndAddNoSpaceParagraphStyle(StyleDefinitionsPart styleDefinitionsPart, string styleId)
        {
            // Access the root element of the styles part.
            Styles styles = styleDefinitionsPart.Styles;
            if (styles == null)
            {
                styleDefinitionsPart.Styles = new Styles();
                styleDefinitionsPart.Styles.Save();
            }


            Style style = new Style() { Type = StyleValues.Paragraph, StyleId = styleId };
            StyleName styleName = new StyleName() { Val = "No Spacing" };
            UIPriority uIPriority = new UIPriority() { Val = 1 };
            PrimaryStyle primaryStyle = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties.Append(spacingBetweenLines);

            style.Append(styleName);
            style.Append(uIPriority);
            style.Append(primaryStyle);
            style.Append(styleParagraphProperties);

            // Add the style to the styles part.
            styles.Append(style);
        }

        //Метод добавления раздела стилей в документ
        public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
        {
            StyleDefinitionsPart part;
            part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Styles root = new Styles();
            root.Save(part);
            return part;
        }

        private Dictionary<string, string> MainProp = new Dictionary<string, string>()
               {
                    { "Name:", "ООО Альфа"},
                    { "ИНН", "77777777"},
                    { "КПП", "701000001"}
               };

        private Dictionary<string, string> ContragentProp = new Dictionary<string, string>()
               {
                    { "Name:", "ООО Бетта"},
                    { "ИНН", "88888888"},
                    { "КПП", "701000001"}
               };

        //Метод добавления однострочной таблицы (шапкаб реквизиты)
        public static void CreateTable(WordprocessingDocument doc, Dictionary<string, string> mainProp, Dictionary<string, string> contragentProp)
        {
            //Получаем корень документа
            MainDocumentPart mainPart = doc.MainDocumentPart;
            
            //Получаем Body
            Body body = mainPart.Document.Body;

            //Создаем таблицу и ее свойства
            Table table = body.AppendChild(new Table());

            TableProperties tableProperties = new TableProperties();
            TableWidth tableWidth = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableLook tableLook = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

            tableProperties.Append(tableWidth);
            tableProperties.Append(tableLook);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "4672" };
            GridColumn gridColumn2 = new GridColumn() { Width = "4673" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);

            TableRow tableRow1 = new TableRow();

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "4672", Type = TableWidthUnitValues.Dxa };

            tableCellProperties1.Append(tableCellWidth1);
            foreach(var prop in mainProp)
            { 
            Paragraph paragraph = GetParagraph(doc, "a3", "24", false, 0, "left", prop.Key + " " + prop.Value);
            tableCell1.Append(paragraph);
            }
            tableCell1.Append(tableCellProperties1);


            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "4673", Type = TableWidthUnitValues.Dxa };

            tableCellProperties2.Append(tableCellWidth2);
            foreach (var prop in contragentProp)
            {
                Paragraph paragraph = GetParagraph(doc, "a3", "24", false, 0, "left", prop.Key + " " + prop.Value);
                tableCell2.Append(paragraph);
            }
            tableCell2.Append(tableCellProperties2);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);

            table.Append(tableProperties);
            table.Append(tableGrid1);
            table.Append(tableRow1);
        }

        //Метод добавления однострочной таблицы (место, дата)
        public static void CreateDateAndPlaceTable(WordprocessingDocument doc, string place, string dateTemplate)
        {
            //Получаем корень документа
            MainDocumentPart mainPart = doc.MainDocumentPart;

            //Получаем Body
            Body body = mainPart.Document.Body;

            //Создаем таблицу и ее свойства
            Table table = body.AppendChild(new Table());

            TableProperties tableProperties = new TableProperties();
            TableWidth tableWidth = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableLook tableLook = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

            tableProperties.Append(tableWidth);
            tableProperties.Append(tableLook);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "4672" };
            GridColumn gridColumn2 = new GridColumn() { Width = "4673" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);

            TableRow tableRow1 = new TableRow();

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "4672", Type = TableWidthUnitValues.Dxa };

            tableCellProperties1.Append(tableCellWidth1);
                Paragraph paragraph1 = GetParagraph(doc, "a3", "24", false, 0, "left", place);
                tableCell1.Append(paragraph1);
            tableCell1.Append(tableCellProperties1);


            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "4673", Type = TableWidthUnitValues.Dxa };

            tableCellProperties2.Append(tableCellWidth2);
            Paragraph paragraph2 = GetParagraph(doc, "a3", "24", false, 0, "right", dateTemplate);
            tableCell2.Append(paragraph2);
            tableCell2.Append(tableCellProperties2);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);

            table.Append(tableProperties);
            table.Append(tableGrid1);
            table.Append(tableRow1);
        }
    }
}

