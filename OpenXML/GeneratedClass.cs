using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using System.Reflection.Metadata;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;

namespace OpenXML
{

    public class GeneratedClass
    {
        public void OpenAndAddTextToWordDocument(string filepath, string txt)
        {
            //Создаем новый документ и конфигурируем его
            WordprocessingDocument wordDoc = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
            Body body = mainPart.Document.AppendChild(new Body());


            for (int i = 0; i < 10; i++)
            {     
            //Создаем разделы и пишем текст
            Paragraph para = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties = new ParagraphMarkRunProperties();
            RunFonts runFonts = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
            FontSize fontSize1 = new FontSize() { Val = "144" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "144" };
            
            paragraphMarkRunProperties.Append(fontSize1);
            paragraphMarkRunProperties.Append(fontSizeComplexScript1);
            paragraphMarkRunProperties.Append(runFonts);
            paragraphProperties.Append(paragraphMarkRunProperties);


            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId1 = new NumberingId() { Val = 3 };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);

            paragraphProperties.Append(numberingProperties1);
            para.Append(paragraphProperties);




            Run run = para.AppendChild(new Run());
            RunProperties runProperties = new RunProperties();
            //Тип и размер шрифта
            RunFonts runFonts1 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
            FontSize fontSize2 = new FontSize() { Val = "144" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "144" };

            runProperties.Append(runFonts1);
            runProperties.Append(fontSize2);
            runProperties.Append(fontSizeComplexScript2);
            run.Append(runProperties);
            run.AppendChild(new Text(txt));
            }

            //Закрываем файл
            wordDoc.Close();
        }
    }
}

