using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ComplianceAssessment
{
    public class StyleDesc // класс со свойствами, описывабщими стиль
    {
        public int LevelSt { get; set; } // уровень текста
        public double AfterSpace { get; set; } // интервал после 
        public double BeforeSpace { get; set; } // интервал перед
        public double BtwSapce { get; set; } // междустрочный интервал
        public double LeftIndent { get; set; } // отступ слева
        public double RigthtIndent { get; set; }    // отступ справа
        public double FirstLineIndent { get; set; } // отступ первой строки
        public string JustPar { get; set; } // выравнивание текста
        
        public double FontSize { get; set; }    // размер ширфта
        public bool Cursive { get; set; }   // курсив
        public bool Underlined { get; set; }    // подчеркнутый
        public bool Bold { get; set; }  // полужирный  
        public string Font { get; set; }    // название шрифта
        public string Color { get; set; }   // цвет шрифта
    }
    public class DocDesc    // класс, описывающий основные параметры создаваемого документа
    {
        public double ColTop { get; set; }  // верхний колонтитул
        public double ColBot { get; set; }  // нижний колонтитул
        public Dictionary<string, double> Margins { get; set; } // поля страницы
        public bool SpecCol { get; set; }   // особый колонтитул первой страницы
        public string Place { get; set; }   // положение номера на странице
        public string NumJust { get; set; } // выравнивание номера на странице

        /// <summary>
        /// Создание шаблона документа на основе введенных параметров
        /// </summary>
        /// <param name="styles">Стили в документе</param>
        /// <param name="desc">Описание документа</param>
        /// <param name="mainPath">Путь к файлу шаблона</param>
        public static void CreateTemplate(Dictionary<int, StyleDesc> styles, DocDesc desc, string mainPath)
        {
            // есть ли в папке со временными файлами шаблон документа
            if (File.Exists(mainPath + @"RulesRepository\Temp\template.docx"))
            {
                // если есть, удалить
                File.Delete(mainPath + @"RulesRepository\Temp\template.docx");
            }
            string path = mainPath + "template.docx";
            // перемещение шаблона, который будет изменен во временнцю папку
            string newPath = path.Replace("template.docx", @"RulesRepository\Temp\template.docx");
            // копирование шаблона документа
            File.Copy(path, newPath);
            // открытие созданного шаблона для проведения модификаций
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(newPath, true))
            {
                if (styles != null)
                {
                    foreach (var s in styles)   // проход по стилям документа
                    {
                        if (s.Key == 1) // если стиль - Обычный текст
                            ChangeStylePrps(wDoc.MainDocumentPart.Document.Descendants<Paragraph>().Last(), s.Value);
                        if (s.Key == 2) // если стиль - Заголовок 1
                            ChangeStylePrps(wDoc.MainDocumentPart.Document.Descendants<Paragraph>().ElementAt(0), s.Value);
                        if (s.Key == 3) // если стиль - Заголовок 2
                            ChangeStylePrps(wDoc.MainDocumentPart.Document.Descendants<Paragraph>().ElementAt(1), s.Value);
                        if (s.Key == 4) // если стиль - Заголовок 3
                            ChangeStylePrps(wDoc.MainDocumentPart.Document.Descendants<Paragraph>().ElementAt(2), s.Value);
                    }
                }
                // изменение параметров раздела
                ChangeSection(desc, wDoc.MainDocumentPart);
                // сохранение изменений в документе
                wDoc.MainDocumentPart.Document.Save();
            }
        }

        // изменение параметров раздела
        private static void ChangeSection(DocDesc descr, MainDocumentPart docPart)
        {
            // изменение полей документа
            PageMargin pageMargin = docPart.Document.Body.Descendants<PageMargin>().First();
            if (descr.Margins["leftMargin"]!=-100)
                pageMargin.Left = Convert.ToUInt32(Math.Round(descr.Margins["leftMargin"] * 567));
            if (descr.Margins["rightMargin"] != -100)
                pageMargin.Right = Convert.ToUInt32(Math.Round(descr.Margins["rightMargin"] * 567));
            if (descr.Margins["topMargin"] != -100)
                pageMargin.Top = Convert.ToInt32(Math.Round(descr.Margins["topMargin"] * 567));
            if (descr.Margins["botMargin"] != -100)
                pageMargin.Bottom = Convert.ToInt32(Math.Round(descr.Margins["botMargin"] * 567));
            // изменение колонтитулов документа
            if (descr.ColTop != -100)
                pageMargin.Header = Convert.ToUInt32(Math.Round(descr.ColTop * 567));
            if (descr.ColBot != -100)
                pageMargin.Footer = Convert.ToUInt32(Math.Round(descr.ColBot * 567));

            // установка свойств раздела
            SectionProperties sectionProps = docPart.Document.Body.Descendants<SectionProperties>().First();
            // установка особого колонтитула первой страницы
            if (descr.SpecCol == true)
                sectionProps.Append(new TitlePage());

            // если пользователь выбрал положение нумерации внизу страницы
            if (descr.Place == "Bottom")
            {
                // вставляется нижний колонтитул
                FooterPart footerPart = docPart.AddNewPart<FooterPart>("rId6");
                GenerateFooterPartContent(footerPart, descr);
                sectionProps.Append(new FooterReference() { Type = HeaderFooterValues.Default, Id = "rId6" });
            }
            else
            {
                // иначе вставляется верхний колонтитул
                HeaderPart headerPart = docPart.AddNewPart<HeaderPart>("rId7");
                GenerateHeaderPartContent(headerPart, descr);
                sectionProps.Append(new FooterReference() { Type = HeaderFooterValues.First, Id = "rId7" });
            }
        }
        // изменение параметров стиля
        private static void ChangeStylePrps(Paragraph st, StyleDesc inputSt)
        {
            // добавление параметров текста
            RunProperties runProperties = new RunProperties();
            // название шрифта
            runProperties.Append(new RunFonts() { Ascii = inputSt.Font, HighAnsi = inputSt.Font, ComplexScript= inputSt.Font });
            if (inputSt.Bold == true) // полужирное начертание
                runProperties.Append(new Bold());
            if (inputSt.Cursive == true) // курсив
                runProperties.Append(new Italic());
            if (inputSt.Underlined == true) // подчеркнутый
                runProperties.Append(new Underline() { Val = UnderlineValues.Single });
            if (inputSt.FontSize != -100) // размер шрифта
            {
                if (inputSt.FontSize == 0) // если задано значение 0, изменить на 1
                    inputSt.FontSize = 1;
                runProperties.Append(new FontSize() { Val = (inputSt.FontSize * 2).ToString() });
            }

            runProperties.Append(new Color() { Val = inputSt.Color });
            st.Descendants<Run>().Last().RsidRunProperties = "0048781C";
            st.Descendants<Run>().Last().PrependChild<RunProperties>(runProperties);

            // свойства текста внутри свойст абзаца
            ParagraphMarkRunProperties parRunProps = new ParagraphMarkRunProperties();
            parRunProps.Append(new RunFonts() { Ascii = inputSt.Font, HighAnsi = inputSt.Font, ComplexScript = inputSt.Font });
            if (inputSt.Bold == true)
            {
                parRunProps.Append(new Bold());
            }
            if (inputSt.Cursive == true)
                parRunProps.Append(new Italic());
            if (inputSt.Underlined == true)
                parRunProps.Append(new Underline() { Val = UnderlineValues.Single });
            if (inputSt.FontSize != -100)
            {
                if (inputSt.FontSize == 0) // если задано значение 0, изменить на 1
                    inputSt.FontSize = 1;
                parRunProps.Append(new FontSize() { Val = (inputSt.FontSize * 2).ToString() });
            }
            // вставка цвета
            parRunProps.Append(new Color() { Val = inputSt.Color });
            ParagraphProperties parProperties = new ParagraphProperties();

            // определение парметров интервальных отступов
            SpacingBetweenLines sp = new SpacingBetweenLines();
            if (inputSt.BeforeSpace != -100)
                sp.Before = Math.Round(inputSt.BeforeSpace * 20).ToString();
            if (inputSt.AfterSpace != -100)
                sp.After = Math.Round(inputSt.AfterSpace * 20).ToString();
            if (inputSt.BtwSapce != -100)
                sp.Line = Math.Round(inputSt.BtwSapce * 240).ToString();
            parProperties.Append(sp);

            // опредеение параметров отступов
            Indentation ind = new Indentation();
            if (inputSt.LeftIndent != -100)
                ind.Left = Math.Round(inputSt.LeftIndent * 567).ToString();
            if (inputSt.RigthtIndent != -100)
                ind.Right = Math.Round(inputSt.RigthtIndent * 567).ToString();
            if (inputSt.FirstLineIndent != -100)
                ind.FirstLine = Math.Round(inputSt.FirstLineIndent * 567).ToString();
            parProperties.Append(ind);

            switch (inputSt.JustPar)
            {
                case "Both":
                    parProperties.Append(new Justification() { Val = JustificationValues.Both});
                    break;
                case "Left":
                    parProperties.Append(new Justification() { Val = JustificationValues.Left });
                    break;
                case "Right":
                    parProperties.Append(new Justification() { Val = JustificationValues.Right });
                    break;
                case "Center":
                    parProperties.Append(new Justification() { Val = JustificationValues.Center });
                    break;
            }

            parProperties.Append(parRunProps);
            st.PrependChild<ParagraphProperties>(parProperties);
        }

        // генерация содержимого нижнего колонтитула
        private static void GenerateFooterPartContent(FooterPart footerPart, DocDesc desc)
        {
            // объявление нового нижнего колонтитула
            Footer footer = new Footer() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" } };
            // добавление пространства имен
            footer.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footer.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            footer.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footer.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footer.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            footer.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            footer.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footer.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footer.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footer.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            SdtBlock sdtBlock = new SdtBlock();
            // определение набора свойств, которые будут применены к колонтитулу
            SdtProperties sdtProperties = new SdtProperties();
            SdtId sdtId = new SdtId() { Val = -228152962 };

            SdtContentDocPartObject sdtContentDocPartObject = new SdtContentDocPartObject();
            // определение параметров нумерации страниц
            DocPartGallery docPartGallery = new DocPartGallery() { Val = "Page Numbers (Bottom of Page)" };
            DocPartUnique docPartUnique = new DocPartUnique();

            sdtContentDocPartObject.Append(docPartGallery);
            sdtContentDocPartObject.Append(docPartUnique);

            sdtProperties.Append(sdtId);
            sdtProperties.Append(sdtContentDocPartObject);

            SdtContentBlock sdtContentBlock = new SdtContentBlock();
            // объявление нового абзаца в колонтитуле
            Paragraph paragraph = new Paragraph() { RsidParagraphAddition = "005A04A2", RsidRunAdditionDefault = "005A04A2" };

            ParagraphProperties paragraphProperties = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "a5" };
            // установка выравнивания нумерации
            switch (desc.NumJust)
            {
                case "Both": // по ширине
                    paragraphProperties.Append(new Justification() { Val = JustificationValues.Both });
                    break;
                case "Left": // слева
                    paragraphProperties.Append(new Justification() { Val = JustificationValues.Left });
                    break;
                case "Right": // справа
                    paragraphProperties.Append(new Justification() { Val = JustificationValues.Right });
                    break;
                case "Center": // по центру
                    paragraphProperties.Append(new Justification() { Val = JustificationValues.Center });
                    break;
            }
            Justification justification = new Justification() { Val = JustificationValues.Center };

            paragraphProperties.Append(paragraphStyleId1);

            Run run = new Run();
            // определение вычисляемого поля с номером страниц
            FieldChar fieldChar = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run.Append(fieldChar);

            Run run2 = new Run();
            FieldCode fieldCode1 = new FieldCode();
            // добавление формулы для нумерации
            fieldCode1.Text = "PAGE   \\* MERGEFORMAT";

            run2.Append(fieldCode1);

            Run run3 = new Run();
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run3.Append(fieldChar2);

            Run run4 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            // добавление начального значения нумерации
            runProperties1.Append(noProof1);
            Text text1 = new Text();
            text1.Text = "1";

            run4.Append(runProperties1);
            run4.Append(text1);

            Run run5 = new Run();
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run5.Append(fieldChar3);
            // добавление в абзац колонтитула свойств
            paragraph.Append(paragraphProperties);
            paragraph.Append(run);
            paragraph.Append(run2);
            paragraph.Append(run3);
            paragraph.Append(run4);
            paragraph.Append(run5);

            sdtContentBlock.Append(paragraph);

            sdtBlock.Append(sdtProperties);
            sdtBlock.Append(sdtContentBlock);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "005A04A2", RsidRunAdditionDefault = "005A04A2" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "a5" };

            paragraphProperties2.Append(paragraphStyleId2);

            paragraph2.Append(paragraphProperties2);

            footer.Append(sdtBlock);
            footer.Append(paragraph2);

            footerPart.Footer = footer;
        }

        // генерация содержимого верхнего колонтитула
        private static void GenerateHeaderPartContent(HeaderPart headerPart, DocDesc desc)
        {
            Header header = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" } };
            header.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            header.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            header.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            SdtBlock sdtBlock = new SdtBlock();
            // определение набора свойств, которые будут применены к колонтитулу
            SdtProperties sdtProperties = new SdtProperties();
            SdtId sdtId = new SdtId() { Val = -1743938990 };
            // определенеи положения номера 
            SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
            DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
            DocPartUnique docPartUnique1 = new DocPartUnique();

            sdtContentDocPartObject1.Append(docPartGallery1);
            sdtContentDocPartObject1.Append(docPartUnique1);

            sdtProperties.Append(sdtId);
            sdtProperties.Append(sdtContentDocPartObject1);

            SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00075615", RsidRunAdditionDefault = "00075615" };

            ParagraphProperties paragraphProperties = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "a5" };
            // установка выравнивания нумерации
            switch (desc.NumJust)
            {
                case "Both": // по ширине
                    paragraphProperties.Append(new Justification() { Val = JustificationValues.Both });
                    break;
                case "Left": // слева
                    paragraphProperties.Append(new Justification() { Val = JustificationValues.Left });
                    break;
                case "Right": // справа
                    paragraphProperties.Append(new Justification() { Val = JustificationValues.Right });
                    break;
                case "Center": // по центру
                    paragraphProperties.Append(new Justification() { Val = JustificationValues.Center });
                    break;
            }
            Justification justification = new Justification() { Val = JustificationValues.Center };

            paragraphProperties.Append(paragraphStyleId1);

            Run run1 = new Run();
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run1.Append(fieldChar1);

            Run run2 = new Run();
            // поле с формулой для нумерации
            FieldCode fieldCode1 = new FieldCode();
            fieldCode1.Text = "PAGE   \\* MERGEFORMAT";

            run2.Append(fieldCode1);

            Run run3 = new Run();
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run3.Append(fieldChar2);

            Run run4 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();
            // установка первого значения нумерации
            runProperties1.Append(noProof1);
            Text text1 = new Text();
            text1.Text = "1";

            run4.Append(runProperties1);
            run4.Append(text1);

            Run run5 = new Run();
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run5.Append(fieldChar3);
            // установка свойст абзаца
            paragraph1.Append(paragraphProperties);
            paragraph1.Append(run1);
            paragraph1.Append(run2);
            paragraph1.Append(run3);
            paragraph1.Append(run4);
            paragraph1.Append(run5);

            sdtContentBlock1.Append(paragraph1);

            sdtBlock.Append(sdtProperties);
            sdtBlock.Append(sdtContentBlock1);
            // добавление информации об абзаце
            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00075615", RsidRunAdditionDefault = "00075615" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "a5" };

            paragraphProperties2.Append(paragraphStyleId2);

            paragraph2.Append(paragraphProperties2);

            header.Append(sdtBlock);
            header.Append(paragraph2);
            // установка верхнего колонтитула
            headerPart.Header = header;
        }
    }
}
