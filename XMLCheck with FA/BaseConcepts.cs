using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ComplianceAssessment
{
    public class BaseConcepts
    {
        public static List<string> styleIdList { get; set; } // список ID стилей в проверяемом документе
        public static Dictionary<string, string> styleNamePlusIdList { get; set; } // словарь со значениями идент стиля и его названия в доке для сравнения
        public static Paragraph parWithDrawingToCompare { get; set; } // абзац с изображением в документе с правилами оформления
        public static Dictionary<string, Style> allStyleListToCompare { get; set; } // стили в шаблонном документе
        public static Dictionary<string, Style> allStyleList { get; set; } // все стили в проверяемом документе
        public static Numbering numberingToCompare { get; set; } // нумерованные и маркированные списки в шаблонном документе
        public static SectionProperties sectionPropsToCompare { get; set; } // свойства раздела

        public static List<List<object>> comentsPars = new List<List<object>>(); // комментарии на абзацы
        public static List<List<object>> comentsRuns = new List<List<object>>(); // комментарии на пробеги текста
        //************************//
        //          ДЛЯ           //
        //    НОМЕРОВ СТРАНИЦ     //
        //************************//

        public static string placeNum { get; set; } // положение номера на странице
        public static string justToCompare { get; set; } // выравнивание номера

        //************************//
        //          ДЛЯ           //
        //         ТАБЛИЦ         //
        //************************//
        // Список свойств таблицы для документа
        public static Dictionary<string, object> dictTableProps { get; set; }
        // Список свойств таблицы для документа для проверки
        public static Table tableToCompare { get; set; }
        //************************//
        //    ДЛЯ АБЗАЦЕВ         //
        //    И РАНОВ             //
        //************************//
        public static Dictionary<string, object> dictParProps { get; set; } // свойства абзаца
        public static Dictionary<string, object> dictToCompareParProps { get; set; } // свойства абзаца в шаблонном документе
        public static Dictionary<string, object> dictRunPrps { get; set; } // свойства пробега текста
        public static Dictionary<string, object> dictRunToComparePrps { get; set; } // свойства пробега текста в шаблонном документе
        public static RunProperties styleToCompareRunProps { get; set; }
        public static IEnumerable<Paragraph> pars { get; set; }

        /// <summary>
        /// Позволяет получить стиль из списка по StyleID
        /// </summary>
        /// <param name="style">StyleID в списке стилей</param>
        /// <param name="StyleList">Список стилей в документе</param>
        public static Style StyleFromStyleList(string styleID, Dictionary<string, Style> styleList)
        {
            Style style = new Style();
            if (styleList.TryGetValue(styleID, out style))   // поиск стиля в документе по styleID
            {
                return style;
            }
            return null;
        }

        public static string ParagraphStyle(Paragraph p)   // определение стиля абзаца
        {
            ParagraphProperties pPr = p.GetFirstChild<ParagraphProperties>();
            if (pPr != null)
            {
                ParagraphStyleId paraStyle = pPr.ParagraphStyleId;
                if (paraStyle != null)
                {
                    return paraStyle.Val.Value; // возвращает значение стиля
                }
                else return "a";    // возвращает стиль Normal
            }
            return "";
        }
        // получение параметров абзаца
        public static ParagraphProperties GetParPropsFromPar(IEnumerable<Paragraph> paragraphs, string styleId)
        {
            dictRunToComparePrps = new Dictionary<string, object>();
            var paragraph = new Paragraph();
            List<Paragraph> newTempList = new List<Paragraph>();
            // добавление во временный список абзацев, у которых есть отметка о стиле
            newTempList = paragraphs.Where(p => p.ParagraphProperties.ParagraphStyleId != null).ToList();     
            
            if (styleId == "a") // если стиль Обычный (так как у него нет отметки о стиле)
                paragraph = paragraphs.Where(p => p.ParagraphProperties.ParagraphStyleId == null).ToList().First();
            else
                paragraph = newTempList.Where(p => p.ParagraphProperties.ParagraphStyleId.Val.Value.Contains(styleId)).ToList().First();
            // свойства пробега текста для абзаца
            styleToCompareRunProps = paragraph.Descendants<Run>().ToList().First().RunProperties;
            dictRunToComparePrps = DictionaryFromType(styleToCompareRunProps);
            return paragraph.ParagraphProperties; // возвращение параметров абзаца
        }
        /// <summary>
        /// Позволяет получить словарь свойств для сравнения
        /// </summary>
        /// <param name="atype">Объект класса, содержащего свойства для сравнения</param>
        /// <returns></returns>
        public static Dictionary<string, object> DictionaryFromType(object atype)
        {
            if (atype == null) return new Dictionary<string, object>();
            Type t = atype.GetType();
            PropertyInfo[] props = t.GetProperties();
            Dictionary<string, object> dict = new Dictionary<string, object>();
            foreach (PropertyInfo prp in props)
            {
                if (prp.CanRead)
                {
                    object value = prp.GetValue(atype, new object[] { });
                    dict.Add(prp.Name, value);
                }
            }
            return dict;
        }
        /// <summary>
        /// Действия над документом с правилами оформления
        /// </summary>
        /// <param name="fileNameCompare"></param>
        public static void DocForComparison(string fileNameCompare)
        {
            // открытие документа с правилами оформления
            using (WordprocessingDocument docForComparison = WordprocessingDocument.Open(fileNameCompare, true))
            {
                FormattingAssemblerSettings settings = new FormattingAssemblerSettings()
                {
                    ClearStyles = false,
                    RemoveStyleNamesFromParagraphAndRunProperties = false,
                    CreateHtmlConverterAnnotationAttributes = false,
                    OrderElementsPerStandard = true,
                    RestrictToSupportedLanguages = true,
                    RestrictToSupportedNumberingFormats = false,
                };
                FormattingAssembler.AssembleFormatting(docForComparison, settings);
                // абзацы из документа
                pars = docForComparison.MainDocumentPart.Document.Descendants<Paragraph>();
                styleIdList = pars.Select(st => ParagraphStyle(st)).ToList(); // список ID стилей
                allStyleListToCompare = docForComparison.MainDocumentPart.StyleDefinitionsPart.Styles.OfType<Style>().Select(pa => pa).ToDictionary(d => d.StyleId.Value);
                styleNamePlusIdList = new Dictionary<string, string>();
                foreach (string styleId in styleIdList)
                {
                    Style st = allStyleListToCompare[styleId];
                    if(!styleNamePlusIdList.ContainsKey(styleId))
                        styleNamePlusIdList.Add(styleId, st.StyleName.Val.Value); // список стилей + ID
                }

                if (docForComparison.MainDocumentPart.Document.Body.Elements<Table>().Count() != 0)
                {
                    // если есть таблица, записать в переменную
                    tableToCompare = docForComparison.MainDocumentPart.Document.Body.Elements<Table>().First();
                }
                else tableToCompare = null;
                // установка свойств раздела 
                sectionPropsToCompare = docForComparison.MainDocumentPart.Document.Body.Descendants<SectionProperties>().First();
                try
                {
                    // если есть изображение, записать в переменную
                    parWithDrawingToCompare = docForComparison.MainDocumentPart.Document.Body.Descendants<Paragraph>().Where(pr => pr.Parent == docForComparison.MainDocumentPart.Document.Body && pr.Descendants<Drawing>().Count() > 0).First();
                }
                catch
                { parWithDrawingToCompare = null; }
                try
                {
                    // нумерация списков
                    numberingToCompare = docForComparison.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                }
                catch { numberingToCompare = null; }
                // обработка колонтитулов
                Footer footer;
                Header header;
                if (docForComparison.MainDocumentPart.FooterParts != null) // если есть нижние колонтитулы
                {
                    for (int i=0; i< docForComparison.MainDocumentPart.FooterParts.Count(); i++)
                    {
                        // определение нижнего колонтитула
                        footer = docForComparison.MainDocumentPart.FooterParts.ElementAt(i).Footer;
                        // положение номера на странице
                        placeNum = (footer.ChildElements.ToList().Exists(s => s.LocalName == "sdt")) ? 
                            footer.Descendants<SdtContentDocPartObject>().First().DocPartGallery.Val.Value.ToString() : "";
                        // выравнивание номера
                        justToCompare = ParagraphDicts.justificationDict[footer.Descendants<ParagraphProperties>().
                            First().Justification.Val.Value.ToString()];
                    }
                }
                
                if (docForComparison.MainDocumentPart.HeaderParts != null) // если есть верхние колонтитулы
                {
                    for (int i = 0; i < docForComparison.MainDocumentPart.HeaderParts.Count(); i++)
                    {
                        // определение верхнего колонтитула
                        header = docForComparison.MainDocumentPart.HeaderParts.ElementAt(i).Header;
                        // положение номера
                        placeNum = (header.ChildElements.ToList().Exists(s => s.LocalName == "sdt")) ? 
                            header.Descendants<SdtContentDocPartObject>().First().DocPartGallery.Val.Value.ToString() : "";
                        justToCompare = ParagraphDicts.justificationDict[header.Descendants<ParagraphProperties>().First().Justification.Val.Value.ToString()];
                    }
                }
            }
        }
        /// <summary>
        /// Добавить комментарий на абзац текста
        /// </summary>
        /// <param name="docPart">часть текста</param>
        /// <param name="comPar">абзац, к которому привязывается комментарий</param>
        /// <param name="comPars">абзацы комментария</param>
        public void AddCommentOnParagraph(MainDocumentPart docPart, List<Paragraph> comPar, List<Paragraph> comPars)
        {
            string id;
            Comment(docPart, comPars, out id);
            comPar.First().InsertBefore(new CommentRangeStart() { Id = id }, comPar.First().GetFirstChild<Run>());
            // Вставить CommentRangeEnd после последнего рана в абзаце
            var cmtEnd = comPar.Last().InsertAfter(new CommentRangeEnd() { Id = id }, comPar.Last().Elements<Run>().Last());
            // Создать ран с CommentReference и вставить.
            comPar.Last().InsertAfter(new Run(new CommentReference() { Id = id }), cmtEnd);
        }
        /// <summary>
        /// Добавить комментарий на пробег текста
        /// </summary>
        /// <param name="docPart">часть текста</param>
        /// <param name="run">пробег текста, к которому привязывается комментарий</param>
        /// <param name="comPars">абзацы комментария</param>
        public void AddCommentOnRun(MainDocumentPart docPart, List<Run> runs, List<Paragraph> comPars)
        {
            string id;
            Comment(docPart, comPars, out id);
            // Текст для комментария 
            // Вставить новый CommentRangeStart перед первым пробегом в абзаце
            runs.First().InsertBefore(new CommentRangeStart() { Id = id }, runs.First().GetFirstChild<Text>());
            //// Вставить CommentRangeEnd после последнего рана в абзаце
            var cmtEnd = runs.Last().InsertAfter(new CommentRangeEnd() { Id = id }, runs.Last().Elements<Text>().Last());
            //// Создать ран с CommentReference и вставить.
            runs.Last().InsertAfter(new Run(new CommentReference() { Id = id }), cmtEnd);
        }
        // комментарий
        private void Comment(MainDocumentPart docPart, List<Paragraph> comPars, out string id)
        {
            Comments comments = null;
            id = "0";
            if (docPart.GetPartsCountOfType<WordprocessingCommentsPart>() > 0)
            {
                comments = docPart.WordprocessingCommentsPart.Comments;
                if (comments.HasChildren == true)
                {
                    id = (comments.Descendants<Comment>().Count().ToString());
                }
            }
            else
            {
                WordprocessingCommentsPart commentPart = docPart.AddNewPart<WordprocessingCommentsPart>();
                commentPart.Comments = new Comments();
                comments = commentPart.Comments;
            }
            

            // Создать комментарий и вставить в документ
            for (int i = 0; i < comPars.Count(); i++)
            {
                comPars.ElementAt(i).ParagraphProperties = new ParagraphProperties(
                                        new NumberingProperties(
                                            new NumberingLevelReference() { Val = 0 },
                                                    new NumberingId() { Val = 114 }));
            }

            Comment cmt = new Comment()
            {
                Id = id.ToString(),
                Author = "ComplianceAssessment System",
                Date = DateTime.Now
            };

            cmt.Append(comPars);
            comments.AppendChild(cmt);
            comments.Save();
        }
    }
}
