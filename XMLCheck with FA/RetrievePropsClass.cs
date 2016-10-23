using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ComplianceAssessment
{
    // класс для извлечение параметров форматирования из документа
    public class RetrievePropsClass: BaseConcepts
    {
        private static Table table;
        private static SectionProperties section;
        private static IEnumerable<Paragraph> parsGet;
        /// <summary>
        /// Извечение параметров форматирования
        /// </summary>
        /// <param name="fileName">Полный путь к файлу документа</param>
        /// <returns></returns>
        public static string GetProperties(string fileName)
        {
            // открытие документа
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(fileName, true))
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
                FormattingAssembler.AssembleFormatting(wDoc, settings);
                // получение списка абзацев
                parsGet = wDoc.MainDocumentPart.Document.Descendants<Paragraph>();
                // полчение информации об используемых стилях
                styleIdList = parsGet.Select(st => ParagraphStyle(st)).ToList(); // ID стилей
                allStyleListToCompare = wDoc.MainDocumentPart.StyleDefinitionsPart.Styles.OfType<Style>().Select(pa => pa).ToDictionary(d => d.StyleId.Value);
                styleNamePlusIdList = new Dictionary<string, string>();

                foreach (string styleId in styleIdList)
                {
                    Style st = allStyleListToCompare[styleId];
                    if (!styleNamePlusIdList.ContainsKey(styleId))
                        styleNamePlusIdList.Add(styleId, st.StyleName.Val.Value); // список стилей + ID
                }

                if (wDoc.MainDocumentPart.Document.Body.Elements<Table>().Count() != 0)
                {
                    // если в документе есть таблицы, получаем  первую таблицу
                    table = wDoc.MainDocumentPart.Document.Body.Elements<Table>().First();
                }
                else table = null;
                // установка свойств раздела 
                section = wDoc.MainDocumentPart.Document.Body.Descendants<SectionProperties>().First();
                try
                {
                    // изображение из документа
                    parWithDrawingToCompare = wDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>().Where(pr => pr.Parent == wDoc.MainDocumentPart.Document.Body && pr.Descendants<Drawing>().Count() > 0).First();
                }
                catch { parWithDrawingToCompare = null; }
                // нумерация списоков
                try
                {
                    numberingToCompare = wDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                }
                catch { numberingToCompare = null; }
                // обработка колонтитулов
                Footer footer;
                Header header;
                if (wDoc.MainDocumentPart.FooterParts != null) // нижний колонтитул
                {
                    for (int i = 0; i < wDoc.MainDocumentPart.FooterParts.Count(); i++)
                    {
                        footer = wDoc.MainDocumentPart.FooterParts.ElementAt(i).Footer;
                        // положение нумерации
                        placeNum = (footer.ChildElements.ToList().Exists(s => s.LocalName == "sdt")) ? footer.Descendants<SdtContentDocPartObject>().First().DocPartGallery.Val.Value.ToString() : "";
                        // выравнивание номера
                        justToCompare = ParagraphDicts.justificationDict[footer.Descendants<ParagraphProperties>().First().Justification.Val.Value.ToString()];
                    }
                }

                if (wDoc.MainDocumentPart.HeaderParts != null) // верхний колонтитул
                {
                    for (int i = 0; i < wDoc.MainDocumentPart.HeaderParts.Count(); i++)
                    {
                        header = wDoc.MainDocumentPart.HeaderParts.ElementAt(i).Header;
                        placeNum = (header.ChildElements.ToList().Exists(s => s.LocalName == "sdt")) ? header.Descendants<SdtContentDocPartObject>().First().DocPartGallery.Val.Value.ToString() : "";
                        justToCompare = ParagraphDicts.justificationDict[header.Descendants<ParagraphProperties>().First().Justification.Val.Value.ToString()];
                    }
                }

                // извлечение параметров
                return GetInfoFromPars(parsGet);

            }
        }

        /// <summary>
        /// Извлечение параметров форматирования из абзацев
        /// </summary>
        /// <param name="pars">Все абзацы документа</param>
        /// <returns></returns>
        private static string GetInfoFromPars(IEnumerable<Paragraph> pars)
        {
            List<String> stProps = new List<string>();
            Dictionary<string, Level> lists;
            string styleName;
            string result = "";
            ParagraphProperties props;
            foreach (var existingStyle in styleNamePlusIdList) // прохождение по всем стилям документа
            {
                // получение свойств абзаца по его стилю
                props = GetParPropsFromPar(pars, existingStyle.Key);
                styleName = "<div class=\"text-info\">Стиль: " + existingStyle.Value + "</div>";
                result += styleName + ParPropsString(props) + FontPropsString(styleToCompareRunProps);
                // проверка на нумерованный/маркированный список
                if (existingStyle.Value == "List Paragraph")
                {
                    lists = GetNumbering(pars, existingStyle.Key); // получение нумерации
                    if (lists != null)
                    {
                        foreach (var el in lists)
                            result += "\n<div class=\"text-info\">Стиль: " + el.Key + "</div>\n" + GetNumInfo(el.Value);
                    }
                }
            }
            if (parWithDrawingToCompare != null)    // проверка на наличие изображения
                result += GetDrawing();
            result += "\n<div class=\"text-info\">Параметры страницы:</div> " + GetSectionProps();
            if (table != null)
                result += "\n<div class=\"text-info\">Параметры таблиц:</div> " + InsideTable() + GetTableMargins();
            return result;
        }
        // словарь со значениями вертикального выравнивания
        public static Dictionary<string, string> alignDict = new Dictionary<string, string>
        {
            ["Bottom"] = "снизу",
            ["Center"] = "по центру",
            ["Top"] = "сверху"
        };
        // извлечение данных из таблицы
        private static string InsideTable()
        {
            // получение количества колонок таблицы
            int cols = table.Descendants<TableGrid>().First().Descendants<GridColumn>().Count();
            // получение первой ячейки заголовка
            TableCell tCellFirst = table.Descendants<TableCell>().First();
            TableCellVerticalAlignment vAlignToCompareHead = tCellFirst.TableCellProperties.TableCellVerticalAlignment;
            // получение первой ячейки из основной части таблицы
            TableCell tCellSecond = table.Descendants<TableCell>().ElementAt(cols);
            TableCellVerticalAlignment vAlignToCompareMain = tCellSecond.TableCellProperties.TableCellVerticalAlignment;

            string head = "";
            if (vAlignToCompareHead != null)
                head = alignDict[vAlignToCompareHead.Val.Value.ToString()];
            head = "; для заголовка - " + head;
            if (vAlignToCompareMain != null)
                head += "; для остальных строк таблицы - " + alignDict[vAlignToCompareMain.Val.Value.ToString()];
            if (head != "")
                head = head.Remove(0, 2);
            head = (head != "") ? "<div>▪ вертикальное выравнивание текста в ячейке " + head + "</div>" : "";

            string headRow = "";
            TableRowProperties tableRowPropsToCompare = table.Descendants<TableRow>().First().TableRowProperties;
            if (tableRowPropsToCompare != null)
            {
                if (tableRowPropsToCompare.HasChildren)
                {
                    // параметры для строк заголовка
                    headRow += (tableRowPropsToCompare.ChildElements.ToList().Exists(s => s.LocalName == "tblHeader")) ? "; повторять как заголовок на каждой странице" : "";
                    headRow += (tableRowPropsToCompare.ChildElements.ToList().Exists(s => s.LocalName == "cantSplit")) ? "; разрешить перенос строк на следующую страницу" : "";
                }
            }
            if (headRow != "")
                headRow = headRow.Remove(0, 2);
            headRow = (headRow != "") ? "<div>▪ параметры строки заголовка: " + headRow + "</div>" : "";

            string bodyRow = "";
            TableRowProperties tableRowPropsToCompareBody = table.Descendants<TableRow>().Last().TableRowProperties;
            if (tableRowPropsToCompareBody != null)
            {
                if (tableRowPropsToCompareBody.HasChildren)
                {
                    // параметры для строк основной части таблицы
                    headRow += (tableRowPropsToCompareBody.ChildElements.ToList().Exists(s => s.LocalName == "tblHeader")) ? "; повторять как заголовок на каждой странице" : "";
                    headRow += (tableRowPropsToCompareBody.ChildElements.ToList().Exists(s => s.LocalName == "cantSplit")) ? "; разрешить перенос строк на следующую страницу" : "";
                }
            }
            if (bodyRow != "")
                bodyRow = bodyRow.Remove(0, 2);
            bodyRow = (bodyRow != "") ? "<div>▪ параметры строк: " + bodyRow + "</div>" : "";

            return head + headRow + bodyRow;
        }
        // извечление полей ячеек таблицы
        private static string GetTableMargins()
        {
            // получение свойств таблицы
            TableProperties dictTableProps = table.Descendants<TableProperties>().First();
            // получение данных о полях ячеей=к таблицы
            TableCellMarginDefault tableMarCells = (dictTableProps.ChildElements.ToList().Exists(s => s.LocalName == "tblCellMar")) ?
                dictTableProps.TableCellMarginDefault : null;
            if (tableMarCells != null)
            {
                string margin = "";
                // верхнее поле
                TopMargin topMargin = (tableMarCells.ChildElements.ToList().Exists(s => s.LocalName == "top")) ?
                    tableMarCells.TopMargin : null;
                if (topMargin != null)
                {
                    margin += (topMargin.Width != null) ?
                        "; верхнее поле " + Math.Round(Convert.ToDouble(topMargin.Width.Value) / 567, 2).ToString() + " см" : "";
                }
                // нижнее поле
                BottomMargin botMargin = (tableMarCells.ChildElements.ToList().Exists(s => s.LocalName == "bottom")) ?
                    tableMarCells.BottomMargin : null;
                if (botMargin != null)
                {
                    margin += (botMargin.Width != null) ?
                         "; нижнее поле " + Math.Round(Convert.ToDouble(botMargin.Width.Value) / 567, 2).ToString() + " см" : "";
                }
                // левое поле
                StartMargin leftMargin = (tableMarCells.ChildElements.ToList().Exists(s => s.LocalName == "start")) ?
                    tableMarCells.StartMargin : null;
                if (leftMargin != null)
                {
                    margin += (leftMargin.Width.Value != null) ?
                        "; левое поле " + Math.Round(Convert.ToDouble(leftMargin.Width.Value) / 567, 2).ToString() + " см" : "";
                }
                // правое поле
                EndMargin rightMargin = (tableMarCells.ChildElements.ToList().Exists(s => s.LocalName == "end")) ?
                    tableMarCells.EndMargin : null;
                if (rightMargin != null)
                {
                    margin += (rightMargin.Width != null) ?
                        "; правое поле " + Math.Round(Convert.ToDouble(rightMargin.Width.Value) / 567, 2).ToString() + " см" : "";
                }
                if (margin != "")
                    margin = margin.Remove(0, 2);
                return (margin != "") ? "<div>▪ поля ячеек: " + margin + "</div>" : "";
            }
            return "";
        }
        // получение свойств раздела
        private static string GetSectionProps()
        {
            string result = "";
            PageMargin pageMargin = section.Descendants<PageMargin>().First();
            // полчение полей страницы
            result += (pageMargin.Bottom != null) ? "нижнее поле до " + Math.Round(Convert.ToDouble(pageMargin.Bottom.Value) / 567, 2).ToString() + " см" : "";
            result += (pageMargin.Top != null) ? "; верхнее поле до " + Math.Round(Convert.ToDouble(pageMargin.Top.Value) / 567, 2).ToString() + " см" : "";
            result += (pageMargin.Left != null) ? "; левое поле до " + Math.Round(Convert.ToDouble(pageMargin.Left.Value) / 567, 2).ToString() + " см" : "";
            result += (pageMargin.Right != null) ? "; правое поле до " + Math.Round(Convert.ToDouble(pageMargin.Right.Value) / 567, 2).ToString() + " см" : "";
            result += (pageMargin.Header != null) ? "; верхний колонтитул " + Math.Round(Convert.ToDouble(pageMargin.Header.Value) / 567, 2).ToString() + " см" : "";
            result += (pageMargin.Footer != null) ? "; нижний колонтитул " + Math.Round(Convert.ToDouble(pageMargin.Footer.Value) / 567, 2).ToString() + " см" : "";

            if (result.StartsWith("; "))
                result = result.Remove(0, 2);
            result = (result != "") ? "<div>▪ поля страницы: " + result + "</div>" : "";

            string page = "";
            PageSize pageSize = section.Descendants<PageSize>().First();
            // получение размера страницы
            page += "ширина - " + Math.Round(Convert.ToDouble(pageSize.Width.Value) / 567, 2).ToString() + " см";
            page += "; высота - " + Math.Round(Convert.ToDouble(pageSize.Height.Value) / 567, 2).ToString() + " см";

            if (page.StartsWith("; "))
                page = page.Remove(0, 2);
            page = (page != "") ? "<div>▪ размер страницы: " + page + "</div>" : null;
            // отметка об особом колонтитуле первой страницы
            string title = "";
            TitlePage tPage = (section.ChildElements.ToList().Exists(s => s.LocalName == "titlePg")) ?
                section.Descendants<TitlePage>().First() : null;
            if (tPage != null)
                if (tPage.Val != null)
                    if (tPage.Val.Value == true)
                        title = "<div>▪ особый колонтитул первой страницы</div>";
            // формат нумерации страниц
            string format = "";
            PageNumberType pageNum = (section.ChildElements.ToList().Exists(s => s.LocalName == "pgNumType")) ?
                section.Descendants<PageNumberType>().First() : null;
            if (pageNum != null)
            {
                if (pageNum.Format != null)
                    format = "формат нумерации - " + ParagraphDicts.numFormatDict[pageNum.Format.Value.ToString()];
                if (pageNum.Start != null)
                    format = "; начать нумерацию с " + pageNum.Start.Value.ToString() + "-ого значения";
            }
            if (format.StartsWith("; "))
                format = format.Remove(0, 2);
            format = (format != "") ? format : "";
            // положение номера на странице
            string place = "";
            if (placeNum != null)
            {
                if (placeNum != "")
                    placeNum = (placeNum.Contains("Bottom")) ? "; номер внизу страницы" : "номер вверху страницы";
            }
            else placeNum = "";
            if (justToCompare != null)
                justToCompare = (justToCompare != "") ? "; выравнивание по " + justToCompare : "";
            else justToCompare = "";
            place = placeNum + justToCompare;
            if (format == "" && place!="")
                place = place.Remove(0, 2);
            place = (place != "") ? place : "";

            string plForm = "";
            if (format != "" || place != "")
                plForm = "<div>▪ нумерация страниц: " + format + place + "</div>";

            string resultingSect = result + page + title + plForm;
            return resultingSect;
        }
        // извлечение параметров изображения
        private static string GetDrawing()
        {
            string dr = "";
            // положение рисунка на странице
            if (parWithDrawingToCompare.Descendants<Drawing>().First().Inline != null)
                dr = "в тексте";
            if (parWithDrawingToCompare.Descendants<Drawing>().First().Anchor != null)
                dr = "обтекание текстом";
            if (dr != "")
                dr = "\n<div class=\"text-info\">Положение рисунка:</div> " + dr;
            return dr;
        }
        // получение параметров нумерации
        public static string GetNumInfo(Level lvl)
        {
            // получение данных о нумерации из уровня списка
            string level = "";
            string regex = @"%\d";
            level = "уровень списка " + lvl.LevelIndex.Value.ToString() + " (0 - начальное значение)";
            level += "; пункт списка имет вид " +
                (lvl.LevelText.Val.Value.Contains("%") ? Regex.Replace(lvl.LevelText.Val.Value, regex, "*")
                : lvl.LevelText.Val.Value);
            level += "; формат номера списка - " + ParagraphDicts.numFormatDict[lvl.NumberingFormat.Val.Value.ToString()];
            level += "; начать список  с " + lvl.StartNumberingValue.Val.ToString() + "-ого значения";
            return level;
        }
        // получение нумерации
        public static Dictionary<string, Level> GetNumbering(IEnumerable<Paragraph> pToCompare, string style)
        {
            Dictionary<string, Level> lvls = new Dictionary<string, Level>(); // уровни нумерации в документе
            // получение параметров нумерации
            List<ParagraphProperties> pPrComp = GetNumPropsFromPars(pToCompare, style);
            if (pPrComp == null) return null;
            List<NumberingProperties> numCompProps = new List<NumberingProperties>();
            // проверка есть ли у абзаца нумерация
            foreach (var props in pPrComp)
                if (props.ChildElements.ToList().Exists(s => s.LocalName == "numPr"))
                    numCompProps.Add(props.NumberingProperties);
            // получение параметров для маркированных списоков
            NumberingProperties np = numCompProps.Find(lvl => levelComp(lvl).NumberingFormat.Val == "bullet");
            Level lvlBullet;
            if (np != null)
            {
                lvlBullet = levelComp(np);
                lvls.Add("Маркированный список", lvlBullet);
            }
            np = numCompProps.First(lvl => levelComp(lvl).NumberingFormat.Val != "bullet");
            Level lvlNumList = levelComp(np); // получение парметров для нумерованных списков
            if (np!=null)
            {
                lvlNumList = levelComp(np);
                lvls.Add("Нумерованный список", lvlNumList);

            }
            return lvls;
        }
        // уровень списка
        public static Level levelComp(NumberingProperties numPrToCompare)
        {
            // получение абстр numId из документа для сравнения
            AbstractNumId abNumid = numberingToCompare.Descendants<NumberingInstance>().First(p => p.NumberID.Value == numPrToCompare.NumberingId.Val).AbstractNumId;
            // получение абстр num из докуента для сравнения
            AbstractNum abNum = numberingToCompare.Descendants<AbstractNum>().First(l => l.AbstractNumberId.Value == abNumid.Val);
            return abNum.Descendants<Level>().First(lev => lev.LevelIndex.Value == numPrToCompare.NumberingLevelReference.Val);
        }
        // получение параметров нумерации из абзацев
        public static List<ParagraphProperties> GetNumPropsFromPars(IEnumerable<Paragraph> paragraphs, string styleId)
        {
            // свойства абзацев
            List<ParagraphProperties> paragraph = new List<ParagraphProperties>();
            List<Paragraph> newTempList = paragraphs.Where(p => p.ParagraphProperties.ParagraphStyleId != null).ToList();
            // получение абзацев с нумерацией
            List<Paragraph> parsNum = newTempList.Where(p => p.ParagraphProperties.ParagraphStyleId.Val.Value.Contains(styleId)).ToList();
            if (parsNum == null)
                return null;
            foreach (var p in parsNum)
                paragraph.Add(p.ParagraphProperties); // сбор всех свойств в список
            return paragraph;
        }
        // свойства шрифта
        private static string FontPropsString(RunProperties rProps)
        {
            string size;
            string color = "";
            string name = "";
            string bold = "";
            string italic = "";
            string underline = "";
            // размер шрифта
            size = (rProps.FontSize != null) ?
                "размер шрифта " + (Convert.ToDouble(rProps.FontSize.Val.Value) / 2).ToString() + " пт" : "";
            if (rProps.Color != null)   // цвет шрифта
            {
                color = "; цвет шрифта - ";
                if (rProps.Color.Val.Value == "000000" || rProps.Color.Val.Value == "auto")
                    color += "черный";
                else color += rProps.Color.Val.Value;
            }
            if (rProps.Italic != null)  // курсив
            {
                if (rProps.Italic.Val != null)
                {
                    if (rProps.Italic.Val.Value == true)
                        italic = "; курсивное начертание";
                }
            }
            if (rProps.Bold != null)    // полужирный
            {
                if (rProps.Bold.Val != null)
                {
                    if (rProps.Bold.Val.Value == true)
                        italic = "; полужирное начертание";
                }
            }
            if (rProps.Underline != null)   // подчеркнутый
            {
                if (rProps.Bold.Val.Value == true)
                    underline = "; подчеркивание на " + FontDicts.linenDict[rProps.Underline.Val.Value.ToString()];
            }
            if (rProps.RunFonts != null)
            {
                name = (rProps.RunFonts.Ascii.Value != null) ? "; название шрифта - " + rProps.RunFonts.Ascii.Value : "";
            }
            // результирующая информация по шрифту
            string result = size + name + color + bold + italic + underline;
            if (result != null)
            {
                if (result.StartsWith("; "))
                    result = result.Replace("; ", "");
                result = "<div>▪ параметры шрифта: " + result + "</div>";
            }
            return result;
        }
        // извлечение параметров абзаца
        private static string ParPropsString(ParagraphProperties props)
        {
            string just;    // выравнивание
            string rIndent = ""; // отступ справа
            string lIndent = ""; // отступ слева
            string fLineIndent = ""; // отступ первой строки
            string afterSpace = ""; // интервал после абзаца
            string beforeSpace = ""; // интервал перед абзацем
            string line = ""; // междустрочный интервал
            string lineRule = ""; // правило междустрочного интервала

            // выравнивание текста
            just = (props.Justification != null) ?
                ParagraphDicts.justificationDict[props.Justification.Val.Value.ToString()] : "левому краю";
            if (props.Indentation != null) // отступы
            {
                rIndent = (props.Indentation.Right != null) ?
                    Math.Round(Convert.ToDouble(props.Indentation.Right.Value) / 567, 2).ToString() : "0";
                lIndent = (props.Indentation.Right != null) ?
                    Math.Round(Convert.ToDouble(props.Indentation.Right.Value) / 567, 2).ToString() : "0";
                fLineIndent = (props.Indentation.FirstLine != null) ?
                    Math.Round(Convert.ToDouble(props.Indentation.FirstLine.Value) / 567, 2).ToString() : "0";
            }
            if (props.SpacingBetweenLines != null) // интервалы
            {
                afterSpace = (props.SpacingBetweenLines.After != null) ?
                    Math.Round(Convert.ToDouble(props.SpacingBetweenLines.After.Value) / 20).ToString() : "0";
                beforeSpace = (props.SpacingBetweenLines.Before != null) ?
                    Math.Round(Convert.ToDouble(props.SpacingBetweenLines.Before.Value) / 20).ToString() : "0";
                line = (props.SpacingBetweenLines.Line != null) ?
                    Math.Round(Convert.ToDouble(props.SpacingBetweenLines.Line.Value) / 240, 2).ToString() +
                     " (" + Math.Round(Convert.ToDouble(props.SpacingBetweenLines.Line.Value) / 20).ToString() + " пт" + ")" : "0";
                lineRule = (props.SpacingBetweenLines.LineRule != null) ?
                    ParagraphDicts.lineRuleDict[props.SpacingBetweenLines.LineRule.Value.ToString()] : "";
            }

            lIndent = (lIndent != "") ? "левый отступ " + lIndent + " см" : "";
            rIndent = (rIndent != "") ? "; правый отступ " + rIndent + " см" : "";
            fLineIndent = (fLineIndent != "") ? "; отступ первой строки " + fLineIndent + " см" : "";
            string indentation = lIndent + rIndent + fLineIndent;
            if (indentation != "")
            {
                if (indentation.StartsWith("; "))
                    indentation = indentation.Replace("; ", "");
                indentation = "<div>▪ " + indentation + "</div>";
            }

            afterSpace = (afterSpace != "") ? "интервал после абзаца " + afterSpace + " пт" : "";
            beforeSpace = (beforeSpace != "") ? "; интервал перед абзацем " + beforeSpace + " пт" : "";
            line = (line != "") ? "; междустрочный интервал " + line : "";
            lineRule = (lineRule != "") ? "; междустрочный интервал - " + lineRule : "";
            string space = afterSpace + beforeSpace + line + lineRule;
            if (space != "")
            {
                if (space.StartsWith("; "))
                    space = space.Replace("; ", "");
                space = "<div>▪ " + space + "</div>";
            }
            return "<div>▪ выравнивание по " + just + "</div>" + indentation + space;
        }
    }
}
