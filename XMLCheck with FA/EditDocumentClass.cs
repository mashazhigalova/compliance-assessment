using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComplianceAssessment
{
    // класс для редактирования документа в соответствии с выбранными правилами
    class EditDocumentClass : BaseConcepts
    {
        // список стилей, используемых в шаблонном документе
        public static List<string> styleNameListToCompare { get; set; }

        string styleToCompare; // стиль в шаблонном документе

        // свойство вертикального выравнивания в ячейках таблицы
        TableCellVerticalAlignment vAlignToCompare { get; set; }
        // свойства ячеек таблицы
        TableCellProperties tCellProps { get; set; }
        /// <summary>
        /// Редактирование документа
        /// </summary>
        /// <param name="docPart"></param>
        /// <param name="fileNameCompare">Документ с правилами оформления</param>
        public void Edit(MainDocumentPart docPart, string fileNameCompare)
        {
            // таблица из шаблонного документа
            Table tableToCompare = new Table();
            // свойства раздела из шаблонного документа
            SectionProperties sectionPropsToCompare;
            // список абзацев
            IEnumerable<Paragraph> parsToCompare;
            // абзацы из исходного документа, которые надо изменить
            IEnumerable<Paragraph> p = docPart.Document.Body.Descendants<Paragraph>().Where(pr => pr.Parent == docPart.Document.Body && pr.InnerText != "");
            // открытие шаблонного документа с правилами оформления
            using (WordprocessingDocument docForComparison = WordprocessingDocument.Open(fileNameCompare, true))
            {
                // вывод всех параметров форматирования на самый верхний уровень
                FormattingAssemblerSettings settings = new FormattingAssemblerSettings()
                {
                    OrderElementsPerStandard = true,
                    RestrictToSupportedLanguages = true,
                    RestrictToSupportedNumberingFormats = false,
                    ClearStyles = false,
                    RemoveStyleNamesFromParagraphAndRunProperties = false,
                    CreateHtmlConverterAnnotationAttributes = false
                };
                FormattingAssembler.AssembleFormatting(docForComparison, settings);

                // абзацы из шаблонного документа
                parsToCompare = docForComparison.MainDocumentPart.Document.Descendants<Paragraph>();
                // список ID стилей из шаблонного документа
                styleIdList = parsToCompare.Select(st => ParagraphStyle(st)).ToList();
                // составление списка из стилей, используемых только в тексте документа (не из таблицы стилей)
                allStyleListToCompare = docForComparison.MainDocumentPart.StyleDefinitionsPart.Styles.OfType<Style>().Select(pa => pa).ToDictionary(d => d.StyleId.Value);
                styleNamePlusIdList = new Dictionary<string, string>();
                // список название стиля + стиль ID
                foreach (string styleId in styleIdList)
                {
                    Style st = allStyleListToCompare[styleId];
                    if (!styleNamePlusIdList.ContainsKey(styleId))
                        styleNamePlusIdList.Add(styleId, st.StyleName.Val.Value);
                }

                if (docForComparison.MainDocumentPart.Document.Body.Elements<Table>().Count() != 0)
                {
                    // табблица из шаблонного документа
                    tableToCompare = docForComparison.MainDocumentPart.Document.Body.Elements<Table>().First();
                }
                try
                {
                    parWithDrawingToCompare = docForComparison.MainDocumentPart.Document.Body.Descendants<Paragraph>().Where(pr => pr.Parent == docForComparison.MainDocumentPart.Document.Body && pr.Descendants<Drawing>().Count() > 0).First();
                    ChangeDrawing(parsToCompare, docPart);
                }
                catch { }
                try
                {
                    // установка свойств раздела 
                    sectionPropsToCompare = docForComparison.MainDocumentPart.Document.Body.Descendants<SectionProperties>().First();
                    docPart.Document.Body.ReplaceChild<SectionProperties>((SectionProperties)sectionPropsToCompare.CloneNode(true), docPart.Document.Body.Descendants<SectionProperties>().First());
                    // удаление ссылок на колонтитулы в доке
                    docPart.Document.Body.Descendants<SectionProperties>().First().RemoveAllChildren<HeaderReference>();
                    docPart.Document.Body.Descendants<SectionProperties>().First().RemoveAllChildren<FooterReference>();

                    Footer footer = null;
                    Header header = null;
                    if (docForComparison.MainDocumentPart.FooterParts != null)
                        footer = (docForComparison.MainDocumentPart.FooterParts.Any(f => f.Footer.ChildElements.ToList().Exists(s => s.LocalName == "sdt"))) ?
                            docForComparison.MainDocumentPart.FooterParts.First(f => f.Footer.ChildElements.ToList().Exists(s => s.LocalName == "sdt")).Footer : null;
                    if (docForComparison.MainDocumentPart.FooterParts != null)
                        header = (docForComparison.MainDocumentPart.HeaderParts.Any(h => h.Header.ChildElements.ToList().Exists(s => s.LocalName == "sdt"))) ?
                              docForComparison.MainDocumentPart.HeaderParts.First(h => h.Header.ChildElements.ToList().Exists(s => s.LocalName == "sdt")).Header : null;


                    ChangeHeaderFooter(docPart, footer, header);
                    // исправление свойств раздела
                }
                catch { }
            }

            foreach (Paragraph par in p)
            {
                ChangeStyle(parsToCompare, docPart, par); // изменение стилей
            }
            if(tableToCompare!=null)
                CheckTableProps(docPart, tableToCompare);
        }

        public static void ChangeHeaderFooter(MainDocumentPart mainDocumentPart, Footer fToCompare, Header hToCompare)
        {
            // Get SectionProperties and Replace HeaderReference and FooterRefernce with new Id
            IEnumerable<SectionProperties> sections = mainDocumentPart.Document.Body.Elements<SectionProperties>();
            if(hToCompare!=null)
            {
                mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);
                HeaderPart headerPart = mainDocumentPart.AddNewPart<HeaderPart>();
                string headerPartId = mainDocumentPart.GetIdOfPart(headerPart);
                headerPart.Header = (Header)hToCompare.CloneNode(true);

                foreach (var section in sections)
                {
                    // Delete existing references to headers and footers
                    section.RemoveAllChildren<HeaderReference>();
                    // Create the new header and footer reference node
                    section.PrependChild<HeaderReference>(new HeaderReference() { Id = headerPartId });
                }
            }

            if (fToCompare != null)
            {
                mainDocumentPart.DeleteParts(mainDocumentPart.FooterParts);
                FooterPart footerPart = mainDocumentPart.AddNewPart<FooterPart>();
                string footerPartId = mainDocumentPart.GetIdOfPart(footerPart);
                footerPart.Footer = (Footer)fToCompare.CloneNode(true);
                foreach (var section in sections)
                {
                    // Delete existing references to headers and footers
                    section.RemoveAllChildren<FooterReference>();

                    // Create the new header and footer reference node
                    section.PrependChild<FooterReference>(new FooterReference() { Id = footerPartId });
                }
            }
        }
        
        // изменение параметров изображения
        private void ChangeDrawing(IEnumerable<Paragraph> pToCompare, MainDocumentPart mainDocumentPart)
        {
            IEnumerable<Paragraph> parsWithDrawing = 
                mainDocumentPart.Document.Body.Descendants<Paragraph>().Where(pr => pr.Parent == mainDocumentPart.Document.Body && pr.Descendants<Drawing>().Count() > 0);

            foreach (Paragraph p in parsWithDrawing)
            {
                p.ParagraphProperties = (ParagraphProperties)parWithDrawingToCompare.ParagraphProperties.CloneNode(true);
            }
        }
        // изменения стиля - параметров шрифтов и абзацев
        public void ChangeStyle(IEnumerable<Paragraph> pToCompare, MainDocumentPart docPart, Paragraph par)
        {
            styleToCompare = ParagraphStyle(par);  // определение стиля абзаца - получение его StyleId
            if (styleToCompare == "") return;
            Dictionary<string, string> idName = new Dictionary<string, string> {[styleToCompare] = allStyleList[styleToCompare].StyleName.Val.Value };
            if (styleNamePlusIdList.ContainsValue(idName[styleToCompare]))
                styleToCompare = styleNamePlusIdList.FirstOrDefault(x => x.Value == idName[styleToCompare]).Key;
            if (styleNamePlusIdList.ContainsKey(styleToCompare))
            {
                try
                {
                    // определение свойств нумерованных списков
                    if (par.ParagraphProperties.ChildElements.ToList().Exists(s => s.LocalName == "numPr"))
                        par.ParagraphProperties.NumberingProperties = (NumberingProperties)ChangeNumProps(pToCompare, par, styleToCompare, docPart).CloneNode(true);
                    // копирование свойств абзаца из шаблонного в проверяемый документ
                    par.ParagraphProperties = (ParagraphProperties)GetParPropsFromPar(pToCompare, styleToCompare).CloneNode(true);
                    foreach (var run in par.Elements<Run>())
                    {
                        // копирование свойств текста из шаблонного в проверяемый докуемент
                        run.RunProperties = (RunProperties)styleToCompareRunProps.CloneNode(true);
                    }
                }
                catch { }
            }
        }
        // изменение параметров нумерованных/маркированных списков
        public NumberingProperties ChangeNumProps(IEnumerable<Paragraph> pToCompare, Paragraph p, string style, MainDocumentPart docPart)
        {
            // параметры нумерации из исходного документа
            NumberingProperties numPr = p.ParagraphProperties.NumberingProperties;
            Numbering numbering = docPart.NumberingDefinitionsPart.Numbering;
            NumberingProperties numPrToCompare = null;
            ParagraphProperties pPr = p.ParagraphProperties;
            List<ParagraphProperties> pPrComp = NumberingCheck.GetNumPropsFromPars(pToCompare, style);
            List<NumberingProperties> numCompProps = new List<NumberingProperties>();
            try
            {               
                // проверка есть ли у абзаца нумерация
                foreach (var props in pPrComp)
                    if (props.ChildElements.ToList().Exists(s => s.LocalName == "numPr"))
                        numCompProps.Add(props.NumberingProperties);

                // если свойств нумерации нет, то убрать список
                if (numCompProps.Count == 0)
                {
                    return null;
                }

                // получение абстр numId из оригинального документа
                AbstractNumId abNumidOriginal = numbering.Descendants<NumberingInstance>().First(po => po.NumberID.Value == numPr.NumberingId.Val).AbstractNumId;
                // получение абстр num из докуента для сравнения
                AbstractNum abNumOriginal = numbering.Descendants<AbstractNum>().First(l => l.AbstractNumberId.Value == abNumidOriginal.Val);
                Level lvlOriginal = abNumOriginal.Descendants<Level>().First(lev => lev.LevelIndex.Value == numPr.NumberingLevelReference.Val);

                if (lvlOriginal.NumberingFormat.Val == "bullet") // проверка маркированных списков
                {
                    foreach (var n in numCompProps)
                        if (NumberingCheck.levelComp(n).NumberingFormat.Val == "bullet")
                            numPrToCompare = n;
                }
                else
                    numPrToCompare = numCompProps.First(lvl => NumberingCheck.levelComp(lvl).NumberingFormat.Val != "bullet");
            }
            catch { }
            return numPrToCompare;
        }
        // исправление таблиц
        public void CheckTableProps(MainDocumentPart docPart, Table tableToCompare)
        {
            // получение всех таблиц из проверяемого документа
            IEnumerable<Table> tables = docPart.Document.Body.Elements<Table>();
            IEnumerable<Paragraph> parsTable;
            foreach (var t in tables)
            {
                // копирование свойств таблицы
                TableProperties tProps = (TableProperties)tableToCompare.Descendants<TableProperties>().First().CloneNode(true);
                // удаление старых свойтств таблицы
                t.RemoveAllChildren<TableProperties>();
                t.PrependChild(tProps); // добавление новых свойств
                if (t.Descendants<TableRow>().Count() > 1)
                {
                    // изменение строки заголовка таблицы
                    TableParsToCompare(1, out parsTable, tableToCompare);
                    if (tableToCompare.Descendants<TableRow>().First().ChildElements.ToList().Exists(s => s.LocalName == "trPr"))
                        t.Descendants<TableRow>().First().TableRowProperties = (TableRowProperties)tableToCompare.Descendants<TableRow>().First().TableRowProperties.CloneNode(true);
                    foreach (var cell in t.Descendants<TableRow>().First().Descendants<TableCell>())
                    {
                        if (vAlignToCompare!=null)
                            cell.TableCellProperties.TableCellVerticalAlignment = (TableCellVerticalAlignment)vAlignToCompare.CloneNode(true); // сиправление вертикального выравниваяни текста
                        cell.TableCellProperties = (TableCellProperties)tCellProps.CloneNode(true);
                        foreach (Paragraph p in cell.Descendants<Paragraph>())
                        {
                            // изменение параметров стиля для заголовка
                            ChangeStyle(parsTable, docPart, p);
                        }
                    }
                    // изменение основной части таблицы
                    TableParsToCompare(2, out parsTable, tableToCompare);
                    for (int i = 1; i < t.Descendants<TableRow>().Count(); i++)
                    {
                        // изменение свойств строк
                        if (tableToCompare.Descendants<TableRow>().Last().ChildElements.ToList().Exists(s => s.LocalName == "trPr"))
                            t.Descendants<TableRow>().ElementAt(i).TableRowProperties = (TableRowProperties)tableToCompare.Descendants<TableRow>().Last().TableRowProperties.CloneNode(true);
                        foreach (var cell in t.Descendants<TableRow>().ElementAt(i).Descendants<TableCell>())
                        {
                            // изменение свойств ячеек
                            cell.TableCellProperties = (TableCellProperties)tCellProps.CloneNode(true);
                            if (vAlignToCompare != null)
                                cell.TableCellProperties.TableCellVerticalAlignment = (TableCellVerticalAlignment)vAlignToCompare.CloneNode(true); // сиправление вертикального выравниваяни текста
                            foreach (Paragraph p in cell.Descendants<Paragraph>())
                                ChangeStyle(parsTable, docPart, p);
                        }
                    }
                }

            }
        }
        // изменение параметров таблицы
        private void TableParsToCompare(int what_part, out IEnumerable<Paragraph> parsInTable, Table tableToCompare)
        {
            parsInTable = null;
            // количество колонок
            int cols = tableToCompare.Descendants<TableGrid>().First().Descendants<GridColumn>().Count();
            switch (what_part)
            {
                case 1:
                    // получение первой ячейки в заголовке
                    TableCell tCellFirst = tableToCompare.Descendants<TableCell>().First();
                    parsInTable = tCellFirst.Descendants<Paragraph>(); // извлечене абзацев
                    vAlignToCompare = tCellFirst.TableCellProperties.TableCellVerticalAlignment; // вертикальное выравнивание
                    tCellProps = tCellFirst.TableCellProperties; // свойства ячейки
                    break;
                case 2:
                    // получение первой ячейки в основной части таблицы
                    TableCell tCellSecond = tableToCompare.Descendants<TableCell>().ElementAt(cols);
                    parsInTable = tCellSecond.Descendants<Paragraph>(); // извлечение абзацев
                    vAlignToCompare = tCellSecond.TableCellProperties.TableCellVerticalAlignment;
                    tCellProps = tCellSecond.TableCellProperties;
                    break;
            }
            // получение текстовых стилей из шаблонной таблицы 
            styleNameListToCompare = parsInTable.Select(st => ParagraphStyle(st)).ToList();
            styleNamePlusIdList = new Dictionary<string, string>();

            foreach (string styleId in styleNameListToCompare)
            {
                Style st = allStyleListToCompare[styleId];
                if (!styleNamePlusIdList.ContainsKey(styleId))
                    styleNamePlusIdList.Add(styleId, st.StyleName.Val.Value);
            }
        }
    }
}
