using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ComplianceAssessment
{
    class TableCheck : BaseConcepts
    {
        // словарь с выравниванием
        public static Dictionary<string, string> alignDict = new Dictionary<string, string>
        {
            ["Bottom"] = "снизу",
            ["Center"] = "по центру",
            ["Top"] = "сверху"
        };
        IEnumerable<Paragraph> paras { get; set; } // абзацы в таблице
        public static TableBorders tableBorders { get; set; }
        // Список свойств таблицы для документа для проверки
        public static TableBorders tableBordersToCompare { get; set; } // границы в таблице
        public static TableCellMarginDefault tableMarCells { get; set; } // поля в ячейках таблице
        public static TableCellMarginDefault tableMarCellsToCompare { get; set; } // поля в таблице из шаблонного документа
        TableCellVerticalAlignment vAlignToCompare { get; set; } // вертикальное выравнивание в шаблонном документе
        // проверка названия таблицы
        private void CheckTableName(Paragraph p, MainDocumentPart docPart)
        {
            if (p.InnerText == "") return; // если абзац не содержит текст
            string regName = "((T|t)able)|((Т|т)абл)";
            List<Paragraph> comPars = new List<Paragraph>();
            string comment = "";
            if (Regex.Match(p.InnerText, regName).Value != "")
            { if (p.InnerText.Length > 200) // если больше 200 символов в таблице
                    comment += "слишком длинное название таблицы"; }
            else comment += "возможно, название таблицы отсутствует (для таблицы далее по тексту)";
            if (comment!="")
                AddCommentOnParagraph(docPart, new List<Paragraph> { p }, new List<Paragraph>
                { new Paragraph(new Run(new Text(comment)))});
        }
        // проверка параметров таблицы
        public void CheckTableProps(MainDocumentPart docPart)
        {
            IEnumerable<Table> tables = docPart.Document.Body.Elements<Table>();
            ParRunCheck pr_check = new ParRunCheck();
            foreach (var t in tables)
            {
                // индекс таблицы в списке элементов документа
                int indTable = docPart.Document.Body.ChildElements.ToList().IndexOf(t);
                if (indTable != 0) // если таблица не первая в списке элементов
                    if (docPart.Document.Body.ChildElements.ToList().ElementAt(indTable - 1).LocalName == "p") // проверка на наличие абзаца над таблицей
                        CheckTableName((Paragraph)docPart.Document.Body.ChildElements.ToList().ElementAt(indTable - 1), docPart);
                TableProperties(t, docPart);
                if (t.Descendants<TableRow>().Count() > 1)
                {
                    paras = TableParsToCompare(1); // проверка строки заголовка
                    RowProps(t.Descendants<TableRow>().First(), tableToCompare.Descendants<TableRow>().First(), docPart); // переделать
                    foreach (var cell in t.Descendants<TableRow>().First().Descendants<TableCell>())
                    {
                        if(cell.InnerText!= "") VerticalAlign(cell, docPart); // проверка вертикального выравнивания текста
                        foreach (Paragraph p in cell.Descendants<Paragraph>().Where(i => i.InnerText != ""))
                        {
                            pr_check.Check(paras, docPart, p); // проверка абзацев и пробегов текста в заголовке
                        }
                    }
                    paras = TableParsToCompare(2); // проверка основного содержимого таблицы
                    for (int i = 1; i < t.Descendants<TableRow>().Count(); i++)
                    { 
                        // проверка свойств строк
                        RowProps(t.Descendants<TableRow>().ElementAt(i), tableToCompare.Descendants<TableRow>().Last(), docPart);
                        foreach (var cell in t.Descendants<TableRow>().ElementAt(i).Descendants<TableCell>())
                        {
                            if (cell.InnerText != "") VerticalAlign(cell, docPart);
                            foreach (Paragraph p in cell.Descendants<Paragraph>().Where(iq => iq.InnerText != ""))
                                pr_check.Check(paras, docPart, p);
                        }
                    }
                }
            }
        }
        // вертикальное выравнивание в ячейке
        private void VerticalAlign(TableCell tCell, MainDocumentPart docPart)
        {
            string com = "";
            List<Paragraph> comPars = new List<Paragraph>();
            TableCellVerticalAlignment vAlign = tCell.TableCellProperties.TableCellVerticalAlignment;
            if (vAlign == null && vAlignToCompare != null)
                com = alignDict[vAlignToCompare.Val.Value.ToString()];
            if (vAlign != null && vAlignToCompare != null)
                if (vAlign.Val.Value != vAlignToCompare.Val.Value)
                    com = alignDict[vAlignToCompare.Val.Value.ToString()];
            comPars.Add((com != "") ? new Paragraph(new Run(new Text("изменить вертикальное выравнивание в ячейке " + com))) : null);
            comPars.RemoveAll(s => s == null);
            if (comPars.Count != 0)
                AddCommentOnParagraph(docPart, new List<Paragraph> { tCell.Descendants<Paragraph>().First() }, comPars);
        }
        // параметры таблиц
        private void TableProperties(Table t, MainDocumentPart docPart)  // свойства таблицы в проверяемом документе
        {
            TableProperties dictTablePropsToCompare = tableToCompare.Descendants<TableProperties>().First();
            List<Paragraph> comPars = new List<Paragraph>();
            TableProperties tPr = t.Descendants<TableProperties>().First();
            tableBorders = tPr.TableBorders;    // получение всех параметров таблицы
            tableMarCells = tPr.TableCellMarginDefault;

            tableBordersToCompare = (dictTablePropsToCompare.ChildElements.ToList().Exists(s => s.LocalName == "tblBorders")) ? 
                dictTablePropsToCompare.TableBorders : null;
            tableMarCellsToCompare = (dictTablePropsToCompare.ChildElements.ToList().Exists(s => s.LocalName == "tblCellMar")) ?
                dictTablePropsToCompare.TableCellMarginDefault : null;

            TablePropertiesCheck tc = new TablePropertiesCheck();
            if(tableBorders!=null && tableBordersToCompare!=null)
                comPars.Add(tc.BordersCheck());
            if (tableMarCells!=null && tableMarCellsToCompare!=null)
                comPars.Add(tc.MarginsCheck());
            comPars.RemoveAll(s => s == null);
            if (comPars.Count != 0)
                AddCommentOnRun(docPart, t.Descendants<Run>().ToList(), comPars);
        }
        // проверка свойств строки
        private void RowProps(TableRow t_row, TableRow t_rowToCompare, MainDocumentPart docPart)
        {
            object val = new object();
            string comRow = "";
            List<Paragraph> comPars = new List<Paragraph>();
            TableRowProperties tableRowPropsToCompare = t_rowToCompare.TableRowProperties;
            TableRowProperties tableRowProps = t_row.TableRowProperties;
            if (tableRowPropsToCompare != null && tableRowProps == null)
            {
                if (tableRowPropsToCompare.HasChildren)
                {
                    // проверка на заголовок и возможность разделения
                    comRow += (tableRowPropsToCompare.ChildElements.ToList().Exists(s => s.LocalName == "tblHeader")) ? "; повторять как заголовок на каждой странице" : "";
                    comRow += (tableRowPropsToCompare.ChildElements.ToList().Exists(s => s.LocalName == "cantSplit")) ? "; разрешить перенос строк на следующую страницу" : "";
                }
            }
            if (comRow.StartsWith("; "))
                comRow = comRow.Remove(0, 2);
            comRow = (comRow != "") ?  "изменить параметры строки: " + comRow + ". " : "";
            comPars.Add((comRow != "") ? new Paragraph(new Run(new Text(comRow))) : null);
            comPars.RemoveAll(s => s == null);
            if (comPars.Count != 0)
                AddCommentOnParagraph(docPart, new List<Paragraph> { t_row.Descendants<Paragraph>().First() }, comPars);
        }

        /// <summary>
        /// Абзацы из таблицы (отдельно для заголовков и основной части)
        /// </summary>
        /// <param name="what_part">часть таблицы (1 - заголовок, 2 - основная часть)</param>
        /// <returns></returns>
        private IEnumerable<Paragraph> TableParsToCompare(int what_part)
        {
            int cols = tableToCompare.Descendants<TableGrid>().First().Descendants<GridColumn>().Count();
            switch (what_part)
            {
                case 1: // в случае заголовка
                    TableCell tCellFirst = tableToCompare.Descendants<TableCell>().First(); // первая ячейка в заголовке
                    paras = tCellFirst.Descendants<Paragraph>(); // абзацы из заголовка
                    vAlignToCompare = tCellFirst.TableCellProperties.TableCellVerticalAlignment;
                    break;
                case 2: // в случае основного содержимого
                    TableCell tCellSecond = tableToCompare.Descendants<TableCell>().ElementAt(cols);
                    paras = tCellSecond.Descendants<Paragraph>();
                    vAlignToCompare = tCellSecond.TableCellProperties.TableCellVerticalAlignment;
                    break;
            }
            styleIdList = paras.Select(st => ParagraphStyle(st)).ToList();
            styleNamePlusIdList = new Dictionary<string, string>();
            // определение стилей в таблице
            foreach (string styleId in styleIdList)
            {
                Style st = allStyleListToCompare[styleId];
                if (!styleNamePlusIdList.ContainsKey(styleId))
                    styleNamePlusIdList.Add(styleId, st.StyleName.Val.Value);
            }
            return paras;
        }

    }
}
