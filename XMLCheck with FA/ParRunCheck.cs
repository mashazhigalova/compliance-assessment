using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ComplianceAssessment
{
    public class ParRunCheck : BaseConcepts
    {
        string style;
        IEnumerable<Paragraph> parsOriginal; // абзацы из проверяемого документа
        IEnumerable<Paragraph> parsWithDrawing; // абзац с изображением
        bool stop; // продолжать проверку разделов?
        SectionPropertiesCheck sectPropsCheck = new SectionPropertiesCheck(); // экземпляр класса проверки раздела
        // вставка маркированного списка в комментарии
        private void Numbering(MainDocumentPart docPart)
        {
            // определение уровня нумерации
            Level lvl = new Level(
                           new NumberingFormat() { Val = NumberFormatValues.Bullet },
                           new LevelText() { Val = "▪ " }
                         )
            { LevelIndex = 0 };
            NumberingDefinitionsPart numberingPart = docPart.NumberingDefinitionsPart;
            if (numberingPart == null) // если части с нумерации в документе не существует, создать
            {
                numberingPart = docPart.AddNewPart<NumberingDefinitionsPart>("NumberingDefinitionsPart001");               
            }
            var element1 = new AbstractNum(lvl){ AbstractNumberId = 115 };

            var element2 = new NumberingInstance(
                         new AbstractNumId() { Val = 115 }
                       ){ NumberID = 114 };
            if (numberingPart.Numbering == null)
            {
                Numbering element = new Numbering(element1, element2);
                element.Save(numberingPart);
            }
            else
            {
                numberingPart.Numbering.InsertAfter(element1, numberingPart.Numbering.Descendants<AbstractNum>().Last());
                numberingPart.Numbering.InsertAfter(element2, numberingPart.Numbering.Descendants<NumberingInstance>().Last());
            }
        }
        // проверка свойств стилей документа
        public void CheckStyleProps(MainDocumentPart docPart, string fileNameCompare)
        {
            TableCheck t = new TableCheck();
            Numbering(docPart); // проверка списков
            parsOriginal = docPart.Document.Body.Descendants<Paragraph>().Where(pr => pr.Parent == docPart.Document.Body && pr.InnerText!="" && pr.Descendants<Drawing>().Count() == 0);
            DocForComparison(fileNameCompare); // получение стилей из документа для сравнения
            // стили из документа, который проверяется
            allStyleList = docPart.StyleDefinitionsPart.Styles.OfType<Style>().Select(pa => pa).ToDictionary(d => d.StyleId.Value);
            foreach (Paragraph par in parsOriginal)
            {
                Check(pars, docPart, par); // проверка оформления каждого абзаца в документе
            }

            t.CheckTableProps(docPart); // проверка таблиц
            AddCommentsOnIdenticalPars(docPart); // вставка комментариев на абзацы с единым оформлением
            AddCommentsOnIdenticalRuns(docPart); // вставка комментарием на пробеги текста с единым оформлением
            ImageCheck(docPart);
        }
        // вставка комментарием на пробеги текста с единым оформлением
        private void AddCommentsOnIdenticalRuns(MainDocumentPart docPart)
        {
            int comCount = comentsRuns.Count; // количество пробегов текста
            int startIndex = 0;
            int endIndex = 0;
            if (comCount == 1) // если пробег текста один
                AddCommentOnRun(docPart, (List<Run>)comentsRuns.First().ElementAt(0), (List<Paragraph>)comentsRuns.First().ElementAt(1));
            for (int d = 0; d < comCount; d++) // если пробегов текста с одним оформлением больше одного
            {
                if (d + 1 < comentsRuns.Count) // проверка на наличие следующего пробега
                {
                    string stringD = "";
                    string stringDplus1 = "";
                    foreach (var pInList in (List<Paragraph>)comentsRuns.ElementAt(d + 1).ElementAt(1))
                        stringDplus1 += pInList.InnerText; // соединенение текста комментария для следуещего пробега текста
                    foreach (var pInList in (List<Paragraph>)comentsRuns.ElementAt(startIndex).ElementAt(1))
                        stringD += pInList.InnerText; // соединенение текста комментария для текущего пробега текста
                    if (stringDplus1 == stringD) // если текст в комментариях одинаков, продолжаем искать расхождение
                    {
                        endIndex = d + 1;
                    }
                    else
                    {
                        if (startIndex == endIndex) // если дошли до конца коллекции
                            AddCommentOnRun(docPart,
                              (List<Run>)comentsRuns.ElementAt(d).ElementAt(0),
                              (List<Paragraph>)comentsRuns.ElementAt(d).ElementAt(1)); // комментарий на одинаковые пробеги
                        else
                        {
                            List<Run> runsJoint = (List<Run>)comentsRuns.ElementAt(startIndex).ElementAt(0); // соединение пробегов текста с единым оформлением
                            runsJoint.AddRange((List<Run>)comentsRuns.ElementAt(endIndex).ElementAt(0));
                            // комментарий на одинаковые пробеги
                            AddCommentOnRun(docPart,
                               runsJoint,
                               (List<Paragraph>)comentsRuns.ElementAt(d).ElementAt(1));
                        }
                        startIndex = d + 1;
                        endIndex = d + 1;
                    }
                }
                else
                {
                    if (startIndex == endIndex)
                        AddCommentOnRun(docPart,
                          (List<Run>)comentsRuns.ElementAt(d).ElementAt(0),
                          (List<Paragraph>)comentsRuns.ElementAt(d).ElementAt(1));
                    else
                    {
                        List<Run> runsJoint = (List<Run>)comentsRuns.ElementAt(startIndex).ElementAt(0);
                        runsJoint.AddRange((List<Run>)comentsRuns.ElementAt(endIndex).ElementAt(0));

                        AddCommentOnRun(docPart,
                           runsJoint,
                           (List<Paragraph>)comentsRuns.ElementAt(d).ElementAt(1));
                    }
                }

            }
            comentsRuns.Clear(); // очистка комментариев
        }
        // добавление комментариев на абзацы с одинаоквым оформением (принцип идентичен методу AddCommentsOnIdenticalRuns)
        private void AddCommentsOnIdenticalPars(MainDocumentPart docPart)
        {
            int comCount = comentsPars.Count;
            int startIndex = 0;
            int endIndex = 0;
            if (comCount == 1)
                AddCommentOnParagraph(docPart, 
                    new List<Paragraph> { (Paragraph)comentsPars.First().ElementAt(0) }, 
                    (List<Paragraph>)comentsPars.First().ElementAt(1));
            for (int d = 0; d < comCount; d++)
            {
                if (d + 1 < comentsPars.Count)
                {
                    string stringD = "";
                    string stringDplus1 = "";
                    foreach (var pInList in (List<Paragraph>)comentsPars.ElementAt(d + 1).ElementAt(1))
                        stringDplus1 += pInList.InnerText;
                    foreach (var pInList in (List<Paragraph>)comentsPars.ElementAt(startIndex).ElementAt(1))
                        stringD += pInList.InnerText;
                    if (stringDplus1 == stringD)
                    {
                        endIndex = d + 1;
                    }
                    else
                    {
                        if (startIndex == endIndex)
                            AddCommentOnParagraph(docPart,
                              new List<Paragraph> { (Paragraph)comentsPars.ElementAt(startIndex).ElementAt(0) },
                              (List<Paragraph>)comentsPars.ElementAt(d).ElementAt(1));
                        else
                            AddCommentOnParagraph(docPart,
                             new List<Paragraph> { (Paragraph)comentsPars.ElementAt(startIndex).ElementAt(0), (Paragraph)comentsPars.ElementAt(endIndex).ElementAt(0) },
                             (List<Paragraph>)comentsPars.ElementAt(d).ElementAt(1));
                        startIndex = d + 1;
                        endIndex = d + 1;
                    }
                }
                else
                {
                    if (startIndex == endIndex)
                        AddCommentOnParagraph(docPart,
                          new List<Paragraph> { (Paragraph)comentsPars.ElementAt(startIndex).ElementAt(0) },
                          (List<Paragraph>)comentsPars.ElementAt(d).ElementAt(1));
                    else
                        AddCommentOnParagraph(docPart,
                         new List<Paragraph> { (Paragraph)comentsPars.ElementAt(startIndex).ElementAt(0), (Paragraph)comentsPars.ElementAt(endIndex).ElementAt(0) },
                         (List<Paragraph>)comentsPars.ElementAt(d).ElementAt(1));
                }

            }
            comentsPars.Clear();
        }
        // проверка изображения в тексте
        private void ImageCheck(MainDocumentPart docPart)
        {
            // обнаружение абзацев с изображениями
            parsWithDrawing = docPart.Document.Body.Descendants<Paragraph>().Where(pr => pr.Parent == docPart.Document.Body && pr.Descendants<Drawing>().Count() > 0);
            string comment = "";

            foreach (Paragraph p in parsWithDrawing) // для каждого абзаца из списка абзацев с изображениями
            {
                if (parWithDrawingToCompare.Descendants<Drawing>().First().Inline != null && p.Descendants<Drawing>().First().Inline == null)
                    comment = "в тексте";
                if (parWithDrawingToCompare.Descendants<Drawing>().First().Anchor != null && p.Descendants<Drawing>().First().Anchor == null)
                    comment = "обтекание текстом";
                if (comment != "")
                    AddCommentOnParagraph(docPart, new List<Paragraph> { p }, new List<Paragraph>() { new Paragraph(new Run(new Text("Положение рисунка: " + comment))) });
                style = ParagraphStyle(p);
                ParProperties(pars, p, docPart); // проверка параметров абзаца
            }
        }
        // проверка абзацев и пробегов текста в соответствии со стилем
        public void Check(IEnumerable<Paragraph> pToCompare, MainDocumentPart docPart, Paragraph par)
        {
            style = ParagraphStyle(par);  // определение стиля абзаца - получение его StyleId
            if (par.InnerText == "") return;
            if (style == "") return;
            Dictionary<string, string> idName = new Dictionary<string, string>
            { [style] = allStyleList[style].StyleName.Val.Value};
            if (styleNamePlusIdList.ContainsValue(idName[style]))
                style = styleNamePlusIdList.FirstOrDefault(x => x.Value == idName[style]).Key;
            if (styleNamePlusIdList.ContainsKey(style)) // если шаблон содержит стиль, проверка выполняется
            {
                ParProperties(pToCompare, par, docPart); // проверка параметров абзаца
                RunProperties(par, docPart);    // проверка параметров пробегов текста
            }
            else AddCommentOnParagraph(docPart, new List<Paragraph> { par }, new List<Paragraph>()
            { new Paragraph(new Run(new Text("к абзацу применен необычный стиль!")))});
        }
        // проверка абзаца
        private void ParProperties(IEnumerable<Paragraph> pToCompare, Paragraph p, MainDocumentPart docPart)  // свойства абзаца в проверяемом документе
        {
            ParagraphProperties pPr = p.ParagraphProperties; // параметры абзаца 
            ParagraphProperties pPrComp = GetParPropsFromPar(pToCompare, style);
            dictParProps = DictionaryFromType(pPr);    // получение всех параметров абзаца
            dictToCompareParProps = DictionaryFromType(pPrComp);

            List<Paragraph> parMas = new List<Paragraph>();
            ParagraphCheck pCheck = new ParagraphCheck();
            NumberingCheck nCheck = new NumberingCheck();
            List<Paragraph> comParsNum = nCheck.NumberingListsCheck(pToCompare, p, style, docPart); // проверка нумерованных и маркированных списков
            if (comParsNum != null) // если списки содержатся
                parMas.AddRange(comParsNum);
            parMas.Add(pCheck.JustificationCheck()); // проверка выравнивания
            parMas.AddRange(pCheck.SpacingBetweenLinesCheck()); // проверка интервалов
            parMas.AddRange(pCheck.IndentCheck()); // проверка отступов
            parMas.RemoveAll(s => s == null);
            if (parMas.Count != 0)
                comentsPars.Add(new List<object> { p, parMas}); // вставка абзацев с комментариями

            try
            {
                // проверка свойств раздела; комментарий ставится на последний абзац раздела
                if (pPr.ChildElements.ToList().Exists(s => s.LocalName == "sectPr"))
                    sectPropsCheck.SectionProps(docPart, pPr.SectionProperties, p);
                else if (parsOriginal != null && !stop)
                {
                    sectPropsCheck.SectionProps(docPart, docPart.Document.Body.Descendants<SectionProperties>().First(), parsOriginal.First());
                    stop = true;
                }
            }
            catch { return; }
        }
        // проверка параметров пробегов текста
        private void RunProperties(Paragraph p, MainDocumentPart docPart)
        {
            FontCheck fCheck = new FontCheck();
            dictRunPrps = new Dictionary<string, object>();
            
            string regLng = @"<w:lang(.*)/>";
            int i = 0;
            IEnumerable<Run> runsInPar = p.Elements<Run>().Where(pr => pr.InnerText != "");
            int runsCount = runsInPar.Count();
            List<Run> runsWithSameForm = new List<Run>();
            
            while (i < runsCount) // получение пробегов с единым форматированием
            {
                if(runsWithSameForm.Count == 0)
                    runsWithSameForm.Add(runsInPar.ElementAt(i));
                string rp1 = runsInPar.ElementAt(i).RunProperties.InnerXml;
                rp1 = rp1.Substring(rp1.IndexOf("<w:rFonts"));
                rp1 = Regex.Replace(rp1, regLng, "");
                string rp2 = "";
                if (i + 1 < runsCount)
                {
                    rp2 = runsInPar.ElementAt(i+1).RunProperties.InnerXml;
                    rp2 = rp2.Substring(rp2.IndexOf("<w:rFonts"));
                    rp2 = Regex.Replace(rp2, regLng, "");
                }
                if (rp1 == rp2)
                {
                    runsWithSameForm.Add(runsInPar.ElementAt(i + 1));
                    i++;
                }
                else
                {
                    List<Paragraph> parsCom = new List<Paragraph>();
                    dictRunPrps = DictionaryFromType(runsInPar.ElementAt(i).RunProperties);
                    parsCom.Add(fCheck.SizeCheck()); // размер шрифта
                    parsCom.Add(fCheck.ColorFontCheck()); // цвет шрифта
                    parsCom.Add(fCheck.ItalicCheck()); // курсив
                    parsCom.Add(fCheck.BoldCheck()); // полужирный
                    parsCom.Add(fCheck.UnderlinedCheck()); // подчеркнутый
                    parsCom.Add(fCheck.FontNameCheck()); // имя шрифта
                    parsCom.RemoveAll(s => s == null);
                    if (parsCom.Count != 0 && runsWithSameForm.Count != 0)
                        comentsRuns.Add(new List<object> { runsWithSameForm, parsCom });
                    runsWithSameForm = new List<Run>();
                    i++;
                }
            }
        }


    }
}
