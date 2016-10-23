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
    class NumberingCheck: BaseConcepts
    {
        Level lvlOriginal; // уровень списка в проверяемом документе
        Level lvlToCompare; // уровень списка в шаблонном документе
        Paragraph comParsNumRemove = new Paragraph();

        // получение параметров форматирования списков
        public void GetProps(IEnumerable<Paragraph> pToCompare, Paragraph p, string style, MainDocumentPart docPart)
        {
            Numbering numbering = docPart.NumberingDefinitionsPart.Numbering;
            NumberingProperties numPr = null;
            ParagraphProperties pPr = p.ParagraphProperties;
            List<ParagraphProperties> pPrComp = GetNumPropsFromPars(pToCompare, style);
            if (pPrComp == null) return;
            List<NumberingProperties> numCompProps = new List<NumberingProperties>();
            // определение свойств нумерованных списков
            if (pPr.ChildElements.ToList().Exists(s => s.LocalName == "numPr"))
                numPr = pPr.NumberingProperties;
            // проверка есть ли у абзаца нумерация
            foreach (var props in pPrComp)
                if (props.ChildElements.ToList().Exists(s => s.LocalName == "numPr"))
                    numCompProps.Add(props.NumberingProperties);

            // если свойств нумерации нет, то убрать список
            if(numPr!=null && numCompProps.Count==0)
            {
                comParsNumRemove = new Paragraph(new Run(new Text("убрать список")));
                return;
            }

            // получение абстр numId из оригинального документа
            AbstractNumId abNumidOriginal = numbering.Descendants<NumberingInstance>().First(po => po.NumberID.Value == numPr.NumberingId.Val).AbstractNumId;
            // получение абстр num из докуента для сравнения
            AbstractNum abNumOriginal = numbering.Descendants<AbstractNum>().First(l => l.AbstractNumberId.Value == abNumidOriginal.Val);
            lvlOriginal = abNumOriginal.Descendants<Level>().First(lev => lev.LevelIndex.Value == numPr.NumberingLevelReference.Val);

            if (lvlOriginal.NumberingFormat.Val == "bullet") // если список маркированный
            {
                foreach (var n in numCompProps)
                    if (levelComp(n).NumberingFormat.Val == "bullet")
                        lvlToCompare = levelComp(n);
            }
            else // если список нумерованный
                lvlToCompare = levelComp(numCompProps.First(lvl => levelComp(lvl).NumberingFormat.Val != "bullet"));
        }
        // получение уровня списка
        public static Level levelComp(NumberingProperties numPrToCompare)
        {
            // получение абстр numId из документа для сравнения
            AbstractNumId abNumid = numberingToCompare.Descendants<NumberingInstance>().First(p => p.NumberID.Value == numPrToCompare.NumberingId.Val).AbstractNumId;
            // получение абстр num из докуента для сравнения
            AbstractNum abNum = numberingToCompare.Descendants<AbstractNum>().First(l => l.AbstractNumberId.Value == abNumid.Val);
            return abNum.Descendants<Level>().First(lev => lev.LevelIndex.Value == numPrToCompare.NumberingLevelReference.Val);
        }
        // получение параметров абзацев из списка
        public static List<ParagraphProperties> GetNumPropsFromPars(IEnumerable<Paragraph> paragraphs, string styleId)
        {
            List<ParagraphProperties> paragraph = new List<ParagraphProperties>();
            List<Paragraph> newTempList = paragraphs.Where(p => p.ParagraphProperties.ParagraphStyleId != null).ToList();
            List<Paragraph> parsNum = newTempList.Where(p => p.ParagraphProperties.ParagraphStyleId.Val.Value.Contains(styleId)).ToList();
            if (parsNum == null) return null;
            foreach (var p in parsNum)
                paragraph.Add(p.ParagraphProperties);
            return paragraph;
        }
        // проверка списков
        public List<Paragraph> NumberingListsCheck(IEnumerable<Paragraph> pToCompare, Paragraph p, string style, MainDocumentPart docPart)
        {
            List<Paragraph> comPars = new List<Paragraph>();
            try
            {
                GetProps(pToCompare, p, style, docPart); // получение параметров форматирования списков (уровней)
                if (comParsNumRemove.InnerText != "")
                    comPars.Add(comParsNumRemove);
                else
                {
                    string level = "";
                    string regex = @"%\d";
                    if (lvlOriginal.LevelIndex.Value != lvlToCompare.LevelIndex.Value) // если уровни списков не равны
                        comPars.Add(
                            new Paragraph(new Run(
                                new Text("уровень списка " + lvlToCompare.LevelIndex.Value.ToString() + " (0 - начальное значение)"))));
                    if (lvlOriginal.LevelText.Val.Value != lvlToCompare.LevelText.Val.Value) // если вид пунктов неодинаков
                    {
                        string lvlText = lvlToCompare.LevelText.Val.Value;
                        comPars.Add(
                            new Paragraph(new Run(
                                new Text("пункт списка имет вид " + (lvlText.Contains("%") ? Regex.Replace(lvlToCompare.LevelText.Val.Value, regex, "*") : lvlText)))));
                    }
                    if (lvlOriginal.NumberingFormat.Val != lvlToCompare.NumberingFormat.Val) // если формат списков не совпадает
                        comPars.Add(
                            new Paragraph(new Run(
                                new Text("формат номера списка - " + ParagraphDicts.numFormatDict[lvlToCompare.NumberingFormat.Val.Value.ToString()]))));
                    if (lvlOriginal.StartNumberingValue.Val != lvlToCompare.StartNumberingValue.Val) // разные начальные значения
                        comPars.Add(
                            new Paragraph(new Run(
                                new Text(level += "начать список  с " + lvlToCompare.StartNumberingValue.Val.ToString() + "-ого значения"))));
                }
            }
            catch { return null; } // обнаружение ошибки
            return comPars; // возвращение комментария по результатам проверки
        }

    }
}
