using DocumentFormat.OpenXml;
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
    class ParagraphDicts // словари с параметрами абзаца
    {
        // выравнивание текста
        public static Dictionary<string, string> justificationDict = new Dictionary<string, string>
        {
            ["Both"] = "ширине страницы",
            ["Left"] = "левому краю",
            ["Right"] = "правому краю",
            ["Center"] = "центру"
        };
        // правило междустрочного интервала
        public static Dictionary<string, string> lineRuleDict = new Dictionary<string, string>
        {
            ["Auto"] = "авто",
            ["Exact"] = "точно",
            ["AtLeast"] = "минимум"
        };
        // формат нумерации
        public static Dictionary<string, string> numFormatDict = new Dictionary<string, string>
        {
            ["CardinalText"] = "количественные числительные (Один, Два, Три..)",
            ["Decimal"] = "числа в десятичной системе (1, 2, 3..)",
            ["DecimalEnclosedCircle"] = "числа, заключенные в окружность",
            ["DecimalEnclosedFullstop"] = "числа с точкой",
            ["DecimalEnclosedParen"] = "числа, заключенные в скобки",
            ["DecimalZero"] = "числа в десятичной системе (01, 02, 03..)",
            ["LowerLetter"] = "буквы в нижнем регистре",
            ["LowerRoman"] = "римские числа в нижнем регистре (i, ii, iii..)",
            ["None"] = "отсутсвует",
            ["OrdinalText"] = "порядковые числительные (Первый, Второй, Третий..)",
            ["UpperLetter"] = "буквы в верхнем регистре",
            ["Bullet"] = "маркер",
            ["UpperRoman"] = "римские числа в верхнем регистре (I, II, III..)"
        };
    }
    // класс для проверки параметров абзаца
    public class ParagraphCheck : BaseConcepts
    {
        private Indentation indentPar; // отступы в проверямом абзаце
        private Indentation indentStyle; // отступы в шаблонном абзаце
        private SpacingBetweenLines spacingPar; // интервалы в проверяемом абзаце
        private SpacingBetweenLines spacingStyle; // интервалы в шаблонном абзаце

        object val = new object();
        /// <summary>
        /// Проверка выравнивания текста в абзаце
        /// </summary>
        public Paragraph JustificationCheck()
        {
            string comment = "";
            Justification jPar = new Justification();
            Justification jStyleToCompare = new Justification();
            object val;
            // извлечение параметра выравнивание из свойств абзаца
            dictParProps.TryGetValue("Justification", out val);
            jPar = (val != null) ? (Justification)val : null;

            dictToCompareParProps.TryGetValue("Justification", out val);
            jStyleToCompare = (val != null) ? (Justification)val : null;
            // сравнение параметров выравнивания
            if (jPar == null && jStyleToCompare != null) // если параметр не определен в проверяемом документе
            {
                if (jStyleToCompare.Val.Value.ToString() != "Left")
                    comment = ParagraphDicts.justificationDict[jStyleToCompare.Val.Value.ToString()];
            }
            else if (jPar != null && jStyleToCompare != null) // если параметры в обоих документах представлены
                if (jPar.Val.Value != jStyleToCompare.Val.Value)
                    comment = ParagraphDicts.justificationDict[jStyleToCompare.Val.Value.ToString()];
            return (comment != "") ? new Paragraph(new Run(new Text("выровнять по " + comment))) : null;
        }
        // проверка отступа слева
        private Paragraph LeftIndent()
        {
            string com = "";
            // отступ слева в шаблонном документе
            double styleIndent = (indentStyle.Left != null) ? 
                Math.Round(Convert.ToDouble(indentStyle.Left.Value) / 567, 2) : 0;
            // отступ слева в проверяемом документе
            double parIndent = (indentPar.Left != null) ? 
                Math.Round(Convert.ToDouble(indentPar.Left.Value) / 567, 2) : 0;
            if (styleIndent != parIndent) // если значения не совпадают, 
                com = styleIndent.ToString(); //записать в комментарий значение из шаблона
            return (com != "") ? 
                new Paragraph(new Run(new Text("изменить левый отступ до " + com + " см"))) : null;
        }
        // проверка отступа справа
        private Paragraph RightIndent()
        {
            string com = "";
            double styleIndent = (indentStyle.Right != null) ? Math.Round(Convert.ToDouble(indentStyle.Right.Value) / 567, 2) : 0;
            double parIndent = (indentPar.Right != null) ? Math.Round(Convert.ToDouble(indentPar.Right.Value) / 567, 2) : 0;
            if (styleIndent != parIndent)
                com = styleIndent.ToString();
            return (com != "") ? new Paragraph(new Run(new Text("изменить правый отступ до " + com + " см"))) : null;
        }
        // проверка отступа/выступа первой строки
        private Paragraph FirstLineIndent()
        {
            string com = "";
            double styleIndent = (indentStyle.FirstLine != null) ? Math.Round(Convert.ToDouble(indentStyle.FirstLine.Value) / 567, 2) : 0;
            double parIndent = (indentPar.FirstLine != null) ? Math.Round(Convert.ToDouble(indentPar.FirstLine.Value) / 567, 2) : 0;
            if (styleIndent != parIndent)
                com = styleIndent.ToString();
            return (com != "") ? new Paragraph(new Run(new Text("изменить отступ первой строки до " + com + " см"))) : null;
        }
        /// <summary>
        /// Проверка отступов 
        /// </summary>
        public List<Paragraph> IndentCheck()
        {
            // список абзацев, которые будут включены в комментарий
            List<Paragraph> parsMas = new List<Paragraph>();
            // составления списка данных об отступах для проверяемого документа
            indentPar = new Indentation();
            dictParProps.TryGetValue("Indentation", out val);
            indentPar = (val != null) ? (Indentation)val : null;
        
            // список информации об отсупах для документа для сравнения
            indentStyle = new Indentation();
            dictToCompareParProps.TryGetValue("Indentation", out val);
            indentStyle = (val != null) ? (Indentation)val : null;
            
            if (indentPar != null && indentStyle != null)
            {
                parsMas.Add((LeftIndent() != null) ? LeftIndent() : null);
                parsMas.Add((RightIndent() != null) ? RightIndent() : null);
                parsMas.Add((FirstLineIndent() != null) ? FirstLineIndent() : null);
            }
            return parsMas;
        }
        /// <summary>
        /// Проверка интервалов
        /// </summary>
        public List<Paragraph> SpacingBetweenLinesCheck()
        {
            // список абзацев, которые будут включены в комментарий
            List<Paragraph> parsMas = new List<Paragraph>();
            // составления списка данных об отступах для проверяемого документа
            spacingPar = new SpacingBetweenLines();
            dictParProps.TryGetValue("SpacingBetweenLines", out val);
            spacingPar = (val != null) ? (SpacingBetweenLines)val : null;

            // список информации об отсупах для документа для сравнения
            spacingStyle = new SpacingBetweenLines();
            dictToCompareParProps.TryGetValue("SpacingBetweenLines", out val);
            spacingStyle = (val != null) ? (SpacingBetweenLines)val : null;
            
            if (spacingPar != null && spacingStyle != null)
            {
                parsMas.Add((AfterSpacing() != null) ? AfterSpacing() : null);
                parsMas.Add((BeforeSpacing() != null) ? BeforeSpacing() : null);
                parsMas.Add((LineRule() != null) ? LineRule() : null);
                parsMas.Add((Line() != null) ? Line() : null);
            }
            return parsMas;
        }
        // проверка интервала после абзаца
        private Paragraph AfterSpacing()
        {
            string com = "";
            double styleSpace = (spacingStyle.After != null) ? Math.Round(Convert.ToDouble(spacingStyle.After.Value) / 20) : 0;
            double parSpace = (spacingPar.After != null) ? Math.Round(Convert.ToDouble(spacingPar.After.Value) / 20) : 0;
            if (styleSpace != parSpace)
                com = styleSpace.ToString();
            return (com != "") ? new Paragraph(new Run(new Text("изменить интервал после абзаца до " + com + " пт"))) : null;
        }
        // проверка интервала перед абзацем
        private Paragraph BeforeSpacing()
        {
            string com = "";
            double styleSpace = (spacingStyle.Before != null) ? Math.Round(Convert.ToDouble(spacingStyle.Before.Value) / 20) : 0;
            double parSpace = (spacingPar.Before != null) ? Math.Round(Convert.ToDouble(spacingPar.Before.Value) / 20) : 0;
            if (styleSpace!=parSpace)
                com = styleSpace.ToString();
            return (com != "") ? new Paragraph(new Run(new Text("изменить интервал перед абзацем до " + com + " пт"))) : null;
        }
        // проверка междутстрочного интервала
        private Paragraph Line()
        {
            string com = "";
            double styleSpace = (spacingStyle.Line != null) ? Math.Round(Convert.ToDouble(spacingStyle.Line.Value) / 240, 2) : 0;
            double parSpace = (spacingPar.Line != null) ? Math.Round(Convert.ToDouble(spacingPar.Line.Value) / 240, 2) : 0;
            if (styleSpace != parSpace)
                com = styleSpace.ToString() + " (" + Math.Round(Convert.ToDouble(spacingStyle.Line.Value) / 20).ToString() + " пт" + ")";
            return (com != "") ? new Paragraph(new Run(new Text("изменить интервал между строками до " + com))) : null;
        }
        // проверка правила междустрочного интервала
        private Paragraph LineRule()
        {
            string com = "";
            if (spacingPar.LineRule == null && spacingStyle.LineRule != null)
                com = ParagraphDicts.lineRuleDict[spacingStyle.LineRule.Value.ToString()]; // извлечение значения из словаря
            else if (spacingPar.LineRule != null && spacingStyle.LineRule != null)
                if (spacingPar.LineRule.Value != spacingStyle.LineRule.Value)
                    com = ParagraphDicts.lineRuleDict[spacingStyle.LineRule.Value.ToString()];
            return (com != "") ? new Paragraph(new Run(new Text("изменить интервал между строками на " + com))) : null;
        }
    }
}
