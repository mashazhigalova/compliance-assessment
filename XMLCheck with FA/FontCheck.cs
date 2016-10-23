using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace ComplianceAssessment
{
    // класс с информацией о шрифтах
    class FontDicts
    {
        // линия подчеркивания
        public static Dictionary<string, string> linenDict = new Dictionary<string, string>
        {
            ["Double"] = "двойную линию",
            ["Single"] = "одинарную линию",
            ["Thick"] = "толстую линию",
            ["Wave"] = "волнистую линию"
        };
    }
    class FontCheck : BaseConcepts
    {
        object val = new object();

        Paragraph mainRunP;
        private void General(string par, out object val)
        {
            val = null;
            dictRunPrps.TryGetValue(par, out val);
        }
        private void GeneralToCompare(string par, out object val)
        {
            val = null;
            dictRunToComparePrps.TryGetValue(par, out val);
        }
        // проверка размера шрифта
        public Paragraph SizeCheck()
        {
            string com = "";
            FontSize size = new FontSize();
            FontSize sizeToCompare = new FontSize();

            General("FontSize", out val);
            size = (val != null) ? (FontSize)val : null;

            GeneralToCompare("FontSize", out val);
            sizeToCompare = (val != null) ? (FontSize)val : null;

            if (size == null && sizeToCompare != null)
                com = (Convert.ToDouble(sizeToCompare.Val.Value) / 2).ToString();
            if (size != null && sizeToCompare != null)
                if (size.Val.Value != sizeToCompare.Val.Value)
                    com = (Convert.ToDouble(sizeToCompare.Val.Value) / 2).ToString();
            return (com != "") ? new Paragraph(new Run(new Text("изменить размер шрифта до " + com + " пт"))) : null;
        }
        // проверка цвета шрифта
        public Paragraph ColorFontCheck()
        {
            mainRunP = new Paragraph();
            Color color = new Color();
            Color colorToCompare = new Color();
            General("Color", out val);
            color = (val != null) ? (Color)val : null;
            GeneralToCompare("Color", out val);
            colorToCompare = (val != null) ? (Color)val : null;

            if (color == null && colorToCompare != null)
            { 
                // если стандартный цвет - черный
                if (colorToCompare.Val.Value != "auto" && colorToCompare.Val.Value != "000000")
                    ColorProc(colorToCompare, out mainRunP);
            }
            else 
            if (color != null && colorToCompare == null)
                    mainRunP = new Paragraph(new Run(new Text("изменить цвет шрифта (на черный)")));
            else 
            if (color != null && colorToCompare != null)
                if (color.Val.Value != colorToCompare.Val.Value)
                    ColorProc(colorToCompare, out mainRunP);
            return (mainRunP.InnerText != "") ? mainRunP : null;
        }

        // назначение цвета к шрифту в комментарии
        private void ColorProc(Color colorToCompare, out Paragraph mainRunP)
        {
            mainRunP = new Paragraph();
            Run runWithColor = new Run();
            Run mainRun = new Run();
            RunProperties runPro = new RunProperties();
            string col = " цвет шрифта";

            mainRun = new Run(new Text("изменить"));
            runPro.Append((Color)colorToCompare.CloneNode(true));
            if (colorToCompare.Val.Value == "auto" || colorToCompare.Val.Value == "000000")
            {
                col += " (на черный)";
            }
            runPro.AppendChild(new Text(col) { Space = SpaceProcessingModeValues.Preserve });
            runWithColor.Append(runPro); // добавление цвета к пробегу текста
            mainRunP.Append(new OpenXmlElement[] { mainRun, runWithColor });
        }
        // проверка курсива
        public Paragraph ItalicCheck()
        {
            string com = "";
            Italic italic = new Italic();
            Italic italicToCompare = new Italic();
            string apply = "применить к шрифту курсивное начертание";
            string remove = "убрать курсивное начертание";

            General("Italic", out val);
            italic = (val != null) ? (Italic)val : null;

            GeneralToCompare("Italic", out val);
            italicToCompare = (val != null) ? (Italic)val : null;

            if (italic == null && italicToCompare != null)
            {
                if (italicToCompare.Val != null)
                    com = (italicToCompare.Val.Value == true) ? apply : "";
                else com = apply;
            }
            else if (italic != null && italicToCompare == null)
            {
                if (italic.Val != null)
                    com = (italic.Val.Value == true) ? remove : "";
                else com = remove;
            }
            else if (italic != null && italicToCompare != null)
                if (italic.Val != null && italicToCompare.Val != null)
                {
                    com = (italicToCompare.Val.Value == true && italic.Val.Value == false) ? apply : "";
                    com = (italicToCompare.Val.Value == false && italic.Val.Value == true) ? remove : "";
                }
                else if (italic.Val == null && italicToCompare.Val != null)
                {
                    com = (italicToCompare.Val.Value == false) ? remove : apply;
                }
                else if (italic.Val != null && italicToCompare.Val == null)
                {
                    com = (italic.Val.Value == true) ? remove : "";
                }
            return (com != "") ? new Paragraph(new Run(new Text(com))) : null;
        }
        // проверка полужирного начертания
        public Paragraph BoldCheck()
        {
            string com = "";
            Bold bold = new Bold();
            Bold boldToCompare = new Bold();
            string apply = "применить к шрифту полужирное начертание";
            string remove = "убрать полужирное начертание";

            General("Bold", out val);
            bold = (val != null) ? (Bold)val : null;

            GeneralToCompare("Bold", out val);
            boldToCompare = (val != null) ? (Bold)val : null;

            if (bold == null && boldToCompare != null)
            {
                if (boldToCompare.Val != null)
                    com = (boldToCompare.Val.Value == true) ? apply : "";
                else com = apply;
            }
            else if (bold != null && boldToCompare == null)
            {
                if (bold.Val != null)
                    com = (bold.Val.Value == true) ? remove : "";
                else com = remove;
            }
            else if (bold != null && boldToCompare != null)
                if (bold.Val != null && boldToCompare.Val != null)
                {
                    com = (boldToCompare.Val.Value == true && bold.Val.Value == false) ? apply : "";
                    com = (boldToCompare.Val.Value == false && bold.Val.Value == true) ? remove : "";
                }
                else if (bold.Val == null && boldToCompare.Val != null)
                {
                    com = (boldToCompare.Val.Value == false) ? remove : apply;
                }
                else if (bold.Val != null && boldToCompare.Val == null)
                {
                    com = (bold.Val.Value == true) ? remove : "";
                }
            return (com != "") ? new Paragraph(new Run(new Text(com))) : null;
        }
        // проверка подчеркивания
        public Paragraph UnderlinedCheck()
        {
            string com = "";
            Underline underline = new Underline();
            Underline underlineToCompare = new Underline();

            General("Underline", out val);
            underline = (val != null) ? (Underline)val : null;

            GeneralToCompare("Underline", out val);
            underlineToCompare = (val != null) ? (Underline)val : null;

            if (underline != null && underlineToCompare == null)
                com = "убрать подчеркивание";
            else if (underline == null && underlineToCompare != null)
                com = "изменить подчеркивание шрифта на " + FontDicts.linenDict[underlineToCompare.Val.Value.ToString()];
            else if (underline != null && underlineToCompare != null)
                if (underline.Val.Value != underlineToCompare.Val.Value)
                    com = "изменить подчеркивание шрифта на " + FontDicts.linenDict[underlineToCompare.Val.Value.ToString()];
            return (com != "") ? new Paragraph(new Run(new Text(com))) : null;
        }
        // проверка названия шрифта
        public Paragraph FontNameCheck()
        {
            string com = "";
            RunFonts fontName = new RunFonts();
            RunFonts underlineToCompare = new RunFonts();

            General("RunFonts", out val);
            fontName = (val != null) ? (RunFonts)val : null;

            GeneralToCompare("RunFonts", out val);
            underlineToCompare = (val != null) ? (RunFonts)val : null;
            string fname = "";
            if (fontName == null && underlineToCompare != null)
                com = underlineToCompare.Ascii.Value;
            if (fontName != null && underlineToCompare != null)
            {
                fname = (fontName.Ascii == null) ? fontName.ComplexScript.Value : fontName.Ascii.Value;
                if (fontName.Ascii.Value != underlineToCompare.Ascii.Value)
                    com = underlineToCompare.Ascii.Value;
            }
            return (com != "") ? new Paragraph(new Run(new Text("изменить название шрифта на " + com))) : null;
        }
    }
}
