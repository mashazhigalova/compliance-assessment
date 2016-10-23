using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComplianceAssessment
{
    // класс для проверки свойств таблицы
    class TablePropertiesCheck: TableCheck
    {
        // словарь с типами линий
        static Dictionary<string, string> valDict = new Dictionary<string, string>
        {
            ["Single"] = "одинарный",
            ["Thick"] = "жирный",
            ["None"] = "невидимый",
            ["Nil"] = "невидимый",
            ["Dashed"] = "пунктирный"
        };

        // проверка границ таблицы
        public Paragraph BordersCheck()
        {
            string com = "";
            com = TopBorderCheck() 
                + LeftBorderCheck() 
                + RightBorderCheck() 
                + BottomBorderCheck() 
                + InsideHBCheck() 
                + InsideVBCheck();
            return (com != "") ? new Paragraph(new Run(new Text(com))) : null;
        }
        // проверка полей ячеек таблицы
        public Paragraph MarginsCheck()
        {
            string com = "";
            com = TopMarginCheck()
                + BotMarginCheck()
                + RightMarginCheck()
                + LeftMarginCheck();
            if (com != "")
                com = com.Remove(0, 2);
            return (com != "") ? new Paragraph(new Run(new Text("изменить значения: " + com))) : null;
        }
        // проверка верхней границы таблицы
        private static string TopBorderCheck()
        {
            object val = new object();
            string comment = "";
            string col = "";
            TopBorder topBorderO = new TopBorder(); // для проверяемого
            TopBorder topBorderC = new TopBorder(); // для проверки

            topBorderO = (tableBorders.ChildElements.ToList().Exists(s => s.LocalName == "top")) ?
                tableBorders.TopBorder : null;
            
            topBorderC = (tableBordersToCompare.ChildElements.ToList().Exists(s => s.LocalName == "top")) ?
                tableBordersToCompare.TopBorder : null;

            if (topBorderO.Size == null && topBorderC.Size != null) // толщина линии
                comment += "толщину линии на " + Math.Round(Convert.ToDouble(topBorderC.Size.Value) / 8, 2).ToString() + " пт";
            else if (topBorderO.Size != null && topBorderC.Size != null)
                comment += (topBorderO.Size.Value != topBorderC.Size.Value) ? "толщину линии на " + Math.Round(Convert.ToDouble(topBorderC.Size.Value) / 8, 2).ToString() + " пт" : "";

            if (topBorderO.Color == null && topBorderC.Color != null) // цвет линии
                comment += "; цвет линии на " + topBorderC.Color;
            else if (topBorderO.Color != null && topBorderC.Color != null)
            {
                if (topBorderC.Color == "auto" || topBorderC.Color == "000000") // если стандарный цвет
                    col = "черный";
                else col = topBorderC.Color;
                comment += (topBorderO.Color.Value != topBorderC.Color.Value) ? "; цвет линии на " + col : "";
            }

            if (topBorderO.Val == null && topBorderC.Val != null)   // тип линии границы
                comment += "; тип линии границы на " + topBorderC.Val.Value;
            else if (topBorderO.Val != null && topBorderC.Val != null)
                comment += (topBorderO.Val.Value != topBorderC.Val.Value) ? "; тип линии границы на " + valDict[topBorderC.Val.Value.ToString()] : "";

            if (comment.StartsWith("; "))
                comment = comment.Remove(0, 2);
            return (comment != "") ? "изменить параметры верхней границы: "  + comment + ". " : "";
        }
        // проверка левой границы
        private static string LeftBorderCheck()
        {
            object val = new object();
            string comment = "";
            string col = "";
            LeftBorder leftBorderO = new LeftBorder(); // для проверяемого
            LeftBorder leftBorderC = new LeftBorder(); // для проверки
            
            leftBorderO = (tableBorders.ChildElements.ToList().Exists(s => s.LocalName == "left")) ?
                tableBorders.LeftBorder : null;
            
            leftBorderC = (tableBordersToCompare.ChildElements.ToList().Exists(s => s.LocalName == "left")) ?
                tableBordersToCompare.LeftBorder : null;

            if (leftBorderO.Size == null && leftBorderC.Size != null)
                comment += "толщину линии на " + Math.Round(Convert.ToDouble(leftBorderC.Size.Value) / 8, 2).ToString() + " пт";
            else if (leftBorderO.Size != null && leftBorderC.Size != null)
                comment += (leftBorderO.Size.Value != leftBorderC.Size.Value) ? "толщину линии на " + Math.Round(Convert.ToDouble(leftBorderC.Size.Value) / 8, 2).ToString() + " пт" : "";

            if (leftBorderO.Color == null && leftBorderC.Color != null)
                comment += "; цвет линии на " + leftBorderC.Color;
            else if (leftBorderO.Color != null && leftBorderC.Color != null)
            {
                if (leftBorderC.Color == "auto" || leftBorderC.Color == "000000")
                    col = "черный";
                else col = leftBorderC.Color;
                comment += (leftBorderO.Color.Value != leftBorderC.Color.Value) ? "; цвет линии на " + col : "";
            }

            if (leftBorderO.Val == null && leftBorderC.Val != null)
                comment += "; тип линии границы на " + leftBorderC.Val.Value;
            else if (leftBorderO.Val != null && leftBorderC.Val != null)
                comment += (leftBorderO.Val.Value != leftBorderC.Val.Value) ? "; тип линии границы на " + valDict[leftBorderC.Val.Value.ToString()] : "";

            if (comment.StartsWith("; "))
                comment = comment.Remove(0, 2);
            return (comment != "") ? "изменить параметры левой границы: " + comment + ". " : "";
        }
        // проверка правой границы
        private static string RightBorderCheck()
        {
            object val = new object();
            string comment = "";
            string col = "";
            RightBorder rightBorderO = new RightBorder(); // для проверяемого
            RightBorder rightBorderC = new RightBorder(); // для проверки
            
            rightBorderO = (tableBorders.ChildElements.ToList().Exists(s => s.LocalName == "right")) ?
                tableBorders.RightBorder : null; ;

            rightBorderC = (tableBordersToCompare.ChildElements.ToList().Exists(s => s.LocalName == "right")) ?
                tableBordersToCompare.RightBorder : null;

            if (rightBorderO.Size == null && rightBorderC.Size != null)
                comment += "толщину линии на " + Math.Round(Convert.ToDouble(rightBorderC.Size.Value) / 8, 2).ToString() + " пт";
            else if (rightBorderO.Size != null && rightBorderC.Size != null)
                comment += (rightBorderO.Size.Value != rightBorderC.Size.Value) ? "толщину линии на " + Math.Round(Convert.ToDouble(rightBorderC.Size.Value) / 8, 2).ToString() + " пт" : "";

            if (rightBorderO.Color == null && rightBorderC.Color != null)
                comment += "; цвет линии на " + rightBorderC.Color;
            else if (rightBorderO.Color != null && rightBorderC.Color != null)
            {
                if (rightBorderC.Color == "auto" || rightBorderC.Color == "000000")
                    col = "черный";
                else col = rightBorderC.Color;
                comment += (rightBorderO.Color.Value != rightBorderC.Color.Value) ? "; цвет линии на " + col : "";
            }

            if (rightBorderO.Val == null && rightBorderC.Val != null)
                comment += "; тип линии границы на " + rightBorderC.Val.Value;
            else if (rightBorderO.Val != null && rightBorderC.Val != null)
                comment += (rightBorderO.Val.Value != rightBorderC.Val.Value) ? "; тип линии границы на " + valDict[rightBorderC.Val.Value.ToString()] : "";

            if (comment.StartsWith("; "))
                comment = comment.Remove(0, 2);
            return (comment != "") ? "изменить параметры правой границы: " + comment + ". " : "";
        }
        // проверка нижней границы
        private static string BottomBorderCheck()
        {
            object val = new object();
            string col = "";
            string comment = "";
            BottomBorder botBorderO = new BottomBorder(); // для проверяемого
            BottomBorder botBorderC = new BottomBorder(); // для проверки

            botBorderO = (tableBorders.ChildElements.ToList().Exists(s => s.LocalName == "bottom")) ?
                tableBorders.BottomBorder : null;

            botBorderC = (tableBordersToCompare.ChildElements.ToList().Exists(s => s.LocalName == "bottom")) ?
                tableBordersToCompare.BottomBorder : null;

            if (botBorderO.Size == null && botBorderC.Size != null)
                comment += "толщину линии на " + Math.Round(Convert.ToDouble(botBorderC.Size.Value) / 8, 2).ToString() + " пт";
            else if (botBorderO.Size != null && botBorderC.Size != null)
                comment += (botBorderO.Size.Value != botBorderC.Size.Value) ? "толщину линии на " + Math.Round(Convert.ToDouble(botBorderC.Size.Value) / 8, 2).ToString() + " пт" : "";

            if (botBorderO.Color == null && botBorderC.Color != null)
                comment += "; цвет линии на " + botBorderC.Color;
            else if (botBorderO.Color != null && botBorderC.Color != null)
            {
                if (botBorderC.Color == "auto" || botBorderC.Color == "000000")
                    col = "черный";
                else col = botBorderC.Color;
                comment += (botBorderO.Color.Value != botBorderC.Color.Value) ? "; цвет линии на " + col : "";
            }

            if (botBorderO.Val == null && botBorderC.Val != null)
                comment += "; тип линии границы на " + botBorderC.Val.Value;
            else if (botBorderO.Val != null && botBorderC.Val != null)
                comment += (botBorderO.Val.Value != botBorderC.Val.Value) ? "; тип линии границы на " + valDict[botBorderC.Val.Value.ToString()] : "";

            if (comment.StartsWith("; "))
                comment = comment.Remove(0, 2);
            return (comment != "") ? "изменить параметры нижней границы: " + comment + ". " : "";
        }
        // проверка внутренней горизонтальной границы
        private static string InsideHBCheck()
        {
            object val = new object();
            string comment = "";
            string col = "";
            InsideHorizontalBorder insideBorderO = new InsideHorizontalBorder(); // для проверяемого
            InsideHorizontalBorder insideBorderC = new InsideHorizontalBorder(); // для проверки

            insideBorderO = (tableBordersToCompare.ChildElements.ToList().Exists(s => s.LocalName == "insideH")) ?
                tableBordersToCompare.InsideHorizontalBorder : null;

            insideBorderC = (tableBordersToCompare.ChildElements.ToList().Exists(s => s.LocalName == "insideH")) ?
                tableBordersToCompare.InsideHorizontalBorder : null;

            if (insideBorderO.Size == null && insideBorderC.Size != null)
                comment += "толщину линии на " + Math.Round(Convert.ToDouble(insideBorderC.Size.Value) / 8, 2).ToString() + " пт";
            else if (insideBorderO.Size != null && insideBorderC.Size != null)
                comment += (insideBorderO.Size.Value != insideBorderC.Size.Value) ? "толщину линии на " + Math.Round(Convert.ToDouble(insideBorderC.Size.Value) / 8, 2).ToString() + " пт" : "";

            if (insideBorderO.Color == null && insideBorderC.Color != null)
                comment += "; цвет линии на " + insideBorderC.Color;
            else if (insideBorderO.Color != null && insideBorderC.Color != null)
            {
                if (insideBorderC.Color == "auto" || insideBorderC.Color == "000000")
                    col = "черный";
                else col = insideBorderC.Color;
                comment += (insideBorderO.Color.Value != insideBorderC.Color.Value) ? "; цвет линии на " + col : "";
            }
            if (insideBorderO.Val == null && insideBorderC.Val != null)
                comment += "; тип линии границы на " + insideBorderC.Val.Value;
            else if (insideBorderO.Val != null && insideBorderC.Val != null)
                comment += (insideBorderO.Val.Value != insideBorderC.Val.Value) ? "; тип линии границы на " + valDict[insideBorderC.Val.Value.ToString()] : "";

            if (comment.StartsWith("; "))
                comment = comment.Remove(0, 2);
            return (comment != "") ? "изменить параметры внутренних горизонтальных границ: " + comment + ". " : "";
        }
        // проверка внутренней вертикальной границы
        private static string InsideVBCheck()
        {
            object val = new object();
            string comment = "";
            string col = "";
            InsideVerticalBorder insideVBorderO = new InsideVerticalBorder(); // для проверяемого
            InsideVerticalBorder insideVBorderC = new InsideVerticalBorder(); // для проверки

            insideVBorderO = (tableBorders.ChildElements.ToList().Exists(s => s.LocalName == "insideV")) ?
                tableBorders.InsideVerticalBorder : null;

            insideVBorderC = (tableBordersToCompare.ChildElements.ToList().Exists(s => s.LocalName == "insideV")) ?
                tableBordersToCompare.InsideVerticalBorder : null;

            if (insideVBorderO.Size == null && insideVBorderC.Size != null)
                comment += "толщину линии на " + Math.Round(Convert.ToDouble(insideVBorderC.Size.Value) / 8, 2).ToString() + " пт";
            else if (insideVBorderO.Size != null && insideVBorderC.Size != null)
                comment += (insideVBorderO.Size.Value != insideVBorderC.Size.Value) ? "толщину линии на " + Math.Round(Convert.ToDouble(insideVBorderC.Size.Value) / 8, 2).ToString() + " пт" : "";
            
            if (insideVBorderO.Color == null && insideVBorderC.Color != null)
                comment += "; цвет линии на " + insideVBorderC.Color;
            else if (insideVBorderO.Color != null && insideVBorderC.Color != null)
            {
                if (insideVBorderC.Color == "auto" || insideVBorderC.Color == "000000")
                    col = "черный";
                else col = insideVBorderC.Color;
                comment += (insideVBorderO.Color.Value != insideVBorderC.Color.Value) ? "; цвет линии на " + col : "";
            }
            if (insideVBorderO.Val == null && insideVBorderC.Val != null)
                comment += "; тип линии границы на " + insideVBorderC.Val.Value;
            else if (insideVBorderO.Val != null && insideVBorderC.Val != null)
                comment += (insideVBorderO.Val.Value != insideVBorderC.Val.Value) ? "; тип линии границы на " + valDict[insideVBorderC.Val.Value.ToString()] : "";

            if (comment.StartsWith("; "))
                comment = comment.Remove(0, 2);
            return (comment != "") ? "изменить параметры внутренних вертикальных границ: " + comment + ". " : "";
        }
        // проверка верхнего поля ячейки
        private static string TopMarginCheck()
        {
            object val = new object();
            string comment = "";
            string top = "; верхнего поля ячеек до ";
            TopMargin topMarginO = new TopMargin(); // для проверяемого
            TopMargin topMarginC = new TopMargin(); // для проверки

            topMarginO = (tableMarCells.ChildElements.ToList().Exists(s => s.LocalName == "top")) ?
                tableMarCells.TopMargin : null;

            topMarginC = (tableMarCellsToCompare.ChildElements.ToList().Exists(s => s.LocalName == "top")) ?
                tableMarCellsToCompare.TopMargin : null;
            if (topMarginC != null) // если в шаблоне значение существует
            {
                if (topMarginO.Width == null && topMarginC.Width != null)
                    comment = top + Math.Round(Convert.ToDouble(topMarginC.Width.Value) / 567, 2).ToString() + " см";
                else if (topMarginO.Width != null && topMarginC.Width != null)
                    comment = (topMarginO.Width.Value != topMarginC.Width.Value) ? top + Math.Round(Convert.ToDouble(topMarginC.Width.Value) / 567, 2).ToString() + " см" : "";
            }
            return comment;
        }
        // проверка нижнего поля ячейки
        public static string BotMarginCheck()
        {
            object val = new object();
            string comment = "";
            string bot = "; нижнего поля ячеек до ";
            BottomMargin botMarginO = new BottomMargin(); // для проверяемого
            BottomMargin botMarginC = new BottomMargin(); // для проверки

            botMarginO = (tableMarCells.ChildElements.ToList().Exists(s => s.LocalName == "bottom")) ?
                tableMarCells.BottomMargin : null;

            botMarginC = (tableMarCellsToCompare.ChildElements.ToList().Exists(s => s.LocalName == "bottom")) ?
                tableMarCellsToCompare.BottomMargin : null;
            if (botMarginC != null)
            {
                if (botMarginO.Width == null && botMarginC.Width != null)
                    comment = bot + Math.Round(Convert.ToDouble(botMarginC.Width.Value) / 567, 2).ToString() + " см";
                else if (botMarginO.Width != null && botMarginC.Width != null)
                    comment = (botMarginO.Width.Value != botMarginC.Width.Value) ? bot + Math.Round(Convert.ToDouble(botMarginC.Width.Value) / 567, 2).ToString() + " см" : "";
            }
            return comment;
        }
        // проверка левого поля ячейки
        public static string LeftMarginCheck()
        {
            object val = new object();
            string comment = "";
            string left = "; левого поля ячеек до ";
            StartMargin leftMarginO = new StartMargin(); // для проверяемого
            StartMargin leftMarginC = new StartMargin(); // для проверки

            leftMarginO = (tableMarCells.ChildElements.ToList().Exists(s => s.LocalName == "start")) ?
                tableMarCells.StartMargin : null;

            leftMarginC = (tableMarCellsToCompare.ChildElements.ToList().Exists(s => s.LocalName == "start")) ?
                tableMarCellsToCompare.StartMargin : null;
            if (leftMarginC != null)
            {
                if (leftMarginO.Width == null && leftMarginC.Width != null)
                    comment = left + Math.Round(Convert.ToDouble(leftMarginC.Width.Value) / 567, 2).ToString() + " см";
                else if (leftMarginO.Width != null && leftMarginC.Width != null)
                    comment = (leftMarginO.Width.Value != leftMarginC.Width.Value) ? left + Math.Round(Convert.ToDouble(leftMarginC.Width.Value) / 567, 2).ToString() + " см" : "";
            }
            return comment;
        }
        // проверка правого поля ячейки
        public static string RightMarginCheck()
        {
            object val = new object();
            string comment = "";
            string right = "; правого поля ячеек до ";
            EndMargin rightMarginO = new EndMargin(); // для проверяемого
            EndMargin rightMarginC = new EndMargin(); // для проверки

            rightMarginO = (tableMarCells.ChildElements.ToList().Exists(s => s.LocalName == "end")) ?
                tableMarCells.EndMargin : null;

            rightMarginC = (tableMarCellsToCompare.ChildElements.ToList().Exists(s => s.LocalName == "end")) ?
                tableMarCellsToCompare.EndMargin : null;
            if (rightMarginC != null)
            {
                if (rightMarginO.Width == null && rightMarginC.Width != null)
                    comment = right + Math.Round(Convert.ToDouble(rightMarginC.Width.Value) / 567, 2).ToString() + " см";
                else if (rightMarginO.Width != null && rightMarginC.Width != null)
                    comment = (rightMarginO.Width.Value != rightMarginC.Width.Value) ? right + Math.Round(Convert.ToDouble(rightMarginC.Width.Value) / 567, 2).ToString() + " см" : "";
            }
            return comment;
        }
    }
}
