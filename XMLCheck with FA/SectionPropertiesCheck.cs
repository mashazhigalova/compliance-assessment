using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComplianceAssessment
{
    // класс для проверки свойств раздела
    class SectionPropertiesCheck: BaseConcepts
    {
        public void SectionProps(MainDocumentPart docPart, SectionProperties sectProps, Paragraph pCom)
        {
            List<Paragraph> comPars = new List<Paragraph>();
            comPars.Add(PageMarginCheck(sectProps));
            comPars.Add(PageSizeCheck(sectProps));
            comPars.Add(TitlePageCheck(sectProps));
            comPars.Add(PageNumTypeCheck(sectProps));
            comPars.Add(CheckNums(docPart));
            comPars.RemoveAll(s => s == null);
            if (comPars.Count != 0)
                AddCommentOnParagraph(docPart, new List<Paragraph> { pCom }, comPars);
        }
        // проверка полей страницы
        private static Paragraph PageMarginCheck(SectionProperties sectProps)
        {
            string comment = "";
            PageMargin pageMarginO = new PageMargin(); // для проверяемого
            PageMargin pageMarginC = new PageMargin(); // для проверки
            
            pageMarginO = sectProps.Descendants<PageMargin>().First();
            pageMarginC = sectionPropsToCompare.Descendants<PageMargin>().First();

            if (pageMarginO.Bottom == null && pageMarginC.Bottom != null)
                comment += "нижнее поле до " + Math.Round(Convert.ToDouble(pageMarginC.Bottom.Value) / 567, 2).ToString() + " см";
            else if (pageMarginO.Bottom != null && pageMarginC.Bottom != null)
                comment += (pageMarginO.Bottom.Value != pageMarginC.Bottom.Value) ? "нижнее поле до " + Math.Round(Convert.ToDouble(pageMarginC.Bottom.Value) / 567, 2).ToString() + " см" : "";

            if (pageMarginO.Top == null && pageMarginC.Top != null)
                comment += "; верхнее поле до " + Math.Round(Convert.ToDouble(pageMarginC.Bottom.Value) / 567, 2).ToString() + " см";
            else if (pageMarginO.Top != null && pageMarginC.Top != null)
                comment += (pageMarginO.Top.Value != pageMarginC.Top.Value) ? "; верхнее поле до " + Math.Round(Convert.ToDouble(pageMarginC.Top.Value) / 567, 2).ToString() + " см" : "";

            if (pageMarginO.Left == null && pageMarginC.Left != null)
                comment += "; левое поле до " + Math.Round(Convert.ToDouble(pageMarginC.Left.Value) / 567, 2).ToString() + " см";
            else if (pageMarginO.Left != null && pageMarginC.Left != null)
                comment += (pageMarginO.Left.Value != pageMarginC.Left.Value) ? "; левое поле до " + Math.Round(Convert.ToDouble(pageMarginC.Left.Value) / 567, 2).ToString() + " см" : "";

            if (pageMarginO.Right == null && pageMarginC.Right != null)
                comment += "; правое поле до " + Math.Round(Convert.ToDouble(pageMarginC.Right.Value) / 567, 2).ToString() + " см";
            else if (pageMarginO.Right != null && pageMarginC.Right != null)
                comment += (pageMarginO.Right.Value != pageMarginC.Right.Value) ? "; правое поле до " + Math.Round(Convert.ToDouble(pageMarginC.Right.Value) / 567, 2).ToString() + " см" : "";

            if (pageMarginO.Header == null && pageMarginC.Header != null)
                comment += "; верхний колонтитул " + Math.Round(Convert.ToDouble(pageMarginC.Header.Value) / 567, 2).ToString() + " см";
            else if (pageMarginO.Header != null && pageMarginC.Header != null)
                comment += (pageMarginO.Header.Value != pageMarginC.Header.Value) ? "; верхний колонтитул " + Math.Round(Convert.ToDouble(pageMarginC.Header.Value) / 567, 2).ToString() + " см" : "";

            if (pageMarginO.Footer == null && pageMarginC.Footer != null)
                comment += "; нижний колонтитул " + Math.Round(Convert.ToDouble(pageMarginC.Footer.Value) / 567, 2).ToString() + " см";
            else if (pageMarginO.Footer != null && pageMarginC.Footer != null)
                comment += (pageMarginO.Footer.Value != pageMarginC.Footer.Value) ? "; нижний колонтитул " + Math.Round(Convert.ToDouble(pageMarginC.Footer.Value) / 567, 2).ToString() + " см" : "";

            if (comment.StartsWith("; "))
                comment = comment.Remove(0, 2);
            return (comment != "") ? new Paragraph(new Run(new Text("изменить поля страницы: " + comment + ". "))) : null;
        }
        // проверка размера страницы
        private static Paragraph PageSizeCheck(SectionProperties sectProps)
        {
            string comment = "";
            PageSize pageSizeO = new PageSize(); // для проверяемого
            PageSize pageSizeC = new PageSize(); // для проверки

            pageSizeO = sectProps.Descendants<PageSize>().First();
            pageSizeC = sectionPropsToCompare.Descendants<PageSize>().First();

            if (pageSizeO.Width == null && pageSizeC.Width != null)
                comment += "ширина - " + Math.Round(Convert.ToDouble(pageSizeC.Width.Value) / 567, 2).ToString() + " см";
            else if (pageSizeO.Width != null && pageSizeC.Width != null)
                comment += (pageSizeO.Width.Value != pageSizeC.Width.Value) ? "ширина - " + Math.Round(Convert.ToDouble(pageSizeC.Width.Value) / 567, 2).ToString() + " см" : "";

            if (pageSizeO.Height == null && pageSizeC.Height != null)
                comment += "; высота - " + Math.Round(Convert.ToDouble(pageSizeC.Height.Value) / 567, 2).ToString() + " см";
            else if (pageSizeO.Height != null && pageSizeC.Height != null)
                comment += (pageSizeO.Height.Value != pageSizeC.Height.Value) ? "; высота - " + Math.Round(Convert.ToDouble(pageSizeC.Height.Value) / 567, 2).ToString() + " см" : "";
            
            if (comment.StartsWith("; "))
                comment = comment.Remove(0, 2);
            return (comment != "") ? new Paragraph(new Run(new Text("изменить размеры страницы: " + comment + ". "))) : null;
        }

        // проверка на то, должна ли первая страница иметь особый колонтитул
        private static Paragraph TitlePageCheck(SectionProperties sectProps)
        {
            string com = "";
            TitlePage tPageO = new TitlePage(); // для проверяемого
            TitlePage tPageC = new TitlePage(); // для проверки
            string apply = "применить особый колонтитул первой страницы";
            string remove = "убрать особый колонтитул первой страницы";

            tPageO = (sectProps.ChildElements.ToList().Exists(s => s.LocalName == "titlePg")) ?
                sectProps.Descendants<TitlePage>().First() : null;
            tPageC = (sectionPropsToCompare.ChildElements.ToList().Exists(s => s.LocalName == "titlePg")) ?
                sectionPropsToCompare.Descendants<TitlePage>().First() : null;

            if (tPageO == null && tPageC != null)
            {
                if (tPageC.Val != null)
                    com = (tPageC.Val.Value == true) ? apply : "";
                else com = apply;
            }
            else if (tPageO != null && tPageC == null)
            {
                if (tPageO.Val != null)
                    com = (tPageO.Val.Value == true) ? remove : "";
                else com = remove;
            }
            else if (tPageO != null && tPageC != null)
                if (tPageO.Val != null && tPageC.Val != null)
                {
                    com = (tPageC.Val.Value == true && tPageO.Val.Value == false) ? apply : "";
                    com = (tPageC.Val.Value == false && tPageO.Val.Value == true) ? remove : "";
                }
                else if (tPageO.Val == null && tPageC.Val != null)
                {
                    com = (tPageC.Val.Value == false) ? remove : apply;
                }
                else if (tPageO.Val != null && tPageC.Val == null)
                {
                    com = (tPageO.Val.Value == true) ? remove : "";
                }
            return (com != "") ? new Paragraph(new Run(new Text(com))) : null;
        }
        // проверка формата номера и начала отсчета
        private static Paragraph PageNumTypeCheck(SectionProperties sectProps)
        {
            string comment = "";
            PageNumberType tPageO = new PageNumberType(); // для проверяемого
            PageNumberType tPageC = new PageNumberType(); // для проверки
            string applyFormatFromCompare = "";
            string applyStartFromCompare = "";

            tPageO = (sectProps.ChildElements.ToList().Exists(s => s.LocalName == "pgNumType")) ?
                sectProps.Descendants<PageNumberType>().First() : null;
            tPageC = (sectionPropsToCompare.ChildElements.ToList().Exists(s => s.LocalName == "pgNumType")) ?
                sectionPropsToCompare.Descendants<PageNumberType>().First() : null;

            if (tPageC != null)
            {
                if (tPageC.Format != null)
                    applyFormatFromCompare = "формат нумерации - " + ParagraphDicts.numFormatDict[tPageC.Format.Value.ToString()];
                if (tPageC.Start != null)
                    applyStartFromCompare = "; начать нумерацию с " + tPageC.Start.Value.ToString() + "-ого значения";

                if (tPageO == null)
                {
                    comment += applyFormatFromCompare + applyStartFromCompare;

                }
                else if (tPageO != null)
                {
                    if (tPageO.Format != null && tPageC.Format != null)
                        comment += (tPageO.Format.Value != tPageO.Format.Value) ? applyFormatFromCompare : "";
                    else if (tPageO.Format == null && tPageC.Format != null)
                        comment += applyFormatFromCompare;

                    if (tPageO.Start != null && tPageC.Start != null)
                        comment += (tPageO.Start.Value != tPageO.Start.Value) ? applyStartFromCompare : "";
                    else if (tPageO.Format == null && tPageC.Format != null)
                        comment += applyStartFromCompare;

                }
            }
            if (comment.StartsWith("; "))
                comment = comment.Remove(0, 2);
            return (comment != "") ? new Paragraph(new Run(new Text("изменить параметры нумерации страниц: " + comment))) : null;
        }

        // проверка положения номера на странице
        private static Paragraph CheckNums(MainDocumentPart docPart)
        {
            Footer footer;
            string placeNumOriginal = "";
            string just= "";
            if (docPart.FooterParts != null) // если в проверяемом документе есть нижний колонтитул
            {
                for (int i = 0; i < docPart.FooterParts.Count(); i++)
                {
                    footer = docPart.FooterParts.ElementAt(i).Footer; // первый нижний колонтитул
                    if (footer.ChildElements.ToList().Exists(s => s.LocalName == "sdt"))
                    {
                        placeNumOriginal = footer.Descendants<SdtContentDocPartObject>().First().DocPartGallery.Val.Value.ToString();
                        just = ParagraphDicts.justificationDict[footer.Descendants<ParagraphProperties>().First().Justification.Val.Value.ToString()];
                    }
                }
            }
            Header header;
            if (docPart.FooterParts != null)
            {
                for (int i = 0; i < docPart.HeaderParts.Count(); i++)
                {
                    header = docPart.HeaderParts.ElementAt(i).Header;
                    if (header.ChildElements.ToList().Exists(s => s.LocalName == "sdt"))
                    {
                        // получение положения нумерации в проверяемом документе
                        placeNumOriginal = header.Descendants<SdtContentDocPartObject>().First().DocPartGallery.Val.Value.ToString();
                        just = ParagraphDicts.justificationDict[header.Descendants<ParagraphProperties>().First().Justification.Val.Value.ToString()];
                    }
                }
            }

            string comment = "";
            if (placeNum != placeNumOriginal) // если положение номера не совпадает
                if (placeNum != "")
                    placeNum = (placeNum.Contains("Bottom")) ? "номер внизу страницы" : "номер вверху страницы";
            just = (just != justToCompare) ? "; выравнивание по " + justToCompare : "";
            comment = placeNum + just;
            if (comment.StartsWith("; "))
                comment = comment.Remove(0, 2);
            return (comment != "") ? new Paragraph(new Run(new Text(comment))) : null;
        }

    }
}
