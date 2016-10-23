using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System.IO;

namespace ComplianceAssessment
{
    public class Check
    {
        /// <summary>
        /// Выполнение проверки документа
        /// </summary>
        /// <param name="fileName">Имя файла для проверки</param>
        /// <param name="fileNameCompare">Имя файла с правилами оформления</param>
        /// <param name="pathOutDocs">Путь к выходным файлам</param>
        /// <param name="pathUpload">Путь к загруженным файлам</param>
        /// <param name="pathRules">Путь к правилам оформления</param>
        public static void Do(string fileName, string fileNameCompare, string pathOutDocs, string pathUpload, string pathRules)
        {
            // удаление всех старых результирующих файлов 
            ClearFolder(pathOutDocs);
            
            // получение полного пути к загруженному файлу
            pathUpload = pathUpload + fileName;
            // получение полного пути к правилам оформления
            fileNameCompare = pathRules + fileNameCompare;
            // открытие загруженного документа для проведения проверки форматирования
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(pathUpload, true))
            {
                // упрощение разметки текстового документа
                SimplifyMarkupSettings markupSettings = new SimplifyMarkupSettings
                {
                    RemoveComments = true,
                    RemoveEndAndFootNotes = true,
                    RemoveLastRenderedPageBreak = true,
                    RemovePermissions = true,
                    RemoveProof = true,
                    RemoveRsidInfo = true,
                    RemoveSmartTags = true,
                    RemoveSoftHyphens = true,
                };
                MarkupSimplifier.SimplifyMarkup(wDoc, markupSettings);

                // перенос всех параметров форматирования на самый высокий уровень - прямое форматирование
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
                MainDocumentPart mainPart = wDoc.MainDocumentPart;
                ParRunCheck a = new ParRunCheck();
                // выполнение проверки параметров документа
                a.CheckStyleProps(mainPart, fileNameCompare);
            }
            // определение имени результирующего исправленного документа
            string fullPathEdit = pathUpload.Replace(".docx", " - ComplianceAssessment.docx");
            if (!File.Exists(fullPathEdit))
                // копирование загруженного документа для исправления
                File.Copy(pathUpload, fullPathEdit);
            // перемещение проверенного файла с комментариями в OutDocs
            File.Copy(pathUpload, pathOutDocs + fileName);
            string newFullPathEdit = fullPathEdit.Replace("Upload", "OutDocs");
            File.Copy(fullPathEdit, newFullPathEdit);
            // открытие документа для его исправления
            using (WordprocessingDocument wEditedDoc = WordprocessingDocument.Open(newFullPathEdit, true))
            {
                // упрощение разметки документа
                SimplifyMarkupSettings markupSettings = new SimplifyMarkupSettings
                {
                    RemoveComments = true,
                };
                MarkupSimplifier.SimplifyMarkup(wEditedDoc, markupSettings);

                MainDocumentPart mainPart = wEditedDoc.MainDocumentPart;
                EditDocumentClass edit = new EditDocumentClass();
                // исправление документа в соответсвии с правилами оформления
                edit.Edit(mainPart, fileNameCompare);
                mainPart.Document.Save();
            }
        }

        // удаление всех файлов по указанному пути
        public static void ClearFolder(string path)
        {
            DirectoryInfo di = new DirectoryInfo(path);
            foreach (var file in di.GetFiles())
                file.Delete();
        }
    }
}
