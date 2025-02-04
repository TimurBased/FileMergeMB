using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using GemBox.Document.Tracking;
using System.Reflection;
using System.Security.Policy;
using System.Windows.Threading;
using System.Threading;


namespace FileMerge
{
    class Program
    {
        static void Main(string[] args)
        {
            string templatePath = Path.GetFullPath("./Files/TemplateFile.docx");
            string newDocFilePath = Path.GetFullPath("./Files/NewFile.docx");
            string file1Path = Path.GetFullPath("./Files/File_1_testWithNumbering.docx");
            string file2Path = Path.GetFullPath("./Files/File_2_testWithNumbering.docx");

            // Вставляемые данные в таблицу
            List<string[]> tableInsertingData= new List<string[]>{
                new string[] { "Арискин Алексей Сергеевич", "Юридический департамент", "Согласовано c комментариями", "Комментарий123456" },
                new string[] { "Сейфут Тимур Маратович", "Юридический департамент", "Согласовано", "" },
                new string[] { "Ивановов Ивааааан Иванович", "Юридический департамент", "Согласовано c комментариями", "Комментарий123456 Комментарий123456 Комментарий123456 Комментарий123456" }
                };
            string url = "https://learn.javascript.ru/cookie";
            try
            {
                
                File.Delete(newDocFilePath);
                File.Copy(templatePath, newDocFilePath, true);

                var fileMap = new Dictionary<string, string>() { 
                    { "<TagExplanatoryNote>", file1Path },
                    { "<DesignSolutionTag>", file2Path }
                };
                // Вставка данных в шаблон
                Console.WriteLine("0) " + System.GC.GetTotalMemory(true).ToString("000,000,000", Thread.CurrentThread.CurrentCulture));
                Console.WriteLine("Start");
                Console.WriteLine("1) " + System.GC.GetTotalMemory(true).ToString("000,000,000", Thread.CurrentThread.CurrentCulture));
                using (WordprocessingDocument newDocFile = WordprocessingDocument.Open(newDocFilePath, true))
                {

                    InsertFileContentAt(newDocFile, fileMap);
                    ReplacePlaceholders(newDocFile, "ToWho", "Наблюдательный совет НКО НКЦ(АО)");
                    ReplacePlaceholders(newDocFile, "WhoQuestion", "О согласовании неаудиторских услуг");
                    ReplacePlaceholders(newDocFile, "WhoSpeaker", "Коковин Сергей Игоревич");
                    ReplacePlaceholders(newDocFile, "OnReview", "КА НКЦ 24.12.24");
                    ReplacePlaceholders(newDocFile, "QuestName", "КА НКЦ 24.12.24");
                    PlaceHyperLink(newDocFile, "QuestNameBM", url);
                    GenerateAgreementTable(newDocFile, "<AgreementTableTag/>", "AgreementTableBM", tableInsertingData);
                    newDocFile.MainDocumentPart.Document.Save();

                }
                Console.WriteLine("2) " + System.GC.GetTotalMemory(true).ToString("000,000,000", Thread.CurrentThread.CurrentCulture));
                Console.WriteLine("End");
                Console.WriteLine("3) " + System.GC.GetTotalMemory(true).ToString("000,000,000", Thread.CurrentThread.CurrentCulture));

                Console.WriteLine("Данные успешно вставлены");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка: " + ex.Message);
            }
            
        }

        private static void InsertFileContentAt(WordprocessingDocument mainDoc, Dictionary<string, string> fileMap)
        {
            MainDocumentPart mainPart = mainDoc.MainDocumentPart;
            var templateBody = mainPart.Document.Body;

            // Перебираем все пары ключ-значение в словаре
            foreach (var kvp in fileMap)
            {
                string tag = kvp.Key;
                string filePath = kvp.Value; 

                // Находим открывающий тег в документе
                var startTag = templateBody.Descendants<Paragraph>()
                    .FirstOrDefault(p => p.InnerText.Trim() == tag);

                int startIndex = templateBody.Elements().ToList().IndexOf(startTag);

                // Генерируем уникальный идентификатор для AltChunk
                string altChunkId = "AltChunkId" + kvp.Key.Substring(1, kvp.Key.Length - 2);

                // Добавляем часть документа как AlternativeFormatImportPart
                AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(
                    AlternativeFormatImportPartType.WordprocessingML, altChunkId);

                using (FileStream fileStream = File.Open(filePath, FileMode.Open))
                {
                    chunk.FeedData(fileStream);
                }

                // Создаем AltChunk элемент и связываем его с идентификатором
                AltChunk altChunk = new AltChunk();
                altChunk.Id = altChunkId;

                // Вставляем AltChunk после найденного тега
                templateBody.InsertAfter(altChunk, startTag);

            }
        }

        public static void ReplacePlaceholders(WordprocessingDocument mainDoc, string targetText, string replacementText)
        {
            var templateBody = mainDoc.MainDocumentPart.Document.Body;
            var TextCollection = templateBody.Descendants<Text>().ToList();

            foreach (var item in TextCollection)
            {
                if (item.Text.Contains(targetText))
                {
                    item.Text = item.Text.Replace(targetText, replacementText);
                    break;
                }
            }
        }

        public static void InsertContentText(WordprocessingDocument mainDoc, string openTag, List<string> replacementText)
        {
            var templateBody = mainDoc.MainDocumentPart.Document.Body;

            // Находим открывающий тег
            var startTag = templateBody.Descendants<Paragraph>()
                .FirstOrDefault(p => p.InnerText.Contains(openTag));

            // Используем StringBuilder для объединения текста
            var sb = new StringBuilder();
            foreach (var text in replacementText)
            {
                if (sb.Length > 0)
                {
                    sb.Append(", "); // Разделитель между текстовыми элементами
                }
                sb.Append(text);
            }

            // Создаем новый абзац и Run с объединённым текстом
            var newRun = new Run(new Text(sb.ToString()));
            var newParagraph = new Paragraph(newRun);

            // Вставляем новый абзац сразу после открывающего тега
            templateBody.InsertAfter(newParagraph, startTag);
        }

        private static void GenerateAgreementTable(WordprocessingDocument mainDoc, string openTag, string agremntTableBookmark, List<string[]> rowData) 
        {
            var mainBody = mainDoc.MainDocumentPart.Document.Body;

            var bookMark = mainBody.Descendants<BookmarkStart>().FirstOrDefault(b => b.Name == agremntTableBookmark);

            var targetTable = bookMark.Ancestors<Table>().FirstOrDefault();

            var startTag = targetTable.Descendants<Paragraph>().FirstOrDefault(p => p.InnerText.Contains(openTag));

            var targetCell = targetTable.Descendants<TableCell>().FirstOrDefault(c => c.InnerText.Contains(openTag));

            var agreementTable = targetCell.Descendants<Table>().FirstOrDefault();

            foreach (var row in rowData)
            {
                var newRow = new TableRow();

                foreach (var item in row)
                {
                    TableCell newCell = new TableCell(
                                            new TableCellProperties(
                                            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                                        ),
                                        new Paragraph(new ParagraphProperties(new SpacingBetweenLines() { After = "0", Before = "0" },
                                                                              new Indentation { FirstLine = "0" },
                                                                              new Justification() { Val = JustificationValues.Center }
                                                                              ),
                                        new Run(new RunProperties() { FontSize = new FontSize() { Val = "16" }, RunFonts = new RunFonts() { HighAnsi = "Times New Roman", Ascii = "Times New Roman", ComplexScript = "Times New Roman" } },
                                        new Text(item))));
                    newRow.AppendChild(newCell);
                }
                agreementTable.AppendChild(newRow);
            }
        }
        private static void PlaceHyperLink(WordprocessingDocument mainDoc, string bookmarkName, string url)
        {
            var mainBody = mainDoc.MainDocumentPart.Document.Body;

            var bookmarkStart = mainBody.Descendants<BookmarkStart>().FirstOrDefault(b => b.Name == bookmarkName);
            var bookmarkEnd = mainBody.Descendants<BookmarkEnd>().FirstOrDefault(b => b.Id == bookmarkStart.Id);
            var parentParagraph = bookmarkStart.Ancestors<Paragraph>().FirstOrDefault();

            var runsBetweenBookmarks = bookmarkStart
                .ElementsAfter() 
                .TakeWhile(e => e != bookmarkEnd)
                .OfType<Run>() 
                .ToList();

            var mainPart = mainDoc.MainDocumentPart;
            var relationshipId = "rId" + Guid.NewGuid().ToString("N");
            mainPart.AddHyperlinkRelationship(new Uri(url), true, relationshipId);

            var hyperlink = new Hyperlink() { Id = relationshipId };

            var hyperlinkStyle = new RunProperties(
                new RunStyle { Val = "Hyperlink"  },
                new Color { Val = "#0000EE" }
            );

            foreach (var run in runsBetweenBookmarks)
            {
                var runProps = run.GetFirstChild<RunProperties>();
                if (runProps == null)
                {
                    runProps = new RunProperties();
                    run.PrependChild(runProps);
                }

                runProps.Append(hyperlinkStyle.CloneNode(true));
                run.Remove();
                hyperlink.Append(run);
            }

            parentParagraph.AppendChild(hyperlink);
        }




    }
}





