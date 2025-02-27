using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.VariantTypes;




namespace FileMerge
{
    class Program
    {
        static void Main(string[] args)
        {
            string templatePath = Path.GetFullPath("./Files/Служебная записка.docm");
            string newDocFilePath = Path.GetFullPath("./Files/NewFile Служебная записка.docm");

            try
            {

                File.Delete(newDocFilePath);
                File.Copy(templatePath, newDocFilePath, true);

                using (WordprocessingDocument newDocFile = WordprocessingDocument.Open(newDocFilePath, true))
                {

                    /* Inserting data with DocProperty */

                    SetNewDataInDocProperty(newDocFile, "@RegData", "12.02.2004");
                    SetNewDataInDocProperty(newDocFile, "@RegNumber", "Регистрационный номер 1231232131");
                    /* Removing data*/

                    //RemoveTextFromBookmark(newDocFile, "RegDataBM");
                    //RemoveTextFromBookmark(newDocFile, "RegNumberBM");
                    //RemoveTextFromBookmark(newDocFile, "Stamp1BM");
                    //RemoveTextFromBookmark(newDocFile, "Stamp2BM");
                    //RemoveTextFromBookmark(newDocFile, "Stamp21BM");
                    //RemoveTextFromBookmark(newDocFile, "Stamp3BM");
                    //RemoveTextFromBookmark(newDocFile, "Stamp4BM");
                    //RemoveTextFromBookmark(newDocFile, "Stamp5BM");

                    /* Inserting data*/

                    //ReplaceTextFromBookMark(newDocFile, "RegDataBM", "99.99.9999" + " ");
                    //ReplaceTextFromBookMark(newDocFile, "RegNumberBM", " " + "9999999999999");
                    //ReplaceTextFromBookMarkInTable(newDocFile, "ShtampTableBM", "<TargetCell/>", "Stamp1BM", "1 ШТАМП");
                    //ReplaceTextFromBookMarkInTable(newDocFile, "ShtampTableBM", "<TargetCell/>", "Stamp2BM", "2 ШТАМП");
                    //ReplaceTextFromBookMarkInTable(newDocFile, "ShtampTableBM", "<TargetCell/>", "Stamp21BM", "2.1 ШТАМП");
                    //ReplaceTextFromBookMarkInTable(newDocFile, "ShtampTableBM", "<TargetCell/>", "Stamp3BM", "3 ШТАМП");
                    //ReplaceTextFromBookMarkInTable(newDocFile, "ShtampTableBM", "<TargetCell/>", "Stamp4BM", "4 ШТАМП");
                    //ReplaceTextFromBookMarkInTable(newDocFile, "ShtampTableBM", "<TargetCell/>", "Stamp5BM", "5 ШТАМП");


                    newDocFile.MainDocumentPart.Document.Save();

                }

                using (WordprocessingDocument newDocFile = WordprocessingDocument.Open(newDocFilePath, true))
                {
                    newDocFile.MainDocumentPart.Document.Save();
                }

                    Console.WriteLine("Данные успешно вставлены");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка: " + ex.Message);
            }

        }


        private static void SetNewDataInDocProperty(WordprocessingDocument mainDoc, string docPropertyName, string data)
        {
            var customPropsPart = mainDoc.CustomFilePropertiesPart;
            var properties = customPropsPart.Properties;
            foreach (var property in properties.Elements<CustomDocumentProperty>())
            {
                if (property.Name == docPropertyName)
                {
                    // Обновляем значение свойства
                    property.Elements<VTLPWSTR>().First().Text = data;
                    break;
                }
            }

        }



        private static void RemoveTextFromBookmark(WordprocessingDocument mainDoc, string bookmarkName)
        {
            var mainbody = mainDoc.MainDocumentPart.Document.Body;

            var bookmarkStart = mainbody.Descendants<BookmarkStart>().FirstOrDefault(b => b.Name == bookmarkName);
            var bookmarkEnd = mainbody.Descendants<BookmarkEnd>().FirstOrDefault(e => e.Id == bookmarkStart.Id);


            var elementsToRemove = bookmarkStart.ElementsAfter()
                                                .TakeWhile(e => e != bookmarkEnd)
                                                .ToList(); 

            foreach (var element in elementsToRemove)
            {
                element.Remove();
            }
        }

        private static void ReplaceTextFromBookMark(WordprocessingDocument mainDoc, string bookmarkName, string data)
        {
            var mainBody = mainDoc.MainDocumentPart.Document.Body;

            var bookmarkStart = mainBody.Descendants<BookmarkStart>().FirstOrDefault(b => b.Name == bookmarkName);
            var bookmarkEnd = mainBody.Descendants<BookmarkEnd>().FirstOrDefault(b => b.Id == bookmarkStart.Id);


            var elementsToRemove = new List<OpenXmlElement>();

            OpenXmlElement currentElement = bookmarkStart.NextSibling();
            while (currentElement != null && !ReferenceEquals(currentElement, bookmarkEnd))
            {
                if (currentElement is Run run)
                {
                    elementsToRemove.Add(run); // Добавляем только элементы <w:r>
                }
                currentElement = currentElement.NextSibling();
            }

            // Удаляем все найденные элементы <w:r>
            foreach (var element in elementsToRemove)
            {
                element.Remove();
            }

            // Создаем новый Run с данным текстом
            var newRun = new Run(new RunProperties (new Text(data)
            {
                Space = SpaceProcessingModeValues.Preserve
            }));

            // Вставляем новый Run после начала закладки
            bookmarkStart.Parent.InsertAfter(newRun, bookmarkStart);
        }


        private static void ReplaceTextFromBookMarkInTable(WordprocessingDocument mainDoc, string tableBookmarkName, string targetTableCell, string bookmarkName, string data)
        {
            var mainBody = mainDoc.MainDocumentPart.Document.Body;

            // Шаг 1: Найти BookmarkStart для таблицы
            var tableBookMark = mainBody.Descendants<BookmarkStart>()
                .FirstOrDefault(b => b.Name == tableBookmarkName);

            // Шаг 2: Найти таблицу, связанную с этим bookmark
            var targetTable = tableBookMark.Ancestors<Table>().FirstOrDefault();

            // Шаг 3: Найти целевую ячейку таблицы
            var targetCell = targetTable.Descendants<TableCell>()
                .FirstOrDefault(c => c.InnerText.Contains(targetTableCell));

            // Шаг 4: Найти BookmarkStart с именем 'bookmarkName' внутри ячейки
            var bookmarkStart = targetCell.Descendants<BookmarkStart>()
                .FirstOrDefault(b => b.Name == bookmarkName);

            // Шаг 5: Найти соответствующий BookmarkEnd
            var bookmarkEnd = targetCell.Descendants<BookmarkEnd>()
                .FirstOrDefault(b => b.Id == bookmarkStart.Id);

            var elementsToRemove = new List<OpenXmlElement>();

            OpenXmlElement currentElement = bookmarkStart.NextSibling();
            while (currentElement != null && !ReferenceEquals(currentElement, bookmarkEnd))
            {
                if (currentElement is Run run)
                {
                    elementsToRemove.Add(run); // Добавляем только элементы <w:r>
                }
                currentElement = currentElement.NextSibling();
            }

            // Удаляем все найденные элементы <w:r>
            foreach (var element in elementsToRemove)
            {
                element.Remove();
            }

            var newRun = new Run(
                new Text(data + " ")
                {
                    Space = SpaceProcessingModeValues.Preserve // Сохраняем пробелы в тексте
                });

            bookmarkStart.Parent.InsertAfter(newRun, bookmarkStart);

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





