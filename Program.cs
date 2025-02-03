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


namespace FileMerge
{
    class Program
    {
        static void Main(string[] args)
        {
            string templatePath = Path.GetFullPath("./Files/TemplateTestNumbering.docx");
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

                var abstractNumIdMap = new Dictionary<string, string>();
                var numIdMap = new Dictionary<int, int>();
                // Вставка данных в шаблон
                using (WordprocessingDocument newDocFile = WordprocessingDocument.Open(newDocFilePath, true))
                {

                    InsertContentAt(newDocFile, "<TagExplanatoryNote>", "<TagExplanatoryNote/>", file1Path, abstractNumIdMap, numIdMap);
                    InsertContentAt(newDocFile, "<DesignSolutionTag>", "<DesignSolutionTag/>", file2Path, abstractNumIdMap, numIdMap);

                    newDocFile.MainDocumentPart.Document.Save();

                }
                

                Console.WriteLine("Данные успешно вставлены");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка: " + ex.Message);
            }
        }

        private static void InsertContentAt(WordprocessingDocument mainDoc, string openTag, string closeTag, string filePath, Dictionary<string, string> abstractNumIdMap, Dictionary<int, int> numIdMap)
        {
            using (WordprocessingDocument insertedFilePath = WordprocessingDocument.Open(filePath, false))
            {
                var templateBody = mainDoc.MainDocumentPart.Document.Body;
                var sourceBody = insertedFilePath.MainDocumentPart.Document.Body;

                // Копируем стили из источника в целевой документ
                var sourceStylesPart = insertedFilePath.MainDocumentPart.StyleDefinitionsPart;
                var targetStylesPart = mainDoc.MainDocumentPart.StyleDefinitionsPart;
                if (sourceStylesPart != null && targetStylesPart != null)
                {
                    var sourceStyles = sourceStylesPart.Styles;
                    var targetStyles = targetStylesPart.Styles;
                    foreach (var style in sourceStyles.Elements<Style>())
                    {
                        if (!targetStyles.Elements<Style>().Any(s => s.StyleId == style.StyleId))
                        {
                            targetStyles.Append(style.CloneNode(true));
                        }
                    }
                }

                // Находим открывающий и закрывающий теги
                var startTag = templateBody.Descendants<Paragraph>()
                    .FirstOrDefault(p => p.InnerText.Contains(openTag));
                var endTag = templateBody.Descendants<Paragraph>()
                    .FirstOrDefault(p => p.InnerText.Contains(closeTag));

                var startIndex = templateBody.Elements().ToList().IndexOf(startTag);
                var endIndex = templateBody.Elements().ToList().IndexOf(endTag);

                // Обрабатываем нумерацию в источнике
                var numberingPart = insertedFilePath.MainDocumentPart.NumberingDefinitionsPart;
                if (numberingPart != null)
                {
                    var sourceNumbering = numberingPart.Numbering;
                    var localAbstractNumIdMap = new Dictionary<string, string>();
                    var localNumIdMap = new Dictionary<int, int>();

                    // Собираем все AbstractNum элементы
                    var abstractNumsToClone = sourceNumbering.Elements<AbstractNum>().ToList();
                    var clonedAbstractNums = new List<AbstractNum>();
                    foreach (var abstractNum in abstractNumsToClone)
                    {
                        // Клонируем текущий AbstractNum
                        var clonedAbstractNum = (AbstractNum)abstractNum.CloneNode(true);

                        // Генерируем новый уникальный идентификатор
                        int newAbstractNumId;
                        if (abstractNumIdMap.Any()) // Проверяем, содержит ли словарь элементы
                        {
                            // Находим максимальный идентификатор в словаре и увеличиваем на 1
                            newAbstractNumId = abstractNumIdMap.Values.Max(id => int.Parse(id)) + 1;
                            abstractNumIdMap.Clear();
                        }
                        else
                        {
                            // Если словарь пуст, начинаем с базового значения (например, текущего идентификатора + 1)
                            newAbstractNumId = abstractNum.AbstractNumberId + 1;
                        }

                        // Обновляем свойство AbstractNumberId у клонированного объекта
                        clonedAbstractNum.AbstractNumberId = newAbstractNumId;

                        // Добавляем отображение старого идентификатора в новый
                        localAbstractNumIdMap[abstractNum.AbstractNumberId] = newAbstractNumId.ToString();

                        // Добавляем только новое значение в abstractNumIdMap
                        abstractNumIdMap[abstractNum.AbstractNumberId] = newAbstractNumId.ToString();

                        // Добавляем клонированный объект в список
                        clonedAbstractNums.Add(clonedAbstractNum);

                    }

                    // Собираем все Num элементы
                    var numsToClone = sourceNumbering.Elements<NumberingInstance>().ToList();
                    var clonedNums = new List<NumberingInstance>();
                    foreach (var numInstance in numsToClone)
                    {
                        // Клонируем текущий NumberingInstance
                        var clonedNumInstance = (NumberingInstance)numInstance.CloneNode(true);

                        // Генерируем новый уникальный идентификатор
                        int newNumId;
                        if (numIdMap.Any()) // Проверяем, содержит ли словарь элементы
                        {
                            // Находим максимальный идентификатор в словаре и увеличиваем на 1
                            newNumId = numIdMap.Values.Max(id => id + 1);
                            numIdMap.Clear(); // Очищаем словарь перед добавлением нового значения
                        }
                        else
                        {
                            // Если словарь пуст, начинаем с базового значения (например, текущего идентификатора + 1)
                            newNumId = numInstance.NumberID.Value + 1;
                        }

                        // Обновляем свойство NumberID у клонированного объекта
                        clonedNumInstance.NumberID = new Int32Value(newNumId);

                        // Обновляем ссылки на AbstractNumId внутри клонированного NumberingInstance
                        foreach (var abstractNumId in clonedNumInstance.Elements<AbstractNumId>())
                        {
                            var oldAbstractNumId = abstractNumId.Val.Value.ToString();
                            if (localAbstractNumIdMap.TryGetValue(oldAbstractNumId, out string newAbstractNumId))
                            {
                                // Заменяем старый AbstractNumId на новый
                                abstractNumId.Val = new Int32Value(int.Parse(newAbstractNumId));
                            }
                        }

                        // Добавляем отображение старого идентификатора в новый
                        localNumIdMap[numInstance.NumberID.Value] = newNumId;

                        // Добавляем только новое значение в numIdMap
                        numIdMap[numInstance.NumberID.Value] = newNumId;

                        // Добавляем клонированный объект в список
                        clonedNums.Add(clonedNumInstance);

                    }

                    // Объединяем локальные словари с общими словарями
                    foreach (var kvp in localAbstractNumIdMap)
                    {
                        abstractNumIdMap[kvp.Key] = kvp.Value;
                    }
                    foreach (var kvp in localNumIdMap)
                    {
                        numIdMap[kvp.Key] = kvp.Value;
                    }
                    // Получаем существующее NumberingDefinitionsPart или создаем новый, если его нет
                    var mainNumberingPart = mainDoc.MainDocumentPart.NumberingDefinitionsPart ?? mainDoc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                    var mainNumbering = new Numbering();

                    // Если существует существующий Numbering, копируем его элементы в новый экземпляр
                    if (mainNumberingPart.Numbering != null)
                    {
                        foreach (var existingAbstractNum in mainNumberingPart.Numbering.Elements<AbstractNum>())
                        {
                            mainNumbering.Append(existingAbstractNum.CloneNode(true));
                        }
                        foreach (var existingNum in mainNumberingPart.Numbering.Elements<NumberingInstance>())
                        {
                            mainNumbering.Append(existingNum.CloneNode(true));
                        }
                    }

                    // Добавляем все AbstractNum элементы в новый документ
                    foreach (var clonedAbstractNum in clonedAbstractNums)
                    {
                        mainNumbering.Append(clonedAbstractNum);
                    }

                    // Добавляем все Num элементы в новый документ
                    foreach (var clonedNum in clonedNums)
                    {
                        mainNumbering.Append(clonedNum);
                    }

                    // Сортируем элементы в правильном порядке
                    SortNumberingElements(mainNumbering);

                    // Сохраняем изменения в NumberingDefinitionsPart
                    mainNumberingPart.Numbering = mainNumbering;

                }
                UpdateNumPrReferences(mainDoc, numIdMap);
                // Вставляем элементы из источника
                var elementsToInsert = sourceBody.Elements().ToList();
                elementsToInsert.Reverse();
                foreach (var element in elementsToInsert)
                {
                    var clonedElement = element.CloneNode(true);
                    templateBody.InsertAfter(clonedElement, startTag);
                }
            }
        }
        private static void SortNumberingElements(Numbering numbering)
        {
            if (numbering == null) return;

            // Собираем все AbstractNum и Num элементы
            var abstractNums = numbering.Elements<AbstractNum>().ToList();
            var nums = numbering.Elements<NumberingInstance>().ToList();

            // Очищаем существующие элементы
            numbering.RemoveAllChildren<AbstractNum>();
            numbering.RemoveAllChildren<NumberingInstance>();

            // Добавляем все AbstractNum элементы в правильном порядке
            foreach (var abstractNum in abstractNums)
            {
                numbering.Append(abstractNum);
            }

            // Добавляем все Num элементы в правильном порядке
            foreach (var num in nums)
            {
                numbering.Append(num);
            }
        }
        private static int GetUniqueNumId(WordprocessingDocument doc)
        {
            var numberingPart = doc.MainDocumentPart.NumberingDefinitionsPart;
            if (numberingPart == null || numberingPart.Numbering == null)
            {
                return 1;
            }

            int maxNumId = 0;
            foreach (var numInstance in numberingPart.Numbering.Elements<NumberingInstance>())
            {
                if (numInstance.NumberID.HasValue && numInstance.NumberID.Value > maxNumId)
                {
                    maxNumId = numInstance.NumberID.Value;
                }
            }
            return maxNumId + 1;
        }

        private static int GetUniqueAbstractNumId(WordprocessingDocument doc)
        {
            var numberingPart = doc.MainDocumentPart.NumberingDefinitionsPart;
            if (numberingPart == null || numberingPart.Numbering == null)
            {
                return 1;
            }

            int maxAbstractNumId = 0;
            foreach (var abstractNum in numberingPart.Numbering.Elements<AbstractNum>())
            {
                if (int.TryParse(abstractNum.AbstractNumberId, out int currentAbstractNumId))
                {
                    if (currentAbstractNumId > maxAbstractNumId)
                    {
                        maxAbstractNumId = currentAbstractNumId;
                    }
                }
            }
            return maxAbstractNumId + 1;
        }

        private static void AddToNumbering(WordprocessingDocument doc, OpenXmlElement element)
        {
            var numberingPart = doc.MainDocumentPart.NumberingDefinitionsPart;
            if (numberingPart == null)
            {
                numberingPart = doc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                var numbering = new Numbering();
                numberingPart.Numbering = numbering;
            }

            numberingPart.Numbering.Append(element);
        }

        private static void UpdateNumPrReferences(WordprocessingDocument mainDoc, Dictionary<int, int> numIdMap)
        {
            var body = mainDoc.MainDocumentPart.Document.Body;

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                var paragraphProperties = paragraph.Elements<ParagraphProperties>().FirstOrDefault();
                if (paragraphProperties != null)
                {
                    var numPr = paragraphProperties.Elements<NumberingProperties>().FirstOrDefault();
                    if (numPr != null)
                    {
                        var numId = numPr.Elements<NumberingId>().FirstOrDefault();

                        if (numId != null)
                        {
                            var oldNumId = numId.Val.Value;
                            if (numIdMap.TryGetValue(oldNumId, out int newNumId))
                            {
                                numId.Val = new Int32Value(newNumId);
                                Console.WriteLine($"Updated numId from {oldNumId} to {newNumId}");
                            }
                            else
                            {
                                Console.WriteLine($"Не удалось найти новое значение для старого numId: {oldNumId}");
                            }
                        }
                    }
                }
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

//ReplacePlaceholders(newDocFile, "ToWho", "Наблюдательный совет НКО НКЦ(АО)");
//ReplacePlaceholders(newDocFile, "WhoQuestion", "О согласовании неаудиторских услуг");
//ReplacePlaceholders(newDocFile, "WhoSpeaker", "Коковин Сергей Игоревич");
//ReplacePlaceholders(newDocFile, "OnReview", "КА НКЦ 24.12.24");
//ReplacePlaceholders(newDocFile, "QuestName", "КА НКЦ 24.12.24");
//PlaceHyperLink(newDocFile, "QuestNameBM", url);
//GenerateAgreementTable(newDocFile, "<AgreementTableTag/>", "AgreementTableBM", tableInsertingData);