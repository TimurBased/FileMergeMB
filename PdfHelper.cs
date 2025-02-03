using System;
using System.IO;
using Spire.Doc;

namespace FileMerge
{
    public static class PdfHelper
    {
        public static void ConvertToPdf(string pathToDocxFile)
        {
            // Загрузка DOCX файла
            Document document = new Document();
            document.LoadFromFile(pathToDocxFile);

            // Сохранение в формате PDF
            document.SaveToFile(pathToDocxFile, FileFormat.PDF);
        }

        public static void ConvertToDocX(string pathToDocFile)
        {
            Document document = new Document();

            document.LoadFromFile(pathToDocFile);

            string outputPath = Path.ChangeExtension(pathToDocFile, ".docx");

            document.SaveToFile(outputPath, FileFormat.Docx);

        }
    }
}
