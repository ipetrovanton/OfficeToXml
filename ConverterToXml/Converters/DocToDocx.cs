using System.IO;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib.WordprocessingML;
using b2xtranslator.WordprocessingMLMapping;
using static b2xtranslator.OpenXmlLib.OpenXmlPackage;

namespace ConverterToXml.Converters
{
    public class DocToDocx
    {
        public void ConvertToDocx(string docPath, string docxPath)
        {
            StructuredStorageReader reader = new StructuredStorageReader(docPath);
            WordDocument doc = new WordDocument(reader);
            WordprocessingDocument docx = WordprocessingDocument.Create(docxPath, DocumentType.Document);
            Converter.Convert(doc, docx);
            docx.Close();
        }

        public Stream ConvertToDocxMemoryStream(Stream stream)
        {
            StructuredStorageReader reader = new StructuredStorageReader(stream);
            WordDocument doc = new WordDocument(reader);
            var docx = WordprocessingDocument.Create("docx", DocumentType.Document);
            Converter.Convert(doc, docx);
            return new MemoryStream(docx.CloseWithoutSavingFile());
        }

        public Stream ConvertByFileToDocxMemoryStream(string docPath)
        {
            StructuredStorageReader reader = new StructuredStorageReader(docPath);
            WordDocument doc = new WordDocument(reader);
            var docx = WordprocessingDocument.Create("docx", DocumentType.Document);
            Converter.Convert(doc, docx);
            return new MemoryStream(docx.CloseWithoutSavingFile());
        }
    }
}
