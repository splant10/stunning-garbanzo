/*
 * Retrieved by Spencer Plant December 13 2015
 * from Brock Nusser at http://stackoverflow.com/questions/83152/reading-pdf-documents-in-net
 */

using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace TextToSpeech
{
    public static class PdfTextGetter
    {
        public static string getPdfText(string path)
        {
            PdfReader reader = new PdfReader(path);
            string text = string.Empty;
            for (int page = 1; page <= reader.NumberOfPages; page++)
            {
                text += PdfTextExtractor.GetTextFromPage(reader, page);
            }
            reader.Close();
            return text;
        }
    }
}
