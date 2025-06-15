using DevExpress.Utils;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf;
using pdf_to_word_converter.IServices;
using Syncfusion.Pdf.Parsing;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace pdf_to_word_converter.Services
{

    public class PdfToWordService : IPdfToWordService
    {
        private readonly ILogger<PdfToWordService> _logger;

        public PdfToWordService(ILogger<PdfToWordService> logger)
        {
            _logger = logger;
        }

        public async Task<byte[]> ConvertPdfToWordAsync(IFormFile pdfFile)
        {
            try
            {
                using var pdfStream = pdfFile.OpenReadStream();
                using var memoryStream = new MemoryStream();

                _logger.LogInformation($"Starting PDF to DOCX conversion for: {pdfFile.FileName}");

                // Extract text from PDF
                var extractedPages = await ExtractTextFromPdfAsync(pdfStream);

                // Create Word document directly (skip HTML step)
                CreateWordDocumentFromPages(memoryStream, extractedPages, pdfFile.FileName);

                _logger.LogInformation("Successfully converted PDF to DOCX");

                return memoryStream.ToArray();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in PDF to Word conversion for file: {FileName}", pdfFile.FileName);
                throw;
            }
        }

        private async Task<List<PageContent>> ExtractTextFromPdfAsync(Stream pdfStream)
        {
            return await Task.Run(() =>
            {
                var pages = new List<PageContent>();

                using (var pdfReader = new PdfReader(pdfStream))
                using (var pdfDocument = new PdfDocument(pdfReader))
                {
                    for (int pageNum = 1; pageNum <= pdfDocument.GetNumberOfPages(); pageNum++)
                    {
                        try
                        {
                            var page = pdfDocument.GetPage(pageNum);
                            var strategy = new LocationTextExtractionStrategy();
                            var pageText = PdfTextExtractor.GetTextFromPage(page, strategy);

                            if (!string.IsNullOrWhiteSpace(pageText))
                            {
                                var pageContent = new PageContent
                                {
                                    PageNumber = pageNum,
                                    RawText = pageText,
                                    ProcessedText = ProcessExtractedText(pageText)
                                };

                                pages.Add(pageContent);
                                _logger.LogDebug($"Extracted {pageText.Length} characters from page {pageNum}");
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, $"Failed to extract text from page {pageNum}");
                        }
                    }
                }

                return pages;
            });
        }

        private string ProcessExtractedText(string rawText)
        {
            if (string.IsNullOrWhiteSpace(rawText))
                return string.Empty;

            var processed = rawText;

            // Clean up common PDF extraction issues
            processed = Regex.Replace(processed, @"\s+", " ");
            processed = Regex.Replace(processed, @"(?<!\.)\n(?![A-Z])", " ");
            processed = Regex.Replace(processed, @"\n{3,}", "\n\n");
            processed = processed.Trim();

            // Create proper paragraphs
            processed = Regex.Replace(processed, @"([.!?])\s*\n\s*([A-Z])", "$1\n\n$2");

            return processed;
        }

        private void CreateWordDocumentFromPages(Stream stream, List<PageContent> pages, string originalFileName)
        {
            using (var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                // Add document title
                AddTitle(body, $"Converted from: {Path.GetFileNameWithoutExtension(originalFileName)}");

                // Add conversion info
                AddParagraph(body, $"Converted on: {DateTime.Now:yyyy-MM-dd HH:mm:ss}", "18", true, true);
                AddParagraph(body, $"Total pages: {pages.Count}", "18", true, true);

                // Add horizontal line
                AddHorizontalRule(body);

                // Process each page
                foreach (var page in pages)
                {
                    // Add page header
                    AddHeading(body, $"Page {page.PageNumber}", "20");

                    // Split content into paragraphs and headings
                    var contentElements = AnalyzeAndSplitContent(page.ProcessedText);

                    foreach (var element in contentElements)
                    {
                        if (element.IsHeading)
                        {
                            AddHeading(body, element.Text, "16");
                        }
                        else
                        {
                            AddParagraph(body, element.Text, "22", false, false);
                        }
                    }

                    // Add page break except for last page
                    if (page.PageNumber < pages.Count)
                    {
                        AddPageBreak(body);
                    }
                }

                mainPart.Document.Save();
            }
        }

        private List<ContentElement> AnalyzeAndSplitContent(string text)
        {
            var elements = new List<ContentElement>();

            if (string.IsNullOrWhiteSpace(text))
                return elements;

            var paragraphs = text.Split(new[] { "\n\n", "\r\n\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var paragraph in paragraphs)
            {
                var cleanParagraph = paragraph.Trim();
                if (!string.IsNullOrWhiteSpace(cleanParagraph))
                {
                    elements.Add(new ContentElement
                    {
                        Text = cleanParagraph,
                        IsHeading = IsLikelyHeading(cleanParagraph)
                    });
                }
            }

            return elements;
        }

        private bool IsLikelyHeading(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var trimmed = text.Trim();

            // Check if it's short and doesn't end with a period
            return trimmed.Length < 100 &&
                   !trimmed.EndsWith(".") &&
                   !trimmed.EndsWith(",") &&
                   trimmed.Split(' ').Length <= 8 &&
                   char.IsUpper(trimmed[0]) &&
                   // Additional checks for common heading patterns
                   (trimmed.Contains(":") ||
                    Regex.IsMatch(trimmed, @"^[A-Z][A-Z\s]+$") || // ALL CAPS
                    Regex.IsMatch(trimmed, @"^\d+\.") || // Numbered
                    trimmed.Length < 50); // Short lines are likely headings
        }

        private void AddTitle(Body body, string text)
        {
            var paragraph = new Paragraph();
            var paragraphProperties = new ParagraphProperties();
            paragraphProperties.AppendChild(new Justification() { Val = JustificationValues.Center });
            paragraph.AppendChild(paragraphProperties);

            var run = new Run();
            var runProperties = new RunProperties();
            runProperties.AppendChild(new Bold());
            runProperties.AppendChild(new FontSize() { Val = "28" });
            runProperties.AppendChild(new Color() { Val = "2c3e50" });
            run.AppendChild(runProperties);
            run.AppendChild(new Text(text));

            paragraph.AppendChild(run);
            body.AppendChild(paragraph);
        }

        private void AddHeading(Body body, string text, string fontSize)
        {
            var paragraph = new Paragraph();
            var run = new Run();

            var runProperties = new RunProperties();
            runProperties.AppendChild(new Bold());
            runProperties.AppendChild(new FontSize() { Val = fontSize });
            runProperties.AppendChild(new Color() { Val = "34495e" });
            run.AppendChild(runProperties);
            run.AppendChild(new Text(text));

            paragraph.AppendChild(run);
            body.AppendChild(paragraph);
        }

        private void AddParagraph(Body body, string text, string fontSize = "22", bool isItalic = false, bool isCentered = false)
        {
            var paragraph = new Paragraph();

            if (isCentered)
            {
                var paragraphProperties = new ParagraphProperties();
                paragraphProperties.AppendChild(new Justification() { Val = JustificationValues.Center });
                paragraph.AppendChild(paragraphProperties);
            }

            var run = new Run();
            var runProperties = new RunProperties();
            runProperties.AppendChild(new FontSize() { Val = fontSize });

            if (isItalic)
            {
                runProperties.AppendChild(new Italic());
                runProperties.AppendChild(new Color() { Val = "7f8c8d" });
            }

            run.AppendChild(runProperties);
            run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

            paragraph.AppendChild(run);
            body.AppendChild(paragraph);
        }

        private void AddHorizontalRule(Body body)
        {
            var paragraph = new Paragraph();
            var paragraphProperties = new ParagraphProperties();

            var paragraphBorders = new ParagraphBorders();
            paragraphBorders.AppendChild(new BottomBorder()
            {
                Val = BorderValues.Single,
                Size = 6,
                Color = "bdc3c7"
            });

            paragraphProperties.AppendChild(paragraphBorders);
            paragraph.AppendChild(paragraphProperties);
            body.AppendChild(paragraph);
        }

        private void AddPageBreak(Body body)
        {
            var paragraph = new Paragraph();
            var run = new Run();
            run.AppendChild(new Break() { Type = BreakValues.Page });
            paragraph.AppendChild(run);
            body.AppendChild(paragraph);
        }
    }

    public class PageContent
    {
        public int PageNumber { get; set; }
        public string RawText { get; set; } = string.Empty;
        public string ProcessedText { get; set; } = string.Empty;
    }

    public class ContentElement
    {
        public string Text { get; set; } = string.Empty;
        public bool IsHeading { get; set; }
    }
}
