namespace pdf_to_word_converter.IServices
{
    public interface IPdfToWordService
    {
        Task<byte[]> ConvertPdfToWordAsync(IFormFile pdfFile);
    }
}
