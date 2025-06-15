using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using pdf_to_word_converter.IServices;

namespace pdf_to_word_converter.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ConversionController : ControllerBase
    {

        private readonly IPdfToWordService _conversionService;
        private readonly ILogger<ConversionController> _logger;

        public ConversionController(IPdfToWordService conversionService, ILogger<ConversionController> logger)
        {
            _conversionService = conversionService;
            _logger = logger;
        }

        [HttpPost("convert")]
        public async Task<IActionResult> ConvertPdfToWord(IFormFile pdfFile)
        {
            try
            {
                // Validation checks
                if (pdfFile == null || pdfFile.Length == 0)
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = "Please select a PDF file to upload.",
                        error = "NO_FILE_SELECTED"
                    });
                }

                // Check file type
                if (!pdfFile.ContentType.Equals("application/pdf", StringComparison.OrdinalIgnoreCase))
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = "Only PDF files are allowed.",
                        error = "INVALID_FILE_TYPE"
                    });
                }

                // Check file size (10MB limit)
                const long maxFileSize = 10 * 1024 * 1024; // 10MB
                if (pdfFile.Length > maxFileSize)
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = "File size must be less than 10MB.",
                        error = "FILE_TOO_LARGE"
                    });
                }

                _logger.LogInformation($"Starting conversion for file: {pdfFile.FileName}, Size: {pdfFile.Length} bytes");

                // Call the service to convert PDF to Word
                var wordFileBytes = await _conversionService.ConvertPdfToWordAsync(pdfFile);

                // Generate output filename
                var originalFileName = Path.GetFileNameWithoutExtension(pdfFile.FileName);
                var outputFileName = $"{originalFileName}_converted.docx";

                _logger.LogInformation($"Conversion completed successfully for: {pdfFile.FileName}");

                // Return the Word file as download
                return File(
                    wordFileBytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    outputFileName
                );
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error converting PDF to Word for file: {FileName}", pdfFile?.FileName ?? "Unknown");

                return StatusCode(500, new
                {
                    success = false,
                    message = "An error occurred during conversion. Please try again.",
                    error = "CONVERSION_FAILED"
                });
            }
        }

        [HttpGet("health")]
        public IActionResult Health()
        {
            return Ok(new
            {
                status = "healthy",
                timestamp = DateTime.UtcNow,
                service = "PDF to Word Converter API",
                version = "1.0.0"
            });
        }

        [HttpGet("info")]
        public IActionResult GetInfo()
        {
            return Ok(new
            {
                service = "PDF to Word Converter",
                version = "1.0.0",
                supportedFormats = new
                {
                    input = new[] { "PDF" },
                    output = new[] { "DOCX" }
                },
                limits = new
                {
                    maxFileSize = "10MB",
                    supportedTypes = new[] { "application/pdf" }
                },
                endpoints = new
                {
                    convert = "/api/conversion/convert",
                    health = "/api/conversion/health",
                    info = "/api/conversion/info"
                }
            });
        }
    }
}


