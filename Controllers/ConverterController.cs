using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using HtmlToOpenXml;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Service.HTML2DOCXConverter.Controllers
{
    [ApiController]
    [Route("convert")]
    public class ConverterController : ControllerBase
    {
        private readonly ILogger<ConverterController> _logger;

        public ConverterController(ILogger<ConverterController> logger)
        {
            _logger = logger;
        }

        [HttpPost]
        public async Task<IActionResult> Post()
        {
            string body = string.Empty;
            using (var reader = new System.IO.StreamReader(Request.Body))
            {
                body = await reader.ReadToEndAsync();
            }

            if (string.IsNullOrWhiteSpace(body))
                return BadRequest();

            using (MemoryStream generatedDocument = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = package.MainDocumentPart;
                    if (mainPart == null)
                    {
                        mainPart = package.AddMainDocumentPart();
                        new Document(new Body()).Save(mainPart);
                    }

                    HtmlConverter converter = new HtmlConverter(mainPart);
                    converter.ParseHtml(body);

                    mainPart.Document.Save();
                    
                }

                return new FileStreamResult(new MemoryStream(generatedDocument.ToArray()), "application/octet-stream");
            }
        }
    }
}
