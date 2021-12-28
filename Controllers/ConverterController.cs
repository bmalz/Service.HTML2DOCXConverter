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
using System.Text.RegularExpressions;

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
                body = await reader.ReadToEndAsync();

            if (string.IsNullOrWhiteSpace(body))
                return BadRequest();

            return new FileStreamResult(new MemoryStream(GenerateDocument(body)), "application/octet-stream");
        }

        [HttpPost("b64")]
        public async Task<IActionResult> PostBase64() {
             string body = string.Empty;

            using (var reader = new System.IO.StreamReader(Request.Body))
                body = await reader.ReadToEndAsync();

            if (string.IsNullOrWhiteSpace(body))
                return BadRequest();
                
            return Ok(Convert.ToBase64String(GenerateDocument(body)));
        }

        private byte[] GenerateDocument(string body) {
            string footer, header = string.Empty;
            header = GetHeader(ref body);
            footer = GetFooter(ref body);

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

                    ApplyFooter(package, footer);
                    ApplyHeader(package, header);

                    mainPart.Document.Save();
                }

                return generatedDocument.ToArray();
            }
        }

        private string GetFooter(ref string body) => GetContentFromTag("footer", ref body);

        private string GetHeader(ref string body) => GetContentFromTag("header", ref body);

        private string GetContentFromTag(string tag, ref string body)
        {
            string pattern = $@"(?:<{tag}>(?<content>(?:.*?\r?\n?)*)<\/{tag}>)+";
            RegexOptions options = RegexOptions.Multiline;
            Regex expression = new Regex(pattern, options);
            Match match = expression.Match(body);
            if (match.Success)
            {
                // Remove tag from body
                Regex regex = new Regex(pattern, options);
                body = regex.Replace(body, "");

                return match.Groups["content"].Value;
            }
            else
                return null;
        }

        private IList<OpenXmlCompositeElement> ConvertHtmlToOpenXml(string input)
        {
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
                    return converter.Parse(input);
                }
            }
        }

        private void ApplyHeader(WordprocessingDocument doc, string input)
        {
            if (!string.IsNullOrWhiteSpace(input))
            {
                MainDocumentPart mainDocPart = doc.MainDocumentPart;
                HeaderPart headerPart = mainDocPart.AddNewPart<HeaderPart>("r97");

                Header header = new Header();
                Paragraph paragraph = new Paragraph() { };
                Run run = new Run();
                run.Append(ConvertHtmlToOpenXml(input));
                paragraph.Append(run);
                header.Append(paragraph);
                headerPart.Header = header;

                SectionProperties sectionProperties = mainDocPart.Document.Body.Descendants<SectionProperties>().FirstOrDefault();
                if (sectionProperties == null)
                {
                    sectionProperties = new SectionProperties() { };
                    mainDocPart.Document.Body.Append(sectionProperties);
                }

                HeaderReference headerReference = new HeaderReference() { Type = DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default, Id = "r97" };
                sectionProperties.InsertAt(headerReference, 0);
            }
        }

        private void ApplyFooter(WordprocessingDocument doc, string input)
        {
            if (!string.IsNullOrWhiteSpace(input))
            {
                MainDocumentPart mainDocPart = doc.MainDocumentPart;
                FooterPart footerPart = mainDocPart.AddNewPart<FooterPart>("r98");

                Footer footer = new Footer();
                Paragraph paragraph = new Paragraph() { };
                Run run = new Run();
                run.Append(ConvertHtmlToOpenXml(input));
                paragraph.Append(run);
                footer.Append(paragraph);
                footerPart.Footer = footer;

                SectionProperties sectionProperties = mainDocPart.Document.Body.Descendants<SectionProperties>().FirstOrDefault();
                if (sectionProperties == null)
                {
                    sectionProperties = new SectionProperties() { };
                    mainDocPart.Document.Body.Append(sectionProperties);
                }

                FooterReference footerReference = new FooterReference() { Type = DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default, Id = "r98" };
                sectionProperties.InsertAt(footerReference, 0);
            }
        }
    }
}
