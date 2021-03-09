extern alias DocIOWPF;
extern alias DocBase;
using DocIOWPF::Syncfusion.DocIO;
using DocIOWPF::Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;

using DocBase::Syncfusion.Pdf;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using SyncfusionAPI.Data;


// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace SyncfusionAPI.Controllers
{
    [Route("[controller]/[action]")]
    [ApiController]
    public class PdfConvertorAPIController : ControllerBase
    {
       
        private readonly IConfiguration configurationString;
      
        private readonly SqlFileStreamDbContext _EDRMScontext;

        public PdfConvertorAPIController( IConfiguration config, SqlFileStreamDbContext EDRMScontext )
        {
           
            configurationString = config;
            _EDRMScontext = EDRMScontext;

        }


        // EDRMS Pdf Convertor   GET api/<PdfConvertorAPIController>/5
        [HttpGet("{id}")]
        public IActionResult PdfConvert(int id)
        {
            var getFile = _EDRMScontext.tbl_FileStream_Data.Where(x => x.ID_Case == id).FirstOrDefault();
            byte[] pdf = getFile.Data;

            MemoryStream pdfStream2 = new MemoryStream(pdf);

            //Open using Syncfusion
            WordDocument document = new WordDocument(pdfStream2, FormatType.Automatic);
            //document.FontSettings.SubstituteFont += SubstitueFont;
            pdfStream2.Dispose();
            pdfStream2 = null;
            // Creates a new instance of DocIORenderer class.
            DocToPDFConverter render = new DocToPDFConverter();
            // Converts Word document into PDF document.
            PdfDocument pdffile = render.ConvertToPDF(document);
            MemoryStream memoryStream = new MemoryStream();
            // Save the PDF document.                    
            //Save using Syncfusion
            pdffile.Save(memoryStream);
            memoryStream.Position = 0;

            pdffile.Close(true);
            string contentType2 = "application/pdf";
            return File(memoryStream, contentType2);
        }















    }
}






//// GET: api/<PdfConvertorAPIController>
//[HttpGet]
//public IEnumerable<string> Get()
//{
//    return new string[] { "value1", "value2" };
//}

//// POST api/<PdfConvertorAPIController>
//[HttpPost]
//public void Post([FromBody] string value)
//{
//}

//// PUT api/<PdfConvertorAPIController>/5
//[HttpPut("{id}")]
//public void Put(int id, [FromBody] string value)
//{
//}

//// DELETE api/<PdfConvertorAPIController>/5
//[HttpDelete("{id}")]
//public void Delete(int id)
//{
//}