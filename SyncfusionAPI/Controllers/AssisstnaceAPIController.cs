
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System.Collections;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Syncfusion.EJ2.Linq;
using Microsoft.AspNetCore.Mvc.Rendering;
using System.IO;
using SyncfusionAPI.Data;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Pdf;
using Syncfusion.DocIORenderer;
//using TuranCore.Extensions;
//using static TuranCore.Models.Enums;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;






// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace SyncfusionAPI.Controllers
{
    [Route("[controller]/[action]")]
    [ApiController]
    public class AssisstnaceAPIController : ControllerBase
    {

        private readonly IConfiguration configurationString;


        private readonly ApplicationDbContext _context;

        public AssisstnaceAPIController(IConfiguration config, ApplicationDbContext context)
        {

            _context = context;
            configurationString = config;
        }

        // https://localhost:44342/AssisstnaceAPI/PrintHospital/414-00124356/1816
        // EDRMS Pdf Convertor   GET api/<PdfConvertorAPIController>/5

        [HttpGet("{IndividualID}/{RequestID}")]
        public IActionResult PrintHospital(string IndividualID, int RequestID)
        {


            var Fa_HO = _context.FA_Hospitals.Where(x => x.RequestID == RequestID).FirstOrDefault();

            var IndividualIDM = new SqlParameter("@thiscasenumber", IndividualID);
            var Assishopital = _context.printdetails.FromSqlRaw("exec sp_PrintDetailsAllDe @thiscasenumber", IndividualIDM).AsNoTracking().AsEnumerable().FirstOrDefault();
            byte[] sp_imageData = Assishopital.Photo;

            //User_Log lg = new User_Log();
            //lg.Log_date = DateTime.Now;
            //lg.User_Name = HttpContext.GetProGresUserName();
            //lg.CaseNumber = Fa_HO.CaseNumber;
            //lg.Action = "Print Hospital Assistance- IndividualID: " + IndividualID;
            //_context.User_Log.Add(lg);
            //_context.SaveChanges();

            // Creating a new document.
            WordDocument document = new WordDocument();
            //Adding a new section to the document.
            WSection section = document.AddSection() as WSection;
            //Set Margin of the section
            section.PageSetup.Margins.All = 72;
            //Set page size of the section
            section.PageSetup.PageSize = new Syncfusion.Drawing.SizeF(612, 792);
            //Create Paragraph styles
            WParagraphStyle style = document.AddParagraphStyle("Normal") as WParagraphStyle;
            style.CharacterFormat.FontName = "Calibri";
            style.CharacterFormat.FontSize = 11f;
            style.ParagraphFormat.BeforeSpacing = 0;
            style.ParagraphFormat.AfterSpacing = 8;
            style.ParagraphFormat.LineSpacing = 13.9f;

            style = document.AddParagraphStyle("Heading 1") as WParagraphStyle;
            style.ApplyBaseStyle("Normal");
            style.CharacterFormat.FontName = "Calibri Light";
            style.CharacterFormat.FontSize = 16f;
            style.CharacterFormat.TextColor = Syncfusion.Drawing.Color.FromArgb(46, 116, 181);
            style.ParagraphFormat.BeforeSpacing = 12;
            style.ParagraphFormat.AfterSpacing = 0;
            style.ParagraphFormat.Keep = true;
            style.ParagraphFormat.KeepFollow = true;
            style.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
            IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();

            // Gets the image stream.
            FileStream imageStream = new FileStream("wwwroot/images/unhcrlogo.jpg", FileMode.Open, FileAccess.Read);
            IWPicture picture = paragraph.AppendPicture(imageStream);
            picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
            picture.VerticalOrigin = VerticalOrigin.Margin;
            picture.VerticalPosition = -45;
            picture.HorizontalOrigin = HorizontalOrigin.Column;
            picture.HorizontalPosition = 0.5f;
            picture.WidthScale = 70;
            picture.HeightScale = 55;



            FileStream imageStream25 = new FileStream("wwwroot/images/apwletter.jpg", FileMode.Open, FileAccess.Read);
            IWPicture picture5 = paragraph.AppendPicture(imageStream25);
            picture5.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
            picture5.VerticalOrigin = VerticalOrigin.Margin;
            //YUKARi almak icin
            //picture5.VerticalPosition = 30;
            //picture5.HorizontalOrigin = HorizontalOrigin.Column;
            // picture5.HorizontalPosition = 103;
            picture5.WidthScale = 85;
            picture5.HeightScale = 85;
            //45-55
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            // paragraph.ParagraphFormat.FirstLineIndent = 36;
            WTextRange textRange = paragraph.AppendText("") as WTextRange;
            textRange.CharacterFormat.FontSize = 11f;
            textRange.CharacterFormat.FontName = "Calibri";

            textRange.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;

            //Appends paragraph.
            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Heading 1");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("BMMYK Birleşmış Milletler Mülteciler Yüksek Komiserliği Sancak Mah. Tiflis Cad. 552. Sokak No3 06550 AnkaraTel: (312) 409 7000  Fax(312) 441 2173  Email turan@unhcr.org") as WTextRange;
            textRange.CharacterFormat.FontSize = 10f;
            textRange.CharacterFormat.FontName = "Calibri";
            textRange.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;


            IWTable table2 = section.AddTable();
            table2.ResetCells(1, 2);
            table2.TableFormat.Borders.BorderType = BorderStyle.None;
            table2.TableFormat.IsAutoResized = true;


            paragraph = table2[0, 0].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            textRange = paragraph.AppendText("BMMYK Dosya numarası:" + Fa_HO.CaseNumber + "\n" + "Seri No: ( " + Fa_HO.RequestID + " )") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;


            paragraph = table2[0, 1].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            textRange = paragraph.AppendText("Veriliş tarihi:  " + DateTime.Now.ToString("dd/mm/yyyy") + "\r") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;




            paragraph = section.AddParagraph();
            paragraph.ParagraphFormat.FirstLineIndent = 0;
            paragraph.BreakCharacterFormat.FontSize = 14f;
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("\n" + Fa_HO.Hospital + "\n") as WTextRange;
            textRange.CharacterFormat.FontSize = 14f;



            if (Fa_HO.Nationality == "IRQ")
            {
                Fa_HO.Nationality = "Irak";
            }
            else if (Fa_HO.Nationality == "IRN")
            {
                Fa_HO.Nationality = "Iran";
            }
            else if (Fa_HO.Nationality == "AFG")
            {
                Fa_HO.Nationality = "Afgan";
            }
            else if (Fa_HO.Nationality == "SOM")
            {
                Fa_HO.Nationality = "Somali";
            }
            else if (Fa_HO.Nationality == "SUD")
            {
                Fa_HO.Nationality = "Sudan";
            }
            else if (Fa_HO.Nationality == "KEN")
            {
                Fa_HO.Nationality = "Kenyan";
            }
            else if (Fa_HO.Nationality == "PAL")
            {
                Fa_HO.Nationality = "Filistin";
            }
            else if (Fa_HO.Nationality == "YEM")
            {
                Fa_HO.Nationality = "Yemen  ";
            }
            else if (Fa_HO.Nationality == "ETH")
            {
                Fa_HO.Nationality = "Etiyopya";
            }
            else if (Fa_HO.Nationality == "UZB")
            {
                Fa_HO.Nationality = "Uzbek";
            }
            else if (Fa_HO.Nationality == "CHI")
            {
                Fa_HO.Nationality = "Cin";
            }
            else if (Fa_HO.Nationality == "KGZ")
            {
                Fa_HO.Nationality = "Kyrgyz";
            }
            else if (Fa_HO.Nationality == "CMR")
            {
                Fa_HO.Nationality = "Cameroon";
            }
            else if (Fa_HO.Nationality == "ICO")
            {
                Fa_HO.Nationality = "Ivorian";
            }
            else if (Fa_HO.Nationality == "ICO")
            {
                Fa_HO.Nationality = "Ivorian";
            }
            else if (Fa_HO.Nationality == "LBR")
            {
                Fa_HO.Nationality = "Filipin";
            }

            //Appends paragraph.
            paragraph = section.AddParagraph();
            // paragraph.ParagraphFormat.FirstLineIndent = 10;
            paragraph.BreakCharacterFormat.FontSize = 12f;
            textRange = paragraph.AppendText(Fa_HO.Nationality + " uyruklu " + Fa_HO.GivenName + " " + Fa_HO.FamilyName + " " + Fa_HO.CaseNumber + " dosya numarası ile Birleşmiş Milletler Mülteciler Yüksek Komiserliği’nde kayıtlıdır.Hastanenizin ilgili bölümünde tahakkuk edecek " + Fa_HO.Amount + " ( " + Fa_HO.AmountInWritten + " ) " + "Türk Lirası’na kadar olan tetkik ve tedavi giderleri Kurumumuz tarafından karşılanacaktır. Bu miktarı geçecek tedavi giderleri ve hastane dışındaki tetkikler için Kurumumuzdan onay alınması gerekmektedir.Onay alınmadan yapılan harcamalardan sorumlu olmayacağımızı belirtmek isteriz. Faturanızın, orijinal mektubumuz ile birlikte adresimize gönderilmesini rica ederiz.\n\n\n") as WTextRange;
            textRange.CharacterFormat.FontSize = 12f;


            paragraph = section.AddParagraph();
            //  paragraph.ParagraphFormat.FirstLineIndent = 10;
            paragraph.BreakCharacterFormat.FontSize = 12f;
            textRange = paragraph.AppendText("BM Mülteciler Yüksek Komiserliği\n") as WTextRange;
            textRange.CharacterFormat.FontSize = 12f;
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;

            section.AddParagraph();

            //Instantiation of DocIORenderer for Word to PDF conversion
            DocIORenderer render = new DocIORenderer();
            //Converts Word document into PDF document
            PdfDocument pdfDocument = render.ConvertToPDF(document);
            //Releases all resources used by the Word document and DocIO Renderer objects
            render.Dispose();
            document.Dispose();
            //Save the document into stream.
            MemoryStream stream = new MemoryStream();
            pdfDocument.Save(stream);
            stream.Position = 0;
            //pdfDocument.Save("DoctoPDF.pdf", Response, HttpReadType.Open);

            //Close the documents.
            pdfDocument.Close(true);
            //Defining the ContentType for pdf file.
            string contentType = "application/pdf";
            //Define the file name.
            string fileName = Fa_HO.CaseNumber+".pdf";

            return File(stream, contentType, fileName);

        }


        // https://localhost:44342/AssisstnaceAPI/PrintAccomodation/414-00124356/14258
        [HttpGet("{IndividualID}/{RequestID}")]
        public IActionResult PrintAccomodation(string IndividualID, int RequestID)
        {

            //User_Log lg = new User_Log();
            //lg.Log_date = DateTime.Now;
            //    lg.User_Name =HttpContext.GetProGresUserName();
            //    lg.CaseNumber = Fa_Ac.CaseNumber;
            //    lg.Action = "Print Accomodation Assistance- IndividualID: " + IndividualID;
            //    _context.User_Log.Add(lg);
            //    _context.SaveChanges();


            var Fa_Ac = _context.FA_Accomodation.Where(x => x.RequestID == RequestID).FirstOrDefault();

            var IndividualIDM = new SqlParameter("@thiscasenumber", IndividualID);


            // Creating a new document.
            WordDocument document = new WordDocument();
            //Adding a new section to the document.
            WSection section = document.AddSection() as WSection;
            //Set Margin of the section
            section.PageSetup.Margins.All = 72;
            //Set page size of the section
            section.PageSetup.PageSize = new Syncfusion.Drawing.SizeF(612, 792);
            //Create Paragraph styles
            WParagraphStyle style = document.AddParagraphStyle("Normal") as WParagraphStyle;
            style.CharacterFormat.FontName = "Calibri";
            style.CharacterFormat.FontSize = 11f;
            style.ParagraphFormat.BeforeSpacing = 0;
            style.ParagraphFormat.AfterSpacing = 8;
            style.ParagraphFormat.LineSpacing = 13.9f;

            style = document.AddParagraphStyle("Heading 1") as WParagraphStyle;
            style.ApplyBaseStyle("Normal");
            style.CharacterFormat.FontName = "Calibri Light";
            style.CharacterFormat.FontSize = 16f;
            style.CharacterFormat.TextColor = Syncfusion.Drawing.Color.FromArgb(46, 116, 181);
            style.ParagraphFormat.BeforeSpacing = 12;
            style.ParagraphFormat.AfterSpacing = 0;
            style.ParagraphFormat.Keep = true;
            style.ParagraphFormat.KeepFollow = true;
            style.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
            IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();

            // Gets the image stream.
            FileStream imageStream = new FileStream("wwwroot/images/unhcrlogo.jpg", FileMode.Open, FileAccess.Read);
            IWPicture picture = paragraph.AppendPicture(imageStream);
            picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
            picture.VerticalOrigin = VerticalOrigin.Margin;
            picture.VerticalPosition = -45;
            picture.HorizontalOrigin = HorizontalOrigin.Column;
            picture.HorizontalPosition = 0.5f;
            picture.WidthScale = 70;
            picture.HeightScale = 55;



            FileStream imageStream25 = new FileStream("wwwroot/images/apwletter.jpg", FileMode.Open, FileAccess.Read);
            IWPicture picture5 = paragraph.AppendPicture(imageStream25);
            picture5.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
            picture5.VerticalOrigin = VerticalOrigin.Margin;
            //YUKARi almak icin
            //picture5.VerticalPosition = 30;
            //picture5.HorizontalOrigin = HorizontalOrigin.Column;
            //picture5.HorizontalPosition = 103;
            picture5.WidthScale = 85;
            picture5.HeightScale = 85;

            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            // paragraph.ParagraphFormat.FirstLineIndent = 36;
            WTextRange textRange = paragraph.AppendText("") as WTextRange;
            textRange.CharacterFormat.FontSize = 11f;
            textRange.CharacterFormat.FontName = "Calibri";

            textRange.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;

            //Appends paragraph.
            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Heading 1");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("BMMYK Birleşmış Milletler Mülteciler Yüksek Komiserliği Sancak Mah. Tiflis Cad. 552. Sokak No3 06550 AnkaraTel: (312) 409 7000  Fax(312) 441 2173  Email turan@unhcr.org") as WTextRange;
            textRange.CharacterFormat.FontSize = 10f;
            textRange.CharacterFormat.FontName = "Calibri";
            textRange.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;


            IWTable table2 = section.AddTable();
            table2.ResetCells(1, 2);
            table2.TableFormat.Borders.BorderType = BorderStyle.None;
            table2.TableFormat.IsAutoResized = true;


            paragraph = table2[0, 0].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            textRange = paragraph.AppendText("BMMYK Dosya numarası:" + Fa_Ac.CaseNumber + "\n" + "Seri No: ( " + Fa_Ac.RequestID + " )") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;


            paragraph = table2[0, 1].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            textRange = paragraph.AppendText("Veriliş tarihi:  " + DateTime.Now.ToString("dd/mm/yyyy") + "\r") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;




            paragraph = section.AddParagraph();
            paragraph.ParagraphFormat.FirstLineIndent = 0;
            paragraph.BreakCharacterFormat.FontSize = 14f;
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("\n" + Fa_Ac.HotelName + "\n") as WTextRange;
            textRange.CharacterFormat.FontSize = 14f;



            DateTime GStartingDate = Convert.ToDateTime(Fa_Ac.StartingDate);
            DateTime GEndingDate = Convert.ToDateTime(Fa_Ac.EndingDate);

            if (Fa_Ac.Nationality == "IRQ")
            {
                Fa_Ac.Nationality = "Irak";
            }
            else if (Fa_Ac.Nationality == "IRN")
            {
                Fa_Ac.Nationality = "Iran";
            }
            else if (Fa_Ac.Nationality == "AFG")
            {
                Fa_Ac.Nationality = "Afgan";
            }
            else if (Fa_Ac.Nationality == "SOM")
            {
                Fa_Ac.Nationality = "Somali";
            }
            else if (Fa_Ac.Nationality == "SUD")
            {
                Fa_Ac.Nationality = "Sudan";
            }
            else if (Fa_Ac.Nationality == "KEN")
            {
                Fa_Ac.Nationality = "Kenyan";
            }
            else if (Fa_Ac.Nationality == "PAL")
            {
                Fa_Ac.Nationality = "Filistin";
            }
            else if (Fa_Ac.Nationality == "YEM")
            {
                Fa_Ac.Nationality = "Yemen  ";
            }
            else if (Fa_Ac.Nationality == "ETH")
            {
                Fa_Ac.Nationality = "Etiyopya";
            }
            else if (Fa_Ac.Nationality == "UZB")
            {
                Fa_Ac.Nationality = "Uzbek";
            }
            else if (Fa_Ac.Nationality == "CHI")
            {
                Fa_Ac.Nationality = "Cin";
            }
            else if (Fa_Ac.Nationality == "KGZ")
            {
                Fa_Ac.Nationality = "Kyrgyz";
            }
            else if (Fa_Ac.Nationality == "CMR")
            {
                Fa_Ac.Nationality = "Cameroon";
            }
            else if (Fa_Ac.Nationality == "ICO")
            {
                Fa_Ac.Nationality = "Ivorian";
            }
            else if (Fa_Ac.Nationality == "ICO")
            {
                Fa_Ac.Nationality = "Ivorian";
            }
            else if (Fa_Ac.Nationality == "LBR")
            {
                Fa_Ac.Nationality = "Filipin";
            }

            //Appends paragraph.
            paragraph = section.AddParagraph();
            // paragraph.ParagraphFormat.FirstLineIndent = 10;
            paragraph.BreakCharacterFormat.FontSize = 12f;
            textRange = paragraph.AppendText("Aşağıda kimlik bilgileri bulunan " + Fa_Ac.Nationality + " uyruklu kişi(ler)  " + GStartingDate.ToString("dd/MM/yyyy") + " - " + GEndingDate.ToString("dd/MM/yyyy") + " tarihlerinde otelinizde kalacaktır.  Belirtilen hizmetlerin ücreti Kurumumuz tarafından karşılanacaktır. Faturanızın, orijinal mektubumuz ile birlikte adresimize gönderilmesini rica ederiz.\n") as WTextRange;
            textRange.CharacterFormat.FontSize = 12f;


            IWTable table6 = section.AddTable();
            table6.ResetCells(8, 2);
            table6.TableFormat.Borders.BorderType = BorderStyle.Hairline;
            table6.TableFormat.IsAutoResized = true;

            //row 1
            paragraph = table6[0, 0].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("Dosya numarası") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;


            paragraph = table6[0, 1].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText(" " + Fa_Ac.CaseNumber + " ") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;




            //second row
            paragraph = table6[1, 0].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText(" Adı / Soyadı ") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;


            paragraph = table6[1, 1].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText(" " + Fa_Ac.GivenName + " " + Fa_Ac.FamilyName + " ") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;




            //thered row
            paragraph = table6[2, 0].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText(" Uyruğu ") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;


            paragraph = table6[2, 1].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText(" " + Fa_Ac.Nationality + " ") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;


            //4
            paragraph = table6[3, 0].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("Sağlanacak hizmetler") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;


            paragraph = table6[3, 1].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("Konaklama") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;

            //5
            paragraph = table6[4, 0].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("Kişi sayısı") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;


            paragraph = table6[4, 1].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("" + Fa_Ac.NumberOfPersons) as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;

            //row 6
            paragraph = table6[5, 0].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("Başlangıç tarihi") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;


            paragraph = table6[5, 1].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText(" " + GStartingDate.ToString("dd/MM/yyyy") + " ") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;

            //row 7
            paragraph = table6[6, 0].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText(" Gideceği şehir ") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;


            paragraph = table6[6, 1].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText(" " + GEndingDate.ToString("dd/MM/yyyy") + " ") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;

            //row 8
            paragraph = table6[7, 0].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("Konaklama süresi") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;


            paragraph = table6[7, 1].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("" + Fa_Ac.TotalDays) as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;





            paragraph = section.AddParagraph();
            //  paragraph.ParagraphFormat.FirstLineIndent = 10;
            paragraph.BreakCharacterFormat.FontSize = 10f;
            textRange = paragraph.AppendText("\nBM Mülteciler Yüksek Komiserliği\n") as WTextRange;
            textRange.CharacterFormat.FontSize = 10f;
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;







            paragraph = section.AddParagraph();
            //  paragraph.ParagraphFormat.FirstLineIndent = 10;
            paragraph.BreakCharacterFormat.FontSize = 11f;
            textRange = paragraph.AppendText(Fa_Ac.HotelName + "\r") as WTextRange;
            textRange = paragraph.AppendText(Fa_Ac.HotelAddress + "\r") as WTextRange;
            textRange = paragraph.AppendText("Tel: 311 10 85\r") as WTextRange;
            textRange.CharacterFormat.FontSize = 11f;
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;






            section.AddParagraph();
            //Instantiation of DocIORenderer for Word to PDF conversion
            DocIORenderer render = new DocIORenderer();
            //Converts Word document into PDF document
            PdfDocument pdfDocument = render.ConvertToPDF(document);
            //Releases all resources used by the Word document and DocIO Renderer objects
            render.Dispose();
            document.Dispose();
            //Save the document into stream.
            MemoryStream stream = new MemoryStream();
            pdfDocument.Save(stream);
            stream.Position = 0;
            //pdfDocument.Save("DoctoPDF.pdf", Response, HttpReadType.Open);


            //Close the documents.
            pdfDocument.Close(true);
            //Defining the ContentType for pdf file.
            string contentType = "application/pdf";
            //Define the file name.
            string fileName = Fa_Ac.CaseNumber+".pdf";


            //Creates a FileContentResult object by using the file contents, content type, and file name.
            return File(stream, contentType, fileName);
            ////Saves the Word document to  MemoryStream
            //MemoryStream stream = new MemoryStream();
            //document.Save(stream, FormatType.Docx);
            //stream.Position = 0;



            ////Opens an existing document from stream through constructor of `WordDocument` class
            //FileStream fileStreamPath = new FileStream(@"Sample.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            ////Creates an empty Word document instance
            ////Loads or opens an existing word document through Open method of WordDocument class
            //document.Open(fileStreamPath, FormatType.Automatic);


            ////Download Word document in the browser
            //return File(stream, "application/msword", "Sample.docx");
        }

        [HttpGet("{IndividualID}/{RequestID}")]
        public IActionResult PrintTransportation(string IndividualID, int RequestID)
        {
            var Fa_Tr = _context.FA_Transportation.Where(x => x.RequestID == RequestID).FirstOrDefault();

            var IndividualIDM = new SqlParameter("@thiscasenumber", IndividualID);



            //User_Log lg = new User_Log();
            //lg.Log_date = DateTime.Now;
            //lg.User_Name = HttpContext.GetProGresUserName();
            //lg.CaseNumber = Fa_Tr.CaseNumber;
            //lg.Action = "Print Transportation Assistance- IndividualID: " + IndividualID;
            //_context.User_Log.Add(lg);
            //_context.SaveChanges();


            //var Assishopital = _context.printdetails.FromSqlRaw("exec sp_PrintDetailsAllDe @thiscasenumber", IndividualIDM).AsNoTracking().AsEnumerable().FirstOrDefault();


            //byte[] sp_imageData = Assishopital.Photo;

            // Creating a new document.
            WordDocument document = new WordDocument();
            //Adding a new section to the document.
            WSection section = document.AddSection() as WSection;
            //Set Margin of the section
            section.PageSetup.Margins.All = 72;
            //Set page size of the section
            section.PageSetup.PageSize = new Syncfusion.Drawing.SizeF(612, 792);
            //Create Paragraph styles
            WParagraphStyle style = document.AddParagraphStyle("Normal") as WParagraphStyle;
            style.CharacterFormat.FontName = "Calibri";
            style.CharacterFormat.FontSize = 11f;
            style.ParagraphFormat.BeforeSpacing = 0;
            style.ParagraphFormat.AfterSpacing = 8;
            style.ParagraphFormat.LineSpacing = 13.9f;

            style = document.AddParagraphStyle("Heading 1") as WParagraphStyle;
            style.ApplyBaseStyle("Normal");
            style.CharacterFormat.FontName = "Calibri Light";
            style.CharacterFormat.FontSize = 16f;
            style.CharacterFormat.TextColor = Syncfusion.Drawing.Color.FromArgb(46, 116, 181);
            style.ParagraphFormat.BeforeSpacing = 12;
            style.ParagraphFormat.AfterSpacing = 0;
            style.ParagraphFormat.Keep = true;
            style.ParagraphFormat.KeepFollow = true;
            style.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
            IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();

            // Gets the image stream.
            FileStream imageStream = new FileStream("wwwroot/images/unhcrlogo.jpg", FileMode.Open, FileAccess.Read);
            IWPicture picture = paragraph.AppendPicture(imageStream);
            picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
            picture.VerticalOrigin = VerticalOrigin.Margin;
            picture.VerticalPosition = -45;
            picture.HorizontalOrigin = HorizontalOrigin.Column;
            picture.HorizontalPosition = 0.5f;
            picture.WidthScale = 70;
            picture.HeightScale = 55;



            FileStream imageStream25 = new FileStream("wwwroot/images/apwletter.jpg", FileMode.Open, FileAccess.Read);
            IWPicture picture5 = paragraph.AppendPicture(imageStream25);
            picture5.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
            picture5.VerticalOrigin = VerticalOrigin.Margin;
            //YUKARi almak icin
            //picture5.VerticalPosition = 30;
            //picture5.HorizontalOrigin = HorizontalOrigin.Column;
            //picture5.HorizontalPosition = 103;
            picture5.WidthScale = 85;
            picture5.HeightScale = 85;

            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            // paragraph.ParagraphFormat.FirstLineIndent = 36;
            WTextRange textRange = paragraph.AppendText("") as WTextRange;
            textRange.CharacterFormat.FontSize = 11f;
            textRange.CharacterFormat.FontName = "Calibri";

            textRange.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;

            //Appends paragraph.
            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Heading 1");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("BMMYK Birleşmış Milletler Mülteciler Yüksek Komiserliği Sancak Mah. Tiflis Cad. 552. Sokak No3 06550 AnkaraTel: (312) 409 7000  Fax(312) 441 2173  Email turan@unhcr.org") as WTextRange;
            textRange.CharacterFormat.FontSize = 10f;
            textRange.CharacterFormat.FontName = "Calibri";
            textRange.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;


            IWTable table2 = section.AddTable();
            table2.ResetCells(1, 2);
            table2.TableFormat.Borders.BorderType = BorderStyle.None;
            table2.TableFormat.IsAutoResized = true;


            paragraph = table2[0, 0].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            textRange = paragraph.AppendText("BMMYK Dosya numarası:" + Fa_Tr.CaseNumber + "\n" + "Seri No: ( " + Fa_Tr.RequestID + " )") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;


            paragraph = table2[0, 1].AddParagraph();
            paragraph.ParagraphFormat.AfterSpacing = 0;
            paragraph.ParagraphFormat.LineSpacing = 9f;
            paragraph.BreakCharacterFormat.FontSize = 9f;
            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            textRange = paragraph.AppendText("Veriliş tarihi:  " + DateTime.Now.ToString("dd/mm/yyyy") + "\r") as WTextRange;
            textRange.CharacterFormat.FontSize = 9f;




            paragraph = section.AddParagraph();
            paragraph.ParagraphFormat.FirstLineIndent = 0;
            paragraph.BreakCharacterFormat.FontSize = 14f;
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            textRange = paragraph.AppendText("\n" + Fa_Tr.NameOfTheFirm + "\n") as WTextRange;
            textRange.CharacterFormat.FontSize = 14f;



            DateTime datenew = Convert.ToDateTime(Fa_Tr.TravelingOn);

            if (Fa_Tr.Nationality == "IRQ")
            {
                Fa_Tr.Nationality = "Irak";
            }
            else if (Fa_Tr.Nationality == "IRN")
            {
                Fa_Tr.Nationality = "Iran";
            }
            else if (Fa_Tr.Nationality == "AFG")
            {
                Fa_Tr.Nationality = "Afgan";
            }
            else if (Fa_Tr.Nationality == "SOM")
            {
                Fa_Tr.Nationality = "Somali";
            }
            else if (Fa_Tr.Nationality == "SUD")
            {
                Fa_Tr.Nationality = "Sudan";
            }
            else if (Fa_Tr.Nationality == "KEN")
            {
                Fa_Tr.Nationality = "Kenyan";
            }
            else if (Fa_Tr.Nationality == "PAL")
            {
                Fa_Tr.Nationality = "Filistin";
            }
            else if (Fa_Tr.Nationality == "YEM")
            {
                Fa_Tr.Nationality = "Yemen  ";
            }
            else if (Fa_Tr.Nationality == "ETH")
            {
                Fa_Tr.Nationality = "Etiyopya";
            }
            else if (Fa_Tr.Nationality == "UZB")
            {
                Fa_Tr.Nationality = "Uzbek";
            }
            else if (Fa_Tr.Nationality == "CHI")
            {
                Fa_Tr.Nationality = "Cin";
            }
            else if (Fa_Tr.Nationality == "KGZ")
            {
                Fa_Tr.Nationality = "Kyrgyz";
            }
            else if (Fa_Tr.Nationality == "CMR")
            {
                Fa_Tr.Nationality = "Cameroon";
            }
            else if (Fa_Tr.Nationality == "ICO")
            {
                Fa_Tr.Nationality = "Ivorian";
            }
            else if (Fa_Tr.Nationality == "ICO")
            {
                Fa_Tr.Nationality = "Ivorian";
            }
            else if (Fa_Tr.Nationality == "LBR")
            {
                Fa_Tr.Nationality = "Filipin";
            }


            //Appends paragraph.
            paragraph = section.AddParagraph();
            // paragraph.ParagraphFormat.FirstLineIndent = 10;
            paragraph.BreakCharacterFormat.FontSize = 12f;
            textRange = paragraph.AppendText("Aşağıda kimlik bilgileri bulunan " + Fa_Tr.Nationality + " uyruklu kişi(ler) " + datenew.ToString("dd/MM/yyyy") + " tarihinde belirtilen şehire gideceklerdir.Bilet ücreti Kurumumuz tarafından karşılanacaktır.\n Faturanızın, orijinal mektubumuz ile birlikte adresimize gönderilmesini rica ederiz.\n") as WTextRange;
            textRange.CharacterFormat.FontSize = 12f;



            string type = Convert.ToString(Fa_Tr.TravelType);
            if (type == "2")
            {
                type = "Gidiş Donüş";

                IWTable table6 = section.AddTable();
                table6.ResetCells(9, 2);
                table6.TableFormat.Borders.BorderType = BorderStyle.Hairline;
                table6.TableFormat.IsAutoResized = true;

                //row 1
                paragraph = table6[0, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("Dosya numarası") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                paragraph = table6[0, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(Fa_Tr.CaseNumber) as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;




                //second row
                paragraph = table6[1, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("Adı / Soyadı") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                paragraph = table6[1, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(Fa_Tr.GivenName + " " + Fa_Tr.FamilyName) as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;




                //thered row
                paragraph = table6[2, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("Uyruğu") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                paragraph = table6[2, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(Fa_Tr.Nationality) as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                //row 4
                paragraph = table6[3, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("Gidiş Tarihi") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                paragraph = table6[3, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(datenew.ToString("dd/MM/yyyy")) as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;

                //row 5
                paragraph = table6[4, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("Gideceği şehir") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                paragraph = table6[4, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(Fa_Tr.TravelingTo) as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;

                //row 6
                paragraph = table6[5, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("Kişi sayısı") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                paragraph = table6[5, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("" + Fa_Tr.NumberOfPersons) as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                //row 7
                paragraph = table6[6, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("Yolculuk türü") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;

                paragraph = table6[6, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(type) as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                //row8

                DateTime datedonus = Convert.ToDateTime(Fa_Tr.RTravelingOn);

                paragraph = table6[7, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("Donuş Tarihi") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;

                paragraph = table6[7, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(datedonus.ToString("dd/MM/yyyy")) as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                //row9
                paragraph = table6[8, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("Donuş şehir") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;

                paragraph = table6[8, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("" + Fa_Tr.RTravelingTo + "") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;

            }
            else
            {
                type = "Tek Gidiş";
                IWTable table6 = section.AddTable();
                table6.ResetCells(7, 2);
                table6.TableFormat.Borders.BorderType = BorderStyle.Hairline;
                table6.TableFormat.IsAutoResized = true;

                //row 1
                paragraph = table6[0, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" Dosya numarası ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                paragraph = table6[0, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" " + Fa_Tr.CaseNumber + " ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;




                //second row
                paragraph = table6[1, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" Adı / Soyadı ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                paragraph = table6[1, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" " + Fa_Tr.GivenName + " " + Fa_Tr.FamilyName + " ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;




                //thered row
                paragraph = table6[2, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" Uyruğu ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                paragraph = table6[2, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" " + Fa_Tr.Nationality + " ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                //row 4
                paragraph = table6[3, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" Gidiş Tarihi ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                paragraph = table6[3, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" " + datenew.ToString("dd/MM/yyyy") + " ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;

                //row 5
                paragraph = table6[4, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" Gideceği şehir ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                paragraph = table6[4, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" " + Fa_Tr.TravelingTo + " ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;

                //row 6
                paragraph = table6[5, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" Kişi sayısı ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                paragraph = table6[5, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" " + Fa_Tr.NumberOfPersons + " ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                //row 7
                paragraph = table6[6, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" Yolculuk türü ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;



                paragraph = table6[6, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText(" " + type + " ") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;
            }


            paragraph = section.AddParagraph();
            //  paragraph.ParagraphFormat.FirstLineIndent = 10;
            paragraph.BreakCharacterFormat.FontSize = 10f;
            textRange = paragraph.AppendText("\nBM Mülteciler Yüksek Komiserliği\n") as WTextRange;
            textRange.CharacterFormat.FontSize = 10f;
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;



            paragraph = section.AddParagraph();
            //  paragraph.ParagraphFormat.FirstLineIndent = 10;
            paragraph.BreakCharacterFormat.FontSize = 11f;
            textRange = paragraph.AppendText(Fa_Tr.NameOfTheFirm + "\r") as WTextRange;
            textRange = paragraph.AppendText("Aşti, Peron " + "23 Konya Yolu\r") as WTextRange;
            textRange = paragraph.AppendText("Tel: 224 15 65\r") as WTextRange;
            textRange.CharacterFormat.FontSize = 11f;
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;






            section.AddParagraph();

            //Instantiation of DocIORenderer for Word to PDF conversion
            DocIORenderer render = new DocIORenderer();
            //Converts Word document into PDF document
            PdfDocument pdfDocument = render.ConvertToPDF(document);
            //Releases all resources used by the Word document and DocIO Renderer objects
            render.Dispose();
            document.Dispose();
            //Save the document into stream.
            MemoryStream stream = new MemoryStream();
            pdfDocument.Save(stream);
            stream.Position = 0;
            //pdfDocument.Save("DoctoPDF.pdf", Response, HttpReadType.Open);


            //Close the documents.
            pdfDocument.Close(true);
            //Defining the ContentType for pdf file.
            string contentType = "application/pdf";
            //Define the file name.
            string fileName = Fa_Tr.CaseNumber + ".pdf";


            //Creates a FileContentResult object by using the file contents, content type, and file name.
            return File(stream, contentType, fileName);

            ////Saves the Word document to  MemoryStream
            //MemoryStream stream = new MemoryStream();
            //document.Save(stream, FormatType.Docx);
            //stream.Position = 0;



            ////Opens an existing document from stream through constructor of `WordDocument` class
            //FileStream fileStreamPath = new FileStream(@"Sample.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            ////Creates an empty Word document instance
            ////Loads or opens an existing word document through Open method of WordDocument class
            //document.Open(fileStreamPath, FormatType.Automatic);


            ////Download Word document in the browser
            //return File(stream, "application/msword", "Sample.docx");
        }


        [HttpGet("{IndividualID}/{casenumber}")]
        public IActionResult WithdrawnLetter(string IndividualID, string casenumber)
        {
            var casenumberP = new SqlParameter("@thiscasenumber", casenumber);
            var WithdrawnL = _context.RSTCaseBioData.FromSqlRaw("exec sp_PrintDetails @thiscasenumber", casenumberP).AsNoTracking().AsEnumerable().FirstOrDefault();
            var IdnID = new SqlParameter("@thiscasenumber", WithdrawnL.IndividualID);

            var WithdrawnIn = _context.printdetails.FromSqlRaw("exec sp_PrintDetailsAllDe @thiscasenumber", IdnID).AsNoTracking().AsEnumerable().FirstOrDefault();

            var DepDetails = _context._DepDetailsTable.FromSqlRaw("exec sp_DepDetails @thiscasenumber", casenumberP).AsNoTracking().ToList();

            //  return View(WithdrawnIn);

            if (WithdrawnIn.ProcessingGroupSize == 0)
            {
                return new EmptyResult();
            }
            else
            {


                // Creating a new document.
                WordDocument document = new WordDocument();
                //Adding a new section to the document.
                WSection section = document.AddSection() as WSection;
                //Set Margin of the section
                section.PageSetup.Margins.All = 72;
                //Set page size of the section
                section.PageSetup.PageSize = new Syncfusion.Drawing.SizeF(612, 792);

                //Create Paragraph styles
                WParagraphStyle style = document.AddParagraphStyle("Normal") as WParagraphStyle;
                style.CharacterFormat.FontName = "Calibri";
                style.CharacterFormat.FontSize = 11f;
                style.ParagraphFormat.BeforeSpacing = 0;
                style.ParagraphFormat.AfterSpacing = 8;
                style.ParagraphFormat.LineSpacing = 13.9f;

                style = document.AddParagraphStyle("Heading 1") as WParagraphStyle;
                style.ApplyBaseStyle("Normal");
                style.CharacterFormat.FontName = "Calibri Light";
                style.CharacterFormat.FontSize = 16f;
                style.CharacterFormat.TextColor = Syncfusion.Drawing.Color.FromArgb(46, 116, 181);
                style.ParagraphFormat.BeforeSpacing = 12;
                style.ParagraphFormat.AfterSpacing = 0;
                style.ParagraphFormat.Keep = true;
                style.ParagraphFormat.KeepFollow = true;
                style.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
                IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();

                // Gets the image stream.
                FileStream imageStream = new FileStream("wwwroot/images/unhcrlogo.jpg", FileMode.Open, FileAccess.Read);
                IWPicture picture = paragraph.AppendPicture(imageStream);
                picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                picture.VerticalOrigin = VerticalOrigin.Margin;
                picture.VerticalPosition = -45;
                picture.HorizontalOrigin = HorizontalOrigin.Column;
                picture.HorizontalPosition = 0.5f;
                picture.WidthScale = 70;
                picture.HeightScale = 55;


                FileStream imageStream25 = new FileStream("wwwroot/images/apwletter.jpg", FileMode.Open, FileAccess.Read);
                IWPicture picture5 = paragraph.AppendPicture(imageStream25);
                picture5.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                picture5.VerticalOrigin = VerticalOrigin.Margin;
                //YUKARi almak icin
                //picture5.VerticalPosition = 150;
                //picture5.HorizontalOrigin = HorizontalOrigin.Column;
                //picture5.HorizontalPosition = 108;
                picture5.WidthScale = 85;
                picture5.HeightScale = 85;



                paragraph.ApplyStyle("Normal");
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                // paragraph.ParagraphFormat.FirstLineIndent = 36;
                WTextRange textRange = paragraph.AppendText("") as WTextRange;
                textRange.CharacterFormat.FontSize = 11f;
                textRange.CharacterFormat.FontName = "Calibri";

                textRange.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;

                //Appends paragraph.
                paragraph = section.AddParagraph();
                paragraph.ApplyStyle("Heading 1");
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("BMMYK Birleşmış Milletler Mülteciler Yüksek Komiserliği Sancak Mah. Tiflis Cad. 552. Sokak No3 06550 AnkaraTel: (312) 409 7000  Fax(312) 441 2173  Email turan@unhcr.org") as WTextRange;
                textRange.CharacterFormat.FontSize = 8f;
                textRange.CharacterFormat.FontName = "Calibri";
                textRange.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;


                IWTable table2 = section.AddTable();
                table2.ResetCells(1, 2);
                table2.TableFormat.Borders.BorderType = BorderStyle.None;
                table2.TableFormat.IsAutoResized = true;


                paragraph = table2[0, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("Reference Number: 2020 / ILC / \r") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;


                paragraph = table2[0, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("Date of Issue (Veriliş Tarihi): " + DateTime.Now.ToString("dd/mm/yyyy") + "\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;




                paragraph = section.AddParagraph();
                paragraph.ParagraphFormat.FirstLineIndent = 0;
                paragraph.BreakCharacterFormat.FontSize = 11f;
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("UNHCR REFUGEE CERTIFICATE\nBMMYK MÜLTECİ BELGESİ") as WTextRange;
                textRange.CharacterFormat.FontSize = 11f;


                IWTable table3 = section.AddTable();
                table3.ResetCells(1, 3);
                table3.TableFormat.Borders.BorderType = BorderStyle.None;
                table3.TableFormat.IsAutoResized = true;


                //Appends paragraph.
                paragraph = table3[0, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";

                textRange = paragraph.AppendText("Name of Applicant (Başvuru sahibinin Adı):\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("UNHCR Registration No. (BMMYK Kayıt No.):\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("Date of Birth (Doğum Tarihi):\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("Place of Birth (Doğum Yeri):\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";

                textRange = paragraph.AppendText(" Nationality (Uyruğu):\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                //Appends paragraph.
                paragraph = table3[0, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;


                //Appends paragraph.
                paragraph = table3[0, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";

                textRange = paragraph.AppendText(WithdrawnIn.GivenName + "\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText(WithdrawnIn.ProcessingGroupNumber + "\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText(WithdrawnIn.DateofBirth.ToString("dd-MM-yyyy") + "\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText(WithdrawnIn.BirthCityTownVillage + "\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText(WithdrawnIn.OriginCountryCode + "\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 7f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                //Appends paragraph.
                paragraph = table3[0, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 9f;
                paragraph.BreakCharacterFormat.FontSize = 9f;



                //Appends paragraph.
                paragraph = table3[0, 2].AddParagraph();
                paragraph.ApplyStyle("Heading 1");
                paragraph.ParagraphFormat.LineSpacing = 9f;
                byte[] sp_imageData3bit = WithdrawnIn.Photo;

                WPicture picture2 = new WPicture(document);
                if (sp_imageData3bit != null)
                {

                    var streammmm2 = new MemoryStream(sp_imageData3bit);

                    Bitmap bmp2;
                    //bmp = new Bitmap(streammmm);
                    //img.Height = 10;
                    Image image2 = Image.FromStream(streammmm2);

                    //Image image = new Bitmap(@"c:\FakePhoto.jpg");

                    int target_height2 = 75;
                    int target_width2 = 75;

                    //Image target_image;
                    Rectangle dest_rect2 = new Rectangle(0, 0, target_width2, target_height2);
                    Bitmap destImage2 = new Bitmap(target_width2, target_height2);

                    destImage2.SetResolution(image2.HorizontalResolution, image2.VerticalResolution);
                    using (var g = Graphics.FromImage(destImage2))
                    {
                        g.CompositingMode = CompositingMode.SourceCopy;
                        g.CompositingQuality = CompositingQuality.HighQuality;
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.SmoothingMode = SmoothingMode.HighQuality;
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        using (var wrapmode = new ImageAttributes())
                        {
                            wrapmode.SetWrapMode(System.Drawing.Drawing2D.WrapMode.TileFlipXY);
                            g.DrawImage(image2, dest_rect2, 0, 0, image2.Width, image2.Height, GraphicsUnit.Pixel, wrapmode);
                        }
                    }
                    ImageConverter converter2 = new ImageConverter();
                    byte[] sp_imageData = (byte[])converter2.ConvertTo(destImage2, typeof(byte[]));



                    picture2.LoadImage(sp_imageData);
                    picture2.Height = 250;
                    picture2.Width = 250;
                    paragraph.Items.Add(picture2);

                    // picture = paragraph.AppendPicture(image21);
                    picture2.TextWrappingStyle = TextWrappingStyle.TopAndBottom;
                    picture2.VerticalOrigin = VerticalOrigin.Paragraph;
                    // picture2.VerticalPosition = 8.2f;
                    picture2.HorizontalOrigin = HorizontalOrigin.Column;
                    //picture2.HorizontalPosition = -14.95f;
                    picture2.WidthScale = 150;
                    picture2.HeightScale = 100;
                }

                else
                {
                    FileStream imageStreamNull = new FileStream("wwwroot/images/user-image.png", FileMode.Open, FileAccess.Read);
                    picture2.LoadImage(imageStreamNull);
                    picture2.Height = 100;
                    picture2.Width = 100;
                    paragraph.Items.Add(picture2);

                    // picture = paragraph.AppendPicture(image21);
                    picture2.TextWrappingStyle = TextWrappingStyle.TopAndBottom;
                    picture2.VerticalOrigin = VerticalOrigin.Paragraph;
                    // picture2.VerticalPosition = 8.2f;
                    picture2.HorizontalOrigin = HorizontalOrigin.Column;
                    //picture2.HorizontalPosition = -14.95f;
                    picture2.WidthScale = 5;
                    picture2.HeightScale = 5;
                }




                paragraph = section.AddParagraph();
                // paragraph.ParagraphFormat.FirstLineIndent = 36;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("To Whom It May Concern") as WTextRange;
                textRange.CharacterFormat.FontSize = 9;




                //Appends paragraph.
                paragraph = section.AddParagraph();
                // paragraph.ParagraphFormat.FirstLineIndent = 10;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                textRange = paragraph.AppendText("This is to certify that the above-named person, national of " + WithdrawnIn.OriginCountryCode + ", has withdrawn his/her application for asylum with our office on " + DateTime.Now + ". Therefore, Mr/Ms " + WithdrawnIn.GivenName + " " + WithdrawnIn.FamilyName + " case has been closed with UNHCR according to his/her wish.") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;

                paragraph = section.AddParagraph();
                //  paragraph.ParagraphFormat.FirstLineIndent = 36;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("İlgili Makama") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;

                //Appends paragraph.
                paragraph = section.AddParagraph();
                ///  paragraph.ParagraphFormat.FirstLineIndent = 10;
                paragraph.BreakCharacterFormat.FontSize = 9f;
                textRange = paragraph.AppendText("İşbu mektupla, " + WithdrawnIn.OriginCountryCode + " uyruklu, yaptığı iltica başvurusunu " + DateTime.Now + " tarihinde geri çektiğini bildiririz. Bu nedenle, Bay/Bayan " + WithdrawnIn.GivenName + " " + WithdrawnIn.FamilyName + " isteği üzerine dosyası kapatılmıştır.") as WTextRange;
                textRange.CharacterFormat.FontSize = 9f;

                int sizeG = DepDetails.Count;

                if (sizeG != 0)
                {
                    IWTable table7 = section.AddTable();
                    table7.ResetCells(sizeG, 4);
                    table7.TableFormat.Borders.BorderType = BorderStyle.Thick;
                    table7.TableFormat.IsAutoResized = true;
                    int i = 0;
                    foreach (var item in DepDetails)
                    {


                        if (i < DepDetails.Count)
                        {
                            paragraph = table7[i, 0].AddParagraph();

                            byte[] sp_imageData2bit = item.Photo;

                            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                            WPicture picture23 = new WPicture(document);

                            if (sp_imageData2bit != null)
                            {
                                var streammmm = new MemoryStream(sp_imageData2bit);
                                Bitmap bmp;
                                Image image = Image.FromStream(streammmm);

                                int target_height = 75;
                                int target_width = 75;

                                Rectangle dest_rect = new Rectangle(0, 0, target_width, target_height);
                                Bitmap destImage = new Bitmap(target_width, target_height);

                                destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);
                                using (var g = Graphics.FromImage(destImage))
                                {
                                    g.CompositingMode = CompositingMode.SourceCopy;
                                    g.CompositingQuality = CompositingQuality.HighQuality;
                                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                    g.SmoothingMode = SmoothingMode.HighQuality;
                                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                                    using (var wrapmode = new ImageAttributes())
                                    {
                                        wrapmode.SetWrapMode(System.Drawing.Drawing2D.WrapMode.TileFlipXY);
                                        g.DrawImage(image, dest_rect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapmode);
                                    }
                                }
                                ImageConverter converter = new ImageConverter();
                                byte[] sp_imageData2 = (byte[])converter.ConvertTo(destImage, typeof(byte[]));
                                paragraph.ApplyStyle("Heading 1");



                                picture23.LoadImage(sp_imageData2);
                                picture23.Height = 10;
                                picture23.Width = 10;
                                paragraph.Items.Add(picture23);


                                picture23.TextWrappingStyle = TextWrappingStyle.Inline;
                                picture23.VerticalOrigin = VerticalOrigin.Paragraph;
                                picture23.HorizontalOrigin = HorizontalOrigin.Page;
                                picture23.WidthScale = 50;
                                picture23.HeightScale = 50;

                            }
                            else
                            {

                                FileStream imageStreamNullloop = new FileStream("wwwroot/images/user-image.png", FileMode.Open, FileAccess.Read);
                                picture23.LoadImage(imageStreamNullloop);
                                picture23.Height = 10;
                                picture23.Width = 10;
                                paragraph.Items.Add(picture23);


                                picture23.TextWrappingStyle = TextWrappingStyle.Inline;
                                picture23.VerticalOrigin = VerticalOrigin.Paragraph;
                                picture23.HorizontalOrigin = HorizontalOrigin.Page;
                                picture23.WidthScale = 5;
                                picture23.HeightScale = 5;


                            }


                            paragraph = table7[i, 1].AddParagraph();
                            paragraph.ParagraphFormat.AfterSpacing = 0;
                            paragraph.ParagraphFormat.LineSpacing = 9f;
                            paragraph.BreakCharacterFormat.FontSize = 9f;
                            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                            textRange = paragraph.AppendText("\r\r" + item.GivenName) as WTextRange;
                            textRange.CharacterFormat.FontSize = 9f;




                            paragraph = table7[i, 2].AddParagraph();
                            paragraph.ParagraphFormat.AfterSpacing = 0;
                            paragraph.ParagraphFormat.LineSpacing = 9f;
                            paragraph.BreakCharacterFormat.FontSize = 9f;
                            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                            textRange = paragraph.AppendText("\r\r" + item.FamilyName) as WTextRange;
                            textRange.CharacterFormat.FontSize = 9f;


                            paragraph = table7[i, 3].AddParagraph();
                            paragraph.ParagraphFormat.AfterSpacing = 0;
                            paragraph.ParagraphFormat.LineSpacing = 9f;
                            paragraph.BreakCharacterFormat.FontSize = 9f;
                            paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                            textRange = paragraph.AppendText("\r\r" + item.RegDate) as WTextRange;
                            textRange.CharacterFormat.FontSize = 9f;
                        }
                        else
                        {

                        }

                        i++;

                    }
                }








                //Instantiation of DocIORenderer for Word to PDF conversion
                DocIORenderer render = new DocIORenderer();
                //Converts Word document into PDF document
                PdfDocument pdfDocument = render.ConvertToPDF(document);
                //Releases all resources used by the Word document and DocIO Renderer objects
                render.Dispose();
                document.Dispose();
                //Save the document into stream.
                MemoryStream stream = new MemoryStream();
                pdfDocument.Save(stream);
                stream.Position = 0;
                //pdfDocument.Save("DoctoPDF.pdf", Response, HttpReadType.Open);


                //Close the documents.
                pdfDocument.Close(true);
                //Defining the ContentType for pdf file.
                string contentType = "application/pdf";
                //Define the file name.
                string fileName = WithdrawnIn.ProcessingGroupNumber +".pdf";


                //Creates a FileContentResult object by using the file contents, content type, and file name.
                return File(stream, contentType, fileName);
            }


        }

        }
}


