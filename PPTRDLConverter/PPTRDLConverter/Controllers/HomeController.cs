using BoldReports.Writer;
using PresentationToRDL;
using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace PPTRDLConverter.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }


        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            string uploadPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Uploads");
            if (!Directory.Exists(uploadPath))
            {
                Directory.CreateDirectory(uploadPath);
            }

            if (file != null && file.ContentLength > 0)
            {
                try
                {
                    // Save uploaded file to server (optional)
                    string filePath = Path.Combine(Server.MapPath("~/Uploads"), Path.GetFileName(file.FileName));
                    file.SaveAs(filePath);

                    // Process PPTX and generate RDL
                    byte[] rdlData = ConvertPptxToRdl(filePath);

                    // Return the generated RDL as a file download
                    return File(rdlData, "application/rdl", "ConvertedReport.rdl");
                }
                catch (Exception ex)
                {
                    return Content("Error: " + ex.Message);
                }
            }
            return Content("No file uploaded.");
        }

        private byte[] ConvertPptxToRdl(string pptFilePath)
        {

            MemoryStream stream = new MemoryStream();
            try
            {
                Presentation presentation = (Presentation)Syncfusion.Presentation.Presentation.Open(pptFilePath);
                PresentationToRdlConverter converter = new PresentationToRdlConverter();
                stream = converter.Convert(presentation, System.IO.Path.GetFileNameWithoutExtension(pptFilePath), "", "", "", false);
                presentation.Close();
            }
            catch (Exception ex)
            {
                stream = null;
                //LoggerManager.LogError("PowerPoint path not found: " + ex.Message);
            }
            finally
            {
                Dispose();
            }

            using (MemoryStream memoryStream = new MemoryStream())
            {
                stream.CopyTo(memoryStream);
                return memoryStream.ToArray();
            }
        }
    }
}