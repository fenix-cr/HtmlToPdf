using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using SelectPdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using TheArtOfDev.HtmlRenderer.PdfSharp;


namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        public ViewResult Index()
        {
            // PDFEmail();
            //  print03();
               return View();
        }

        [HttpPost]
        [ValidateInput(false)]
        public void GenerarReporte(String htmlcode, String opcion)
        {
            switch (opcion)
            {
                case "print_NReco":
                    print_NReco(htmlcode);
                    break;
                case "print_iText7":
                    print_iText7(htmlcode);
                    break;
                case "print_PdfSharp":
                    print_PdfSharp(htmlcode);
                    break;
                case "print_SelectPdf":
                    print_SelectPdf(htmlcode);
                    break;
                case "print_TheArtOfDev":
                    print_TheArtOfDev(htmlcode);
                    break;
                default:
                    break;
            }
          
        }

        //iText7
        private void print_iText7(String htmlcode)
        {
            Byte[] res = null;
            using (MemoryStream ms = new MemoryStream())
            {
                iText.Html2pdf.HtmlConverter.ConvertToPdf(htmlcode, ms);
                res = ms.ToArray();
                RetornarPdf(res);
            }

        }

        //NReco
        private void print_NReco(String htmlcode)
        {

            var htmlToPdf = new NReco.PdfGenerator.HtmlToPdfConverter();
            var res = htmlToPdf.GeneratePdf(htmlcode);

            RetornarPdf(res);
        }

        //PdfSharp
        private void print_PdfSharp(String htmlcode)
        {
            Byte[] res = null;
            using (MemoryStream ms = new MemoryStream())
            {
                var pdfFile = TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerator.GeneratePdf(htmlcode, PdfSharp.PageSize.Letter, 30);
                pdfFile.Save(ms);
                res = ms.ToArray();
            }


            RetornarPdf(res);
        }

        //SelectPdf
        private void print_SelectPdf(String htmlcode)
        {
            // instantiate a html to pdf converter object
            SelectPdf.HtmlToPdf converter = new SelectPdf.HtmlToPdf();

            converter.Options.MarginTop = 10;
            converter.Options.MarginBottom = 10;
            converter.Options.MarginLeft = 10;
            converter.Options.MarginRight = 10;

            // create a new pdf document converting an url
            SelectPdf.PdfDocument doc = converter.ConvertHtmlString(htmlcode);
            Byte[] res = null;

            using (MemoryStream memoryStream = new MemoryStream())
            {
                res = doc.Save();
                memoryStream.Close();
                RetornarPdf(res);
                // close pdf document
                doc.Close();
            }


        }

        //TheArtOfDev
        private void print_TheArtOfDev(String htmlcode)
        {
            var sbd = @"<!DOCTYPE html><html><head><title>HTML Page with text,image and CSS to PDF</title>     <META NAME='ROBOTS' CONTENT='NOINDEX, NOFOLLOW'><style type='text/css'>  #text{    max-width: 100%;    width: 35%;    text-align: center;    margin: 50px;    border: 5px solid orange;    background: whitesmoke;  }  #text h2{    background: orange;    color: white;  }#btn{      padding:10px;      border: 0px;      margin: 50px;      cursor: pointer;    }  </style>      html, body{      overflow-x: hidden;    }</style><META NAME='ROBOTS' CONTENT='NOINDEX, NOFOLLOW'></head><body>    <script type='text/javascript'></script><button id='btn'>Convert to CSS</button><div id='text'><h2>HTML Page with text,image and CSS to PDF</h2><img src='https://codingstatus.com/wp-content/uploads/2019/12/codingstatus.jpg' width='100%'><p>Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industrys standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum</p></div><script src='https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js'></script><script src='https://cdnjs.cloudflare.com/ajax/libs/html2canvas/0.4.1/html2canvas.js'></script><script src='https://cdnjs.cloudflare.com/ajax/libs/jspdf/1.0.272/jspdf.debug.js'></script><script src='custom.js'></script></body></html>";
            var sb = htmlcode;

            Byte[] res = null;
            using (MemoryStream memoryStream = new MemoryStream())
            {
                var pdf = TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerator.GeneratePdf(sb, PdfSharp.PageSize.A4);
                pdf.Save(memoryStream);
                res = memoryStream.ToArray();

                memoryStream.Close();

                RetornarPdf(res);
            }
        }

        //iTextSharp
        private void print_iTextSharp5()
        {
            Document doc = new Document(iTextSharp.text.PageSize.LETTER, 10, 10, 42, 35);

            using (MemoryStream memoryStream = new MemoryStream())
            {
                PdfWriter writer = PdfWriter.GetInstance(doc, memoryStream);
                doc.Open();
                HTMLWorker html = new HTMLWorker(doc);
                /*StyleSheet css = new StyleSheet();*/ //Not supported
                /*css.LoadTagStyle("div", "color", "red");*/
                //css.LoadStyle("div", "color", "green");
                string simple = "<html><body><h1 style='color: green;'>Heading in Green</h1><div style='color: red;'>Sample text in red color.</div></body></html>";
                html.Parse(new StringReader(simple));
                //css.LoadTagStyle("DIV", "color", "red");
                /*html.SetStyleSheet(css);*/
                doc.Close();

                byte[] bytes = memoryStream.ToArray();
                memoryStream.Close();

                // Clears all content output from the buffer stream                 
                Response.Clear();
                // Gets or sets the HTTP MIME type of the output stream.
                Response.ContentType = "application/pdf";
                // Adds an HTTP header to the output stream
                Response.AddHeader("Content-Disposition", "attachment; filename=Invoice.pdf");

                //Gets or sets a value indicating whether to buffer output and send it after
                // the complete response is finished processing.
                Response.Buffer = true;
                // Sets the Cache-Control header to one of the values of System.Web.HttpCacheability.
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                // Writes a string of binary characters to the HTTP output stream. it write the generated bytes .
                Response.BinaryWrite(bytes);
                // Sends all currently buffered output to the client, stops execution of the
                // page, and raises the System.Web.HttpApplication.EndRequest event.

                Response.End();
                // Closes the socket connection to a client. it is a necessary step as you must close the response after doing work.its best approach.
                Response.Close();
            }
        }

        private void RetornarPdf(Byte[] res)
        {
            try
            {
                // Clears all content output from the buffer stream                 
                Response.Clear();
                // Gets or sets the HTTP MIME type of the output stream.
                Response.ContentType = "application/pdf";
                // Adds an HTTP header to the output stream
                Response.AddHeader("Content-Disposition", "attachment; filename=Invoice.pdf");

                //Gets or sets a value indicating whether to buffer output and send it after
                // the complete response is finished processing.
                Response.Buffer = true;
                // Sets the Cache-Control header to one of the values of System.Web.HttpCacheability.
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                // Writes a string of binary characters to the HTTP output stream. it write the generated bytes .
                Response.BinaryWrite(res);
                // Sends all currently buffered output to the client, stops execution of the
                // page, and raises the System.Web.HttpApplication.EndRequest event.

                Response.End();
                // Closes the socket connection to a client. it is a necessary step as you must close the response after doing work.its best approach.
                Response.Close();
            }
            catch (Exception ex)
            {


            }

        }

        public ActionResult Reporte()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

    }
}