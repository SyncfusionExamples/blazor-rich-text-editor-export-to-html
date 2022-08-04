using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Syncfusion.Pdf;
using Syncfusion.HtmlConverter;
using Microsoft.JSInterop;
using Syncfusion.EJ.Export;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;


namespace Export_HTML_File.Data
{
 

    public class ExportService
    {
        public void ExportToHtml(string value)
        {
  
            WordDocument document = GetDocument(value);
            //Saves the Word document to MemoryStream
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Html);
            stream.Position = 0;
            FileStream outputStream = new FileStream("record1.html", FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
            document.Save(outputStream, FormatType.Html);
            document.Close();
            outputStream.Flush();
            outputStream.Dispose();
            // You can upload this stream to the azure
            
        }


        public WordDocument GetDocument(string htmlText)
        {
            WordDocument document = null;
            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream, System.Text.Encoding.Default);
            htmlText = htmlText.Replace("\"", "'");
            XmlConversion XmlText = new XmlConversion(htmlText);
            XhtmlConversion XhtmlText = new XhtmlConversion(XmlText);
            writer.Write(XhtmlText.ToString());
            writer.Flush();
            stream.Position = 0;
            document = new WordDocument(stream, FormatType.Html, XHTMLValidationType.None);
            return document;
        }

        public async Task SaveAs(IJSRuntime js, string filename, byte[] data)
        {
            await js.InvokeAsync<object>(
                "saveAsFile",
                filename,
                Convert.ToBase64String(data));
        }

        public MemoryStream ExportAsPdf(string content)
        {
            try
            {
                // Initialize the HTML to PDF converter
                HtmlToPdfConverter htmlConverter = new HtmlToPdfConverter();

                WebKitConverterSettings settings = new WebKitConverterSettings();

                // Used to load resources before convert
                // Map your local path here
                string baseUrl = @"C:/Users/TempKaruna/source/repos/RTEtoPDF/wwwroot/images";

                // Set WebKit path
                // Map your local path installed location here
                settings.WebKitPath = @"C:/Program Files (x86)/Syncfusion/HTMLConverter/18.1.0.52/QtBinariesDotNetCore";

                // Set additional delay; units in milliseconds;
                settings.AdditionalDelay = 3000;

                // Assign WebKit settings to HTML converter
                htmlConverter.ConverterSettings = settings;

                // Convert HTML string to PDF
                PdfDocument document = htmlConverter.Convert(content, baseUrl);

                // Save the document into stream.
                MemoryStream stream = new MemoryStream();

                document.Save(stream);

                stream.Position = 0;

                // Close the document.
                document.Close(true);

                return stream;
            }
            catch
            {
                return new MemoryStream();
            }
        }
    }
}
