using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;

namespace Doc2Pdf.Controllers
{
    public class Data
    {
        public string Chave { get; set; }
        public string Valor { get; set; }
    }

    public class Document
    {
        public int DocumentID { get; set; }
    }

    public class DocumentController : ApiController
    {
        //string nome = $@"c:\\pocs\{Guid.NewGuid()}.pdf";

        public Task<HttpResponseMessage> PostFile()
        {
            string fileConverted;

            HttpRequestMessage request = this.Request;
            if (!request.Content.IsMimeMultipartContent())
            {
                throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
            }
            try
            {

                string root = System.Web.HttpContext.Current.Server.MapPath("~/App_Data/uploads");
                var provider = new MultipartFormDataStreamProvider(root);

                var task = Request.Content.ReadAsMultipartAsync(provider).
                ContinueWith<HttpResponseMessage>(o =>
                {
                    if (o.IsFaulted || o.IsCanceled)
                        throw new HttpResponseException(HttpStatusCode.InternalServerError);

                    FileInfo finfo = new FileInfo(provider.FileData.First().LocalFileName);

                    string guid = Guid.NewGuid().ToString();

                    var fileName = guid + "_" + provider.FileData.First().Headers.ContentDisposition.FileName.Replace("\"", "");
                    var fullName = Path.Combine(root, fileName);
                    fileConverted = fullName.Replace(fullName.Substring(fullName.LastIndexOf(".") + 1), "pdf");
                    File.Move(finfo.FullName, fullName);

                    ConverterDocumento(fullName, fileConverted);

                    return new HttpResponseMessage()
                    {
                        Content = new StringContent(fileConverted.Substring(fileConverted.LastIndexOf("\\") + 1))
                    };
                });
                return task;

            }
            catch (Exception ex)
            {

                return System.Threading.Tasks.Task.Factory.StartNew(() => new HttpResponseMessage()
                {
                    Content = new StringContent(ex.Message)
                });
            }

        }
        public HttpResponseMessage Get(string id)
        {
            //Dictionary<string, string> itens = model.ToDictionary(c => c.Chave, c => c.Valor);
            //var fullPath = System.Web.Hosting.HostingEnvironment.MapPath(@"~/App_Data/teste.docx");

            //PreencherContrato(fullPath, new Dictionary<string, string> {
            //    {"[data]", DateTime.Today.ToString() },
            //    {"[autor]", "José Costa" }
            //});

            string root = System.Web.HttpContext.Current.Server.MapPath("~/App_Data/uploads");
            var path = Path.Combine(root, id);

            var pdfContent = new MemoryStream(System.IO.File.ReadAllBytes(path+".pdf "));
            var stream = new MemoryStream();

            var result = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(pdfContent.ToArray())
            };

            result.Content.Headers.ContentDisposition =
                new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
                {
                    FileName = id+".pdf"
                };
            result.Content.Headers.ContentType =
                new MediaTypeHeaderValue("application/octet-stream");

            return result;
        }

        public void ConverterDocumento(string filePath, string path)
        {
            object missing = Type.Missing;

            //Abre a aplicação Word e faz uma cópia do documento mapeado
            Microsoft.Office.Interop.Word.Application oApp = new Microsoft.Office.Interop.Word.Application();

            object template = filePath;
            Microsoft.Office.Interop.Word.Document oDoc = oApp.Documents.Add(ref template, ref missing, ref missing, ref missing);

            oDoc.ExportAsFixedFormat(path, WdExportFormat.wdExportFormatPDF);

            oDoc.Close(false); // Close the Word Document.
            oApp.Quit(false); // Close Word Application.

            // Release all Interop objects.
            if (oDoc != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
            if (oApp != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp);
            oDoc = null;
            oApp = null;
            GC.Collect();
        }


        public void PreencherContrato(string templatePath, Dictionary<string, string> keyValues)
        {
            object missing = Type.Missing;

            //Abre a aplicação Word e faz uma cópia do documento mapeado
            Microsoft.Office.Interop.Word.Application oApp = new Microsoft.Office.Interop.Word.Application();

            object template = templatePath;
            Microsoft.Office.Interop.Word.Document oDoc = oApp.Documents.Add(ref template, ref missing, ref missing, ref missing);

            //Troca o conteúdo de alguns tags
            Microsoft.Office.Interop.Word.Range oRng = oDoc.Range(ref missing, ref missing);

            foreach (var item in keyValues)
            {
                oRng = oDoc.Range(ref missing, ref missing);
                SubstituirValores(ref oRng, item.Key, item.Value);
            }

            //oDoc.ExportAsFixedFormat(nome, WdExportFormat.wdExportFormatPDF);

            oDoc.Close(false); // Close the Word Document.
            oApp.Quit(false); // Close Word Application.

            // Release all Interop objects.
            if (oDoc != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
            if (oApp != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp);
            oDoc = null;
            oApp = null;
            GC.Collect();
        }


        private void SubstituirValores(ref Microsoft.Office.Interop.Word.Range oRng, object findText, object replaceWith)
        {
            object missing = Type.Missing;
            object MatchWholeWord = true;
            object Forward = true;
            object MachAllWordForms = true;

            oRng.Find.Execute(ref findText, ref missing, ref MatchWholeWord, ref missing, ref missing, ref missing, ref Forward,
            ref missing, ref missing, ref replaceWith, ref missing, ref missing, ref missing, ref missing, ref missing);
        }

    }
}
