using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Doc2Pdf.Console
{
    class Program
    {
        static void Main(string[] args)
        {
 
            Microsoft.Office.Interop.Word.Document wordDocument = null;
            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            wordDocument = appWord.Documents.Open("c:\\pocs\\teste.docx");
            wordDocument.ExportAsFixedFormat(@"c:\pocs\\teste.pdf", WdExportFormat.wdExportFormatPDF);


            //Application word = new Application();
            ////Document doc = word.Documents.Open($@"c:\Modelo Estudo Preliminar.docx");

            //object missing = Type.Missing;

            ////Abre a aplicação Word e faz uma cópia do documento mapeado
            //Microsoft.Office.Interop.Word.Application oApp = new Microsoft.Office.Interop.Word.Application();

            //object template = $@"c:\Modelo Estudo Preliminar.docx"; //@"C:\\Users\\Luiz\\Documents\\Visual Studio 2005\\Projects\\PreencherWord";
            //Microsoft.Office.Interop.Word.Document doc = oApp.Documents.Add(ref template, ref missing, ref missing, ref missing);

            //doc.Activate();
            //doc.SaveAs2(@"c:\document.pdf", WdSaveFormat.wdFormatPDF);
            //doc.Close();
        }
    }
}
