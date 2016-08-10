using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Configuration;

namespace GoikenIndar
{
    class Code
    {
        internal void extract(string ActivePath)
        {

            DirectoryInfo dinfo1 = new DirectoryInfo(ActivePath);
            //To bind the filenames in the listbox.
            foreach (FileInfo fInfo in dinfo1.GetFiles())//To bind the list box only when it is loaded for first time.
            {
                if (fInfo.Extension == ".msg")
                {
                    // Create a new Application Class
                    Application app = new Application();
                    //// Create a MailItem object
                    MailItem item = (MailItem)app.CreateItemFromTemplate(fInfo.FullName, Type.Missing);
                    // Access different fields of the message

                    Microsoft.Office.Interop.Outlook.Attachments attachments = item.Attachments;
                    if (attachments != null && attachments.Count > 0)
                    {
                        for (int i = 1; i <= attachments.Count; i++)
                        {
                            Microsoft.Office.Interop.Outlook.Attachment attachment = attachments[i];
                            if (attachment.Type == Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue)
                            {
                                if (!item.HTMLBody.Contains(attachment.FileName))
                                {
                                    string filename = Path.Combine(ActivePath + @"\", attachment.FileName);
                                    attachment.SaveAsFile(filename); 
                                }
                                
                                
                                

                            }
                        }
                    }
                    //app.Quit();
                }
            }
            
            
        }

        internal void reponse(string p)
        {
            // Create a new Application Class
            Application app = new Application();
            //// Create a MailItem object
            //MailItem item = (MailItem)app.CreateItemFromTemplate(p, Type.Missing);

            string adjuntatzeko = @"C:\Users\igarmendia.LASER\Documents\visual studio 2010\Projects\GoikenIndar\GoikenIndar\bin\Debug\pdf\adjuntatzeko.pdf";
            Outlook.Application outlookApp = new Outlook.Application();
            //Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
            //mailItem.Subject = item.Subject;
            //mailItem.To = item.SenderEmailAddress;
            ////mailItem.To = "alopetegi@semakprocesados.com";

            //mailItem.Body = item.Body;
            //mailItem.Importance = Outlook.OlImportance.olImportanceNormal;
            //mailItem.Display(false);

            Outlook.MailItem mailItem = (Outlook.MailItem) app.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = "This is the subject";
            mailItem.To = "someone@example.com";
            mailItem.Body = "This is the message.";
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;

            //merge both the  files
            byte[] _byte2 = File.ReadAllBytes(@"C:\Users\igarmendia.LASER\Documents\visual studio 2010\Projects\GoikenIndar\GoikenIndar\bin\Debug\pdf\GO17967.pdf");//file in Lotes folder
            //byte[] _byte2 = File.ReadAllBytes(Path2);
            byte[] _byte1 = File.ReadAllBytes(p);//file in codigo folder

            List<byte[]> sourceFiles = new List<byte[]>();
            sourceFiles.Add(_byte1);
            sourceFiles.Add(_byte2);
            byte[] _byte3 = MergeFiles(sourceFiles, "");
            PdfReader reader = new PdfReader(_byte3);
            PdfContentByte content;
            PdfStamper pdfStamper = new PdfStamper(reader, new FileStream(adjuntatzeko, FileMode.Create));
            //to imprint the Lote number
            //for (int i = 0; i < reader.NumberOfPages; i++)
            //{
            //    content = pdfStamper.GetOverContent(i + 1);
            //    content.SaveState();
            //    content.SetColorFill(BaseColor.RED);
            //    content.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 25f);
            //    content.BeginText();
            //    content.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 20, 20, 0);
            //    content.EndText();
            //    content.RestoreState();
            //}

            // property changes that you can not edit the PDF
            pdfStamper.FormFlattening = false;

            pdfStamper.Close();

            //System.IO.File.Copy(p, System.IO.Path.Combine(@"C:\Users\igarmendia.LASER\Documents\visual studio 2010\Projects\GoikenIndar\GoikenIndar\bin\Debug\pdf\ ", "adjuntatzeko.pdf"), true);//copy to download-certificate folder
            mailItem.Attachments.Add(adjuntatzeko, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
            mailItem.Display(false);



        }

        private byte[] MergeFiles(List<byte[]> sourceFiles, string LogPath)
        {
            Document document = new Document();
            MemoryStream output = new MemoryStream();

            try
            {
                // Initialize pdf writer
                PdfWriter writer = PdfWriter.GetInstance(document, output);
                //writer.PageEvent = new PdfPageEvents();

                // Open document to write
                document.Open();
                PdfContentByte content = writer.DirectContent;

                // Iterate through all pdf documents
                for (int fileCounter = 0; fileCounter < sourceFiles.Count; fileCounter++)
                {
                    // Create pdf reader
                    PdfReader reader = new PdfReader(sourceFiles[fileCounter]);
                    int numberOfPages = reader.NumberOfPages;

                    // Iterate through all pages
                    for (int currentPageIndex = 1; currentPageIndex <= numberOfPages; currentPageIndex++)
                    {
                        // Determine page size for the current page
                        document.SetPageSize(reader.GetPageSizeWithRotation(currentPageIndex));

                        // Create page
                        document.NewPage();
                        PdfImportedPage importedPage =
                          writer.GetImportedPage(reader, currentPageIndex);


                        // Determine page orientation
                        int pageOrientation = reader.GetPageRotation(currentPageIndex);
                        if ((pageOrientation == 90) || (pageOrientation == 270))
                        {
                            content.AddTemplate(importedPage, 0, -1f, 1f, 0, 0, reader.GetPageSizeWithRotation(currentPageIndex).Height);
                        }
                        else
                        {
                            content.AddTemplate(importedPage, 1f, 0, 0, 1f, 0, 0);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                //We create a Log file to register the log details
                FileStream fs1 = null;
                //string logFile = LogPath+"\\Log\\log" + DateTime.Now.ToString("dd-MM-yyyyHHmm") + ".txt";
                string logFile = @"Log\logCertificateMerge" + DateTime.Now.ToString("dd-MM-yyyyHHmm") + ".txt";
                if (!File.Exists(logFile))
                {
                    fs1 = File.Create(logFile);
                    fs1.Dispose();
                    fs1.Close();
                }
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(logFile, true))
                {
                    //we save the exception in the log file.
                    file.WriteLine(ex.Message);
                }

                //tareaEjecutada("(log error certificado) ");
            }
            finally
            {
                document.Close();
            }
            return output.GetBuffer();
        }

        internal void responder(string albaran, string certificado, string appDirectory)
        {
            Application app = new Application();
            //// Create a MailItem object
            //MailItem item = (MailItem)app.CreateItemFromTemplate(p, Type.Missing);

            string adjuntatzeko = appDirectory+"\\adjuntatzeko_"+albaran;
            Outlook.Application outlookApp = new Outlook.Application();
            //Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
            //mailItem.Subject = item.Subject;
            //mailItem.To = item.SenderEmailAddress;
            ////mailItem.To = "alopetegi@semakprocesados.com";

            //mailItem.Body = item.Body;
            //mailItem.Importance = Outlook.OlImportance.olImportanceNormal;
            //mailItem.Display(false);

            Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = "This is the subject";
            mailItem.To = "someone@example.com";
            mailItem.Body = "This is the message.";
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;

            //merge both the  files
            byte[] _byte2 = File.ReadAllBytes(appDirectory+"\\"+certificado);//file in Lotes folder
            //byte[] _byte2 = File.ReadAllBytes(Path2);
            byte[] _byte1 = File.ReadAllBytes(appDirectory + "\\" + albaran);//file in codigo folder

            List<byte[]> sourceFiles = new List<byte[]>();
            sourceFiles.Add(_byte1);
            sourceFiles.Add(_byte2);
            byte[] _byte3 = MergeFiles(sourceFiles, "");
            PdfReader reader = new PdfReader(_byte3);
            PdfContentByte content;
            PdfStamper pdfStamper = new PdfStamper(reader, new FileStream(adjuntatzeko, FileMode.Create));
            //to imprint the Lote number
            //for (int i = 0; i < reader.NumberOfPages; i++)
            //{
            //    content = pdfStamper.GetOverContent(i + 1);
            //    content.SaveState();
            //    content.SetColorFill(BaseColor.RED);
            //    content.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 25f);
            //    content.BeginText();
            //    content.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 20, 20, 0);
            //    content.EndText();
            //    content.RestoreState();
            //}

            // property changes that you can not edit the PDF
            pdfStamper.FormFlattening = false;

            pdfStamper.Close();

            //System.IO.File.Copy(p, System.IO.Path.Combine(@"C:\Users\igarmendia.LASER\Documents\visual studio 2010\Projects\GoikenIndar\GoikenIndar\bin\Debug\pdf\ ", "adjuntatzeko.pdf"), true);//copy to download-certificate folder
            mailItem.Attachments.Add(adjuntatzeko, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
            mailItem.Display(false);
        }

        internal void responder(string[] args, string appDirectory)
        {
            Application app = new Application();
            string certificado = "";
            string albaran = args[0].ToString();
            //string certificadoPath = @"P:\0-Certificados\CHAPA\LOTES\";
            string certificadoPath = ConfigurationManager.AppSettings.Get("certificadoPath"); ;
            //// Create a MailItem object
            //MailItem item = (MailItem)app.CreateItemFromTemplate(p, Type.Missing);

            string adjuntatzeko = appDirectory + "\\adjuntatzeko_" + albaran;
            File.Copy(appDirectory +"\\"+albaran,adjuntatzeko);
            Outlook.Application outlookApp = new Outlook.Application();
            //Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
            //mailItem.Subject = item.Subject;
            //mailItem.To = item.SenderEmailAddress;
            ////mailItem.To = "alopetegi@semakprocesados.com";

            //mailItem.Body = item.Body;
            //mailItem.Importance = Outlook.OlImportance.olImportanceNormal;
            //mailItem.Display(false);

            Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = "This is the subject";
            mailItem.To = "someone@example.com";
            mailItem.Body = "This is the message.";
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;

            for (int i = 1;  i<args.Count() ; i++)
            {
                //merge both files
                byte[] _byte2 = File.ReadAllBytes(certificadoPath + "\\" + args[i]);//file in Lotes folder
                //byte[] _byte2 = File.ReadAllBytes(Path2);
                byte[] _byte1 = File.ReadAllBytes(adjuntatzeko);//file in codigo folder

                List<byte[]> sourceFiles = new List<byte[]>();
                sourceFiles.Add(_byte1);
                sourceFiles.Add(_byte2);
                byte[] _byte3 = MergeFiles(sourceFiles, "");
                PdfReader reader = new PdfReader(_byte3);
                PdfContentByte content;
                PdfStamper pdfStamper = new PdfStamper(reader, new FileStream(adjuntatzeko, FileMode.Create));
               
                pdfStamper.FormFlattening = false;

                pdfStamper.Close();
            }

            //System.IO.File.Copy(p, System.IO.Path.Combine(@"C:\Users\igarmendia.LASER\Documents\visual studio 2010\Projects\GoikenIndar\GoikenIndar\bin\Debug\pdf\ ", "adjuntatzeko.pdf"), true);//copy to download-certificate folder
            mailItem.Attachments.Add(adjuntatzeko, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
            mailItem.Display(false);
        }
    }
}
