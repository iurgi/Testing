using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;

namespace GoikenIndar
{
    class Program
    {
        static void Main(string[] args)
        {
            string albaran = "";
            string certificado = "";
            //string appDirectory = Path.GetDirectoryName(Application.ExecutablePath);
            string appDirectory = Directory.GetCurrentDirectory();

            //try
            //{
            //    albaran = args[0];
            //    certificado = args[1];

            //}
            //catch (System.Exception e)
            //{
            //    albaran = "Albarán_ 881706_INDAR.pdf";
            //    certificado = "GO17967.pdf";
            //}
            if (args.Count() == 0)
            {

                List<string> list = new List<string>();

                list.Add("Albarán_ 881706_INDAR.pdf");
                list.Add("GO17967.pdf");
                list.Add("GO17968.pdf");
                

                // convert it to an array if you want to
                args = list.ToArray();


                //args[0] = "Albarán_ 881706_INDAR.pdf";
                //args[1] = "GO17967.pdf";
                //args[2] = "GO17968.pdf";
            }

            Code kode = new Code();
            //kode.responder(albaran, certificado,appDirectory+"\\pdf\\");
            if (args.Count() > 0)
            {
                kode.responder(args, appDirectory + "\\pdf");
            }
            else
            {
                Console.WriteLine("No hay suficientes parametros");
                Console.ReadKey();
            }
        }

        private static void responder()
        {
            Code kode = new Code();
            kode.reponse(@"C:\Users\igarmendia.LASER\Documents\Visual Studio 2010\Projects\GoikenIndar\GoikenIndar\bin\Debug\pdf\Albarán_ 881706_INDAR.pdf");
        }
    }
}
