using HtmlToPdf;
using System;
using System.IO;

namespace HtmlToPdf_NetFramework
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Methods methods = new Methods();
            string selectedFile = null;
            int tekrarSayisi = 0;
            do
            {
                if (tekrarSayisi != 0)
                {
                    Console.Clear();
                    Console.WriteLine("HATA: Yanlış girdi girdiniz. Lütfen talimatları dikkate alınız.");
                }
                Console.WriteLine("Adım 1: Merhaba, Html To Pdf otomasyonuna hoşgeldiniz.");
                methods.GetFileNames("Pdf Formats");
                Console.Write("Pdf 'e dönüştürüp, indirmek istediğiniz dosyanın yanındaki numarayı ya da ismini eksiksiz yazınız: ");
                string fileSelectionInput = Console.ReadLine();
                selectedFile = methods.Choice(fileSelectionInput);
                Console.WriteLine();
                tekrarSayisi++;
            } while (selectedFile is null);

            // 4. Adım: Seçilen dosyayı pdf 'e dönüştür.
            var fileDirectory = methods.GetFileDirectory("Pdf Formats");
            var uploadDirectory = methods.GetFileDirectory("Uploads");
            if (methods.CanAccessFile(Path.Combine(fileDirectory, selectedFile + ".htm"))) methods.ConvertHtmlToPdf(fileDirectory, selectedFile, uploadDirectory);

            /* Programın hemen kapanmaması istendiğinde bu yorum satırları deaktive edilebilir.
            // 5. Adım: Programın hemen kapanmaması adına bir okuma işlemi koy.
            Console.WriteLine("\n");
            Console.ReadLine();
            */
        }
    }
}
