using PdfSharp;
using PdfSharp.Pdf;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using TheArtOfDev.HtmlRenderer.PdfSharp;

namespace HtmlToPdf
{
    public class Methods
    {
        /// <summary>
        /// İsmi verilen klasör içerisindeki dosyaların isimlerini döndürür.
        /// </summary>
        /// <param name="nameOfTheFile">İçerisi aracanacak olan klasörün adı</param>
        /// <returns></returns>
        public void GetFileNames(string nameOfTheFile)
        {
            string pdfFormatsDirectory = GetFileDirectory(nameOfTheFile);
            if (Directory.Exists(pdfFormatsDirectory))
            {
                Console.WriteLine($"Adım 2: {nameOfTheFile} klasörüne giriliyor.");
                var files = Directory.GetFiles(pdfFormatsDirectory).ToList();
                if (files.Count > 0)
                {
                    int i = 0;
                    string fileName = null;
                    Console.WriteLine($"Adım 3: {nameOfTheFile} Klasöründeki tüm dosyalar listeleniyor.\n");
                    foreach (var file in files)
                    {
                        fileName = Path.GetFileName(file.Remove(file.Length - 4));
                        Console.WriteLine($"[{++i}] {fileName}");
                    }
                    Console.WriteLine();
                }
                else Console.WriteLine($"Adım 3: {nameOfTheFile} içerisinde herhangi bir dosya bulunmadı.");
            }
            else Console.WriteLine($"Adım 2: {nameOfTheFile} klasörü mevcut değil.");
        }

        /// <summary>
        /// İsmi verilen klasör eğer proje içerisinde mevcut ise klasör yolunu getirir.
        /// </summary>
        /// <param name="fileName">Aranacak dosyanın adı</param>
        /// <returns></returns>
        public string GetFileDirectory(string fileName)
        {
            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string projectName = "HtmlToPdf"; //Assembly.GetExecutingAssembly().GetName().Name;
            string solutionDirectory = FindSolutionDirectory(currentDirectory, projectName);
            string pdfFormatsDirectory = Path.Combine(solutionDirectory, fileName);
            return pdfFormatsDirectory;
        }

        /// <summary>
        /// Proje çözümünün bulunduğu klasörü yolunu getirir.
        /// </summary>
        /// <param name="currentDirectory">Program.cs 'in bulunduğu klasör yolu</param>
        /// <param name="projectName">Projenin adı</param>
        /// <returns></returns>
        public string FindSolutionDirectory(string currentDirectory, string projectName)
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(currentDirectory);
            while (directoryInfo != null && !directoryInfo.Name.Equals(projectName, StringComparison.OrdinalIgnoreCase))
            {
                directoryInfo = directoryInfo.Parent;
            }
            return directoryInfo?.FullName;
        }

        /// <summary>
        /// Yapılan seçimin ilgili dosyalardan bir tanesinin ismine eşit olup olmadığını doğrular.
        /// </summary>
        /// <param name="choice">Alınan girdi, kontrol edilecek olan metin</param>
        /// <returns></returns>
        public string Choice(string choice)
        {
            for (int i = 1; i <= 6; i++)
            {
                if (isIt(choice, i.ToString()) || isIt(choice, $"invitation {i}") || isIt(choice, $"invitation - {i}") || isIt(choice, $"invitation-{i}"))
                    return $"invitation-{i}";
            }
            return null;
        }

        /// <summary>
        /// Verilen iki metnin birbirine eşit olup olmadığını doğrular.
        /// </summary>
        /// <param name="value1">Ana metin</param>
        /// <param name="value2">Eşit olup olmadığı kontrol edilecek olan metin</param>
        /// <returns></returns>
        public bool isIt(string value1, string value2)
        {
            return value1 == value2 ? true : false;
        }

        /// <summary>
        /// Yolu verilen .htm dosyasını .pdf dönüştürüp yine yolu verilen klasöre yükler.
        /// </summary>
        /// <param name="htmlFilePath">Dönüştürülecek olan dosyanın yolu</param>
        /// <param name="outputPdfPath">Dönüşüm tamamlandıktan sonra yüklenecek olan klasörün yolu</param>
        /// <param name="fileName">Pdf 'e dönüştürülecek dosyanın ismi</param>
        public void ConvertHtmlToPdf(string htmlFilePath, string fileName, string outputPdfPath)
        {
            try
            {
                var htmlFileName = fileName.Contains(".htm") ? fileName : fileName + ".htm";
                var outputFileName = outputPdfPath + "\\" + Guid.NewGuid().ToString() + ".pdf";
                Console.WriteLine("Adım 4.1: Dosya adı uzantısı ayarlandı.");
                PdfDocument pdf = PdfGenerator.GeneratePdf(File.ReadAllText($"{htmlFilePath}\\{htmlFileName}"), PageSize.A4);
                Console.WriteLine("Adım 4.2: Dosya okuma işlemi tamamlandı.");
                pdf.Save(outputFileName);
                Console.WriteLine("Adım 4.3: Dosya kaydetme işlemi tamamlandı.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Adım 4: Hata ile karşılaşıldı. İşte hata:\n{ex.Message}\n\n\n");
                throw;
            }
        }

        /// <summary>
        /// Adresi verilen klasöre ya da dosyaya olan erişimi sorgular.
        /// </summary>
        /// <param name="filePath">Dosya ya da klasör yolu</param>
        /// <returns></returns>
        public bool CanAccessFile(string filePath)
        {
            try
            {
                using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ah, hayır! Bir hata ile karşılaşıldı. İşte hata:\n{ex.Message}");
                return false;
            }
        }
    }
}
