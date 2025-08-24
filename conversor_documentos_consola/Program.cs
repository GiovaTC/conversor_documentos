using System;
using System.IO;
using Xceed.Words.NET;
using ClosedXML.Excel;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Ingrese la ruta del archivo a convertir:");
        string rutaArchivo = Console.ReadLine();

        if (!File.Exists(rutaArchivo))
        {
            Console.WriteLine("El archivo no existe.");
            return;
        }

        Console.WriteLine("Seleccione el formato al que desea convertir:");
        Console.WriteLine("1 - TXT");
        Console.WriteLine("2 - DOCX");
        Console.WriteLine("3 - XLSX");
        Console.WriteLine("4 - PDF");
        string opcion = Console.ReadLine();
        
        
        string contenido = File.ReadAllText(rutaArchivo);
        string nuevoArchivo = "";

        switch (opcion)
        {
            case "1":
                nuevoArchivo = Path.ChangeExtension(rutaArchivo, ".txt");
                File.WriteAllText(nuevoArchivo, contenido);
                Console.WriteLine("Archivo convertido a TXT");
                break;

            case "2":
                nuevoArchivo = Path.ChangeExtension(rutaArchivo, ".docx");
                using (var doc = DocX.Create(nuevoArchivo))
                {
                    doc.InsertParagraph(contenido);
                    doc.Save();
                }
                Console.WriteLine("Archivo convertido a DOCX");
                break;
            
            case "3":
                nuevoArchivo = Path.ChangeExtension(rutaArchivo, ".xlsx");
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Hoja1");
                    worksheet.Cell("A1").Value = contenido;
                    workbook.SaveAs(nuevoArchivo);
                }
                Console.WriteLine("Archivo convertido a XLSX");
                break;

            case "4":
                nuevoArchivo = Path.ChangeExtension(rutaArchivo, ".pdf");
                using (var writer = new PdfWriter(nuevoArchivo))
                {
                    using (var pdf = new PdfDocument(writer))
                    {
                        var document = new Document(pdf);
                        document.Add(new Paragraph(contenido));
                        document.Close();
                    }
                }
                Console.WriteLine("Archivo convertido a PDF");
                break;

            default:
                Console.WriteLine("Opción inválida.");
                break;
        }

        if (!string.IsNullOrEmpty(nuevoArchivo))
        {
            Console.WriteLine($"Archivo guardado como: {nuevoArchivo}");
        }
    }
}