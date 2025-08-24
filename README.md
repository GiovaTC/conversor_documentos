# Conversor de Documentos en C# (.NET) – Visual Studio 2022

<img width="2554" height="1079" alt="image" src="https://github.com/user-attachments/assets/db50b637-3e4b-4670-86fe-33488cf4763e" />

Este proyecto es una aplicación de consola en C# que permite convertir documentos entre diferentes formatos: `.txt`, `.docx`, `.xlsx`, y `.pdf`. Al convertir, imprime en pantalla el tipo de archivo al que fue convertido.

---

## 🧰 Requisitos

- Visual Studio 2022
- .NET Core / .NET 6 o superior
- Paquetes NuGet:

### 📦 Dependencias necesarias:

Instálalas desde la consola del administrador de paquetes NuGet:

```bash
Install-Package Xceed.Words.NET
Install-Package ClosedXML
Install-Package itext7
📝 Instrucciones
1. Crear el Proyecto
Abre Visual Studio 2022.

Crea un nuevo proyecto del tipo "Aplicación de consola (.NET Core)" o .NET 6/7.

Nombra el proyecto como ConversorDocumentos.

2. Agregar los Paquetes NuGet
Instala los siguientes paquetes usando NuGet:

Xceed.Words.NET – para manipular archivos .docx

ClosedXML – para manipular archivos .xlsx

itext7 – para generar archivos .pdf

💻 Código Fuente
csharp

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

✅ Funcionalidades
📥 Carga un archivo desde ruta especificada.

🔄 Permite elegir entre 4 formatos de salida.

📤 Guarda el archivo convertido en la misma ubicación.

🖨️ Informa por consola el tipo de archivo generado.

🚫 Limitaciones
Solo trabaja con contenido de texto plano.

No convierte formatos complejos (tablas, estilos, imágenes).

La entrada debe ser un archivo de texto legible como string (por ejemplo .txt).

