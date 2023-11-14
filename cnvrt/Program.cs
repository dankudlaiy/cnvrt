using iText.IO.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Font;
using OfficeOpenXml;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Font;
using Xceed.Words.NET;

namespace cnvrt;

internal static class Program
{
    private static bool _xlmsg = true;
    
    [STAThread]
    private static void Main()
    {
        Console.WriteLine("Please choose folder for files converting");
        
        ApplicationConfiguration.Initialize();

        using var folderBrowserDialog = new FolderBrowserDialog();
        
        var result = folderBrowserDialog.ShowDialog();
            
        if (result == DialogResult.OK)
        {
            var selectedFolder = folderBrowserDialog.SelectedPath;

            Console.WriteLine($"Selected Folder: {selectedFolder}");
                
            foreach (var filePath in Directory.GetFiles(selectedFolder))
            {
                if (IsExcelFile(filePath))
                {
                    Console.Write($"File: {filePath}");
                
                    try
                    {
                        ConvertExcelToPdf(filePath, $"{filePath}.pdf");
                    
                        Console.WriteLine(" | Converted");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($" | Error | {ex.Message}");
                    }
                }

                if (!IsWordFile(filePath)) continue;
                
                Console.Write($"File: {filePath}");
                
                try
                {
                    var left = Console.CursorLeft;
                    
                    ConvertDocxToPdf(filePath, $"{filePath}.pdf");
                    
                    if(_xlmsg) ClearXlMsg(left);
                    
                    Console.WriteLine(" | Converted");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($" | Error | {ex.Message}");
                }
            }
            
            Console.WriteLine("\nend");
            Console.ReadLine();
        }
        else
        {
            Console.WriteLine("Folder selection canceled by the user.");
        }
    }
    
    private static void ConvertExcelToPdf(string excelFilePath, string pdfFilePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        using var package = new ExcelPackage(new FileInfo(excelFilePath));
        
        using var pdfDocument = new PdfDocument(new PdfWriter(pdfFilePath));
        
        using var document = new Document(pdfDocument);
        
        document.SetFont(GetFont());
        
        foreach (var worksheet in package.Workbook.Worksheets)
        {
            document.Add(new Paragraph(worksheet.Name));

            var table = new Table(worksheet.Dimension.Columns);
            
            for (var row = 1; row <= worksheet.Dimension.Rows; row++)
            {
                for (var col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Text;
                    table.AddCell(cellValue);
                }
            }

            document.Add(table);
            document.Add(new Paragraph("\n"));
        }
    }

    private static void ConvertDocxToPdf(string docxFilePath, string pdfFilePath)
    {
        using var doc = DocX.Load(docxFilePath);
        
        using var pdfWriter = new PdfWriter(pdfFilePath);
        
        using var pdfDocument = new PdfDocument(pdfWriter);
        
        using var document = new Document(pdfDocument);

        document.SetFont(GetFont());
        
        foreach (var paragraph in doc.Paragraphs)
        {
            document.Add(new Paragraph(paragraph.Text));
        }
    }

    private static bool IsExcelFile(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLower();
        
        return extension is ".xlsx" or ".xls";
    }
    
    private static bool IsWordFile(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLower();
        
        return extension is ".doc" or ".docx";
    }
    
    private static void ClearXlMsg(int left)
    {
        for (var i = 0; i < 4; i++)
        {
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            Console.Write(new string(' ', 67));
        }
        
        Console.SetCursorPosition(left, Console.CursorTop - 1);
        
        Console.Write(new string(' ', 67));
        
        Console.SetCursorPosition(left, Console.CursorTop);

        _xlmsg = false;
    }

    private static PdfFont GetFont()
    {
        var dir = Directory.GetCurrentDirectory();

        var path = Path.Combine(dir, "arial-unicode-ms.ttf");

        var fontProgram = FontProgramFactory.CreateFont(path);

        return PdfFontFactory.CreateFont(fontProgram, PdfEncodings.IDENTITY_H, 
            PdfFontFactory.EmbeddingStrategy.FORCE_EMBEDDED);
    }
}