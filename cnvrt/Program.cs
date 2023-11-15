using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace cnvrt;

internal static class Program
{
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
            
            var excelApp = new Excel.Application();
            
            var wordApp = new Word.Application();
                
            foreach (var filePath in Directory.GetFiles(selectedFolder))
            {
                if (IsExcelFile(filePath))
                {
                    Console.Write($"File: {filePath}");
                
                    try
                    {
                        ConvertExcelToPdf(excelApp, filePath, $"{filePath}.pdf");
                    
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
                    ConvertDocxToPdf(wordApp, filePath, $"{filePath}.pdf");

                    Console.WriteLine(" | Converted");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($" | Error | {ex.Message}");
                }
            }

            excelApp.Quit();
            wordApp.Quit();
            
            Marshal.ReleaseComObject(excelApp);
            Marshal.ReleaseComObject(wordApp);

            Console.WriteLine("\nend");
            Console.ReadLine();
        }
        else
        {
            Console.WriteLine("Folder selection canceled by the user.");
        }
    }
    
    private static void ConvertExcelToPdf(Excel.Application excelApp, string excelFilePath, string pdfFilePath)
    {
        var workbooks = excelApp.Workbooks;
        var workbook = workbooks.Open(excelFilePath);
        
        workbook.PrintOut(1, 2147483647, 1, false, 
            "Microsoft Print to PDF", true, false, pdfFilePath);
        
        workbook.Close();
        
        Marshal.ReleaseComObject(workbook);
        Marshal.ReleaseComObject(workbooks);
    }

    private static void ConvertDocxToPdf(Word.Application wordApp, string docxFilePath, string pdfFilePath)
    {
        var documents = wordApp.Documents;
        var document = documents.Open(docxFilePath);
        
        document.PrintOut(Type.Missing, Type.Missing, Type.Missing, pdfFilePath, "Microsoft Print to PDF");
        
        document.Close();
        
        Marshal.ReleaseComObject(document);
        Marshal.ReleaseComObject(documents);
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
}