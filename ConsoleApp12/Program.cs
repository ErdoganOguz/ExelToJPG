using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        try
        {
            // Path to the Excel file
            string excelFilePath = @"C:\path\to\your\excel\file.xlsx";
            // Path to the output folder
            string outputFolder = @"C:\path\to\your\output\folder\";

            // Initialize Excel application
            Excel.Application excelApp = new Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excelFilePath);

            int pageIndex = 1;
            // Loop through each worksheet in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                string sheetName = worksheet.Name;
                // Define the range to be captured
                Range r = worksheet.Range["A1:S8"];
                r.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                // Check if there is data in the clipboard
                if (Clipboard.GetDataObject() != null)
                {
                    IDataObject data = Clipboard.GetDataObject();
                    // Check if the data is in bitmap format
                    if (data.GetDataPresent(DataFormats.Bitmap))
                    {
                        // Retrieve the image from the clipboard
                        System.Drawing.Image image = (Image)data.GetData(DataFormats.Bitmap, true);
                        // Save the image as a JPEG file
                        image.Save($@"{outputFolder}{sheetName}.jpg", ImageFormat.Jpeg);
                        Console.WriteLine($"Saved sheet \"{sheetName}\"");
                    }
                    else
                    {
                        Console.WriteLine("No image in Clipboard !!");
                    }
                }
                else
                {
                    Console.WriteLine("Clipboard Empty !!");
                }

                pageIndex++;
            }

        }
        catch (Exception ex)
        {
            // Print any exceptions that occur
            Console.WriteLine(ex);
        }
        Console.ReadKey();
    }

    // Method to export a worksheet as a PDF file
    static void CaptureExcelWorksheet(Worksheet worksheet, string pdfFilePath)
    {
        worksheet.PageSetup.Zoom = false;
        worksheet.PageSetup.FitToPagesWide = 1;
        worksheet.PageSetup.FitToPagesTall = 1;
        worksheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFilePath);
    }
}
