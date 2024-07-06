using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ios = System.Runtime.InteropServices;

class Program
{
    [STAThread]

    static void Main(string[] args)
    {
        try
        {
            string excelFilePath = @"C:\path\to\your\excel\file.xlsx";
            string outputFolder = @"C:\path\to\your\output\folder\";

            Excel.Application excelApp = new Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excelFilePath);

            int pageIndex = 1;
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                string sheetName = worksheet.Name;
                Range r = worksheet.Range["A1:S8"];
                r.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                if (Clipboard.GetDataObject() != null)
                {
                    IDataObject data = Clipboard.GetDataObject();
                    if (data.GetDataPresent(DataFormats.Bitmap))
                    {
                        System.Drawing.Image image = (Image)data.GetData(DataFormats.Bitmap, true);
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
            Console.WriteLine(ex);
        }
        Console.ReadKey();
    }

    static void CaptureExcelWorksheet(Worksheet worksheet, string pdfFilePath)
    {
        worksheet.PageSetup.Zoom = false;
        worksheet.PageSetup.FitToPagesWide = 1;
        worksheet.PageSetup.FitToPagesTall = 1;
        worksheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFilePath);
    }
}
