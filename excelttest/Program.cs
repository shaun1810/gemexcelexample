using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Drawing.Imaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;

namespace excelttest
{
    class Program
    {
        public static DirectoryInfo outputDir = new DirectoryInfo($"{Environment.SpecialFolder.DesktopDirectory}");
        static void Main(string[] args)
        {
            var fileTimestamp = DateTime.UtcNow.ToFileTimeUtc();
            FileInfo newFile = GetFileInfo("AlgoSettings"+fileTimestamp+ ".xlsx");
            
            //Create the workbook
            ExcelPackage pck = new ExcelPackage(newFile);
            //Add the Content sheet
            var ws = pck.Workbook.Worksheets.Add("Algolia Settings");
            
            //Headers
            ws.Cells["B1"].Value = "Name";
            ws.Cells["C1"].Value = "Size";
            ws.Cells["D1"].Value = "Created";
            ws.Cells["E1"].Value = "Last modified";
            ws.Select("A2");
            //height is 20 pixels 
            //Add the textbox
            //Done! save the sheet
            pck.Save();
        }

        public static FileInfo GetFileInfo(string file, bool deleteIfExists = true)
        {
            var fi = new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+ "\\" +file);
            if (deleteIfExists && fi.Exists)
            {
                fi.Delete();  // ensures we create a new workbook
            }
            return fi;
        }
    }
}
