using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MakeScreenShotsDocument;

namespace MakeWordDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateWordDoc cd = new CreateWordDoc();
            cd.leftBrowser = "IE9";
            cd.leftScreenShotPath = "E:\\LillyWeb\\CialisMD\\PreScreenShots\\IE9\\Uploaded\\";
            // leave cp.rightBrowser blank to have only the left side populated
            cd.rightBrowser = "CRM";
            cd.rightScreenShotPath = "E:\\LillyWeb\\CialisMD\\PreScreenShots\\CRM\\Uploaded\\";
            cd.ExcelPathAndFile = "E:\\LillyWeb\\CialisMD\\PreScreenShots\\IE9\\Uploaded\\list.xlsx";
            cd.HeaderPage = "E:\\LillyWeb\\CialisMD\\PreScreenShots\\IE9\\Page1.docx";
            cd.CreateDoc();
        }
    }
}
