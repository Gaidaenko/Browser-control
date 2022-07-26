using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Net.Http;
using System.IO;
using System.Text;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Threading;
using System.Linq;
using OpenQA.Selenium.Support.UI;
using System.Drawing;
using CsQuery.ExtensionMethods.Internal;

namespace Fill_browser_page
{
    public class FileSelected 
    {       
        public void xlsxSelected()
        {
            Excel.Application xlsApp = new Excel.Application();

            Workbook ObjWorkBook = xlsApp.Workbooks.Open(Fields.FileName, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);   
            Worksheet ObjWorkSheet;
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];                                                      

            Excel.Range Rng, CheckingRow;
            Rng = xlsApp.get_Range("A2", "AF20");
            var dataArr = (object[,])Rng.Value;

            CheckingRow = xlsApp.get_Range("B2");

            if (dataArr[Fields.row, 2] == null)
            {
                MessageBox.Show(Fields.statusShowComplete);               
                xlsApp.Quit();
                return;
            }

            if (CheckingRow.Value != null)
            {
                Fields.AuthorName = dataArr[Fields.row, 2].ToString();                                                                    
                Fields.DeliveryAddress = dataArr[Fields.row, 5].ToString();       
                Fields.Description = dataArr[Fields.row, 8].ToString();           
                Fields.TotalWeight = dataArr[Fields.row, 14].ToString();          
                Fields.Cost = dataArr[Fields.row, 15].ToString();                 
                Fields.ItemLength = dataArr[Fields.row, 19].ToString();           
                Fields.ItemWidth = dataArr[Fields.row, 20].ToString();                               
                Fields.ItemHeight = dataArr[Fields.row, 21].ToString();           
                Fields.Recipient = dataArr[Fields.row, 22].ToString();            
                Fields.Locality = dataArr[Fields.row, 27].ToString();             
                Fields.Branch = dataArr[Fields.row, 28].ToString();

                xlsApp.Quit();

                AuthBrowser authBrowser = new AuthBrowser();
                authBrowser.Browser();          
            }               
        }      
    }
}
