using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
//using System.Windows.Forms;
using System.Net.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;



namespace ConsoleApp3
{
    class Program
    {
        static Excel.Application xlApp;
        static Excel.Workbook xlWorkBook;
        static Excel.Worksheet xlWorkSheet;
        static Excel.Range range;

        static int columnCnt = 0;
        static int rowCnt = 0;
        static int output = 0;
        static int output1 = 0;



        static void Main(string[] args)
        {

            string connecString = ConfigurationManager.AppSettings["MyPath"];

            if (File.Exists(connecString))
            {
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(connecString);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;
                columnCnt = range.Columns.Count;
                rowCnt = range.Rows.Count;



              int IndexOfCol1 = DeleteColumn("id");
              int IndexOfCol2 = DeleteColumn("age");

             // Console.WriteLine(IndexOfCol1);
             //Console.WriteLine(IndexOfCol2);

                if (IndexOfCol1 > 0 && IndexOfCol2 > 0)
                {
                    ((Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Columns[IndexOfCol1]).EntireColumn.Delete(null);

                    ((Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Columns[IndexOfCol2-1]).EntireColumn.Delete(null);

                    int FormatCol1 = FormatColumn("invoice");
                    int FormatCol2 = FormatColumn("date1");
                    

                    //Console.WriteLine(FormatCol1);
                    //Console.WriteLine(FormatCol2);

                    if (FormatCol1 > 0 && FormatCol2 > 0)
                    {

                        xlWorkSheet.Columns[FormatCol1].NumberFormat = Constants.FormatInvoice;
                        xlWorkSheet.Columns[FormatCol2].NumberFormat = Constants.FormatDate;

                        Display();

                        string filename = ConfigurationManager.AppSettings["MyPath1"] + DateTime.Now.ToString("dd_MM_yyyy");
                        xlWorkBook.SaveAs(filename + ".xlsx");

                        Marshal.ReleaseComObject(range);

                        Marshal.ReleaseComObject(xlWorkSheet);//close and release

                        xlWorkBook.Close();

                        Marshal.ReleaseComObject(xlWorkBook);//quit and release

                        xlApp.Quit();

                        Marshal.ReleaseComObject(xlApp);




                        //try
                        //{
                        //    MailMessage mail = new MailMessage();
                        //    SmtpClient SmtpServer = new SmtpClient("smtp-mail.outlook.com");
                        //    mail.From = new MailAddress("Supriya.Thakur@cognizant.com");
                        //    mail.To.Add("Tanya.Gupta3@cognizant.com");
                        //    mail.Subject = "KLARNA MAIL";
                        //    mail.Body = "Hi Klarna Support,\n" + " In the attached spreadsheet please find today's list of possibly failed refunds. As usual please could you review each of these and let us know if they are at a failed or success status so we can action accordingly. Any refunds which have failed we will retry on our side and any refunds which are successful we will set to complete within Back Office. Please endeavor to reply back to this email within 24 hours so we can action accordingly. If you are not able to find the customer account in both BO and ASOS Report, then drop an email like below. ";


                        //    System.Net.Mail.Attachment attachment;
                        //    attachment = new System.Net.Mail.Attachment(ConfigurationManager.AppSettings["MyPath1"] + DateTime.Now.ToString("dd_MM_yyyy") + ".xlsx");
                        //    mail.Attachments.Add(attachment);

                        //    SmtpServer.Port = 587;
                        //    SmtpServer.Credentials = new System.Net.NetworkCredential("Supriya.Thakur@cognizant.com", "**********");
                        //    SmtpServer.EnableSsl = true;

                        //    SmtpServer.Send(mail);
                        //    Console.WriteLine("MAIL SEND");
                        //}
                        //catch (Exception ex)
                        //{
                        //    Console.WriteLine("MAIL IS NOT SEND \n" + ex.ToString());
                        //}





                    }
                    else
                    {

                        Console.WriteLine(Constants.FormattingColumnNotExists);
                    }
                }
                else
                {

                    Console.WriteLine(Constants.DeletingColumnNotexists);

                }
            }
            else
            {

                Console.WriteLine(Constants.FileNotExists);

            }
            Console.ReadKey();
        }

        public static void Display()
        {
            for (int row = 1; row <= rowCnt; row++)

            {
                for (int col = 1; col <= columnCnt; col++)
                {
                    if (col == 1)

                        Console.Write("\r\n");

                    if (range.Cells[row, col] != null && range.Cells[row, col].Value2 != null)

                        Console.Write(range.Cells[row, col].Value2.ToString() + "\t");
                }
            }

        }
        public static int DeleteColumn(string str)
        {

            for (int col = 1; col <= columnCnt; col++)
            {

                if (xlWorkSheet.Cells[1, col].value == str)
                {
                    output1 = col;
                    return output1;
                }

            }

            return 0;
        }
        public static int FormatColumn(string str)
        {

            for (int col = 1; col <= columnCnt; col++)

            {
                if (xlWorkSheet.Cells[1, col].value == str)
                {
                    output = col;
                    return output;
                }
            }

            return 0;
        }

    }
}
