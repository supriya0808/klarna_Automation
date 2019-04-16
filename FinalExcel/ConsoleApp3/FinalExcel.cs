using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Net.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;



namespace ConsoleApp3
{
    class Program
    {
        public  static Excel.Application xlApp;
        public static Excel.Workbook xlWorkBook;
        public static Excel.Worksheet xlWorkSheet;
        public static Excel.Range range;

        static int columnCnt = 0;
        static int rowCnt = 0;
        static int output = 0;
        static int output1 = 0;

        public static  string CsvFileName = ConfigurationManager.AppSettings["csvFileName"];
        public static string TsvFileName = ConfigurationManager.AppSettings["tsvFileName"];
        public static string ExcelFileName = ConfigurationManager.AppSettings["excelFileName"];


        static void Main(string[] args)
        {



            FileInfo fi1 = new FileInfo(CsvFileName);
            Delete(TsvFileName);
            ConvertCSVtoTabDelimited(fi1);


            FileInfo fi2 = new FileInfo(TsvFileName);
            Delete(ExcelFileName);
            ConvertTSVtoEXCEL(fi2);



            if (File.Exists(ExcelFileName))
            {
                string filenam = ConfigurationManager.AppSettings["MyPath1"] + DateTime.Now.ToString("dd_MM_yyyy") + ".xlsx";
                Delete(filenam);
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(ExcelFileName);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;
                columnCnt = range.Columns.Count;
                rowCnt = range.Rows.Count;



              int IndexOfCol1 = DeleteColumn("TargetTeam");
              int IndexOfCol2 = DeleteColumn("batchNumber");

                if (IndexOfCol1 > 0 && IndexOfCol2 > 0)
                {
                    ((Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Columns[IndexOfCol1]).EntireColumn.Delete(null);

                    ((Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Columns[IndexOfCol2-1]).EntireColumn.Delete(null);

                    int FormatCol1 = FormatColumn("InvoiceNumber");
                    int FormatCol2 = FormatColumn("ReceiptId");
                    int FormatCol3 = FormatColumn("dateentered");
                    int FormatCol4 = FormatColumn("VoidHeaderId");


                
                    

                    if (FormatCol1 >0 && FormatCol2 >0 && FormatCol3 >0 && FormatCol4 >0)
                    {

                        xlWorkSheet.Columns[FormatCol1].NumberFormat = Constants.Format;
                        xlWorkSheet.Columns[FormatCol2].NumberFormat = Constants.Format;
                        xlWorkSheet.Columns[FormatCol3].NumberFormat = Constants.FormatDate;
                        xlWorkSheet.Columns[FormatCol4].NumberFormat = Constants.Format;

                        Display();

                        string filename = ConfigurationManager.AppSettings["MyPath1"] + DateTime.Now.ToString("dd_MM_yyyy");
                        xlWorkBook.SaveAs(filename + ".xlsx");

                        Marshal.ReleaseComObject(range);

                        Marshal.ReleaseComObject(xlWorkSheet);//close and release

                        xlWorkBook.Close();

                        Marshal.ReleaseComObject(xlWorkBook);//quit and release

                        xlApp.Quit();

                        Marshal.ReleaseComObject(xlApp);

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

        public static void Delete(string fn)
        {
            if (File.Exists(fn))
            {
                File.Delete(fn);
            }
        }


        private static void ConvertCSVtoTabDelimited(FileInfo fi)// csv-tsv
        {
            try
            {
                string NewFileName = Path.Combine(Path.GetDirectoryName(fi.FullName), Path.GetFileNameWithoutExtension(fi.FullName) + ".tsv");
                System.IO.File.WriteAllText(NewFileName, System.IO.File.ReadAllText(fi.FullName).Replace(",", "\t"));
            }
            catch (Exception ex)
            {
                Console.WriteLine("File: " + fi.FullName + Environment.NewLine + Environment.NewLine + ex.ToString(), "Error Converting File");
            }
        }

        public static void ConvertTSVtoEXCEL(FileInfo fi1)
        {
            string worksheetsName = "KLARNA";


            var format = new ExcelTextFormat();
            format.Delimiter = '\t';
            format.EOL = "\r";

            using (ExcelPackage package = new ExcelPackage(new FileInfo(ExcelFileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
                worksheet.Cells["A1"].LoadFromText(new FileInfo(TsvFileName), format);
                package.Save();
            }


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


//public async Task sendemail(string Toemail, string filePath)
//{
//    try
//    {
//        await Task.Run(() =>
//        {
//            MailMessage mail = new MailMessage();
//            SmtpClient SmtpServer = new SmtpClient("smtp.sendgrid.net");
//            mail.From = new MailAddress("xxxxxxxx@gmail.com");
//            mail.To.Add(Toemail);
//            mail.Subject = "Test Mail - 1";
//            mail.Body = "mail with attachment";
//            System.Net.Mail.Attachment attachment;
//            attachment = new System.Net.Mail.Attachment(filePath);
//            mail.Attachments.Add(attachment);
//            SmtpServer.Port = 25;
//            SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["ApiKey"], ConfigurationManager.AppSettings["ApiKeyPass"]);
//            SmtpServer.EnableSsl = true;
//            SmtpServer.Send(mail);
//        });
//    }
//    catch (Exception ex)
//    {
//        throw ex;
//    }
//}

//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using Outlook = Microsoft.Office.Interop.Outlook;
//using System.Configuration;
//using Office = Microsoft.Office.Core;
//using Excel = Microsoft.Office.Interop.Excel;
//using System.Net;
//using System.Net.Mail;

//namespace mailsending
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {


//            try
//            {

//                Outlook.Application App = new Outlook.Application();

//                Outlook.MailItem mail = (Outlook.MailItem)App.CreateItem(Outlook.OlItemType.olMailItem);
//                Microsoft.Office.Interop.Outlook._MailItem oMailItem = (Microsoft.Office.Interop.Outlook._MailItem)App.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
//                oMailItem.Subject = "Subject of Mail";
//                //MailMessage message = new MailMessage(from, to);
//                mail.HTMLBody = "Hi Klarna Support,In the attached spreadsheet please find today's list of possibly failed refunds.As usual please could you review each of these and let us know if they are at a failed or success status so we can actionaccordingly.Any refunds which have failed we will retry on our side and any refunds which are successful we will set to completewithin Back Office.Please endeavor to reply back to this email within 24 hours so we can action accordingly.If you are not able to find the customer account in both BO and ASOS Report, then drop an email like below. ";
//                //SmtpClient client = new SmtpClient("smtp-mail.outlook.com");
//                String sDisplayName = "MyAttachment";
//                int iPosition = (int)mail.Body.Length + 1;
//                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;

//                Outlook.Attachment Attach = mail.Attachments.Add
//                                             (@"C:\Users\764986\Desktop\ye.xlxs", iAttachType, iPosition, sDisplayName);

//                mail.Subject = "Your Subject will go here.";

//                Outlook.Recipients oRecips = (Outlook.Recipients)mail.Recipients;

//                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("Supriya.Thakur@cognizant.com");
//                mail.To = "me@abc.com;test@def.com";
//                //mail.Cc = "con@def.com";//All the mail lists have to be separated by the ';'
//                //oMailItem.Body = "Body of the mail";
//                //oMailItem.To = lblEmailIDValue.Text.Trim();
//                oMailItem.CC = "sample@gmail.com";
//                //oMailItem.Send();
//                oRecip.Resolve();

//                mail.Send();

//                oRecip = null;
//                oRecips = null;
//                mail = null;
//                App = null;
//            }
//            catch
//            {

//            }
//        }
//        }
//}
////using System;

////// You will need to add a reference to this library:
////using System.Net.Mail;

////namespace SmtpMailConnections
////{
////    public class OutlookDotComMail
////    {
////        string _sender = "";
////        string _password = "";
////        public OutlookDotComMail(string sender, string password)
////        {
////            _sender = sender;
////            _password = password;
////        }

////        public void SendMail(string recipient, string subject, string message)
////        {
////            SmtpClient client = new SmtpClient("smtp-mail.outlook.com");

////            client.Port = 587;
////            client.DeliveryMethod = SmtpDeliveryMethod.Network;
////            client.UseDefaultCredentials = false;
////            System.Net.NetworkCredential credentials =
////                new System.Net.NetworkCredential(_sender, _password);
////            client.EnableSsl = true;
////            client.Credentials = credentials;

////            try
////            {
////                var mail = new MailMessage(_sender.Trim(), recipient.Trim());
////                mail.Subject = subject;
////                mail.Body = message;
////                client.Send(mail);
////            }
////            catch (Exception ex)
////            {
////                Console.WriteLine(ex.Message);
////                throw ex;
////            }
////        }
////    }
////}}


////public class simpletest
////{
////    public async Task sendemail(string Toemail, string filePath)
////    {
////        try
////        {
////            await Task.Run(() =>
////            {
////                MailMessage mail = new MailMessage();
////                SmtpClient SmtpServer = new SmtpClient("smtp.sendgrid.net");
////                mail.From = new MailAddress("xxxxxxxx@gmail.com");
////                mail.To.Add(Toemail);
////                mail.Subject = "Test Mail - 1";
////                mail.Body = "mail with attachment";
////                System.Net.Mail.Attachment attachment;
////                attachment = new System.Net.Mail.Attachment(filePath);
////                mail.Attachments.Add(attachment);
////                SmtpServer.Port = 25;
////                SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["ApiKey"], ConfigurationManager.AppSettings["ApiKeyPass"]);
////                SmtpServer.EnableSsl = true;
////                SmtpServer.Send(mail);
////            });
////        }
////        catch (Exception ex)
////        {
////            throw ex;
////        }
////    }
////}


////MailAddress from = new MailAddress("b@test.com", "B");
////MailAddress to = new MailAddress("j@test.com", "J");
////MailMessage message = new MailMessage(from, to);
////// message.Subject = "Using the SmtpClient class.";
////message.Subject = "Using the SmtpClient class.";
////message.Body = @"Using this feature, you can send an e-mail message from an application very easily.";
////// Add a carbon copy recipient.
////MailAddress copy = new MailAddress("N@test.com");
////message.CC.Add(copy);
////MailAddress Bcopy = new MailAddress("L@test.com");
////message.BCC.Add(Bcopy);
////SmtpClient client = new SmtpClient(server);
////// Include credentials if the server requires them.
////client.Credentials = CredentialCache.DefaultNetworkCredentials;
////Console.WriteLine("Sending an e-mail message to {0} by using the SMTP host {1}.", to.Address, client.Host);

////      try {
////        client.Send(message);
////      }
////      catch (Exception ex) {
////        Console.WriteLine("Exception caught in CreateCopyMessage(): {0}", 
////                    ex.ToString() );
////  	  }