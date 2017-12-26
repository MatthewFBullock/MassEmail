using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Net.Mail;


namespace MassEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            //open the excel document
            var excel_directory = @"C:\Users\matth\OneDrive\Alumni Chapter\Program for Texting Alumni\MassEmail\MassEmail\bin\Debug\Portland_Oregon_Alumni_List.xlsx";
            //var excel_directory = @"C:\Users\matth\OneDrive\Alumni Chapter\Program for Texting Alumni\MassEmail\MassEmail\bin\Debug\Test.xlsx";
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excel_directory);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            List<string> Member = new List<string>();

            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1)
                        Console.Write("\r\n");

                    //write the value to the console
                    try
                    {
                        if (xlRange.Cells[i, j] != null)
                            Member.Add(xlRange.Cells[i, j].Value2.ToString());
                    }
                    catch (Exception)
                    {
                        Member.Add("null");
                        continue;
                    }
                    //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    //    Member.Add("null");
                    
                    
                    
                }
                if (Member[7].Equals("null"))
                {
                    Member.Clear();
                    continue;
                }
                if (Member[1].Equals("null") || Member[3].Equals("null") || int.Parse(Member[7]) > 2017)
                {
                    Member.Clear();
                    continue;
                }
                SendMail(Member[1], Member[3]);
                Member.Clear();
            }



            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("Press ENTER to continue...");
            Console.ReadLine();
        }

        private static void SendMail(string email_address, string outbound_name)
        {
            var fromAddress = new MailAddress("matthew.bullock52@gmail.com", "Matthew Bullock");
            var toAddress = new MailAddress(email_address, outbound_name);
            const string fromPassword = "Matt!0614";
            //declares attachment
            System.Net.Mail.Attachment attachment;
            //adds the attachment to memory
            attachment = new System.Net.Mail.Attachment(@"C:\Users\matth\OneDrive\Alumni Chapter\Meetings\PDFs\111417 Meeting Minutes.pdf");
            const string subject = "Delta Chi Portland Alumni Chapter - TopGolf Function Registration";
            string body = "Happy Saturdays " + outbound_name + ",\n\n" +
                "I hope your Thanksgiving holidays treated you well! If you are interested in attending our first function, which was voted on during our first meeting, registration can be done so here: https://goo.gl/forms/LbSH08RiqN9DfOU53 \n\n" + 
                
                "If you have any questions, feel free to reach out to me. Also, be sure to join our Facebook group page if you haven't done so already! https://www.facebook.com/groups/150651385533420/ \n\n" + 
                "ItB,\n\n" +
                "Matt Bullock\n" +
                "(303) 549-9597\n" +
                "Oregon State '17";

            var smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
            };
            try
            {
                using (var message = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = body
                })
                {
                    //adds attachement 
                    //message.Attachments.Add(attachment);
                    //sends the email
                    smtp.Send(message);
                    Console.WriteLine("Message sent to {0} successfully!", outbound_name);
                }
            }
            catch (Exception E)
            {
                var future_datetime = DateTime.Now.AddMinutes(30);
                do
                {
                    var datetime_output = future_datetime - DateTime.Now;
                    Console.WriteLine(datetime_output);
                } while (DateTime.Now < future_datetime);
                //Console.WriteLine(E.Message);
                //Console.WriteLine("Press ENTER to continue...");
                //Console.ReadLine();
            }
            
        }
    }
}
