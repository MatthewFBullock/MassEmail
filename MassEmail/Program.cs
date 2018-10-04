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
using System.Timers;


namespace MassEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            var CurrentDirectory = Environment.CurrentDirectory;
            //open the excel document
            //var excel_directory = CurrentDirectory + "\\Portland_Oregon_Alumni_List - Copy.xlsx";
            //var excel_directory = CurrentDirectory + "\\Portland_Oregon_Alumni_List.xlsx";
            var excel_directory = CurrentDirectory + "\\Test.xlsx";
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
                SendMail(Member[1], Member[3], ref rowCount);
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

        private static void SendMail(string email_address, string outbound_name, ref int _rowCount)
        {
            var fromAddress = new MailAddress("matthew.bullock52@gmail.com", "Matthew Bullock");
            var toAddress = new MailAddress(email_address, outbound_name);
            const string fromPassword = "Matt!0614";
            //declares attachment
            System.Net.Mail.Attachment attachment;
            //adds the attachment to memory
            attachment = new System.Net.Mail.Attachment(@"C:\Users\matth\Google Drive\Alumni Chapter\Meetings\PDFs\031318 Meeting Minutes.pdf");
            const string subject = "Delta Chi Portland Alumni Chapter - September Monthly Meeting";
            string body = "Brother " + outbound_name + ",\n\n" +
                "We are going to be having our monthly chapter meeting this next Tuesday at 6pm. I look forward to catching up with everyone and sharing the cool stuff we learned at convention, along with resources and direction that has been advised to our chapter since we all last met.\n\n" +
                
                "Location: 1411 SW Morrison St. Suite 200\n\n" +  

                "As always, be sure to join our Facebook group page if you haven't done so already! https://www.facebook.com/groups/150651385533420/ \n\n" + 
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
                //// Create a timer with a two second interval.
                //System.Timers.Timer aTimer = new System.Timers.Timer(15000);
                //// Hook up the Elapsed event for the timer. 
                //aTimer.Elapsed += OnTimedEvent;
                //aTimer.AutoReset = false;
                //aTimer.Enabled = true;
                //Console.WriteLine("The application started at {0:HH:mm:ss}", DateTime.Now);


                DateTime _aTimer = DateTime.Now;
                Timer.ReferenceTimer(ref _aTimer);
                if (true)
                {

                }


                //var future_datetime = DateTime.Now.AddMinutes(30);
                //do
                //{
                //    var datetime_output = future_datetime - DateTime.Now;
                //    Console.WriteLine(datetime_output);
                //} while (DateTime.Now < future_datetime);
                //Console.WriteLine(E.Message);
                //Console.WriteLine("Press ENTER to continue...");
                //Console.ReadLine();

                _rowCount = _rowCount - 1; // sets count back so we don't miss an alumnus!
            }            
        }
    }
}
