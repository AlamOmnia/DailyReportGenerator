using System;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;

namespace Demo_Excel_Export
{
    public partial class runMonthly : Form
    {
        public runMonthly()
        {
            InitializeComponent();
            this.startDate.Value = DateTime.Today.AddDays(-1);
            this.endDate.Value = DateTime.Today;
        }

        private void run_Click(object sender, EventArgs e)
            
        {
            DateTime start = startDate.Value.AddDays(-1);
            DateTime end =  endDate.Value.AddHours(5);
            DateTime ans1 = startDate.Value; 
            DateTime ans2 = endDate.Value;
            Daily da = new Daily();
            da.ExportReport(start,end,ans1,ans2);
        }

        private void sendMail_Click(object sender, EventArgs e)
        {
            try
            {

                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress("omnia.alam@gmail.com");
                    mail.To.Add("omnia.alam@gmail.com");
                    mail.Subject = "Test-1";
                    mail.Body = "<h1>Hello</h1>";
                    mail.IsBodyHtml = true;
                    mail.Attachments.Add(new Attachment("D:\\Daily International incoming and outgoingTraffic Report of Purple ICX for BTRC(15-Aug).xlsx"));

                    using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                    {
                        smtp.EnableSsl = true;
                        smtp.UseDefaultCredentials = false;
                        smtp.Credentials = new NetworkCredential("omnia.alam@gmail.com", "poltadg45id");
                        smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                        smtp.Send(mail);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }



        }

      
    }
}
