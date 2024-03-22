using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace MailSender
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("MailSender")]
    [ComVisible(true)]
    public class Business
    {
        public string Host { get; set; }
        public int Porta { get; set; }
        public string Usuario { get; set; }
        public string Senha { get; set; }
        public string Remetente { get; set; }
        public string Destinatarios { get; set; }
        public string Assunto { get; set; }
        public string Mensagem { get; set; }
        public string Anexos { get; set; }
        public string SendEmail()
        {
            string resultado = "";

            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                SmtpClient client = new SmtpClient();
                client.Host = Host;
                client.Port = Porta;
                client.EnableSsl = true;
                client.Credentials = new NetworkCredential(Usuario, Senha);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                //client.UseDefaultCredentials = true;
                MailMessage mail = new MailMessage();
                mail.Sender = new MailAddress(Remetente, Assunto);
                mail.From = new MailAddress(Remetente, Assunto);

                if (Anexos != null)
                {
                    foreach (var item in Anexos.Split(';'))
                    {
                        mail.Attachments.Add(new Attachment(item));
                    }
                }

                foreach (var item in Destinatarios.Split(';'))
                {
                    mail.To.Add(new MailAddress(item));
                }

                mail.Subject = Assunto;
                mail.Body = Mensagem;
                mail.IsBodyHtml = true;
                mail.Priority = MailPriority.High;
                //
                client.Send(mail);
                resultado = "Email enviado com sucesso!";
            }
            catch (Exception ex)
            {
                resultado = $"Falha ao enviar e-mail {ex.Message} {ex.InnerException}";
            }
            return resultado;
        }
    }
}
