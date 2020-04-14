using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using SBO.Hub.Models;
using System;
using System.Collections.Generic;

namespace SBO.Hub.Helpers
{
    public class EmailHelper
    {
        SmtpClient Client;
        public bool Connected { get; set; } = false;
        private EmailConfigurationModel EmailConfigModel;

        public EmailHelper(EmailConfigurationModel emailConfigModel)
        {
            EmailConfigModel = emailConfigModel;
            if (EmailConfigModel == null || String.IsNullOrEmpty(EmailConfigModel.Server))
            {
                throw new Exception("Configuração de envio de e-mail não encontrada, favor verifique o cadastro");
            }
        }

        public void Connect()
        {
            Client = new SmtpClient();

            SecureSocketOptions secure = EmailConfigModel.SSL == "Y" ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.None;

            Client.Connect(EmailConfigModel.Server, EmailConfigModel.Port, SecureSocketOptions.StartTls);
            Client.Authenticate(EmailConfigModel.Email, EmailConfigModel.Password);
            Connected = true;
        }

        public void Disconnect()
        {
            Client.Disconnect(true);
            Client.Dispose();
            Connected = false;
        }

        public string SendEmail(string subject, string body, string email, List<string> attachments = null)
        {
            string msg = String.Empty;
            try
            {
                var message = new MimeMessage();

                message.From.Add(new MailboxAddress(EmailConfigModel.Email));
                message.To.Add(new MailboxAddress(email));
                message.Subject = subject;

                BodyBuilder bodyBuilder = new BodyBuilder();
                bodyBuilder.HtmlBody = body;

                if (attachments != null)
                {
                    foreach (var att in attachments)
                    {
                        bodyBuilder.Attachments.Add(att);
                    }
                }

                message.Body = bodyBuilder.ToMessageBody();

                Client.Send(message);
            }
            catch (Exception ex)
            {
                msg = ex.Message;
            }
            return msg;
        }
    }
}
