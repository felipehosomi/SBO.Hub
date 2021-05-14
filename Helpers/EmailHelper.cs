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
            Client.ServerCertificateValidationCallback = (s, c, h, e) => true;

            SecureSocketOptions secure = EmailConfigModel.SSL == "Y" ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.None;

            Client.Connect(EmailConfigModel.Server, EmailConfigModel.Port, SecureSocketOptions.StartTls);
            Client.Authenticate(EmailConfigModel.Email, EmailConfigModel.Password);
            Connected = true;
        }

        public void Disconnect()
        {
            if (Connected)
            {
                Client.Disconnect(true);
                Client.Dispose();
            }
            Connected = false;
        }

        public string SendEmail(string subject, string body, List<string> emails, List<string> ccEmails = null, List<string> bccEmails = null, List<string> attachments = null)
        {
            string msg = String.Empty;
            try
            {
                var message = new MimeMessage();

                if (!String.IsNullOrEmpty(EmailConfigModel.Name))
                {
                    message.From.Add(new MailboxAddress(EmailConfigModel.Name, EmailConfigModel.Email));
                }
                else
                {
                    message.From.Add(new MailboxAddress(EmailConfigModel.Email));
                }

                foreach (var email in emails)
                {
                    message.To.Add(new MailboxAddress(email.Trim()));
                }
                if (ccEmails != null)
                {
                    foreach (var email in ccEmails)
                    {
                        message.Cc.Add(new MailboxAddress(email.Trim()));
                    }
                }
                if (bccEmails != null)
                {
                    foreach (var email in bccEmails)
                    {
                        message.Bcc.Add(new MailboxAddress(email.Trim()));
                    }
                }

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

        public string SendEmail(string subject, string body, string email, List<string> attachments = null)
        {
            if (String.IsNullOrEmpty(email))
            {
                return "";
            }

            if (!this.Connected)
            {
                this.Connect();
            }
            
            List<string> list = new List<string>();
            list.Add(email);
            return this.SendEmail(subject, body, list, attachments: attachments);
        }
    }
}
