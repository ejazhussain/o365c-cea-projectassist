using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using O365C.Agent.ProjectAssist.Bot.Helpers;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace O365C.Agent.ProjectAssist.Bot.Services
{
    public interface IEmailService
    {
        Task<bool> SendEmailAsync(string accessToken, string fromEmail, string toEmail, string subject, string body);
    }

    public class EmailService : IEmailService
    {
        public async Task<bool> SendEmailAsync(string accessToken, string fromEmail, string toEmail, string subject, string body)
        {
            if (string.IsNullOrWhiteSpace(accessToken))
                throw new ArgumentException("Access token cannot be null or empty.", nameof(accessToken));
            if (string.IsNullOrWhiteSpace(fromEmail))
                throw new ArgumentException("From email cannot be null or empty.", nameof(fromEmail));
            if (string.IsNullOrWhiteSpace(toEmail))
                throw new ArgumentException("To email cannot be null or empty.", nameof(toEmail));
            if (string.IsNullOrWhiteSpace(subject))
                throw new ArgumentException("Subject cannot be null or empty.", nameof(subject));
            if (string.IsNullOrWhiteSpace(body))
                throw new ArgumentException("Body cannot be null or empty.", nameof(body));

            try
            {
                var graphClient = GraphAuthHelper.CreateGraphClientWithAccessToken(accessToken);
                var toRecipients = new List<Recipient>
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = toEmail
                        }
                    }
                };

                var requestBody = new SendMailPostRequestBody
                {
                    Message = new Message
                    {
                        Subject = subject,
                        Body = new ItemBody
                        {
                            ContentType = BodyType.Html,
                            Content = body
                        },
                        ToRecipients = toRecipients
                    },
                    SaveToSentItems = true
                };

                await graphClient.Users[fromEmail].SendMail.PostAsync(requestBody);
                return true;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error in SendEmailAsync: {ex.Message}");
                return false;
            }
        }
    }
}
