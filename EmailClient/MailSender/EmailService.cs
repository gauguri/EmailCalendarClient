using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.IO;

namespace EmailCalendarsClient.MailSender
{
    public class EmailService
    {
        private static readonly IReadOnlyDictionary<string, string> KnownMimeTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            [".jpg"] = "image/jpeg",
            [".jpeg"] = "image/jpeg",
            [".png"] = "image/png",
            [".gif"] = "image/gif",
            [".bmp"] = "image/bmp",
            [".tif"] = "image/tiff",
            [".tiff"] = "image/tiff"
        };

        MessageAttachmentsCollectionPage MessageAttachmentsCollectionPage = new MessageAttachmentsCollectionPage();

        public Message CreateStandardEmail(string recipient, string header, string body)
        {
            var message = new Message
            {
                Subject = header,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = body
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = recipient
                        }
                    }
                },
                Attachments = MessageAttachmentsCollectionPage
            };

            return message;
        }

        public Message CreateHtmlEmail(string recipient, string header, string body)
        {
            var message = new Message
            {
                Subject = header,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = body
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = recipient
                        }
                    }
                },
                Attachments = MessageAttachmentsCollectionPage
            };

            return message;
        }

        public FileAttachment AddAttachment(byte[] rawData, string filePath, bool isInline = false, string contentId = null, string contentType = null)
        {
            if (rawData == null)
            {
                throw new ArgumentNullException(nameof(rawData));
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("File path must be provided.", nameof(filePath));
            }

            var attachment = new FileAttachment
            {
                Name = Path.GetFileName(filePath),
                ContentBytes = EncodeTobase64Bytes(rawData),
                ContentType = contentType ?? GetMimeType(filePath)
            };

            if (isInline)
            {
                attachment.IsInline = true;
                attachment.ContentId = string.IsNullOrEmpty(contentId) ? Guid.NewGuid().ToString() : contentId;
            }
            else if (!string.IsNullOrEmpty(contentId))
            {
                attachment.ContentId = contentId;
            }

            MessageAttachmentsCollectionPage.Add(attachment);
            return attachment;
        }

        public FileAttachment AddInlineAttachment(byte[] rawData, string filePath, string contentId = null, string contentType = null)
        {
            return AddAttachment(rawData, filePath, true, contentId, contentType);
        }

        public void ClearAttachments()
        {
            MessageAttachmentsCollectionPage.Clear();
        }

        static public byte[] EncodeTobase64Bytes(byte[] rawData)
        {
            string base64String = System.Convert.ToBase64String(rawData);
            var returnValue = Convert.FromBase64String(base64String);
            return returnValue;
        }

        private static string GetMimeType(string filePath)
        {
            var extension = Path.GetExtension(filePath);

            if (!string.IsNullOrEmpty(extension) && KnownMimeTypes.TryGetValue(extension, out var mimeType))
            {
                return mimeType;
            }

            return "application/octet-stream";
        }
    }
}
