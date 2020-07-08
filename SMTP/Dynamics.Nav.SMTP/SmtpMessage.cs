using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Microsoft.Dynamics.Nav.SMTP
{
    public class SmtpMessage : IDisposable
    {
        private const string Success = "";

        private MailMessage message;

        private SmtpClient smtpClient;

        private string fromName = string.Empty;

        private string fromAddress = string.Empty;

        private List<LinkedResource> linkedResources = new List<LinkedResource>();

        public enum MailPriorities : ushort
        {
            Normal = 0,
            Low = 1,
            High = 2
        }

        public MailPriorities Priority
        {
            get
            {
                return (MailPriorities)this.message.Priority;
            }
            set
            {
                this.message.Priority = (MailPriority)value;
            }
        }

        public string ReplyTo
        {
            set
            {
                this.message.ReplyToList.Clear();
                if (!string.IsNullOrEmpty(value))
                {
                    this.message.ReplyToList.Add(SmtpMessage.FormatToComma(value));
                }
            }
        }

        public string Bcc
        {
            get
            {
                return SmtpMessage.FormatToSemicolon(this.message.Bcc.ToString());
            }
            set
            {
                this.message.Bcc.Clear();
                if (!string.IsNullOrEmpty(value))
                {
                    this.message.Bcc.Add(SmtpMessage.FormatToComma(value));
                }
            }
        }

        public string Body
        {
            get
            {
                return this.message.Body;
            }
            set
            {
                this.message.Body = value;
            }
        }

        public string CC
        {
            get
            {
                return SmtpMessage.FormatToSemicolon(this.message.CC.ToString());
            }
            set
            {
                this.message.CC.Clear();
                if (!string.IsNullOrEmpty(value))
                {
                    this.message.CC.Add(SmtpMessage.FormatToComma(value));
                }
            }
        }

        public string FromAddress
        {
            get
            {
                return this.fromAddress;
            }
            set
            {
                this.fromAddress = value;
            }
        }

        public string FromName
        {
            get
            {
                return this.fromName;
            }
            set
            {
                this.fromName = value;
            }
        }

        public bool HtmlFormatted
        {
            get
            {
                return this.message.IsBodyHtml;
            }
            set
            {
                this.message.IsBodyHtml = value;
            }
        }

        public string Subject
        {
            get
            {
                return this.message.Subject;
            }
            set
            {
                this.message.Subject = value;
            }
        }

        public int Timeout
        {
            get
            {
                return this.smtpClient.Timeout;
            }
            set
            {
                this.smtpClient.Timeout = value;
            }
        }

        public string To
        {
            get
            {
                return SmtpMessage.FormatToSemicolon(this.message.To.ToString());
            }
            set
            {
                this.message.To.Clear();
                if (!string.IsNullOrEmpty(value))
                {
                    this.message.To.Add(SmtpMessage.FormatToComma(value));
                }
            }
        }

        public SmtpMessage()
        {
            MailMessage mailMessage = new MailMessage()
            {
                BodyEncoding = Encoding.UTF8,
                SubjectEncoding = Encoding.UTF8
            };
            this.message = mailMessage;
            this.smtpClient = new SmtpClient()
            {
                DeliveryFormat = SmtpDeliveryFormat.International
            };
        }

        public SmtpMessage(Encoding bodyEncoding, Encoding subjectEncoding)
        {
            MailMessage mailMessage = new MailMessage()
            {
                BodyEncoding = bodyEncoding,
                SubjectEncoding = subjectEncoding
            };
            this.message = mailMessage;
            this.smtpClient = new SmtpClient()
            {
                DeliveryFormat = SmtpDeliveryFormat.International
            };
        }

        public string AddAttachment(Stream attachmentStream, string attachmentName)
        {
            string str;
            try
            {
                this.message.Attachments.Add(new Attachment(attachmentStream, attachmentName));
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AddAttachments(string files)
        {
            string str;
            try
            {
                files = SmtpMessage.FormatToSemicolon(files);
                string[] strArrays = files.Split(new char[] { ';' });
                for (int i = 0; i < (int)strArrays.Length; i++)
                {
                    string str1 = strArrays[i];
                    if (!string.IsNullOrEmpty(str1.Trim()))
                    {
                        this.message.Attachments.Add(new Attachment(str1));
                    }
                }
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AddAttachmentWithName(string files, string fileNames)
        {
            string str;
            try
            {
                char[] chrArray = new char[] { ';' };
                string[] strArrays = files.Split(chrArray);
                chrArray = new char[] { ';' };
                string[] strArrays1 = fileNames.Split(chrArray);
                bool length = (int)strArrays.Length == (int)strArrays1.Length;
                for (int i = 0; i < (int)strArrays.Length; i++)
                {
                    string str1 = strArrays[i].Trim();
                    if (!string.IsNullOrEmpty(str1))
                    {
                        if ((!length ? true : string.IsNullOrWhiteSpace(strArrays1[i])))
                        {
                            this.message.Attachments.Add(new Attachment(str1));
                        }
                        else
                        {
                            AttachmentCollection attachments = this.message.Attachments;
                            Attachment attachment = new Attachment(str1)
                            {
                                Name = strArrays1[i].Trim()
                            };
                            attachments.Add(attachment);
                        }
                    }
                }
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AddReplyTo(string recipients)
        {
            string str;
            try
            {
                this.message.ReplyToList.Add(SmtpMessage.FormatToComma(recipients));
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AddBCC(string recipients)
        {
            string str;
            try
            {
                this.message.Bcc.Add(SmtpMessage.FormatToComma(recipients));
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AddBCCWithDisplayName(string recipientname, string recipientaddress)
        {
            string str;
            try
            {
                this.message.Bcc.Add(string.Format("{0} <{1}>", recipientname, SmtpMessage.RemoveSemicolonComma(recipientaddress)));
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AddCC(string recipients)
        {
            string str;
            try
            {
                this.message.CC.Add(SmtpMessage.FormatToComma(recipients));
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AddCCWithDisplayName(string recipientname, string recipientaddress)
        {
            string str;
            try
            {
                this.message.CC.Add(string.Format("{0} <{1}>", recipientname, SmtpMessage.RemoveSemicolonComma(recipientaddress)));
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AddRecipients(string recipients)
        {
            string str;
            try
            {
                this.message.To.Add(SmtpMessage.FormatToComma(recipients));
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AddRecipientsWithDisplayName(string recipientname, string recipientaddress)
        {
            string str;
            try
            {
                this.message.To.Add(string.Format("{0} <{1}>", recipientname, SmtpMessage.RemoveSemicolonComma(recipientaddress)));
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AddHeader(string name, string value)
        {
            string str;
            try
            {
                this.message.Headers.Add(name, value);
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AppendBody(string text)
        {
            string str;
            try
            {
                MailMessage mailMessage = this.message;
                mailMessage.Body = string.Concat(mailMessage.Body, text);
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AddAlternateView(string content, string contenttype)
        {
            string str;
            if (string.IsNullOrWhiteSpace(content) || string.IsNullOrWhiteSpace(contenttype))
                return string.Empty;

            try
            {
                AlternateView alternate;
                if (string.IsNullOrWhiteSpace(contenttype))
                {
                    alternate = AlternateView.CreateAlternateViewFromString(content);
                }
                else
                {
                    ContentType mimeType = new System.Net.Mime.ContentType(contenttype);
                    alternate = AlternateView.CreateAlternateViewFromString(content, mimeType);
                }

                foreach (LinkedResource linkedResource in linkedResources)
                    alternate.LinkedResources.Add(linkedResource);

                MailMessage mailMessage = this.message;
                mailMessage.AlternateViews.Add(alternate);
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string DeliveryNotification(DeliveryNotificationOptions value)
        {
            string str;
            try
            {
                this.message.DeliveryNotificationOptions = value;
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AddLinkedResource(Stream resourceStream, string contentType, string contentId)
        {
            string str;
            try
            {
                LinkedResource linkedResource = new LinkedResource(resourceStream, new ContentType(contentType));
                linkedResource.ContentId = contentId;
                linkedResources.Add(linkedResource);
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }

        public string AddLinkedResources(string files, string contentType, string contentId)
        {
            string str;
            if (string.IsNullOrWhiteSpace(files))
                return string.Empty;

            try
            {
                LinkedResource linkedResource = new LinkedResource(files, new ContentType(contentType));
                linkedResource.ContentId = contentId;
                linkedResources.Add(linkedResource);
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }
        
        public string ConvertImageToBase64String(string files)
        {
            string base64String = string.Empty;
            string contentType = string.Empty;

            if (string.IsNullOrWhiteSpace(files))
                return string.Empty;

            try
            {
                if (!File.Exists(files))
                    throw new Exception(string.Format("File {0} does not exists", files));

                contentType = string.Format("data:image/{0};base64,", Path.GetExtension(files).TrimStart('.').ToLower());
                using (Image _image = Image.FromFile(files))
                {
                    using (MemoryStream _mStream = new MemoryStream())
                    {
                        _image.Save(_mStream, _image.RawFormat);
                        byte[] _imageBytes = _mStream.ToArray();
                        base64String = Convert.ToBase64String(_imageBytes);
                    }
                }
            }
            catch (Exception exception)
            {
                throw new Exception(SmtpMessage.ErrorMessage(exception));
            }
            return string.Format("{0}{1}", contentType, base64String).TrimStart().TrimEnd();
        }

        public bool ConvertBase64ImagesToContentId()
        {
            if (!this.HtmlFormatted || this.Body == null)
                return true;
            Regex regex = new Regex("data:(.*);base64,(.*)");
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.PreserveWhitespace = true;
            xmlDocument.XmlResolver = (XmlResolver)null;
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.DtdProcessing = DtdProcessing.Prohibit;
            settings.XmlResolver = (XmlResolver)null;
            using (Stream input = (Stream)new MemoryStream(Encoding.UTF8.GetBytes(this.Body)))
            {
                using (XmlReader reader = XmlReader.Create(input, settings))
                {
                    try
                    {
                        xmlDocument.Load(reader);
                    }
                    catch (XmlException ex)
                    {
                        return false;
                    }
                }
            }
            XmlNodeList xmlNodeList = xmlDocument.SelectNodes("//img");
            if (xmlNodeList == null)
                return true;
            bool flag = true;
            for (int index = 0; index < xmlNodeList.Count; ++index)
            {
                XmlElement xmlElement = (XmlElement)xmlNodeList[index];
                if (xmlElement.HasAttribute("src"))
                {
                    string attribute = xmlElement.GetAttribute("src");
                    Match match = regex.Match(attribute);
                    if (!string.IsNullOrEmpty(match.Value))
                    {
                        string str1 = match.Groups[1].Value;
                        string s = match.Groups[2].Value;
                        byte[] bytes;
                        try
                        {
                            bytes = Convert.FromBase64String(s);
                        }
                        catch (FormatException ex)
                        {
                            flag = false;
                            continue;
                        }
                        string str2 = Path.ChangeExtension(Path.GetTempFileName(), "jpg");
                        try
                        {
                            System.IO.File.WriteAllBytes(str2, bytes);
                        }
                        catch (IOException ex)
                        {
                            flag = false;
                            continue;
                        }
                        string withoutExtension = Path.GetFileNameWithoutExtension(str2);
                        Attachment attachment = new Attachment(str2);
                        attachment.ContentDisposition.Inline = true;
                        attachment.ContentDisposition.DispositionType = "inline";
                        attachment.ContentId = withoutExtension;
                        attachment.ContentType.MediaType = str1;
                        attachment.ContentType.Name = str2;
                        this.message.Attachments.Add(attachment);
                        xmlElement.SetAttribute("src", string.Format((IFormatProvider)CultureInfo.InvariantCulture, "cid:{0}", (object)withoutExtension));
                    }
                }
            }
            this.Body = xmlDocument.OuterXml;
            return flag;
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposedManagedResources)
        {
            if (disposedManagedResources)
            {
                if (this.message != null)
                {
                    this.message.Dispose();
                    this.message = null;
                }
                if (this.smtpClient != null)
                {
                    this.smtpClient.Dispose();
                    this.smtpClient = null;
                }
            }
        }

        private static string ErrorMessage(Exception e)
        {
            string str;
            str = (null == e.InnerException ? e.Message : string.Concat(e.Message, Environment.NewLine, e.InnerException.Message));
            return str;
        }

        private static string FormatToComma(string adresses)
        {
            return adresses.Replace(";", ",");
        }

        private static string FormatToSemicolon(string adresses)
        {
            return adresses.Replace(",", ";");
        }

        private static string RemoveSemicolonComma(string adresses)
        {
            return adresses.Replace(",", "").Replace(";", "");
        }

        public string Send(string server, int serverPort, bool authentication, string userName, string password, bool enableSsl)
        {
            string str;
            try
            {
                this.message.From = new MailAddress(this.FromAddress, this.FromName);
                this.smtpClient.Host = server;
                this.smtpClient.Port = serverPort;
                this.smtpClient.EnableSsl = enableSsl;
                if (!authentication)
                {
                    this.smtpClient.Credentials = null;
                    this.smtpClient.UseDefaultCredentials = false;
                }
                else if (string.IsNullOrEmpty(userName))
                {
                    this.smtpClient.Credentials = null;
                    this.smtpClient.UseDefaultCredentials = true;
                }
                else
                {
                    this.smtpClient.UseDefaultCredentials = false;
                    this.smtpClient.Credentials = new NetworkCredential(userName, password);
                }
                this.ConvertBase64ImagesToContentId();
                this.smtpClient.Send(this.message);
                str = "";
            }
            catch (Exception exception)
            {
                str = SmtpMessage.ErrorMessage(exception);
            }
            return str;
        }
    }
}