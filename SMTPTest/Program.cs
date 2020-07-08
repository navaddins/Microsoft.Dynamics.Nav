using Microsoft.Dynamics.Nav.SMTP;
namespace SMTPTest
{
    class Program
    {
        static void Main(string[] args)
        {
            SmtpMessage _smtpMessage = new SmtpMessage();

            string _mailserver = "smtp.office365.com";
            int _mailport = 587;
            string _username = "";
            string _password = "";

            bool _usessl = false;
            bool _useHtml = false;

            string _fromName = "FromName";
            string _fromAddress = "From@outlook.com";

            string _toname = "ToName";
            string _toaddress = "To@outlook.com";

            string _bccname = "BccName";
            string _bccaddress = "Bcc@outlook.com";

            string _ccname = "CcName";
            string _ccaddress = "Cc@outlook.com";

            string _replytoaddress = "ReplyTo@outlook.com"; // Reply to this email if other party click on reply
            string _notifyonreceiptname = "Disposition-Notification-To";
            string _notifyonreceiptaddress = "NotifyTo@outlook.com"; // Read Message will reply to this email address

            string _attachfile = "";

            string _subject = "This is email subject";
            string _bodytext = "This is some plain text.";

            string _bodytextimage = "<img src=\"{0}\" alt=\"Girl in a jacket\" />";
            string _bodytextimages = string.Empty;
            _bodytextimages += string.Format(_bodytextimage, _smtpMessage.ConvertImageToBase64String(@""));
            _bodytextimages += string.Format(_bodytextimage, _smtpMessage.ConvertImageToBase64String(@""));

            string _content = "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">";
            _content += "<HTML><HEAD><META http-equiv=Content-Type content=\"image/jpeg; charset=iso-8859-1\">";
            _content += "</HEAD><BODY><DIV><FONT face=Arial color=#ff0000 size=2>This is some html text.";
            _content += "</FONT><br><img src=cid:Image1 alt=\"Girl in a jacket\"></img><br><img src=cid:Image2 alt=></img></DIV></BODY></HTML>";

            string _contentType = "text/html";
            SmtpMessage.MailPriorities _mailPriorities = SmtpMessage.MailPriorities.High;


            // Sender Name and Address
            _smtpMessage.FromName = _fromName;
            _smtpMessage.FromAddress = _fromAddress;


            // Send To Name and Address
            if (string.IsNullOrWhiteSpace(_toname))
                _smtpMessage.AddRecipients(_toaddress);
            else
                _smtpMessage.AddRecipientsWithDisplayName(_toname, _toaddress);

            // Cc
            if (string.IsNullOrWhiteSpace(_ccname))
                _smtpMessage.AddCC(_ccaddress);
            else
                _smtpMessage.AddCCWithDisplayName(_ccname, _ccaddress);

            // Bcc
            if (string.IsNullOrWhiteSpace(_bccname))
                _smtpMessage.AddBCC(_bccaddress);
            else
                _smtpMessage.AddBCCWithDisplayName(_bccname, _bccaddress);

            if (!string.IsNullOrWhiteSpace(_replytoaddress))
                _smtpMessage.AddReplyTo(_replytoaddress);

            // Add Header
            // For Receipt return message
            if (!string.IsNullOrWhiteSpace(_notifyonreceiptaddress))
                _smtpMessage.AddHeader(_notifyonreceiptname, _notifyonreceiptaddress);

            // Attach the file
            if (!string.IsNullOrWhiteSpace(_attachfile))
            {
                if (System.IO.File.Exists(_attachfile))
                {
                    _smtpMessage.AddAttachments(_attachfile);
                }
            }

            // Body
            _smtpMessage.Body = _bodytext;
            _smtpMessage.AppendBody(_bodytextimages);

            // EMail will be receive as html if we client is support html format otherwise will get plain text
            // You can skip below AddLinkResources and AddAlternateView if your email body is html and you must set _useHtml to true
            // Add Linked Resource for alternate body as HTML
            _smtpMessage.AddLinkedResources(@"", "image/jpeg", "Image1");
            _smtpMessage.AddLinkedResources(@"", "image/jpeg", "Image2");

            // The alternate body as HTML
            if (!string.IsNullOrWhiteSpace(_content))
                _smtpMessage.AddAlternateView(_content, _contentType);

            _smtpMessage.Subject = _subject;
            _smtpMessage.HtmlFormatted = _useHtml;
            _smtpMessage.Priority = _mailPriorities;
            string _result = _smtpMessage.Send(_mailserver, _mailport, true, _username, _password, _usessl);

            System.Console.WriteLine(_result);
            if (!string.IsNullOrWhiteSpace(_result))
                System.Console.ReadLine();
        }
    }
}
