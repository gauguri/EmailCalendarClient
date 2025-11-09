using EmailCalendarsClient.MailSender;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Win32;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace GraphEmailClient
{
    public partial class MainWindow : Window
    {
        AadGraphApiDelegatedClient _aadGraphApiDelegatedClient = new AadGraphApiDelegatedClient();
        EmailService _emailService = new EmailService();

        const string SignInString = "Sign In";
        const string ClearCacheString = "Clear Cache";

        public MainWindow()
        {
            InitializeComponent();
            _aadGraphApiDelegatedClient.InitClient();
        }

        private async void SignIn(object sender = null, RoutedEventArgs args = null)
        {
            var accounts = await _aadGraphApiDelegatedClient.GetAccountsAsync();

            if (SignInButton.Content.ToString() == ClearCacheString)
            {
                await _aadGraphApiDelegatedClient.RemoveAccountsAsync();

                SignInButton.Content = SignInString;
                UserName.Content = "Not signed in";
                return;
            }

            try
            {
                var account = await _aadGraphApiDelegatedClient.SignIn();

                Dispatcher.Invoke(() =>
                {
                    SignInButton.Content = ClearCacheString;
                    SetUserName(account);
                });
            }
            catch (MsalException ex)
            {
                if (ex.ErrorCode == "access_denied")
                {
                    // The user canceled sign in, take no action.
                }
                else
                {
                    // An unexpected error occurred.
                    string message = ex.Message;
                    if (ex.InnerException != null)
                    {
                        message += "Error Code: " + ex.ErrorCode + "Inner Exception : " + ex.InnerException.Message;
                    }

                    MessageBox.Show(message);
                }

                Dispatcher.Invoke(() =>
                {
                    UserName.Content = "Not signed in";
                });
            }
        }

        private async void SendEmail(object sender, RoutedEventArgs e)
        {
            var message = _emailService.CreateStandardEmail(EmailRecipientText.Text, 
                EmailHeader.Text, EmailBody.Text);

            await _aadGraphApiDelegatedClient.SendEmailAsync(message);
            _emailService.ClearAttachments();
        }

        private async void SendHtmlEmail(object sender, RoutedEventArgs e)
        {
            var signature = EmailSignature.Text;
            var body = EmailBody.Text;

            var messageHtml = string.IsNullOrWhiteSpace(signature)
                ? _emailService.CreateHtmlEmail(EmailRecipientText.Text,
                    EmailHeader.Text, body)
                : _emailService.CreateHtmlEmail(EmailRecipientText.Text,
                    EmailHeader.Text, BuildHtmlBody(body, signature));

            await _aadGraphApiDelegatedClient.SendEmailAsync(messageHtml);
            _emailService.ClearAttachments();
        }

        private void AddAttachment(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == true)
            {
                byte[] data = System.IO.File.ReadAllBytes(dlg.FileName);
                _emailService.AddAttachment(data, dlg.FileName);
            }
        }

        private void AddInlineImage(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Image files (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff)|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff|All files (*.*)|*.*"
            };

            if (dlg.ShowDialog() != true)
            {
                return;
            }

            byte[] data;
            try
            {
                data = System.IO.File.ReadAllBytes(dlg.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to read the selected file.\n{ex.Message}", "Inline Image", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            FileAttachment attachment;
            try
            {
                attachment = _emailService.AddInlineAttachment(data, dlg.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to add inline image.\n{ex.Message}", "Inline Image", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (attachment == null || string.IsNullOrEmpty(attachment.ContentId))
            {
                MessageBox.Show("Unable to create inline image attachment.", "Inline Image", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var imgTag = $"<img src=\"cid:{attachment.ContentId}\" alt=\"{attachment.Name}\" />";
            var existingSignature = EmailSignature.Text ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(existingSignature))
            {
                existingSignature = existingSignature.TrimEnd();

                if (!existingSignature.EndsWith("<br/>", StringComparison.OrdinalIgnoreCase))
                {
                    existingSignature += "<br/>";
                }

                existingSignature += Environment.NewLine;
            }

            EmailSignature.Text = existingSignature + imgTag;
            EmailSignature.CaretIndex = EmailSignature.Text.Length;
            EmailSignature.Focus();
        }

        private const int CsvEmailBatchLimit = 20;
        private static readonly TimeSpan DelayBetweenCsvEmails = TimeSpan.FromSeconds(3);

        private async void SendEmailsFromCsv(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                Title = "Select CSV file"
            };

            if (dlg.ShowDialog() != true)
            {
                return;
            }

            int successCount = 0;
            int failureCount = 0;
            var errorSamples = new List<string>();
            int lineNumber = 0;
            var recipientIndex = 0;
            var subjectIndex = 1;
            var bodyIndex = 2;
            var headerEvaluated = false;
            int attemptedSends = 0;
            var batchLimitReached = false;
            var cancellationToken = CancellationToken.None;

            try
            {
                using var parser = new TextFieldParser(dlg.FileName)
                {
                    TextFieldType = FieldType.Delimited,
                    HasFieldsEnclosedInQuotes = true
                };
                parser.SetDelimiters(",");

                while (!parser.EndOfData)
                {
                    lineNumber++;
                    string[] fields;

                    try
                    {
                        fields = parser.ReadFields();
                    }
                    catch (MalformedLineException parseEx)
                    {
                        failureCount++;
                        AppendCsvError(errorSamples, lineNumber, $"Malformed CSV: {parseEx.Message}");
                        continue;
                    }

                    if (fields == null)
                    {
                        continue;
                    }

                    if (!headerEvaluated)
                    {
                        headerEvaluated = true;

                        if (TryMapHeader(fields, out var mappedRecipientIndex, out var mappedSubjectIndex, out var mappedBodyIndex))
                        {
                            if (mappedRecipientIndex >= 0)
                            {
                                recipientIndex = mappedRecipientIndex;
                            }

                            if (mappedSubjectIndex >= 0)
                            {
                                subjectIndex = mappedSubjectIndex;
                            }

                            if (mappedBodyIndex >= 0)
                            {
                                bodyIndex = mappedBodyIndex;
                            }

                            continue;
                        }
                    }

                    if (fields.Length <= Math.Max(recipientIndex, Math.Max(subjectIndex, bodyIndex)))
                    {
                        failureCount++;
                        AppendCsvError(errorSamples, lineNumber, "Expected columns for recipient, subject, and body.");
                        continue;
                    }

                    var recipient = GetFieldValue(fields, recipientIndex)?.Trim();
                    var subject = GetFieldValue(fields, subjectIndex) ?? string.Empty;
                    var body = GetFieldValue(fields, bodyIndex) ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(recipient))
                    {
                        failureCount++;
                        AppendCsvError(errorSamples, lineNumber, "Recipient address is missing.");
                        continue;
                    }

                    try
                    {
                        var signature = EmailSignature.Text;
                        var message = string.IsNullOrWhiteSpace(signature)
                            ? _emailService.CreateStandardEmail(recipient, subject, body)
                            : _emailService.CreateHtmlEmail(recipient, subject, BuildHtmlBody(body, signature));
                        await _aadGraphApiDelegatedClient.SendEmailAsync(message, cancellationToken);
                        successCount++;
                    }
                    catch (Exception sendEx)
                    {
                        failureCount++;
                        AppendCsvError(errorSamples, lineNumber, sendEx.Message);
                    }

                    attemptedSends++;

                    if (attemptedSends >= CsvEmailBatchLimit)
                    {
                        batchLimitReached = !parser.EndOfData;
                        break;
                    }

                    try
                    {
                        await Task.Delay(DelayBetweenCsvEmails, cancellationToken);
                    }
                    catch (TaskCanceledException)
                    {
                        break;
                    }
                }

                var summary = BuildCsvSummary(successCount, failureCount, errorSamples, batchLimitReached);
                MessageBox.Show(summary, "CSV Email Send", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to process CSV file.\n{ex.Message}", "CSV Email Send", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _emailService.ClearAttachments();
            }
        }

        private static void AppendCsvError(ICollection<string> errors, int lineNumber, string message)
        {
            if (errors.Count < 5)
            {
                errors.Add($"Line {lineNumber}: {message}");
            }
        }

        private static string BuildCsvSummary(int successCount, int failureCount, ICollection<string> errorSamples, bool batchLimitReached)
        {
            var builder = new StringBuilder();
            builder.AppendLine($"Emails sent successfully: {successCount}");
            builder.AppendLine($"Emails failed: {failureCount}");

            if (errorSamples.Count > 0)
            {
                builder.AppendLine();
                builder.AppendLine("Sample errors:");
                foreach (var error in errorSamples)
                {
                    builder.AppendLine(error);
                }

                if (failureCount > errorSamples.Count)
                {
                    builder.AppendLine("...");
                }
            }

            if (batchLimitReached)
            {
                builder.AppendLine();
                builder.AppendLine($"Processing stopped after {CsvEmailBatchLimit} emails to avoid throttling. Re-run the import to send the remaining messages.");
            }

            return builder.ToString();
        }

        private static string GetFieldValue(IReadOnlyList<string> fields, int index)
        {
            return index >= 0 && index < fields.Count ? fields[index] : string.Empty;
        }

        private static bool TryMapHeader(IReadOnlyList<string> fields, out int recipientIndex, out int subjectIndex, out int bodyIndex)
        {
            recipientIndex = -1;
            subjectIndex = -1;
            bodyIndex = -1;

            if (fields.Count == 0)
            {
                return false;
            }

            for (var i = 0; i < fields.Count; i++)
            {
                var value = fields[i];

                if (recipientIndex < 0 && Matches(value, "recipient", "email", "emailaddress", "to"))
                {
                    recipientIndex = i;
                    continue;
                }

                if (subjectIndex < 0 && Matches(value, "subject", "title"))
                {
                    subjectIndex = i;
                    continue;
                }

                if (bodyIndex < 0 && Matches(value, "body", "message", "content"))
                {
                    bodyIndex = i;
                }
            }

            if (recipientIndex < 0)
            {
                return false;
            }

            return true;
        }

        private static bool Matches(string value, params string[] candidates)
        {
            foreach (var candidate in candidates)
            {
                if (string.Equals(value, candidate, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static string BuildHtmlBody(string body, string signature)
        {
            var bodyContent = body ?? string.Empty;

            if (!ContainsLikelyHtml(bodyContent))
            {
                bodyContent = WebUtility.HtmlEncode(bodyContent);
                bodyContent = bodyContent.Replace("\r\n", "<br/>").Replace("\n", "<br/>");
            }

            var signatureContent = NormalizeSignatureContent(signature);

            return string.Concat(bodyContent, "<br/><br/>", signatureContent);
        }

        private static bool ContainsLikelyHtml(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return value.Contains("<") && value.Contains(">");
        }

        private static string NormalizeSignatureContent(string signature)
        {
            if (string.IsNullOrEmpty(signature))
            {
                return string.Empty;
            }

            if (ContainsBlockLevelHtml(signature))
            {
                return signature;
            }

            var normalized = signature.Replace("\r\n", "\n").Replace("\r", "\n");
            return normalized.Replace("\n", "<br/>");
        }

        private static bool ContainsBlockLevelHtml(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string[] blockLevelTags =
            {
                "<article", "<div", "<footer", "<header", "<h1", "<h2", "<h3", "<h4", "<h5", "<h6",
                "<li", "<ol", "<p", "<section", "<table", "<tbody", "<td", "<tfoot", "<th", "<thead", "<tr", "<ul"
            };

            foreach (var tag in blockLevelTags)
            {
                if (value.IndexOf(tag, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private void SetUserName(IAccount userInfo)
        {
            string userName = null;

            if (userInfo != null)
            {
                userName = userInfo.Username;
            }

            if (userName == null)
            {
                userName = "Not identified";
            }

            UserName.Content = userName;
        }
    }
}
