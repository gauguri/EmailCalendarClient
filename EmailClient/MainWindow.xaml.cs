using EmailCalendarsClient.MailSender;
using Microsoft.Identity.Client;
using Microsoft.Win32;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
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
            var messageHtml = _emailService.CreateHtmlEmail(EmailRecipientText.Text,
                EmailHeader.Text, EmailBody.Text);

            await _aadGraphApiDelegatedClient.SendEmailAsync(messageHtml);
            _emailService.ClearAttachments();
        }

        private void AddAttachment(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == true)
            {
                byte[] data = File.ReadAllBytes(dlg.FileName);
                _emailService.AddAttachment(data, dlg.FileName);
            }
        }

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

            try
            {
                using var parser = new TextFieldParser(dlg.FileName)
                {
                    TextFieldType = FieldType.Delimited,
                    HasFieldsEnclosedInQuotes = true
                };
                parser.SetDelimiters(",");

                _emailService.ClearAttachments();

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

                    if (lineNumber == 1 && LooksLikeHeader(fields))
                    {
                        continue;
                    }

                    if (fields.Length < 3)
                    {
                        failureCount++;
                        AppendCsvError(errorSamples, lineNumber, "Expected at least three columns (recipient, subject, body).");
                        continue;
                    }

                    var recipient = fields[0]?.Trim();
                    var subject = fields[1] ?? string.Empty;
                    var body = fields[2] ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(recipient))
                    {
                        failureCount++;
                        AppendCsvError(errorSamples, lineNumber, "Recipient address is missing.");
                        continue;
                    }

                    try
                    {
                        var message = _emailService.CreateStandardEmail(recipient, subject, body);
                        await _aadGraphApiDelegatedClient.SendEmailAsync(message);
                        successCount++;
                    }
                    catch (Exception sendEx)
                    {
                        failureCount++;
                        AppendCsvError(errorSamples, lineNumber, sendEx.Message);
                    }
                }

                var summary = BuildCsvSummary(successCount, failureCount, errorSamples);
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

        private static string BuildCsvSummary(int successCount, int failureCount, ICollection<string> errorSamples)
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

            return builder.ToString();
        }

        private static bool LooksLikeHeader(IReadOnlyList<string> fields)
        {
            if (fields.Count < 3)
            {
                return false;
            }

            static bool Matches(string value, params string[] candidates)
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

            return Matches(fields[0], "recipient", "email", "emailaddress")
                && Matches(fields[1], "subject", "title")
                && Matches(fields[2], "body", "message", "content");
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
