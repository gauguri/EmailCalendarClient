using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using GraphEmailClient;

namespace EmailCalendarsClient.MailSender
{
    public class AadGraphApiDelegatedClient
    {
        private readonly HttpClient _httpClient = new HttpClient();
        private IPublicClientApplication _app;

        private static readonly string AadInstance = ConfigurationManager.AppSettings["AADInstance"];
        private static readonly string Tenant = ConfigurationManager.AppSettings["Tenant"];
        private static readonly string ClientId = ConfigurationManager.AppSettings["ClientId"];
        private static readonly string Scope = ConfigurationManager.AppSettings["Scope"];

        private static readonly string Authority = string.Format(CultureInfo.InvariantCulture, AadInstance, Tenant);
        private static readonly string[] Scopes = { Scope };

        public void InitClient()
        {
            _app = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority(Authority)
                .WithRedirectUri("http://localhost:65419") // needed only for the system browser
                .Build();

            TokenCacheHelper.EnableSerialization(_app.UserTokenCache);
        }

        public async Task<IAccount> SignIn()
        {
            try
            {
                var result = await AcquireTokenSilent();
                return result.Account;
            }
            catch (MsalUiRequiredException)
            {
                return await AcquireTokenInteractive().ConfigureAwait(false);
            }
        }

        private async Task<IAccount> AcquireTokenInteractive()
        {
            var accounts = (await _app.GetAccountsAsync()).ToList();

            var builder = _app.AcquireTokenInteractive(Scopes)
                .WithAccount(accounts.FirstOrDefault())
                .WithUseEmbeddedWebView(false)
                .WithPrompt(Microsoft.Identity.Client.Prompt.SelectAccount);

            var result = await builder.ExecuteAsync().ConfigureAwait(false);

            return result.Account;
        }

        public async Task<AuthenticationResult> AcquireTokenSilent()
        {
            var accounts = await GetAccountsAsync();
            var result = await _app.AcquireTokenSilent(Scopes, accounts.FirstOrDefault())
                    .ExecuteAsync()
                    .ConfigureAwait(false);

            return result;
        }

        public async Task<IList<IAccount>> GetAccountsAsync()
        {
            var accounts = await _app.GetAccountsAsync();
            return accounts.ToList();
        }

        public async Task RemoveAccountsAsync()
        {
            IList<IAccount> accounts = await GetAccountsAsync();

            // Clears the library cache. Does not affect the browser cookies.
            while (accounts.Any())
            {
                await _app.RemoveAsync(accounts.First());
                accounts = await GetAccountsAsync();
            }
        }

        public async Task SendEmailAsync(Message message, CancellationToken cancellationToken = default)
        {
            var result = await AcquireTokenSilent();

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var graphClient = new GraphServiceClient(_httpClient)
            {
                AuthenticationProvider = new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                    await Task.FromResult<object>(null);
                })
            };

            var saveToSentItems = true;

            await ExecuteWithThrottlingRetries(async () =>
            {
                await graphClient.Me
                    .SendMail(message, saveToSentItems)
                    .Request()
                    .PostAsync(cancellationToken)
                    .ConfigureAwait(false);
            }, cancellationToken).ConfigureAwait(false);
        }

        private static async Task ExecuteWithThrottlingRetries(Func<Task> operation, CancellationToken cancellationToken)
        {
            const int maxAttempts = 3;
            var delay = TimeSpan.Zero;

            for (var attempt = 0; attempt < maxAttempts; attempt++)
            {
                if (delay > TimeSpan.Zero)
                {
                    await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
                }

                try
                {
                    await operation().ConfigureAwait(false);
                    return;
                }
                catch (ServiceException ex) when (IsThrottlingStatus(ex))
                {
                    delay = GetRetryDelay(ex, attempt);
                }
            }

            await operation().ConfigureAwait(false);
        }

        private static bool IsThrottlingStatus(ServiceException exception)
        {
            return exception.StatusCode == System.Net.HttpStatusCode.TooManyRequests
                   || exception.StatusCode == System.Net.HttpStatusCode.ServiceUnavailable;
        }

        private static TimeSpan GetRetryDelay(ServiceException exception, int attempt)
        {
            var retryAfter = exception.ResponseHeadersRetryAfter?.Delta;

            if (retryAfter.HasValue && retryAfter.Value > TimeSpan.Zero)
            {
                return retryAfter.Value;
            }

            var exponentialBackoffSeconds = Math.Min(30, Math.Pow(2, attempt) * 3);
            return TimeSpan.FromSeconds(exponentialBackoffSeconds);
        }

    }
}
