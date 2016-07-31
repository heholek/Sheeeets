using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;

namespace Sheeeets
{
    public class OnErrorEventArgs : EventArgs
    {
        public Exception Error;
    }
    public class Session
    {
        public string[] Scopes { get; } = { SheetsService.Scope.Spreadsheets };

        public string ApplicationName = "Sheeeets";

        public UserCredential UserCredentials { get; private set; }
        public ServiceAccountCredential ServiceCredentials { get; private set; }

        public SheetsService Service { get; private set; }

        public bool Authenticating { get; private set; }
        public bool Authenticated { get; private set; }
        public bool AuthenticationError { get; private set; }

        public event EventHandler OnAuthenticationCompleted;
        public event EventHandler<OnErrorEventArgs> OnAuthenticationError;

        public void AuthenticateUser(string authfile, string credpath = null)
        {
            Authenticating = true;
            AuthenticationError = false;
            FileStream stream;
            var authTask = new Task(() =>
            {
                try
                {
                    stream = new FileStream(authfile, FileMode.Open, FileAccess.Read);
                    if (string.IsNullOrWhiteSpace(credpath))
                    {
                        credpath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                        var filefriendlyappname = Path.GetInvalidFileNameChars().Aggregate(ApplicationName, (current, c) => current.Replace(c, '-'));
                        credpath = Path.Combine(credpath, ".credentials/sheets.googleapis.com-" + filefriendlyappname + ".json");
                    }
                    var authbrokertask = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credpath, true));

                    authbrokertask.ContinueWith(task =>
                    {
                        try
                        {
                            UserCredentials = task.Result;
                            Service = new SheetsService(new BaseClientService.Initializer()
                            {
                                HttpClientInitializer = UserCredentials,
                                ApplicationName = ApplicationName
                            });
                            Authenticating = false;
                            Authenticated = true;
                            _docCache.Clear();
                            OnAuthenticationCompleted?.Invoke(this, EventArgs.Empty);
                        }
                        catch (Exception e)
                        {
                            Authenticating = false;
                            AuthenticationError = true;
                            var args = new OnErrorEventArgs
                            {
                                Error = e
                            };
                            OnAuthenticationError?.Invoke(this, args);
                        }
                    });
                    //authbrokertask.Start();
                }
                catch (Exception e)
                {
                    AuthenticationError = true;
                    var args = new OnErrorEventArgs
                    {
                        Error = e
                    };
                    OnAuthenticationError?.Invoke(this, args);
                }
            });
            authTask.Start();
        }

        public void AuthenticateService(string keyfile, string email)
        {
            Authenticating = true;
            AuthenticationError = false;

            var authTask = new Task(() =>
            {
                try
                {
                    var certificate = new X509Certificate2(keyfile, "notasecret", X509KeyStorageFlags.Exportable);
                    ServiceCredentials = new ServiceAccountCredential( new ServiceAccountCredential.Initializer(email) {Scopes = Scopes}.FromCertificate(certificate));
                    Service = new SheetsService( new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = ServiceCredentials,
                        ApplicationName = ApplicationName
                    });
                    Authenticating = false;
                    Authenticated = true;
                    _docCache.Clear();
                    OnAuthenticationCompleted?.Invoke(this, EventArgs.Empty);
                }
                catch (Exception e)
                {
                    AuthenticationError = true;
                    var args = new OnErrorEventArgs
                    {
                        Error = e
                    };
                    OnAuthenticationError?.Invoke(this, args);
                }
            });
            authTask.Start();
        }
        private readonly Dictionary<string, Spreadsheet> _docCache = new Dictionary<string, Spreadsheet>();
        public Spreadsheet this[string id]
        {
            get
            {
                if (_docCache.ContainsKey(id)) return _docCache[id];
                var doc = new Spreadsheet(id, this);
                _docCache.Add(id,doc);
                return doc;
            }
        }
    }
}
