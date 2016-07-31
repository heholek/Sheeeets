using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VVVV.Core.Logging;
using VVVV.PluginInterfaces.V2;

namespace Sheeeets.Nodes
{
    [PluginInfo(
        Name = "ServiceSession",
        Category = "Sheeeets",
        Author = "microdee"
    )]
    public class SheeeetsSessionNode : IPluginEvaluate
    {
        [Import] public ILogger FLogger;

        [Input("Authenticating File", StringType = StringType.Filename, FileMask = "All Files (*.*)|*.*")]
        public ISpread<string> FAuthFilePath;

        [Input("Application Name", DefaultString = "Sheeeets")]
        public ISpread<string> FAppName;
        [Input("Email")]
        public ISpread<string> FUser;

        [Input("Authenticate", IsBang = true, IsSingle = true)]
        public ISpread<bool> FAuthenticate;

        [Output("Session")]
        public ISpread<Session> FSession;
        [Output("Authenticating")]
        public ISpread<bool> FAuthInProg;
        [Output("Authenticated")]
        public ISpread<bool> FAuthSuccess;
        [Output("Authentication Error")]
        public ISpread<bool> FAuthError;

        public void Evaluate(int SpreadMax)
        {
            if (FAuthenticate[0])
            {
                if (File.Exists(FAuthFilePath[0]))
                {
                    FSession[0] = new Session
                    {
                        ApplicationName = FAppName[0]
                    };
                    FSession[0].OnAuthenticationError += (sender, args) =>
                    {
                        FLogger.Log(args.Error, LogType.Error);
                    };
                    FSession[0].AuthenticateService(FAuthFilePath[0], FUser[0]);
                }
            }
            if (FSession[0] != null)
            {
                FAuthInProg[0] = FSession[0].Authenticating;
                FAuthSuccess[0] = FSession[0].Authenticated;
                FAuthError[0] = FSession[0].AuthenticationError;
            }
        }
    }
}
