using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VVVV.PluginInterfaces.V2;
using VVVV.Core.Logging;

namespace Sheeeets.Nodes
{
    [PluginInfo(
        Name = "Spreadsheet",
        Category = "Sheeeets",
        Author = "microdee"
    )]
    public class SheeeetsSpreadsheetNode : IPluginEvaluate
    {
        [Import] public ILogger FLogger;

        [Input("Session")] public Pin<object> FSession;
        [Input("Spreadsheet ID")] public ISpread<string> FSID;
        [Input("Request Data", IsBang = true)] public ISpread<bool> FRequest;

        [Output("Spreadsheet")] public ISpread<Spreadsheet> FSpreadsheet;
        [Output("Sheets")] public ISpread<ISpread<Sheet>> FSheets;
        [Output("Requesting")] public ISpread<bool> FRequesting;
        [Output("Data Available")] public ISpread<bool> FDataAvail;
        [Output("Request Error")] public ISpread<bool> FRequestError;

        private readonly List<string> EventSubscribedSpreadsheets = new List<string>();

        public void Evaluate(int SpreadMax)
        {
            if (FSession.IsConnected && (FSession.SliceCount > 0))
            {
                if (FSession[0] is Session)
                {
                    var session = (Session)FSession[0];
                    if (session.Authenticated)
                    {
                        FSpreadsheet.SliceCount = FSID.SliceCount;
                        FSheets.SliceCount = FSID.SliceCount;
                        FRequesting.SliceCount = FSID.SliceCount;
                        FDataAvail.SliceCount = FSID.SliceCount;
                        FRequestError.SliceCount = FSID.SliceCount;
                        for (int i = 0; i < FSID.SliceCount; i++)
                        {
                            var sprsht = session[FSID[i]];
                            if (!EventSubscribedSpreadsheets.Contains(FSID[i]))
                            {
                                EventSubscribedSpreadsheets.Add(FSID[i]);
                                sprsht.OnRequestError += (sender, args) =>
                                {
                                    FLogger.Log(args.Error, LogType.Error);
                                };
                            }
                            FSpreadsheet[i] = sprsht;
                            FDataAvail[i] = sprsht.DataAvailable;
                            FRequesting[i] = sprsht.Requesting;
                            FRequestError[i] = sprsht.RequestError;
                            if (FRequest[i])
                            {
                                sprsht.RequestSpreadsheetAsync();
                            }
                            if (sprsht.DataAvailable)
                            {
                                FSheets[i].SliceCount = sprsht.Count;
                                for (int j = 0; j < sprsht.Count; j++)
                                {
                                    FSheets[i][j] = sprsht[j];
                                }
                            }
                            else
                            {
                                FSheets[i].SliceCount = 0;
                            }
                        }
                    }
                    else
                    {
                        FSpreadsheet.SliceCount = 0;
                        FRequestError.SliceCount = 0;
                        FSheets.SliceCount = 0;
                        FRequesting.SliceCount = 0;
                        FDataAvail.SliceCount = 0;
                    }
                }
                else
                {
                    FSpreadsheet.SliceCount = 0;
                    FRequestError.SliceCount = 0;
                    FSheets.SliceCount = 0;
                    FRequesting.SliceCount = 0;
                    FDataAvail.SliceCount = 0;
                }
            }
            else
            {
                FSpreadsheet.SliceCount = 0;
                FRequestError.SliceCount = 0;
                FSheets.SliceCount = 0;
                FRequesting.SliceCount = 0;
                FDataAvail.SliceCount = 0;
            }
        }
    }
}
